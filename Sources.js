/**
 * Direct source loader (replaces the IMPORTRANGE-fed tabs of the common sheet).
 * Uses the Sheets advanced service (Sheets.Spreadsheets.Values.batchGet).
 *
 * Every loader returns the same 2D shape the old common-sheet tab had: header rows
 * first (as listed in headerRows), then the stacked data rows. For sources listed in
 * SOURCES, every row carries two trailing meta cells [idIndex, sourceRow] so that
 * hyperlinks can be resolved later for just the rows that are shown.
 *
 * Speed: one batchGet per spreadsheet, only the project's bucket spreadsheet is read,
 * and results are filtered to the project's ISO before being cached (10 min).
 */

// The 9 project spreadsheets that were stacked by IMPORTRANGE, in the original order.
const PROJECT_SPREADSHEET_IDS = [
  '1nGg31RHluzyBVIM16DpPXZ4Rd7n3TeSxWvYowUShnrI',
  '1VhZnUCRWyYPUfeeMJE91vSWb3GL7AG_7_yXMDl7VsTU',
  '1vnLf_Yyeb4oUg-Gs1Rnu5X73kPfkwdMD7OiVX1XdigE',
  '1UT-W4K_i4rNYxMBB0DW23Qla1QPkHPaHoJKsp7u5DD0',
  '1YhTCYrxnz35KTdUGf7smIzkudFRxm1QTDFMO74lESlY',
  '1GIKkjeUZU5vgqYmQ2hItilFVQjC9kTLepT4Rl6L9Q_E',
  '1R6WToiXNpZzaY2BbgGrMAVVQ-IsLcfZ7Xv-mxmMKzJc',
  '1mc9wxCCpaxkgDQjZzeHR-DjoVacc8ss6A4s-x2ipyfs',
  '144Y_Vnns_NuP2v0TfYFFApzrjAVrp-67c9kj9OzDTvQ'
];

const MOM_SPREADSHEET_ID = '13ZZ-xJvQOvEXDGiEKxL7b1nhzIgfr2gZvtmcHJLaOOw';

/**
 * Key = tab name the rest of the code already uses.
 *  ids        spreadsheets to stack (in order)
 *  tab        source tab name
 *  blocks     column blocks, joined side by side (rows start at source row 1)
 *  headerRows source rows (1-based) kept as header rows at the top of the result
 *  dataStart  first source data row (1-based)
 *  skipIds    spreadsheets that do not have this tab
 *  byIso      rows can be reduced to one ISO before caching
 *
 * headerRows / dataStart come from the old IMPORTRANGE start rows - check them with
 * verifySources().
 */
const SOURCES = {
  'PROJECT PLANNING':       { ids: PROJECT_SPREADSHEET_IDS, tab: 'PROJECT PLANNING', blocks: ['A:T'],  headerRows: [1],    dataStart: 2, byIso: true },
  'Manpower Cost':          { ids: PROJECT_SPREADSHEET_IDS, tab: 'Manpower Cost',    blocks: ['A:C'],  headerRows: [1],    dataStart: 2, byIso: true },
  // Common sheet: =QUERY(VSTACK(IFERROR(IMPORTRANGE(<each of the 9>, "INVENTORY!A3:X50000"),)...), "SELECT * WHERE Col1 IS NOT NULL", 0)
  // Both header rows are typed there; data starts at source row 3.
  'INVENTORY':              { ids: PROJECT_SPREADSHEET_IDS, tab: 'INVENTORY',        blocks: ['A:X'],  headerRows: [1, 2], dataStart: 3, byIso: true,
                              headers: [
                                ['', '', '', '', 'QNTY', '', '', '', '', 'VALUE(RS)', '', '', '', '', 'STATUS', '', 'INVENTORY IN TRANSIT', '', '', '', '', '', '', ''],
                                ['ISO NO.', 'MILESTONE', 'TASK', 'PRODUCT NAME', 'BOQ/TENDER1', 'INDENT1', 'REC.1', 'CONSU.1', 'AVAIL.1', 'BOQ/TENDER2', 'INDENT2', 'REC.2', 'CONSU.2', 'AVAIL.2', 'MILESTONE', 'TASK', 'INDENT NO.', 'Qty', 'PO NO', 'VENDOR', 'DISPATCH DET.', 'EXPECTED DATE OF REC.', 'ISO No.', 'Planned Indent Qty']
                              ],
                              requireFirstCol: true },
  'Auditor Sheet Manpower': { ids: PROJECT_SPREADSHEET_IDS, tab: 'Auditor Sheet',    blocks: ['A:J'],  headerRows: [1, 2], dataStart: 3, byIso: true },
  'Auditor Sheet Item':     { ids: PROJECT_SPREADSHEET_IDS, tab: 'Auditor Sheet',    blocks: ['K:U'],  headerRows: [1, 2], dataStart: 3, byIso: true },
  'Auditor Sheet Drawing':  { ids: PROJECT_SPREADSHEET_IDS, tab: 'Auditor Sheet',    blocks: ['W:AB'], headerRows: [1, 2], dataStart: 3, byIso: true },
  // Same as the common sheet's =LET(VSTACK(IMPORTRANGE(... "WBS Sheet!A4:M")...), FILTER(data, col 1 <> "")):
  // headers are typed there (not read from the source) and rows without a first cell are dropped.
  'WBS Sheet':              { ids: PROJECT_SPREADSHEET_IDS, tab: 'WBS Sheet',        blocks: ['A:M'],  headerRows: [3],    dataStart: 4, byIso: true,
                              headers: [['ISO No.', 'CLIENT NAME', 'MILESTONE/BOQ ITEM', 'TASK NAME', 'BOQ Item', 'QTY', 'PRECEDING WORK', 'TIME REQ.-DAYS', 'Engineer Name', 'UNIQUE ID', 'PLANNED START DATE', 'ACTUAL END DATE', 'Delay']],
                              requireFirstCol: true,
                              skipIds: ['1UT-W4K_i4rNYxMBB0DW23Qla1QPkHPaHoJKsp7u5DD0'] },
  'Project Details Form':   { ids: ['1GrciDwR0Qcc1HKrm34bOeGKwEjyqH9BlDuTFDRM1N8U'], tab: 'Project_Details', blocks: ['A:AC'], headerRows: [1], dataStart: 2, byIso: true },
  'Expenses':               { ids: ['1uHi7hEW69NyNG2hNrUmY6Qh6XFKoBKeCvE9zGbT6bKo'], tab: 'Journal',         blocks: ['A:G'],  headerRows: [1], dataStart: 2, byIso: true },
  'Billing':                { ids: ['1mQyuZNKhM4d2n_ARCBd_IUcuwEnrUNl-I_lzF_lfozg'], tab: 'FMS.',            blocks: ['A:N', 'AE:AE'], headerRows: [8], dataStart: 9, byIso: true },
  'MOM':                    { ids: [MOM_SPREADSHEET_ID],                              tab: 'Response',        blocks: ['A:AS'], headerRows: [1], dataStart: 2, byIso: true }
};

/**
 * Tabs still read from the COMMON sheet (the old IMPORTRANGE tabs) because their direct
 * mapping (typed header rows, helper columns) is not confirmed yet. Remove a name from here
 * to switch that tab to its source spreadsheet(s) in SOURCES once it matches the old data.
 */
const LOCAL_TABS = {
  'PROJECT PLANNING': true, 'Manpower Cost': true, 'INVENTORY': true,
  'Auditor Sheet Manpower': true, 'Auditor Sheet Item': true, 'Auditor Sheet Drawing': true,
  'WBS Sheet': true, 'Billing': true, 'Expenses': true, 'MOM': true,
  'Project Details Form': true // the common sheet's copy has the extra 'item' column the Product field reads
};
/** Header rows at the top of a common-sheet tab (default 1). */
const LOCAL_HEADER_ROWS = { 'INVENTORY': 2, 'Auditor Sheet Manpower': 2, 'Auditor Sheet Item': 2, 'Auditor Sheet Drawing': 2 };

/**
 * Rows of one ISO can live in several of the 9 project spreadsheets (e.g. Manpower Cost,
 * Audit), so the default is to read all 9. Set true only if every tab of a project is
 * confirmed to sit in its bucket's spreadsheet.
 */
const ROUTE_BY_BUCKET = false;

/** Tabs rebuilt from raw helper columns instead of read as laid out (see loadPlanningDirect_). */
const DERIVED_TABS = { 'PROJECT PLANNING': true };

/** forceSource ('direct' | 'local') overrides LOCAL_TABS, e.g. for compareSources(). */
function isDirect_(name, forceSource) {
  if (forceSource) return !!SOURCES[name] && forceSource === 'direct';
  return !!SOURCES[name] && !LOCAL_TABS[name];
}

/** Number of header rows at the top of the loaded result for a tab. */
function headerCount_(name, forceSource) {
  return isDirect_(name, forceSource) ? SOURCES[name].headerRows.length : (LOCAL_HEADER_ROWS[name] || 1);
}

const SOURCE_CACHE_TTL = 600; // seconds; data is up to 10 minutes old
const CACHE_CHUNK = 90000;    // CacheService limit is 100KB per value
const CACHE_PREFIX = 'src3:';
const META_COLS = 2;          // trailing [idIndex, sourceRow] on every SOURCES row

/**
 * Loads the given tab names. Names without an entry in SOURCES (e.g. FINANCE/ACCOUNTS)
 * fall back to the local tab of localSs.
 *  projectIds  restrict the 9 stacked project sources to these spreadsheets (or null = all)
 *  opts        { formatted: true => display strings, iso: keep only this ISO's rows }
 * @return {Object<string, any[][]>}
 */
function loadSources_(names, localSs, projectIds, opts) {
  opts = opts || {};
  const t0 = Date.now();
  const out = {};
  const missing = [];

  names.forEach(name => {
    const src = SOURCES[name];
    if (!src || !isDirect_(name, opts.source)) {
      out[name] = loadLocal_(localSs, name, opts);
      return;
    }
    const hit = cacheable_(src, opts) ? cacheGet_(cacheKey_(name, projectIds, opts)) : null;
    if (hit) out[name] = hit; else missing.push(name);
  });

  if (missing.length) {
    // Throws if a source cannot be read, so incomplete data is never cached.
    const fresh = loadSourcesFresh_(missing, localSs, projectIds, !!opts.formatted, opts.iso);
    missing.forEach(name => {
      const src = SOURCES[name];
      let data = fresh[name];
      if (src.byIso && opts.iso) data = filterByIso_(name, data, opts.iso, 'direct');
      out[name] = data;
      if (cacheable_(src, opts)) cachePut_(cacheKey_(name, projectIds, opts), data);
    });
  }

  Logger.log('loadSources_: ' + (names.length - missing.length) + ' cached, ' + missing.length + ' fetched in ' + (Date.now() - t0) + ' ms');
  return out;
}

/** Whole 15k-row tabs are not worth caching; only ISO-filtered (small) results are. */
function cacheable_(src, opts) {
  return !src.byIso || !!opts.iso;
}

/** Only the 9 stacked project sources depend on projectIds; the others are single-ID. */
function idsFor_(src, projectIds) {
  return (projectIds && src.ids === PROJECT_SPREADSHEET_IDS) ? projectIds : src.ids;
}

function cacheKey_(name, projectIds, opts) {
  const src = SOURCES[name];
  let key = name;
  if (projectIds && src.ids === PROJECT_SPREADSHEET_IDS) key += '@' + projectIds.join('|');
  if (opts.formatted) key += '|f';
  if (src.byIso && opts.iso) key += '|' + opts.iso;
  return key;
}

/** ISO column header, tolerant of "ISO No.", "ISO NO", "ISO NO." (but not "ISO No 2"). */
function isoColIdx_(hdr) {
  for (let i = 0; i < hdr.length; i++) {
    if (String(hdr[i]).trim().toUpperCase().replace(/\.$/, '') === 'ISO NO') return i;
  }
  return -1;
}

function filterByIso_(name, data, iso, forceSource) {
  const hIdx = headerCount_(name, forceSource) - 1;
  const ci = isoColIdx_(data[hIdx] || []);
  if (ci === -1) return data;
  return data.slice(0, hIdx + 1).concat(data.slice(hIdx + 1).filter(r => cleanIso(r[ci]) === iso));
}

/** Dates from SpreadsheetApp become serial numbers, like the Sheets API, so cached JSON stays usable. */
function dateToSerial_(v) {
  return Object.prototype.toString.call(v) === '[object Date]' ? (v.getTime() - v.getTimezoneOffset() * 60000) / 86400000 + 25569 : v;
}

/**
 * Reads a tab of the common sheet. Rows of tabs listed in SOURCES get the same trailing
 * [idIndex, sourceRow] meta as direct rows, and are cut down to one ISO and cached like them.
 */
function loadLocal_(ss, name, opts) {
  const src = SOURCES[name];
  if (!src) { // e.g. FINANCE/ACCOUNTS: the raw tab
    const raw = ss.getSheetByName(name);
    return raw ? (opts.formatted ? raw.getDataRange().getDisplayValues() : raw.getDataRange().getValues()) : [];
  }
  const cacheOk = !src.byIso || !!opts.iso;
  const key = 'local|' + name + (opts.formatted ? '|f' : '') + (src.byIso && opts.iso ? '|' + opts.iso : '');
  if (cacheOk) { const hit = cacheGet_(key); if (hit) return hit; }

  const sh = ss.getSheetByName(name);
  if (!sh) return [];
  const range = sh.getDataRange();
  const vals = opts.formatted ? range.getDisplayValues() : range.getValues();
  const hc = headerCount_(name, 'local');
  const out = [];
  for (let r = 0; r < vals.length; r++) {
    const row = opts.formatted ? vals[r] : vals[r].map(dateToSerial_);
    if (r < hc) { out.push(row.concat(['', ''])); continue; }
    if (row.some(c => c !== '' && c !== null)) out.push(row.concat([0, r + 1]));
  }
  const data = (src.byIso && opts.iso) ? filterByIso_(name, out, opts.iso, 'local') : out;
  if (cacheOk) cachePut_(key, data);
  return data;
}

/** Link replacement for common-sheet rows: reads the rich text of just the link columns. */
function resolveLinksLocal_(ss, name, hdr, rows, linkNames) {
  const sh = ss.getSheetByName(name);
  if (!sh || !rows.length) return;
  const last = sh.getLastRow();
  linkNames.forEach(n => {
    const ci = hdr.findIndex(h => String(h).trim() === n);
    if (ci < 0 || ci >= hdr.length - META_COLS) return;
    const rt = sh.getRange(1, ci + 1, last, 1).getRichTextValues();
    rows.forEach(row => {
      const sr = row[row.length - 1];
      const cell = sr && rt[sr - 1] && rt[sr - 1][0];
      const link = cell && cell.getLinkUrl();
      if (link) row[ci] = link;
    });
  });
}

function cacheGet_(key) {
  try {
    const cache = CacheService.getScriptCache();
    const n = Number(cache.get(CACHE_PREFIX + key + ':n'));
    if (!n) return null;
    const keys = [];
    for (let i = 0; i < n; i++) keys.push(CACHE_PREFIX + key + ':' + i);
    const parts = cache.getAll(keys);
    let b64 = '';
    for (let i = 0; i < n; i++) {
      if (!parts[keys[i]]) return null;
      b64 += parts[keys[i]];
    }
    const json = Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(b64), 'application/x-gzip')).getDataAsString();
    return JSON.parse(json);
  } catch (e) {
    return null;
  }
}

function cachePut_(key, data) {
  try {
    const b64 = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(JSON.stringify(data))).getBytes());
    const entries = {};
    let n = 0;
    for (let i = 0; i < b64.length; i += CACHE_CHUNK) entries[CACHE_PREFIX + key + ':' + (n++)] = b64.substring(i, i + CACHE_CHUNK);
    entries[CACHE_PREFIX + key + ':n'] = String(n);
    CacheService.getScriptCache().putAll(entries, SOURCE_CACHE_TTL);
  } catch (e) {
    Logger.log('cachePut_ failed for ' + key + ': ' + e.message);
  }
}

/** Cache wrapper for small computed results (tables) keyed by ISO. */
function cachedResult_(key, compute) {
  const hit = cacheGet_('res|' + key);
  if (hit) return hit;
  const value = compute();
  cachePut_('res|' + key, value);
  return value;
}

/**
 * Run from the editor to force the next load to re-read the sources. Cached entries are
 * keyed by ISO and expire on their own after 10 minutes; this bumps nothing, so for an
 * immediate full refresh change CACHE_PREFIX (e.g. 'src4:') and push.
 */
function clearSourceCache() {
  const cache = CacheService.getScriptCache();
  cache.removeAll(['src3:projIdx:n', 'src3:bucketTitles:n', 'src3:projIdx:0', 'src3:bucketTitles:0']);
  Logger.log('Index and title caches cleared. For everything else, bump CACHE_PREFIX in Sources.js.');
}

/** Spreadsheet titles (e.g. "PMS-STP CIVIL") for the 9 project spreadsheets, cached. */
function projectSpreadsheetTitles_() {
  const hit = cacheGet_('bucketTitles');
  if (hit) return hit;
  const titles = {};
  let complete = true;
  PROJECT_SPREADSHEET_IDS.forEach(id => {
    try { titles[id] = Sheets.Spreadsheets.get(id, { fields: 'properties.title' }).properties.title; }
    catch (e) { complete = false; Logger.log('Title lookup failed for ' + id + ': ' + e.message); }
  });
  if (complete) cachePut_('bucketTitles', titles);
  return titles;
}

function normBucket_(s) {
  return String(s || '').toUpperCase().replace(/^PMS/, '').replace(/[^A-Z0-9]/g, '');
}

/**
 * Spreadsheet IDs for a bucket such as "STP CIVIL", matched on the spreadsheet title
 * ("PMS-STP CIVIL"). Returns null when there is no single match, meaning "read all 9".
 */
function bucketSpreadsheetIds_(bucket) {
  if (!ROUTE_BY_BUCKET) return null;
  const want = normBucket_(bucket);
  if (!want) return null;
  const titles = projectSpreadsheetTitles_();
  const hits = PROJECT_SPREADSHEET_IDS.filter(id => titles[id] && normBucket_(titles[id]) === want);
  if (hits.length !== 1) {
    Logger.log('Bucket "' + bucket + '" matched ' + hits.length + ' spreadsheets; reading all. Titles: ' + JSON.stringify(titles));
    return null;
  }
  return hits;
}

/**
 * Tiny project lookup (project name -> ISO + bucket), cached. Lets the table loaders find
 * the ISO and the right spreadsheet without loading Project Details every time.
 * @return {{byName: Object, byIso: Object, pairs: string[][]}}
 */
function projectIndex_() {
  const hit = cacheGet_('projIdx');
  if (hit) return hit;
  const rows = loadSourcesFresh_(['Project Details Form'], null, null, false)['Project Details Form'];
  if (!rows || rows.length < 2) throw new Error('Project Details Form returned no rows.');
  const hdr = rows[0].map(h => String(h).trim());
  const n = hdr.indexOf('PROJECT NAME');
  const i = isoColIdx_(hdr);
  const b = hdr.indexOf('BUCKET FIELD');
  if (n === -1 || i === -1) throw new Error("'PROJECT NAME' or 'ISO No.' column not found in Project Details Form.");
  const idx = { byName: {}, byIso: {}, pairs: [] };
  const seen = {};
  rows.slice(1).forEach(r => {
    const name = r[n];
    if (!name) return;
    const iso = cleanIso(r[i]);
    const bucket = b > -1 ? String(r[b] || '') : '';
    if (!idx.byName[name]) idx.byName[name] = { iso: iso, bucket: bucket };
    if (iso && !idx.byIso[iso]) idx.byIso[iso] = { name: name, bucket: bucket };
    if (!seen[name + '\u0001' + bucket]) { seen[name + '\u0001' + bucket] = 1; idx.pairs.push([name, bucket]); }
  });
  cachePut_('projIdx', idx);
  return idx;
}

const CHUNK_ROWS = 4000;     // rows per range
const BATCH_CELLS = 200000;  // cells per batchGet call
const RETRY_WAITS_MS = [1000, 3000];

/** Retries transient Sheets API errors, but only while a failed attempt was quick. */
function sheetsRetry_(label, fn) {
  for (let i = 0; ; i++) {
    const t = Date.now();
    try {
      return fn();
    } catch (e) {
      const transient = /unavailable|timed out|timeout|internal error|backend error|rate limit|quota|(500|502|503|504)/i.test(e.message);
      if (!transient || i >= RETRY_WAITS_MS.length || Date.now() - t > 20000) throw e;
      Logger.log(label + ': attempt ' + (i + 1) + ' failed (' + e.message + '), retrying');
      Utilities.sleep(RETRY_WAITS_MS[i]);
    }
  }
}

/** Allocated row count per tab of a spreadsheet. */
function gridRows_(id) {
  const meta = sheetsRetry_('gridRows ' + id, () => Sheets.Spreadsheets.get(id, { fields: 'sheets.properties(title,gridProperties.rowCount)' }));
  const rows = {};
  meta.sheets.forEach(s => { rows[s.properties.title] = s.properties.gridProperties.rowCount; });
  return rows;
}

/**
 * Reads every request from one spreadsheet as row-chunked ranges, batched under a cell
 * budget, so no single call is huge. Returns one matrix per request, row-aligned to the
 * sheet (row 1 = index 0), with trailing blank rows trimmed.
 * reqs: [{ tab, block }]
 */
function readSpreadsheetRanges_(id, reqs, renderOption) {
  const t0 = Date.now();
  let sizes = null;
  try { sizes = gridRows_(id); }
  catch (e) { throw new Error('Could not read source spreadsheet ' + id + ' (' + e.message + '). Check the ID and that the account running the web app can open it.'); }

  // Plan chunks.
  const chunks = []; // { reqIdx, start, end, last, range, cells }
  const results = reqs.map(() => []);
  reqs.forEach((r, ri) => {
    const rc = sizes[r.tab];
    if (!rc) { Logger.log('Tab "' + r.tab + '" not found in ' + id + '; skipped'); return; }
    const parts = r.block.split(':');
    const width = blockWidth_(r.block);
    for (let start = 1; start <= rc; start += CHUNK_ROWS) {
      const end = Math.min(start + CHUNK_ROWS - 1, rc);
      chunks.push({
        reqIdx: ri, start: start, end: end, last: end === rc, tab: r.tab, c1: parts[0], c2: parts[1],
        range: "'" + r.tab + "'!" + parts[0] + start + ':' + parts[1] + end,
        cells: (end - start + 1) * width
      });
    }
  });

  // Group chunks into batches under the cell budget.
  const batches = [];
  let cur = [], curCells = 0;
  chunks.forEach(c => {
    if (cur.length && curCells + c.cells > BATCH_CELLS) { batches.push(cur); cur = []; curCells = 0; }
    cur.push(c); curCells += c.cells;
  });
  if (cur.length) batches.push(cur);

  batches.forEach((batch, bi) => {
    const t = Date.now();
    let valueRanges;
    try {
      valueRanges = sheetsRetry_('batchGet ' + id, () =>
        Sheets.Spreadsheets.Values.batchGet(id, { ranges: batch.map(c => c.range), valueRenderOption: renderOption }).valueRanges);
    } catch (e) {
      Logger.log('batchGet failed for ' + id + ' (batch ' + (bi + 1) + '/' + batches.length + ', ' + batch.length + ' ranges) after ' + (Date.now() - t) + ' ms: ' + e.message + ' - trying SpreadsheetApp');
      // Some tabs the Sheets API cannot serve (heavy formulas, "service unavailable"); the
      // SpreadsheetApp service reads the same cells through a different path.
      try {
        valueRanges = batch.map(c => ({ values: readChunkViaSpreadsheetApp_(id, c, renderOption) }));
        Logger.log('SpreadsheetApp fallback ok for ' + id + ' (' + batch.length + ' ranges) in ' + (Date.now() - t) + ' ms');
      } catch (e2) {
        throw new Error('Could not read source spreadsheet ' + id + ' while reading ' + batch[0].range + '. Sheets API: ' + e.message + ' | SpreadsheetApp: ' + e2.message);
      }
    }
    batch.forEach((c, k) => {
      const vals = (valueRanges[k] && valueRanges[k].values) || [];
      const m = results[c.reqIdx];
      // Keep rows aligned with the sheet: a non-final chunk is padded to its full height.
      for (let r = 0; r < vals.length; r++) m.push(vals[r]);
      if (!c.last) while (m.length < c.end) m.push([]);
    });
  });

  Logger.log('read ' + id + ': ' + chunks.length + ' chunks in ' + batches.length + ' calls, ' + (Date.now() - t0) + ' ms');
  return results;
}

const _ssHandles = {};

/** Same shape as a Sheets API values response (trailing blank rows trimmed, dates as serials). */
function readChunkViaSpreadsheetApp_(id, c, renderOption) {
  const ss = _ssHandles[id] || (_ssHandles[id] = SpreadsheetApp.openById(id));
  const sh = ss.getSheetByName(c.tab);
  if (!sh) throw new Error('tab "' + c.tab + '" not found');
  const col = colToNum_(c.c1);
  const width = colToNum_(c.c2) - col + 1;
  const rows = Math.min(c.end, sh.getMaxRows()) - c.start + 1;
  if (rows < 1) return [];
  const range = sh.getRange(c.start, col, rows, width);
  const vals = renderOption === 'FORMATTED_VALUE' ? range.getDisplayValues() : range.getValues();
  const out = vals.map(row => row.map(v =>
    Object.prototype.toString.call(v) === '[object Date]' ? (v.getTime() - v.getTimezoneOffset() * 60000) / 86400000 + 25569 : v));
  let last = out.length;
  while (last > 0 && out[last - 1].every(v => v === '' || v === null)) last--;
  return out.slice(0, last);
}


function loadSourcesFresh_(names, localSs, projectIds, formatted, iso) {
  const out = {};
  const requests = [];

  names.forEach(name => {
    const src = SOURCES[name];
    if (!src) {
      const sh = localSs.getSheetByName(name);
      out[name] = sh ? sh.getDataRange().getValues() : [];
      return;
    }
    if (DERIVED_TABS[name]) {
      out[name] = loadPlanningDirect_(idsFor_(src, projectIds), iso, formatted);
      return;
    }
    idsFor_(src, projectIds).forEach(id => {
      if ((src.skipIds || []).indexOf(id) > -1) return;
      src.blocks.forEach((block, bi) => requests.push({ name: name, id: id, bi: bi, tab: src.tab, block: block }));
    });
  });

  // One batchGet per spreadsheet, covering every tab we need from it.
  const byId = {};
  requests.forEach(r => { (byId[r.id] = byId[r.id] || []).push(r); });

  const grids = {}; // name -> id -> block index -> matrix
  Object.keys(byId).forEach(id => {
    const reqs = byId[id];
    const matrices = readSpreadsheetRanges_(id, reqs, formatted ? 'FORMATTED_VALUE' : 'UNFORMATTED_VALUE');
    reqs.forEach((r, i) => {
      const g = grids[r.name] = grids[r.name] || {};
      const blocks = g[r.id] = g[r.id] || [];
      blocks[r.bi] = matrices[i];
    });
  });

  names.forEach(name => {
    const src = SOURCES[name];
    if (!src || DERIVED_TABS[name]) return;
    const widths = src.blocks.map(blockWidth_);
    let header = null;
    const data = [];

    idsFor_(src, projectIds).forEach(id => {
      const blocks = (grids[name] || {})[id];
      if (!blocks) return;
      const idIdx = src.ids.indexOf(id);
      const rowCount = Math.max.apply(null, blocks.map(b => (b || []).length).concat(0));
      const grid = [];
      for (let r = 0; r < rowCount; r++) {
        let row = [];
        blocks.forEach((b, bi) => {
          const cells = ((b || [])[r] || []).slice(0, widths[bi]);
          while (cells.length < widths[bi]) cells.push('');
          row = row.concat(cells);
        });
        grid.push(row);
      }
      if (!header) {
        if (src.headers) header = src.headers.map(h => h.concat(['', '']));
        else if (grid.length >= Math.max.apply(null, src.headerRows)) header = src.headerRows.map(h => grid[h - 1].concat(['', '']));
      }
      for (let r = src.dataStart - 1; r < grid.length; r++) {
        const row = grid[r];
        const keep = src.requireFirstCol ? (row[0] !== '' && row[0] !== null) : row.some(c => c !== '' && c !== null);
        if (keep) data.push(row.concat([idIdx, r + 1]));
      }
    });

    out[name] = (header || []).concat(data);
  });

  return out;
}

/**
 * Replaces the text of link cells with their URL, but only for the rows passed in.
 * One Sheets API call per 40 cells; the response holds just the hyperlink field.
 * rows must come from a SOURCES load (they carry the trailing [idIndex, sourceRow] meta).
 */
function resolveLinks_(name, hdr, rows, linkNames) {
  const src = SOURCES[name];
  if (!src || !rows.length) return;
  if (!isDirect_(name)) { resolveLinksLocal_(SpreadsheetApp.getActiveSpreadsheet(), name, hdr, rows, linkNames); return; }
  const cols = linkNames
    .map(n => hdr.findIndex(h => String(h).trim() === n))
    .filter(i => i > -1 && i < hdr.length - META_COLS);
  if (!cols.length) return;

  const byId = {};
  rows.forEach((row, ri) => {
    const id = src.ids[row[row.length - META_COLS]];
    const sr = row[row.length - 1];
    if (!id || !sr) return;
    cols.forEach(ci => {
      const letter = colLetter_(src, ci);
      if (letter) (byId[id] = byId[id] || []).push({ ri: ri, ci: ci, a1: "'" + src.tab + "'!" + letter + sr });
    });
  });

  Object.keys(byId).forEach(id => {
    const items = byId[id];
    for (let i = 0; i < items.length; i += 40) {
      const chunk = items.slice(i, i + 40);
      try {
        const res = Sheets.Spreadsheets.get(id, { ranges: chunk.map(c => c.a1), fields: 'sheets.data.rowData.values.hyperlink' });
        const data = (res.sheets && res.sheets[0] && res.sheets[0].data) || [];
        chunk.forEach((c, k) => {
          const rd = data[k] && data[k].rowData;
          const link = rd && rd[0] && rd[0].values && rd[0].values[0] && rd[0].values[0].hyperlink;
          if (link) rows[c.ri][c.ci] = link;
        });
      } catch (e) {
        Logger.log('resolveLinks_ failed for ' + id + ': ' + e.message);
      }
    }
  });
}

/** Sheets API returns dates as serial numbers; the chart code expects Date objects. */
function toDate_(v) {
  if (v instanceof Date) return v;
  if (typeof v === 'number' && v > 0) return new Date(Math.round((v - 25569) * 86400000));
  return null;
}

function blockRange_(tab, block) {
  const parts = block.split(':');
  return "'" + tab + "'!" + parts[0] + '1:' + parts[1];
}

function blockWidth_(block) {
  const parts = block.split(':');
  return colToNum_(parts[1]) - colToNum_(parts[0]) + 1;
}

function colToNum_(letters) {
  let n = 0;
  for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
  return n;
}

function numToCol_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

/** Column letter in the source tab for the idx-th column of a (possibly multi-block) source. */
function colLetter_(src, idx) {
  let off = idx;
  for (let i = 0; i < src.blocks.length; i++) {
    const w = blockWidth_(src.blocks[i]);
    if (off < w) return numToCol_(colToNum_(src.blocks[i].split(':')[0]) + off);
    off -= w;
  }
  return null;
}

/** values.batchGet with retry; returns one matrix per range. */
function valuesBatch_(id, ranges, renderOption) {
  const vr = sheetsRetry_('batchGet ' + id, () =>
    Sheets.Spreadsheets.Values.batchGet(id, { ranges: ranges, valueRenderOption: renderOption }).valueRanges);
  return ranges.map((r, i) => (vr[i] && vr[i].values) || []);
}

/** Reads one column of a tab for rows 1..rowCount in chunks; returns a flat array (index = row-1). */
function readColumns_(id, tab, letters, rowCount, renderOption) {
  const cols = letters.map(() => []);
  const ranges = [], meta = [];
  for (let start = 1; start <= rowCount; start += CHUNK_ROWS) {
    const end = Math.min(start + CHUNK_ROWS - 1, rowCount);
    letters.forEach((L, li) => { ranges.push("'" + tab + "'!" + L + start + ':' + L + end); meta.push({ li: li, start: start, end: end, last: end === rowCount }); });
  }
  const perCall = Math.max(1, Math.floor(BATCH_CELLS / CHUNK_ROWS));
  for (let i = 0; i < ranges.length; i += perCall) {
    const m = valuesBatch_(id, ranges.slice(i, i + perCall), renderOption);
    m.forEach((vals, k) => {
      const c = meta[i + k];
      const arr = cols[c.li];
      vals.forEach(r => arr.push(r.length ? r[0] : ''));
      if (!c.last) while (arr.length < c.end) arr.push('');
    });
  }
  return cols;
}

function fmtDate_(v) {
  if (typeof v === 'number' && v > 0) {
    const d = new Date(Math.round((v - 25569) * 86400000));
    const p = n => (n < 10 ? '0' : '') + n;
    return p(d.getUTCDate()) + '/' + p(d.getUTCMonth() + 1) + '/' + d.getUTCFullYear();
  }
  return v === null || v === undefined ? '' : String(v);
}

function fmtNum_(v) {
  return typeof v === 'number' ? String(Number(v.toFixed(10))) : (v === null || v === undefined ? '' : String(v));
}

/**
 * PROJECT PLANNING without its formulas. In each project spreadsheet the visible columns
 * B..S are FILTER()s over plain helper columns (W..AR) where AR = "Project", and T is
 * round(rate from RESOURCES * worked hours). This reads the ISO and type columns, picks the
 * matching rows, fetches only those rows' helper cells and rebuilds A..T, so the slow formulas
 * are never evaluated. Same shape as the old tab: header row first, then rows + [idIndex, 0].
 */
function loadPlanningDirect_(ids, iso, formatted) {
  const tab = 'PROJECT PLANNING';
  const src = SOURCES[tab];
  const W = colToNum_('W');
  // output column -> helper column (null = counter / computed)
  const map = [null, 'AA', 'AB', 'AC', 'AF', 'AE', 'W', 'AD', 'AF', 'W', 'X', 'AG', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AP', 'AQ'];
  const dateCols = { 1: true, 12: true };
  const rd = (row, L) => { const v = row[colToNum_(L) - W]; return v === undefined || v === null ? '' : v; };
  let header = null;
  const data = [];

  ids.forEach(id => {
    if ((src.skipIds || []).indexOf(id) > -1) return;
    const idIdx = src.ids.indexOf(id);
    const sizes = gridRows_(id);
    const rc = sizes[tab];
    if (!rc) { Logger.log('Tab "' + tab + '" not found in ' + id + '; skipped'); return; }
    const t = Date.now();

    if (!header) header = valuesBatch_(id, ["'" + tab + "'!A1:T1"], 'FORMATTED_VALUE')[0][0] || [];

    // 1. which rows? (ISO + type columns only)
    const cols = readColumns_(id, tab, ['AC', 'AR'], rc, 'UNFORMATTED_VALUE');
    const hits = [];
    for (let r = 0; r < rc; r++) {
      if (String(cols[1][r] || '').trim().toLowerCase() !== 'project') continue;
      if (iso && cleanIso(cols[0][r]) !== iso) continue;
      hits.push(r + 1);
    }

    // 2. fetch just those rows (nearby rows share one range)
    const runs = [];
    hits.forEach(r => {
      const last = runs[runs.length - 1];
      if (last && r - last[1] <= 20) last[1] = r; else runs.push([r, r]);
    });
    const rows = {}; // sheet row -> helper cells
    for (let i = 0; i < runs.length; i += 40) {
      const batch = runs.slice(i, i + 40);
      const m = valuesBatch_(id, batch.map(b => "'" + tab + "'!W" + b[0] + ':AR' + b[1]), 'UNFORMATTED_VALUE');
      m.forEach((vals, k) => vals.forEach((row, j) => { rows[batch[k][0] + j] = row; }));
    }

    // 3. Amount = round(rate * worked hours); rate = RESOURCES column M for the name in column B
    const rate = {};
    if (hits.length && sizes['RESOURCES']) {
      try {
        const rc2 = sizes['RESOURCES'];
        const rcols = readColumns_(id, 'RESOURCES', ['B', 'M'], rc2, 'UNFORMATTED_VALUE');
        for (let r = rc2 - 1; r >= 0; r--) {
          const k = String(rcols[0][r] || '').trim().toLowerCase();
          if (k) rate[k] = rcols[1][r];   // iterating backwards so the first match wins, like VLOOKUP
        }
      } catch (e) {
        Logger.log('RESOURCES rate lookup failed for ' + id + ': ' + e.message + ' (Amount left as 0)');
      }
    }

    let n = 0;
    hits.forEach(r => {
      const row = rows[r];
      if (!row) return;
      const out = [++n];
      for (let c = 1; c < map.length; c++) {
        const v = rd(row, map[c]);
        out.push(formatted ? (dateCols[c] ? fmtDate_(v) : fmtNum_(v)) : v);
      }
      const worked = parseFloat(rd(row, 'AN')) || 0;
      const rt = parseFloat(rate[String(rd(row, 'AM')).trim().toLowerCase()]) || 0;
      const amount = Math.round(rt * worked);
      out.push(formatted ? String(amount) : amount);
      out.push(idIdx, 0);
      data.push(out);
    });
    Logger.log('planning ' + id + ': ' + rc + ' rows scanned, ' + hits.length + ' matched in ' + (Date.now() - t) + ' ms');
  });

  return [(header || []).concat(['', ''])].concat(data);
}

/**
 * Run from the editor. For the project below, loads the Planning, Manpower Cost and WBS Sheet tabs from
 * the common sheet AND from the source spreadsheets and reports whether they agree (row counts,
 * identical rows, differing rows), so a tab is only switched to direct reads (remove it from
 * LOCAL_TABS) once it matches.
 */
function compareSources() {
  const projectName = 'Hyatt regency';
  const ctx = projectContext_(projectName);
  const norm = c => {
    if (typeof c === 'number') return (c > 20000 && c < 80000) ? String(Math.floor(c)) : String(Math.round(c * 1e6) / 1e6);
    return String(c).trim();
  };
  [['PROJECT PLANNING', 1], ['Manpower Cost', 0], ['WBS Sheet', 0], ['INVENTORY', 0]].forEach(pair => {
    const name = pair[0], skipCols = pair[1];
    let loc, dir;
    const t1 = Date.now();
    try { loc = loadSources_([name], ctx.ss, null, { iso: ctx.isoNo, source: 'local' })[name]; }
    catch (e) { Logger.log(name + ': common-sheet read FAILED: ' + e.message); return; }
    const t2 = Date.now();
    try { dir = loadSources_([name], ctx.ss, null, { iso: ctx.isoNo, source: 'direct' })[name]; }
    catch (e) { Logger.log(name + ': direct read FAILED: ' + e.message); return; }
    const t3 = Date.now();
    const hl = headerCount_(name, 'local'), hd = headerCount_(name, 'direct');
    const sig = (rows) => rows.map(r => r.slice(skipCols, r.length - META_COLS).map(norm).join('|'));
    const a = sig(loc.slice(hl)), b = sig(dir.slice(hd));
    const count = {};
    b.forEach(x => { count[x] = (count[x] || 0) + 1; });
    let same = 0; const onlyLocal = [];
    a.forEach((x, i) => { if (count[x] > 0) { count[x]--; same++; } else onlyLocal.push(x); });
    const onlyDirect = [];
    Object.keys(count).forEach(x => { for (let k = 0; k < count[x]; k++) onlyDirect.push(x); });
    Logger.log('== ' + name + ' for ' + ctx.isoNo + ': common ' + a.length + ' rows (' + (t2 - t1) + ' ms) | direct ' + b.length + ' rows (' + (t3 - t2) + ' ms) | identical ' + same);
    Logger.log('   header common: ' + JSON.stringify(loc[hl - 1].slice(0, loc[hl - 1].length - META_COLS)));
    Logger.log('   header direct: ' + JSON.stringify(dir[hd - 1].slice(0, dir[hd - 1].length - META_COLS)));
    onlyLocal.slice(0, 3).forEach(x => Logger.log('   only in common: ' + x.slice(0, 200)));
    onlyDirect.slice(0, 3).forEach(x => Logger.log('   only in direct: ' + x.slice(0, 200)));
    Logger.log('   RESULT: ' + (a.length === b.length && same === a.length ? 'MATCH - safe to switch' : 'DIFFERENT - do not switch yet'));
  });
}

/**
 * Run from the editor for one problem spreadsheet: shows which account the script runs
 * as, and what each access path (Sheets API, SpreadsheetApp, DriveApp) says about the ID.
 */
function probeAccess() {
  const id = PROJECT_SPREADSHEET_IDS[0];
  Logger.log('ID in code: [' + id + '] length ' + id.length);
  Logger.log('Effective user: ' + Session.getEffectiveUser().getEmail());
  try { Logger.log('Sheets API: ' + Sheets.Spreadsheets.get(id, { fields: 'properties.title' }).properties.title); }
  catch (e) { Logger.log('Sheets API failed: ' + e.message); }
  try { Logger.log('SpreadsheetApp: ' + SpreadsheetApp.openById(id).getName()); }
  catch (e) { Logger.log('SpreadsheetApp failed: ' + e.message); }
  try { Logger.log('DriveApp: ' + DriveApp.getFileById(id).getName()); }
  catch (e) { Logger.log('DriveApp failed: ' + e.message); }
}

/**
 * Cheap metadata-only call: logs the grid size of every tab we read, so you can see
 * which tab/spreadsheet is huge (rows x columns) before fetching any data.
 */
function diagnoseSources() {
  const ids = {};
  Object.keys(SOURCES).forEach(n => SOURCES[n].ids.forEach(id => { (ids[id] = ids[id] || {})[SOURCES[n].tab] = true; }));
  Object.keys(ids).forEach(id => {
    const t = Date.now();
    try {
      const meta = Sheets.Spreadsheets.get(id, { fields: 'properties.title,sheets.properties(title,gridProperties(rowCount,columnCount))' });
      const lines = meta.sheets
        .filter(s => ids[id][s.properties.title])
        .map(s => s.properties.title + ' ' + s.properties.gridProperties.rowCount + 'x' + s.properties.gridProperties.columnCount);
      Logger.log(id + ' "' + meta.properties.title + '" (' + (Date.now() - t) + ' ms): ' + lines.join(' | '));
    } catch (e) {
      Logger.log(id + ' FAILED after ' + (Date.now() - t) + ' ms: ' + e.message);
    }
  });
}

/**
 * Logs the header rows and row count of every source so you can check headerRows /
 * dataStart against what the code expects. Reads everything: can take a while.
 */
function verifySources() {
  const names = Object.keys(SOURCES);
  const data = loadSources_(names, SpreadsheetApp.getActiveSpreadsheet());
  names.forEach(n => {
    const hdr = headerCount_(n);
    Logger.log(n + ': ' + (data[n].length - hdr) + ' rows | header: ' + JSON.stringify(data[n].slice(0, hdr).map(r => r.slice(0, -META_COLS))));
  });
}

/**
 * Times the read of every tab we use from one project spreadsheet (default: the first,
 * STP CIVIL), one tab at a time, so a slow or failing tab is easy to spot.
 */
function profileBucket() {
  const id = PROJECT_SPREADSHEET_IDS[0];
  const seen = {};
  Object.keys(SOURCES).forEach(name => {
    const src = SOURCES[name];
    if (src.ids.indexOf(id) === -1) return;
    src.blocks.forEach(block => {
      const key = src.tab + '|' + block;
      if (seen[key]) return;
      seen[key] = true;
      const t = Date.now();
      try {
        const m = readSpreadsheetRanges_(id, [{ tab: src.tab, block: block }], 'UNFORMATTED_VALUE')[0];
        Logger.log(src.tab + ' ' + block + ': ' + m.length + ' rows in ' + (Date.now() - t) + ' ms');
      } catch (e) {
        Logger.log(src.tab + ' ' + block + ': FAILED after ' + (Date.now() - t) + ' ms: ' + e.message);
      }
    });
  });
}

/**
 * Reads only rows 2-3 (40 cells) of STP CIVIL's PROJECT PLANNING as formulas and logs them,
 * so you can see whether the tab holds typed values or formulas (IMPORTRANGE / QUERY /
 * ARRAYFORMULA...). If this is still running after about a minute, press Stop and tell me.
 */
function inspectPlanning() {
  const id = PROJECT_SPREADSHEET_IDS[0];
  const t = Date.now();
  const read = (range, render) => Sheets.Spreadsheets.Values.get(id, "'PROJECT PLANNING'!" + range, { valueRenderOption: render }).values || [];
  try {
    // One labelled line per column: letter | header | row 2 | row 3, formulas as typed.
    const f = read('A1:AZ3', 'FORMULA');
    const width = Math.max.apply(null, f.map(r => r.length));
    for (let c = 0; c < width; c++) {
      Logger.log(numToCol_(c + 1) + ' | hdr: ' + String((f[0] || [])[c] === undefined ? '' : f[0][c]).slice(0, 60) +
        ' | r2: ' + String((f[1] || [])[c] === undefined ? '' : f[1][c]).slice(0, 160) +
        ' | r3: ' + String((f[2] || [])[c] === undefined ? '' : f[2][c]).slice(0, 80));
    }
    Logger.log('inspectPlanning ok in ' + (Date.now() - t) + ' ms');
  } catch (e) {
    Logger.log('inspectPlanning FAILED after ' + (Date.now() - t) + ' ms: ' + e.message);
  }
  try {
    const m = Sheets.Spreadsheets.get(id, { fields: 'sheets.properties(title,gridProperties(rowCount,columnCount))' });
    const res = m.sheets.filter(s => s.properties.title === 'RESOURCES')[0];
    Logger.log('RESOURCES tab: ' + (res ? res.properties.gridProperties.rowCount + 'x' + res.properties.gridProperties.columnCount : 'not found'));
  } catch (e) {
    Logger.log('RESOURCES lookup failed: ' + e.message);
  }
}

/**
 * Finds out how the STP CIVIL PROJECT PLANNING tab can be read. Tries several range sizes
 * with each value render option through the Sheets API, then the same through
 * SpreadsheetApp, and logs time and rows or the error for each. Stops after 4 minutes.
 */
function probePlanning() {
  const id = PROJECT_SPREADSHEET_IDS[0];
  const tab = 'PROJECT PLANNING';
  const t0 = Date.now();
  const over = () => Date.now() - t0 > 240000;
  const ranges = ['A1:T5', 'A1:T200', 'A1:T1000', 'A1:T4000', 'U1:AR200'];
  ['UNFORMATTED_VALUE', 'FORMATTED_VALUE', 'FORMULA'].forEach(render => {
    ranges.forEach(r => {
      if (over()) return;
      const t = Date.now();
      try {
        const v = Sheets.Spreadsheets.Values.get(id, "'" + tab + "'!" + r, { valueRenderOption: render }).values || [];
        Logger.log('API ' + render + ' ' + r + ': ' + v.length + ' rows in ' + (Date.now() - t) + ' ms');
      } catch (e) {
        Logger.log('API ' + render + ' ' + r + ': FAILED after ' + (Date.now() - t) + ' ms: ' + e.message);
      }
    });
  });
  [200, 1000, 4000].forEach(n => {
    if (over()) return;
    const t = Date.now();
    try {
      const sh = SpreadsheetApp.openById(id).getSheetByName(tab);
      const v = sh.getRange(1, 1, n, 20).getValues();
      Logger.log('SpreadsheetApp A1:T' + n + ': ' + v.length + ' rows in ' + (Date.now() - t) + ' ms');
    } catch (e) {
      Logger.log('SpreadsheetApp A1:T' + n + ': FAILED after ' + (Date.now() - t) + ' ms: ' + e.message);
    }
  });
  Logger.log('probePlanning finished in ' + (Date.now() - t0) + ' ms');
}
