/**
 * Table / modal loaders for the dashboard. Each one finds the project's ISO and bucket
 * from the cached project index, loads only that ISO's rows from the source spreadsheet
 * (see Sources.js), and returns the same { headers, rows } shapes the page already expects.
 */

function projectContext_(projectName) {
  const p = projectIndex_().byName[projectName];
  if (!p) throw new Error("Project '" + projectName + "' not found.");
  if (!p.iso) throw new Error('ISO No. not found for this project.');
  return {
    ss: SpreadsheetApp.getActiveSpreadsheet(),
    projectName: projectName,
    isoNo: p.iso,
    projectIds: bucketSpreadsheetIds_(p.bucket)
  };
}

function contextFromIso_(isoNo) {
  const iso = cleanIso(isoNo);
  const p = projectIndex_().byIso[iso];
  if (!p) throw new Error('ISO No. ' + iso + ' not found.');
  return {
    ss: SpreadsheetApp.getActiveSpreadsheet(),
    projectName: p.name,
    isoNo: iso,
    projectIds: bucketSpreadsheetIds_(p.bucket)
  };
}

/** Rows of one tab for the project's ISO (header rows first). */
function projectTab_(ctx, name, formatted) {
  return loadSources_([name], ctx.ss, ctx.projectIds, { formatted: !!formatted, iso: ctx.isoNo })[name];
}

/** Index of the last header row of a tab (INVENTORY and the audit tabs have two). */
function headerIdx_(name) {
  return SOURCES[name] ? headerCount_(name) - 1 : 0;
}

function apiError_(fnName, e) {
  Logger.log('Error in ' + fnName + ': ' + e.toString());
  return { error: 'An error occurred on the server: ' + e.message };
}

/**
 * Generic table: one tab, one ISO, optional row predicate and link columns.
 *  opts.key    distinguishes cached variants (filters)
 *  opts.pred   (hdr) => (row) => boolean; may throw if a required column is missing
 *  opts.links  header names whose cells should be replaced by their hyperlink
 */
function buildTable_(ctx, name, opts) {
  opts = opts || {};
  return cachedResult_(['tbl', name, ctx.isoNo, opts.key || ''].join('|'), function () {
    const t = projectTab_(ctx, name, true);
    const hIdx = headerIdx_(name);
    if (!t.length) throw new Error("Required sheet '" + name + "' not found or empty.");
    const hdr = (t[hIdx] || []).map(h => String(h).trim());
    const ci = isoColIdx_(hdr, name);
    if (ci === -1) throw new Error("'ISO No.' column not found in '" + name + "'.");

    let rows = t.slice(hIdx + 1).filter(r => cleanIso(r[ci]) === ctx.isoNo);
    if (opts.pred) rows = rows.filter(opts.pred(hdr));
    if (opts.links) resolveLinks_(name, hdr, rows, opts.links);

    const width = hdr.length - (SOURCES[name] ? META_COLS : 0);
    return { headers: hdr.slice(0, width), rows: rows.map(r => r.slice(0, width)) };
  });
}

function colIdxOrThrow_(hdr, colName, tabName) {
  const i = hdr.indexOf(colName);
  if (i === -1) throw new Error("Required column '" + colName + "' not found in '" + tabName + "'.");
  return i;
}

function positive_(v) {
  return parseFloat(String(v).replace(/,/g, '')) > 0;
}

function statusOf_(v) {
  return (v || '').toString().toUpperCase().trim();
}

// ---------------------------------------------------------------------------
// Plain tables
// ---------------------------------------------------------------------------

function getWbsDataForProject(projectName) {
  try { return buildTable_(projectContext_(projectName), 'WBS Sheet'); }
  catch (e) { return apiError_('getWbsDataForProject', e); }
}

function getProjectPlanningData(projectName) {
  try { return buildTable_(projectContext_(projectName), 'PROJECT PLANNING'); }
  catch (e) { return apiError_('getProjectPlanningData', e); }
}

function getInventoryData(projectName) {
  try { return buildTable_(projectContext_(projectName), 'INVENTORY', { links: ['PO NO'] }); }
  catch (e) { return apiError_('getInventoryData', e); }
}

function getManpowerCostData(projectName) {
  try { return buildTable_(projectContext_(projectName), 'Manpower Cost'); }
  catch (e) { return apiError_('getManpowerCostData', e); }
}

function getFinanceData(projectName) {
  try { return buildTable_(projectContext_(projectName), 'FINANCE/ACCOUNTS'); }
  catch (e) { return apiError_('getFinanceData', e); }
}

function getAuditManpowerData(projectName) {
  try { return buildTable_(projectContext_(projectName), 'Auditor Sheet Manpower'); }
  catch (e) { return apiError_('getAuditManpowerData', e); }
}

function getAuditItemData(projectName) {
  try { return buildTable_(projectContext_(projectName), 'Auditor Sheet Item'); }
  catch (e) { return apiError_('getAuditItemData', e); }
}

function getAuditDrawingData(projectName) {
  try { return buildTable_(projectContext_(projectName), 'Auditor Sheet Drawing', { links: ['Drawing Attachment', 'Final Drawing'] }); }
  catch (e) { return apiError_('getAuditDrawingData', e); }
}

function getExpenseDetailsData(projectName) {
  try { return buildTable_(projectContext_(projectName), 'Expenses'); }
  catch (e) { return apiError_('getExpenseDetailsData', e); }
}

function getBillingData(projectName) {
  try { return buildTable_(projectContext_(projectName), 'Billing', { links: ['Supp Doc'] }); }
  catch (e) { return apiError_('getBillingData', e); }
}

// ---------------------------------------------------------------------------
// Filtered tables
// ---------------------------------------------------------------------------

/** Product Used (Rs.) modal: INVENTORY rows with CONSU.2 > 0. */
function getProductUsedData(projectName) {
  try {
    return buildTable_(projectContext_(projectName), 'INVENTORY', {
      key: 'consu2',
      pred: hdr => { const i = hdr.indexOf('CONSU.2'); return row => i !== -1 && parseFloat(row[i]) > 0; }
    });
  } catch (e) { return apiError_('getProductUsedData', e); }
}

/** INVENTORY rows where the given column (e.g. 'INDENT1', 'REC.1') is > 0. */
function getFilteredInventoryData(projectName, filterColumn) {
  try {
    return buildTable_(projectContext_(projectName), 'INVENTORY', {
      key: 'col:' + filterColumn,
      pred: hdr => { const i = colIdxOrThrow_(hdr, filterColumn, 'INVENTORY'); return row => positive_(row[i]); }
    });
  } catch (e) { return apiError_('getFilteredInventoryData', e); }
}

/** Tasks by status. HOLD includes blanks, to match the KPI math. */
function getFilteredTasks(projectName, filterType) {
  try {
    return buildTable_(projectContext_(projectName), 'PROJECT PLANNING', {
      key: 'status:' + filterType,
      pred: hdr => {
        const s = colIdxOrThrow_(hdr, 'Status', 'PROJECT PLANNING');
        if (['DONE', 'IN-PROGRESS', 'HOLD'].indexOf(filterType) === -1) return () => false;
        return row => {
          const v = statusOf_(row[s]);
          if (filterType === 'HOLD') return v !== 'DONE' && v !== 'IN-PROGRESS' && v !== 'IN PROGRESS';
          if (filterType === 'IN-PROGRESS') return v === 'IN-PROGRESS' || v === 'IN PROGRESS';
          return v === filterType;
        };
      }
    });
  } catch (e) { return apiError_('getFilteredTasks', e); }
}

function getPlannedHrsData(projectName) {
  try {
    return buildTable_(projectContext_(projectName), 'PROJECT PLANNING', {
      key: 'planned',
      pred: hdr => { const i = colIdxOrThrow_(hdr, 'Required Hours', 'PROJECT PLANNING'); return row => parseFloat(row[i]) > 0; }
    });
  } catch (e) { return apiError_('getPlannedHrsData', e); }
}

function getActualHrsData(projectName) {
  try {
    return buildTable_(projectContext_(projectName), 'PROJECT PLANNING', {
      key: 'actual',
      pred: hdr => { const i = colIdxOrThrow_(hdr, 'Worked (Hrs)', 'PROJECT PLANNING'); return row => parseFloat(row[i]) > 0; }
    });
  } catch (e) { return apiError_('getActualHrsData', e); }
}

function getFilteredBillingData(projectName, filterType) {
  try {
    return buildTable_(projectContext_(projectName), 'Billing', {
      key: 'status:' + filterType,
      pred: hdr => {
        const s = colIdxOrThrow_(hdr, 'Status5', 'Billing');
        if (filterType === 'DONE') return row => statusOf_(row[s]) === 'DONE';
        if (filterType === 'PENDING') return row => statusOf_(row[s]) !== 'DONE';
        return () => false;
      }
    });
  } catch (e) { return apiError_('getFilteredBillingData', e); }
}

// ---------------------------------------------------------------------------
// Project details (single row, with links) and MOM
// ---------------------------------------------------------------------------

function getProjectDetailsFormData(projectName) {
  try {
    const ctx = projectContext_(projectName);
    return cachedResult_(['details', ctx.isoNo, projectName].join('|'), function () {
      const name = 'Project Details Form';
      const t = projectTab_(ctx, name, true);
      if (!t.length) throw new Error("'Project Details Form' returned no rows.");
      const hdr = t[0].map(h => String(h).trim());
      const nameIdx = hdr.indexOf('PROJECT NAME');
      if (nameIdx === -1) throw new Error("'PROJECT NAME' column not found.");
      const row = t.slice(1).find(r => r[nameIdx] === projectName);
      if (!row) throw new Error("Project '" + projectName + "' not found in 'Project Details Form'.");

      resolveLinks_(name, hdr, [row], ['SITE PICS', 'INFO-SHEET', 'PROJECT DRAWING', 'CUSTOMER PO COPY', 'CLIENT SITE VISITS/MOM']);
      const width = hdr.length - META_COLS;
      return { headers: hdr.slice(0, width), row: row.slice(0, width) };
    });
  } catch (e) { return apiError_('getProjectDetailsFormData', e); }
}

/** MOM rows for an ISO with image / PDF / signature cells replaced by their links. */
function momTable_(isoNo) {
  const iso = cleanIso(isoNo);
  return cachedResult_(['mom', iso].join('|'), function () {
    const t = loadSources_(['MOM'], SpreadsheetApp.getActiveSpreadsheet(), null, { formatted: true, iso: iso })['MOM'];
    if (!t.length) return { headers: [], rows: [] };
    const hdr = t[0].map(h => String(h).trim());
    const ci = isoColIdx_(hdr, 'MOM');
    if (ci === -1) throw new Error("'ISO No.' column not found in 'MOM' sheet.");
    const rows = t.slice(1).filter(r => cleanIso(r[ci]) === iso);
    resolveLinks_('MOM', hdr, rows, [
      'Upload Image 1', 'Upload Image 2', 'Upload Image 3', 'Upload Image 4', 'Upload Image 5',
      'PDF Link', 'Client Signature (Client Attendee 1)', 'Mentor Signature (Mentor Attendee 1)'
    ]);
    const width = hdr.length - META_COLS;
    return { headers: hdr.slice(0, width), rows: rows.map(r => r.slice(0, width)) };
  });
}

function getMomData(isoNo) {
  try {
    if (!isoNo) return { error: 'ISO No. was not provided to the server.' };
    return momTable_(isoNo);
  } catch (e) { return apiError_('getMomData', e); }
}

// ---------------------------------------------------------------------------
// Summaries computed from PROJECT PLANNING / WBS Sheet
// ---------------------------------------------------------------------------

/** Per-engineer total worked hours, highest first. */
function getEngineerWorkSummary(projectName) {
  try {
    const ctx = projectContext_(projectName);
    const t = projectTab_(ctx, 'PROJECT PLANNING', false);
    const hdr = (t[0] || []).map(h => String(h).trim());
    const iso = isoColIdx_(hdr);
    const doer = hdr.indexOf('Engineer Name (Doer)');
    const worked = hdr.indexOf('Worked (Hrs)');
    if (iso === -1 || doer === -1 || worked === -1) return { error: "Required columns not found in 'PROJECT PLANNING' sheet." };

    const hours = {};
    t.slice(1).filter(r => cleanIso(r[iso]) === ctx.isoNo).forEach(r => {
      const h = parseFloat(r[worked]) || 0;
      if (r[doer] && h > 0) hours[r[doer]] = (hours[r[doer]] || 0) + h;
    });
    return Object.keys(hours).map(n => ({ name: n, hours: hours[n] })).sort((a, b) => b.hours - a.hours);
  } catch (e) { return apiError_('getEngineerWorkSummary', e); }
}

/** Task counts per milestone (pending = in-progress + hold). */
function getMilestoneDetailsData(projectName) {
  try {
    const ctx = projectContext_(projectName);
    const t = projectTab_(ctx, 'PROJECT PLANNING', false);
    const hdr = (t[0] || []).map(h => String(h).trim());
    const iso = isoColIdx_(hdr);
    const mile = hdr.indexOf('Milestone');
    const status = hdr.indexOf('Status');
    if (iso === -1 || mile === -1 || status === -1) return { error: "Required columns not found in 'PROJECT PLANNING' sheet." };

    const stats = {};
    t.slice(1).filter(r => cleanIso(r[iso]) === ctx.isoNo).forEach(r => {
      const m = r[mile];
      if (!m) return;
      const s = stats[m] = stats[m] || { done: 0, inProgress: 0, hold: 0 };
      const v = statusOf_(r[status]);
      if (v === 'DONE') s.done++;
      else if (v === 'IN-PROGRESS') s.inProgress++;
      else if (v === 'HOLD') s.hold++;
    });
    return Object.keys(stats).map(m => {
      const s = stats[m];
      return { milestone: m, total: s.done + s.inProgress + s.hold, done: s.done, pending: s.inProgress + s.hold };
    });
  } catch (e) { return apiError_('getMilestoneDetailsData', e); }
}

/** Milestone names that are COMPLETED (every WBS row has an actual end date) or PENDING. */
function getFilteredMilestones(projectName, filterType) {
  try {
    const ctx = projectContext_(projectName);
    const t = projectTab_(ctx, 'WBS Sheet', false);
    const hdr = (t[0] || []).map(h => String(h).trim());
    const iso = isoColIdx_(hdr);
    const mile = hdr.indexOf('MILESTONE/BOQ ITEM');
    const end = hdr.indexOf('ACTUAL END DATE');
    if (iso === -1 || mile === -1 || end === -1) return { error: "Required columns not found in 'WBS Sheet'." };

    const groups = {};
    t.slice(1).filter(r => cleanIso(r[iso]) === ctx.isoNo).forEach(r => {
      const m = String(r[mile]).trim();
      if (!m || m === 'N/A') return;
      (groups[m] = groups[m] || []).push(r);
    });
    return Object.keys(groups).filter(m => {
      const complete = groups[m].every(r => r[end] && String(r[end]).trim() !== '' && String(r[end]).toUpperCase() !== 'N/A');
      return filterType === 'COMPLETED' ? complete : filterType === 'PENDING' ? !complete : false;
    }).sort();
  } catch (e) { return apiError_('getFilteredMilestones', e); }
}

// ---------------------------------------------------------------------------
// Legacy chart endpoints that take an ISO directly (not called by the current page)
// ---------------------------------------------------------------------------

function getTaskStatusChartData(isoNo) {
  try {
    const ctx = contextFromIso_(isoNo);
    const t = projectTab_(ctx, 'PROJECT PLANNING', false);
    const hdr = (t[0] || []).map(h => String(h).trim());
    const iso = isoColIdx_(hdr);
    const task = hdr.indexOf('Task');
    const pct = hdr.indexOf('Status %');
    const planned = hdr.indexOf('Planned Date');
    if (iso === -1 || task === -1 || pct === -1) return [];

    const chartData = t.slice(1)
      .filter(r => cleanIso(r[iso]) === ctx.isoNo && r[task])
      .map(r => {
        let d = planned > -1 ? toDate_(r[planned]) : null;
        if (!d && planned > -1 && typeof r[planned] === 'string') {
          const m = r[planned].match(/(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
          if (m) d = new Date(m[3], m[2] - 1, m[1]);
        }
        const valid = d && !isNaN(d);
        return {
          milestone: r[task],
          planned: parseFloat(r[pct]) || 0,
          actual: valid ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd-MM-yyyy') : 'N/A',
          sortKey: valid ? d.getTime() : null
        };
      });
    chartData.sort((a, b) => a.sortKey === null ? 1 : b.sortKey === null ? -1 : a.sortKey - b.sortKey);
    return chartData.map(c => ({ milestone: c.milestone, planned: c.planned, actual: c.actual }));
  } catch (e) {
    Logger.log('Error in getTaskStatusChartData: ' + e.toString());
    return [];
  }
}

function getMilestoneComboChartData(isoNo) {
  try {
    const ctx = contextFromIso_(isoNo);
    const planning = projectTab_(ctx, 'PROJECT PLANNING', false);
    const billing = projectTab_(ctx, 'Billing', false);
    return calculateMilestoneChartInMemory(ctx.isoNo, planning, billing);
  } catch (e) {
    Logger.log('Error in getMilestoneComboChartData: ' + e.toString());
    return [];
  }
}
