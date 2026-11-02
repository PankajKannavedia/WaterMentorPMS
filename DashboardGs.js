/**
 * Progressive dashboard loading. The page calls getDashboardSection() once per section,
 * all at the same time, and fills each card as its section returns, so a slow source
 * (Expenses, PROJECT PLANNING) no longer blocks the cards that do not need it.
 * Each section loads only the tabs it needs, for one ISO (see Sources.js).
 */

const DASH_TABS = [
  'Project Details Form', 'PROJECT PLANNING', 'Manpower Cost', 'INVENTORY',
  'Auditor Sheet Manpower', 'Auditor Sheet Item', 'Auditor Sheet Drawing',
  'FINANCE/ACCOUNTS', 'Billing', 'WBS Sheet', 'Expenses'
];

/** tabs: sources to load (Project Details is always loaded). keys: KPI values this section owns. */
const DASH_SECTIONS = {
  header:    { tabs: [], keys: ['orderAmount', 'projectManager', 'siteIncharge', 'paymentTerms', 'poCopy'] },
  expenses:  { tabs: ['Expenses'], keys: ['expensesRs'] },
  inventory: { tabs: ['INVENTORY', 'Manpower Cost'], keys: ['manpowerUsedRs', 'productUsedRs', 'indentPlanned', 'actualIndent', 'received', 'used'] },
  planning:  { tabs: ['PROJECT PLANNING', 'Billing'], keys: ['engWorked', 'plannedManHrs', 'actualManHrs', 'totalTasks', 'tasksDone', 'tasksInProgress', 'tasksOnHold'] },
  audit:     { tabs: ['Auditor Sheet Manpower', 'Auditor Sheet Item', 'Auditor Sheet Drawing'], keys: ['auditManpower', 'auditProduct', 'auditDrawings'] },
  wbs:       { tabs: ['WBS Sheet'], keys: ['totalMilestones', 'completedMilestones', 'pendingMilestones', 'milestoneDonePercent'] },
  billing:   { tabs: ['Billing', 'FINANCE/ACCOUNTS'], keys: ['totalBillingMilestones', 'completedBillingMilestones', 'overallBillingDonePercent', 'overallBillingPendingPercent', 'paymentReceivedPercent'] },
  mom:       { tabs: [], keys: [] }
};

/**
 * Run from the editor (optionally pass a project name, default "Hyatt regency"). For each
 * source, logs how many rows matched the project's ISO, how long it took, and the header
 * row, so a wrong header/start row or an ISO mismatch is easy to see. A failing source is
 * logged and skipped; the rest still run.
 */
function debugProject(projectName) {
  projectName = projectName || 'Hyatt regency';
  const idx = projectIndex_();
  const p = idx.byName[projectName];
  Logger.log('index: ' + JSON.stringify(p) + ' | projects indexed: ' + Object.keys(idx.byName).length);
  if (!p) return;
  const ctx = projectContext_(projectName);
  Logger.log('bucket "' + p.bucket + '" -> spreadsheets: ' + JSON.stringify(ctx.projectIds) + ' | ISO: ' + ctx.isoNo);
  // PROJECT PLANNING (the tab that has hung) goes last so every other source reports first.
  const order = DASH_TABS.filter(n => n !== 'PROJECT PLANNING').concat(['MOM', 'PROJECT PLANNING']);
  order.forEach(name => {
    const t = Date.now();
    try {
      const d = loadSources_([name], ctx.ss, ctx.projectIds, { iso: ctx.isoNo })[name];
      const h = headerIdx_(name);
      const hdr = (d[h] || []).map(x => String(x).trim());
      const shown = name === 'Project Details Form' ? hdr : hdr.slice(0, 16);
      Logger.log(name + ': ' + Math.max(0, d.length - h - 1) + ' rows for this ISO in ' + (Date.now() - t) + ' ms | header: ' + JSON.stringify(shown));
    } catch (e) {
      Logger.log(name + ': FAILED after ' + (Date.now() - t) + ' ms: ' + e.message);
    }
  });
}

/**
 * READ-ONLY. Run from the editor of the script bound to the common sheet. For every tab of
 * this spreadsheet it logs size, whether the app code still reads it, and which other tabs'
 * formulas mention it (first 60 rows of each tab are scanned), then lists the tabs that are
 * safe to hide: not read by the code and not needed (directly or through a chain of formulas)
 * by a tab that is. Nothing is changed. Formulas can also use INDIRECT or named ranges, which
 * this cannot see, so hide first (hideUnusedTabs) and delete only after a few days.
 */
function auditCommonSheet() {
  const plan = commonSheetPlan_();
  plan.tabs.forEach(t => {
    Logger.log((t.needed ? 'KEEP   ' : 'UNUSED ') + t.name + ' (' + t.rows + 'x' + t.cols + (t.hidden ? ', hidden' : '') + ')' +
      (t.inUse ? ' | code reads it: ' + t.inUse : '') +
      (t.referencedBy.length ? ' | formulas in: ' + t.referencedBy.join(', ') : '') +
      (!t.inUse && t.needed ? ' | kept only because a needed tab refers to it' : ''));
  });
  Logger.log('Safe to hide (' + plan.unused.length + '): ' + JSON.stringify(plan.unused));
}

/** Hides the unused tabs found by auditCommonSheet. Reversible: showAllTabs() unhides every tab. */
function hideUnusedTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const plan = commonSheetPlan_();
  let hidden = 0;
  plan.unused.forEach(name => {
    const sh = ss.getSheetByName(name);
    const visible = ss.getSheets().filter(s => !s.isSheetHidden()).length;
    if (sh && !sh.isSheetHidden() && visible > 1) { sh.hideSheet(); hidden++; }
  });
  Logger.log('Hid ' + hidden + ' tab(s): ' + JSON.stringify(plan.unused));
}

function showAllTabs() {
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(sh => { if (sh.isSheetHidden()) sh.showSheet(); });
  Logger.log('All tabs visible again.');
}

function commonSheetPlan_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const names = sheets.map(s => s.getName());

  // Tabs the server code still reads from this spreadsheet.
  const inUse = { 'Doer list': 'login', 'Project Details Form': 'Product / Site incharge columns and lookups', 'FINANCE/ACCOUNTS': 'payment received %' };
  Object.keys(SOURCES).forEach(n => { if (!isDirect_(n)) inUse[n] = 'dashboard data (still read from this sheet)'; });

  // Which tabs' formulas mention which tab.
  const referencedBy = {};
  names.forEach(n => { referencedBy[n] = []; });
  sheets.forEach(sh => {
    const rows = Math.min(sh.getLastRow(), 60), cols = sh.getLastColumn();
    if (!rows || !cols) return;
    const text = sh.getRange(1, 1, rows, cols).getFormulas().map(r => r.join('\n')).join('\n');
    names.forEach(n => {
      if (n !== sh.getName() && (text.indexOf(n + '!') > -1 || text.indexOf("'" + n + "'!") > -1)) referencedBy[n].push(sh.getName());
    });
  });

  // needed = read by the code, or referenced by a needed tab (repeat until nothing changes).
  const needed = {};
  names.forEach(n => { if (inUse[n]) needed[n] = true; });
  let changed = true;
  while (changed) {
    changed = false;
    names.forEach(n => {
      if (!needed[n] && referencedBy[n].some(y => needed[y])) { needed[n] = true; changed = true; }
    });
  }

  const tabs = sheets.map(sh => ({
    name: sh.getName(), rows: sh.getLastRow(), cols: sh.getLastColumn(), hidden: sh.isSheetHidden(),
    inUse: inUse[sh.getName()] || '', referencedBy: referencedBy[sh.getName()], needed: !!needed[sh.getName()]
  }));
  return { tabs: tabs, unused: tabs.filter(t => !t.needed).map(t => t.name) };
}

/**
 * Run from the editor of the script bound to the COMMON sheet. Prints, for every tab the
 * dashboard reads, the first three rows exactly as the old IMPORTRANGE-based tabs had them
 * (row 1-2 are often typed headers that do not exist in the source spreadsheets), plus the
 * formula in A1:A3 so the import's start row is visible. Copy the log back as text.
 */
function dumpCommonHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tabs = ['PROJECT PLANNING', 'Manpower Cost', 'INVENTORY', 'Auditor Sheet Manpower', 'Auditor Sheet Item',
    'Auditor Sheet Drawing', 'WBS Sheet', 'Billing', 'Project Details Form', 'Expenses', 'MOM'];
  tabs.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) { Logger.log('== ' + name + ': tab not found'); return; }
    const cols = Math.min(sh.getLastColumn(), 45);
    const vals = sh.getRange(1, 1, 3, cols).getDisplayValues();
    const f = sh.getRange(1, 1, 3, 1).getFormulas().map(r => String(r[0]).slice(0, 110));
    Logger.log('== ' + name + ' (' + sh.getLastRow() + ' rows x ' + sh.getLastColumn() + ' cols)');
    Logger.log('A1:A3 formulas: ' + JSON.stringify(f));
    vals.forEach((row, i) => Logger.log('row ' + (i + 1) + ': ' + JSON.stringify(row.map(c => String(c).slice(0, 40)))));
  });
}

/**
 * Run from the editor: calls each table loader the page uses, one at a time, with a real
 * project name (the editor cannot pass arguments), and logs rows / time / error for each.
 * Edit the name below to test another project.
 */
function testTables() {
  const projectName = 'Hyatt regency';
  const calls = [
    ['getProjectDetailsFormData', () => getProjectDetailsFormData(projectName)],
    ['getExpenseDetailsData', () => getExpenseDetailsData(projectName)],
    ['getBillingData', () => getBillingData(projectName)],
    ['getFinanceData', () => getFinanceData(projectName)],
    ['getManpowerCostData', () => getManpowerCostData(projectName)],
    ['getInventoryData', () => getInventoryData(projectName)],
    ['getAuditManpowerData', () => getAuditManpowerData(projectName)],
    ['getAuditItemData', () => getAuditItemData(projectName)],
    ['getAuditDrawingData', () => getAuditDrawingData(projectName)],
    ['getWbsDataForProject', () => getWbsDataForProject(projectName)],
    ['getProjectPlanningData', () => getProjectPlanningData(projectName)],
    ['getMomData', () => getMomData(projectIndex_().byName[projectName].iso)]
  ];
  calls.forEach(c => {
    const t = Date.now();
    const r = c[1]();
    const n = r && r.rows ? r.rows.length : r && r.row ? 1 : 0;
    Logger.log(c[0] + ': ' + (r && r.error ? 'ERROR ' + r.error : n + ' rows') + ' in ' + (Date.now() - t) + ' ms');
  });
}

/**
 * @param {string} projectName
 * @param {string} section one of DASH_SECTIONS
 * @return {{details?: Object, kpis?: Object, timelineChart?: Array, milestoneChart?: Array, momData?: Object, error?: string}}
 */
function getDashboardSection(projectName, section) {
  try {
    const def = DASH_SECTIONS[section];
    if (!def) return { error: 'Unknown dashboard section: ' + section };
    const t0 = Date.now();
    const ctx = projectContext_(projectName);

    if (section === 'mom') return { momData: momTable_(ctx.isoNo) };

    const loaded = loadSources_(['Project Details Form'].concat(def.tabs), ctx.ss, ctx.projectIds, { iso: ctx.isoNo });
    // calculateKpisInMemory tolerates empty tabs, so every tab not loaded here is just [].
    const allData = {};
    DASH_TABS.forEach(n => { allData[n] = loaded[n] || []; });

    const details = allData['Project Details Form'];
    if (details.length < 2) return { error: 'Project Details Form has no row for this project.' };

    const out = {};
    if (section === 'header') {
      const h = details[0].map(x => String(x).trim());
      const row = details.slice(1).find(r => r[h.indexOf('PROJECT NAME')] === projectName);
      if (!row) return { error: 'Project not found.' };
      out.details = {
        isoNo: ctx.isoNo,
        projectAdd: row[h.indexOf('PROJECT ADD')],
        contactPerson: row[h.indexOf('CONTACT PERSON')],
        item: row[h.indexOf('item')]
      };
    }

    const kpis = calculateKpisInMemory(ctx.isoNo, allData);
    out.kpis = {};
    def.keys.forEach(k => { out.kpis[k] = kpis[k] === undefined ? null : kpis[k]; });

    // Cards use the same reads as their detail modals (formatted INVENTORY, PROJECT PLANNING milestones).
    if (section === 'inventory') {
      const inv = projectTab_(ctx, 'INVENTORY', true);
      const hIdx = headerIdx_('INVENTORY');
      const hdr = (inv[hIdx] || []).map(h => String(h).trim());
      const ci = isoColIdx_(hdr, 'INVENTORY');
      const rows = ci === -1 ? [] : inv.slice(hIdx + 1).filter(r => cleanIso(r[ci]) === ctx.isoNo);
      const sum = col => {
        const i = hdr.indexOf(col);
        return i === -1 ? 0 : rows.reduce((s, r) => s + (parseFloat(String(r[i]).replace(/,/g, '')) || 0), 0);
      };
      out.kpis.indentPlanned = sum('Planned Indent Qty');
      out.kpis.actualIndent = sum('INDENT1');
      out.kpis.received = sum('REC.1');
      out.kpis.used = sum('CONSU.1');
      Logger.log('inventory KPI: rows=' + rows.length + ' isoCol=' + ci + ' hdr=' + JSON.stringify(hdr.slice(0, hdr.length - META_COLS)) +
        ' sums=' + [out.kpis.indentPlanned, out.kpis.actualIndent, out.kpis.received, out.kpis.used].join('/'));
    }
    if (section === 'wbs' && !out.kpis.totalMilestones) {
      const t = projectTab_(ctx, 'PROJECT PLANNING', false);
      const hdr = (t[0] || []).map(h => String(h).trim());
      const iso = isoColIdx_(hdr), mile = hdr.indexOf('Milestone'), st = hdr.indexOf('Status');
      if (iso > -1 && mile > -1 && st > -1) {
        const m = {};
        t.slice(1).filter(r => cleanIso(r[iso]) === ctx.isoNo && r[mile]).forEach(r => {
          const g = m[r[mile]] = m[r[mile]] || { all: true };
          if (statusOf_(r[st]) !== 'DONE') g.all = false;
        });
        const names = Object.keys(m), done = names.filter(n => m[n].all).length;
        out.kpis.totalMilestones = names.length;
        out.kpis.completedMilestones = done;
        out.kpis.pendingMilestones = names.length - done;
        out.kpis.milestoneDonePercent = names.length ? Math.round(done / names.length * 100) : 0;
      }
    }

    if (section === 'wbs') out.timelineChart = calculateTimelineChartInMemory(ctx.isoNo, allData['WBS Sheet']);
    if (section === 'planning') out.milestoneChart = calculateMilestoneChartInMemory(ctx.isoNo, allData['PROJECT PLANNING'], allData['Billing']);

    Logger.log('getDashboardSection ' + section + ' for ' + ctx.isoNo + ': ' + (Date.now() - t0) + ' ms');
    return out;
  } catch (e) {
    Logger.log('Error in getDashboardSection(' + section + '): ' + e.toString());
    return { error: e.message };
  }
}
