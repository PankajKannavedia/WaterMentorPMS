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

    if (section === 'wbs') out.timelineChart = calculateTimelineChartInMemory(ctx.isoNo, allData['WBS Sheet']);
    if (section === 'planning') out.milestoneChart = calculateMilestoneChartInMemory(ctx.isoNo, allData['PROJECT PLANNING'], allData['Billing']);

    Logger.log('getDashboardSection ' + section + ' for ' + ctx.isoNo + ': ' + (Date.now() - t0) + ' ms');
    return out;
  } catch (e) {
    Logger.log('Error in getDashboardSection(' + section + '): ' + e.toString());
    return { error: e.message };
  }
}
