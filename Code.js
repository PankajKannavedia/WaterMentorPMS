/**
 * Serves the web app and passes the app's URL to the HTML template.
 * This allows for safe redirection on login/logout.
 */
function doGet(e) {
  const userCache = CacheService.getUserCache();
  const isLoggedIn = userCache.get('isLoggedIn'); // Check the cache
  const scriptUrl = ScriptApp.getService().getUrl();

  // Clear any stale login state if needed
  Logger.log('Login status: ' + isLoggedIn);
  Logger.log('Script URL: ' + scriptUrl);

  if (isLoggedIn === 'true') {
    const template = HtmlService.createTemplateFromFile('Index');
    template.url = scriptUrl;
    return template.evaluate()
      .setTitle('PMS')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else {
    const template = HtmlService.createTemplateFromFile('Login');
    template.url = scriptUrl;
    return template.evaluate()
      .setTitle('PMS Login');
  }
}

/**
 * Authenticates a user against the 'Doer list' sheet.
 */
function authenticateUser(username, password) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Doer list');
    if (!sheet) return { success: false, message: 'Configuration error: "Doer list" sheet not found.' };

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();

    const nameIndex = headers.indexOf('Name');
    const passwordIndex = headers.indexOf('Password');
    const roleIndex = headers.indexOf('Role');
    // --- UPDATED: Look for ISO column ---
    const isoIndex = headers.indexOf('ISO No.'); 

    const userRow = data.find(row => row[nameIndex].toString().trim() === username.trim());

    if (userRow && userRow[passwordIndex].toString() === password) {
      const userCache = CacheService.getUserCache();
      userCache.put('isLoggedIn', 'true', 14400); 

      const userProperties = PropertiesService.getUserProperties();
      userProperties.setProperty('username', userRow[nameIndex].toString());
      
      if (roleIndex !== -1) {
        userProperties.setProperty('userRole', userRow[roleIndex].toString());
      }
      
      // --- UPDATED: Store the ISO Number ---
      if (isoIndex !== -1 && userRow[isoIndex]) {
        // userProperties.setProperty('assignedIso', userRow[isoIndex].toString());
        
        // Clean the ISO before storing it in the user session
        userProperties.setProperty('assignedIso', cleanIso(userRow[isoIndex]));
      } else {
        userProperties.deleteProperty('assignedIso');
      }

      return { success: true, message: 'Login successful.' };
    } else {
      return { success: false, message: 'Invalid username or password.' };
    }
  } catch (e) {
    Logger.log(`Authentication Error: ${e.toString()}`);
    return { success: false, message: 'An error occurred during authentication.' };
  }
}

/**
 * Retrieves the current user's information.
 */
function getUserInfo() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    return { 
      username: userProperties.getProperty('username'), 
      userRole: userProperties.getProperty('userRole') || 'N/A',
      assignedIso: userProperties.getProperty('assignedIso') || null // Return the ISO
    };
  } catch (e) {
    return null;
  }
}

/**
 * Fetches the assigned project for a specific user based on the 'Doer list'.
 * 1. Finds the ISO No. for the username in 'Doer list'.
 * 2. Finds the Project Name for that ISO No. in 'Project Details Form'.
 * @param {string} username The logged-in username.
 * @returns {object} { projectName: "...", isoNo: "..." } or null if not found.
 */
function getCustomerAssignedProject(username) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Get ISO No. from 'Doer list'
    const doerSheet = ss.getSheetByName('Doer list');
    if (!doerSheet) return null;
    
    const doerData = doerSheet.getDataRange().getValues();
    const doerHeaders = doerData.shift(); // Headers: Name, Password, Role, ISO No.
    
    const nameIdx = doerHeaders.indexOf('Name');
    const isoIdx = doerHeaders.indexOf('ISO No.');
    
    if (nameIdx === -1 || isoIdx === -1) return null;
    
    const doerRow = doerData.find(row => row[nameIdx].toString().trim() === username);
    if (!doerRow) return null;
    
    const assignedIso = doerRow[isoIdx];
    if (!assignedIso) return null;

    // 2. Get Project Name from 'Project Details Form' using that ISO No.
    const detailsSheet = ss.getSheetByName('Project Details Form');
    if (!detailsSheet) return null;
    
    const detailsData = detailsSheet.getDataRange().getValues();
    const detailsHeaders = detailsData.shift();
    const projNameIdx = detailsHeaders.indexOf('PROJECT NAME');
    const detailsIsoIdx = detailsHeaders.indexOf('ISO No.');
    
    if (projNameIdx === -1 || detailsIsoIdx === -1) return null;
    
    const projectRow = detailsData.find(row => row[detailsIsoIdx] === assignedIso);
    
    if (projectRow) {
      return {
        projectName: projectRow[projNameIdx],
        isoNo: assignedIso
      };
    }
    
    return null;

  } catch (e) {
    Logger.log('Error in getCustomerAssignedProject: ' + e.toString());
    return null;
  }
}

/**
 * NEW HELPER: Finds the Project Name based on the ISO No.
 * This is needed because your Dashboard functions require Project Name to load data.
 */
function getProjectNameFromIso(isoNo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Project Details Form');
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  
  const isoColIdx = headers.indexOf('ISO No.');
  const projectColIdx = headers.indexOf('PROJECT NAME');

  if (isoColIdx === -1 || projectColIdx === -1) return null;

  const projectRow = data.find(row => row[isoColIdx] === isoNo);
  
  // Return the Project Name if found
  return projectRow ? projectRow[projectColIdx] : null; 
}


/**
 * Logs the user out by deleting their session properties.
 */
function logout() {
  try {
    const userCache = CacheService.getUserCache();
    userCache.remove('isLoggedIn'); // Remove the session from the cache

    const userProperties = PropertiesService.getUserProperties();
    userProperties.deleteAllProperties(); // Also clear username/role
    
    Logger.log('User logged out successfully');
    return { success: true };


  } catch (e) {
    Logger.log('Logout error: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Includes the content of another HTML file into a template.
 * This allows for component-based HTML structure (e.g., sidebars, navbars).
 * @param {string} filename The name of the HTML file to include (without .html).
 * @return {string} The HTML content of the file.
 */
function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}

/**
 * Modified: Fetches projects that are pending OR marked as 'NR'
 */
function getPendingProjects() {
  try {
    const externalSpreadsheetId = '1RnTKOiplVFLiZouq86fBdOm3PmaGpG0izXDwCZl0ow8';
    const sheetName = 'Main';
    const ss = SpreadsheetApp.openById(externalSpreadsheetId);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [['Error'], ['External sheet not found.']];

    const headerRow = 7;
    if (sheet.getLastRow() < headerRow) return [['Error'], ['No data in external sheet.']];

    const range = sheet.getRange(headerRow, 1, sheet.getLastRow() - headerRow + 1, sheet.getLastColumn());
    const values = range.getDisplayValues();
    const headers = values.shift().map(h => h.trim());
    const statusIndex = headers.indexOf('Status');
    const isoNoIndex = headers.indexOf('ISO No.');
    const bucketIndex = headers.indexOf('Bucket');

    if (statusIndex === -1 || isoNoIndex === -1 || bucketIndex === -1) {
      return [['Error'], ['Required headers "Status", "ISO No.", or "Bucket" not found.']];
    }

    const filteredRows = values.filter(row => {
      const statusValue = row[statusIndex].trim();
      const isoNoValue = row[isoNoIndex].trim();
      const bucketValue = row[bucketIndex].trim();
      
      // MODIFIED: Include both blank buckets and 'NR' buckets
      return statusValue === '' && isoNoValue !== '' && (bucketValue === '' || bucketValue === 'NR');
    });

    return [headers, ...filteredRows];

  } catch (e) {
    return [['Error'], [e.message]];
  }
}

/**
 * Finds a project by its ISO No. in the external sheet and updates its Bucket column.
 * @param {string} isoNo The unique ISO Number to identify the row.
 * @param {string} bucket The bucket value to set.
 */
function updateProjectBucket(isoNo, bucket) {
  try {
    const externalSpreadsheetId = '1RnTKOiplVFLiZouq86fBdOm3PmaGpG0izXDwCZl0ow8';
    const sheetName = 'Main';

    const ss = SpreadsheetApp.openById(externalSpreadsheetId);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error('External sheet not found.');

    const headerRow = 7;
    const range = sheet.getRange(headerRow, 1, sheet.getLastRow() - headerRow + 1, sheet.getLastColumn());
    const values = range.getValues();
    const headers = values.shift().map(h => h.trim());

    const isoNoIndex = headers.indexOf('ISO No.');
    const bucketColIndex = headers.indexOf('Bucket');

    if (isoNoIndex === -1 || bucketColIndex === -1) {
      throw new Error('Could not find required columns (ISO No. or Bucket).');
    }

    const rowIndex = values.findIndex(row => row[isoNoIndex] === isoNo);

    if (rowIndex === -1) {
      throw new Error('Project with ISO No. ' + isoNo + ' not found.');
    }

    const sheetRow = rowIndex + headerRow + 1;
    sheet.getRange(sheetRow, bucketColIndex + 1).setValue(bucket);

    return "Success";

  } catch (e) {
    return "Error: " + e.message;
  }
}

/**
 * Gets a list of project names filtered by a specific bucket.
 * @param {string} bucketName The name of the bucket to filter by (e.g., 'STP CIVIL').
 * @returns {string[]} An array of project names that match the bucket.
 */
function getProjectNamesByBucket(bucketName) {
  try {
    // Cached (name, bucket) pairs from Project Details Form; see projectIndex_().
    const names = projectIndex_().pairs
      .filter(p => p[1] === bucketName)
      .map(p => p[0]);

    // Return a unique list of names
    return [...new Set(names)];

  } catch (e) {
    Logger.log('Error in getProjectNamesByBucket: ' + e.toString());
    return []; // Return empty array on error
  }
}

/**
 * MASTER FUNCTION (OPTIMIZED): Fetches ALL sheets (including MOM) in one go.
 */
function getAllDashboardData(projectName) {
  const result = {
    details: null,
    kpis: null,
    timelineChart: [],
    milestoneChart: [],
    momData: { headers: [], rows: [] } // Added container for MOM data
  };

  try {
    // 1. The project index (cached) gives the ISO and bucket, so only that ISO's rows
    //    from that bucket's spreadsheet are read, in one batchGet, and cached for 10 min.
    const ctx = projectContext_(projectName);
    const sheetNames = [
      'Project Details Form', 'PROJECT PLANNING', 'Manpower Cost',
      'INVENTORY', 'Auditor Sheet Manpower', 'Auditor Sheet Item',
      'Auditor Sheet Drawing', 'FINANCE/ACCOUNTS', 'Billing', 'WBS Sheet', 'Expenses'
    ];
    const allData = loadSources_(sheetNames, ctx.ss, ctx.projectIds, { iso: ctx.isoNo });

    // 2. Process Project Details
    const detailsData = allData['Project Details Form'];
    if (detailsData.length > 0) {
      const headers = detailsData[0].map(h => String(h).trim());
      const projectIndex = headers.indexOf('PROJECT NAME');
      const isoIndex = headers.indexOf('ISO No.');
      
      const projectRow = detailsData.slice(1).find(row => row[projectIndex] === projectName);
      
      if (projectRow) {
        // const isoNo = projectRow[isoIndex];
        const isoNo = cleanIso(projectRow[isoIndex]);

        result.details = {
          isoNo: isoNo,
          projectAdd: projectRow[headers.indexOf('PROJECT ADD')],
          contactPerson: projectRow[headers.indexOf('CONTACT PERSON')],
          item: projectRow[headers.indexOf('item')]
        };

        if (!isoNo) return { error: 'ISO No. is empty for project "' + projectName + '" in Project Details Form.' };

        if (isoNo) {
          // 3. Process KPIs and Charts
          result.kpis = calculateKpisInMemory(isoNo, allData, result.details.orderAmount);
          result.timelineChart = calculateTimelineChartInMemory(isoNo, allData['WBS Sheet']);
          result.milestoneChart = calculateMilestoneChartInMemory(isoNo, allData['PROJECT PLANNING'], allData['Billing']);

          // 4. MOM rows for this ISO with image/PDF/signature links (a MOM failure must not
          //    take the whole dashboard down).
          try {
            result.momData = momTable_(isoNo);
          } catch (momErr) {
            Logger.log('MOM load failed: ' + momErr.toString());
          }
        }
      } else {
        return { error: "Project not found." };
      }
    }
    return result;

  } catch (e) {
    Logger.log("Error in getAllDashboardData: " + e.toString());
    return { error: e.toString() };
  }
}

/**
 * OPTIMIZED KPI CALCULATOR
 * Calculates all KPIs using data already loaded in memory. No spreadsheet calls.
 */
function calculateKpisInMemory(isoNo, allData, orderAmountFromDetails) {
  const kpis = {};
  
  // Helper to sum in memory
  const sumInMemory = (sheetName, headerRowIdx, colName, isoColName) => {
    const data = allData[sheetName];
    if (!data || data.length <= headerRowIdx) return 0;
    
    const headers = data[headerRowIdx - 1].map(h => String(h).trim());
    const valIdx = headers.indexOf(colName);
    let isoIdx = headers.indexOf(isoColName); if (isoIdx === -1) isoIdx = isoColIdx_(headers);
    
    if (valIdx === -1 || isoIdx === -1) return 0;

    return data.slice(headerRowIdx)
      .filter(row => cleanIso(row[isoIdx]) === isoNo)
      .reduce((sum, row) => sum + (parseFloat(row[valIdx]) || 0), 0);
  };

  // --- 1. Project Info (Sum Order Amount + Handle Commas) ---
  const detailsData = allData['Project Details Form'];
  const dHeaders = detailsData[0].map(h => String(h).trim());
  const dIsoIdx = isoColIdx_(dHeaders);
  const dRateIdx = dHeaders.indexOf('BOQ RATE(Rs)');
  const dManagerIdx = dHeaders.indexOf('PROJECT MANAGER');

  const dSiteInchargeIdx = dHeaders.indexOf('SITE INCHARGE'); 
  
  // Find ALL rows for this project to sum the amount
  const projectRows = detailsData.slice(1).filter(r => cleanIso(r[dIsoIdx]) === isoNo);
  
  // Calculate Total Order Amount
  let totalOrderAmount = 0;
  if (dRateIdx > -1) {
    totalOrderAmount = projectRows.reduce((sum, row) => {
      // 1. Convert to string to safely use replace
      // 2. Remove all commas (e.g., "1,50,000" -> "150000")
      // 3. Parse to float
      let rawValue = String(row[dRateIdx] || 0);
      let cleanValue = parseFloat(rawValue.replace(/,/g, ''));
      
      return sum + (isNaN(cleanValue) ? 0 : cleanValue);
    }, 0);
  }
  kpis.orderAmount = totalOrderAmount;

  const firstRow = projectRows.length > 0 ? projectRows[0] : null;
  
  kpis.projectManager = (firstRow && dManagerIdx > -1) ? firstRow[dManagerIdx] : 'N/A';

  kpis.siteIncharge = (firstRow && dSiteInchargeIdx > -1) ? firstRow[dSiteInchargeIdx] : 'N/A';
  
  // Payment Terms
  if (firstRow) {
      kpis.paymentTerms = {
        advance: firstRow[dHeaders.indexOf('ADVANCE (%)')] || 0,
        againstPi: firstRow[dHeaders.indexOf('AGAINST PI (%)')] || 0,
        installTest: firstRow[dHeaders.indexOf('INSTALLATION & TESTING (%)')] || 0,
        commission: firstRow[dHeaders.indexOf('COMMISSIONING & HANDOVER (%)')] || 0,
        security: firstRow[dHeaders.indexOf('SECURITY AMOUNT (%)')] || 0,
      };
      
      const poLinkCell = firstRow[dHeaders.indexOf('CUSTOMER PO COPY')];
      kpis.poCopy = (typeof poLinkCell === 'string' && poLinkCell.startsWith('http')) ? poLinkCell : '#';
  }

  // --- 2. Expenses ---
  kpis.manpowerUsedRs = sumInMemory("Manpower Cost", 1, "Manpower Cost", "ISO No.");
  kpis.productUsedRs = sumInMemory("INVENTORY", 2, "CONSU.2", "ISO No.");

  // --- 2. Expenses (UPDATED for Debit Amount) ---
  const expenseSheetData = allData['Expenses'];
  if(expenseSheetData && expenseSheetData.length > 0) {
      const eHeaders = expenseSheetData[0].map(h => String(h).trim());
      
      // Robust search for ISO column (handles "ISO No." or "ISO No")
      let eIsoIdx = eHeaders.indexOf('ISO No.');
      if (eIsoIdx === -1) eIsoIdx = eHeaders.indexOf('ISO No');

      // Search for Debit Amount column
      const eDebitIdx = eHeaders.indexOf('Debit Amount');
      
      if (eIsoIdx > -1 && eDebitIdx > -1) {
          const targetIso = String(isoNo).trim();
          
          kpis.expensesRs = expenseSheetData.slice(1)
            .filter(r => {
                // Extract clean ISO (removes " - Client Name" if present)
                const rowIso = String(r[eIsoIdx]).split(' - ')[0].trim();
                return rowIso === targetIso;
            })
            .reduce((sum, r) => {
                // Remove commas and sum the Debit Amount
                let rawVal = String(r[eDebitIdx] || 0);
                let val = parseFloat(rawVal.replace(/,/g, ''));
                return sum + (isNaN(val) ? 0 : val);
            }, 0);
      } else { 
        Logger.log("Expenses Error: 'ISO No.' or 'Debit Amount' column not found.");
        kpis.expensesRs = 0; 
      }
  } else { kpis.expensesRs = 0; }

  // --- 3. Product Status ---
  kpis.indentPlanned = sumInMemory("INVENTORY", 2, "Planned Indent Qty", "ISO No.");
  kpis.actualIndent = sumInMemory("INVENTORY", 2, "INDENT1", "ISO No.");
  kpis.received = sumInMemory("INVENTORY", 2, "REC.1", "ISO No.");
  kpis.used = sumInMemory("INVENTORY", 2, "CONSU.1", "ISO No.");

  // --- 4. Manpower & Engineer ---
  const pData = allData['PROJECT PLANNING'];
  if (pData.length > 0) {
      const pHeaders = pData[0].map(h => String(h).trim());
      const pIsoIdx = isoColIdx_(pHeaders);
      const doerIdx = pHeaders.indexOf("Engineer Name (Doer)");
      const reqHrsIdx = pHeaders.indexOf("Required Hours");
      const workedHrsIdx = pHeaders.indexOf("Worked (Hrs)");
      
      if (pIsoIdx > -1) {
          const projRows = pData.slice(1).filter(r => cleanIso(r[pIsoIdx]) === isoNo);
          
          if (doerIdx > -1) {
             const doers = new Set(projRows.map(r => r[doerIdx]).filter(String));
             kpis.engWorked = doers.size;
          }
          if (reqHrsIdx > -1) kpis.plannedManHrs = projRows.reduce((sum, r) => sum + (parseFloat(r[reqHrsIdx])||0), 0);
          if (workedHrsIdx > -1) kpis.actualManHrs = projRows.reduce((sum, r) => sum + (parseFloat(r[workedHrsIdx])||0), 0);
      }
  }

  // --- 5. Audit ---
  kpis.auditManpower = sumInMemory("Auditor Sheet Manpower", 2, "Deviation", "ISO No.");
  kpis.auditProduct = sumInMemory("Auditor Sheet Item", 2, "Deviation", "ISO No.");
  kpis.auditDrawings = sumInMemory("Auditor Sheet Drawing", 2, "Deviation", "ISO NO.");

  
  // --- 6. Milestone Summary ---
  const wbsData = allData['WBS Sheet'];
  if (wbsData && wbsData.length > 0) {
    const wHeaders = wbsData[0].map(h => String(h).trim());
    const wIsoIdx = isoColIdx_(wHeaders);
    const wMileIdx = wHeaders.indexOf("MILESTONE/BOQ ITEM"); // Matches your screenshot column C
    const wActualEndIdx = wHeaders.indexOf("ACTUAL END DATE"); // Matches your screenshot column K (assumed)

    if (wIsoIdx > -1 && wMileIdx > -1) {
      const targetIso = String(isoNo).trim();
      const projWbsRows = wbsData.slice(1).filter(r => cleanIso(r[wIsoIdx]) === targetIso);

      // Group by Unique Milestone Name to get the count of 12
      const milestoneGroups = {}; 
      projWbsRows.forEach(row => {
        const mName = String(row[wMileIdx]).trim();
        if (!mName || mName === "N/A" || mName === "") return;
        if (!milestoneGroups[mName]) milestoneGroups[mName] = [];
        milestoneGroups[mName].push(row);
      });

      const uniqueMilestoneNames = Object.keys(milestoneGroups);
      kpis.totalMilestones = uniqueMilestoneNames.length; // This will now show 12

      let mCompleted = 0;
      uniqueMilestoneNames.forEach(name => {
        const tasksInThisMilestone = milestoneGroups[name];
        // A milestone is "Done" if every row has an Actual End Date
        const isAllDone = tasksInThisMilestone.every(t => {
          const dateVal = t[wActualEndIdx];
          return dateVal && String(dateVal).trim() !== "" && String(dateVal).toUpperCase() !== "N/A";
        });
        if (isAllDone) mCompleted++;
      });

      kpis.completedMilestones = mCompleted;
      kpis.pendingMilestones = kpis.totalMilestones - mCompleted;
      kpis.milestoneDonePercent = kpis.totalMilestones > 0 ? Math.round((mCompleted / kpis.totalMilestones) * 100) : 0;
    }
  }

  // --- 7. Task Status (Keep this pulling from PROJECT PLANNING) ---
  if (pData && pData.length > 0) {
      const pHeaders = pData[0].map(h => String(h).trim());
      const pIsoIdx = isoColIdx_(pHeaders);
      const statusIdx = pHeaders.indexOf("Status");

      if (pIsoIdx > -1 && statusIdx > -1) {
          const targetIsoStr = String(isoNo).trim();
          
          // 1. Filter all rows that match this ISO
          const projTasks = pData.slice(1).filter(row => {
            const rowIso = cleanIso(row[pIsoIdx]);
            return rowIso === targetIsoStr;
          });

          let done = 0, hold = 0, inprog = 0;
          
          // 2. Process every task found for this ISO
          projTasks.forEach(r => {
              const s = String(r[statusIdx] || "").trim().toUpperCase();

              if (s === 'DONE') {
                  done++;
              } else if (s === 'IN-PROGRESS' || s === 'IN PROGRESS') {
                  inprog++;
              } else {
                  // UPDATED: This counts 'HOLD' as well as any EMPTY/BLANK cells as HOLD
                  hold++;
              }
          });

          // 3. Update the KPIs
          kpis.totalTasks = done + hold + inprog; // This will now show the actual total
          kpis.tasksDone = done;
          kpis.tasksInProgress = inprog;
          kpis.tasksOnHold = hold;
      } else {
          Logger.log("Error: Could not find 'ISO No.' or 'Status' column in PROJECT PLANNING");
      }
  }

  // --- 8. Billing FMS (FIXED & ROBUST) ---
  const bData = allData['Billing'];
  if (bData && bData.length > 1) {
      const bHeaders = bData[0].map(h => String(h).trim());
      
      // 1. Find Headers (Robust Search)
      // Look for "ISO NO" OR "ISO No."
      let bIsoIdx = bHeaders.indexOf("ISO NO");
      if (bIsoIdx === -1) bIsoIdx = bHeaders.indexOf("ISO No.");
      
      const bMilestoneIdx = bHeaders.indexOf("Milestones");
      
      // Look for "Status5" OR "Status"
      let bStatusIdx = bHeaders.indexOf("Status5");
      if (bStatusIdx === -1) bStatusIdx = bHeaders.indexOf("Status");

      if (bIsoIdx > -1 && bMilestoneIdx > -1 && bStatusIdx > -1) {
          const targetIso = String(isoNo).trim(); // This is already cleaned by your previous update
          
          // 2. Filter Rows using cleanIso on the sheet data
          const bRows = bData.slice(1).filter(r => {
             // We clean the row data on the fly (e.g. remove "- COMPANY NAME")
             return cleanIso(r[bIsoIdx]) === targetIso;
          });

          const totalB = bRows.length;
          
          if (totalB > 0) {
              // Count Unique Milestones
              const uniqueM = new Set(bRows.map(r => String(r[bMilestoneIdx]).trim()).filter(Boolean));
              kpis.totalBillingMilestones = uniqueM.size;
              
              // Count Completed Milestones
              let completedMCount = 0;
              uniqueM.forEach(m => {
                  const mRows = bRows.filter(r => String(r[bMilestoneIdx]).trim() === m);
                  // Check if ALL rows for this milestone are DONE
                  const allDone = mRows.every(r => String(r[bStatusIdx]).trim().toUpperCase() === 'DONE');
                  if(allDone && mRows.length > 0) completedMCount++;
              });
              kpis.completedBillingMilestones = completedMCount;

              // Calculate Percentages
              const doneB = bRows.filter(r => String(r[bStatusIdx]).trim().toUpperCase() === 'DONE').length;
              
              // Avoid division by zero
              if (totalB > 0) {
                  kpis.overallBillingDonePercent = Math.round((doneB / totalB) * 100);
                  kpis.overallBillingPendingPercent = 100 - kpis.overallBillingDonePercent;
              }
          }
      } else {
        Logger.log("Billing Error: Could not find ISO, Milestones, or Status5/Status columns.");
      }
  }

  // --- 9. Finance ---
  const fData = allData['FINANCE/ACCOUNTS'];
  if (fData && fData.length > 1) {
      const fHeaders = fData[0].map(h => String(h).trim());
      const fIsoIdx = isoColIdx_(fHeaders);
      const payIdx = fHeaders.indexOf("% of  Pymt recd");
      
      if (fIsoIdx > -1 && payIdx > -1) {
          // Use slice(1) to skip header, find first match
          const fRow = fData.slice(1).find(r => cleanIso(r[fIsoIdx]) === isoNo);
          kpis.paymentReceivedPercent = fRow ? (fRow[payIdx] || 0) : 0;
      }
  }

  // --- 10. Profit Loss ---
  const orderAmt = parseFloat(kpis.orderAmount) || 0;
  const costs = (kpis.manpowerUsedRs || 0) + (kpis.productUsedRs || 0) + (kpis.expensesRs || 0);
  kpis.profitLoss = orderAmt - costs;

  return kpis;
}

/**
 * OPTIMIZED CHART CALCULATOR (Timeline)
 */
function calculateTimelineChartInMemory(isoNo, wbsData) {
  if (!wbsData || wbsData.length < 2) return [];
  const headers = wbsData[0].map(h => String(h).trim());
  const isoIdx = isoColIdx_(headers);
  const mileIdx = headers.indexOf('MILESTONE/BOQ ITEM');
  const planIdx = headers.indexOf('PLANNED START DATE');
  const actIdx = headers.indexOf('ACTUAL END DATE');

  if (isoIdx === -1) return [];

  return wbsData.slice(1)
    .filter(row => cleanIso(row[isoIdx]) === isoNo)
    .map(row => {
      const pDate = row[planIdx];
      const aDate = row[actIdx];
      
      // Basic date check
      const pD = toDate_(pDate), aD = toDate_(aDate);
      const fPDate = pD ? Utilities.formatDate(pD, Session.getScriptTimeZone(), "dd/MM/yyyy") : null;
      const fADate = (aD && aD.getTime() > 0) ? Utilities.formatDate(aD, Session.getScriptTimeZone(), "dd/MM/yyyy") : null;

      return {
        milestone: row[mileIdx],
        planned: fPDate,
        actual: fADate
      };
    })
    .filter(d => d.milestone && d.planned);
}

/**
 * CORRECTED: Fetches and combines data for the Milestone Progress Combo Chart.
 * Uses cleanIso to ensure data is found even if suffixes exist in the Billing sheet.
 */
function calculateMilestoneChartInMemory(isoNo, planningData, billingData) {
  if (!planningData || !billingData || !isoNo) return [];
  
  const targetIso = cleanIso(isoNo);

  // --- 1. Process Planning Data ---
  const pHeaders = planningData[0].map(h => String(h).trim());
  const pIsoIdx = isoColIdx_(pHeaders);
  const pMileIdx = pHeaders.indexOf("Milestone");
  const pStatIdx = pHeaders.indexOf("Status %");
  
  // Filter planning tasks using cleanIso
  const pTasks = pIsoIdx > -1 ? planningData.slice(1).filter(r => cleanIso(r[pIsoIdx]) === targetIso) : [];

  // --- 2. Process Billing Data ---
  const bHeaders = billingData[0].map(h => String(h).trim());
  
  // Robust Billing Header Search (Same as KPI logic)
  let bIsoIdx = bHeaders.indexOf("ISO NO");
  if (bIsoIdx === -1) bIsoIdx = bHeaders.indexOf("ISO No.");
  
  const bMileIdx = bHeaders.indexOf("Milestones");
  
  let bStatusIdx = bHeaders.indexOf("Status5");
  if (bStatusIdx === -1) bStatusIdx = bHeaders.indexOf("Status");

  // Filter billing tasks using cleanIso
  const bTasks = bIsoIdx > -1 ? billingData.slice(1).filter(r => cleanIso(r[bIsoIdx]) === targetIso) : [];

  // --- 3. Combine Milestone Names for the X-Axis ---
  const allMiles = new Set([
    ...pTasks.map(r => String(r[pMileIdx] || "").trim()),
    ...bTasks.map(r => String(r[bMileIdx] || "").trim())
  ].filter(Boolean));

  const chartData = [];
  allMiles.forEach(m => {
      // Calculate Task Progress (Brown Bar)
      const myTasks = pTasks.filter(r => String(r[pMileIdx] || "").trim() === m);
      let tProg = 0;
      if (myTasks.length > 0) {
          const sumPct = myTasks.reduce((s, r) => s + (parseFloat(r[pStatIdx]) || 0), 0);
          // Check if data is decimal (0.8) or whole (80)
          tProg = sumPct / myTasks.length;
          if (tProg <= 1 && tProg > 0) tProg = tProg * 100; 
      }

      // Calculate Billing Status (Yellow Bar)
      const myBills = bTasks.filter(r => String(r[bMileIdx] || "").trim() === m);
      let bStat = 0;
      if (myBills.length > 0) {
          const done = myBills.filter(r => String(r[bStatusIdx] || "").trim().toUpperCase() === 'DONE').length;
          bStat = (done / myBills.length) * 100;
      }

      chartData.push({
          milestone: m,
          taskProgress: tProg,
          billingStatus: bStat
      });
  });

  // Sort by Task Progress descending so the chart looks organized
  return chartData.sort((a, b) => b.taskProgress - a.taskProgress);
}

// /**
//  * OPTIMIZED CHART CALCULATOR (Milestone Combo)
//  */
// function calculateMilestoneChartInMemory(isoNo, planningData, billingData) {
//   if (!planningData || !billingData) return [];
  
//   // Process Planning
//   const pHeaders = planningData[0].map(h => String(h).trim());
//   const pIsoIdx = pHeaders.indexOf("ISO No.");
//   const pMileIdx = pHeaders.indexOf("Milestone");
//   const pStatIdx = pHeaders.indexOf("Status %");
//   const pTasks = pIsoIdx > -1 ? planningData.slice(1).filter(r => r[pIsoIdx] === isoNo) : [];

//   // Process Billing
//   const bHeaders = billingData[0].map(h => String(h).trim());
//   const bIsoIdx = bHeaders.indexOf("ISO NO");
//   const bMileIdx = bHeaders.indexOf("Milestones");
//   const bStatIdx = bHeaders.indexOf("Status5");
//   const bTasks = bIsoIdx > -1 ? billingData.slice(1).filter(r => r[bIsoIdx] === isoNo) : [];

//   // Combine Milestones
//   const allMiles = new Set([
//     ...pTasks.map(r => r[pMileIdx]),
//     ...bTasks.map(r => r[bMileIdx])
//   ].filter(Boolean));

//   const chartData = [];
//   allMiles.forEach(m => {
//       // Task Progress
//       const myTasks = pTasks.filter(r => r[pMileIdx] === m);
//       let tProg = 0;
//       if (myTasks.length > 0) {
//           const sumPct = myTasks.reduce((s, r) => s + (parseFloat(r[pStatIdx]) || 0), 0);
//           tProg = (sumPct / myTasks.length) * 100; // Assuming raw data is 0.5 for 50%, if 50, remove *100
//           // Wait, earlier we found your data is likely 0-1 or 0-100?
//           // If your data is 0.8 for 80%, keep *100. If it's 80, remove *100.
//           // Based on previous code "parseFloat(row[statusPercentIdx]) || 0", let's assume it's already 0-100 based on your chart fix.
//           // ADJUSTMENT: If your sheet has "80%", parseFloat is 80. If "0.8", it is 0.8.
//           // Let's assume standard number 80.
//           tProg = sumPct / myTasks.length; 
//       }

//       // Billing Status
//       const myBills = bTasks.filter(r => r[bMileIdx] === m);
//       let bStat = 0;
//       if (myBills.length > 0) {
//           const done = myBills.filter(r => String(r[bStatIdx]).trim().toUpperCase() === 'DONE').length;
//           bStat = (done / myBills.length) * 100;
//       }

//       chartData.push({
//           milestone: m,
//           taskProgress: tProg,
//           billingStatus: bStat
//       });
//   });

//   return chartData;
// }

/**
 * Gets specific details for a selected project name.
 */
function getProjectDetails(projectName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Project Details Form');
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const headers = data.shift();

  const projectIndex = headers.indexOf('PROJECT NAME');
  const isoIndex = headers.indexOf('ISO No.');
  const addIndex = headers.indexOf('PROJECT ADD');
  const personIndex = headers.indexOf('CONTACT PERSON');
  const itemIndex = headers.indexOf('item');

  const projectRow = data.find(row => row[projectIndex] === projectName);

  if (projectRow) {
    return {
      isoNo: projectRow[isoIndex],
      projectAdd: projectRow[addIndex],
      contactPerson: projectRow[personIndex],
      item: projectRow[itemIndex]
    };
  }
  return null;
}

/**
 * Fetches and prepares data for the D3.js timeline chart, using ISO No.
 * @param {string} isoNo The ISO Number of the project.
 */
function getChartData(isoNo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wbsSheet = ss.getSheetByName('WBS Sheet');

  if (!wbsSheet || !isoNo) return [];

  const wbsData = wbsSheet.getDataRange().getValues();
  const wbsHeaders = wbsData.shift();

  const wbsIsoIndex = wbsHeaders.indexOf('ISO No.');
  const milestoneIndex = wbsHeaders.indexOf('MILESTONE/BOQ ITEM');
  const plannedStartIndex = wbsHeaders.indexOf('PLANNED START DATE');
  const actualEndIndex = wbsHeaders.indexOf('ACTUAL END DATE');
  if (wbsIsoIndex === -1) return [];

  const chartData = wbsData
    .filter(row => row[wbsIsoIndex] === isoNo)
    .map(row => {
      const plannedDateValue = row[plannedStartIndex];
      const actualEndDateValue = row[actualEndIndex];

      const plannedDate = (plannedDateValue instanceof Date) ?
      Utilities.formatDate(plannedDateValue, Session.getScriptTimeZone(), "dd/MM/yyyy") :
      null;

      const actualEndDate = (actualEndDateValue instanceof Date && actualEndDateValue.getTime() > 0) ?
      Utilities.formatDate(actualEndDateValue, Session.getScriptTimeZone(), "dd/MM/yyyy") :
      null;

      return {
        milestone: row[milestoneIndex],
        planned: plannedDate,
        actual: actualEndDate
      };
    })
    .filter(d => d.milestone && d.planned);
  return chartData;
}

/**
 * Gathers a comprehensive set of Key Performance Indicators (KPIs) for a given project.
 * @param {string} projectName The name of the project to fetch KPIs for.
 * @returns {object} An object containing all the calculated KPIs.
 */
function getProjectKpis(projectName = 'Hyatt regency') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const kpis = {}; // Object to hold all our results

  // 1. Get the ISO No. for the project (our primary key)
  const detailsSheet = ss.getSheetByName('Project Details Form');
  if (!detailsSheet) return { error: "Project Details Form not found." };
  const detailsData = detailsSheet.getDataRange().getValues();
  const detailsHeaders = detailsData.shift();
  const projectIndex = detailsHeaders.indexOf('PROJECT NAME');
  const isoIndex = detailsHeaders.indexOf('ISO No.');

  const projectRow = detailsData.find(row => row[projectIndex] === projectName);
  if (!projectRow) return { error: `Project '${projectName}' not found.` };
  const isoNo = projectRow[isoIndex];

  // 2. Fetch data from "Project Details Form"
  kpis.orderAmount = projectRow[detailsHeaders.indexOf('BOQ RATE(Rs)')] || 0;
  kpis.projectManager = projectRow[detailsHeaders.indexOf('PROJECT MANAGER')] || 'N/A';

  const poCopyIndex = detailsHeaders.indexOf('CUSTOMER PO COPY');
  let poLinkUrl = '#';
  const projectRowIndexInData = detailsData.findIndex(row => row[projectIndex] === projectName);
  if (poCopyIndex !== -1 && projectRowIndexInData !== -1) {
    const sheetRowNumber = projectRowIndexInData + 2;
    const sheetColumnNumber = poCopyIndex + 1;
    const poCopyCell = detailsSheet.getRange(sheetRowNumber, sheetColumnNumber);
    const poLink = poCopyCell.getRichTextValue();
    if (poLink) {
        poLinkUrl = poLink.getLinkUrl() || '#';
    }
  }
  kpis.poCopy = poLinkUrl;

  kpis.paymentTerms = {
    advance: projectRow[detailsHeaders.indexOf('ADVANCE (%)')] || 0,
    againstPi: projectRow[detailsHeaders.indexOf('AGAINST PI (%)')] || 0,
    installTest: projectRow[detailsHeaders.indexOf('INSTALLATION & TESTING (%)')] || 0,
    commission: projectRow[detailsHeaders.indexOf('COMMISSIONING & HANDOVER (%)')] || 0,
    security: projectRow[detailsHeaders.indexOf('SECURITY AMOUNT (%)')] || 0,
  };

  // Helper function to sum a column in a sheet based on ISO No.
  const sumColumnByIso = (sheetName, headerRowIndex, columnToSum, isoColumn) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return 0;
    const data = sheet.getDataRange().getValues();
    if (data.length < headerRowIndex) return 0;
    const headers = data[headerRowIndex - 1];
    const colIdx = headers.indexOf(columnToSum);
    const isoColIdx = headers.indexOf(isoColumn);
    if (colIdx === -1 || isoColIdx === -1) return 0;

    return data
      .slice(headerRowIndex)
      .filter(row => row[isoColIdx] === isoNo)
      .reduce((sum, row) => sum + (parseFloat(row[colIdx]) || 0), 0);
  };

  // 3. Calculate Expense Data
  kpis.manpowerUsedRs = sumColumnByIso("Manpower Cost", 1, "Manpower Cost", "ISO No.");
  kpis.productUsedRs = sumColumnByIso("INVENTORY", 2, "CONSU.2", "ISO No.");
  kpis.expensesRs = sumColumnByIso("Expenses", 1, "Amount", "ISO No.");

  // 4. Product Status Data
  kpis.indentPlanned = sumColumnByIso("INVENTORY", 2, "Qty", "ISO No.");      // NEW: From 'Qty' column
  kpis.actualIndent = sumColumnByIso("INVENTORY", 2, "INDENT1", "ISO No.");   // RENAMED: from 'indent'
  kpis.received = sumColumnByIso("INVENTORY", 2, "REC.1", "ISO No.");
  kpis.used = sumColumnByIso("INVENTORY", 2, "CONSU.1", "ISO No.");

  // 5. Manpower & Engineer Data
  const planningSheet = ss.getSheetByName("PROJECT PLANNING");
  let engWorked = 0, plannedManHrs = 0, actualManHrs = 0;
  if (planningSheet) {
    const pData = planningSheet.getDataRange().getValues();
    const pHeaders = pData.shift();
    const pIsoIdx = pHeaders.indexOf("ISO No.");
    // --- THIS IS THE CORRECTED LINE ---
    const doerNameIdx = pHeaders.indexOf("Engineer Name (Doer)"); // Targeting the correct column
    const reqHrsIdx = pHeaders.indexOf("Required Hours");
    const workedHrsIdx = pHeaders.indexOf("Worked (Hrs)");

    if (pIsoIdx > -1 && doerNameIdx > -1) {
        const projectPlanningData = pData.filter(r => r[pIsoIdx] === isoNo);
        
        // --- THIS IS THE CORRECTED LOGIC ---
        // It now maps the correct column, filters out any blank names, and then creates the unique Set.
        const uniqueEngineers = new Set(projectPlanningData.map(r => r[doerNameIdx]).filter(String));
        
        engWorked = uniqueEngineers.size;
        plannedManHrs = projectPlanningData.reduce((sum, r) => sum + (parseFloat(r[reqHrsIdx]) || 0), 0);
        actualManHrs = projectPlanningData.reduce((sum, r) => sum + (parseFloat(r[workedHrsIdx]) || 0), 0);
    }
  }
  kpis.engWorked = engWorked;
  kpis.plannedManHrs = plannedManHrs;
  kpis.actualManHrs = actualManHrs;

  // 6. Audit Data
  // Headers are on row 2 for all audit sheets. Data starts on row 3.
  // The slice() method in the helper correctly handles this.
  kpis.auditManpower = sumColumnByIso("Auditor Sheet Manpower", 2, "Deviation", "ISO No.");
  kpis.auditProduct = sumColumnByIso("Auditor Sheet Item", 2, "Deviation", "ISO No."); 
  kpis.auditDrawings = sumColumnByIso("Auditor Sheet Drawing", 2, "Deviation", "ISO No.");

  // 7. Task Status Data (4-COLUMN LOGIC)
  let totalTasks = 0, doneCount = 0, holdCount = 0, inProgressCount = 0;
  if (planningSheet) {
    const pData = planningSheet.getDataRange().getValues();
    const pHeaders = pData.shift();
    const pIsoIdx = pHeaders.indexOf("ISO No.");
    const statusIdx = pHeaders.indexOf("Status");

    if (pIsoIdx > -1 && statusIdx > -1) {
        const projectTasks = pData.filter(r => r[pIsoIdx] === isoNo);
        projectTasks.forEach(task => {
            const status = (task[statusIdx] || '').toString().toUpperCase().trim();
            switch (status) {
                case 'DONE': doneCount++; break;
                case 'HOLD': holdCount++; break;
                case 'IN-PROGRESS': inProgressCount++; break;
            }
        });
        totalTasks = doneCount + holdCount + inProgressCount;
    }
  }
  kpis.totalTasks = totalTasks;
  kpis.tasksDone = doneCount;
  kpis.tasksInProgress = inProgressCount; // Now a separate value
  kpis.tasksOnHold = holdCount;         // Now a separate value


  // ====================================================================
  // ===== 8. Finance Data  =====
  // ====================================================================
  // --- Part 1: Payment Received % ---
  const financeSheet = ss.getSheetByName("FINANCE/ACCOUNTS");
  kpis.paymentReceivedPercent = 0;
  if (financeSheet) {
    const fData = financeSheet.getDataRange().getValues();
    if (fData.length > 1) {
        const fHeaders = fData[0];
        const fIsoIdx = fHeaders.indexOf("ISO No.");
        const paymentRecdIdx = fHeaders.indexOf("% of  Pymt recd");
        if (fIsoIdx > -1) {
            const dataRows = fData.slice(1);
            const financeRow = dataRows.find(r => r[fIsoIdx] === isoNo);
            if (financeRow) {
              kpis.paymentReceivedPercent = financeRow[paymentRecdIdx] || 0;
            }
        }
    }
  }

  // ====================================================================
  // ============= 9. MILESTONE SUMMARY STATUS DATA =====================
  // ====================================================================
  kpis.totalMilestones = 0;
  kpis.completedMilestones = 0;
  kpis.pendingMilestones = 0;

  const wbsData = allData['WBS Sheet'];
  if (wbsData && wbsData.length > 0) {
    const wHeaders = wbsData[0].map(h => String(h).trim());
    const wIsoIdx = wHeaders.indexOf("ISO No.");
    const wMileIdx = wHeaders.indexOf("MILESTONE/BOQ ITEM");
    const wActualEndIdx = wHeaders.indexOf("ACTUAL END DATE");

    if (wIsoIdx > -1 && wMileIdx > -1 && wActualEndIdx > -1) {
      // 1. Filter rows for this project (Using TRIM for robust matching)
      const projectIsoStr = String(isoNo).trim();
      const projWbsRows = wbsData.slice(1).filter(r => String(r[wIsoIdx]).trim() === projectIsoStr);

      
      // 2. Group all 236+ rows by their UNIQUE Milestone Name
      const milestoneGroups = {}; 
      projWbsRows.forEach(row => {
        const mName = String(row[wMileIdx]).trim();
        if (!mName) return; // Skip empty rows
        if (!milestoneGroups[mName]) milestoneGroups[mName] = [];
        milestoneGroups[mName].push(row);
      });

      const uniqueMilestoneNames = Object.keys(milestoneGroups);
      let completedCount = 0;

      // 3. Evaluate each Unique Milestone group
      uniqueMilestoneNames.forEach(name => {
        const tasksInThisMilestone = milestoneGroups[name];
        
        // A milestone is considered "DONE" only if EVERY row inside it 
        // has an Actual End Date filled in.
        const isAllDone = tasksInThisMilestone.every(t => {
          const dateVal = t[wActualEndIdx];
          return dateVal && String(dateVal).trim() !== "" && String(dateVal).toUpperCase() !== "N/A";
        });

        if (isAllDone) completedCount++;
      });

      // 4. Update the KPIs with unique counts
      kpis.totalMilestones = uniqueMilestoneNames.length;
      kpis.completedMilestones = completedCount;
      kpis.pendingMilestones = kpis.totalMilestones - completedCount;
    }
  }
  
  // ====================================================================
  // ===== END OF WBS SUMMARY LOGIC =====
  // ====================================================================

  // ====================================================================
  // ===== 10. NEW: BILLING FMS SUMMARY DATA =====
  // ====================================================================
  kpis.totalBillingMilestones = 0;
  kpis.completedBillingMilestones = 0; // NEW value to return
  kpis.overallBillingDonePercent = 0;
  kpis.overallBillingPendingPercent = 0;
  const billingSheet = ss.getSheetByName("Billing");
  if (billingSheet) {
    const bData = billingSheet.getDataRange().getValues();
    if (bData.length > 1) {
      const bHeaders = bData[0].map(h => h.trim());
      const bIsoIdx = bHeaders.indexOf("ISO NO");
      const bMilestoneIdx = bHeaders.indexOf("Milestones");
      const bStatusIdx = bHeaders.indexOf("Status5");

      if (bIsoIdx > -1 && bMilestoneIdx > -1 && bStatusIdx > -1) {
        const projectBillings = bData.slice(1).filter(r => r[bIsoIdx] === isoNo);
        const totalBillings = projectBillings.length;

        if (totalBillings > 0) {
          const milestoneGroups = {};
          projectBillings.forEach(row => {
              const milestoneName = row[bMilestoneIdx];
              if (!milestoneName) return;
              if (!milestoneGroups[milestoneName]) {
                  milestoneGroups[milestoneName] = [];
              }
              milestoneGroups[milestoneName].push(row);
          });

          kpis.totalBillingMilestones = Object.keys(milestoneGroups).length;
          
          let completedMilestoneCount = 0;
          for (const milestoneName in milestoneGroups) {
              const billingsForMilestone = milestoneGroups[milestoneName];
              const allDone = billingsForMilestone.every(b => (b[bStatusIdx] || '').toString().trim().toUpperCase() === 'DONE');
              if (allDone) {
                  completedMilestoneCount++;
              }
          }
          kpis.completedBillingMilestones = completedMilestoneCount;

          // Overall percentage calculation remains the same
          const doneBillings = projectBillings.filter(r => (r[bStatusIdx] || '').toString().trim().toUpperCase() === 'DONE').length;
          const donePercent = Math.round((doneBillings / totalBillings) * 100);
          
          kpis.overallBillingDonePercent = donePercent;
          kpis.overallBillingPendingPercent = 100 - donePercent;
        }
      }
    }
  }

  // ====================================================================
  // ===== 11. NEW: PROFIT / LOSS CALCULATION =====
  // ====================================================================
  const orderAmount = parseFloat(kpis.orderAmount) || 0;
  const totalCosts = (kpis.manpowerUsedRs || 0) + (kpis.productUsedRs || 0) + (kpis.expensesRs || 0);
  kpis.profitLoss = orderAmount - totalCosts;

  
  // ====================================================================
  // ===== END OF NEW SECTION =====
  // ====================================================================
  
  return kpis;
}

/**
 * Helper to clean the ISO string.
 * Converts "MWEPL/2025-26/02641 - NATUROVILLE HOSPITALITY" -> "MWEPL/2025-26/02641"
 */
function cleanIso(rawIso) {
  if (!rawIso) return "";
  // Convert to string, split by " - " (space hyphen space), take the first part, and trim whitespace
  return String(rawIso).split(' - ')[0].trim();
}