function doGet(e) {
  return HtmlService.createTemplateFromFile('Dashboard')
    .evaluate()
    .setTitle('Dashboard Summary')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Fetches data from Close Cases and Forwarded Cases and returns aggregated dashboard metrics.
 */
function getSheetByNameTrimmed(ss, targetName) {
  if (!ss) return null;
  const direct = ss.getSheetByName(targetName);
  if (direct) return direct;
  
  const clean = targetName.toLowerCase().replace(/\s+/g, '');
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const sName = sheets[i].getName().toLowerCase().replace(/\s+/g, '');
    if (sName === clean) {
      return sheets[i];
    }
  }
  return null;
}

/**
 * Normalizes employee names: trims extra spaces and standardizes to Title Case (e.g. "ANSHUL" -> "Anshul", "anshul" -> "Anshul").
 */
function normalizeEmployeeName(rawName) {
  if (rawName === undefined || rawName === null) return '';
  let name = String(rawName).trim();
  if (!name || name === '-' || name === 'N/A' || name === 'null' || name === 'undefined') {
    return '';
  }
  
  // Collapse multiple spaces
  name = name.replace(/\s+/g, ' ');
  
  // Ignore header text values
  const lower = name.toLowerCase();
  if (lower.includes('closed') || lower.includes('forwarded') || lower === 'employee') {
    return '';
  }
  
  // Standard Title Case to merge UPPERCASE/lowercase versions without altering spelling
  return name.split(' ')
    .map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
    .join(' ');
}

/**
 * Fetches data from Close Cases and Forwarded Cases and returns aggregated dashboard metrics.
 */
/**
 * Parses diverse date formats from Google Sheets into a standard Date object
 */
function parseDateValue(rawDate) {
  if (!rawDate) return null;
  if (rawDate instanceof Date) {
    if (isNaN(rawDate.getTime())) return null;
    return new Date(rawDate.getFullYear(), rawDate.getMonth(), rawDate.getDate());
  }
  const str = String(rawDate).trim();
  if (!str || str === '-' || str === 'N/A') return null;
  
  // Format: YYYY-MM-DD or YYYY/MM/DD
  if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}/.test(str)) {
    const parts = str.split(/[-/]/);
    return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
  }
  
  // Format: M/D/YYYY or D/M/YYYY
  if (/^\d{1,2}[-/]\d{1,2}[-/]\d{2,4}/.test(str)) {
    const parts = str.split(/[-/]/);
    let year = parseInt(parts[2], 10);
    if (year < 100) year += 2000;
    const month = parseInt(parts[0], 10) - 1;
    const day = parseInt(parts[1], 10);
    return new Date(year, month, day);
  }
  
  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
  }
  return null;
}

/**
 * Fetches data from Close Cases and Forwarded Cases and returns aggregated dashboard metrics.
 */
/**
 * Formats a Date object into DD-MMM format (e.g. 01-Aug)
 */
function formatDateStr(d) {
  if (!d) return '';
  const day = String(d.getDate()).padStart(2, '0');
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return `${day}-${months[d.getMonth()]}`;
}

/**
 * Fetches data from Close Cases and Forwarded Cases and returns aggregated dashboard metrics.
 */
function getDashboardData(filterEmp = 'All Employees', filterAct = 'All Activities', dateFrom = '', dateTo = '') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const closeCasesSheet = getSheetByNameTrimmed(ss, "Close Cases");
  const forwardedCasesSheet = getSheetByNameTrimmed(ss, "Forwarded cases");
  
  let allEntries = [];
  let totalMasterRows = 0;
  
  // Section Column definitions (Header lists for UI)
  const structure = {
    ccS1Cols: [],
    ccS2Cols: [],
    fcS1Cols: [],
    fcS2Cols: []
  };

  // 1. Process Close Cases Sheet
  if (closeCasesSheet) {
    const data = closeCasesSheet.getDataRange().getValues();
    if (data.length > 1) {
      totalMasterRows += (data.length - 1);
      const headers = data[0].map(h => String(h || '').trim());
      
      const empCols = [];
      headers.forEach((h, idx) => {
        const lower = h.toLowerCase().replace(/[^a-z]/g, '');
        if (lower.includes('closedby') || lower === 'closedby') {
          empCols.push(idx);
        }
      });
      const s1EmpIdx = empCols.length > 0 ? empCols[0] : 5;
      const s2EmpIdx = empCols.length > 1 ? empCols[1] : 10;
      
      for (let c = 1; c < s1EmpIdx && c < headers.length; c++) {
        if (headers[c]) structure.ccS1Cols.push({ name: headers[c], index: c });
      }
      for (let c = s1EmpIdx + 1; c < headers.length; c++) {
        if (c !== s2EmpIdx && headers[c]) structure.ccS2Cols.push({ name: headers[c], index: c });
      }

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const dObj = parseDateValue(row[0]);
        const dStr = dObj ? formatDateStr(dObj) : (row[0] ? String(row[0]).trim() : '');
        const dTime = dObj ? dObj.getTime() : 0;
        
        // Section 1 Entry: Employee in Col F (s1EmpIdx)
        const emp1 = normalizeEmployeeName(row[s1EmpIdx]);
        if (emp1 !== '') {
          const filled = {};
          structure.ccS1Cols.forEach(col => {
            const val = row[col.index];
            if (val !== undefined && val !== null && String(val).trim() !== '') {
              filled[col.name] = true;
            }
          });
          
          allEntries.push({
            sheet: 'Close Cases',
            section: 'CC_S1',
            employee: emp1,
            date: row[0],
            dateStr: dStr,
            dateTime: dTime,
            cols: filled
          });
        }
        
        // Section 2 Entry: Employee in Col K (s2EmpIdx)
        const emp2 = normalizeEmployeeName(row[s2EmpIdx]);
        if (emp2 !== '') {
          const filled = {};
          structure.ccS2Cols.forEach(col => {
            const val = row[col.index];
            if (val !== undefined && val !== null && String(val).trim() !== '') {
              filled[col.name] = true;
            }
          });
          
          allEntries.push({
            sheet: 'Close Cases',
            section: 'CC_S2',
            employee: emp2,
            date: row[0],
            dateStr: dStr,
            dateTime: dTime,
            cols: filled
          });
        }
      }
    }
  }

  // 2. Process Forwarded Cases Sheet
  if (forwardedCasesSheet) {
    const data = forwardedCasesSheet.getDataRange().getValues();
    if (data.length > 1) {
      totalMasterRows += (data.length - 1);
      const headers = data[0].map(h => String(h || '').trim());
      
      const empCols = [];
      headers.forEach((h, idx) => {
        const lower = h.toLowerCase().replace(/[^a-z]/g, '');
        if (lower.includes('forwardedby') || lower.includes('forwardby')) {
          empCols.push(idx);
        }
      });
      const s1EmpIdx = empCols.length > 0 ? empCols[0] : 6;
      const s2EmpIdx = empCols.length > 1 ? empCols[1] : 12;
      
      for (let c = 1; c < s1EmpIdx && c < headers.length; c++) {
        if (headers[c]) structure.fcS1Cols.push({ name: headers[c], index: c });
      }
      for (let c = s1EmpIdx + 1; c < headers.length; c++) {
        if (c !== s2EmpIdx && headers[c]) structure.fcS2Cols.push({ name: headers[c], index: c });
      }

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const dObj = parseDateValue(row[0]);
        const dStr = dObj ? formatDateStr(dObj) : (row[0] ? String(row[0]).trim() : '');
        const dTime = dObj ? dObj.getTime() : 0;
        
        // Section 1 Entry: Employee in Col G (s1EmpIdx)
        const emp1 = normalizeEmployeeName(row[s1EmpIdx]);
        if (emp1 !== '') {
          const filled = {};
          structure.fcS1Cols.forEach(col => {
            const val = row[col.index];
            if (val !== undefined && val !== null && String(val).trim() !== '') {
              filled[col.name] = true;
            }
          });
          
          allEntries.push({
            sheet: 'Forwarded Cases',
            section: 'FC_S1',
            employee: emp1,
            date: row[0],
            dateStr: dStr,
            dateTime: dTime,
            cols: filled
          });
        }
        
        // Section 2 Entry: Employee in Col M (s2EmpIdx)
        const emp2 = normalizeEmployeeName(row[s2EmpIdx]);
        if (emp2 !== '') {
          const filled = {};
          structure.fcS2Cols.forEach(col => {
            const val = row[col.index];
            if (val !== undefined && val !== null && String(val).trim() !== '') {
              filled[col.name] = true;
            }
          });
          
          allEntries.push({
            sheet: 'Forwarded Cases',
            section: 'FC_S2',
            employee: emp2,
            date: row[0],
            dateStr: dStr,
            dateTime: dTime,
            cols: filled
          });
        }
      }
    }
  }

  // Extract all unique employees for the dropdown BEFORE filtering
  const allUniqueEmployees = new Set();
  allEntries.forEach(entry => {
    if (entry.employee) allUniqueEmployees.add(entry.employee);
  });
  const employeeNames = Array.from(allUniqueEmployees).sort();

  // Section-wise Detection of Unknown / Incomplete Activities
  let unknownEntries = [];
  
  // 1. Close Cases Incomplete Check
  if (closeCasesSheet) {
    const data = closeCasesSheet.getDataRange().getValues();
    if (data.length > 1) {
      const headers = data[0].map(h => String(h || '').trim());
      const empCols = [];
      headers.forEach((h, idx) => {
        const lower = h.toLowerCase().replace(/[^a-z]/g, '');
        if (lower.includes('closedby') || lower === 'closedby') empCols.push(idx);
      });
      const s1EmpIdx = empCols.length > 0 ? empCols[0] : 5;
      const s2EmpIdx = empCols.length > 1 ? empCols[1] : 10;

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rawDate = row[0];
        const hasDate = (rawDate !== undefined && rawDate !== null && String(rawDate).trim() !== '' && parseDateValue(rawDate) !== null);
        
        // Check Section 1
        let s1HasData = false;
        for (let c = 1; c < s1EmpIdx && c < row.length; c++) {
          if (row[c] && String(row[c]).trim() !== '') { s1HasData = true; break; }
        }
        const emp1 = normalizeEmployeeName(row[s1EmpIdx]);
        if (s1HasData && !emp1) {
          unknownEntries.push({
            sheet: 'Close Cases',
            section: 'Section 1 (Col A-F)',
            rowNumber: i + 1,
            reason: 'Missing Employee (Col F)',
            detail: 'Activity data present but CLOSED-BY is blank'
          });
        } else if (emp1 && !hasDate) {
          unknownEntries.push({
            sheet: 'Close Cases',
            section: 'Section 1 (Col A-F)',
            rowNumber: i + 1,
            reason: 'Missing Date (Col A)',
            detail: `Employee: ${emp1}`
          });
        }

        // Check Section 2
        let s2HasData = false;
        for (let c = s1EmpIdx + 1; c < headers.length && c < row.length; c++) {
          if (c !== s2EmpIdx && row[c] && String(row[c]).trim() !== '') { s2HasData = true; break; }
        }
        const emp2 = normalizeEmployeeName(row[s2EmpIdx]);
        if (s2HasData && !emp2) {
          unknownEntries.push({
            sheet: 'Close Cases',
            section: 'Section 2 (Col G-K)',
            rowNumber: i + 1,
            reason: 'Missing Employee (Col K)',
            detail: 'Activity data present but CLOSED BY is blank'
          });
        } else if (emp2 && !hasDate) {
          unknownEntries.push({
            sheet: 'Close Cases',
            section: 'Section 2 (Col G-K)',
            rowNumber: i + 1,
            reason: 'Missing Date (Col A)',
            detail: `Employee: ${emp2}`
          });
        }
      }
    }
  }

  // 2. Forwarded Cases Incomplete Check
  if (forwardedCasesSheet) {
    const data = forwardedCasesSheet.getDataRange().getValues();
    if (data.length > 1) {
      const headers = data[0].map(h => String(h || '').trim());
      const empCols = [];
      headers.forEach((h, idx) => {
        const lower = h.toLowerCase().replace(/[^a-z]/g, '');
        if (lower.includes('forwardedby') || lower.includes('forwardby')) empCols.push(idx);
      });
      const s1EmpIdx = empCols.length > 0 ? empCols[0] : 6;
      const s2EmpIdx = empCols.length > 1 ? empCols[1] : 12;

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rawDate = row[0];
        const hasDate = (rawDate !== undefined && rawDate !== null && String(rawDate).trim() !== '' && parseDateValue(rawDate) !== null);
        
        // Check Section 1
        let s1HasData = false;
        for (let c = 1; c < s1EmpIdx && c < row.length; c++) {
          if (row[c] && String(row[c]).trim() !== '') { s1HasData = true; break; }
        }
        const emp1 = normalizeEmployeeName(row[s1EmpIdx]);
        if (s1HasData && !emp1) {
          unknownEntries.push({
            sheet: 'Forwarded Cases',
            section: 'Section 1 (Col A-G)',
            rowNumber: i + 1,
            reason: 'Missing Employee (Col G)',
            detail: 'Activity data present but Forwarded-By is blank'
          });
        } else if (emp1 && !hasDate) {
          unknownEntries.push({
            sheet: 'Forwarded Cases',
            section: 'Section 1 (Col A-G)',
            rowNumber: i + 1,
            reason: 'Missing Date (Col A)',
            detail: `Employee: ${emp1}`
          });
        }

        // Check Section 2
        let s2HasData = false;
        for (let c = s1EmpIdx + 1; c < headers.length && c < row.length; c++) {
          if (c !== s2EmpIdx && row[c] && String(row[c]).trim() !== '') { s2HasData = true; break; }
        }
        const emp2 = normalizeEmployeeName(row[s2EmpIdx]);
        if (s2HasData && !emp2) {
          unknownEntries.push({
            sheet: 'Forwarded Cases',
            section: 'Section 2 (Col H-M)',
            rowNumber: i + 1,
            reason: 'Missing Employee (Col M)',
            detail: 'Activity data present but Forwarded by is blank'
          });
        } else if (emp2 && !hasDate) {
          unknownEntries.push({
            sheet: 'Forwarded Cases',
            section: 'Section 2 (Col H-M)',
            rowNumber: i + 1,
            reason: 'Missing Date (Col A)',
            detail: `Employee: ${emp2}`
          });
        }
      }
    }
  }

  // Apply Employee Filter (case-insensitive)
  let filteredEntries = allEntries;
  if (filterEmp && filterEmp !== 'All Employees') {
    const cleanFilterEmp = filterEmp.toLowerCase().trim();
    filteredEntries = filteredEntries.filter(entry => entry.employee.toLowerCase().trim() === cleanFilterEmp);
  }
  
  // Apply Activity / Section / Column Filter
  if (filterAct && filterAct !== 'All Activities') {
    if (filterAct === 'Close Cases') {
      filteredEntries = filteredEntries.filter(entry => entry.sheet === 'Close Cases');
    } else if (filterAct === 'Forwarded Cases') {
      filteredEntries = filteredEntries.filter(entry => entry.sheet === 'Forwarded Cases');
    } else if (filterAct === 'CC_S1') {
      filteredEntries = filteredEntries.filter(entry => entry.section === 'CC_S1');
    } else if (filterAct === 'CC_S2') {
      filteredEntries = filteredEntries.filter(entry => entry.section === 'CC_S2');
    } else if (filterAct === 'FC_S1') {
      filteredEntries = filteredEntries.filter(entry => entry.section === 'FC_S1');
    } else if (filterAct === 'FC_S2') {
      filteredEntries = filteredEntries.filter(entry => entry.section === 'FC_S2');
    } else if (filterAct.startsWith('CC_S1:')) {
      const colName = filterAct.replace('CC_S1:', '');
      filteredEntries = filteredEntries.filter(entry => entry.section === 'CC_S1' && entry.cols[colName]);
    } else if (filterAct.startsWith('CC_S2:')) {
      const colName = filterAct.replace('CC_S2:', '');
      filteredEntries = filteredEntries.filter(entry => entry.section === 'CC_S2' && entry.cols[colName]);
    } else if (filterAct.startsWith('FC_S1:')) {
      const colName = filterAct.replace('FC_S1:', '');
      filteredEntries = filteredEntries.filter(entry => entry.section === 'FC_S1' && entry.cols[colName]);
    } else if (filterAct.startsWith('FC_S2:')) {
      const colName = filterAct.replace('FC_S2:', '');
      filteredEntries = filteredEntries.filter(entry => entry.section === 'FC_S2' && entry.cols[colName]);
    }
  }

  // Apply Date From Filter (Column A)
  if (dateFrom) {
    const fromD = parseDateValue(dateFrom);
    if (fromD) {
      filteredEntries = filteredEntries.filter(entry => {
        const d = parseDateValue(entry.date);
        return d && d.getTime() >= fromD.getTime();
      });
    }
  }

  // Apply Date To Filter (Column A)
  if (dateTo) {
    const toD = parseDateValue(dateTo);
    if (toD) {
      filteredEntries = filteredEntries.filter(entry => {
        const d = parseDateValue(entry.date);
        return d && d.getTime() <= toD.getTime();
      });
    }
  }

  // Calculate metrics from filtered entries
  const uniqueFilteredEmployees = new Set();
  const employeeWorkload = {};
  
  filteredEntries.forEach(entry => {
    const emp = entry.employee || 'Unknown';
    uniqueFilteredEmployees.add(emp);
    
    if (!employeeWorkload[emp]) {
      employeeWorkload[emp] = 0;
    }
    employeeWorkload[emp]++;
  });

  // Calculate top performers
  let topPerformers = Object.keys(employeeWorkload).map(emp => {
    return {
      name: emp,
      workload: employeeWorkload[emp]
    };
  }).sort((a, b) => b.workload - a.workload);

  const topPerformerName = topPerformers.length > 0 ? topPerformers[0].name : 'N/A';
  
  // Calculate completion percentage
  const maxWorkload = topPerformers.length > 0 ? topPerformers[0].workload : 1;
  topPerformers = topPerformers.slice(0, 10).map(p => {
    return {
      name: p.name,
      workload: p.workload,
      completion: Math.round((p.workload / maxWorkload) * 100)
    };
  });

  // Calculate employee leaderboard rankings
  const totalFilteredCount = filteredEntries.length;
  const leaderboard = Object.keys(employeeWorkload).map(emp => {
    const count = employeeWorkload[emp];
    const pct = totalFilteredCount > 0 ? ((count / totalFilteredCount) * 100).toFixed(1) + '%' : '0.0%';
    return {
      name: emp,
      total: count,
      contribution: pct
    };
  }).sort((a, b) => b.total - a.total);

  leaderboard.forEach((item, idx) => {
    item.rank = idx + 1;
  });

  // Calculate complete employee profile drill-down data
  const employeeProfiles = {};
  const allCategoryNames = new Set();
  structure.ccS1Cols.forEach(c => allCategoryNames.add(c.name));
  structure.ccS2Cols.forEach(c => allCategoryNames.add(c.name));
  structure.fcS1Cols.forEach(c => allCategoryNames.add(c.name));
  structure.fcS2Cols.forEach(c => allCategoryNames.add(c.name));

  const empDailyMaps = {};
  const empActMaps = {};

  allEntries.forEach(e => {
    const emp = e.employee;
    if (!emp) return;

    if (!empDailyMaps[emp]) empDailyMaps[emp] = {};
    if (!empActMaps[emp]) empActMaps[emp] = {};

    if (e.dateStr) {
      if (!empDailyMaps[emp][e.dateStr]) {
        empDailyMaps[emp][e.dateStr] = { label: e.dateStr, time: e.dateTime, count: 0 };
      }
      empDailyMaps[emp][e.dateStr].count++;
    }

    if (e.cols) {
      Object.keys(e.cols).forEach(colName => {
        empActMaps[emp][colName] = (empActMaps[emp][colName] || 0) + 1;
      });
    }
  });

  leaderboard.forEach(item => {
    const emp = item.name;
    const totalCount = item.total;
    const dailyMap = empDailyMaps[emp] || {};
    const actMap = empActMaps[emp] || {};

    const dailyTrend = Object.values(dailyMap).sort((a, b) => a.time - b.time);

    const breakdown = Array.from(allCategoryNames).map(act => {
      const count = actMap[act] || 0;
      const pct = totalCount > 0 ? ((count / totalCount) * 100).toFixed(1) + '%' : '0.0%';
      return {
        activity: act,
        count: count,
        share: pct
      };
    }).sort((a, b) => b.count - a.count);

    employeeProfiles[emp] = {
      employee: emp,
      rank: '#' + item.rank,
      totalActivities: totalCount,
      dailyTrend: dailyTrend,
      breakdown: breakdown
    };
  });

  // 1. Weekday Heatmap Calculation
  const weekdayCounts = { 0: 0, 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0 };
  const teamActCounts = {};

  filteredEntries.forEach(entry => {
    const d = parseDateValue(entry.date);
    if (d) {
      const day = d.getDay();
      weekdayCounts[day] = (weekdayCounts[day] || 0) + 1;
    }
    if (entry.cols) {
      Object.keys(entry.cols).forEach(colName => {
        teamActCounts[colName] = (teamActCounts[colName] || 0) + 1;
      });
    }
  });

  const weekdayNames = ['SUN', 'MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT'];
  const maxDayCount = Math.max(...Object.values(weekdayCounts), 1);
  const weekdayHeatmap = weekdayNames.map((name, idx) => {
    const cnt = weekdayCounts[idx] || 0;
    return {
      day: name,
      dayIndex: idx,
      count: cnt,
      intensity: maxDayCount > 0 ? (cnt / maxDayCount) : 0
    };
  });

  // 2. Bottom 5 Performers
  const bottomPerformers = leaderboard.length > 5 ? leaderboard.slice(-5).reverse() : [...leaderboard].reverse();

  // 3. Activity Breakdown Shares (Team wide)
  const totalTeamAct = filteredEntries.length;
  const activityBreakdown = Object.keys(teamActCounts).map(colName => {
    const cnt = teamActCounts[colName];
    const pct = totalTeamAct > 0 ? ((cnt / totalTeamAct) * 100).toFixed(1) + '%' : '0.0%';
    return {
      activity: colName,
      count: cnt,
      share: pct,
      numericShare: totalTeamAct > 0 ? (cnt / totalTeamAct) * 100 : 0
    };
  }).sort((a, b) => b.count - a.count);

  return {
    employees: uniqueFilteredEmployees.size,
    employeeNames: employeeNames,
    structure: {
      ccS1Cols: structure.ccS1Cols.map(c => c.name),
      ccS2Cols: structure.ccS2Cols.map(c => c.name),
      fcS1Cols: structure.fcS1Cols.map(c => c.name),
      fcS2Cols: structure.fcS2Cols.map(c => c.name)
    },
    activities: filteredEntries.length,
    topPerformer: topPerformerName,
    unknownActivities: unknownEntries.length,
    unknownEntries: unknownEntries,
    masterRows: totalMasterRows,
    topPerformers: topPerformers,
    leaderboard: leaderboard,
    weekdayHeatmap: weekdayHeatmap,
    bottomPerformers: bottomPerformers,
    activityBreakdown: activityBreakdown,
    dailyTrendTotal: filteredEntries.length
  };
}

/**
 * Returns drill-down metrics for a specific employee
 */
function getEmployeeProfileData(targetEmp = '', actFilter = 'All Activities') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const closeCasesSheet = getSheetByNameTrimmed(ss, "Close Cases");
  const forwardedCasesSheet = getSheetByNameTrimmed(ss, "Forwarded cases");

  const structure = {
    ccS1Cols: [],
    ccS2Cols: [],
    fcS1Cols: [],
    fcS2Cols: []
  };

  const allEmployeeCounts = {};
  const empDailyCounts = {};
  const actCounts = {};
  const allCategoryNames = new Set();

  // Helper to process row
  function processRow(sheetName, section, empVal, dateVal, rowCols, headerCols) {
    const emp = normalizeEmployeeName(empVal);
    if (!emp) return;

    allEmployeeCounts[emp] = (allEmployeeCounts[emp] || 0) + 1;

    const isTarget = targetEmp && (emp.toLowerCase() === targetEmp.toLowerCase());
    if (!isTarget && targetEmp !== '') return;

    if (actFilter && actFilter !== 'All Activities') {
      if (actFilter === 'Close Cases' && sheetName !== 'Close Cases') return;
      if (actFilter === 'Forwarded Cases' && sheetName !== 'Forwarded Cases') return;
      if (actFilter === 'CC_S1' && section !== 'CC_S1') return;
      if (actFilter === 'CC_S2' && section !== 'CC_S2') return;
      if (actFilter === 'FC_S1' && section !== 'FC_S1') return;
      if (actFilter === 'FC_S2' && section !== 'FC_S2') return;
    }

    const dObj = parseDateValue(dateVal);
    if (dObj) {
      const dStr = formatDateStr(dObj);
      const dTime = dObj.getTime();
      if (!empDailyCounts[dStr]) {
        empDailyCounts[dStr] = { label: dStr, time: dTime, count: 0 };
      }
      empDailyCounts[dStr].count++;
    }

    headerCols.forEach(col => {
      const val = rowCols[col.index];
      if (val !== undefined && val !== null && String(val).trim() !== '') {
        actCounts[col.name] = (actCounts[col.name] || 0) + 1;
      }
    });
  }

  // 1. Close Cases
  if (closeCasesSheet) {
    const data = closeCasesSheet.getDataRange().getValues();
    if (data.length > 1) {
      const headers = data[0].map(h => String(h || '').trim());
      const empCols = [];
      headers.forEach((h, idx) => {
        const lower = h.toLowerCase().replace(/[^a-z]/g, '');
        if (lower.includes('closedby') || lower === 'closedby') empCols.push(idx);
      });
      const s1EmpIdx = empCols.length > 0 ? empCols[0] : 5;
      const s2EmpIdx = empCols.length > 1 ? empCols[1] : 10;

      for (let c = 1; c < s1EmpIdx && c < headers.length; c++) {
        if (headers[c]) {
          structure.ccS1Cols.push({ name: headers[c], index: c });
          allCategoryNames.add(headers[c]);
        }
      }
      for (let c = s1EmpIdx + 1; c < headers.length; c++) {
        if (c !== s2EmpIdx && headers[c]) {
          structure.ccS2Cols.push({ name: headers[c], index: c });
          allCategoryNames.add(headers[c]);
        }
      }

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        processRow('Close Cases', 'CC_S1', row[s1EmpIdx], row[0], row, structure.ccS1Cols);
        processRow('Close Cases', 'CC_S2', row[s2EmpIdx], row[0], row, structure.ccS2Cols);
      }
    }
  }

  // 2. Forwarded Cases
  if (forwardedCasesSheet) {
    const data = forwardedCasesSheet.getDataRange().getValues();
    if (data.length > 1) {
      const headers = data[0].map(h => String(h || '').trim());
      const empCols = [];
      headers.forEach((h, idx) => {
        const lower = h.toLowerCase().replace(/[^a-z]/g, '');
        if (lower.includes('forwardedby') || lower.includes('forwardby')) empCols.push(idx);
      });
      const s1EmpIdx = empCols.length > 0 ? empCols[0] : 6;
      const s2EmpIdx = empCols.length > 1 ? empCols[1] : 12;

      for (let c = 1; c < s1EmpIdx && c < headers.length; c++) {
        if (headers[c]) {
          structure.fcS1Cols.push({ name: headers[c], index: c });
          allCategoryNames.add(headers[c]);
        }
      }
      for (let c = s1EmpIdx + 1; c < headers.length; c++) {
        if (c !== s2EmpIdx && headers[c]) {
          structure.fcS2Cols.push({ name: headers[c], index: c });
          allCategoryNames.add(headers[c]);
        }
      }

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        processRow('Forwarded Cases', 'FC_S1', row[s1EmpIdx], row[0], row, structure.fcS1Cols);
        processRow('Forwarded Cases', 'FC_S2', row[s2EmpIdx], row[0], row, structure.fcS2Cols);
      }
    }
  }

  const sortedEmployees = Object.keys(allEmployeeCounts).sort((a, b) => allEmployeeCounts[b] - allEmployeeCounts[a]);
  const employeeNames = Object.keys(allEmployeeCounts).sort();

  let currentTarget = targetEmp;
  if (!currentTarget || !allEmployeeCounts[currentTarget]) {
    currentTarget = sortedEmployees.length > 0 ? sortedEmployees[0] : '';
    if (!targetEmp && currentTarget) {
      return getEmployeeProfileData(currentTarget, actFilter);
    }
  }

  const rankIdx = sortedEmployees.findIndex(e => e.toLowerCase() === currentTarget.toLowerCase());
  const rankStr = rankIdx !== -1 ? `#${rankIdx + 1}` : '#-';
  const totalCount = allEmployeeCounts[currentTarget] || 0;

  const dailyTrend = Object.values(empDailyCounts).sort((a, b) => a.time - b.time);

  allCategoryNames.forEach(name => {
    if (!actCounts[name]) actCounts[name] = 0;
  });

  const breakdown = Object.keys(actCounts).map(act => {
    const count = actCounts[act];
    const pct = totalCount > 0 ? ((count / totalCount) * 100).toFixed(1) + '%' : '0.0%';
    return {
      activity: act,
      count: count,
      share: pct
    };
  }).sort((a, b) => b.count - a.count);

  return {
    employee: currentTarget,
    rank: rankStr,
    totalActivities: totalCount,
    dailyTrend: dailyTrend,
    breakdown: breakdown,
    employeeList: employeeNames
  };
}

/**
 * Generates direct download URL for the spreadsheet as XLSX (excluding Config sheet)
 */
function getSpreadsheetExportUrl() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  // If Config exists, create a clean temporary export workbook with only data sheets
  const configSheet = ss.getSheetByName("Config");
  if (configSheet && sheets.length > 1) {
    try {
      const tempSS = SpreadsheetApp.create('Master_Report_' + Utilities.formatDate(new Date(), 'GMT', 'yyyyMMdd'));
      sheets.forEach(s => {
        if (s.getName().toLowerCase().trim() !== 'config') {
          s.copyTo(tempSS).setName(s.getName());
        }
      });
      const defSheet = tempSS.getSheetByName('Sheet1');
      if (defSheet && tempSS.getSheets().length > 1) {
        tempSS.deleteSheet(defSheet);
      }
      DriveApp.getFileById(tempSS.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return 'https://docs.google.com/spreadsheets/d/' + tempSS.getId() + '/export?format=xlsx';
    } catch (e) {
      // Fallback to direct export URL
      return 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx';
    }
  }
  
  return 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx';
}

/**
 * Returns all data sheets (excluding Config) in one ZIP package, combined CSV, and individual files
 */
function getMasterReportCsvData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const files = [];
  const blobs = [];
  const combinedLines = [];
  
  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (name.toLowerCase().trim() === 'config') return; // Exclude Config sheet
    
    const data = sheet.getDataRange().getValues();
    if (data.length > 0) {
      combinedLines.push(`--- SHEET: ${name} ---`);
      
      const csvRows = data.map(row => {
        return row.map(cell => {
          let str = (cell === null || cell === undefined) ? '' : String(cell);
          if (str.includes('"') || str.includes(',') || str.includes('\n') || str.includes('\r')) {
            str = '"' + str.replace(/"/g, '""') + '"';
          }
          return str;
        }).join(',');
      });
      
      const filename = name.trim() + '.csv';
      const content = csvRows.join('\r\n');
      
      files.push({
        filename: filename,
        content: content
      });
      
      combinedLines.push(content);
      combinedLines.push('');
      
      blobs.push(Utilities.newBlob(content, 'text/csv', filename));
    }
  });
  
  let zipBase64 = null;
  if (blobs.length > 0) {
    const zipBlob = Utilities.zip(blobs, 'Master_Report_All_Sheets.zip');
    zipBase64 = Utilities.base64Encode(zipBlob.getBytes());
  }
  
  return {
    files: files,
    zipBase64: zipBase64,
    combinedCsv: combinedLines.join('\r\n')
  };
}

/**
 * Returns date-wise aggregated matrix for the Monthly Report
 */
function getMonthlyReportData(dateFrom = '', dateTo = '') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const closeCasesSheet = getSheetByNameTrimmed(ss, "Close Cases");
  const forwardedCasesSheet = getSheetByNameTrimmed(ss, "Forwarded cases");

  const fromD = dateFrom ? parseDateValue(dateFrom) : null;
  const toD = dateTo ? parseDateValue(dateTo) : null;

  const activityColumns = [];
  const dateMap = {};

  function processSheet(sheet, isForwarded) {
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    const headers = data[0].map(h => String(h || '').trim());
    const empCols = [];
    headers.forEach((h, idx) => {
      const lower = h.toLowerCase().replace(/[^a-z]/g, '');
      if (isForwarded) {
        if (lower.includes('forwardedby') || lower.includes('forwardby')) empCols.push(idx);
      } else {
        if (lower.includes('closedby') || lower === 'closedby') empCols.push(idx);
      }
    });

    const s1EmpIdx = empCols.length > 0 ? empCols[0] : (isForwarded ? 6 : 5);
    const s2EmpIdx = empCols.length > 1 ? empCols[1] : (isForwarded ? 12 : 10);

    const s1Cols = [];
    for (let c = 1; c < s1EmpIdx && c < headers.length; c++) {
      if (headers[c] && headers[c] !== '') {
        const colName = headers[c];
        s1Cols.push({ index: c, name: colName });
        if (!activityColumns.includes(colName)) activityColumns.push(colName);
      }
    }

    const s2Cols = [];
    for (let c = s1EmpIdx + 1; c < headers.length; c++) {
      if (c !== s2EmpIdx && headers[c] && headers[c] !== '') {
        const colName = headers[c];
        s2Cols.push({ index: c, name: colName });
        if (!activityColumns.includes(colName)) activityColumns.push(colName);
      }
    }

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rawDate = row[0];
      const dObj = parseDateValue(rawDate);
      if (!dObj) continue;

      if (fromD && dObj.getTime() < fromD.getTime()) continue;
      if (toD && dObj.getTime() > toD.getTime()) continue;

      const dStr = formatDateStr(dObj);
      const dTime = dObj.getTime();

      if (!dateMap[dStr]) {
        dateMap[dStr] = { date: dStr, time: dTime, counts: {} };
      }

      s1Cols.forEach(col => {
        const val = row[col.index];
        if (val !== undefined && val !== null && String(val).trim() !== '') {
          dateMap[dStr].counts[col.name] = (dateMap[dStr].counts[col.name] || 0) + 1;
        }
      });

      s2Cols.forEach(col => {
        const val = row[col.index];
        if (val !== undefined && val !== null && String(val).trim() !== '') {
          dateMap[dStr].counts[col.name] = (dateMap[dStr].counts[col.name] || 0) + 1;
        }
      });
    }
  }

  processSheet(closeCasesSheet, false);
  processSheet(forwardedCasesSheet, true);

  const rows = Object.values(dateMap).sort((a, b) => a.time - b.time);

  return {
    columns: activityColumns,
    rows: rows
  };
}

/**
 * Returns employee-wise aggregated matrix for the Custom Date Report
 */
function getCustomDateReportData(dateFrom = '', dateTo = '') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const closeCasesSheet = getSheetByNameTrimmed(ss, "Close Cases");
  const forwardedCasesSheet = getSheetByNameTrimmed(ss, "Forwarded cases");

  const fromD = dateFrom ? parseDateValue(dateFrom) : null;
  const toD = dateTo ? parseDateValue(dateTo) : null;

  const activityColumns = [];
  const empMap = {};

  function processSheet(sheet, isForwarded) {
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    const headers = data[0].map(h => String(h || '').trim());
    const empCols = [];
    headers.forEach((h, idx) => {
      const lower = h.toLowerCase().replace(/[^a-z]/g, '');
      if (isForwarded) {
        if (lower.includes('forwardedby') || lower.includes('forwardby')) empCols.push(idx);
      } else {
        if (lower.includes('closedby') || lower === 'closedby') empCols.push(idx);
      }
    });

    const s1EmpIdx = empCols.length > 0 ? empCols[0] : (isForwarded ? 6 : 5);
    const s2EmpIdx = empCols.length > 1 ? empCols[1] : (isForwarded ? 12 : 10);

    const s1Cols = [];
    for (let c = 1; c < s1EmpIdx && c < headers.length; c++) {
      if (headers[c] && headers[c] !== '') {
        const colName = headers[c];
        s1Cols.push({ index: c, name: colName });
        if (!activityColumns.includes(colName)) activityColumns.push(colName);
      }
    }

    const s2Cols = [];
    for (let c = s1EmpIdx + 1; c < headers.length; c++) {
      if (c !== s2EmpIdx && headers[c] && headers[c] !== '') {
        const colName = headers[c];
        s2Cols.push({ index: c, name: colName });
        if (!activityColumns.includes(colName)) activityColumns.push(colName);
      }
    }

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rawDate = row[0];
      const dObj = parseDateValue(rawDate);
      
      if (fromD && (!dObj || dObj.getTime() < fromD.getTime())) continue;
      if (toD && (!dObj || dObj.getTime() > toD.getTime())) continue;

      const emp1 = normalizeEmployeeName(row[s1EmpIdx]);
      if (emp1) {
        if (!empMap[emp1]) empMap[emp1] = { name: emp1, counts: {} };
        s1Cols.forEach(col => {
          const val = row[col.index];
          if (val !== undefined && val !== null && String(val).trim() !== '') {
            empMap[emp1].counts[col.name] = (empMap[emp1].counts[col.name] || 0) + 1;
          }
        });
      }

      const emp2 = normalizeEmployeeName(row[s2EmpIdx]);
      if (emp2) {
        if (!empMap[emp2]) empMap[emp2] = { name: emp2, counts: {} };
        s2Cols.forEach(col => {
          const val = row[col.index];
          if (val !== undefined && val !== null && String(val).trim() !== '') {
            empMap[emp2].counts[col.name] = (empMap[emp2].counts[col.name] || 0) + 1;
          }
        });
      }
    }
  }

  processSheet(closeCasesSheet, false);
  processSheet(forwardedCasesSheet, true);

  const rows = Object.values(empMap).sort((a, b) => a.name.localeCompare(b.name));

  return {
    columns: activityColumns,
    rows: rows
  };
}

/**
 * Helper to safely open source Google Spreadsheet by URL or Spreadsheet ID
 */
function openSourceSpreadsheet(sourceInput) {
  if (!sourceInput) {
    throw new Error("Please enter a valid Source Spreadsheet URL or ID in Config!B2");
  }
  const input = String(sourceInput).trim();
  if (input.startsWith("http://") || input.startsWith("https://")) {
    return SpreadsheetApp.openByUrl(input);
  } else {
    return SpreadsheetApp.openById(input);
  }
}

/**
 * Ensures the 'Config' sheet exists with required baseline template structure
 */
function ensureConfigSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  let config = ss.getSheetByName("Config");
  if (!config) {
    config = ss.insertSheet("Config");
    config.getRange("A1:B5").setValues([
      ["CONFIG FIELD", "VALUE"],
      ["Source Spreadsheet URL", ""],
      ["Last Sync Time", "Never"],
      ["Sync Status", "Ready"],
      ["Total Rows Imported", 0]
    ]);
    config.getRange("A8:D10").setValues([
      ["Import Type", "Source Tab", "Destination Tab", "Clear Before Import"],
      ["Close Cases", "August Close Cases", "Close Cases", true],
      ["Forwarded Cases", "August Forwarded cases", "Forwarded cases", true]
    ]);
    config.getRange("A15:D15").setValues([
      ["Sync Date Time", "Close Rows", "Forwarded Rows", "Status"]
    ]);
    config.getRange("A1:B1").setFontWeight("bold").setBackground("#e0e7ff");
    config.getRange("A8:D8").setFontWeight("bold").setBackground("#e0e7ff");
    config.getRange("A15:D15").setFontWeight("bold").setBackground("#e0e7ff");
  }
  return config;
}

/**
 * Retrieves current sync configuration, mappings, and historical audit logs from Config sheet
 */
function getSyncStatusData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ensureConfigSheet(ss);

  const sourceUrl = String(config.getRange("B2").getDisplayValue() || '').trim();
  const lastSyncRaw = config.getRange("B3").getValue();
  const lastSync = (lastSyncRaw instanceof Date) ? Utilities.formatDate(lastSyncRaw, Session.getScriptTimeZone() || 'GMT', "dd-MMM-yyyy hh:mm:ss a") : String(lastSyncRaw || 'Never');
  const status = String(config.getRange("B4").getDisplayValue() || 'Ready').trim();
  const totalImported = config.getRange("B5").getValue() || 0;

  // Read mappings from A9:D14
  const mappingVals = config.getRange("A9:D14").getValues();
  const mappings = [];
  mappingVals.forEach(m => {
    if (m[0] || m[1] || m[2]) {
      mappings.push({
        importType: String(m[0] || ''),
        sourceTab: String(m[1] || ''),
        destinationTab: String(m[2] || ''),
        clearBeforeImport: (m[3] === true || String(m[3]).toUpperCase() === 'TRUE')
      });
    }
  });

  // Read historical sync logs from row 16 onwards
  const lastRow = config.getLastRow();
  const logs = [];
  if (lastRow >= 16) {
    const logData = config.getRange(16, 1, lastRow - 15, 4).getValues();
    logData.forEach(r => {
      if (r[0] || r[3]) {
        const timeVal = (r[0] instanceof Date) ? Utilities.formatDate(r[0], Session.getScriptTimeZone() || 'GMT', "dd-MMM-yyyy hh:mm:ss a") : String(r[0] || '-');
        logs.push({
          timestamp: timeVal,
          closeRows: r[1] || 0,
          forwardedRows: r[2] || 0,
          totalRows: (Number(r[1]) || 0) + (Number(r[2]) || 0),
          status: String(r[3] || 'Success')
        });
      }
    });
  }

  // Attempt to fetch source tabs if URL is available
  let sourceTabs = [];
  let sourceConnected = false;
  if (sourceUrl) {
    try {
      const sourceSS = openSourceSpreadsheet(sourceUrl);
      sourceTabs = sourceSS.getSheets().map(s => s.getName());
      sourceConnected = true;
    } catch (e) {
      sourceConnected = false;
    }
  }

  return {
    sourceUrl: sourceUrl,
    lastSync: lastSync,
    status: status,
    totalImported: totalImported,
    mappings: mappings,
    sourceTabs: sourceTabs,
    sourceConnected: sourceConnected,
    schedule: getScheduleConfiguration(),
    logs: logs.reverse()
  };
}

/**
 * Refreshes available sheet tab dropdowns from the source spreadsheet
 */
function executeRefreshSourceTabs(customSourceUrl = '') {
  return executeCoreRefreshSourceTabs(customSourceUrl);
}

/**
 * Executes full data sync according to Code.js logic
 */
function executeDataSync(customSourceUrl = '') {
  return executeCoreDataSync(customSourceUrl);
}

/**
 * Saves source URL and tab mappings back to Config sheet
 */
function saveSyncMappings(sourceUrl, mappingList) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ensureConfigSheet(ss);

  if (sourceUrl !== undefined && sourceUrl !== null) {
    config.getRange("B2").setValue(String(sourceUrl).trim());
  }

  if (Array.isArray(mappingList) && mappingList.length > 0) {
    mappingList.forEach((m, idx) => {
      const row = 9 + idx;
      if (row <= 14) {
        if (m.importType !== undefined) config.getRange(row, 1).setValue(m.importType);
        if (m.sourceTab !== undefined) config.getRange(row, 2).setValue(m.sourceTab);
        if (m.destinationTab !== undefined) config.getRange(row, 3).setValue(m.destinationTab);
        if (m.clearBeforeImport !== undefined) config.getRange(row, 4).setValue(m.clearBeforeImport);
      }
    });
  }

  return { success: true };
}

/**
 * Retrieves background schedule config
 */
function getAutoSyncSchedule() {
  return getScheduleConfiguration();
}

/**
 * Configures background schedule trigger
 */
function saveAutoSyncSchedule(isEnabled, scheduleType, timeValue) {
  return saveScheduleConfiguration(isEnabled, scheduleType, timeValue);
}

/**
 * Deactivates background schedule triggers
 */
function deleteAutoSyncSchedule() {
  deleteSyncTriggers(false);
  return getScheduleConfiguration();
}

// ============================================================================
// WEBAPP EMAIL SCHEDULER CONTROLLER ENDPOINTS
// ============================================================================

/**
 * Reads all active email schedules and dispatcher status
 */
function getEmailSchedulerData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureEmailScheduleSheet(ss);
  const data = sheet.getDataRange().getValues();
  const schedules = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0]) {
      schedules.push({
        id: String(row[0] || `SCH-${3000 + i}`),
        reportType: String(row[1] || "Yesterday's Master Report"),
        formatTemplate: String(row[2] || "Raw Master Report (All columns)"),
        customSubject: String(row[3] || "Review Daily reports (Yesterday)"),
        recipients: String(row[4] || ""),
        frequency: String(row[5] || "Daily"),
        targetHour: String(row[6] || "10:00 AM"),
        isActive: (row[7] === true || String(row[7]).toUpperCase() === "TRUE"),
        lastSent: String(row[8] || "-"),
        status: String(row[9] || "Pending"),
        createdAt: String(row[10] || "")
      });
    }
  }

  // Check if hourly dispatcher trigger is running
  const triggers = ScriptApp.getProjectTriggers();
  const isDispatcherActive = triggers.some(t => t.getHandlerFunction() === "processScheduledEmailDispatcher");

  return {
    schedules: schedules,
    isDispatcherActive: isDispatcherActive
  };
}

/**
 * Saves or updates an email schedule
 */
function saveEmailSchedule(sched) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureEmailScheduleSheet(ss);
  const data = sheet.getDataRange().getValues();
  const timeZone = Session.getScriptTimeZone() || 'Asia/Kolkata';

  let targetRow = -1;
  const schedId = sched.id || `SCH-${Math.floor(1000 + Math.random() * 9000)}`;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(schedId)) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow === -1) {
    targetRow = sheet.getLastRow() + 1;
  }

  sheet.getRange(targetRow, 1, 1, 11).setValues([[
    schedId,
    sched.reportType || "Yesterday's Master Report",
    sched.formatTemplate || "Raw Master Report (All columns)",
    sched.customSubject || "",
    sched.recipients || "",
    sched.frequency || "Daily",
    sched.targetHour || "9:00 AM",
    (sched.isActive !== undefined ? sched.isActive : true),
    sched.lastSent || "-",
    sched.status || "Pending",
    sched.createdAt || Utilities.formatDate(new Date(), timeZone, "M/d/yyyy, h:mm:ss a")
  ]]);

  // Ensure hourly dispatcher is initialized if schedule is active
  if (sched.isActive) {
    const triggers = ScriptApp.getProjectTriggers();
    if (!triggers.some(t => t.getHandlerFunction() === "processScheduledEmailDispatcher")) {
      setupEmailDispatcherTrigger();
    }
  }

  return getEmailSchedulerData();
}

/**
 * Deletes an email schedule
 */
function deleteEmailSchedule(scheduleId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureEmailScheduleSheet(ss);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(scheduleId)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }

  return getEmailSchedulerData();
}

/**
 * Immediately dispatches an email schedule test run
 */
function executeSendEmailScheduleNow(sched) {
  return sendScheduledReportEmail(sched);
}

/**
 * Toggles hourly dispatcher trigger
 */
function toggleEmailDispatcher(enable) {
  if (enable) {
    return setupEmailDispatcherTrigger();
  } else {
    return deleteEmailDispatcherTrigger();
  }
}

// ============================================================================
// WEBAPP ACCESS CONTROL CONTROLLER ENDPOINTS
// ============================================================================

/**
 * Reads all registered access control users
 */
function getAccessControlList() {
  return getAccessUsersData();
}

/**
 * Saves or updates an access control user
 */
function saveAccessControlUser(user) {
  return saveAccessUser(user);
}

/**
 * Deletes an access control user
 */
function deleteAccessControlUser(email) {
  return deleteAccessUser(email);
}

/**
 * Authenticates user login
 */
function authenticateLogin(email, password) {
  return validateUserLogin(email, password);
}
