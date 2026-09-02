/**
 * ============================================================================
 * RECON TEAM CID MASTER REPORTING - DATA IMPORT & SYNCHRONIZATION ENGINE
 * ============================================================================
 * Handles live synchronization between operational source Google Spreadsheets
 * and Master Report destination sheets based on dynamic Config tab mappings.
 */

/**
 * Creates custom menu in Google Sheets UI upon opening workbook
 */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("Master Import")
      .addItem("⚡ Sync Data Now", "syncData")
      .addItem("🔄 Refresh Source Tabs", "refreshSourceTabs")
      .addSeparator()
      .addItem("📧 Process Email Schedules Now", "processScheduledEmailDispatcher")
      .addItem("📊 Open Dashboard WebApp", "openDashboardWebApp")
      .addSeparator()
      .addItem("⏰ Setup Daily Auto-Sync (6 AM)", "setupDailySyncTrigger")
      .addItem("🚫 Remove Auto-Sync Triggers", "deleteSyncTriggers")
      .addToUi();
  } catch (e) {
    // getUi() may not be available in certain execution contexts
  }
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
 * Core function to refresh source tab dropdowns and validation in Config sheet
 */
function executeCoreRefreshSourceTabs(customSourceUrl) {
  const masterSS = SpreadsheetApp.getActiveSpreadsheet();
  const config = ensureConfigSheet(masterSS);

  if (customSourceUrl) {
    config.getRange("B2").setValue(String(customSourceUrl).trim());
  }

  const sourceUrl = customSourceUrl ? String(customSourceUrl).trim() : String(config.getRange("B2").getDisplayValue() || '').trim();
  if (!sourceUrl) {
    throw new Error("Please enter Source Spreadsheet URL in Config!B2");
  }

  const sourceSS = openSourceSpreadsheet(sourceUrl);
  const sheetNames = sourceSS.getSheets().map(sheet => sheet.getName());

  if (sheetNames.length > 0) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(sheetNames, true)
      .setAllowInvalid(false)
      .build();

    // Apply data validation to mapping source tab cells (B9:B14)
    config.getRange("B9:B14").setDataValidation(rule);
  }

  return {
    success: true,
    tabs: sheetNames,
    count: sheetNames.length
  };
}

/**
 * Core function to execute data synchronization between Source and Master sheets
 */
function executeCoreDataSync(customSourceUrl) {
  const masterSS = SpreadsheetApp.getActiveSpreadsheet();
  const config = ensureConfigSheet(masterSS);

  if (customSourceUrl) {
    config.getRange("B2").setValue(String(customSourceUrl).trim());
  }

  const sourceUrl = String(config.getRange("B2").getDisplayValue() || '').trim();
  if (!sourceUrl) {
    throw new Error("Source Spreadsheet URL is missing in Config!B2");
  }

  const syncTime = new Date();

  try {
    const sourceSS = openSourceSpreadsheet(sourceUrl);

    // Read mapping rows (A9:D14)
    const mappings = config.getRange("A9:D14").getValues();

    let totalImported = 0;
    let closeRows = 0;
    let forwardedRows = 0;
    let processedTabs = 0;

    mappings.forEach(mapping => {
      const importType = String(mapping[0] || '').trim();
      const sourceTab = String(mapping[1] || '').trim();
      const destinationTab = String(mapping[2] || '').trim();
      const clearBeforeImport = mapping[3];

      if (!sourceTab || !destinationTab) return;

      const sourceSheet = sourceSS.getSheetByName(sourceTab);
      let destinationSheet = masterSS.getSheetByName(destinationTab);

      if (!sourceSheet) {
        throw new Error(`Source tab "${sourceTab}" not found in source spreadsheet.`);
      }

      // Create destination sheet if it doesn't exist
      if (!destinationSheet) {
        destinationSheet = masterSS.insertSheet(destinationTab);
      }

      const sourceData = sourceSheet.getDataRange().getValues();
      if (!sourceData || sourceData.length <= 1) {
        // Source has only headers or is empty
        return;
      }

      // Copy headers to destination sheet if destination sheet is currently completely empty
      if (destinationSheet.getLastRow() === 0 && sourceData.length > 0) {
        const headerRow = [sourceData[0]];
        destinationSheet.getRange(1, 1, 1, headerRow[0].length).setValues(headerRow);
      }

      // Clear existing data but preserve row 1 headers
      const shouldClear = (clearBeforeImport === true || String(clearBeforeImport).toUpperCase() === "TRUE");
      if (shouldClear) {
        const lastRow = destinationSheet.getLastRow();
        if (lastRow > 1) {
          destinationSheet.getRange(2, 1, lastRow - 1, destinationSheet.getMaxColumns()).clearContent();
        }
      }

      // Extract rows (skip header row 0)
      const rows = sourceData.slice(1).filter(r => r.some(cell => cell !== '' && cell !== null && cell !== undefined));
      if (rows.length === 0) return;

      const reqCols = rows[0].length;
      const maxCols = destinationSheet.getMaxColumns();
      if (maxCols < reqCols) {
        destinationSheet.insertColumnsAfter(maxCols, reqCols - maxCols);
      }

      const destStartRow = destinationSheet.getLastRow() + 1;
      destinationSheet.getRange(destStartRow, 1, rows.length, reqCols).setValues(rows);

      totalImported += rows.length;
      processedTabs++;

      const lowerType = importType.toLowerCase();
      if (lowerType.includes("close")) {
        closeRows += rows.length;
      } else if (lowerType.includes("forward")) {
        forwardedRows += rows.length;
      }
    });

    // Update status cells in Config sheet
    config.getRange("B3").setValue(syncTime);
    config.getRange("B4").setValue("Success");
    config.getRange("B5").setValue(totalImported);

    // Sync History Audit Log (from row 16 onwards)
    const logRow = Math.max(config.getLastRow() + 1, 16);
    config.getRange(logRow, 1, 1, 4).setValues([[
      syncTime,
      closeRows,
      forwardedRows,
      "Success"
    ]]);

    const formattedTime = Utilities.formatDate(syncTime, Session.getScriptTimeZone() || 'GMT', "dd-MMM-yyyy hh:mm:ss a");

    return {
      success: true,
      closeRows: closeRows,
      forwardedRows: forwardedRows,
      totalImported: totalImported,
      processedTabs: processedTabs,
      timestamp: formattedTime
    };

  } catch (error) {
    config.getRange("B3").setValue(syncTime);
    config.getRange("B4").setValue("Failed");

    const logRow = Math.max(config.getLastRow() + 1, 16);
    config.getRange(logRow, 1, 1, 4).setValues([[
      syncTime,
      0,
      0,
      "Failed: " + error.message
    ]]);

    throw new Error(error.message);
  }
}

/**
 * Interactive UI Handler: Refresh source tab dropdowns in Sheets UI
 */
function refreshSourceTabs() {
  try {
    const result = executeCoreRefreshSourceTabs();
    SpreadsheetApp.getUi().alert(
      "Source Tabs Refreshed",
      `Success!\n\n${result.count} tabs discovered and loaded into Config dropdowns:\n` + result.tabs.join(", "),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      "Unable to access source spreadsheet",
      error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Interactive UI Handler: Live Data Sync in Sheets UI
 */
function syncData() {
  try {
    const res = executeCoreDataSync();
    SpreadsheetApp.getUi().alert(
      "Import Completed Successfully",
      `🎉 Live synchronization finished at ${res.timestamp}\n\n` +
      `• Close Cases: ${res.closeRows.toLocaleString()} rows\n` +
      `• Forwarded Cases: ${res.forwardedRows.toLocaleString()} rows\n` +
      `• Total Rows Imported: ${res.totalImported.toLocaleString()} rows`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      "Import Failed",
      "⚠️ Data synchronization encountered an error:\n\n" + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Opens Dashboard WebApp in a modal dialog or opens URL
 */
function openDashboardWebApp() {
  try {
    const html = HtmlService.createTemplateFromFile('Dashboard')
      .evaluate()
      .setTitle('Recon Team CID Master Reporting')
      .setWidth(1200)
      .setHeight(800);
    SpreadsheetApp.getUi().showModalDialog(html, 'Recon Team Dashboard');
  } catch (e) {
    SpreadsheetApp.getUi().alert("Please deploy this project as a WebApp or check doGet() configuration.");
  }
}

/**
 * Core function executed by time-driven trigger:
 * 1. Refreshes source tabs dynamically
 * 2. Synchronizes data to Master sheets
 */
function executeScheduledAutoSync() {
  console.log("⏰ Scheduled auto-refresh and data sync starting...");
  try {
    // Step 1: Refresh source tabs dynamically
    executeCoreRefreshSourceTabs();
    // Step 2: Synchronize data to destination sheets
    const result = executeCoreDataSync();
    console.log("✅ Scheduled auto-sync completed successfully:", JSON.stringify(result));
    return result;
  } catch (err) {
    console.error("⚠️ Scheduled auto-sync failed:", err.message);
    throw err;
  }
}

/**
 * Retrieves the current auto-sync schedule and active trigger status
 */
function getScheduleConfiguration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ensureConfigSheet(ss);
  
  // Read schedule settings from Config or defaults
  let isEnabled = false;
  let scheduleType = "daily"; // "daily" or "interval"
  let scheduleTime = "6"; // 6 AM default
  let triggerDesc = "Disabled";

  const triggers = ScriptApp.getProjectTriggers();
  const syncTrigger = triggers.find(t => 
    t.getHandlerFunction() === "executeScheduledAutoSync" || 
    t.getHandlerFunction() === "executeCoreDataSync"
  );

  if (syncTrigger) {
    isEnabled = true;
    triggerDesc = "Active (Automated Background Trigger)";
  }

  return {
    isEnabled: isEnabled,
    scheduleType: scheduleType,
    scheduleTime: scheduleTime,
    triggerDesc: triggerDesc,
    totalTriggers: triggers.length
  };
}

/**
 * Saves schedule settings and configures Google Apps Script time-based trigger
 * @param {boolean} isEnabled
 * @param {string} scheduleType "daily" | "interval"
 * @param {number|string} timeValue Hour (0-23) for daily, or hours interval (1, 2, 4, 6, 12)
 */
function saveScheduleConfiguration(isEnabled, scheduleType, timeValue) {
  // Always clear existing sync triggers first
  deleteSyncTriggers(false);

  if (isEnabled) {
    const val = Number(timeValue) || 6;
    if (scheduleType === "interval") {
      // Every X hours
      let hours = val;
      if (![1, 2, 4, 6, 8, 12].includes(hours)) hours = 2;
      ScriptApp.newTrigger("executeScheduledAutoSync")
        .timeBased()
        .everyHours(hours)
        .create();
    } else {
      // Daily at specific hour
      let hour = (val >= 0 && val <= 23) ? val : 6;
      ScriptApp.newTrigger("executeScheduledAutoSync")
        .timeBased()
        .atHour(hour)
        .everyDays(1)
        .create();
    }
  }

  return getScheduleConfiguration();
}

/**
 * Removes all active sync triggers
 * @param {boolean} showAlert Whether to show UI alert if called from menu
 */
function deleteSyncTriggers(showAlert) {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === "executeScheduledAutoSync" || fn === "executeCoreDataSync" || fn === "syncData") {
      ScriptApp.deleteTrigger(t);
    }
  });

  if (showAlert) {
    try {
      SpreadsheetApp.getUi().alert(
        "Triggers Removed",
        "All automated sync triggers have been removed.",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (e) {}
  }
}

/**
 * Menu helper for quick 6 AM daily sync setup
 */
function setupDailySyncTrigger() {
  saveScheduleConfiguration(true, "daily", 6);
  try {
    SpreadsheetApp.getUi().alert(
      "Auto-Sync Scheduled",
      "✅ Automated background refresh and data sync has been scheduled for 6:00 AM daily.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {}
}

// ============================================================================
// EMAIL REPORTING SCHEDULER ENGINE
// ============================================================================

/**
 * Ensures the 'EmailSchedules' sheet exists to store automated email report configurations
 */
function ensureEmailScheduleSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("EmailSchedules");
  if (!sheet) {
    sheet = ss.insertSheet("EmailSchedules");
    const headers = [
      "Schedule ID",
      "Report Type",
      "Format Template",
      "Custom Subject",
      "Recipients",
      "Frequency",
      "Target Hour",
      "Active",
      "Last Sent",
      "Status",
      "Created At"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e0e7ff");
    sheet.setFrozenRows(1);

    // Add baseline demo schedule matching user requirements
    sheet.getRange(2, 1, 1, headers.length).setValues([[
      "SCH-3055",
      "Yesterday's Master Report",
      "Raw Master Report (All columns)",
      "Review Daily reports (Yesterday)",
      "barjuVishal@gmail.com",
      "Daily",
      "10:00 AM",
      true,
      "9/2/2026, 10:40:49 AM",
      "Sent Successfully",
      new Date().toISOString()
    ]]);
  }
  return sheet;
}

/**
 * Helper to generate CSV attachment string for email reports
 */
function generateEmailReportCsv(reportType, formatTemplate, customFrom, customTo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timeZone = Session.getScriptTimeZone() || 'Asia/Kolkata';
  const today = new Date();

  let fromDate = null;
  let toDate = null;

  if (reportType === "Yesterday's Master Report") {
    const yest = new Date(today);
    yest.setDate(yest.getDate() - 1);
    fromDate = Utilities.formatDate(yest, timeZone, "yyyy-MM-dd");
    toDate = fromDate;
  } else if (reportType === "Today's Master Report") {
    fromDate = Utilities.formatDate(today, timeZone, "yyyy-MM-dd");
    toDate = fromDate;
  } else if (reportType === "Current Month's Master Report") {
    const firstDay = new Date(today.getFullYear(), today.getMonth(), 1);
    fromDate = Utilities.formatDate(firstDay, timeZone, "yyyy-MM-dd");
    toDate = Utilities.formatDate(today, timeZone, "yyyy-MM-dd");
  } else if (reportType === "Custom Date Range" && customFrom && customTo) {
    fromDate = customFrom;
    toDate = customTo;
  }

  // Generate Raw Master Report (Combined Close Cases + Forwarded Cases)
  const rows = [];
  const headers = ["Activity Source", "Date", "Booking ID / CID", "Employee Name", "Activity / Category", "Status / Remarks"];
  rows.push(headers);

  const closeSheet = ss.getSheetByName("Close Cases");
  const fwdSheet = ss.getSheetByName("Forwarded cases") || ss.getSheetByName("Forwarded Cases");

  if (closeSheet && closeSheet.getLastRow() > 1) {
    const data = closeSheet.getDataRange().getDisplayValues();
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      rows.push(["Close Cases", r[1] || '', r[0] || '', r[3] || '', r[2] || '', r[4] || 'Closed']);
    }
  }

  if (fwdSheet && fwdSheet.getLastRow() > 1) {
    const data = fwdSheet.getDataRange().getDisplayValues();
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      rows.push(["Forwarded Cases", r[1] || '', r[0] || '', r[3] || '', r[2] || '', r[4] || 'Forwarded']);
    }
  }

  return rows.map(r => r.map(c => `"${String(c).replace(/"/g, '""')}"`).join(",")).join("\n");
}

/**
 * Dispatches scheduled report email with CSV attachment
 */
function sendScheduledReportEmail(sched) {
  if (!sched || !sched.recipients) {
    throw new Error("Recipient email address is required.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timeZone = Session.getScriptTimeZone() || 'Asia/Kolkata';
  const todayStr = Utilities.formatDate(new Date(), timeZone, "dd-MMM-yyyy");

  // Format custom subject with tags: {template}, {type}, {date}, {range}, {id}
  let subject = sched.customSubject || `Daily Report - {type} ({date})`;
  subject = subject
    .replace(/{template}/gi, sched.formatTemplate || 'Master Report')
    .replace(/{type}/gi, sched.reportType || 'Master Report')
    .replace(/{date}/gi, todayStr)
    .replace(/{range}/gi, todayStr)
    .replace(/{id}/gi, sched.id || 'SCH-1001');

  // Generate CSV attachment
  const csvContent = generateEmailReportCsv(sched.reportType, sched.formatTemplate, sched.customFrom, sched.customTo);
  const fileName = `Recon_Report_${(sched.reportType || 'Daily').replace(/[^a-zA-Z0-9]/g, '_')}_${todayStr}.csv`;
  const attachment = Utilities.newBlob(csvContent, "text/csv", fileName);

  const htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 650px; margin: 0 auto; color: #1e293b; border: 1px solid #e2e8f0; border-radius: 10px; overflow: hidden;">
      <div style="background: linear-gradient(135deg, #4f46e5 0%, #3730a3 100%); color: white; padding: 24px; text-align: center;">
        <h2 style="margin: 0; font-size: 20px;">S.T.E.P.S Global Solutions</h2>
        <p style="margin: 4px 0 0 0; opacity: 0.9; font-size: 13px;">Recon Team CID Master Reporting System</p>
      </div>
      <div style="padding: 24px; background: white;">
        <p style="font-size: 15px; font-weight: bold; margin-top: 0;">Automated Report Dispatch</p>
        <p style="font-size: 13px; color: #475569; line-height: 1.5;">
          Your scheduled automated report <strong>${sched.reportType || 'Master Report'}</strong> has been generated and is attached below in CSV format.
        </p>
        <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 16px; margin: 16px 0; font-size: 13px;">
          <table style="width: 100%; border-collapse: collapse;">
            <tr><td style="padding: 4px 0; color: #64748b; width: 140px;">Schedule ID:</td><td><strong>${sched.id || 'SCH-3055'}</strong></td></tr>
            <tr><td style="padding: 4px 0; color: #64748b;">Report Type:</td><td><strong>${sched.reportType}</strong></td></tr>
            <tr><td style="padding: 4px 0; color: #64748b;">Format Template:</td><td>${sched.formatTemplate}</td></tr>
            <tr><td style="padding: 4px 0; color: #64748b;">Generated At:</td><td>${Utilities.formatDate(new Date(), timeZone, "dd-MMM-yyyy hh:mm:ss a")}</td></tr>
          </table>
        </div>
        <p style="font-size: 12px; color: #94a3b8; margin-bottom: 0;">
          This automated report was dispatched by Google Apps Script background dispatcher.
        </p>
      </div>
    </div>
  `;

  MailApp.sendEmail({
    to: sched.recipients,
    subject: subject,
    htmlBody: htmlBody,
    attachments: [attachment]
  });

  // Update schedule status in sheet
  const emailSheet = ensureEmailScheduleSheet(ss);
  const data = emailSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(sched.id)) {
      const nowFormatted = Utilities.formatDate(new Date(), timeZone, "M/d/yyyy, h:mm:ss a");
      emailSheet.getRange(i + 1, 9).setValue(nowFormatted);
      emailSheet.getRange(i + 1, 10).setValue("Sent Successfully");
      break;
    }
  }

  return { success: true, timestamp: todayStr, subject: subject };
}

/**
 * Hourly dispatcher trigger function to process due email schedules
 */
function processScheduledEmailDispatcher() {
  console.log("📧 Dispatcher checking scheduled email reports...");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureEmailScheduleSheet(ss);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const timeZone = Session.getScriptTimeZone() || 'Asia/Kolkata';
  const currentHour = Number(Utilities.formatDate(new Date(), timeZone, "H"));

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id = row[0];
    const reportType = row[1];
    const formatTemplate = row[2];
    const customSubject = row[3];
    const recipients = row[4];
    const frequency = row[5];
    const targetHourStr = String(row[6] || '');
    const isActive = (row[7] === true || String(row[7]).toUpperCase() === "TRUE");

    if (!isActive || !recipients) continue;

    // Parse target hour (e.g. "9:00 AM" -> 9, "10:00 PM" -> 22)
    let targetH = 9;
    if (targetHourStr.includes("PM")) {
      const h = parseInt(targetHourStr);
      targetH = (h === 12) ? 12 : h + 12;
    } else if (targetHourStr.includes("AM")) {
      const h = parseInt(targetHourStr);
      targetH = (h === 12) ? 0 : h;
    } else {
      targetH = parseInt(targetHourStr) || 9;
    }

    if (frequency === "Hourly" || targetH === currentHour) {
      try {
        sendScheduledReportEmail({
          id: id,
          reportType: reportType,
          formatTemplate: formatTemplate,
          customSubject: customSubject,
          recipients: recipients
        });
        console.log(`✅ Dispatched report schedule ${id} to ${recipients}`);
      } catch (e) {
        console.error(`⚠️ Failed to dispatch schedule ${id}:`, e.message);
        sheet.getRange(i + 1, 10).setValue("Failed: " + e.message);
      }
    }
  }
}

/**
 * Sets up hourly dispatcher trigger
 */
function setupEmailDispatcherTrigger() {
  deleteEmailDispatcherTrigger();
  ScriptApp.newTrigger("processScheduledEmailDispatcher")
    .timeBased()
    .everyHours(1)
    .create();
  return { isDispatcherActive: true };
}

/**
 * Deletes email dispatcher triggers
 */
function deleteEmailDispatcherTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "processScheduledEmailDispatcher") {
      ScriptApp.deleteTrigger(t);
    }
  });
  return { isDispatcherActive: false };
}

// ============================================================================
// ACCESS CONTROL SYSTEM ENGINE
// ============================================================================

/**
 * Ensures 'AccessControl' sheet exists with default credentials
 */
function ensureAccessControlSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("AccessControl");
  if (!sheet) {
    sheet = ss.insertSheet("AccessControl");
    const headers = ["Email Address", "Password", "Role", "Status", "Created At"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#fee2e2");
    sheet.setFrozenRows(1);

    // Add baseline admin user
    sheet.getRange(2, 1, 1, headers.length).setValues([[
      "admin@stepsglobal.com",
      "admin",
      "Admin (Full Access)",
      "Active",
      new Date().toISOString()
    ]]);
  }
  return sheet;
}

/**
 * Retrieves all registered users from AccessControl sheet
 */
function getAccessUsersData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureAccessControlSheet(ss);
  const data = sheet.getDataRange().getValues();
  const users = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0]) {
      users.push({
        email: String(row[0] || '').trim(),
        password: String(row[1] || '').trim(),
        role: String(row[2] || 'Admin (Full Access)').trim(),
        status: String(row[3] || 'Active').trim(),
        createdAt: String(row[4] || '')
      });
    }
  }

  return { users: users };
}

/**
 * Adds or updates a user in AccessControl sheet
 */
function saveAccessUser(user) {
  if (!user || !user.email) {
    throw new Error("Email address is required.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureAccessControlSheet(ss);
  const data = sheet.getDataRange().getValues();
  const timeZone = Session.getScriptTimeZone() || 'Asia/Kolkata';

  let targetRow = -1;
  const userEmail = String(user.email).trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === userEmail) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow === -1) {
    targetRow = sheet.getLastRow() + 1;
  }

  sheet.getRange(targetRow, 1, 1, 5).setValues([[
    user.email.trim(),
    user.password || 'password123',
    user.role || 'Admin (Full Access)',
    user.status || 'Active',
    user.createdAt || Utilities.formatDate(new Date(), timeZone, "M/d/yyyy, h:mm:ss a")
  ]]);

  return getAccessUsersData();
}

/**
 * Deletes a user by email from AccessControl sheet
 */
function deleteAccessUser(email) {
  if (!email) throw new Error("Email is required for deletion.");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureAccessControlSheet(ss);
  const data = sheet.getDataRange().getValues();
  const userEmail = String(email).trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === userEmail) {
      sheet.deleteRow(i + 1);
      break;
    }
  }

  return getAccessUsersData();
}

/**
 * Authenticates user credentials against AccessControl sheet
 */
function validateUserLogin(email, password) {
  if (!email || !password) {
    return { success: false, message: "Email and password are required." };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureAccessControlSheet(ss);
  const data = sheet.getDataRange().getValues();

  const inputEmail = String(email).trim().toLowerCase();
  const inputPass = String(password).trim();

  // Baseline admin fallback
  if (inputEmail === "admin@stepsglobal.com" && inputPass === "admin") {
    return {
      success: true,
      user: {
        email: "admin@stepsglobal.com",
        role: "Admin (Full Access)"
      }
    };
  }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowEmail = String(row[0] || '').trim().toLowerCase();
    const rowPass = String(row[1] || '').trim();
    const rowRole = String(row[2] || 'Member (Read-Only Reports/Dashboard)').trim();
    const rowStatus = String(row[3] || 'Active').trim();

    if (rowEmail === inputEmail) {
      if (rowStatus.toLowerCase() === 'inactive') {
        return { success: false, message: "This account has been deactivated. Please contact an admin." };
      }
      if (rowPass === inputPass) {
        return {
          success: true,
          user: {
            email: row[0],
            role: rowRole
          }
        };
      } else {
        return { success: false, message: "Incorrect password. Please try again." };
      }
    }
  }

  return { success: false, message: "Invalid email or password. Please check your credentials." };
}