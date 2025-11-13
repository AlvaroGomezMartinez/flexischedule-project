/**
 * COGNOS Report Consolidation System
 * Main entry point for Google Apps Script automation
 */

// Email configuration for the three COGNOS reports
const EMAIL_CONFIG = {
  attendance: {
    subject: 'A new version of My ATT - Attendance Bulletin is available',
    sheetName: 'BHS attendance',
    attendanceCodeColumn: 'K',  // Attendance code column
    periodColumn: 'J',           // Period column for filtering
    periodFilter: '02'           // Only import 2nd period rows (stored as text)
  },
  courses: {
    subject: 'A new version of My Student CY List - Course, Teacher & Room is available',
    sheetName: '2nd period default',
    teacherColumn: 'G',  // Teacher name column (after reordering)
    originalColumns: ['Student Name', 'Student Id', 'Grade', '9th Grd Entry', 'Period', 'Description', 'Room', 'Instructor', 'Instructor ID', 'Instructor Email'],
    targetColumns: ['Student Id', 'Student Name', 'Grade', 'Period', 'Description', 'Room', 'Instructor', 'Instructor Id', 'Instructor Email']
  },
  contacts: {
    subject: 'A new version of My Student CY List - Student Email/Contact Info - Next Year Option is available',
    sheetName: 'contact info',
    emailColumns: {
      student: 'M',
      guardian1: 'F',
      guardian2: 'J'
    }
  }
};

// Main configuration for sheet names and column mappings
const CONFIG = {
  // Sheet names
  sheetNames: {
    attendance: 'BHS attendance',
    courses: '2nd period default',
    contacts: 'contact info',
    mailOut: 'Mail Out',
    flexAbsencesPattern: /\d+\.\d+\s+flex absences/i  // Matches "11.3 flex absences"
  },
  
  // Column mappings for enrichment
  columnMappings: {
    flexAbsences: {
      flexiSchedData: 'A:L',      // User-pasted FlexiSched data (L = FlexiSched Comment)
      attendanceCode: 'M',         // From BHS attendance or #N/A
      teacherName: 'N',            // From 2nd period default
      studentEmail: 'O',           // From contact info
      guardian1Email: 'P',         // From contact info
      guardian2Email: 'Q'          // From contact info
    },
    bhsAttendance: {
      attendanceCode: 'K'  // Attendance code is in column K
    },
    secondPeriod: {
      teacherName: 'G'  // Teacher name is in column G (after reordering)
    },
    contactInfo: {
      studentEmail: 'M',
      guardian1Email: 'F',
      guardian2Email: 'J'
    }
  },
  
  // Student ID column (assumed to be column A in all sheets)
  studentIdColumn: 'A',
  
  // Flex absences sheet headers
  flexAbsencesHeaders: {
    M: 'Attendance Code',
    N: '2nd Period Teacher',
    O: 'Student Email',        // From contact info column N
    P: 'Guardian 1 Email',     // From contact info column G
    Q: 'Guardian 2 Email'      // From contact info column K
  }
};

// ============================================================================
// UI MODULE
// ============================================================================

/**
 * Creates custom menu when spreadsheet opens
 * This function is automatically triggered by Google Sheets
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Flex Absence Tracker')
    .addItem('Create today\'s flex absences sheet', 'createTodaysFlexAbsencesSheet')
    .addItem('Import COGNOS Reports from GMail', 'importCognosReports')
    .addItem('Add data to Flex Absences sheet', 'enrichFlexAbsences')
    .addItem('Sync Comments from Mail Out sheet', 'syncComments')
    .addSeparator()
    .addItem('Test Header Setup (Debug)', 'testSetupHeaders')
    .addItem('Help', 'openHelpDocument')
    .addToUi();
}

/**
 * Opens the FlexiSched Directions Google Doc in a new browser tab
 */
function openHelpDocument() {
  const docId = '1rbjRHQ_4XxhCxDrPdaPX97d0frXToJT0FhZVa8slFR4';
  const url = `https://docs.google.com/document/d/${docId}/edit`;
  
  const html = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_blank">
        <style>
          body {
            font-family: Arial, sans-serif;
            padding: 20px;
            text-align: center;
            margin: 0;
          }
          .hidden {
            display: none;
          }
          .link-container {
            margin-top: 15px;
          }
          a {
            color: #4285f4;
            text-decoration: none;
            font-weight: bold;
            font-size: 16px;
          }
          a:hover {
            text-decoration: underline;
          }
          p {
            color: #666;
            font-size: 14px;
            margin: 10px 0;
          }
          .close-btn {
            margin-top: 15px;
            padding: 8px 16px;
            background-color: #4285f4;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
          }
          .close-btn:hover {
            background-color: #357ae8;
          }
        </style>
      </head>
      <body>
        <div id="message">
          <p>Click the link below to open the help document:</p>
          <div class="link-container">
            <a href="${url}" target="_blank" onclick="linkClicked()">📖 Open FlexiSched Help</a>
          </div>
          <button class="close-btn" onclick="google.script.host.close()">Close</button>
        </div>
        <script>
          var opened = window.open('${url}', '_blank');
          
          function linkClicked() {
            setTimeout(function() {
              google.script.host.close();
            }, 500);
          }
          
          // If popup opened successfully, auto-close after a moment
          if (opened) {
            setTimeout(function() {
              google.script.host.close();
            }, 1000);
          }
        </script>
      </body>
    </html>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(350)
    .setHeight(150);
  
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, 'FlexiSched Help');
}

/**
 * Displays a success message as a toast notification
 * @param {string} message - The success message to display
 */
function showSuccessMessage(message) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Success', 5);
}

/**
 * Displays an error message as an alert dialog
 * @param {string} message - The error message to display
 */
function showErrorMessage(message) {
  const ui = SpreadsheetApp.getUi();
  ui.alert('Error', message, ui.ButtonSet.OK);
}

// ============================================================================
// SHEET CREATION MODULE
// ============================================================================

/**
 * Formats a date as "M.D" for sheet naming (e.g., "11.3" for November 3rd)
 * @param {Date} date - The date to format
 * @returns {string} Formatted date string in "M.D" format
 */
function formatDateForSheetName(date) {
  const month = date.getMonth() + 1;  // getMonth() returns 0-11
  const day = date.getDate();
  return `${month}.${day}`;
}

/**
 * Creates a new flex absences sheet with today's date
 * Checks for duplicate sheet names and sets up column headers
 * Requirements: 7.2, 7.3, 7.4
 */
function createTodaysFlexAbsencesSheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const today = new Date();
    const dateStr = formatDateForSheetName(today);
    const sheetName = `${dateStr} flex absences`;
    
    // Check if sheet already exists
    const existingSheet = spreadsheet.getSheetByName(sheetName);
    if (existingSheet) {
      showSuccessMessage(`Sheet "${sheetName}" already exists.`);
      return;
    }
    
    // Create new sheet
    const newSheet = spreadsheet.insertSheet(sheetName);
    
    // Set up headers
    setupFlexAbsencesHeaders(newSheet);
    
    showSuccessMessage(`Created new sheet: "${sheetName}"`);
    
  } catch (error) {
    Logger.log(`Error creating flex absences sheet: ${error.message}`);
    showErrorMessage(`Failed to create flex absences sheet: ${error.message}`);
  }
}

/**
 * Sets up initial template for a flex absences sheet
 * Provides user instructions and sets enrichment headers in row 2 (L-Q)
 * FlexiSched data will be pasted by user and will include its own headers in rows 1-2
 * Requirements: 7.5, 7.6
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to set up headers for
 */
function setupFlexAbsencesHeaders(sheet) {
  try {
    Logger.log('Setting up flex absences headers...');
    
    // Set instruction in row 1 for user guidance
    sheet.getRange('A1').setValue('Paste FlexiSched data here (will overwrite this row and row 2)');
    sheet.getRange('A1').setFontStyle('italic');
    sheet.getRange('A1').setFontColor('#666666');
    
    // Set enriched data headers in row 2 (columns M-Q)
    // These will be restored automatically when enrichment runs if overwritten
    const headers = CONFIG.flexAbsencesHeaders;
    Logger.log('Headers config:', headers);
    
    sheet.getRange('M2').setValue(headers.M);  // Attendance Code
    sheet.getRange('N2').setValue(headers.N);  // 2nd Period Teacher
    sheet.getRange('O2').setValue(headers.O);  // Student Email
    sheet.getRange('P2').setValue(headers.P);  // Guardian 1 Email
    sheet.getRange('Q2').setValue(headers.Q);  // Guardian 2 Email
    
    Logger.log('Headers set in row 2, columns M-Q');
    
    // Format enrichment headers (bold)
    sheet.getRange('M2:Q2').setFontWeight('bold');
    Logger.log('Headers formatted as bold');
    
    // Add a note to cell A1 explaining the process
    const instructionNote = 'Instructions:\n' +
      '1. Delete all data from this sheet\n' +
      '2. Paste FlexiSched report data (includes 2 header rows)\n' +
      '3. Click "Add data to Flex Absences sheet" to enrich data\n\n' +
      'The enrichment headers (M-Q) will be automatically restored if overwritten.';
    sheet.getRange('A1').setNote(instructionNote);
    
    Logger.log('Setup complete for flex absences headers');
    
  } catch (error) {
    Logger.log(`Error in setupFlexAbsencesHeaders: ${error.message}`);
    throw error;
  }
}

// ============================================================================
// SHEET MANAGER MODULE
// ============================================================================

/**
 * Returns an existing sheet by name or creates a new one if it doesn't exist
 * Requirements: 2.5, 3.1
 * @param {string} sheetName - The name of the sheet to get or create
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The existing or newly created sheet
 */
function getOrCreateSheet(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    Logger.log(`Created new sheet: ${sheetName}`);
  }
  
  return sheet;
}

/**
 * Finds and returns the flex absences sheet matching date pattern and "flex absences" suffix
 * Requirements: 2.5, 3.1
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null} The flex absences sheet or null if not found
 */
function getFlexAbsencesSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const pattern = CONFIG.sheetNames.flexAbsencesPattern;
  
  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    const sheetName = sheet.getName();
    
    if (pattern.test(sheetName)) {
      Logger.log(`Found flex absences sheet: ${sheetName}`);
      return sheet;
    }
  }
  
  Logger.log('No flex absences sheet found matching pattern');
  return null;
}

/**
 * Clears data from a sheet starting at the specified row
 * Preserves headers (assumes row 1 contains headers)
 * Requirements: 2.5, 3.1
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to clear data from
 * @param {number} startRow - The row number to start clearing from (1-indexed)
 */
function clearSheetData(sheet, startRow) {
  const lastRow = sheet.getLastRow();
  
  // If startRow is beyond the last row, nothing to clear
  if (startRow > lastRow) {
    Logger.log(`No data to clear in sheet ${sheet.getName()} starting from row ${startRow}`);
    return;
  }
  
  const lastColumn = sheet.getLastColumn();
  
  // If there are no columns, nothing to clear
  if (lastColumn === 0) {
    Logger.log(`No columns in sheet ${sheet.getName()}`);
    return;
  }
  
  const numRows = lastRow - startRow + 1;
  const range = sheet.getRange(startRow, 1, numRows, lastColumn);
  range.clearContent();
  
  Logger.log(`Cleared ${numRows} rows from sheet ${sheet.getName()} starting at row ${startRow}`);
}

/**
 * Writes a 2D array of data to a sheet starting at the specified row
 * Requirements: 2.6, 2.7
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to write data to
 * @param {Array<Array>} data - 2D array of data to write
 * @param {number} startRow - The row number to start writing at (1-indexed)
 */
function writeDataToSheet(sheet, data, startRow) {
  if (!data || data.length === 0) {
    Logger.log('No data to write to sheet');
    return;
  }
  
  const numRows = data.length;
  const numCols = data[0].length;
  
  const range = sheet.getRange(startRow, 1, numRows, numCols);
  range.setValues(data);
  
  Logger.log(`Wrote ${numRows} rows and ${numCols} columns to sheet ${sheet.getName()} starting at row ${startRow}`);
}

/**
 * Adds a note to a specific cell in a sheet
 * Requirements: 2.6, 2.7
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet containing the cell
 * @param {string} cell - The cell address in A1 notation (e.g., "A1")
 * @param {string} note - The note text to add to the cell
 */
function addNoteToCell(sheet, cell, note) {
  const range = sheet.getRange(cell);
  range.setNote(note);
  
  Logger.log(`Added note to cell ${cell} in sheet ${sheet.getName()}`);
}

/**
 * Sets up column headers for a sheet
 * Requirements: 2.6, 2.7
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to set headers for
 * @param {Array<string>} headers - Array of header strings
 */
function setSheetHeaders(sheet, headers) {
  if (!headers || headers.length === 0) {
    Logger.log('No headers to set');
    return;
  }
  
  const range = sheet.getRange(1, 1, 1, headers.length);
  range.setValues([headers]);
  range.setFontWeight('bold');
  
  Logger.log(`Set ${headers.length} headers in sheet ${sheet.getName()}`);
}

// ============================================================================
// GMAIL SERVICE MODULE
// ============================================================================

/**
 * Gets the active user's email address
 * Requirements: 1.1, 8.1, 8.2, 8.3
 * @returns {string} The user's email address
 */
function getUserEmail() {
  const userEmail = Session.getActiveUser().getEmail();
  Logger.log(`Detected user email: ${userEmail}`);
  return userEmail;
}

/**
 * Searches Gmail for COGNOS report emails from the user's own email address
 * Returns the most recent email for each of the three report types
 * Requirements: 1.1, 8.1, 8.2, 8.3
 * @returns {Object} Object with keys 'attendance', 'courses', 'contacts' containing email data or null
 */
function searchCognosEmails() {
  try {
    const userEmail = getUserEmail();
    const results = {
      attendance: null,
      courses: null,
      contacts: null
    };
    
    // Search for each report type
    for (const reportType in EMAIL_CONFIG) {
      const config = EMAIL_CONFIG[reportType];
      const subject = config.subject;
      
      // Build search query: from user's email AND subject matches
      const searchQuery = `from:${userEmail} subject:"${subject}"`;
      Logger.log(`Searching Gmail with query: ${searchQuery}`);
      
      // Search Gmail and get the most recent thread
      const threads = GmailApp.search(searchQuery, 0, 1);
      
      if (threads.length > 0) {
        const messages = threads[0].getMessages();
        
        if (messages.length > 0) {
          // Get the most recent message in the thread
          const message = messages[messages.length - 1];
          
          results[reportType] = {
            messageId: message.getId(),
            subject: message.getSubject(),
            from: message.getFrom(),
            date: message.getDate(),
            reportType: reportType
          };
          
          Logger.log(`Found ${reportType} report: ${message.getSubject()} from ${message.getDate()}`);
        }
      } else {
        Logger.log(`No emails found for ${reportType} report with subject: ${subject}`);
      }
    }
    
    return results;
    
  } catch (error) {
    Logger.log(`Error searching COGNOS emails: ${error.message}`);
    throw error;
  }
}

/**
 * Retrieves Excel attachments from a Gmail message
 * Requirements: 1.2, 1.3, 1.4
 * @param {string} messageId - The Gmail message ID
 * @returns {Array<Object>} Array of attachment objects with blob, filename, and mimeType
 */
function getAttachments(messageId) {
  try {
    const message = GmailApp.getMessageById(messageId);
    
    if (!message) {
      throw new Error(`Message not found with ID: ${messageId}`);
    }
    
    const attachments = message.getAttachments();
    const excelAttachments = [];
    
    // Filter for Excel files (.xlsx format)
    for (let i = 0; i < attachments.length; i++) {
      const attachment = attachments[i];
      const mimeType = attachment.getContentType();
      const filename = attachment.getName();
      
      // Check for Excel 2007+ format (.xlsx)
      if (mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
          filename.toLowerCase().endsWith('.xlsx')) {
        
        excelAttachments.push({
          blob: attachment,
          filename: filename,
          mimeType: mimeType
        });
        
        Logger.log(`Found Excel attachment: ${filename}`);
      }
    }
    
    if (excelAttachments.length === 0) {
      throw new Error(`No Excel attachments found in message ${messageId}`);
    }
    
    return excelAttachments;
    
  } catch (error) {
    Logger.log(`Error retrieving attachments from message ${messageId}: ${error.message}`);
    throw error;
  }
}

/**
 * Identifies which COGNOS report type an email contains based on subject
 * Requirements: 1.2, 1.3, 1.4
 * @param {Object} email - Email object with subject property
 * @returns {string|null} Report type ('attendance', 'courses', 'contacts') or null if not recognized
 */
function identifyReportType(email) {
  if (!email || !email.subject) {
    Logger.log('Invalid email object provided to identifyReportType');
    return null;
  }
  
  const subject = email.subject;
  
  // Check against each configured report type
  for (const reportType in EMAIL_CONFIG) {
    const config = EMAIL_CONFIG[reportType];
    
    if (subject.includes(config.subject)) {
      Logger.log(`Identified report type: ${reportType} for subject: ${subject}`);
      return reportType;
    }
  }
  
  Logger.log(`Could not identify report type for subject: ${subject}`);
  return null;
}

// ============================================================================
// EXCEL CONVERTER MODULE
// ============================================================================

/**
 * Converts an Excel blob to a 2D array
 * Requirements: 2.4
 * @param {GoogleAppsScript.Base.Blob} blob - The Excel file blob
 * @returns {Array<Array>} 2D array of cell values
 */
function convertExcelToArray(blob) {
  try {
    // Convert Excel blob to temporary spreadsheet
    const tempFile = Drive.Files.insert({
      title: 'temp_excel_' + new Date().getTime(),
      mimeType: MimeType.GOOGLE_SHEETS
    }, blob, {
      convert: true
    });
    
    // Open the converted spreadsheet
    const tempSpreadsheet = SpreadsheetApp.openById(tempFile.id);
    const tempSheet = tempSpreadsheet.getSheets()[0];
    
    // Get all data from the sheet
    const lastRow = tempSheet.getLastRow();
    const lastColumn = tempSheet.getLastColumn();
    
    let data = [];
    if (lastRow > 0 && lastColumn > 0) {
      const range = tempSheet.getRange(1, 1, lastRow, lastColumn);
      data = range.getValues();
    }
    
    // Clean up: delete the temporary file
    Drive.Files.remove(tempFile.id);
    
    Logger.log(`Converted Excel blob to array with ${data.length} rows`);
    return data;
    
  } catch (error) {
    Logger.log(`Error converting Excel to array: ${error.message}`);
    throw new Error(`Failed to convert Excel file: ${error.message}`);
  }
}

/**
 * Parses Excel data and extracts headers and data rows
 * Requirements: 2.4
 * @param {GoogleAppsScript.Base.Blob} excelBlob - The Excel file blob
 * @returns {Object} Object with 'headers' array and 'rows' 2D array
 */
function parseExcelData(excelBlob) {
  try {
    const allData = convertExcelToArray(excelBlob);
    
    if (!allData || allData.length === 0) {
      Logger.log('No data found in Excel file');
      return {
        headers: [],
        rows: []
      };
    }
    
    // First row is headers
    const headers = allData[0];
    
    // Remaining rows are data
    const rows = allData.slice(1);
    
    Logger.log(`Parsed Excel data: ${headers.length} columns, ${rows.length} data rows`);
    
    return {
      headers: headers,
      rows: rows
    };
    
  } catch (error) {
    Logger.log(`Error parsing Excel data: ${error.message}`);
    throw error;
  }
}

/**
 * Filters data rows by a specific column value
 * Used to filter attendance data to only 2nd period (column J = "02")
 * Requirements: 2.1, 2.2
 * @param {Array<Array>} data - 2D array with headers in first row and data in subsequent rows
 * @param {number} periodColumn - Zero-based column index to filter on (e.g., 9 for column J)
 * @param {string} periodValue - The value to filter for (e.g., "02")
 * @returns {Array<Array>} Filtered 2D array with headers and matching rows
 */
function filterByPeriod(data, periodColumn, periodValue) {
  try {
    if (!data || data.length === 0) {
      Logger.log('No data to filter');
      return [];
    }
    
    // First row is headers
    const headers = data[0];
    const dataRows = data.slice(1);
    
    // Filter rows where the specified column matches the period value
    const filteredRows = dataRows.filter(row => {
      // Convert to string for comparison (handles both string and number types)
      const cellValue = row[periodColumn] ? String(row[periodColumn]).trim() : '';
      return cellValue === periodValue;
    });
    
    Logger.log(`Filtered ${dataRows.length} rows to ${filteredRows.length} rows where column ${periodColumn} = "${periodValue}"`);
    
    // Return headers plus filtered rows
    return [headers].concat(filteredRows);
    
  } catch (error) {
    Logger.log(`Error filtering by period: ${error.message}`);
    throw error;
  }
}

/**
 * Reorders columns in data array based on column mapping
 * Used to reorder and exclude columns for the courses report
 * Requirements: 2.1, 2.2
 * @param {Array<Array>} data - 2D array with headers in first row and data in subsequent rows
 * @param {Object} columnMapping - Object with 'originalColumns' and 'targetColumns' arrays
 * @returns {Array<Array>} Reordered 2D array with headers and data
 */
function reorderColumns(data, columnMapping) {
  try {
    if (!data || data.length === 0) {
      Logger.log('No data to reorder');
      return [];
    }
    
    if (!columnMapping || !columnMapping.originalColumns || !columnMapping.targetColumns) {
      Logger.log('Invalid column mapping provided');
      throw new Error('Column mapping must include originalColumns and targetColumns arrays');
    }
    
    const headers = data[0];
    const dataRows = data.slice(1);
    
    // Create a map of original column names to their indices
    const columnIndexMap = {};
    for (let i = 0; i < headers.length; i++) {
      const headerName = String(headers[i]).trim();
      columnIndexMap[headerName] = i;
    }
    
    // Build array of column indices in the target order
    const targetIndices = [];
    for (let i = 0; i < columnMapping.targetColumns.length; i++) {
      const targetColumn = columnMapping.targetColumns[i];
      const originalIndex = columnIndexMap[targetColumn];
      
      if (originalIndex === undefined) {
        Logger.log(`Warning: Target column "${targetColumn}" not found in original data`);
        targetIndices.push(-1);  // Mark as missing
      } else {
        targetIndices.push(originalIndex);
      }
    }
    
    // Reorder headers
    const reorderedHeaders = columnMapping.targetColumns.slice();
    
    // Reorder data rows
    const reorderedRows = dataRows.map(row => {
      return targetIndices.map(index => {
        return index >= 0 ? row[index] : '';  // Use empty string for missing columns
      });
    });
    
    Logger.log(`Reordered data from ${headers.length} columns to ${reorderedHeaders.length} columns`);
    
    // Return reordered headers plus reordered data rows
    return [reorderedHeaders].concat(reorderedRows);
    
  } catch (error) {
    Logger.log(`Error reordering columns: ${error.message}`);
    throw error;
  }
}

// ============================================================================
// COGNOS REPORT IMPORT ORCHESTRATION
// ============================================================================

/**
 * Main orchestration function for importing COGNOS reports from Gmail
 * Searches for three COGNOS reports, extracts Excel attachments, applies special processing,
 * and imports data to respective sheets with timestamp notes
 * Requirements: 1.5, 1.6, 2.1, 2.2, 2.3, 2.5, 2.6, 2.7
 */
function importCognosReports() {
  try {
    Logger.log('Starting COGNOS report import...');
    
    // Search Gmail for the three COGNOS reports
    const emailResults = searchCognosEmails();
    
    // Track which reports were found and imported
    const importResults = {
      attendance: { found: false, imported: false, error: null },
      courses: { found: false, imported: false, error: null },
      contacts: { found: false, imported: false, error: null }
    };
    
    // Process each report type
    for (const reportType in EMAIL_CONFIG) {
      const config = EMAIL_CONFIG[reportType];
      const email = emailResults[reportType];
      
      if (!email) {
        // Report not found
        const errorMsg = `Report not found: ${config.subject}`;
        Logger.log(errorMsg);
        importResults[reportType].error = errorMsg;
        
        // Add error note to sheet A1
        const sheet = getOrCreateSheet(config.sheetName);
        const timestamp = new Date().toLocaleString();
        addNoteToCell(sheet, 'A1', `${timestamp}: ${errorMsg}`);
        
        continue;
      }
      
      importResults[reportType].found = true;
      
      try {
        // Extract Excel attachment from email
        Logger.log(`Processing ${reportType} report from email: ${email.subject}`);
        const attachments = getAttachments(email.messageId);
        
        if (attachments.length === 0) {
          throw new Error('No Excel attachments found in email');
        }
        
        // Use the first Excel attachment
        const attachment = attachments[0];
        Logger.log(`Processing attachment: ${attachment.filename}`);
        
        // Parse Excel data
        const parsedData = parseExcelData(attachment.blob);
        let processedData = [parsedData.headers].concat(parsedData.rows);
        
        // Apply special processing based on report type
        if (reportType === 'attendance') {
          // Filter to only 2nd period rows (column J = "02")
          // Column J is index 9 (zero-based)
          Logger.log('Applying 2nd period filter to attendance data...');
          processedData = filterByPeriod(processedData, 9, config.periodFilter);
          
        } else if (reportType === 'courses') {
          // Reorder columns and exclude "9th Grd Entry"
          Logger.log('Reordering columns for courses data...');
          processedData = reorderColumns(processedData, {
            originalColumns: config.originalColumns,
            targetColumns: config.targetColumns
          });
        }
        // contacts report needs no special processing
        
        // Get or create the target sheet
        const sheet = getOrCreateSheet(config.sheetName);
        
        // Clear existing data (preserve headers by clearing from row 2 onward)
        clearSheetData(sheet, 2);
        
        // Write data to sheet (starting from row 1 to include headers)
        writeDataToSheet(sheet, processedData, 1);
        
        // Add timestamp note to cell A1
        const timestamp = new Date().toLocaleString();
        const successNote = `Successfully imported on ${timestamp} from email dated ${email.date.toLocaleString()}`;
        addNoteToCell(sheet, 'A1', successNote);
        
        importResults[reportType].imported = true;
        Logger.log(`Successfully imported ${reportType} report to ${config.sheetName}`);
        
      } catch (error) {
        // Error processing this specific report
        const errorMsg = `Error importing ${reportType} report: ${error.message}`;
        Logger.log(errorMsg);
        importResults[reportType].error = errorMsg;
        
        // Add error note to sheet A1
        const sheet = getOrCreateSheet(config.sheetName);
        const timestamp = new Date().toLocaleString();
        addNoteToCell(sheet, 'A1', `${timestamp}: ${errorMsg}`);
      }
    }
    
    // Count successfully imported reports
    let importedCount = 0;
    const missingReports = [];
    const failedReports = [];
    
    for (const reportType in importResults) {
      const result = importResults[reportType];
      
      if (result.imported) {
        importedCount++;
      } else if (!result.found) {
        missingReports.push(EMAIL_CONFIG[reportType].subject);
      } else if (result.error) {
        failedReports.push(reportType);
      }
    }
    
    // Display results to user
    if (importedCount === 3) {
      showSuccessMessage(`Successfully imported all 3 COGNOS reports!`);
    } else if (importedCount > 0) {
      let message = `Imported ${importedCount} of 3 reports.`;
      
      if (missingReports.length > 0) {
        message += `\n\nMissing reports:\n- ${missingReports.join('\n- ')}`;
      }
      
      if (failedReports.length > 0) {
        message += `\n\nFailed to import: ${failedReports.join(', ')}`;
      }
      
      message += '\n\nCheck cell A1 notes in each sheet for details.';
      showErrorMessage(message);
    } else {
      showErrorMessage('Failed to import any reports. Check cell A1 notes in each sheet for details.');
    }
    
    Logger.log(`Import complete. Imported ${importedCount} of 3 reports.`);
    
  } catch (error) {
    Logger.log(`Critical error in importCognosReports: ${error.message}`);
    showErrorMessage(`Failed to import COGNOS reports: ${error.message}`);
  }
}

// ============================================================================
// ENRICHMENT SERVICE MODULE
// ============================================================================

/**
 * Identifies which column contains the student ID in a sheet
 * Assumes student ID is in column A by default, but can be configured
 * Requirements: 4.3, 4.5, 4.6
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to analyze
 * @returns {number} Zero-based column index containing student ID (default: 0 for column A)
 */
function getStudentIdColumn(sheet) {
  // For this implementation, student ID is always in column A (index 0)
  // This function exists for future flexibility if column positions change
  const studentIdColumn = 0;  // Column A
  
  Logger.log(`Student ID column for sheet ${sheet.getName()}: ${studentIdColumn} (Column A)`);
  return studentIdColumn;
}

/**
 * Builds a lookup map from a sheet mapping student ID to specified data columns
 * Requirements: 4.3, 4.5, 4.6
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to build map from
 * @param {number} idColumn - Zero-based column index containing student ID
 * @param {Array<number>} dataColumns - Array of zero-based column indices to extract
 * @returns {Map<string, Array>} Map of student ID (as string) to array of data values
 */
function buildStudentMap(sheet, idColumn, dataColumns) {
  try {
    const studentMap = new Map();
    
    // Get all data from the sheet
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    
    if (lastRow <= 1) {
      Logger.log(`Sheet ${sheet.getName()} has no data rows (only headers or empty)`);
      return studentMap;
    }
    
    if (lastColumn === 0) {
      Logger.log(`Sheet ${sheet.getName()} has no columns`);
      return studentMap;
    }
    
    // Get data starting from row 2 (skip headers in row 1)
    const dataRange = sheet.getRange(2, 1, lastRow - 1, lastColumn);
    const data = dataRange.getValues();
    
    // Build the map
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const studentId = String(row[idColumn]).trim();
      
      // Skip rows with empty student ID
      if (!studentId || studentId === '') {
        continue;
      }
      
      // Extract data from specified columns
      const dataValues = dataColumns.map(colIndex => {
        return colIndex < row.length ? row[colIndex] : '';
      });
      
      // Store in map (if duplicate IDs exist, last one wins)
      studentMap.set(studentId, dataValues);
    }
    
    Logger.log(`Built student map from sheet ${sheet.getName()}: ${studentMap.size} students`);
    return studentMap;
    
  } catch (error) {
    Logger.log(`Error building student map from sheet ${sheet.getName()}: ${error.message}`);
    throw error;
  }
}

/**
 * Adds attendance codes from BHS attendance sheet to flex absences sheet column M
 * If no matching student found, adds #N/A
 * Requirements: 4.3, 4.4, 4.5, 4.6
 * @param {GoogleAppsScript.Spreadsheet.Sheet} flexSheet - The flex absences sheet
 * @param {Map<string, Array>} attendanceMap - Map of student ID to [attendance code]
 */
function addAttendanceCodes(flexSheet, attendanceMap) {
  try {
    const lastRow = flexSheet.getLastRow();
    
    if (lastRow <= 2) {
      Logger.log('No data rows in flex absences sheet to enrich (need at least 3 rows for FlexiSched data)');
      return;
    }
    
    // Get student IDs from column A (starting from row 3, skipping FlexiSched headers in rows 1-2)
    const studentIdRange = flexSheet.getRange(3, 1, lastRow - 2, 1);
    const studentIds = studentIdRange.getValues();
    
    // Prepare attendance codes to write to column M
    const attendanceCodes = [];
    
    for (let i = 0; i < studentIds.length; i++) {
      const studentId = String(studentIds[i][0]).trim();
      
      if (!studentId || studentId === '') {
        // Empty row, leave attendance code empty
        attendanceCodes.push(['']);
        continue;
      }
      
      // Lookup student in attendance map
      if (attendanceMap.has(studentId)) {
        const data = attendanceMap.get(studentId);
        const attendanceCode = data[0] || '';  // First element is attendance code
        attendanceCodes.push([attendanceCode]);
      } else {
        // Student not found in BHS attendance - mark as #N/A
        attendanceCodes.push(['#N/A']);
      }
    }
    
    // Ensure header is set for column M before writing data
    const headerM = flexSheet.getRange('M2');
    if (!headerM.getValue() || headerM.getValue() !== CONFIG.flexAbsencesHeaders.M) {
      headerM.setValue(CONFIG.flexAbsencesHeaders.M);
      headerM.setFontWeight('bold');
      Logger.log('Set header for column M: ' + CONFIG.flexAbsencesHeaders.M);
    }
    
    // Write attendance codes to column M (column 13) starting from row 3
    const targetRange = flexSheet.getRange(3, 13, attendanceCodes.length, 1);
    targetRange.setValues(attendanceCodes);
    
    Logger.log(`Added attendance codes to ${attendanceCodes.length} rows in flex absences sheet`);
    
  } catch (error) {
    Logger.log(`Error adding attendance codes: ${error.message}`);
    throw error;
  }
}

/**
 * Adds teacher names from 2nd period default sheet to flex absences sheet column N
 * Requirements: 4.3, 4.4, 4.5, 4.6
 * @param {GoogleAppsScript.Spreadsheet.Sheet} flexSheet - The flex absences sheet
 * @param {Map<string, Array>} teacherMap - Map of student ID to [teacher name]
 */
function addTeacherNames(flexSheet, teacherMap) {
  try {
    const lastRow = flexSheet.getLastRow();
    
    if (lastRow <= 2) {
      Logger.log('No data rows in flex absences sheet to enrich (need at least 3 rows for FlexiSched data)');
      return;
    }
    
    // Get student IDs from column A (starting from row 3, skipping FlexiSched headers in rows 1-2)
    const studentIdRange = flexSheet.getRange(3, 1, lastRow - 2, 1);
    const studentIds = studentIdRange.getValues();
    
    // Prepare teacher names to write to column N
    const teacherNames = [];
    
    for (let i = 0; i < studentIds.length; i++) {
      const studentId = String(studentIds[i][0]).trim();
      
      if (!studentId || studentId === '') {
        // Empty row, leave teacher name empty
        teacherNames.push(['']);
        continue;
      }
      
      // Lookup student in teacher map
      if (teacherMap.has(studentId)) {
        const data = teacherMap.get(studentId);
        const teacherName = data[0] || '';  // First element is teacher name
        teacherNames.push([teacherName]);
      } else {
        // Student not found in 2nd period default - leave empty
        teacherNames.push(['']);
      }
    }
    
    // Ensure header is set for column N before writing data
    const headerN = flexSheet.getRange('N2');
    if (!headerN.getValue() || headerN.getValue() !== CONFIG.flexAbsencesHeaders.N) {
      headerN.setValue(CONFIG.flexAbsencesHeaders.N);
      headerN.setFontWeight('bold');
      Logger.log('Set header for column N: ' + CONFIG.flexAbsencesHeaders.N);
    }
    
    // Write teacher names to column N (column 14) starting from row 3
    const targetRange = flexSheet.getRange(3, 14, teacherNames.length, 1);
    targetRange.setValues(teacherNames);
    
    Logger.log(`Added teacher names to ${teacherNames.length} rows in flex absences sheet`);
    
  } catch (error) {
    Logger.log(`Error adding teacher names: ${error.message}`);
    throw error;
  }
}

/**
 * Adds contact information (emails) from contact info sheet to flex absences sheet columns O-Q
 * Column mapping: Contact Info N→Flex O, Contact Info G→Flex P, Contact Info K→Flex Q
 * Requirements: 4.3, 4.4, 4.5, 4.6
 * @param {GoogleAppsScript.Spreadsheet.Sheet} flexSheet - The flex absences sheet
 * @param {Map<string, Array>} contactMap - Map of student ID to [student email(N), guardian1 email(G), guardian2 email(K)]
 */
function addContactInfo(flexSheet, contactMap) {
  try {
    const lastRow = flexSheet.getLastRow();
    
    if (lastRow <= 2) {
      Logger.log('No data rows in flex absences sheet to enrich (need at least 3 rows for FlexiSched data)');
      return;
    }
    
    // Get student IDs from column A (starting from row 3, skipping FlexiSched headers in rows 1-2)
    const studentIdRange = flexSheet.getRange(3, 1, lastRow - 2, 1);
    const studentIds = studentIdRange.getValues();
    
    // Prepare contact info to write to columns P-R
    const contactInfo = [];
    let foundCount = 0;
    let notFoundCount = 0;
    
    Logger.log(`Processing ${studentIds.length} students for contact info`);
    Logger.log(`Contact map has ${contactMap.size} entries`);
    
    // Show first few student IDs we're looking for
    Logger.log('Student IDs from flex sheet (first few):');
    for (let i = 0; i < Math.min(3, studentIds.length); i++) {
      const studentId = String(studentIds[i][0]).trim();
      Logger.log(`  "${studentId}" (type: ${typeof studentIds[i][0]})`);
    }
    
    for (let i = 0; i < studentIds.length; i++) {
      const studentId = String(studentIds[i][0]).trim();
      
      if (!studentId || studentId === '') {
        // Empty row, leave contact info empty
        contactInfo.push(['', '', '']);
        continue;
      }
      
      // Lookup student in contact map
      if (contactMap.has(studentId)) {
        const data = contactMap.get(studentId);
        // data[0] = Column N (Student Email) → goes to Column O
        // data[1] = Column G (Guardian 1 Email) → goes to Column P  
        // data[2] = Column K (Guardian 2 Email) → goes to Column Q
        const studentEmail = data[0] || '';    // Column N → Column O
        const guardian1Email = data[1] || '';  // Column G → Column P
        const guardian2Email = data[2] || '';  // Column K → Column Q
        contactInfo.push([studentEmail, guardian1Email, guardian2Email]);
        foundCount++;
        
        // Log first few matches for debugging
        if (foundCount <= 3) {
          Logger.log(`Student ${studentId} found: O=${studentEmail}, P=${guardian1Email}, Q=${guardian2Email}`);
        }
      } else {
        // Student not found in contact info - leave empty
        contactInfo.push(['', '', '']);
        notFoundCount++;
        
        // Log first few misses for debugging
        if (notFoundCount <= 3) {
          Logger.log(`Student ${studentId} NOT found in contact map`);
        }
      }
    }
    
    Logger.log(`Contact info results: ${foundCount} found, ${notFoundCount} not found`);
    
    // Ensure headers are set for columns O-Q before writing data
    const headerO = flexSheet.getRange('O2');
    const headerP = flexSheet.getRange('P2');
    const headerQ = flexSheet.getRange('Q2');
    
    if (!headerO.getValue() || headerO.getValue() !== CONFIG.flexAbsencesHeaders.O) {
      headerO.setValue(CONFIG.flexAbsencesHeaders.O);
      headerO.setFontWeight('bold');
      Logger.log('Set header for column O: ' + CONFIG.flexAbsencesHeaders.O);
    }
    
    if (!headerP.getValue() || headerP.getValue() !== CONFIG.flexAbsencesHeaders.P) {
      headerP.setValue(CONFIG.flexAbsencesHeaders.P);
      headerP.setFontWeight('bold');
      Logger.log('Set header for column P: ' + CONFIG.flexAbsencesHeaders.P);
    }
    
    if (!headerQ.getValue() || headerQ.getValue() !== CONFIG.flexAbsencesHeaders.Q) {
      headerQ.setValue(CONFIG.flexAbsencesHeaders.Q);
      headerQ.setFontWeight('bold');
      Logger.log('Set header for column Q: ' + CONFIG.flexAbsencesHeaders.Q);
    }
    
    // Write contact info to columns O-Q (columns 15-17) starting from row 3
    const targetRange = flexSheet.getRange(3, 15, contactInfo.length, 3);
    targetRange.setValues(contactInfo);
    
    Logger.log(`Added contact info to ${contactInfo.length} rows in flex absences sheet`);
    
  } catch (error) {
    Logger.log(`Error adding contact info: ${error.message}`);
    throw error;
  }
}

/**
 * Identifies students who skipped their flex class (have #N/A in column M)
 * Requirements: 5.1, 5.2, 5.3, 5.4
 * @param {GoogleAppsScript.Spreadsheet.Sheet} flexSheet - The flex absences sheet
 * @returns {Array<Object>} Array of skipper objects with rowIndex and rowData (columns A-Q)
 */
function identifySkippers(flexSheet) {
  try {
    const skippers = [];
    const lastRow = flexSheet.getLastRow();
    
    if (lastRow <= 2) {
      Logger.log('No data rows in flex absences sheet to check for skippers (need at least 3 rows for FlexiSched data)');
      return skippers;
    }
    
    // Get all data from columns A-Q (columns 1-17) starting from row 3 (skipping FlexiSched headers)
    const dataRange = flexSheet.getRange(3, 1, lastRow - 2, 17);
    const data = dataRange.getValues();
    
    // Check each row for #N/A in column M (index 12 in the array)
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const attendanceCode = String(row[12]).trim();  // Column M is index 12
      
      // Check if attendance code is #N/A
      if (attendanceCode === '#N/A') {
        skippers.push({
          rowIndex: i + 3,  // Actual row number in sheet (accounting for two header rows)
          rowData: row      // All data from columns A-Q
        });
      }
    }
    
    Logger.log(`Identified ${skippers.length} skippers in flex absences sheet`);
    return skippers;
    
  } catch (error) {
    Logger.log(`Error identifying skippers: ${error.message}`);
    throw error;
  }
}

/**
 * Copies skipper data to the Mail Out sheet
 * Clears existing Mail Out sheet data before adding new skippers
 * Requirements: 5.1, 5.2, 5.3, 5.4
 * @param {GoogleAppsScript.Spreadsheet.Sheet} flexSheet - The flex absences sheet (for header reference)
 * @param {Array<Object>} skipperRows - Array of skipper objects with rowData
 */
function copySkippersToMailOut(flexSheet, skipperRows) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const mailOutSheet = getOrCreateSheet(CONFIG.sheetNames.mailOut);
    
    // Clear existing data in Mail Out sheet (from row 2 onward, preserve headers)
    clearSheetData(mailOutSheet, 2);
    
    if (skipperRows.length === 0) {
      Logger.log('No skippers to copy to Mail Out sheet');
      return;
    }
    
    // Set correct Mail Out headers (17 columns, not 18)
    const mailOutHeaders = [
      'ID', 'First Name', 'Last Name', 'Grad Year', 'Flex Name', 'Type', 'Request', 
      'Day', 'Period', 'Date', 'Flex Status', 'Comment', 'Brennan Attendance', 
      '2nd Period Teacher', 'Student Email', 'Guardian 1 Email', 'Guardian 2 Email'
    ];
    
    // Check if headers need to be set
    const currentHeaders = mailOutSheet.getRange(1, 1, 1, 17).getValues()[0];
    const hasHeaders = currentHeaders.some(header => header !== '');
    
    if (!hasHeaders) {
      mailOutSheet.getRange(1, 1, 1, 17).setValues([mailOutHeaders]);
      mailOutSheet.getRange(1, 1, 1, 17).setFontWeight('bold');
      Logger.log('Set headers in Mail Out sheet');
    }
    
    // Prepare skipper data for writing - both sheets now have 17 columns
    const skipperData = skipperRows.map(skipper => {
      const flexRow = skipper.rowData; // 17 columns (A-Q)
      
      // Direct mapping since both sheets have same structure:
      // A-L: FlexiSched data + Comment
      // M: Attendance Code → Brennan Attendance  
      // N: 2nd Period Teacher
      // O: Student Email
      // P: Guardian 1 Email
      // Q: Guardian 2 Email
      
      return flexRow; // Direct copy since structures match
    });
    
    // Write skipper data to Mail Out sheet starting at row 2
    writeDataToSheet(mailOutSheet, skipperData, 2);
    
    Logger.log(`Copied ${skipperRows.length} skippers to Mail Out sheet`);
    
  } catch (error) {
    Logger.log(`Error copying skippers to Mail Out sheet: ${error.message}`);
    throw error;
  }
}

/**
 * Checks if enrichment headers are missing in row 2 and restores them if needed
 * This handles the case where users paste FlexiSched data and overwrite the enrichment headers
 * Requirements: 7.5, 7.6
 * @param {GoogleAppsScript.Spreadsheet.Sheet} flexSheet - The flex absences sheet
 */
function ensureEnrichmentHeaders(flexSheet) {
  try {
    Logger.log('Ensuring enrichment headers are present...');
    const lastColumn = flexSheet.getLastColumn();
    Logger.log(`Sheet has ${lastColumn} columns`);
    
    // Get current headers in row 2, columns M-Q (columns 13-17)
    // We'll read what's there, even if the sheet doesn't have all columns yet
    const headerRange = flexSheet.getRange(2, 13, 1, 5);
    const currentHeaders = headerRange.getValues()[0];
    Logger.log('Current headers in M2:Q2:', currentHeaders);
    
    // Check if any of the enrichment headers are missing or empty
    const expectedHeaders = [
      CONFIG.flexAbsencesHeaders.M,  // 'Attendance Code'
      CONFIG.flexAbsencesHeaders.N,  // '2nd Period Teacher'
      CONFIG.flexAbsencesHeaders.O,  // 'Student Email'
      CONFIG.flexAbsencesHeaders.P,  // 'Guardian 1 Email'
      CONFIG.flexAbsencesHeaders.Q   // 'Guardian 2 Email'
    ];
    
    let needsUpdate = false;
    for (let i = 0; i < expectedHeaders.length; i++) {
      const currentHeader = String(currentHeaders[i] || '').trim();
      const expectedHeader = expectedHeaders[i];
      
      if (currentHeader !== expectedHeader) {
        needsUpdate = true;
        Logger.log(`Header mismatch in column ${String.fromCharCode(77 + i)}: expected "${expectedHeader}", found "${currentHeader}"`);
      }
    }
    
    if (needsUpdate) {
      Logger.log('Restoring missing enrichment headers in row 2, columns M-Q');
      Logger.log('Expected headers:', expectedHeaders);
      
      // Set the enrichment headers in row 2
      headerRange.setValues([expectedHeaders]);
      headerRange.setFontWeight('bold');
      
      Logger.log('Successfully restored enrichment headers');
      
      // Verify they were set
      const verifyHeaders = headerRange.getValues()[0];
      Logger.log('Verified headers after setting:', verifyHeaders);
    } else {
      Logger.log('Enrichment headers are already present and correct');
    }
    
  } catch (error) {
    Logger.log(`Error ensuring enrichment headers: ${error.message}`);
    // Don't throw error - this is not critical, enrichment can still proceed
  }
}

/**
 * Test function to manually set up headers in the flex absences sheet
 * This can be used for debugging header setup issues
 */
function testSetupHeaders() {
  try {
    Logger.log('Starting manual header setup test...');
    
    const flexSheet = getFlexAbsencesSheet();
    if (!flexSheet) {
      showErrorMessage('Could not find flex absences sheet. Please create one first.');
      return;
    }
    
    Logger.log(`Found flex absences sheet: ${flexSheet.getName()}`);
    
    // Force setup headers
    setupFlexAbsencesHeaders(flexSheet);
    
    // Also force ensure headers
    ensureEnrichmentHeaders(flexSheet);
    
    showSuccessMessage('Header setup test completed. Check the execution log for details.');
    
  } catch (error) {
    Logger.log(`Error in testSetupHeaders: ${error.message}`);
    showErrorMessage(`Header setup test failed: ${error.message}`);
  }
}

/**
 * Main orchestration function for enriching flex absences data
 * Adds attendance codes, teacher names, and contact info, then identifies skippers
 * Requirements: 3.4, 4.3, 5.5
 */
function enrichFlexAbsences() {
  try {
    Logger.log('Starting flex absences enrichment...');
    
    // Get the flex absences sheet
    const flexSheet = getFlexAbsencesSheet();
    
    if (!flexSheet) {
      showErrorMessage('Could not find flex absences sheet. Please create a sheet with a name like "11.3 flex absences".');
      return;
    }
    
    Logger.log(`Found flex absences sheet: ${flexSheet.getName()}`);
    
    // Check if there's data to enrich (need at least 3 rows: 2 FlexiSched headers + 1 data row)
    const lastRow = flexSheet.getLastRow();
    if (lastRow <= 2) {
      showErrorMessage('No data found in flex absences sheet. Please paste FlexiSched data first (need at least 3 rows: 2 headers + data).');
      return;
    }
    
    // Ensure enrichment headers are present in row 2 (restore if overwritten by FlexiSched paste)
    ensureEnrichmentHeaders(flexSheet);
    
    // Get the three source sheets
    const attendanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheetNames.attendance);
    const coursesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheetNames.courses);
    const contactsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheetNames.contacts);
    
    // Verify all source sheets exist
    if (!attendanceSheet) {
      showErrorMessage('BHS attendance sheet not found. Please import COGNOS reports first.');
      return;
    }
    
    if (!coursesSheet) {
      showErrorMessage('2nd period default sheet not found. Please import COGNOS reports first.');
      return;
    }
    
    if (!contactsSheet) {
      showErrorMessage('contact info sheet not found. Please import COGNOS reports first.');
      return;
    }
    
    // Build lookup maps from source sheets
    Logger.log('Building student lookup maps...');
    
    // Build attendance map: student ID -> [attendance code from column K]
    // Column K is index 10 (zero-based)
    const attendanceIdColumn = getStudentIdColumn(attendanceSheet);
    const attendanceMap = buildStudentMap(attendanceSheet, attendanceIdColumn, [10]);
    
    // Build teacher map: student ID -> [teacher name from column G]
    // Column G is index 6 (zero-based)
    const coursesIdColumn = getStudentIdColumn(coursesSheet);
    const teacherMap = buildStudentMap(coursesSheet, coursesIdColumn, [6]);
    
    // Build contact map: student ID -> [student email (N), guardian1 email (G), guardian2 email (K)]
    // Student ID is in column B (index 1), not A
    // Column N is index 13, Column G is index 6, Column K is index 10 (zero-based)
    // Mapping: N→P, G→Q, K→R
    const contactsIdColumn = 1;  // Column B for contact sheet
    Logger.log('Building contact map from columns: N(13)→P, G(6)→Q, K(10)→R');
    
    // Debug: Check contact sheet data
    const contactLastRow = contactsSheet.getLastRow();
    const contactLastCol = contactsSheet.getLastColumn();
    Logger.log(`Contact sheet has ${contactLastRow} rows and ${contactLastCol} columns`);
    
    if (contactLastRow > 1) {
      // Show first few rows of contact data for debugging
      const sampleRange = contactsSheet.getRange(1, 1, Math.min(4, contactLastRow), Math.min(13, contactLastCol));
      const sampleData = sampleRange.getValues();
      Logger.log('Contact sheet sample data (first few rows):');
      sampleData.forEach((row, index) => {
        Logger.log(`  Row ${index + 1}: [${row.slice(0, 5).join(', ')}...] (showing first 5 columns)`);
      });
      
      // Show student ID column specifically
      if (contactLastRow > 1) {
        const idRange = contactsSheet.getRange(2, 1, Math.min(3, contactLastRow - 1), 1);
        const idData = idRange.getValues();
        Logger.log('Student IDs in contact sheet (first few):');
        idData.forEach((row, index) => {
          Logger.log(`  Row ${index + 2}: "${row[0]}" (type: ${typeof row[0]})`);
        });
      }
    }
    
    const contactMap = buildStudentMap(contactsSheet, contactsIdColumn, [13, 6, 10]);
    
    // Log some sample contact data for debugging
    if (contactMap.size > 0) {
      const sampleEntries = Array.from(contactMap.entries()).slice(0, 3);
      Logger.log('Sample contact map entries:');
      sampleEntries.forEach(([id, emails]) => {
        Logger.log(`  ${id}: [${emails.join(', ')}]`);
      });
    }
    
    Logger.log(`Built maps: ${attendanceMap.size} attendance records, ${teacherMap.size} teacher records, ${contactMap.size} contact records`);
    
    // Enrich the flex absences sheet
    Logger.log('Enriching flex absences data...');
    
    // Add attendance codes to column L (or #N/A if not found)
    addAttendanceCodes(flexSheet, attendanceMap);
    
    // Add teacher names to column M
    addTeacherNames(flexSheet, teacherMap);
    
    // Add contact info to columns O-Q
    addContactInfo(flexSheet, contactMap);
    
    Logger.log('Enrichment complete. Identifying skippers...');
    
    // Identify students with #N/A attendance codes (skippers)
    const skippers = identifySkippers(flexSheet);
    
    // Copy skippers to Mail Out sheet
    copySkippersToMailOut(flexSheet, skippers);
    
    // Display success message
    const message = `Enrichment complete!\n\nIdentified ${skippers.length} skipper(s) and copied to Mail Out sheet.`;
    showSuccessMessage(message);
    
    Logger.log(`Enrichment complete. ${skippers.length} skippers identified.`);
    
  } catch (error) {
    Logger.log(`Error in enrichFlexAbsences: ${error.message}`);
    showErrorMessage(`Failed to enrich flex absences data: ${error.message}`);
  }
}

// ============================================================================
// COMMENT SYNC SERVICE MODULE
// ============================================================================

/**
 * Builds a map of student ID to comment from the Mail Out sheet
 * Requirements: 6.1, 6.2, 6.3, 6.4, 6.5
 * @param {GoogleAppsScript.Spreadsheet.Sheet} mailOutSheet - The Mail Out sheet
 * @param {number} idColumn - Zero-based column index containing student ID
 * @returns {Map<string, string>} Map of student ID to comment text
 */
function buildCommentMap(mailOutSheet, idColumn) {
  try {
    const commentMap = new Map();
    
    // Get all data from the sheet
    const lastRow = mailOutSheet.getLastRow();
    const lastColumn = mailOutSheet.getLastColumn();
    
    if (lastRow <= 1) {
      Logger.log(`Mail Out sheet has no data rows (only headers or empty)`);
      return commentMap;
    }
    
    if (lastColumn === 0) {
      Logger.log(`Mail Out sheet has no columns`);
      return commentMap;
    }
    
    // Get data starting from row 2 (skip headers in row 1)
    // We need columns A (student ID) and L (comments, column 12)
    const dataRange = mailOutSheet.getRange(2, 1, lastRow - 1, Math.max(lastColumn, 12));
    const data = dataRange.getValues();
    
    // Build the map - column L is index 11 (zero-based)
    const commentColumnIndex = 11;
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const studentId = String(row[idColumn]).trim();
      
      // Skip rows with empty student ID
      if (!studentId || studentId === '') {
        continue;
      }
      
      // Get comment from column L (index 11)
      const comment = commentColumnIndex < row.length ? String(row[commentColumnIndex]) : '';
      
      // Only add to map if there's a comment (skip empty comments)
      if (comment && comment.trim() !== '') {
        commentMap.set(studentId, comment);
      }
    }
    
    Logger.log(`Built comment map from Mail Out sheet: ${commentMap.size} students with comments`);
    return commentMap;
    
  } catch (error) {
    Logger.log(`Error building comment map from Mail Out sheet: ${error.message}`);
    throw error;
  }
}

/**
 * Updates comments in column L (FlexiSched Comment) of the Flex Absences sheet from the comment map
 * Preserves existing comments for students not in the map
 * Requirements: 6.1, 6.2, 6.3, 6.4, 6.5
 * @param {GoogleAppsScript.Spreadsheet.Sheet} flexSheet - The flex absences sheet
 * @param {Map<string, string>} commentMap - Map of student ID to comment text
 */
function updateFlexComments(flexSheet, commentMap) {
  try {
    const lastRow = flexSheet.getLastRow();
    
    if (lastRow <= 2) {
      Logger.log('No data rows in flex absences sheet to update comments (need at least 3 rows for FlexiSched data)');
      return;
    }
    
    // Get student IDs from column A (starting from row 3, skipping FlexiSched headers in rows 1-2)
    const studentIdRange = flexSheet.getRange(3, 1, lastRow - 2, 1);
    const studentIds = studentIdRange.getValues();
    
    // Get existing comments from column L (column 12) starting from row 3 - FlexiSched Comment column
    const existingCommentsRange = flexSheet.getRange(3, 12, lastRow - 2, 1);
    const existingComments = existingCommentsRange.getValues();
    
    // Prepare updated comments to write to column L (FlexiSched Comment column)
    const updatedComments = [];
    let updatedCount = 0;
    let unmatchedCount = 0;
    
    for (let i = 0; i < studentIds.length; i++) {
      const studentId = String(studentIds[i][0]).trim();
      
      if (!studentId || studentId === '') {
        // Empty row, preserve existing comment
        updatedComments.push([existingComments[i][0]]);
        continue;
      }
      
      // Check if student has a comment in the comment map
      if (commentMap.has(studentId)) {
        const newComment = commentMap.get(studentId);
        updatedComments.push([newComment]);
        updatedCount++;
      } else {
        // Student not in Mail Out sheet, preserve existing comment
        updatedComments.push([existingComments[i][0]]);
        unmatchedCount++;
      }
    }
    
    // Write updated comments to column L (column 12) starting from row 3 - FlexiSched Comment column
    const targetRange = flexSheet.getRange(3, 12, updatedComments.length, 1);
    targetRange.setValues(updatedComments);
    
    Logger.log(`Updated comments in flex absences sheet: ${updatedCount} comments synced, ${unmatchedCount} students not in Mail Out`);
    return updatedCount;
    
  } catch (error) {
    Logger.log(`Error updating flex comments: ${error.message}`);
    throw error;
  }
}

/**
 * Main orchestration function for syncing comments from Mail Out sheet to Flex Absences sheet
 * Copies comments from column L of Mail Out to column L (FlexiSched Comment) of Flex Absences
 * Requirements: 6.1, 6.2, 6.3, 6.4, 6.5
 */
function syncComments() {
  try {
    Logger.log('Starting comment sync from Mail Out to Flex Absences...');
    
    // Get the flex absences sheet
    const flexSheet = getFlexAbsencesSheet();
    
    if (!flexSheet) {
      showErrorMessage('Could not find flex absences sheet. Please create a sheet with a name like "11.3 flex absences".');
      return;
    }
    
    Logger.log(`Found flex absences sheet: ${flexSheet.getName()}`);
    
    // Check if there's data in flex absences sheet (need at least 3 rows for FlexiSched data)
    const flexLastRow = flexSheet.getLastRow();
    if (flexLastRow <= 2) {
      showErrorMessage('No data found in flex absences sheet (need at least 3 rows: 2 headers + data).');
      return;
    }
    
    // Get the Mail Out sheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const mailOutSheet = spreadsheet.getSheetByName(CONFIG.sheetNames.mailOut);
    
    if (!mailOutSheet) {
      showErrorMessage('Mail Out sheet not found. Please run enrichment first.');
      return;
    }
    
    // Check if there's data in Mail Out sheet
    const mailOutLastRow = mailOutSheet.getLastRow();
    if (mailOutLastRow <= 1) {
      showErrorMessage('No data found in Mail Out sheet. Please run enrichment first.');
      return;
    }
    
    Logger.log('Building comment map from Mail Out sheet...');
    
    // Build comment map from Mail Out sheet
    // Student ID is in column A (index 0)
    const idColumn = 0;
    const commentMap = buildCommentMap(mailOutSheet, idColumn);
    
    if (commentMap.size === 0) {
      showSuccessMessage('No comments found in Mail Out sheet to sync.');
      return;
    }
    
    Logger.log(`Found ${commentMap.size} comments in Mail Out sheet`);
    
    // Update comments in Flex Absences sheet
    Logger.log('Updating comments in Flex Absences sheet...');
    const updatedCount = updateFlexComments(flexSheet, commentMap);
    
    // Display success message
    const message = `Comment sync complete!\n\nSynced ${updatedCount} comment(s) from Mail Out to Flex Absences sheet.`;
    showSuccessMessage(message);
    
    Logger.log(`Comment sync complete. ${updatedCount} comments synced.`);
    
  } catch (error) {
    Logger.log(`Error in syncComments: ${error.message}`);
    showErrorMessage(`Failed to sync comments: ${error.message}`);
  }
}
