/**
 * COGNOS Report Consolidation System
 * Main entry point for Google Apps Script automation
 */

// Email configuration for the three COGNOS reports
const EMAIL_CONFIG = {
  attendance: {
    subject: 'My ATT - Attendance Bulletin (1)',
    sheetName: 'BHS attendance',
    attendanceCodeColumn: 'K',  // Attendance code column
    periodColumn: 'J',           // Period column for filtering
    periodFilter: '02'           // Only import 2nd period rows (stored as text)
  },
  courses: {
    subject: 'My Student CY List - Courses, Teacher & Room',
    sheetName: '2nd period default',
    teacherColumn: 'G',  // Teacher name column (after reordering)
    originalColumns: ['Student Name', 'Student Id', 'Grade', '9th Grd Entry', 'Period', 'Description', 'Room', 'Instructor', 'Instructor ID', 'Instructor Email'],
    targetColumns: ['Student Id', 'Student Name', 'Grade', 'Period', 'Description', 'Room', 'Instructor', 'Instructor Id', 'Instructor Email']
  },
  contacts: {
    subject: 'My Student CY List - Student Email/Contact Info - Next Year Option (1)',
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
      flexiSchedData: 'A:K',      // User-pasted FlexiSched data
      attendanceCode: 'L',         // From BHS attendance or #N/A
      teacherName: 'M',            // From 2nd period default
      comments: 'N',               // User-entered after FormMule
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
    L: 'Attendance Code',
    M: '2nd Period Teacher',
    N: 'Comments',
    O: 'Student Email',
    P: 'Guardian 1 Email',
    Q: 'Guardian 2 Email'
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
    .addToUi();
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
 * Sets up column headers for a flex absences sheet
 * Sets headers for columns A-Q (A-K for FlexiSched data, L-Q for enriched data)
 * Requirements: 7.5, 7.6
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to set up headers for
 */
function setupFlexAbsencesHeaders(sheet) {
  // Set FlexiSched data headers (columns A-K)
  const flexiSchedHeaders = ['ID', 'First Name', 'Last Name', 'Grade', 'Flex Name', 
                              'Type', 'Request', 'Day', 'Period', 'Date', 'Flex Status'];
  sheet.getRange('A1:K1').setValues([flexiSchedHeaders]);
  
  // Set enriched data headers (columns L-Q)
  const headers = CONFIG.flexAbsencesHeaders;
  sheet.getRange('L1').setValue(headers.L);  // Attendance Code
  sheet.getRange('M1').setValue(headers.M);  // 2nd Period Teacher
  sheet.getRange('N1').setValue(headers.N);  // Comments
  sheet.getRange('O1').setValue(headers.O);  // Student Email
  sheet.getRange('P1').setValue(headers.P);  // Guardian 1 Email
  sheet.getRange('Q1').setValue(headers.Q);  // Guardian 2 Email
  
  // Format entire header row (bold)
  sheet.getRange('A1:Q1').setFontWeight('bold');
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

// Placeholder functions for future implementation
function enrichFlexAbsences() {
  showErrorMessage('Enrichment functionality not yet implemented.');
}

function syncComments() {
  showErrorMessage('Comment sync functionality not yet implemented.');
}
