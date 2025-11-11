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
 * Adds attendance codes from BHS attendance sheet to flex absences sheet column L
 * If no matching student found, adds #N/A
 * Requirements: 4.3, 4.4, 4.5, 4.6
 * @param {GoogleAppsScript.Spreadsheet.Sheet} flexSheet - The flex absences sheet
 * @param {Map<string, Array>} attendanceMap - Map of student ID to [attendance code]
 */
function addAttendanceCodes(flexSheet, attendanceMap) {
  try {
    const lastRow = flexSheet.getLastRow();
    
    if (lastRow <= 1) {
      Logger.log('No data rows in flex absences sheet to enrich');
      return;
    }
    
    // Get student IDs from column A (starting from row 2, skipping headers)
    const studentIdRange = flexSheet.getRange(2, 1, lastRow - 1, 1);
    const studentIds = studentIdRange.getValues();
    
    // Prepare attendance codes to write to column L
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
    
    // Write attendance codes to column L (column 12)
    const targetRange = flexSheet.getRange(2, 12, attendanceCodes.length, 1);
    targetRange.setValues(attendanceCodes);
    
    Logger.log(`Added attendance codes to ${attendanceCodes.length} rows in flex absences sheet`);
    
  } catch (error) {
    Logger.log(`Error adding attendance codes: ${error.message}`);
    throw error;
  }
}

/**
 * Adds teacher names from 2nd period default sheet to flex absences sheet column M
 * Requirements: 4.3, 4.4, 4.5, 4.6
 * @param {GoogleAppsScript.Spreadsheet.Sheet} flexSheet - The flex absences sheet
 * @param {Map<string, Array>} teacherMap - Map of student ID to [teacher name]
 */
function addTeacherNames(flexSheet, teacherMap) {
  try {
    const lastRow = flexSheet.getLastRow();
    
    if (lastRow <= 1) {
      Logger.log('No data rows in flex absences sheet to enrich');
      return;
    }
    
    // Get student IDs from column A (starting from row 2, skipping headers)
    const studentIdRange = flexSheet.getRange(2, 1, lastRow - 1, 1);
    const studentIds = studentIdRange.getValues();
    
    // Prepare teacher names to write to column M
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
    
    // Write teacher names to column M (column 13)
    const targetRange = flexSheet.getRange(2, 13, teacherNames.length, 1);
    targetRange.setValues(teacherNames);
    
    Logger.log(`Added teacher names to ${teacherNames.length} rows in flex absences sheet`);
    
  } catch (error) {
    Logger.log(`Error adding teacher names: ${error.message}`);
    throw error;
  }
}

/**
 * Adds contact information (emails) from contact info sheet to flex absences sheet columns O-Q
 * Requirements: 4.3, 4.4, 4.5, 4.6
 * @param {GoogleAppsScript.Spreadsheet.Sheet} flexSheet - The flex absences sheet
 * @param {Map<string, Array>} contactMap - Map of student ID to [student email, guardian1 email, guardian2 email]
 */
function addContactInfo(flexSheet, contactMap) {
  try {
    const lastRow = flexSheet.getLastRow();
    
    if (lastRow <= 1) {
      Logger.log('No data rows in flex absences sheet to enrich');
      return;
    }
    
    // Get student IDs from column A (starting from row 2, skipping headers)
    const studentIdRange = flexSheet.getRange(2, 1, lastRow - 1, 1);
    const studentIds = studentIdRange.getValues();
    
    // Prepare contact info to write to columns O-Q
    const contactInfo = [];
    
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
        const studentEmail = data[0] || '';    // First element is student email
        const guardian1Email = data[1] || '';  // Second element is guardian 1 email
        const guardian2Email = data[2] || '';  // Third element is guardian 2 email
        contactInfo.push([studentEmail, guardian1Email, guardian2Email]);
      } else {
        // Student not found in contact info - leave empty
        contactInfo.push(['', '', '']);
      }
    }
    
    // Write contact info to columns O-Q (columns 15-17)
    const targetRange = flexSheet.getRange(2, 15, contactInfo.length, 3);
    targetRange.setValues(contactInfo);
    
    Logger.log(`Added contact info to ${contactInfo.length} rows in flex absences sheet`);
    
  } catch (error) {
    Logger.log(`Error adding contact info: ${error.message}`);
    throw error;
  }
}

/**
 * Identifies students who skipped their flex class (have #N/A in column L)
 * Requirements: 5.1, 5.2, 5.3, 5.4
 * @param {GoogleAppsScript.Spreadsheet.Sheet} flexSheet - The flex absences sheet
 * @returns {Array<Object>} Array of skipper objects with rowIndex and rowData (columns A-Q)
 */
function identifySkippers(flexSheet) {
  try {
    const skippers = [];
    const lastRow = flexSheet.getLastRow();
    
    if (lastRow <= 1) {
      Logger.log('No data rows in flex absences sheet to check for skippers');
      return skippers;
    }
    
    // Get all data from columns A-Q (columns 1-17) starting from row 2
    const dataRange = flexSheet.getRange(2, 1, lastRow - 1, 17);
    const data = dataRange.getValues();
    
    // Check each row for #N/A in column L (index 11 in the array)
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const attendanceCode = String(row[11]).trim();  // Column L is index 11
      
      // Check if attendance code is #N/A
      if (attendanceCode === '#N/A') {
        skippers.push({
          rowIndex: i + 2,  // Actual row number in sheet (accounting for header row)
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
    
    // Get headers from flex absences sheet (columns A-Q)
    const flexHeaders = flexSheet.getRange(1, 1, 1, 17).getValues()[0];
    
    // Set headers in Mail Out sheet if not already set
    const mailOutHeaders = mailOutSheet.getRange(1, 1, 1, 17).getValues()[0];
    const hasHeaders = mailOutHeaders.some(header => header !== '');
    
    if (!hasHeaders) {
      mailOutSheet.getRange(1, 1, 1, 17).setValues([flexHeaders]);
      mailOutSheet.getRange(1, 1, 1, 17).setFontWeight('bold');
      Logger.log('Set headers in Mail Out sheet');
    }
    
    // Prepare skipper data for writing (extract rowData from each skipper object)
    const skipperData = skipperRows.map(skipper => skipper.rowData);
    
    // Write skipper data to Mail Out sheet starting at row 2
    writeDataToSheet(mailOutSheet, skipperData, 2);
    
    Logger.log(`Copied ${skipperRows.length} skippers to Mail Out sheet`);
    
  } catch (error) {
    Logger.log(`Error copying skippers to Mail Out sheet: ${error.message}`);
    throw error;
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
    
    // Check if there's data to enrich
    const lastRow = flexSheet.getLastRow();
    if (lastRow <= 1) {
      showErrorMessage('No data found in flex absences sheet. Please paste FlexiSched data first.');
      return;
    }
    
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
    
    // Build contact map: student ID -> [student email (M), guardian1 email (F), guardian2 email (J)]
    // Column M is index 12, Column F is index 5, Column J is index 9 (zero-based)
    const contactsIdColumn = getStudentIdColumn(contactsSheet);
    const contactMap = buildStudentMap(contactsSheet, contactsIdColumn, [12, 5, 9]);
    
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

function syncComments() {
  showErrorMessage('Comment sync functionality not yet implemented.');
}
