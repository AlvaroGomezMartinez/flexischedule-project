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
