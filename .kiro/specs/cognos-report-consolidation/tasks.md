# Implementation Plan

- [x] 1. Set up CLASP development environment
  - Initialize CLASP project for pushing to Google Apps Script
  - Create .claspignore file to exclude non-GAS files (spec files, README, etc.)
  - Configure CLASP to connect to the bound Google Sheets spreadsheet
  - _Requirements: Setup_

- [x] 2. Set up project structure and configuration
  - Create Code.gs file as the main entry point
  - Define CONFIG constant with email patterns, sheet names, and column mappings
  - Define EMAIL_CONFIG constant with the three COGNOS report configurations
  - _Requirements: 1.1, 2.1, 2.2, 2.3_

- [x] 3. Implement UI Module with custom menu
  - Write `onOpen()` function to create custom menu when spreadsheet opens
  - Create menu structure with four items: "Create today's flex absences sheet", "Import COGNOS Reports from GMail", "Add data to Flex Absences sheet", "Sync Comments from Mail Out sheet"
  - Write `showSuccessMessage(message)` function to display toast notifications
  - Write `showErrorMessage(message)` function to display error alerts
  - _Requirements: 4.1, 4.2, 7.1_

- [x] 4. Implement Sheet Creation Module
- [x] 4.1 Create date formatting and sheet creation functions
  - Write `formatDateForSheetName(date)` function to format date as "M.D" (e.g., "11.3")
  - Write `createTodaysFlexAbsencesSheet()` function to create new sheet with today's date
  - Implement duplicate sheet name checking
  - _Requirements: 7.2, 7.3, 7.4_

- [x] 4.2 Set up flex absences sheet headers
  - Write `setupFlexAbsencesHeaders(sheet)` function to set column headers A-Q
  - Set headers: L="Attendance Code", M="2nd Period Teacher", N="Comments", O="Student Email", P="Guardian 1 Email", Q="Guardian 2 Email"
  - _Requirements: 7.5, 7.6_

- [x] 5. Implement Sheet Manager Module
- [x] 5.1 Create basic sheet management functions
  - Write `getOrCreateSheet(sheetName)` function to return existing sheet or create new one
  - Write `getFlexAbsencesSheet()` function to find sheet matching date pattern and "flex absences" suffix
  - Write `clearSheetData(sheet, startRow)` function to clear data from specified row onward
  - _Requirements: 2.5, 3.1_

- [x] 5.2 Create sheet data manipulation functions
  - Write `writeDataToSheet(sheet, data, startRow)` function to write 2D array to sheet
  - Write `addNoteToCell(sheet, cell, note)` function to add notes to cells
  - Write `setSheetHeaders(sheet, headers)` function to set up column headers
  - _Requirements: 2.6, 2.7_

- [x] 6. Implement Gmail Service Module
- [x] 6.1 Create user email detection and Gmail search functions
  - Write `getUserEmail()` function using `Session.getActiveUser().getEmail()`
  - Write `searchCognosEmails()` function to search Gmail with `from:` filter for user's own email
  - Implement search queries for all three COGNOS report subjects
  - _Requirements: 1.1, 8.1, 8.2, 8.3_

- [x] 6.2 Create email attachment retrieval functions
  - Write `getAttachments(messageId)` function to retrieve Excel attachments from emails
  - Write `identifyReportType(email)` function to determine which of the three reports an email contains
  - Implement error handling for missing emails and attachments
  - _Requirements: 1.2, 1.3, 1.4_

- [x] 7. Implement Excel Converter Module
- [x] 7.1 Create Excel parsing functions
  - Write `convertExcelToArray(blob)` function to convert Excel blob to 2D array
  - Write `parseExcelData(excelBlob)` function to extract data and headers from Excel files
  - _Requirements: 2.4_

- [x] 7.2 Create data transformation functions
  - Write `filterByPeriod(data, periodColumn, periodValue)` function to filter rows where column J = "02"
  - Write `reorderColumns(data, columnMapping)` function to reorder columns and exclude "9th Grd Entry" for courses report
  - _Requirements: 2.1, 2.2_

- [x] 8. Implement COGNOS report import orchestration
  - Create main `importCognosReports()` function that orchestrates the import workflow
  - Search Gmail for three COGNOS reports using user's email
  - Extract Excel attachments from each email
  - Apply special processing (filtering, reordering) based on report type
  - Import data to respective sheets: BHS attendance, 2nd period default, contact info
  - Add timestamp notes to cell A1 of each imported sheet
  - Display success message with count of reports imported
  - _Requirements: 1.5, 1.6, 2.1, 2.2, 2.3, 2.5, 2.6, 2.7_

- [x] 9. Implement Enrichment Service Module
- [x] 9.1 Create student lookup map building functions
  - Write `getStudentIdColumn(sheet)` function to identify which column contains student ID
  - Write `buildStudentMap(sheet, idColumn, dataColumns)` function to create student ID → data lookup map
  - Build maps from BHS attendance (attendance codes), 2nd period default (teacher names), and contact info (emails)
  - _Requirements: 4.3, 4.5, 4.6_

- [x] 9.2 Create data enrichment functions
  - Write `addAttendanceCodes(flexSheet, attendanceMap)` function to add attendance codes or #N/A to column L
  - Write `addTeacherNames(flexSheet, teacherMap)` function to add teacher names to column M
  - Write `addContactInfo(flexSheet, contactMap)` function to add emails to columns O-Q
  - _Requirements: 4.3, 4.4, 4.5, 4.6_

- [x] 9.3 Create skipper identification functions
  - Write `identifySkippers(flexSheet)` function to find rows with #N/A in column L
  - Write `copySkippersToMailOut(flexSheet, skipperRows)` function to copy skipper data to Mail Out sheet
  - Clear existing Mail Out sheet data before adding new skippers
  - _Requirements: 5.1, 5.2, 5.3, 5.4_

- [x] 9.4 Create main enrichment orchestration function
  - Write `enrichFlexAbsences()` function to orchestrate the entire enrichment workflow
  - Preserve FlexiSched data in columns A-K during enrichment
  - Display success message with count of skippers identified
  - _Requirements: 3.4, 4.3, 5.5_

- [x] 10. Complete Comment Sync Service Module
  - Complete the `updateFlexComments(flexSheet, commentMap)` function (currently incomplete - missing closing braces and logic)
  - Write `syncComments()` function to orchestrate comment sync from Mail Out to Flex Absences
  - Display success message with count of comments synced
  - _Requirements: 6.1, 6.2, 6.3, 6.4, 6.5_




