# Requirements Document

## Introduction

This document specifies the requirements for a Google Apps Script (GAS) automation system that identifies students who skipped their flex classes by consolidating COGNOS attendance data with FlexiSched absence reports. The system automates the retrieval of three Excel reports emailed daily at 11:30 AM from COGNOS, imports their data into separate sheets, enriches manually-entered FlexiSched absence data with attendance codes and contact information, identifies students without legitimate absence codes (skippers), and prepares data for parent notification via FormMule.

## Glossary

- **GAS**: Google Apps Script - the scripting platform for Google Workspace automation
- **COGNOS**: The business intelligence system that generates and emails three daily reports
- **FlexiSched**: The external program that tracks flex class attendance; user manually copies absence data from this system
- **FormMule**: The email distribution system used to communicate with parents about student absences
- **Flex Absences Sheet**: The sheet with a date-based name (e.g., "11.3 flex absences") where FlexiSched absence data is manually entered
- **BHS Attendance Sheet**: The sheet containing COGNOS attendance data with legitimate absence codes in column K
- **2nd Period Default Sheet**: The sheet containing student 2nd period teacher assignments in column G
- **Contact Info Sheet**: The sheet containing student and guardian email addresses
- **Mail Out Sheet**: The sheet containing only students who skipped (have #N/A in attendance code column) for FormMule processing
- **Report Email**: An email sent by COGNOS containing an attached Excel spreadsheet
- **Attendance Code**: A code in column K of BHS Attendance Sheet indicating a legitimate absence reason
- **Skipper**: A student who appears in FlexiSched absences but has no matching attendance code in COGNOS data (indicated by #N/A)

## Requirements

### Requirement 1

**User Story:** As a daily report administrator, I want to manually import three specific COGNOS reports from my Gmail through a menu action, so that I have control over when the data is imported.

#### Acceptance Criteria

1. WHEN the user selects the import menu item, THE GAS SHALL search the user's Gmail for emails with subjects "My Student CY List - Courses, Teacher & Room", "My Student CY List - Student Email/Contact Info - Next Year Option (1)", and "My ATT - Attendance Bulletin (1)".
2. THE GAS SHALL identify and retrieve exactly three Report Emails with Excel spreadsheet attachments from the user's own email address.
3. WHEN a Report Email is found, THE GAS SHALL extract the Excel attachment from that email.
4. IF fewer than three Report Emails are found, THEN THE GAS SHALL log an error message indicating which reports are missing, display a notification to the user, and add the error message as a note in cell A1 of each affected sheet.
5. THE GAS SHALL process only the most recent Report Emails matching the search criteria.
6. WHEN the import completes successfully, THE GAS SHALL display a success notification to the user with the count of reports imported.

### Requirement 2

**User Story:** As a daily report administrator, I want each COGNOS report's data automatically imported into its designated sheet, so that the data is ready for enrichment and consolidation.

#### Acceptance Criteria

1. THE GAS SHALL import "My ATT - Attendance Bulletin (1)" data into the "BHS attendance" sheet filtering only rows where column J contains "02" (2nd period stored as text), with attendance codes stored in column K.
2. THE GAS SHALL import "My Student CY List - Courses, Teacher & Room" data into the "2nd period default" sheet with columns reordered as Student Id, Student Name, Grade, Period, Description, Room, Instructor, Instructor Id, Instructor Email, excluding the "9th Grd Entry" column.
3. THE GAS SHALL import "My Student CY List - Student Email/Contact Info - Next Year Option (1)" data into the "contact info" sheet with student email in column M, Guardian 1 email in column F, and Guardian 2 email in column J.
4. WHEN Excel data is extracted from a Report Email, THE GAS SHALL convert the Excel data to a format compatible with Google Sheets.
5. THE GAS SHALL clear existing data in each sheet before importing new data.
6. THE GAS SHALL preserve the column headers from the original Excel reports in each sheet.
7. WHEN the automated import is complete, THE GAS SHALL log a confirmation message indicating successful import of all three reports and add a note in cell A1 of each sheet with the date and time of successful import.

### Requirement 3

**User Story:** As a daily report administrator, I want to manually paste FlexiSched absence data into a date-named sheet, so that I can identify which students were absent from their flex classes.

#### Acceptance Criteria

1. THE GAS SHALL provide a sheet with a name containing the current date and ending with "flex absences" (e.g., "11.3 flex absences") where the user can paste FlexiSched data.
2. THE GAS SHALL preserve manually-entered FlexiSched data in columns A through K when importing COGNOS reports.
3. THE Flex Absences Sheet SHALL accept data in columns A through K from the FlexiSched absence report.
4. THE GAS SHALL not overwrite or modify manually-entered FlexiSched data in columns A through K during the enrichment process.

### Requirement 4

**User Story:** As a daily report administrator, I want to trigger data enrichment through a custom menu, so that I can add attendance codes, teacher names, and contact information to the FlexiSched absence data.

#### Acceptance Criteria

1. THE GAS SHALL create a custom menu in the Google Sheets user interface.
2. THE custom menu SHALL include menu items for importing COGNOS reports and enriching FlexiSched data.
3. WHEN the user selects the enrichment menu item, THE GAS SHALL lookup each student from the Flex Absences Sheet in the "BHS attendance" sheet and add the attendance code from column K into column L of the Flex Absences Sheet.
4. IF no matching student is found in "BHS attendance", THEN THE GAS SHALL add "#N/A" in column L of the Flex Absences Sheet.
5. WHEN enriching data, THE GAS SHALL lookup each student in the "2nd period default" sheet and add the teacher name from column G into column M of the Flex Absences Sheet.
6. WHEN enriching data, THE GAS SHALL lookup each student in the "contact info" sheet and add student email from column M into column O, Guardian 1 email from column F into column P, and Guardian 2 email from column J into column Q of the Flex Absences Sheet.

### Requirement 5

**User Story:** As a daily report administrator, I want students who skipped their flex class automatically identified and copied to the Mail Out sheet, so that I can send notifications to their parents via FormMule.

#### Acceptance Criteria

1. WHEN the enrichment process completes, THE GAS SHALL identify all rows in the Flex Absences Sheet where column L contains "#N/A".
2. WHEN a student with "#N/A" in column L is identified, THE GAS SHALL copy all data from columns A through Q of that row to the "Mail Out" sheet.
3. THE GAS SHALL clear existing data in the "Mail Out" sheet before adding newly identified skippers.
4. THE "Mail Out" sheet SHALL contain only students who skipped their flex class (those without legitimate attendance codes).
5. THE GAS SHALL preserve column N in the "Mail Out" sheet for user comments after FormMule processing.

### Requirement 6

**User Story:** As a daily report administrator, I want to copy comments from the Mail Out sheet back to the Flex Absences sheet after adding them, so that I maintain a complete record of all communications.

#### Acceptance Criteria

1. THE custom menu SHALL include a menu item for syncing comments from "Mail Out" to the Flex Absences Sheet.
2. WHEN the user selects the comment sync menu item, THE GAS SHALL match each student in the "Mail Out" sheet to the corresponding row in the Flex Absences Sheet.
3. WHEN a matching student is found, THE GAS SHALL copy the comment from column N of "Mail Out" to column N of the Flex Absences Sheet.
4. THE GAS SHALL preserve existing comments in the Flex Absences Sheet for students not in the "Mail Out" sheet.
5. IF no matching student is found, THEN THE GAS SHALL log a warning but continue processing other students.

### Requirement 7

**User Story:** As a daily report administrator, I want to automatically create a new flex absences sheet for today's date, so that I have a properly formatted sheet ready for FlexiSched data entry.

#### Acceptance Criteria

1. THE custom menu SHALL include a menu item for creating today's flex absences sheet.
2. WHEN the user selects the create sheet menu item, THE GAS SHALL format today's date as "M.D" (e.g., "11.3" for November 3rd).
3. THE GAS SHALL create a new sheet with the name "[Date] flex absences" where [Date] is formatted as "M.D".
4. IF a sheet with that name already exists, THEN THE GAS SHALL display a message indicating the sheet already exists and not create a duplicate.
5. WHEN creating a new flex absences sheet, THE GAS SHALL set up column headers for columns A through Q.
6. THE GAS SHALL set column L header to "Attendance Code", column M to "2nd Period Teacher", column N to "Comments", column O to "Student Email", column P to "Guardian 1 Email", and column Q to "Guardian 2 Email".

### Requirement 8

**User Story:** As a daily report administrator, I want the system to automatically detect my email address and search only my Gmail, so that I don't need to manually configure email settings.

#### Acceptance Criteria

1. WHEN the import operation begins, THE GAS SHALL automatically detect the active user's email address using `Session.getActiveUser().getEmail()`.
2. THE GAS SHALL search Gmail using a `from:` filter with the detected user email address.
3. THE GAS SHALL only retrieve COGNOS reports sent from the user's own email address.
4. THE GAS SHALL not require any manual configuration of email addresses or user settings.
5. THE system SHALL work automatically for any user who has access to the spreadsheet and runs the import operation.

### Requirement 9

**User Story:** As a daily report administrator, I want clear error messages and status updates, so that I can troubleshoot issues and confirm successful operations.

#### Acceptance Criteria

1. WHEN a manual operation begins, THE GAS SHALL display a status message indicating the operation in progress.
2. WHEN a manual operation completes successfully, THE GAS SHALL display a success message with the count of records processed.
3. IF an error occurs during automated import, THEN THE GAS SHALL log the error to the Apps Script execution log and display a notification to the user with the error status.
4. IF an error occurs during manual enrichment, THEN THE GAS SHALL log the error to the Apps Script execution log and display an error message identifying which sheet or operation caused the error.
5. THE GAS SHALL log all operations and errors to the Apps Script execution log for debugging purposes.
