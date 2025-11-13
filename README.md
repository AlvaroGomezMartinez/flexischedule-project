# FlexiSchedule Attendance Automation

A Google Apps Script automation system that identifies students who skipped their flex classes by consolidating COGNOS attendance data with FlexiSched absence reports.

## Overview

This system automates the daily workflow of:
1. Importing three COGNOS attendance reports from Gmail
2. Enriching manually-entered FlexiSched absence data with attendance codes and contact information
3. Identifying students without legitimate absence codes (skippers)
4. Preparing data for parent notification via FormMule

## Features

- **Automated Report Import**: Retrieves three daily COGNOS Excel reports from Gmail with automatic filtering and column reordering
- **Data Enrichment**: Automatically adds attendance codes, teacher names, and contact information to absence records
- **Skipper Identification**: Flags students who were absent without legitimate excuse codes (shows "#N/A" in attendance code column)
- **Comment Synchronization**: Syncs comments between Mail Out sheet (column L) and Flex Absences sheet (column L)
- **Date-Based Sheet Creation**: Automatically creates properly formatted sheets for each day's absences
- **Header Restoration**: Automatically restores enrichment headers if overwritten when pasting FlexiSched data

## Setup

### Prerequisites

- Google Workspace account with access to Gmail and Google Sheets
- COGNOS system configured to email three daily reports:
  - "A new version of My ATT - Attendance Bulletin is available"
  - "A new version of My Student CY List - Course, Teacher & Room is available"
  - "A new version of My Student CY List - Student Email/Contact Info - Next Year Option is available"

### Installation

1. Create a new Google Sheet or open your existing FlexiSchedule spreadsheet
2. Open the Apps Script editor (Extensions > Apps Script)
3. Copy the contents of `Code.js` into the script editor
4. Save the project
5. Refresh your Google Sheet to see the custom "Flex Absence Tracker" menu

### Required Sheets

The system expects the following sheets in your spreadsheet:
- **[Date] flex absences** (e.g., "11.3 flex absences") - Created automatically or manually
- **BHS attendance** - Populated by COGNOS import (filtered to 2nd period only)
- **2nd period default** - Populated by COGNOS import (columns reordered automatically)
- **contact info** - Populated by COGNOS import (Student ID is in column B, not A)
- **Mail Out** - Automatically populated with skippers

**Note**: The contact info sheet has a different structure than other sheets - Student ID is in column B (column A contains "Current Building").

## Usage

### Daily Workflow

1. **Create Today's Sheet** (if needed)
   - Select `Flex Absence Tracker > Create today's flex absences sheet`
   - A new sheet will be created with today's date (e.g., "11.3 flex absences")

2. **Import COGNOS Reports**
   - Select `Flex Absence Tracker > Import COGNOS Reports from GMail`
   - The system will search your Gmail for the three daily reports
   - Data will be imported into BHS attendance, 2nd period default, and contact info sheets

3. **Enter FlexiSched Data**
   - Manually paste absence data from FlexiSched into columns A-L of today's flex absences sheet
   - The FlexiSched report includes two header rows, so data will start on row 3
   - Column L contains the FlexiSched Comment field

4. **Enrich Data**
   - Select `Flex Absence Tracker > Add data to Flex Absences sheet`
   - The system will automatically restore enrichment headers if they were overwritten
   - The system will add:
     - Attendance codes (column M)
     - 2nd period teacher names (column N)
     - Student and guardian email addresses (columns O-Q)
   - Students with "#N/A" in the attendance code column are automatically copied to the Mail Out sheet

5. **Process with FormMule**
   - Use the Mail Out sheet with FormMule to send parent notifications
   - Add any comments in column L of the Mail Out sheet

6. **Sync Comments**
   - Select `Flex Absence Tracker > Sync Comments from Mail Out sheet`
   - Comments from column L of the Mail Out sheet will be copied back to column L of the flex absences sheet

## Sheet Structure

For detailed column-by-column documentation of all sheets, see [SHEETS_STRUCTURE.md](SHEETS_STRUCTURE.md).

### Quick Overview

**Flex Absences Sheet**: Contains columns A-Q with:
- **A-L**: FlexiSched data (manually pasted from FlexiSched report, includes 2 header rows)
  - Column L contains FlexiSched Comment field
- **M-Q**: Enriched data (auto-populated by the script)
  - M: Attendance Code (from BHS attendance sheet)
  - N: 2nd Period Teacher (from 2nd period default sheet)
  - O: Student Email (from contact info sheet column N)
  - P: Guardian 1 Email (from contact info sheet column G)
  - Q: Guardian 2 Email (from contact info sheet column K)

**Mail Out Sheet**: Contains only students who skipped (those with "#N/A" in attendance code column M), ready for FormMule processing. Has 17 columns (A-Q) matching the flex absences structure.

## Permissions

The script requires the following OAuth scopes:
- `spreadsheets` - Access to Google Sheets
- `gmail.readonly` - Read access to Gmail for report retrieval
- `drive` - Access to Drive for Excel file processing
- `script.container.ui` - Display custom menus and dialogs
- `userinfo.email` - Access to user's email address for report filtering

## Troubleshooting

### Reports Not Found
- Verify the COGNOS reports were sent to your email address
- Check that the email subjects match exactly:
  - "A new version of My ATT - Attendance Bulletin is available"
  - "A new version of My Student CY List - Course, Teacher & Room is available"
  - "A new version of My Student CY List - Student Email/Contact Info - Next Year Option is available"
- Ensure the reports are from your own email address

### Missing Data After Enrichment
- Verify all required sheets exist (BHS attendance, 2nd period default, contact info)
- Check that COGNOS reports were imported successfully
- Ensure student IDs in the contact info sheet are in column B (not column A)
- Review the Apps Script execution log for errors (Extensions > Apps Script > Executions)

### Sheet Already Exists Error
- A sheet with today's date already exists
- Use the existing sheet or rename/delete it before creating a new one

## Technical Details

- **Runtime**: Google Apps Script V8
- **Time Zone**: America/Chicago
- **Exception Logging**: Stackdriver

## Project Structure

```
.
├── Code.js              # Main script file
├── appsscript.json      # Apps Script manifest
├── .clasp.json          # Clasp configuration
├── .claspignore         # Clasp ignore rules
├── SHEETS_STRUCTURE.md  # Detailed sheet structure documentation
└── README.md            # This file
```

## Development

This project uses [clasp](https://github.com/google/clasp) for local development and deployment.

### Deploy with clasp

```bash
# Login to clasp
clasp login

# Push changes to Apps Script
clasp push

# Open the project in the Apps Script editor
clasp open
```

## License

This project is intended for internal use at NISD.
