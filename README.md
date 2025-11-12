# FlexiSchedule Attendance Automation

A Google Apps Script automation system that identifies students who skipped their flex classes by consolidating COGNOS attendance data with FlexiSched absence reports.

## Overview

This system automates the daily workflow of:
1. Importing three COGNOS attendance reports from Gmail
2. Enriching manually-entered FlexiSched absence data with attendance codes and contact information
3. Identifying students without legitimate absence codes (skippers)
4. Preparing data for parent notification via FormMule

## Features

- **Automated Report Import**: Retrieves three daily COGNOS Excel reports from Gmail
- **Data Enrichment**: Automatically adds attendance codes, teacher names, and contact information to absence records
- **Skipper Identification**: Flags students who were absent without legitimate excuse codes
- **Comment Synchronization**: Syncs comments between Mail Out and Flex Absences sheets
- **Date-Based Sheet Creation**: Automatically creates properly formatted sheets for each day's absences

## Setup

### Prerequisites

- Google Workspace account with access to Gmail and Google Sheets
- COGNOS system configured to email three daily reports:
  - "My ATT - Attendance Bulletin (1)"
  - "My Student CY List - Courses, Teacher & Room"
  - "My Student CY List - Student Email/Contact Info - Next Year Option (1)"

### Installation

1. Create a new Google Sheet or open your existing FlexiSchedule spreadsheet
2. Open the Apps Script editor (Extensions > Apps Script)
3. Copy the contents of `Code.js` into the script editor
4. Save the project
5. Refresh your Google Sheet to see the custom "FlexiSchedule" menu

### Required Sheets

The system expects the following sheets in your spreadsheet:
- **[Date] flex absences** (e.g., "11.3 flex absences") - Created automatically or manually
- **BHS attendance** - Populated by COGNOS import
- **2nd period default** - Populated by COGNOS import
- **contact info** - Populated by COGNOS import
- **Mail Out** - Automatically populated with skippers

## Usage

### Daily Workflow

1. **Create Today's Sheet** (if needed)
   - Select `FlexiSchedule > Create Today's Flex Absences Sheet`
   - A new sheet will be created with today's date (e.g., "11.3 flex absences")

2. **Import COGNOS Reports**
   - Select `FlexiSchedule > Import COGNOS Reports`
   - The system will search your Gmail for the three daily reports
   - Data will be imported into BHS attendance, 2nd period default, and contact info sheets

3. **Enter FlexiSched Data**
   - Manually paste absence data from FlexiSched into columns A-K of today's flex absences sheet

4. **Enrich Data**
   - Select `FlexiSchedule > Enrich Flex Absences Data`
   - The system will add:
     - Attendance codes (column L)
     - 2nd period teacher names (column M)
     - Student and guardian email addresses (columns O-Q)
   - Students with "#N/A" in the attendance code column are automatically copied to the Mail Out sheet

5. **Process with FormMule**
   - Use the Mail Out sheet with FormMule to send parent notifications
   - Add any comments in column N of the Mail Out sheet

6. **Sync Comments**
   - Select `FlexiSchedule > Sync Comments from Mail Out`
   - Comments from the Mail Out sheet will be copied back to the flex absences sheet

## Sheet Structure

For detailed column-by-column documentation of all sheets, see [SHEETS_STRUCTURE.md](SHEETS_STRUCTURE.md).

### Quick Overview

**Flex Absences Sheet**: Contains columns A-Q with FlexiSched data (A-K manually entered) and enriched data (L-Q auto-populated with attendance codes, teacher names, and contact information).

**Mail Out Sheet**: Contains only students who skipped (those with "#N/A" in attendance code column), ready for FormMule processing.

## Permissions

The script requires the following OAuth scopes:
- `spreadsheets.currentonly` - Access to the current spreadsheet
- `gmail.readonly` - Read access to Gmail for report retrieval
- `drive` - Access to Drive for Excel file processing

## Troubleshooting

### Reports Not Found
- Verify the COGNOS reports were sent to your email address
- Check that the email subjects match exactly:
  - "My Student CY List - Courses, Teacher & Room"
  - "My Student CY List - Student Email/Contact Info - Next Year Option (1)"
  - "My ATT - Attendance Bulletin (1)"
- Ensure the reports are from your own email address

### Missing Data After Enrichment
- Verify all required sheets exist (BHS attendance, 2nd period default, contact info)
- Check that COGNOS reports were imported successfully
- Review the Apps Script execution log for errors (View > Logs in the script editor)

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
