# Google Sheets Structure Documentation

This document describes the structure of all sheets used in the FlexiSchedule Attendance Automation system.

## Flex Absences Sheet

Daily sheet created with format `[Date] flex absences` (e.g., "11.3 flex absences")

| Column | Header | Source | Description |
|--------|--------|--------|-------------|
| A | ID | Manual Entry | Student ID from FlexiSched |
| B | First Name | Manual Entry | Student first name |
| C | Last Name | Manual Entry | Student last name |
| D | Grade | Manual Entry | Student grade level |
| E | Flex Name | Manual Entry | Name of flex class |
| F | Type | Manual Entry | Absence type |
| G | Request | Manual Entry | Request details |
| H | Day | Manual Entry | Day of week |
| I | Period | Manual Entry | Period number |
| J | Date | Manual Entry | Date of absence |
| K | Flex Status | Manual Entry | Status from FlexiSched |
| L | Brennan Attendance | Auto-populated | Attendance code from BHS attendance sheet |
| M | 2nd Period Teacher | Auto-populated | Teacher name from 2nd period default sheet |
| N | Comment | Manual/Synced | Comments (synced from Mail Out sheet) |
| O | Student Email | Auto-populated | From contact info sheet |
| P | Guardian 1 Email | Auto-populated | From contact info sheet |
| Q | Guardian 2 Email | Auto-populated | From contact info sheet |

**Note**: Columns A-K are manually entered by copying data from FlexiSched reports. Columns L-Q are automatically populated by the script.

---

## BHS Attendance Sheet

Populated by COGNOS import from "My ATT - Attendance Bulletin (1)" report.

| Column | Header | Description |
|--------|--------|-------------|
| A | Stu Id | Student ID number |
| B | Student Name | Full student name |
| C | Gr | Grade level |
| D | Team | Student team assignment |
| E | Hrm | Homeroom |
| F | Rm | Room number |
| G | Course | Course name |
| H | Section | Section number |
| I | Description | Course description |
| J | Pd | Period |
| K | Cd | Attendance code |
| L | Source | Data source |
| M | Student Email | Student email address |
| N | Guardian Email | Guardian email address |
| O | Count distinct(Student Id) | Student count |

---

## 2nd Period Default Sheet

Populated by COGNOS import from "My Student CY List - Courses, Teacher & Room" report.

| Column | Header | Description |
|--------|--------|-------------|
| A | Student Id | Student ID number |
| B | Student Name | Full student name |
| C | Grade | Grade level |
| D | Period | Period number |
| E | Description | Course description |
| F | Room | Room number |
| G | Instructor | Teacher name |
| H | Instructor ID | Teacher ID |
| I | Instructor Email | Teacher email address |

---

## Contact Info Sheet

Populated by COGNOS import from "My Student CY List - Student Email/Contact Info - Next Year Option (1)" report.

| Column | Header | Description |
|--------|--------|-------------|
| A | Student ID | Student ID number |
| B | Student Name | Full student name |
| C | Grade Level | Student grade level |
| D | Notification | Notification preference |
| E | Guardian 1 | Guardian 1 name |
| F | Guardian 1 Email | Guardian 1 email address |
| G | Guardian 1 Cell | Guardian 1 cell phone |
| H | Guardian 1 Home | Guardian 1 home phone |
| I | Guardian 2 | Guardian 2 name |
| J | Guardian 2 Email | Guardian 2 email address |
| K | Guardian 2 Cell | Guardian 2 cell phone |
| L | Guardian 2 Home | Guardian 2 home phone |
| M | Student Email | Student email address |

---

## Mail Out Sheet

Automatically populated with students who have "#N/A" in the Brennan Attendance column (skippers). Used for FormMule parent notification processing.

| Column | Header | Source | Description |
|--------|--------|--------|-------------|
| A | ID | From Flex Absences | Student ID |
| B | First Name | From Flex Absences | Student first name |
| C | Last Name | From Flex Absences | Student last name |
| D | Grad Year | From Flex Absences | Graduation year (Grade) |
| E | Flex Name | From Flex Absences | Name of flex class |
| F | Type | From Flex Absences | Absence type |
| G | Request | From Flex Absences | Request details |
| H | Day | From Flex Absences | Day of week |
| I | Period | From Flex Absences | Period number |
| J | Date | From Flex Absences | Date of absence |
| K | Flex Status | From Flex Absences | Status from FlexiSched |
| L | Brennan Attendance | From Flex Absences | Shows "#N/A" for skippers |
| M | 2nd Period Teacher | From Flex Absences | Teacher name |
| N | Comment | Manual Entry | Comments for parent notification |
| O | Student Email | From Flex Absences | Student email address |
| P | Guardian 1 Email | From Flex Absences | Guardian 1 email address |
| Q | Guardian 2 Email | From Flex Absences | Guardian 2 email address |
| R | [Date] - Send Status | FormMule | Email send status (e.g., "11.3 - Send Status") |

**Note**: Column R header changes based on the date (e.g., "11.3 - Send Status" for November 3rd).

---

## Data Flow

1. **COGNOS Reports** → Import into BHS attendance, 2nd period default, and contact info sheets
2. **FlexiSched Report** → Manual entry into columns A-K of Flex Absences sheet
3. **Enrichment Script** → Populates columns L-Q of Flex Absences sheet using VLOOKUP-style matching
4. **Skipper Detection** → Students with "#N/A" in column L are copied to Mail Out sheet
5. **FormMule Processing** → Mail Out sheet used to send parent notifications
6. **Comment Sync** → Comments from Mail Out sheet (column N) synced back to Flex Absences sheet
