# FlexiSched Directions

## Overview

The FlexiSched Attendance Automation system helps you identify students who skipped their flex classes by consolidating COGNOS attendance data with FlexiSched absence reports. This guide walks you through the daily workflow.

---

## Daily Workflow

### Step 1: Create Today's Flex Absences Sheet

1. Open the **Flex Absence Tracker** menu in Google Sheets
2. Select **Create today's flex absences sheet**
3. A new sheet will be created with today's date (e.g., "11.3 flex absences")
4. If the sheet already exists, you'll see a message confirming it

**Note:** This step only needs to be done once per day. If you've already created today's sheet, skip to Step 2.

---

### Step 2: Import COGNOS Reports

1. Ensure you've received three COGNOS reports in your Gmail:
   - "My ATT - Attendance Bulletin (1)"
   - "My Student CY List - Courses, Teacher & Room"
   - "My Student CY List - Student Email/Contact Info - Next Year Option (1)"

2. Open the **Flex Absence Tracker** menu
3. Select **Import COGNOS Reports from GMail**
4. The system will automatically:
   - Search your Gmail for the three reports
   - Extract the Excel attachments
   - Import data into the following sheets:
     - **BHS attendance** (filtered to 2nd period only)
     - **2nd period default** (with reordered columns)
     - **contact info**

5. Check cell A1 in each sheet for import status notes

**Troubleshooting:**
- If reports aren't found, verify they were sent to your email address
- Check that the email subjects match exactly
- Ensure the reports are from your own email address

---

### Step 3: Enter FlexiSched Data

1. Open FlexiSched and generate your absence report
2. Copy the absence data from FlexiSched
3. Navigate to today's flex absences sheet (e.g., "11.3 flex absences")
4. Paste the data into **columns A-K** starting at row 2
   - Column A: Student ID
   - Column B: First Name
   - Column C: Last Name
   - Column D: Grade
   - Column E: Flex Name
   - Column F: Type
   - Column G: Request
   - Column H: Day
   - Column I: Period
   - Column J: Date
   - Column K: Flex Status

**Note:** Do not paste data into columns L-Q. These will be automatically populated in the next step.

---

### Step 4: Add Data to Flex Absences Sheet

1. Open the **Flex Absence Tracker** menu
2. Select **Add data to Flex Absences sheet**
3. The system will automatically populate:
   - **Column L:** Attendance Code (from BHS attendance sheet)
     - If a student isn't found, "#N/A" will appear
   - **Column M:** 2nd Period Teacher (from 2nd period default sheet)
   - **Column O:** Student Email (from contact info sheet)
   - **Column P:** Guardian 1 Email (from contact info sheet)
   - **Column Q:** Guardian 2 Email (from contact info sheet)

4. Students with "#N/A" in the Attendance Code column are automatically identified as **skippers**
5. All skippers are copied to the **Mail Out** sheet for parent notification

**What's a skipper?**
A skipper is a student who was absent from their flex class but doesn't have a legitimate attendance code in the system (meaning they were present at school but skipped their flex).

---

### Step 5: Process with FormMule

1. Navigate to the **Mail Out** sheet
2. Review the list of skippers
3. Add any necessary comments in **Column N** (Comments)
4. Use FormMule to send parent notifications based on the Mail Out sheet data
5. FormMule will add a send status column (e.g., "11.3 - Send Status") to track which emails were sent

**Note:** The Mail Out sheet contains only students who skipped (those with "#N/A" in the attendance code).

---

### Step 6: Sync Comments

After processing with FormMule, if you added comments in the Mail Out sheet:

1. Open the **Flex Absence Tracker** menu
2. Select **Sync Comments from Mail Out sheet**
3. Comments from Column N of the Mail Out sheet will be copied back to Column N of the flex absences sheet
4. This keeps your records synchronized

---

## Sheet Structure Reference

### Flex Absences Sheet (e.g., "11.3 flex absences")

| Columns | Description | Source |
|---------|-------------|--------|
| A-K | FlexiSched data | Manual entry (paste from FlexiSched) |
| L | Attendance Code | Auto-populated from BHS attendance |
| M | 2nd Period Teacher | Auto-populated from 2nd period default |
| N | Comments | Manual entry or synced from Mail Out |
| O | Student Email | Auto-populated from contact info |
| P | Guardian 1 Email | Auto-populated from contact info |
| Q | Guardian 2 Email | Auto-populated from contact info |

### Mail Out Sheet

Contains only students who skipped (those with "#N/A" in attendance code). This sheet is used for FormMule parent notifications.

---

## Tips and Best Practices

1. **Run imports early:** Import COGNOS reports as soon as you receive them each morning
2. **Check cell A1 notes:** After importing, check cell A1 in each sheet for import status and timestamps
3. **Review skippers:** Before sending notifications, review the Mail Out sheet to ensure all skippers are legitimate
4. **Add context in comments:** Use Column N to add any relevant context before sending parent notifications
5. **Keep sheets organized:** The system automatically creates dated sheets, so you have a historical record

---

## Common Issues

### "Could not find flex absences sheet"
- Make sure you've created today's sheet using the menu
- Check that the sheet name follows the format "M.D flex absences" (e.g., "11.3 flex absences")

### "No Excel attachments found"
- Verify the COGNOS reports have Excel (.xlsx) attachments
- Check that you're looking at the correct emails

### "No data found in flex absences sheet"
- Make sure you've pasted FlexiSched data into columns A-K before running enrichment

### Missing attendance codes or teacher names
- Verify COGNOS reports were imported successfully
- Check that the BHS attendance and 2nd period default sheets have data
- Ensure student IDs match between sheets

---

## Need More Help?

For detailed technical documentation, reference this project on GitHub at: https://github.com/AlvaroGomezMartinez/flexischedule-project

For technical issues or questions about the system, contact your Academic Technology Coach.
