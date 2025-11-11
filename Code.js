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
