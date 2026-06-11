// ─── CONFIGURATION ────────────────────────────────────────────────────────────
// Fill in these values after you create the Sheet and Forms.
// Everything else in this project reads from here.

var CONFIG = {

  // Google Sheet ID — from the URL of your Portal Tracker sheet
  // e.g. https://docs.google.com/spreadsheets/d/SHEET_ID/edit
  SHEET_ID: 'PASTE_SHEET_ID_HERE',

  // Google Form ID for homework submission (Form A)
  HOMEWORK_FORM_ID: 'PASTE_HOMEWORK_FORM_ID_HERE',

  // Google Form ID for parent approval (Form B)
  APPROVAL_FORM_ID: 'PASTE_APPROVAL_FORM_ID_HERE',

  // Drive folder ID where student submissions are stored
  SUBMISSIONS_FOLDER_ID: 'PASTE_SUBMISSIONS_FOLDER_ID_HERE',

  // Drive folder ID for answer keys (teacher-only)
  ANSWER_KEYS_FOLDER_ID: 'PASTE_ANSWER_KEYS_FOLDER_ID_HERE',

  // Base URL for GitHub Pages
  BASE_URL: 'https://misra-ravi.github.io/physics-foundation/',

  // Portal admin email (receives error alerts)
  ADMIN_EMAIL: 'misra.ravikant@gmail.com',

  // Display name shown in emails
  PORTAL_NAME: 'IB Physics Portal',
};

// Sheet tab names — do not change after setup
var SHEETS = {
  ROSTER:   'Roster',
  UNITS:    'Units',
  PROGRESS: 'Progress',
  EMAIL_LOG: 'EmailLog',
};
