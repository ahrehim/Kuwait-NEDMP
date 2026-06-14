/**
 * EPA NEDMP Verification Tool - Google Sheets backend
 *
 * Deploy this file as a Google Apps Script Web App.
 * First run setupDatabase() from Apps Script to create/initialize the Google Sheet.
 * Then deploy as:
 *   Execute as: Me
 *   Who has access: Anyone with the link
 *
 * https://script.google.com/macros/s/AKfycbztByF6hrhvkyJ3oFfSR7xCxHw2ZO70hV8hCu6aWYUFYQX6bcHDMK9xgVlKzVXj9RLFZQ/exec
 */

const APP_VERSION = 'v21';
const DATABASE_NAME = 'EPA_NEDMP_Verification_Database';
const RECORDS_SHEET_NAME = 'Verification Records';
const PASSWORDS_SHEET_NAME = 'Department Passwords';
const README_SHEET_NAME = 'README';
const SUMMARY_SHEET_NAME = 'Department Summary';

const RECORD_HEADERS = [
  'Submission ID',
  'Received At',
  'Backend Version',
  'Client Version',
  'Source URL',
  'Submission Action',
  'Timestamp',
  'Department',
  'Department Arabic',
  'Department Canonical ID',
  'Department Canonical Name',
  'Department Canonical Name Arabic',
  'Verifier Name',
  'Reviewed Department IDs',
  'Reviewed Departments',
  'Reviewed Departments Arabic',
  'Section ID',
  'Section Name',
  'Section Name Arabic',
  'Question ID',
  'Question No',
  'Question Text',
  'Question Text Arabic',
  'Decision',
  'Original Proposed EPA Integrated Response',
  'Original Proposed EPA Integrated Response Arabic',
  'Department Preferred Integrated Result',
  'Department Preferred Integrated Result Arabic',
  'Agreement Level',
  'Agreement Level Arabic',
  'Aggregation Method',
  'Aggregation Method Arabic',
  'Explanation',
  'Language'
];

const PASSWORD_HEADERS = ['department', 'password', 'departmentId', 'canonicalId', 'role'];
const DEPARTMENT_PASSWORDS = [
  {
    "department": "Chemical Safety Dep",
    "password": "ChemicalSafety2026",
    "departmentId": "D01",
    "canonicalId": "CD01",
    "role": "department"
  },
  {
    "department": "Industrial and Commercial Waste Section",
    "password": "IndustrialWaste2026",
    "departmentId": "D02",
    "canonicalId": "CD02",
    "role": "department"
  },
  {
    "department": "Municipal Waste Section",
    "password": "MunicipalWaste2026",
    "departmentId": "D03",
    "canonicalId": "CD03",
    "role": "department"
  },
  {
    "department": "Medical Waste Section",
    "password": "MedicalWaste2026",
    "departmentId": "D04",
    "canonicalId": "CD04",
    "role": "department"
  },
  {
    "department": "Air Quality Monitoring Department",
    "password": "AirQuality2026",
    "departmentId": "D05",
    "canonicalId": "CD05",
    "role": "department"
  },
  {
    "department": "Waste Management Department",
    "password": "WasteManagement2026",
    "departmentId": "D06",
    "canonicalId": "CD06",
    "role": "department"
  },
  {
    "department": "Water Quality Monitoring Department – Sea Water Monitoring Section",
    "password": "SeaWaterQuality2026",
    "departmentId": "D07",
    "canonicalId": "CD07",
    "role": "department"
  },
  {
    "department": "Center of Analytical Laboratories",
    "password": "AnalyticalLabs2026",
    "departmentId": "D08",
    "canonicalId": "CD08",
    "role": "department"
  },
  {
    "department": "Environmental Monitoring Information System (eMISK) – Database Section",
    "password": "eMISKDatabase2026",
    "departmentId": "D10",
    "canonicalId": "CD09",
    "role": "department"
  },
  {
    "department": "Admin",
    "password": "admin2026",
    "departmentId": "",
    "canonicalId": "",
    "role": "admin"
  }
];

function setupDatabase() {
  const ss = getOrCreateSpreadsheet_();
  ensureDatabase_(ss);
  Logger.log('EPA NEDMP verification database ready: ' + ss.getUrl());
  return ss.getUrl();
}

function doPost(e) {
  try {
    const payload = parsePayload_(e);
    const action = String(payload.action || 'submitVerification');
    if (action !== 'submitVerification' && action !== 'syncRecords' && action !== 'finalSubmission') {
      throw new Error('Unsupported action: ' + action);
    }
    const ss = getOrCreateSpreadsheet_();
    ensureDatabase_(ss);
    const result = appendRecords_(ss, payload, action);
    return json_({
      ok: true,
      action: action,
      inserted: result.inserted,
      skippedDuplicates: result.skippedDuplicates,
      spreadsheetUrl: ss.getUrl()
    });
  } catch (error) {
    return json_({ok: false, error: String(error && error.message ? error.message : error)});
  }
}

function doGet(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || 'status');
    const ss = getOrCreateSpreadsheet_();
    ensureDatabase_(ss);
    if (action === 'setup') {
      return json_({ok: true, spreadsheetUrl: ss.getUrl(), message: 'Database initialized.'});
    }
    return HtmlService.createHtmlOutput(
      '<h2>EPA NEDMP Verification Backend</h2>' +
      '<p>Status: online</p>' +
      '<p>Database: <a target="_blank" href="' + ss.getUrl() + '">' + ss.getUrl() + '</a></p>'
    );
  } catch (error) {
    return json_({ok: false, error: String(error && error.message ? error.message : error)});
  }
}

function parsePayload_(e) {
  const raw = (e && e.postData && e.postData.contents) ||
    (e && e.parameter && e.parameter.payload) ||
    '{}';
  return JSON.parse(raw || '{}');
}

function appendRecords_(ss, payload, action) {
  const records = Array.isArray(payload.records) ? payload.records : (payload.record ? [payload.record] : []);
  if (!records.length) return {inserted: 0, skippedDuplicates: 0};

  const sheet = ensureSheet_(ss, RECORDS_SHEET_NAME, RECORD_HEADERS);
  const existingIds = existingSubmissionIds_(sheet);
  const rows = [];
  let skippedDuplicates = 0;

  records.forEach(function(record) {
    const submissionId = safe_(record.submissionId) || Utilities.getUuid();
    if (existingIds.has(submissionId)) {
      skippedDuplicates += 1;
      return;
    }
    existingIds.add(submissionId);
    rows.push(recordToRow_(record, payload, action, submissionId));
  });

  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, RECORD_HEADERS.length).setValues(rows);
    sheet.autoResizeColumns(1, RECORD_HEADERS.length);
    updateSummary_(ss);
  }
  return {inserted: rows.length, skippedDuplicates: skippedDuplicates};
}

function recordToRow_(record, payload, action, submissionId) {
  return [
    submissionId,
    new Date(),
    APP_VERSION,
    safe_(payload.appVersion),
    safe_(payload.sourceUrl),
    action,
    safe_(record.timestamp),
    safe_(record.verifierDepartment),
    safe_(record.verifierDepartmentAr),
    safe_(record.verifierDepartmentCanonicalId),
    safe_(record.verifierDepartmentCanonicalName),
    safe_(record.verifierDepartmentCanonicalNameAr),
    safe_(record.verifierName),
    safe_(record.reviewedDepartmentIds),
    safe_(record.reviewedDepartments),
    safe_(record.reviewedDepartmentsAr),
    safe_(record.sectionId),
    safe_(record.sectionName),
    safe_(record.sectionNameAr),
    safe_(record.questionId),
    safe_(record.questionNo),
    safe_(record.question),
    safe_(record.questionAr),
    safe_(record.decision),
    safe_(record.originalIntegratedResponse),
    safe_(record.originalIntegratedResponseAr),
    safe_(record.preferredIntegratedResult),
    safe_(record.preferredIntegratedResultAr),
    safe_(record.agreementLevel),
    safe_(record.agreementLevelAr),
    safe_(record.aggregationMethod),
    safe_(record.aggregationMethodAr),
    safe_(record.explanation),
    safe_(payload.language)
  ];
}

function existingSubmissionIds_(sheet) {
  const ids = new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return ids;
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  values.forEach(function(row) {
    if (row[0]) ids.add(String(row[0]));
  });
  return ids;
}

function getOrCreateSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const existingId = props.getProperty('SPREADSHEET_ID');
  if (existingId) {
    try {
      return SpreadsheetApp.openById(existingId);
    } catch (error) {
      props.deleteProperty('SPREADSHEET_ID');
    }
  }

  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    props.setProperty('SPREADSHEET_ID', active.getId());
    return active;
  }

  const ss = SpreadsheetApp.create(DATABASE_NAME);
  props.setProperty('SPREADSHEET_ID', ss.getId());
  return ss;
}

function ensureDatabase_(ss) {
  const recordsSheet = ensureSheet_(ss, RECORDS_SHEET_NAME, RECORD_HEADERS);
  recordsSheet.setFrozenRows(1);
  recordsSheet.getRange(1, 1, 1, RECORD_HEADERS.length).setFontWeight('bold').setBackground('#dff0f7');

  const passwordSheet = ensureSheet_(ss, PASSWORDS_SHEET_NAME, PASSWORD_HEADERS);
  passwordSheet.clearContents();
  const passwordRows = [PASSWORD_HEADERS].concat(DEPARTMENT_PASSWORDS.map(function(entry) {
    return PASSWORD_HEADERS.map(function(header) { return safe_(entry[header]); });
  }));
  passwordSheet.getRange(1, 1, passwordRows.length, PASSWORD_HEADERS.length).setValues(passwordRows);
  passwordSheet.setFrozenRows(1);
  passwordSheet.getRange(1, 1, 1, PASSWORD_HEADERS.length).setFontWeight('bold').setBackground('#dff0f7');
  passwordSheet.autoResizeColumns(1, PASSWORD_HEADERS.length);

  const readme = ensureSheet_(ss, README_SHEET_NAME, ['Item', 'Value']);
  readme.clearContents();
  readme.getRange(1, 1, 6, 2).setValues([
    ['Item', 'Value'],
    ['Purpose', 'Unified online database for EPA NEDMP departmental verification submissions'],
    ['Records sheet', RECORDS_SHEET_NAME],
    ['Backend version', APP_VERSION],
    ['Created/updated', new Date()],
    ['Web app status', 'Deploy this Apps Script as a Web App and paste the deployment URL into NEDMP-Verification-Config.js']
  ]);
  readme.setFrozenRows(1);
  readme.autoResizeColumns(1, 2);

  updateSummary_(ss);
}

function ensureSheet_(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    const needsHeader = headers.some(function(header, index) { return current[index] !== header; });
    if (needsHeader) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

function updateSummary_(ss) {
  const summary = ensureSheet_(ss, SUMMARY_SHEET_NAME, ['Metric', 'Value']);
  summary.clearContents();
  summary.getRange(1, 1, 8, 2).setValues([
    ['Metric', 'Value'],
    ['Total verification records', '=COUNTA(\'' + RECORDS_SHEET_NAME + '\'!A2:A)'],
    ['Unique departments', '=COUNTUNIQUE(\'' + RECORDS_SHEET_NAME + '\'!H2:H)'],
    ['Unique questions verified', '=COUNTUNIQUE(\'' + RECORDS_SHEET_NAME + '\'!T2:T)'],
    ['Agree records', '=COUNTIF(\'' + RECORDS_SHEET_NAME + '\'!X2:X,"Agree")'],
    ['Disagree records', '=COUNTIF(\'' + RECORDS_SHEET_NAME + '\'!X2:X,"Disagree")'],
    ['Last received timestamp', '=MAX(\'' + RECORDS_SHEET_NAME + '\'!B2:B)'],
    ['Database URL', ss.getUrl()]
  ]);
  summary.setFrozenRows(1);
  summary.autoResizeColumns(1, 2);
}

function safe_(value) {
  if (value === null || value === undefined) return '';
  return String(value);
}

function json_(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
