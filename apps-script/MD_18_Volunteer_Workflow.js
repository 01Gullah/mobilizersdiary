/*
 * THE MOBILIZER'S DIARY — VOLUNTEER INTEREST WORKFLOW
 *
 * Uses the existing CMS_ID Script Property and writes only to the separate
 * Volunteer_Applications tab. It does not write to student reports,
 * newsletter subscribers, or opportunity recommendations.
 *
 * Run MD18_setupVolunteerWorkflow() once. It creates the tab automatically
 * and safely upgrades the previous version by adding State_of_Study and
 * Institution_Type columns after Institution, preserving existing rows.
 *
 * API route: volunteerApplication -> MD18_submitVolunteerApplication_(data)
 * This module does not send email; it returns a VOL reference on screen.
 */

var MD18_VOL_CONFIG_ = {
  SHEET_NAME: 'Volunteer_Applications',
  HEADERS: [
    'Application_ID', 'Received_At', 'Full_Name', 'Email', 'Phone',
    'Institution', 'State_of_Study', 'Institution_Type', 'Study_Level',
    'Interest_Area', 'Skills_Experience', 'Availability', 'Motivation',
    'Consent', 'Application_Status', 'Admin_Notes'
  ],
  STATE_OPTIONS: [
    'Abia', 'Adamawa', 'Akwa Ibom', 'Anambra', 'Bauchi', 'Bayelsa', 'Benue',
    'Borno', 'Cross River', 'Delta', 'Ebonyi', 'Edo', 'Ekiti', 'Enugu',
    'Federal Capital Territory (FCT)', 'Gombe', 'Imo', 'Jigawa', 'Kaduna',
    'Kano', 'Katsina', 'Kebbi', 'Kogi', 'Kwara', 'Lagos', 'Nasarawa',
    'Niger', 'Ogun', 'Ondo', 'Osun', 'Oyo', 'Plateau', 'Rivers', 'Sokoto',
    'Taraba', 'Yobe', 'Zamfara'
  ],
  INSTITUTION_TYPES: [
    'Federal University', 'State University', 'Private University',
    'Federal Polytechnic', 'State Polytechnic', 'Private Polytechnic',
    'College of Education', 'College/School of Health or Health Technology',
    'Other Tertiary Institution'
  ],
  STATUS_VALUES: ['New', 'Under Review', 'Contacted', 'Accepted', 'Waitlisted', 'Declined'],
  BRAND_GREEN: '#123F3B',
  BRAND_GOLD: '#C59B35'
};

/** Run once to create/validate, upgrade, and format the private review tab. */
function MD18_setupVolunteerWorkflow() {
  var spreadsheet = MD18_getCmsSpreadsheet_();
  var sheet = MD18_ensureSheet_(spreadsheet);
  MD18_styleSheet_(sheet);
  MD18_addStatusDropdown_(sheet);
  var result = {
    ok: true,
    workbook: spreadsheet.getName(),
    volunteerSheet: sheet.getName(),
    stateOptions: MD18_VOL_CONFIG_.STATE_OPTIONS.length,
    institutionTypeOptions: MD18_VOL_CONFIG_.INSTITUTION_TYPES.length,
    applicationStatusOptions: MD18_VOL_CONFIG_.STATUS_VALUES
  };
  Logger.log(JSON.stringify(result));
  return result;
}

/** Web API handler. Stores a volunteer enquiry privately and returns its reference. */
function MD18_submitVolunteerApplication_(data) {
  data = data || {};
  if (MD18_clean_(data.hp, 200)) return { accepted: true, ref: '' };

  var application = {
    fullName: MD18_clean_(data.fullName, 120),
    email: MD18_clean_(data.email, 180).toLowerCase(),
    phone: MD18_clean_(data.phone, 40),
    institution: MD18_clean_(data.institution, 180),
    state: MD18_clean_(data.stateOfStudy, 80),
    institutionType: MD18_clean_(data.institutionType, 100),
    studyLevel: MD18_clean_(data.studyLevel, 100),
    interestArea: MD18_clean_(data.interestArea, 120),
    skills: MD18_clean_(data.skills, 1800),
    availability: MD18_clean_(data.availability, 100),
    motivation: MD18_clean_(data.motivation, 2500),
    consent: data.consent === true || String(data.consent).toLowerCase() === 'true'
  };

  if (!application.fullName || !application.email || !application.state ||
      !application.institutionType || !application.interestArea ||
      !application.availability || !application.motivation) {
    throw new Error('Please complete all required volunteer fields.');
  }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(application.email)) {
    throw new Error('Please enter a valid email address.');
  }
  if (MD18_VOL_CONFIG_.STATE_OPTIONS.indexOf(application.state) < 0) {
    throw new Error('Please choose a state or FCT from the list.');
  }
  if (MD18_VOL_CONFIG_.INSTITUTION_TYPES.indexOf(application.institutionType) < 0) {
    throw new Error('Please choose an institution type from the list.');
  }
  if (!application.consent) {
    throw new Error('Please agree to the volunteer enquiry notice before submitting.');
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var spreadsheet = MD18_getCmsSpreadsheet_();
    var sheet = MD18_ensureSheet_(spreadsheet);
    var idInfo = MD18_nextId_(sheet);
    var row = [
      idInfo.id,
      new Date(),
      application.fullName,
      application.email,
      application.phone,
      application.institution,
      application.state,
      application.institutionType,
      application.studyLevel,
      application.interestArea,
      application.skills,
      application.availability,
      application.motivation,
      'Yes',
      'New',
      ''
    ];
    sheet.appendRow(row);
    PropertiesService.getScriptProperties().setProperty(idInfo.counterKey, String(idInfo.sequence));
    MD18_styleDataRow_(sheet, sheet.getLastRow());
    Logger.log('Received volunteer interest ' + idInfo.id);
    return { accepted: true, ref: idInfo.id };
  } finally {
    lock.releaseLock();
  }
}

function MD18_getCmsSpreadsheet_() {
  var id = PropertiesService.getScriptProperties().getProperty('CMS_ID');
  if (!id) throw new Error('CMS_ID is missing. Use the existing CMS workbook; do not create a separate workbook.');
  return SpreadsheetApp.openById(id);
}

function MD18_ensureSheet_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(MD18_VOL_CONFIG_.SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(MD18_VOL_CONFIG_.SHEET_NAME);
    sheet.getRange(1, 1, 1, MD18_VOL_CONFIG_.HEADERS.length).setValues([MD18_VOL_CONFIG_.HEADERS]);
    return sheet;
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, MD18_VOL_CONFIG_.HEADERS.length).setValues([MD18_VOL_CONFIG_.HEADERS]);
    return sheet;
  }

  var currentWidth = MD18_VOL_CONFIG_.HEADERS.length;
  if (sheet.getLastColumn() >= currentWidth) {
    var current = sheet.getRange(1, 1, 1, currentWidth).getDisplayValues()[0]
      .map(function (value) { return String(value || '').trim(); });
    var currentMatches = MD18_VOL_CONFIG_.HEADERS.every(function (header, index) {
      return current[index] === header;
    });
    if (currentMatches) return sheet;
  }

  // Upgrade the previous 14-column layout without overwriting applicant data.
  var legacyHeaders = [
    'Application_ID', 'Received_At', 'Full_Name', 'Email', 'Phone',
    'Institution', 'Study_Level', 'Interest_Area', 'Skills_Experience',
    'Availability', 'Motivation', 'Consent', 'Application_Status', 'Admin_Notes'
  ];
  if (sheet.getLastColumn() >= legacyHeaders.length) {
    var legacy = sheet.getRange(1, 1, 1, legacyHeaders.length).getDisplayValues()[0]
      .map(function (value) { return String(value || '').trim(); });
    var legacyMatches = legacyHeaders.every(function (header, index) {
      return legacy[index] === header;
    });
    if (legacyMatches) {
      sheet.insertColumnsAfter(6, 2);
      sheet.getRange(1, 7, 1, 2).setValues([['State_of_Study', 'Institution_Type']]);
      return sheet;
    }
  }

  throw new Error('The ' + MD18_VOL_CONFIG_.SHEET_NAME + ' tab has unexpected headers. No existing data was changed. Expected the current volunteer headers or the previous 14-column layout.');
}

function MD18_addStatusDropdown_(sheet) {
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(MD18_VOL_CONFIG_.STATUS_VALUES, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 15, Math.max(sheet.getMaxRows() - 1, 1), 1).setDataValidation(rule);
}

function MD18_styleSheet_(sheet) {
  var width = MD18_VOL_CONFIG_.HEADERS.length;
  sheet.setFrozenRows(1);
  sheet.setTabColor(MD18_VOL_CONFIG_.BRAND_GOLD);
  sheet.getRange(1, 1, 1, width)
    .setBackground(MD18_VOL_CONFIG_.BRAND_GREEN)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setWrap(true);
  sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 1), width).setVerticalAlignment('top');
  sheet.autoResizeColumns(1, width);
}

function MD18_styleDataRow_(sheet, rowNumber) {
  sheet.getRange(rowNumber, 1, 1, MD18_VOL_CONFIG_.HEADERS.length)
    .setVerticalAlignment('top')
    .setWrap(true);
}

function MD18_nextId_(sheet) {
  var year = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy');
  var key = 'MD18_VOL_COUNTER_' + year;
  var stored = Number(PropertiesService.getScriptProperties().getProperty(key) || 0);
  var highest = 0;
  if (sheet.getLastRow() >= 2) {
    var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
    var pattern = new RegExp('^VOL-' + year + '-(\\d+)$');
    values.forEach(function (item) {
      var match = pattern.exec(String(item[0] || '').trim());
      if (match) highest = Math.max(highest, Number(match[1]));
    });
  }
  var sequence = Math.max(stored, highest) + 1;
  return {
    id: 'VOL-' + year + '-' + String(sequence).padStart(4, '0'),
    sequence: sequence,
    counterKey: key
  };
}

function MD18_clean_(value, maxLength) {
  return String(value == null ? '' : value)
    .replace(/[\x00-\x1F\x7F]/g, ' ')
    .trim()
    .substring(0, maxLength);
}
