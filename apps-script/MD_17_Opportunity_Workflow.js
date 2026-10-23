/*
 * THE MOBILIZER'S DIARY — DEDICATED OPPORTUNITY WORKFLOW
 *
 * Uses the existing CMS_ID Script Property and existing Opportunities CMS tab.
 * It does not write recommendations into Student_Voices or Submission_Log.
 *
 * Run MD17_setupOpportunityWorkflow() once to create/format the review tab,
 * validate the Opportunities tab, and install the approval-sync trigger.
 *
 * API actions added in MD_13_WEB_API.gs:
 *   recommendOpportunity -> MD17_recommendOpportunity_(data)
 *   publicOpportunities -> MD17_getPublicOpportunities_()
 */

var MD17_OPP_CONFIG_ = {
  RECOMMENDATIONS_SHEET: 'Opportunity_Recommendations',
  OPPORTUNITIES_SHEET: 'Opportunities',
  RECOMMENDATION_HEADERS: [
    'Recommendation_ID', 'Received_At', 'Title', 'Category', 'Organisation',
    'Source_URL', 'Deadline', 'Eligibility', 'Additional_Details',
    'Submitter_Name', 'Submitter_Email', 'Consent', 'Verification_Status',
    'Editorial_Status', 'Editor_Notes', 'Published_Opp_ID', 'Published_At'
  ],
  OPPORTUNITY_HEADERS: [
    'Opp_ID', 'Title', 'Organisation', 'Category', 'Deadline', 'Eligibility',
    'Apply_URL', 'Verification_Status', 'Editorial_Status', 'Issue_Number'
  ],
  BRAND_GREEN: '#123F3B',
  BRAND_GOLD: '#C59B35'
};

/** Run once. Creates tabs/dropdowns with code and installs one edit trigger. */
function MD17_setupOpportunityWorkflow() {
  var ss = MD17_getCmsSpreadsheet_();
  var recommendations = MD17_ensureSheet_(ss, MD17_OPP_CONFIG_.RECOMMENDATIONS_SHEET, MD17_OPP_CONFIG_.RECOMMENDATION_HEADERS);
  var opportunities = MD17_ensureSheet_(ss, MD17_OPP_CONFIG_.OPPORTUNITIES_SHEET, MD17_OPP_CONFIG_.OPPORTUNITY_HEADERS);
  MD17_styleSheet_(recommendations);
  MD17_styleSheet_(opportunities);
  MD17_addReviewDropdowns_(recommendations);
  MD17_installEditTrigger_(ss);
  var result = {
    ok: true,
    workbook: ss.getName(),
    recommendationSheet: recommendations.getName(),
    opportunitiesSheet: opportunities.getName(),
    editTriggerInstalled: true
  };
  Logger.log(JSON.stringify(result));
  return result;
}

/** Public API handler. Validates and records an opportunity recommendation. */
function MD17_recommendOpportunity_(data) {
  data = data || {};
  if (MD17_clean_(data.hp, 200)) {
    return { accepted: true, ref: '' };
  }

  var title = MD17_clean_(data.title, 180);
  var category = MD17_clean_(data.type || data.category, 100);
  var organisation = MD17_clean_(data.organisation, 180);
  var sourceUrl = MD17_clean_(data.sourceUrl || data.url, 500);
  var deadline = MD17_clean_(data.deadline, 100);
  var eligibility = MD17_clean_(data.eligibility, 2500);
  var notes = MD17_clean_(data.notes, 2500);
  var contactName = MD17_clean_(data.contactName, 120);
  var contactEmail = MD17_clean_(data.contactEmail, 180).toLowerCase();
  var consent = data.consent === true || String(data.consent).toLowerCase() === 'true';

  if (!title || !category || !organisation || !sourceUrl || !eligibility) {
    throw new Error('Please complete all required opportunity details.');
  }
  if (!MD17_isHttpUrl_(sourceUrl)) {
    throw new Error('The official opportunity link must begin with https:// or http://.');
  }
  if (!consent) throw new Error('Please confirm the recommendation statement before submitting.');
  if ((contactName && !contactEmail) || (!contactName && contactEmail)) {
    throw new Error('Enter both your name and email for follow-up, or leave both blank.');
  }
  if (contactEmail && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(contactEmail)) {
    throw new Error('Please enter a valid email address or leave it blank.');
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = MD17_getCmsSpreadsheet_();
    var review = MD17_ensureSheet_(ss, MD17_OPP_CONFIG_.RECOMMENDATIONS_SHEET, MD17_OPP_CONFIG_.RECOMMENDATION_HEADERS);
    var headers = MD17_headerMap_(review, MD17_OPP_CONFIG_.RECOMMENDATION_HEADERS);
    var stamp = new Date();
    var idInfo = MD17_nextId_(review, headers.Recommendation_ID, 'ROP');
    var row = [];
    row[headers.Recommendation_ID] = idInfo.id;
    row[headers.Received_At] = stamp;
    row[headers.Title] = title;
    row[headers.Category] = category;
    row[headers.Organisation] = organisation;
    row[headers.Source_URL] = sourceUrl;
    row[headers.Deadline] = deadline;
    row[headers.Eligibility] = eligibility;
    row[headers.Additional_Details] = notes;
    row[headers.Submitter_Name] = contactName;
    row[headers.Submitter_Email] = contactEmail;
    row[headers.Consent] = 'Yes';
    row[headers.Verification_Status] = 'Pending';
    row[headers.Editorial_Status] = 'Pending';
    row[headers.Editor_Notes] = '';
    row[headers.Published_Opp_ID] = '';
    row[headers.Published_At] = '';
    for (var i = 0; i < MD17_OPP_CONFIG_.RECOMMENDATION_HEADERS.length; i++) {
      if (typeof row[i] === 'undefined') row[i] = '';
    }
    review.appendRow(row);
    PropertiesService.getScriptProperties().setProperty(idInfo.counterKey, String(idInfo.sequence));
    MD17_styleDataRow_(review, review.getLastRow(), MD17_OPP_CONFIG_.RECOMMENDATION_HEADERS.length);
    Logger.log('Received opportunity recommendation ' + idInfo.id);
    return { accepted: true, ref: idInfo.id };
  } finally {
    lock.releaseLock();
  }
}

/** Public API handler. Returns approved listings only; never returns submitter details. */
function MD17_getPublicOpportunities_() {
  var ss = MD17_getCmsSpreadsheet_();
  var sheet = ss.getSheetByName(MD17_OPP_CONFIG_.OPPORTUNITIES_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var headers = MD17_headerMap_(sheet, MD17_OPP_CONFIG_.OPPORTUNITY_HEADERS);
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues();
  var output = [];
  rows.forEach(function (row) {
    var verified = String(row[headers.Verification_Status] || '').trim().toLowerCase() === 'verified';
    var approved = String(row[headers.Editorial_Status] || '').trim().toLowerCase() === 'approved';
    var applyUrl = String(row[headers.Apply_URL] || '').trim();
    var title = String(row[headers.Title] || '').trim();
    var organisation = String(row[headers.Organisation] || '').trim();
    if (!verified || !approved || !title || !organisation || !MD17_isHttpUrl_(applyUrl)) return;
    output.push({
      id: String(row[headers.Opp_ID] || '').trim(),
      title: title,
      organisation: organisation,
      category: String(row[headers.Category] || '').trim(),
      deadline: String(row[headers.Deadline] || '').trim(),
      eligibility: String(row[headers.Eligibility] || '').trim(),
      applyUrl: applyUrl
    });
  });
  return output;
}

/** Installable edit trigger. Publishes or hides an opportunity when editors change review statuses. */
function MD17_onOpportunityRecommendationEdit(e) {
  if (!e || !e.range || !e.source) return;
  var editedSheet = e.range.getSheet();
  if (editedSheet.getName() !== MD17_OPP_CONFIG_.RECOMMENDATIONS_SHEET || e.range.getRow() < 2) return;
  var reviewHeaders = MD17_headerMap_(editedSheet, MD17_OPP_CONFIG_.RECOMMENDATION_HEADERS);
  var firstCol = e.range.getColumn();
  var lastCol = firstCol + e.range.getNumColumns() - 1;
  var verifyCol = reviewHeaders.Verification_Status + 1;
  var editorialCol = reviewHeaders.Editorial_Status + 1;
  if (!((firstCol <= verifyCol && lastCol >= verifyCol) || (firstCol <= editorialCol && lastCol >= editorialCol))) return;

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = e.source;
    var review = editedSheet;
    var opportunities = MD17_ensureSheet_(ss, MD17_OPP_CONFIG_.OPPORTUNITIES_SHEET, MD17_OPP_CONFIG_.OPPORTUNITY_HEADERS);
    var oppHeaders = MD17_headerMap_(opportunities, MD17_OPP_CONFIG_.OPPORTUNITY_HEADERS);
    var startRow = e.range.getRow();
    var endRow = startRow + e.range.getNumRows() - 1;
    for (var rowNumber = startRow; rowNumber <= endRow; rowNumber++) {
      MD17_syncReviewRow_(review, opportunities, reviewHeaders, oppHeaders, rowNumber);
    }
  } finally {
    lock.releaseLock();
  }
}

function MD17_syncReviewRow_(review, opportunities, reviewHeaders, oppHeaders, rowNumber) {
  if (rowNumber > review.getLastRow()) return;
  var reviewRow = review.getRange(rowNumber, 1, 1, MD17_OPP_CONFIG_.RECOMMENDATION_HEADERS.length).getValues()[0];
  var recommendationId = String(reviewRow[reviewHeaders.Recommendation_ID] || '').trim();
  if (!recommendationId) return;
  var verification = String(reviewRow[reviewHeaders.Verification_Status] || '').trim();
  var editorial = String(reviewRow[reviewHeaders.Editorial_Status] || '').trim();
  var verified = verification.toLowerCase() === 'verified';
  var approved = editorial.toLowerCase() === 'approved';
  var publishedId = String(reviewRow[reviewHeaders.Published_Opp_ID] || '').trim();
  var publicRowNumber = publishedId ? MD17_findIdRow_(opportunities, oppHeaders.Opp_ID, publishedId) : 0;

  if (!verified || !approved) {
    if (publicRowNumber) {
      opportunities.getRange(publicRowNumber, oppHeaders.Verification_Status + 1).setValue(verification || 'Pending');
      opportunities.getRange(publicRowNumber, oppHeaders.Editorial_Status + 1).setValue(editorial || 'Pending');
      Logger.log('Listing ' + publishedId + ' hidden from public feed; review status changed.');
    }
    return;
  }

  var year = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy');
  if (!publishedId) {
    var idInfo = MD17_nextId_(opportunities, oppHeaders.Opp_ID, 'OP');
    publishedId = idInfo.id;
    publicRowNumber = opportunities.getLastRow() + 1;
    PropertiesService.getScriptProperties().setProperty(idInfo.counterKey, String(idInfo.sequence));
  } else if (!publicRowNumber) {
    publicRowNumber = opportunities.getLastRow() + 1;
  }
  var publishedAt = reviewRow[reviewHeaders.Published_At] || new Date();
  var publicRow = [];
  publicRow[oppHeaders.Opp_ID] = publishedId;
  publicRow[oppHeaders.Title] = reviewRow[reviewHeaders.Title];
  publicRow[oppHeaders.Organisation] = reviewRow[reviewHeaders.Organisation];
  publicRow[oppHeaders.Category] = reviewRow[reviewHeaders.Category];
  publicRow[oppHeaders.Deadline] = reviewRow[reviewHeaders.Deadline];
  publicRow[oppHeaders.Eligibility] = reviewRow[reviewHeaders.Eligibility];
  publicRow[oppHeaders.Apply_URL] = reviewRow[reviewHeaders.Source_URL];
  publicRow[oppHeaders.Verification_Status] = 'Verified';
  publicRow[oppHeaders.Editorial_Status] = 'Approved';
  publicRow[oppHeaders.Issue_Number] = '';
  for (var i = 0; i < MD17_OPP_CONFIG_.OPPORTUNITY_HEADERS.length; i++) {
    if (typeof publicRow[i] === 'undefined') publicRow[i] = '';
  }
  opportunities.getRange(publicRowNumber, 1, 1, publicRow.length).setValues([publicRow]);
  review.getRange(rowNumber, reviewHeaders.Published_Opp_ID + 1).setValue(publishedId);
  review.getRange(rowNumber, reviewHeaders.Published_At + 1).setValue(publishedAt);
  MD17_styleDataRow_(opportunities, publicRowNumber, publicRow.length);
  Logger.log('Recommendation ' + recommendationId + ' published as ' + publishedId + ' (' + year + ').');
}

function MD17_installEditTrigger_(ss) {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'MD17_onOpportunityRecommendationEdit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger('MD17_onOpportunityRecommendationEdit').forSpreadsheet(ss).onEdit().create();
}

function MD17_nextId_(sheet, idColumnIndex, prefix) {
  var year = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy');
  var key = 'MD17_' + prefix + '_COUNTER_' + year;
  var stored = Number(PropertiesService.getScriptProperties().getProperty(key) || 0);
  var highest = 0;
  if (sheet.getLastRow() >= 2) {
    var ids = sheet.getRange(2, idColumnIndex + 1, sheet.getLastRow() - 1, 1).getDisplayValues();
    var pattern = new RegExp('^' + prefix + '-' + year + '-(\\d+)$');
    ids.forEach(function (item) {
      var match = pattern.exec(String(item[0] || '').trim());
      if (match) highest = Math.max(highest, Number(match[1]));
    });
  }
  var sequence = Math.max(stored, highest) + 1;
  return { id: prefix + '-' + year + '-' + String(sequence).padStart(4, '0'), sequence: sequence, counterKey: key };
}

function MD17_findIdRow_(sheet, idColumnIndex, id) {
  if (sheet.getLastRow() < 2) return 0;
  var values = sheet.getRange(2, idColumnIndex + 1, sheet.getLastRow() - 1, 1).getDisplayValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === id) return i + 2;
  }
  return 0;
}

function MD17_getCmsSpreadsheet_() {
  var id = PropertiesService.getScriptProperties().getProperty('CMS_ID');
  if (!id) throw new Error('CMS_ID is missing. Run the existing workspace setup; do not create a separate workbook.');
  return SpreadsheetApp.openById(id);
}

function MD17_ensureSheet_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    var actual = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0].map(function (v) { return String(v).trim(); });
    for (var i = 0; i < headers.length; i++) {
      if (actual[i] !== headers[i]) {
        throw new Error('The ' + name + ' tab has unexpected headers. No existing data was changed. Expected column ' + (i + 1) + ' to be ' + headers[i] + '.');
      }
    }
  }
  return sheet;
}

function MD17_headerMap_(sheet, expectedHeaders) {
  var row = sheet.getRange(1, 1, 1, expectedHeaders.length).getDisplayValues()[0].map(function (v) { return String(v).trim(); });
  var map = {};
  expectedHeaders.forEach(function (header) {
    var index = row.indexOf(header);
    if (index < 0) throw new Error('Missing expected column ' + header + ' in tab ' + sheet.getName() + '.');
    map[header] = index;
  });
  return map;
}

function MD17_addReviewDropdowns_(sheet) {
  var rows = Math.max(sheet.getMaxRows() - 1, 1);
  var verificationRule = SpreadsheetApp.newDataValidation().requireValueInList(['Pending', 'Verified', 'Not Verified'], true).setAllowInvalid(false).build();
  var editorialRule = SpreadsheetApp.newDataValidation().requireValueInList(['Pending', 'Under Review', 'Approved', 'Rejected', 'Needs Info'], true).setAllowInvalid(false).build();
  sheet.getRange(2, 13, rows, 1).setDataValidation(verificationRule);
  sheet.getRange(2, 14, rows, 1).setDataValidation(editorialRule);
}

function MD17_styleSheet_(sheet) {
  var width = sheet.getLastColumn();
  if (width < 1) return;
  sheet.setFrozenRows(1);
  sheet.setTabColor(MD17_OPP_CONFIG_.BRAND_GOLD);
  sheet.getRange(1, 1, 1, width).setBackground(MD17_OPP_CONFIG_.BRAND_GREEN).setFontColor('#FFFFFF').setFontWeight('bold').setWrap(true);
  sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 1), width).setVerticalAlignment('top');
  sheet.autoResizeColumns(1, width);
}

function MD17_styleDataRow_(sheet, rowNumber, width) {
  sheet.getRange(rowNumber, 1, 1, width).setVerticalAlignment('top').setWrap(true);
}

function MD17_clean_(value, maxLength) {
  return String(value == null ? '' : value).replace(/[\u0000-\u001F\u007F]/g, ' ').trim().substring(0, maxLength);
}

function MD17_isHttpUrl_(value) {
  return /^https?:\/\/[^\s]+$/i.test(String(value || '').trim());
}


/**
 * Publish one recommendation from the authenticated Admin Portal.
 * The governance layer must authorize this action before calling it.
 */
function MD17_publishRecommendationFromPortal_(recommendationId, actorEmail) {
  recommendationId = MD17_clean_(recommendationId, 120);
  actorEmail = MD17_clean_(actorEmail, 180).toLowerCase();
  if (!recommendationId) throw new Error('Recommendation reference is required.');

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = MD17_getCmsSpreadsheet_();
    var review = MD17_ensureSheet_(ss, MD17_OPP_CONFIG_.RECOMMENDATIONS_SHEET, MD17_OPP_CONFIG_.RECOMMENDATION_HEADERS);
    var opportunities = MD17_ensureSheet_(ss, MD17_OPP_CONFIG_.OPPORTUNITIES_SHEET, MD17_OPP_CONFIG_.OPPORTUNITY_HEADERS);
    var reviewHeaders = MD17_headerMap_(review, MD17_OPP_CONFIG_.RECOMMENDATION_HEADERS);
    var oppHeaders = MD17_headerMap_(opportunities, MD17_OPP_CONFIG_.OPPORTUNITY_HEADERS);
    var rowNumber = MD17_findIdRow_(review, reviewHeaders.Recommendation_ID, recommendationId);
    if (!rowNumber) throw new Error('Opportunity recommendation not found: ' + recommendationId);

    var row = review.getRange(rowNumber, 1, 1, MD17_OPP_CONFIG_.RECOMMENDATION_HEADERS.length).getValues()[0];
    var sourceUrl = String(row[reviewHeaders.Source_URL] || '').trim();
    if (!MD17_isHttpUrl_(sourceUrl)) {
      throw new Error('The opportunity must have a valid official source URL before publication.');
    }
    // Portal approval is the authoritative editorial decision. Persist the
    // public statuses here so no manual CMS edit is required.
    review.getRange(rowNumber, reviewHeaders.Verification_Status + 1).setValue('Verified');
    review.getRange(rowNumber, reviewHeaders.Editorial_Status + 1).setValue('Approved');
    row[reviewHeaders.Verification_Status] = 'Verified';
    row[reviewHeaders.Editorial_Status] = 'Approved';

    var publishedId = String(row[reviewHeaders.Published_Opp_ID] || '').trim();
    var publicRow = publishedId ? MD17_findIdRow_(opportunities, oppHeaders.Opp_ID, publishedId) : 0;
    if (!publishedId) {
      var idInfo = MD17_nextId_(opportunities, oppHeaders.Opp_ID, 'OP');
      publishedId = idInfo.id;
      PropertiesService.getScriptProperties().setProperty(idInfo.counterKey, String(idInfo.sequence));
    }
    if (!publicRow) publicRow = opportunities.getLastRow() + 1;

    var publicValues = [];
    publicValues[oppHeaders.Opp_ID] = publishedId;
    publicValues[oppHeaders.Title] = row[reviewHeaders.Title];
    publicValues[oppHeaders.Organisation] = row[reviewHeaders.Organisation];
    publicValues[oppHeaders.Category] = row[reviewHeaders.Category];
    publicValues[oppHeaders.Deadline] = row[reviewHeaders.Deadline];
    publicValues[oppHeaders.Eligibility] = row[reviewHeaders.Eligibility];
    publicValues[oppHeaders.Apply_URL] = row[reviewHeaders.Source_URL];
    publicValues[oppHeaders.Verification_Status] = 'Verified';
    publicValues[oppHeaders.Editorial_Status] = 'Approved';
    publicValues[oppHeaders.Issue_Number] = row[reviewHeaders.Published_At] || '';
    for (var i = 0; i < MD17_OPP_CONFIG_.OPPORTUNITY_HEADERS.length; i++) {
      if (typeof publicValues[i] === 'undefined') publicValues[i] = '';
    }

    opportunities.getRange(publicRow, 1, 1, publicValues.length).setValues([publicValues]);
    review.getRange(rowNumber, reviewHeaders.Published_Opp_ID + 1).setValue(publishedId);
    review.getRange(rowNumber, reviewHeaders.Published_At + 1).setValue(row[reviewHeaders.Published_At] || new Date());
    MD17_styleDataRow_(opportunities, publicRow, publicValues.length);

    return { ok: true, publishedOppId: publishedId, recommendationId: recommendationId, actorEmail: actorEmail };
  } finally {
    lock.releaseLock();
  }
}
