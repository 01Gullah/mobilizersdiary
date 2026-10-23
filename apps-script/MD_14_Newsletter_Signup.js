/**
 * The Mobilizer's Diary — newsletter signup and unsubscribe API module
 * File name in Apps Script: MD_14_Newsletter_Signup.gs
 *
 * Connects to the existing doPost/apiRoute_ dispatcher:
 *   case 'subscribe': return subscribeNewsletter(d);
 *   case 'unsubscribe': return unsubscribeNewsletter(d.token);
 *
 * This file deliberately does NOT define doPost(e) or apiRoute_().
 * It uses the existing CMS_ID Script Property created by setupWorkspace().
 * It creates the Newsletter_Subscribers tab automatically if it is missing.
 *
 * Gmail/Brevo Script Properties (never put secret values in GitHub or chat):
 *   BREVO_API_KEY       Brevo API v3 key
 *   BREVO_SENDER_EMAIL  Sender address verified in Brevo
 *   BREVO_REPLY_TO_EMAIL (optional)
 *   NEWSLETTER_PORTAL_URL (optional; defaults to the public portal URL)
 *
 * Status is written as "Subscribed" / "Unsubscribed". The campaign queue
 * sender remains a separate module and must only queue subscribed recipients.
 */

var MD14_NEWSLETTER_CONFIG_ = {
  SHEET_NAME: 'Newsletter_Subscribers',
  HEADERS: [
    'Subscriber_ID', 'Timestamp', 'Email', 'Name', 'Subscription Status',
    'Source', 'Consent_Given', 'Consent_Text', 'Unsubscribe_Token',
    'Unsubscribed_At', 'Notes'
  ],
  CONSENT_TEXT: "I agree to receive The Mobilizer's Diary newsletter at this email address. I can unsubscribe at any time.",
  SENDER_NAME: "The Mobilizer's Diary",
  BREVO_ENDPOINT: 'https://api.brevo.com/v3/smtp/email',
  GMAIL_QUOTA_THRESHOLD: 5,
  DEFAULT_PORTAL_URL: 'https://01gullah.github.io/mobilizersdiary/'
};

/** API handler called by the existing action router for action="subscribe". */
function subscribeNewsletter(d) {
  d = d || {};

  // Silently accept the hidden honeypot field so bots do not learn the filter.
  if (String(d.hp || '').trim()) return { status: 'subscribed' };

  var email = MD14_cleanText_(d.email, 150).toLowerCase();
  if (!MD14_isValidEmail_(email) || /^[=+\-@]/.test(email)) {
    throw new Error('Please enter a valid email address.');
  }
  if (d.consent !== true && String(d.consent).toLowerCase() !== 'true') {
    throw new Error('Please tick the box to subscribe.');
  }
  var name = MD14_safeCellText_(MD14_cleanText_(d.name, 100));
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  var sheet;
  var columns;
  var subscriberRow = 0;
  var token = '';
  var alreadySubscribed = false;
  var now = new Date();

  try {
    sheet = MD14_getNewsletterSheet_(true);
    columns = MD14_getNewsletterColumns_(sheet);
    var lastRow = sheet.getLastRow();
    var lastMatch = 0;
    var activeMatch = 0;

    if (lastRow >= 2) {
      var emails = sheet.getRange(2, columns.email + 1, lastRow - 1, 1).getDisplayValues();
      var statuses = sheet.getRange(2, columns.status + 1, lastRow - 1, 1).getDisplayValues();
      for (var i = emails.length - 1; i >= 0; i--) {
        if (String(emails[i][0] || '').trim().toLowerCase() !== email) continue;
        var matchRow = i + 2;
        if (!lastMatch) lastMatch = matchRow;
        if (MD14_isSubscribedStatus_(statuses[i][0])) {
          activeMatch = matchRow;
          break;
        }
      }
    }

    if (activeMatch) {
      alreadySubscribed = true;
      subscriberRow = activeMatch;
    } else if (lastMatch) {
      // Re-subscription: record fresh consent and rotate the unsubscribe token.
      subscriberRow = lastMatch;
      token = MD14_newToken_();
      MD14_setCell_(sheet, subscriberRow, columns.timestamp, now);
      MD14_setCell_(sheet, subscriberRow, columns.email, email);
      MD14_setCell_(sheet, subscriberRow, columns.name, name);
      MD14_setCell_(sheet, subscriberRow, columns.status, 'Subscribed');
      MD14_setCell_(sheet, subscriberRow, columns.source, 'Portal');
      MD14_setCell_(sheet, subscriberRow, columns.consentGiven, 'Yes');
      MD14_setCell_(sheet, subscriberRow, columns.consentText, MD14_NEWSLETTER_CONFIG_.CONSENT_TEXT);
      MD14_setCell_(sheet, subscriberRow, columns.token, token);
      MD14_setCell_(sheet, subscriberRow, columns.unsubscribedAt, '');
      MD14_setCell_(sheet, subscriberRow, columns.notes, 'Re-subscribed with fresh consent; welcome email pending.');
    } else {
      token = MD14_newToken_();
      var newRow = [];
      for (var c = 0; c < sheet.getLastColumn(); c++) newRow.push('');
      newRow[columns.id] = 'NL-' + Utilities.getUuid().replace(/-/g, '').substring(0, 12).toUpperCase();
      newRow[columns.timestamp] = now;
      newRow[columns.email] = email;
      newRow[columns.name] = name;
      newRow[columns.status] = 'Subscribed';
      newRow[columns.source] = 'Portal';
      newRow[columns.consentGiven] = 'Yes';
      newRow[columns.consentText] = MD14_NEWSLETTER_CONFIG_.CONSENT_TEXT;
      newRow[columns.token] = token;
      newRow[columns.unsubscribedAt] = '';
      newRow[columns.notes] = 'Welcome email pending.';
      sheet.appendRow(newRow);
      subscriberRow = sheet.getLastRow();
    }
  } finally {
    lock.releaseLock();
  }

  if (alreadySubscribed) return { status: 'already' };

  // The subscription is saved before email delivery. If both providers fail,
  // keep the consent record, log the failure, and do not falsely undo signup.
  try {
    var method = MD14_sendWelcomeEmail_(email, name, token);
    MD14_setCell_(sheet, subscriberRow, columns.notes, 'Welcome email sent via ' + method + '.');
    Logger.log('Newsletter welcome email sent via ' + method + ' for subscriber row ' + subscriberRow + '.');
    return { status: 'subscribed', welcomeEmailSent: true, deliveryMethod: method };
  } catch (mailError) {
    var message = MD14_safeError_(mailError);
    MD14_setCell_(sheet, subscriberRow, columns.notes, 'Subscription saved; welcome email failed: ' + message);
    Logger.log('Newsletter signup saved for row ' + subscriberRow + ', but welcome email failed: ' + message);
    return { status: 'subscribed', welcomeEmailSent: false };
  }
}

/** API handler called by the existing action router for action="unsubscribe". */
function unsubscribeNewsletter(token) {
  token = String(token || '').trim();
  if (!/^[A-Za-z0-9_-]{40,100}$/.test(token)) {
    throw new Error('This unsubscribe link is not valid.');
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = MD14_getNewsletterSheet_(false);
    if (!sheet || sheet.getLastRow() < 2) throw new Error('This unsubscribe link is not valid.');
    var columns = MD14_getNewsletterColumns_(sheet);
    var tokens = sheet.getRange(2, columns.token + 1, sheet.getLastRow() - 1, 1).getDisplayValues();
    for (var i = 0; i < tokens.length; i++) {
      if (String(tokens[i][0] || '') === token) {
        var row = i + 2;
        MD14_setCell_(sheet, row, columns.status, 'Unsubscribed');
        MD14_setCell_(sheet, row, columns.unsubscribedAt, new Date());
        MD14_setCell_(sheet, row, columns.notes, 'Unsubscribed using the email link.');
        return { status: 'unsubscribed' };
      }
    }
    throw new Error('This unsubscribe link is not valid.');
  } finally {
    lock.releaseLock();
  }
}

/** Optional one-time setup. Creates the tab if needed without replacing existing data. */
function MD14_setupNewsletterTab() {
  var sheet = MD14_getNewsletterSheet_(true);
  Logger.log('Newsletter subscriber tab ready: ' + sheet.getName());
  return 'Newsletter subscriber tab ready.';
}

/** Read-only diagnostic; it does not create or modify the sheet. */
function MD14_checkNewsletterSetup() {
  var spreadsheet = MD14_getCmsSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(MD14_NEWSLETTER_CONFIG_.SHEET_NAME);
  Logger.log('CMS spreadsheet: ' + spreadsheet.getName());
  Logger.log('Newsletter_Subscribers tab: ' + (sheet ? 'found' : 'not created yet; first valid signup will create it'));
  Logger.log('Brevo API key configured: ' + Boolean(PropertiesService.getScriptProperties().getProperty('BREVO_API_KEY')));
  Logger.log('Brevo sender configured: ' + Boolean(PropertiesService.getScriptProperties().getProperty('BREVO_SENDER_EMAIL')));
  return { sheetFound: Boolean(sheet) };
}

function MD14_getCmsSpreadsheet_() {
  var id = PropertiesService.getScriptProperties().getProperty('CMS_ID');
  if (!id) throw new Error('CMS_ID is not configured. Run the existing workspace setup and confirm its Script Property.');
  return SpreadsheetApp.openById(id);
}

function MD14_getNewsletterSheet_(createIfMissing) {
  var spreadsheet = MD14_getCmsSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(MD14_NEWSLETTER_CONFIG_.SHEET_NAME);
  if (!sheet && createIfMissing) {
    sheet = spreadsheet.insertSheet(MD14_NEWSLETTER_CONFIG_.SHEET_NAME);
  }
  if (!sheet) return null;

  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0 || !String(sheet.getRange(1, 1).getDisplayValue() || '').trim()) {
    var headers = MD14_NEWSLETTER_CONFIG_.HEADERS;
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#123F3B').setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setTabColor('#C59B35');
    sheet.setColumnWidths(1, headers.length, 150);
    sheet.setColumnWidth(3, 240);
    sheet.setColumnWidth(8, 380);
    sheet.setColumnWidth(11, 300);
    sheet.getRange(2, 3, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('@');
  }
  // Validate schema before any data is written. Existing data is never cleared.
  MD14_getNewsletterColumns_(sheet);
  return sheet;
}

function MD14_getNewsletterColumns_(sheet) {
  var lastColumn = sheet.getLastColumn();
  if (lastColumn < 1 || sheet.getLastRow() < 1) throw new Error('Newsletter subscriber sheet needs a header row.');
  var headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0]
    .map(function (value) { return String(value || '').trim().toLowerCase(); });

  function findHeader(names) {
    for (var i = 0; i < names.length; i++) {
      var index = headers.indexOf(names[i].toLowerCase());
      if (index >= 0) return index;
    }
    throw new Error('Newsletter subscriber sheet is missing required header: ' + names[0]);
  }

  return {
    id: findHeader(['subscriber_id', 'subscriber id']),
    timestamp: findHeader(['timestamp', 'date received']),
    email: findHeader(['email']),
    name: findHeader(['name', 'display_name', 'display name']),
    status: findHeader(['subscription status', 'status']),
    source: findHeader(['source']),
    consentGiven: findHeader(['consent_given', 'consent given']),
    consentText: findHeader(['consent_text', 'consent text']),
    token: findHeader(['unsubscribe_token', 'unsubscribe token']),
    unsubscribedAt: findHeader(['unsubscribed_at', 'unsubscribed at']),
    notes: findHeader(['notes'])
  };
}

function MD14_setCell_(sheet, row, zeroBasedColumn, value) {
  sheet.getRange(row, zeroBasedColumn + 1).setValue(value);
}

function MD14_isSubscribedStatus_(value) {
  var status = String(value || '').trim().toLowerCase();
  // Accept "Active" for compatibility with the earlier proposed sheet schema.
  return status === 'subscribed' || status === 'active';
}

function MD14_newToken_() {
  return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
}

function MD14_cleanText_(value, maxLength) {
  return String(value == null ? '' : value)
    .replace(/[\u0000-\u001F\u007F]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, maxLength);
}

function MD14_safeCellText_(value) {
  value = String(value || '');
  return /^[=+\-@\t\r]/.test(value) ? "'" + value : value;
}

function MD14_isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || '').trim());
}

function MD14_sendWelcomeEmail_(email, name, token) {
  var props = PropertiesService.getScriptProperties();
  var portalUrl = props.getProperty('NEWSLETTER_PORTAL_URL') || MD14_NEWSLETTER_CONFIG_.DEFAULT_PORTAL_URL;
  portalUrl = String(portalUrl).trim().replace(/\/+$/, '');
  if (portalUrl.indexOf('https://') !== 0) throw new Error('NEWSLETTER_PORTAL_URL must use HTTPS.');
  var unsubscribeUrl = portalUrl + '/?unsubscribe=' + encodeURIComponent(token);
  var greeting = name ? 'Hello ' + MD14_escapeHtml_(name) + ',' : 'Hello,';
  var html = '<div style="margin:0;padding:24px;background:#f7f4ec;font-family:Arial,sans-serif;color:#123F3B">' +
    '<div style="max-width:600px;margin:0 auto;padding:24px;background:#ffffff;border-top:5px solid #C59B35">' +
    '<p style="margin:0 0 18px;color:#123F3B;font-weight:700;letter-spacing:2px">THE MOBILIZER\'S DIARY</p>' +
    '<h1 style="margin:0 0 16px;color:#123F3B;font-size:24px">Welcome to the newsletter</h1>' +
    '<p>' + greeting + '</p>' +
    '<p>Thank you for subscribing to The Mobilizer\'s Diary. The newsletter is free, short, and written for students.</p>' +
    '<p>Each issue shares student experiences, constructive solutions, campus highlights, and useful opportunities.</p>' +
    '<p>Your subscription is separate from any report or support request you have made.</p>' +
    '<hr style="border:0;border-top:1px solid #ded8c7;margin:24px 0">' +
    '<p style="font-size:12px;color:#53635f">You can <a href="' + MD14_escapeHtml_(unsubscribeUrl) + '" style="color:#123F3B">unsubscribe at any time</a>.</p>' +
    '</div></div>';
  var text = (name ? 'Hello ' + name + ',' : 'Hello,') +
    '\n\nThank you for subscribing to The Mobilizer\'s Diary. The newsletter is free, short, and written for students.' +
    '\n\nEach issue shares student experiences, constructive solutions, campus highlights, and useful opportunities.' +
    '\n\nYour subscription is separate from any report or support request you have made.' +
    '\n\nUnsubscribe at any time: ' + unsubscribeUrl;
  return MD14_sendHtmlEmailWithFallback_(email, 'Welcome to The Mobilizer\'s Diary Newsletter', html, text);
}

function MD14_sendHtmlEmailWithFallback_(email, subject, htmlBody, textBody) {
  var remainingQuota = 0;
  try {
    remainingQuota = MailApp.getRemainingDailyQuota();
  } catch (quotaError) {
    Logger.log('Newsletter quota check failed; trying Brevo: ' + MD14_safeError_(quotaError));
    remainingQuota = 0;
  }

  if (remainingQuota > MD14_NEWSLETTER_CONFIG_.GMAIL_QUOTA_THRESHOLD) {
    try {
      var options = { htmlBody: htmlBody, name: MD14_NEWSLETTER_CONFIG_.SENDER_NAME };
      var replyTo = PropertiesService.getScriptProperties().getProperty('BREVO_REPLY_TO_EMAIL');
      if (replyTo) options.replyTo = replyTo;
      GmailApp.sendEmail(email, subject, textBody, options);
      return 'Gmail';
    } catch (gmailError) {
      Logger.log('Gmail welcome email failed; trying Brevo: ' + MD14_safeError_(gmailError));
      try {
        return MD14_sendViaBrevo_(email, subject, htmlBody, textBody);
      } catch (brevoFallbackError) {
        throw new Error('Gmail failed: ' + MD14_safeError_(gmailError) + ' | Brevo fallback failed: ' + MD14_safeError_(brevoFallbackError));
      }
    }
  }

  Logger.log('Gmail quota is at or below ' + MD14_NEWSLETTER_CONFIG_.GMAIL_QUOTA_THRESHOLD + '; trying Brevo for welcome email.');
  return MD14_sendViaBrevo_(email, subject, htmlBody, textBody);
}

function MD14_sendViaBrevo_(email, subject, htmlBody, textBody) {
  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty('BREVO_API_KEY');
  var senderEmail = String(props.getProperty('BREVO_SENDER_EMAIL') || '').trim();
  if (!apiKey) throw new Error('BREVO_API_KEY is not configured in Apps Script Script Properties.');
  if (!MD14_isValidEmail_(senderEmail)) throw new Error('BREVO_SENDER_EMAIL must be a valid, Brevo-verified sender address.');

  var payload = {
    sender: { name: MD14_NEWSLETTER_CONFIG_.SENDER_NAME, email: senderEmail },
    to: [{ email: email }],
    subject: subject,
    htmlContent: htmlBody,
    textContent: textBody
  };
  var replyTo = String(props.getProperty('BREVO_REPLY_TO_EMAIL') || '').trim();
  if (replyTo) {
    if (!MD14_isValidEmail_(replyTo)) throw new Error('BREVO_REPLY_TO_EMAIL is not a valid email address.');
    payload.replyTo = { email: replyTo };
  }

  var response = UrlFetchApp.fetch(MD14_NEWSLETTER_CONFIG_.BREVO_ENDPOINT, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'api-key': apiKey, 'accept': 'application/json' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  var code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Brevo returned HTTP ' + code + ': ' + String(response.getContentText() || '').substring(0, 300));
  }
  Logger.log('Brevo accepted the newsletter welcome email (HTTP ' + code + ').');
  return 'Brevo';
}

function MD14_escapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function MD14_safeError_(error) {
  var message = error && error.message ? error.message : String(error || 'Unknown error');
  return String(message).substring(0, 700);
}
