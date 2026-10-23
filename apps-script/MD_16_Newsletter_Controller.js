/*
 * THE MOBILIZER'S DIARY — AUTOMATED NEWSLETTER CONTROLLER
 *
 * Add this as one new Apps Script file named MD_16_Newsletter_Controller.gs.
 * It uses the existing CMS_ID Script Property and Newsletter_Subscribers tab.
 * It automatically creates and formats Newsletter_Campaigns and NewsletterQueue.
 * It does NOT replace setupWorkspace(), doPost(e), apiRoute_(), or MD_14 handlers.
 *
 * WORKFLOW
 * 1) Run MD16_setupNewsletterController() once. It creates both control tabs.
 * 2) Run MD16_createCampaignDraft() to create a campaign row.
 * 3) In Newsletter_Campaigns, fill Issue_Title, Subject, HTML_Body (optional if
 *    the MD_Launch_Edition_Email template file exists), and Text_Body. Verified +
 *    Approved rows from Opportunities are auto-featured in the email. Custom HTML
 *    may use {{MD16_OPPORTUNITIES}} as its placement marker; otherwise cards append.
 * 4) Set Approval_Status to Approved using the dropdown after editorial review.
 * 5) Run MD16_prepareApprovedCampaigns(). It adds one queue row per currently
 *    subscribed contact with recorded consent and a valid unsubscribe token.
 * 6) Review NewsletterQueue, then run MD16_sendPendingNewsletterBatch(). This
 *    sends at most 5 pending rows per run. Re-run only when you intend to send
 *    the next batch. No trigger or automatic mass send is installed.
 *
 * REQUIRED Script Properties already used by this project:
 *   CMS_ID                 ID of the existing CMS workbook (setupWorkspace creates it)
 *   BREVO_API_KEY          Regular Brevo API v3 key (used when Gmail quota is low/fails)
 *   BREVO_SENDER_EMAIL     Sender address verified in Brevo (needed for Brevo fallback)
 * Optional:
 *   BREVO_REPLY_TO_EMAIL
 *   NEWSLETTER_PORTAL_URL  defaults to the existing GitHub Pages portal
 *
 * SAFETY
 * - Queue preparation includes only status=Subscribed AND Consent_Given=Yes rows.
 * - Immediately before sending, live subscription status/consent/token are checked again.
 * - Every message receives an individualized unsubscribe URL.
 * - Failed sends are not automatically retried. Provider-accepted messages that
 *   cannot be recorded become Needs Review to avoid accidental duplicates.
 */

var MD16_CONFIG_ = {
  CAMPAIGNS_SHEET: 'Newsletter_Campaigns',
  SUBSCRIBERS_SHEET: 'Newsletter_Subscribers',
  OPPORTUNITIES_SHEET: 'Opportunities',
  QUEUE_SHEET: 'NewsletterQueue',
  CAMPAIGN_HEADERS: [
    'Campaign_ID', 'Issue_Title', 'Subject', 'HTML_Body', 'Text_Body',
    'Approval_Status', 'Queue_Status', 'Created_At', 'Approved_At',
    'Prepared_At', 'Notes'
  ],
  QUEUE_HEADERS: [
    'Queue_ID', 'Campaign_ID', 'Email', 'Subscription_Status', 'Consent_Given',
    'Unsubscribe_Token', 'Subject', 'HTML_Body', 'Text_Body', 'Status',
    'Sent_At', 'Error', 'Delivery_Method', 'Created_At'
  ],
  SENDER_NAME: "The Mobilizer's Diary",
  BREVO_ENDPOINT: 'https://api.brevo.com/v3/smtp/email',
  GMAIL_QUOTA_THRESHOLD: 5,
  DEFAULT_PORTAL_URL: 'https://01gullah.github.io/mobilizersdiary/',
  DEFAULT_OPPORTUNITIES_PAGE_URL: 'https://01gullah.github.io/mobilizersdiary/opportunities.html',
  MAX_NEWSLETTER_OPPORTUNITIES: 3,
  DEFAULT_BATCH_SIZE: 9,
  MAX_BATCH_SIZE: 25,
  BRAND_GREEN: '#123F3B',
  BRAND_GOLD: '#C59B35'
};

/** Run once to create/validate both controller tabs, styling and dropdowns. */
function MD16_setupNewsletterController() {
  var ss = MD16_getSpreadsheet_();
  var campaigns = MD16_ensureSheet_(ss, MD16_CONFIG_.CAMPAIGNS_SHEET, MD16_CONFIG_.CAMPAIGN_HEADERS);
  var queue = MD16_ensureSheet_(ss, MD16_CONFIG_.QUEUE_SHEET, MD16_CONFIG_.QUEUE_HEADERS);
  MD16_applyCampaignFormatting_(campaigns);
  MD16_applyQueueFormatting_(queue);
  var subscribers = ss.getSheetByName(MD16_CONFIG_.SUBSCRIBERS_SHEET);
  Logger.log('Newsletter controller ready in workbook: ' + ss.getName());
  Logger.log('Campaigns tab: ready; queue tab: ready; subscribers tab: ' + (subscribers ? 'found' : 'not created yet (MD_14 creates it on first valid signup or via MD14_setupNewsletterTab).'));
  Logger.log('Opportunities tab: ' + (ss.getSheetByName(MD16_CONFIG_.OPPORTUNITIES_SHEET) ? 'found; verified + approved listings will be featured' : 'not found; newsletter opportunity section will be empty until MD17 creates the tab.'));
  return { ok: true, spreadsheetName: ss.getName(), campaignsSheet: campaigns.getName(), queueSheet: queue.getName(), subscribersSheetFound: Boolean(subscribers) };
}

/** Read-only check; does not create or alter workbook tabs. */
function MD16_checkNewsletterController() {
  var ss = MD16_getSpreadsheet_();
  var result = {
    workbook: ss.getName(),
    campaignsSheetFound: Boolean(ss.getSheetByName(MD16_CONFIG_.CAMPAIGNS_SHEET)),
    queueSheetFound: Boolean(ss.getSheetByName(MD16_CONFIG_.QUEUE_SHEET)),
    subscribersSheetFound: Boolean(ss.getSheetByName(MD16_CONFIG_.SUBSCRIBERS_SHEET)),
    opportunitiesSheetFound: Boolean(ss.getSheetByName(MD16_CONFIG_.OPPORTUNITIES_SHEET)),
    opportunitiesPageUrl: MD16_getOpportunitiesPageUrl_(),
    brevoKeyConfigured: Boolean(PropertiesService.getScriptProperties().getProperty('BREVO_API_KEY')),
    brevoSenderConfigured: Boolean(PropertiesService.getScriptProperties().getProperty('BREVO_SENDER_EMAIL'))
  };
  Logger.log(JSON.stringify(result));
  return result;
}

/** Add a clean Draft row and return its campaign ID. */
function MD16_createCampaignDraft() {
  var ss = MD16_getSpreadsheet_();
  var sh = MD16_ensureSheet_(ss, MD16_CONFIG_.CAMPAIGNS_SHEET, MD16_CONFIG_.CAMPAIGN_HEADERS);
  MD16_applyCampaignFormatting_(sh);
  var id = 'MD-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss') + '-' + Utilities.getUuid().substring(0, 6).toUpperCase();
  var row = [id, '', '', '', '', 'Draft', 'Not Prepared', new Date(), '', '', ''];
  sh.appendRow(row);
  var rowNumber = sh.getLastRow();
  sh.getRange(rowNumber, 1, 1, MD16_CONFIG_.CAMPAIGN_HEADERS.length).setVerticalAlignment('top');
  Logger.log('Draft created in ' + MD16_CONFIG_.CAMPAIGNS_SHEET + ' row ' + rowNumber + ': ' + id);
  return { ok: true, campaignId: id, row: rowNumber, instruction: 'Open Newsletter_Campaigns and complete this row. After review, select Approved in Approval_Status.' };
}

/**
 * Prepare the queue for every Approved campaign not yet Prepared.
 * Idempotent: existing Campaign_ID + Email pairs are never duplicated.
 */
function MD16_prepareApprovedCampaigns() {
  var ss = MD16_getSpreadsheet_();
  var campaigns = MD16_ensureSheet_(ss, MD16_CONFIG_.CAMPAIGNS_SHEET, MD16_CONFIG_.CAMPAIGN_HEADERS);
  var queue = MD16_ensureSheet_(ss, MD16_CONFIG_.QUEUE_SHEET, MD16_CONFIG_.QUEUE_HEADERS);
  MD16_applyCampaignFormatting_(campaigns);
  MD16_applyQueueFormatting_(queue);
  var subscribers = ss.getSheetByName(MD16_CONFIG_.SUBSCRIBERS_SHEET);
  if (!subscribers) throw new Error('Newsletter_Subscribers tab is not present yet. It is created by MD_14_Newsletter_Signup.gs after a valid signup, or run MD14_setupNewsletterTab() once.');

  var c = MD16_headerMap_(campaigns, MD16_CONFIG_.CAMPAIGN_HEADERS);
  var s = MD16_subscriberHeaderMap_(subscribers);
  var q = MD16_headerMap_(queue, MD16_CONFIG_.QUEUE_HEADERS);
  var campaignLast = campaigns.getLastRow();
  if (campaignLast < 2) return { ok: true, prepared: 0, message: 'No campaign rows yet. Run MD16_createCampaignDraft() first.' };

  var existing = {};
  if (queue.getLastRow() >= 2) {
    var priorRows = queue.getRange(2, 1, queue.getLastRow() - 1, queue.getLastColumn()).getDisplayValues();
    priorRows.forEach(function (r) {
      var cid = String(r[q.Campaign_ID] || '').trim();
      var email = String(r[q.Email] || '').trim().toLowerCase();
      if (cid && email) existing[cid + '|' + email] = true;
    });
  }

  var subscribed = MD16_getEligibleSubscribers_(subscribers, s);
  var data = campaigns.getRange(2, 1, campaignLast - 1, campaigns.getLastColumn()).getValues();
  var preparedTotal = 0;
  var now = new Date();
  for (var i = 0; i < data.length; i++) {
    var sheetRow = i + 2;
    var row = data[i];
    var campaignId = String(row[c.Campaign_ID] || '').trim();
    var approval = String(row[c.Approval_Status] || '').trim().toLowerCase();
    var queueStatus = String(row[c.Queue_Status] || '').trim().toLowerCase();
    if (!campaignId || approval !== 'approved' || queueStatus === 'prepared' || queueStatus === 'complete') continue;

    var title = String(row[c.Issue_Title] || '').trim();
    var subject = String(row[c.Subject] || '').trim();
    var sourceHtml = String(row[c.HTML_Body] || '').trim();
    var sourceText = String(row[c.Text_Body] || '').trim();
    if (!subject) throw new Error('Campaign ' + campaignId + ' is Approved but Subject is blank (row ' + sheetRow + ').');
    if (!sourceHtml && !MD16_hasEmailTemplate_()) throw new Error('Campaign ' + campaignId + ' has no HTML_Body, and HTML file MD_Launch_Edition_Email is not available. Add email HTML to the campaign row or add that template file.');
    if (!subscribed.length) {
      campaigns.getRange(sheetRow, c.Queue_Status + 1).setValue('No Eligible Subscribers');
      campaigns.getRange(sheetRow, c.Notes + 1).setValue('No active subscriber with explicit consent and a valid unsubscribe token was found. No email was queued.');
      continue;
    }

    var append = [];
    var opportunityItems = MD16_getVerifiedOpportunitiesForNewsletter_();
    var opportunitiesPageUrl = MD16_getOpportunitiesPageUrl_();
    for (var j = 0; j < subscribed.length; j++) {
      var person = subscribed[j];
      var key = campaignId + '|' + person.email;
      if (existing[key]) continue;
      var unsubscribeUrl = MD16_unsubscribeUrl_(person.token);
      var html = sourceHtml
        ? MD16_appendUnsubscribeHtml_(MD16_injectOpportunitySection_(sourceHtml, opportunityItems, opportunitiesPageUrl), unsubscribeUrl)
        : MD16_renderTemplate_(title, unsubscribeUrl, opportunityItems, opportunitiesPageUrl);
      var text = sourceText
        ? MD16_appendOpportunityText_(sourceText, opportunityItems, opportunitiesPageUrl)
        : MD16_htmlToText_(html);
      text += '\n\nYou received this because you subscribed to The Mobilizer\'s Diary. Unsubscribe: ' + unsubscribeUrl;
      var queueId = 'Q-' + Utilities.getUuid().replace(/-/g, '').substring(0, 14).toUpperCase();
      var newRow = [];
      for (var k = 0; k < MD16_CONFIG_.QUEUE_HEADERS.length; k++) newRow.push('');
      newRow[q.Queue_ID] = queueId;
      newRow[q.Campaign_ID] = campaignId;
      newRow[q.Email] = person.email;
      newRow[q.Subscription_Status] = 'Subscribed';
      newRow[q.Consent_Given] = 'Yes';
      newRow[q.Unsubscribe_Token] = person.token;
      newRow[q.Subject] = subject;
      newRow[q.HTML_Body] = html;
      newRow[q.Text_Body] = text;
      newRow[q.Status] = 'Pending';
      newRow[q.Created_At] = now;
      append.push(newRow);
      existing[key] = true;
    }
    if (append.length) {
      queue.getRange(queue.getLastRow() + 1, 1, append.length, MD16_CONFIG_.QUEUE_HEADERS.length).setValues(append);
      preparedTotal += append.length;
    }
    campaigns.getRange(sheetRow, c.Queue_Status + 1).setValue('Prepared');
    campaigns.getRange(sheetRow, c.Prepared_At + 1).setValue(now);
    if (!row[c.Approved_At]) campaigns.getRange(sheetRow, c.Approved_At + 1).setValue(now);
    campaigns.getRange(sheetRow, c.Notes + 1).setValue('Queue prepared. Review queue rows before sending. Existing campaign/email pairs were not duplicated.');
  }
  Logger.log('Newsletter controller prepared ' + preparedTotal + ' new queue row(s). No email was sent.');
  return { ok: true, prepared: preparedTotal, message: 'Queue prepared only; no email was sent. Review NewsletterQueue before sending.' };
}

/** Send up to 5 Pending queue rows from Approved campaigns. Never retries failed rows. */
function MD16_sendPendingNewsletterBatch(maxRows) {
  var limit = Number(maxRows) || MD16_CONFIG_.DEFAULT_BATCH_SIZE;
  limit = Math.max(1, Math.min(Math.floor(limit), MD16_CONFIG_.MAX_BATCH_SIZE));
  var ss = MD16_getSpreadsheet_();
  var queue = ss.getSheetByName(MD16_CONFIG_.QUEUE_SHEET);
  var campaigns = ss.getSheetByName(MD16_CONFIG_.CAMPAIGNS_SHEET);
  if (!queue || !campaigns) throw new Error('Controller tabs missing. Run MD16_setupNewsletterController() first.');
  var q = MD16_headerMap_(queue, MD16_CONFIG_.QUEUE_HEADERS);
  var c = MD16_headerMap_(campaigns, MD16_CONFIG_.CAMPAIGN_HEADERS);
  if (queue.getLastRow() < 2) return { ok: true, attempted: 0, message: 'Queue is empty.' };

  var approvedCampaigns = {};
  if (campaigns.getLastRow() >= 2) {
    var campaignData = campaigns.getRange(2, 1, campaigns.getLastRow() - 1, campaigns.getLastColumn()).getValues();
    campaignData.forEach(function (r) {
      if (String(r[c.Approval_Status] || '').trim().toLowerCase() === 'approved') approvedCampaigns[String(r[c.Campaign_ID] || '').trim()] = true;
    });
  }
  var ids = queue.getRange(2, q.Queue_ID + 1, queue.getLastRow() - 1, 1).getDisplayValues();
  var results = [];
  for (var i = 0; i < ids.length && results.length < limit; i++) {
    var rowNumber = i + 2;
    var current = queue.getRange(rowNumber, 1, 1, queue.getLastColumn()).getValues()[0];
    if (String(current[q.Status] || '').trim().toLowerCase() !== 'pending') continue;
    if (!approvedCampaigns[String(current[q.Campaign_ID] || '').trim()]) {
      queue.getRange(rowNumber, q.Status + 1).setValue('Needs Review');
      queue.getRange(rowNumber, q.Error + 1).setValue('Campaign is no longer marked Approved; no email was sent.');
      continue;
    }
    results.push(MD16_sendOneQueueRow_(ss, queue, rowNumber, q, current));
  }
  Logger.log('Newsletter send batch finished: attempted ' + results.length + ' row(s).');
  return { ok: true, attempted: results.length, results: results, note: 'Only Pending rows under Approved campaigns are considered. Failed rows are not automatically retried.' };
}

/** Send one explicitly selected queue row; the row must still be Pending and approved. */
function MD16_sendOneNewsletterRow(rowNumber) {
  var ss = MD16_getSpreadsheet_();
  var queue = ss.getSheetByName(MD16_CONFIG_.QUEUE_SHEET);
  var campaigns = ss.getSheetByName(MD16_CONFIG_.CAMPAIGNS_SHEET);
  if (!queue || !campaigns) throw new Error('Controller tabs missing. Run MD16_setupNewsletterController() first.');
  var n = Number(rowNumber);
  if (!Number.isInteger(n) || n < 2 || n > queue.getLastRow()) throw new Error('Enter a valid NewsletterQueue row number (2 or greater).');
  var q = MD16_headerMap_(queue, MD16_CONFIG_.QUEUE_HEADERS);
  var c = MD16_headerMap_(campaigns, MD16_CONFIG_.CAMPAIGN_HEADERS);
  var vals = campaigns.getLastRow() >= 2 ? campaigns.getRange(2, 1, campaigns.getLastRow() - 1, campaigns.getLastColumn()).getValues() : [];
  var approved = false;
  var campaignId = String(queue.getRange(n, q.Campaign_ID + 1).getDisplayValue() || '').trim();
  vals.forEach(function (r) {
    if (String(r[c.Campaign_ID] || '').trim() === campaignId && String(r[c.Approval_Status] || '').trim().toLowerCase() === 'approved') approved = true;
  });
  if (!approved) throw new Error('The campaign linked to this row is not Approved. No email was sent.');
  var row = queue.getRange(n, 1, 1, queue.getLastColumn()).getValues()[0];
  return MD16_sendOneQueueRow_(ss, queue, n, q, row);
}

function MD16_sendOneQueueRow_(ss, queue, rowNumber, q, row) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  var accepted = false;
  try {
    // Re-read after acquiring the lock so a second concurrent invocation cannot
    // send a row that another invocation already moved out of Pending.
    row = queue.getRange(rowNumber, 1, 1, queue.getLastColumn()).getValues()[0];
    var status = String(row[q.Status] || '').trim().toLowerCase();
    if (status !== 'pending') return { row: rowNumber, skipped: true, status: status };
    var email = String(row[q.Email] || '').trim().toLowerCase();
    var subject = String(row[q.Subject] || '').trim();
    var html = String(row[q.HTML_Body] || '').trim();
    var text = String(row[q.Text_Body] || '').trim();
    if (!MD16_isEmail_(email) || !subject || !html) throw new Error('Queue row has an invalid email, blank subject, or blank HTML body.');
    if (!MD16_confirmCurrentConsent_(ss, email, String(row[q.Unsubscribe_Token] || ''))) {
      queue.getRange(rowNumber, q.Status + 1).setValue('Cancelled');
      queue.getRange(rowNumber, q.Error + 1).setValue('Subscriber is no longer actively subscribed with matching consent/token. No email was sent.');
      return { row: rowNumber, status: 'Cancelled', message: 'Subscription changed or consent/token no longer matches.' };
    }
    queue.getRange(rowNumber, q.Status + 1).setValue('Sending');
    queue.getRange(rowNumber, q.Error + 1).clearContent();
    SpreadsheetApp.flush();

    var method = MD16_sendWithFallback_(email, subject, html, text || MD16_htmlToText_(html));
    accepted = true;
    try {
      queue.getRange(rowNumber, q.Status + 1).setValue('Sent');
      queue.getRange(rowNumber, q.Sent_At + 1).setValue(new Date());
      queue.getRange(rowNumber, q.Error + 1).clearContent();
      queue.getRange(rowNumber, q.Delivery_Method + 1).setValue(method);
      SpreadsheetApp.flush();
    } catch (writeError) {
      try {
        queue.getRange(rowNumber, q.Status + 1).setValue('Needs Review');
        queue.getRange(rowNumber, q.Error + 1).setValue('Provider accepted email; verify delivery before retrying. ' + MD16_error_(writeError));
      } catch (ignored) {}
      return { row: rowNumber, status: 'Needs Review', message: 'Provider accepted email, but the sheet could not confirm Sent. Verify before retrying.' };
    }
    return { row: rowNumber, status: 'Sent', method: method };
  } catch (error) {
    var message = MD16_error_(error);
    if (!accepted) {
      try {
        queue.getRange(rowNumber, q.Status + 1).setValue('Failed');
        queue.getRange(rowNumber, q.Error + 1).setValue(message);
        SpreadsheetApp.flush();
      } catch (statusError) { Logger.log('Could not record send failure: ' + MD16_error_(statusError)); }
    }
    Logger.log('Newsletter queue row ' + rowNumber + ' failed: ' + message);
    return { row: rowNumber, status: accepted ? 'Needs Review' : 'Failed', error: message };
  } finally {
    lock.releaseLock();
  }
}

function MD16_sendWithFallback_(email, subject, html, text) {
  var quota = 0;
  try { quota = MailApp.getRemainingDailyQuota(); }
  catch (e) { Logger.log('Could not read Gmail quota; using Brevo. ' + MD16_error_(e)); }
  if (quota > MD16_CONFIG_.GMAIL_QUOTA_THRESHOLD) {
    try {
      var options = { htmlBody: html, name: MD16_CONFIG_.SENDER_NAME };
      var replyTo = PropertiesService.getScriptProperties().getProperty('BREVO_REPLY_TO_EMAIL');
      if (replyTo) options.replyTo = replyTo;
      GmailApp.sendEmail(email, subject, text, options);
      return 'Gmail';
    } catch (gmailError) {
      Logger.log('Gmail send failed; trying Brevo. ' + MD16_error_(gmailError));
      MD16_sendViaBrevo_(email, subject, html, text);
      return 'Brevo (Gmail fallback)';
    }
  }
  MD16_sendViaBrevo_(email, subject, html, text);
  return 'Brevo (quota fallback)';
}

function MD16_sendViaBrevo_(email, subject, html, text) {
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('BREVO_API_KEY');
  var sender = String(props.getProperty('BREVO_SENDER_EMAIL') || '').trim();
  if (!key) throw new Error('BREVO_API_KEY is missing from Script Properties.');
  if (!MD16_isEmail_(sender)) throw new Error('BREVO_SENDER_EMAIL must be set to the verified Brevo sender address.');
  var payload = { sender: { name: MD16_CONFIG_.SENDER_NAME, email: sender }, to: [{ email: email }], subject: subject, htmlContent: html, textContent: text };
  var replyTo = String(props.getProperty('BREVO_REPLY_TO_EMAIL') || '').trim();
  if (replyTo) payload.replyTo = { email: replyTo };
  var response = UrlFetchApp.fetch(MD16_CONFIG_.BREVO_ENDPOINT, {
    method: 'post', contentType: 'application/json',
    headers: { 'api-key': key, 'accept': 'application/json' },
    payload: JSON.stringify(payload), muteHttpExceptions: true
  });
  var code = response.getResponseCode();
  if (code < 200 || code >= 300) throw new Error('Brevo returned HTTP ' + code + ': ' + String(response.getContentText() || '').substring(0, 500));
}

function MD16_getEligibleSubscribers_(sheet, s) {
  var result = [];
  if (sheet.getLastRow() < 2) return result;
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  values.forEach(function (r) {
    var email = String(r[s.Email] || '').trim().toLowerCase();
    var status = String(r[s.Subscription_Status] || '').trim().toLowerCase();
    var consent = String(r[s.Consent_Given] || '').trim().toLowerCase();
    var token = String(r[s.Unsubscribe_Token] || '').trim();
    if (status === 'subscribed' && (consent === 'yes' || consent === 'true') && MD16_isEmail_(email) && /^[A-Za-z0-9_-]{40,100}$/.test(token)) {
      result.push({ email: email, token: token });
    }
  });
  return result;
}

function MD16_confirmCurrentConsent_(ss, email, token) {
  var sheet = ss.getSheetByName(MD16_CONFIG_.SUBSCRIBERS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return false;
  var s = MD16_subscriberHeaderMap_(sheet);
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    if (String(r[s.Email] || '').trim().toLowerCase() !== email) continue;
    var status = String(r[s.Subscription_Status] || '').trim().toLowerCase();
    var consent = String(r[s.Consent_Given] || '').trim().toLowerCase();
    var currentToken = String(r[s.Unsubscribe_Token] || '').trim();
    if (status === 'subscribed' && (consent === 'yes' || consent === 'true') && token && currentToken === token) return true;
  }
  return false;
}

function MD16_subscriberHeaderMap_(sheet) {
  var h = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(function (x) { return String(x || '').trim().toLowerCase(); });
  function find(options) {
    for (var i = 0; i < options.length; i++) { var n = h.indexOf(options[i].toLowerCase()); if (n >= 0) return n; }
    throw new Error('Newsletter_Subscribers is missing a required column: ' + options[0]);
  }
  return {
    Email: find(['Email']),
    Subscription_Status: find(['Subscription Status', 'Status']),
    Consent_Given: find(['Consent_Given', 'Consent Given']),
    Unsubscribe_Token: find(['Unsubscribe_Token', 'Unsubscribe Token'])
  };
}

function MD16_unsubscribeUrl_(token) {
  var base = String(PropertiesService.getScriptProperties().getProperty('NEWSLETTER_PORTAL_URL') || MD16_CONFIG_.DEFAULT_PORTAL_URL).trim();
  if (base.indexOf('https://') !== 0) throw new Error('NEWSLETTER_PORTAL_URL must use HTTPS.');
  base = base.replace(/\/+$/, '');
  return base + (base.indexOf('?') >= 0 ? '&' : '?') + 'unsubscribe=' + encodeURIComponent(token);
}

function MD16_appendUnsubscribeHtml_(html, url) {
  var safe = MD16_escapeHtml_(url);
  var footer = '<div style="margin:24px auto 0;padding:16px 12px;border-top:1px solid #ded8c7;color:#53635f;font:12px Arial,sans-serif;text-align:center">You received this because you subscribed to The Mobilizer\'s Diary. <a href="' + safe + '" style="color:#123F3B">Unsubscribe</a>.</div>';
  if (/<\/body\s*>/i.test(html)) return html.replace(/<\/body\s*>/i, footer + '</body>');
  return html + footer;
}

function MD16_renderTemplate_(issueTitle, unsubscribeUrl, opportunityItems, opportunitiesPageUrl) {
  var template;
  try {
    template = HtmlService.createTemplateFromFile('MD_Launch_Edition_Email');
  } catch (e) {
    throw new Error('HTML template MD_Launch_Edition_Email is missing.');
  }

  var portalUrl = String(
    PropertiesService.getScriptProperties().getProperty('NEWSLETTER_PORTAL_URL') ||
    MD16_CONFIG_.DEFAULT_PORTAL_URL
  ).trim();

  template.edition = {
    issueTitle: issueTitle || '',
    headline: issueTitle || '',
    founderDesk: {
      imageUrl: MD16_getFounderPortraitUrl_(),
      title: '',
      body: '',
      signatureName: '',
      signatureTitle: ''
    },
    historyCorner: [],
    studentVoices: [],
    podcast: [{ Episode_Title: 'Podcast coming soon', Guest: '', Release_Date: '', Duration: '', Summary: 'We are preparing conversations and stories from students, organisers and community leaders. This space will feature them in a future edition.', Audio_URL: '' }],
    opportunities: Array.isArray(opportunityItems) ? opportunityItems : MD16_getVerifiedOpportunitiesForNewsletter_(),
    opportunitiesPageUrl: opportunitiesPageUrl || MD16_getOpportunitiesPageUrl_(),
    portalUrl: portalUrl,
    unsubscribeUrl: unsubscribeUrl || ''
  };

  return template.evaluate().getContent();
}

/** Portrait is opt-in: leave NEWSLETTER_FOUNDER_PORTRAIT_URL blank to omit it. */
function MD16_getFounderPortraitUrl_() {
  var url = String(
    PropertiesService.getScriptProperties().getProperty('NEWSLETTER_FOUNDER_PORTRAIT_URL') || ''
  ).trim();
  if (!url) return '';
  if (url.indexOf('https://') !== 0) {
    throw new Error('NEWSLETTER_FOUNDER_PORTRAIT_URL must use HTTPS.');
  }
  return url;
}

/** Link to the dedicated opportunities page, configured in Script Properties. */
function MD16_getOpportunitiesPageUrl_() {
  var url = String(
    PropertiesService.getScriptProperties().getProperty('NEWSLETTER_OPPORTUNITIES_PAGE_URL') ||
    MD16_CONFIG_.DEFAULT_OPPORTUNITIES_PAGE_URL
  ).trim();
  if (url.indexOf('https://') !== 0) {
    throw new Error('NEWSLETTER_OPPORTUNITIES_PAGE_URL must use HTTPS.');
  }
  return url;
}

/** Read-only: select a maximum of three current CMS listings that are both Verified and Approved. */
function MD16_getVerifiedOpportunitiesForNewsletter_() {
  var ss = MD16_getSpreadsheet_();
  var sheet = ss.getSheetByName(MD16_CONFIG_.OPPORTUNITIES_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var headers = MD16_headerMap_(sheet, [
    'Opp_ID','Title','Organisation','Category','Deadline','Eligibility','Apply_URL',
    'Verification_Status','Editorial_Status'
  ]);
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues();
  var output = [];
  rows.forEach(function (row) {
    var verified = String(row[headers.Verification_Status] || '').trim().toLowerCase() === 'verified';
    var approved = String(row[headers.Editorial_Status] || '').trim().toLowerCase() === 'approved';
    var applyUrl = String(row[headers.Apply_URL] || '').trim();
    var title = String(row[headers.Title] || '').trim();
    var organisation = String(row[headers.Organisation] || '').trim();
    if (!verified || !approved || !title || !organisation || !MD16_isHttpUrl_(applyUrl)) return;
    var category = String(row[headers.Category] || '').trim();
    var deadline = String(row[headers.Deadline] || '').trim();
    var eligibility = String(row[headers.Eligibility] || '').trim();
    output.push({
      id: String(row[headers.Opp_ID] || '').trim(),
      title: title,
      subtitle: category,
      meta: [organisation, deadline ? 'Deadline: ' + deadline : ''].filter(Boolean).join(' · '),
      body: eligibility.substring(0, 600),
      url: applyUrl
    });
  });
  return output.slice(0, MD16_CONFIG_.MAX_NEWSLETTER_OPPORTUNITIES);
}

/** Manual, read-only diagnostic; logs approved CMS listings but never sends or queues email. */
function MD16_checkOpportunityIntegration() {
  var items = MD16_getVerifiedOpportunitiesForNewsletter_();
  var result = {
    ok: true,
    opportunitiesSheetFound: Boolean(MD16_getSpreadsheet_().getSheetByName(MD16_CONFIG_.OPPORTUNITIES_SHEET)),
    featuredCount: items.length,
    featured: items.map(function (item) { return { id: item.id, title: item.title }; }),
    opportunitiesPageUrl: MD16_getOpportunitiesPageUrl_(),
    note: 'Read-only check. No campaign was prepared and no email was sent.'
  };
  Logger.log(JSON.stringify(result));
  return result;
}

/** Render the real template with current CMS opportunities for inspection; never queues or sends. */
function MD16_testOpportunityNewsletterRender() {
  var items = MD16_getVerifiedOpportunitiesForNewsletter_();
  var pageUrl = MD16_getOpportunitiesPageUrl_();
  var html = MD16_renderTemplate_('Preview only — no email sent', '', items, pageUrl);
  var allItemsRendered = items.every(function (item) {
    return html.indexOf(MD16_escapeHtml_(item.title)) >= 0 && html.indexOf(MD16_escapeHtml_(item.url)) >= 0;
  });
  var result = {
    ok: true,
    opportunitiesSheetFound: Boolean(MD16_getSpreadsheet_().getSheetByName(MD16_CONFIG_.OPPORTUNITIES_SHEET)),
    templateRendered: /<html[\s>]/i.test(html),
    featuredCount: items.length,
    featuredTitles: items.map(function (item) { return item.title; }),
    opportunitySectionRendered: items.length > 0 && html.indexOf('Opportunities') >= 0,
    allFeaturedLinksAndTitlesRendered: items.length === 0 ? true : allItemsRendered,
    opportunitiesPageUrlConfigured: Boolean(pageUrl),
    noQueueCreated: true,
    noEmailSent: true,
    note: items.length ? 'Review this result; the approved items were rendered into the email template.' : 'Template rendered, but there are no Verified + Approved listings to display yet.'
  };
  Logger.log(JSON.stringify(result));
  return result;
}

/** Place dynamic opportunity cards at a marker, or append them before </body> in custom HTML. */
function MD16_injectOpportunitySection_(html, items, pageUrl) {
  html = String(html || '');
  var section = MD16_renderOpportunitySection_(items, pageUrl);
  var marker = '{{MD16_OPPORTUNITIES}}';
  if (html.indexOf(marker) >= 0) return html.split(marker).join(section);
  if (!section) return html;
  if (/<\/body\s*>/i.test(html)) return html.replace(/<\/body\s*>/i, section + '</body>');
  return html + section;
}

function MD16_renderOpportunitySection_(items, pageUrl) {
  items = Array.isArray(items) ? items : [];
  if (!items.length) return '';
  var page = String(pageUrl || '');
  var html = '<table role="presentation" border="0" cellpadding="0" cellspacing="0" width="100%" style="margin-top:24px;border-top:1px solid #eee9dc"><tr><td style="padding:22px 24px 8px">' +
    '<p style="margin:0 0 7px;color:#a27b1e;font-size:12px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase">Opportunities</p>';
  if (page) html += '<p style="margin:0 0 14px"><a href="' + MD16_escapeHtml_(page) + '" style="color:#123f3b;font-size:13px;font-weight:700">Browse all opportunities for students</a></p>';
  items.forEach(function (item) {
    html += '<div style="margin:0 0 17px;padding:14px;border:1px solid #ded8c7;border-top:3px solid #c59b35;background:#fffefb">' +
      '<h2 style="margin:0 0 7px;color:#123f3b;font-family:Georgia,serif;font-size:21px;line-height:1.25">' + MD16_escapeHtml_(item.title) + '</h2>';
    if (item.subtitle) html += '<p style="margin:0 0 5px;color:#214b45;font-size:13px;font-weight:700">' + MD16_escapeHtml_(item.subtitle) + '</p>';
    if (item.meta) html += '<p style="margin:0 0 8px;color:#73817d;font-size:12px;line-height:1.5">' + MD16_escapeHtml_(item.meta) + '</p>';
    if (item.body) html += '<p style="margin:0;color:#263d39;font-size:14px;line-height:1.6">' + MD16_escapeHtml_(item.body) + '</p>';
    if (MD16_isHttpUrl_(item.url)) html += '<p style="margin:10px 0 0"><a href="' + MD16_escapeHtml_(item.url) + '" style="color:#123f3b;font-weight:700">View official opportunity</a></p>';
    html += '</div>';
  });
  return html + '</td></tr></table>';
}

function MD16_appendOpportunityText_(text, items, pageUrl) {
  text = String(text || '');
  items = Array.isArray(items) ? items : [];
  var marker = '{{MD16_OPPORTUNITIES}}';
  var block = '';
  if (items.length) {
    block = '\n\nOPPORTUNITIES\n';
    items.forEach(function (item) {
      block += '\n' + item.title + (item.subtitle ? ' — ' + item.subtitle : '') + '\n' + item.meta +
        (item.body ? '\n' + item.body : '') + (item.url ? '\nApply: ' + item.url : '') + '\n';
    });
    if (pageUrl) block += '\nBrowse all opportunities: ' + pageUrl;
  }
  if (text.indexOf(marker) >= 0) return text.split(marker).join(block);
  return block ? text + block : text;
}

function MD16_hasEmailTemplate_() {
  try { HtmlService.createTemplateFromFile('MD_Launch_Edition_Email'); return true; }
  catch (e) { return false; }
}

function MD16_htmlToText_(html) {
  return String(html || '').replace(/<\s*br\s*\/?>/gi, '\n').replace(/<\/p\s*>/gi, '\n\n').replace(/<[^>]*>/g, ' ').replace(/&nbsp;/gi, ' ').replace(/&amp;/gi, '&').replace(/&lt;/gi, '<').replace(/&gt;/gi, '>').replace(/&quot;/gi, '"').replace(/&#39;/gi, "'").replace(/[ \t]+/g, ' ').replace(/\n\s*\n\s*\n/g, '\n\n').trim();
}

function MD16_escapeHtml_(value) {
  return String(value || '').replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/'/g, '&#39;');
}

function MD16_isEmail_(value) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || '').trim()); }
function MD16_isHttpUrl_(value) { return /^https?:\/\/[^\s]+$/i.test(String(value || '').trim()); }
function MD16_error_(error) { return String(error && error.message ? error.message : error || 'Unknown error').substring(0, 800); }

function MD16_getSpreadsheet_() {
  var id = PropertiesService.getScriptProperties().getProperty('CMS_ID');
  if (!id) throw new Error('CMS_ID is missing. Do not create a new workbook manually; run the existing setupWorkspace() function and confirm it creates CMS_ID.');
  return SpreadsheetApp.openById(id);
}

function MD16_ensureSheet_(ss, name, requiredHeaders) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0 || sh.getLastColumn() === 0 || !String(sh.getRange(1, 1).getDisplayValue() || '').trim()) {
    sh.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
  } else {
    var existing = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(function (x) { return String(x || '').trim().toLowerCase(); });
    var missing = requiredHeaders.filter(function (h) { return existing.indexOf(h.toLowerCase()) < 0; });
    if (missing.length) throw new Error('Tab ' + name + ' already exists but is missing required header(s): ' + missing.join(', ') + '. No existing data was changed.');
  }
  return sh;
}

function MD16_headerMap_(sheet, headers) {
  var current = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(function (x) { return String(x || '').trim().toLowerCase(); });
  var result = {};
  headers.forEach(function (h) {
    var n = current.indexOf(h.toLowerCase());
    if (n < 0) throw new Error('Tab ' + sheet.getName() + ' is missing required header: ' + h);
    result[h] = n;
  });
  return result;
}

function MD16_applyCampaignFormatting_(sheet) {
  var n = MD16_CONFIG_.CAMPAIGN_HEADERS.length;
  sheet.setFrozenRows(1); sheet.setTabColor(MD16_CONFIG_.BRAND_GOLD);
  sheet.getRange(1, 1, 1, n).setBackground(MD16_CONFIG_.BRAND_GREEN).setFontColor('#FFFFFF').setFontWeight('bold').setWrap(true);
  sheet.setColumnWidths(1, n, 150); sheet.setColumnWidth(2, 240); sheet.setColumnWidth(3, 280);
  sheet.setColumnWidth(4, 420); sheet.setColumnWidth(5, 360); sheet.setColumnWidth(11, 300);
  var rows = Math.max(sheet.getMaxRows() - 1, 1);
  var approval = SpreadsheetApp.newDataValidation().requireValueInList(['Draft', 'In Review', 'Approved', 'Rejected'], true).setAllowInvalid(false).build();
  var queueStatus = SpreadsheetApp.newDataValidation().requireValueInList(['Not Prepared', 'Prepared', 'No Eligible Subscribers', 'Complete'], true).setAllowInvalid(false).build();
  sheet.getRange(2, 6, rows, 1).setDataValidation(approval);
  sheet.getRange(2, 7, rows, 1).setDataValidation(queueStatus);
  sheet.getRange(2, 4, rows, 2).setWrap(true).setVerticalAlignment('top');
}

function MD16_applyQueueFormatting_(sheet) {
  var n = MD16_CONFIG_.QUEUE_HEADERS.length;
  sheet.setFrozenRows(1); sheet.setTabColor(MD16_CONFIG_.BRAND_GREEN);
  sheet.getRange(1, 1, 1, n).setBackground(MD16_CONFIG_.BRAND_GREEN).setFontColor('#FFFFFF').setFontWeight('bold').setWrap(true);
  sheet.setColumnWidths(1, n, 150); sheet.setColumnWidth(3, 250); sheet.setColumnWidth(8, 420); sheet.setColumnWidth(9, 360); sheet.setColumnWidth(12, 320);
  var rows = Math.max(sheet.getMaxRows() - 1, 1);
  var status = SpreadsheetApp.newDataValidation().requireValueInList(['Pending', 'Sending', 'Sent', 'Failed', 'Needs Review', 'Cancelled'], true).setAllowInvalid(false).build();
  sheet.getRange(2, 10, rows, 1).setDataValidation(status);
  sheet.getRange(2, 8, rows, 2).setWrap(true).setVerticalAlignment('top');
}


/**
 * Return a Failed queue row to Pending after an administrator has reviewed it.
 * This does not send the email; the normal batch action sends it later.
 */
function MD16_retryFailedQueueRow(rowNumber, reviewNote, actorEmail) {
  return MD16_resolveQueueRow_(rowNumber, 'Failed', 'Pending', reviewNote, actorEmail);
}

/**
 * Resolve a Needs Review row without guessing whether delivery occurred.
 * decision must be either Sent or Pending.
 */
function MD16_resolveNeedsReviewQueueRow(rowNumber, decision, reviewNote, actorEmail) {
  decision = String(decision || '').trim();
  if (decision !== 'Sent' && decision !== 'Pending') {
    throw new Error('Needs Review can only be resolved as Sent or Pending.');
  }
  return MD16_resolveQueueRow_(rowNumber, 'Needs Review', decision, reviewNote, actorEmail);
}

function MD16_resolveQueueRow_(rowNumber, expectedStatus, nextStatus, reviewNote, actorEmail) {
  var ss = MD16_getSpreadsheet_();
  var queue = ss.getSheetByName(MD16_CONFIG_.QUEUE_SHEET);
  if (!queue) throw new Error('NewsletterQueue tab is missing.');
  var q = MD16_headerMap_(queue, MD16_CONFIG_.QUEUE_HEADERS);
  var n = Number(rowNumber);
  if (!Number.isInteger(n) || n < 2 || n > queue.getLastRow()) throw new Error('Enter a valid NewsletterQueue row number.');
  var current = String(queue.getRange(n, q.Status + 1).getDisplayValue() || '').trim();
  if (current !== expectedStatus) throw new Error('This queue row is no longer ' + expectedStatus + '. Refresh the queue and review its current status.');
  var note = String(reviewNote || '').replace(/[\x00-\x1F\x7F]/g, ' ').trim().substring(0, 2000);
  if (!note) throw new Error('Enter a review note before resolving this queue row.');
  queue.getRange(n, q.Status + 1).setValue(nextStatus);
  queue.getRange(n, q.Error + 1).setValue(note);
  if (nextStatus === 'Pending') queue.getRange(n, q.Sent_At + 1).clearContent();
  if (nextStatus === 'Sent' && !queue.getRange(n, q.Sent_At + 1).getValue()) queue.getRange(n, q.Sent_At + 1).setValue(new Date());
  SpreadsheetApp.flush();
  if (typeof MD19_audit_ === 'function') MD19_audit_(actorEmail || 'ADMIN', 'RESOLVE_NEWSLETTER_QUEUE_ROW', 'NewsletterQueue', String(queue.getRange(n, q.Queue_ID + 1).getDisplayValue() || n), { from: expectedStatus, to: nextStatus, note: note }, 'Success');
  return { ok: true, row: n, status: nextStatus, message: nextStatus === 'Pending' ? 'Row returned to Pending. It will be eligible for a later send batch.' : 'Row marked Sent. It will not be sent again.' };
}

/** Read-only queue view for the Admin Portal. Email and content are returned for authenticated admins only. */
function MD16_listQueueForPortal(statusFilter) {
  var ss = MD16_getSpreadsheet_();
  var queue = ss.getSheetByName(MD16_CONFIG_.QUEUE_SHEET);
  if (!queue || queue.getLastRow() < 2) return { ok: true, items: [] };
  var q = MD16_headerMap_(queue, MD16_CONFIG_.QUEUE_HEADERS);
  var filter = String(statusFilter || 'all').trim().toLowerCase();
  var rows = queue.getRange(2, 1, queue.getLastRow() - 1, queue.getLastColumn()).getDisplayValues();
  var items = [];
  rows.forEach(function (row, i) {
    var status = String(row[q.Status] || '').trim();
    if (filter !== 'all' && status.toLowerCase() !== filter) return;
    items.push({ rowNumber: i + 2, queueId: row[q.Queue_ID], campaignId: row[q.Campaign_ID], email: row[q.Email], subject: row[q.Subject], status: status, error: row[q.Error], deliveryMethod: row[q.Delivery_Method], createdAt: row[q.Created_At], sentAt: row[q.Sent_At], html: row[q.HTML_Body], text: row[q.Text_Body] });
  });
  return { ok: true, items: items };
}