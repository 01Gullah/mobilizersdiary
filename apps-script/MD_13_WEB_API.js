/*
 * THE MOBILIZER'S DIARY — WEB APP API DISPATCHER
 *
 * Complete replacement for MD_13_WEB_API.gs.
 * Preserves public API actions and forwards all admin actions to MD19_route_.
 * Print access actions (print*, adminPrint*) are forwarded to MD_25.
 */
function doPost(e) {
  var out;
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    out = {
      ok: true,
      result: apiRoute_(String(body.action || ''), body.data || {})
    };
  } catch (err) {
    out = {
      ok: false,
      error: String(err && err.message ? err.message : err).slice(0, 500)
    };
  }
  return ContentService
    .createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function apiRoute_(action, d) {
  d = d || {};
  switch (action) {
    case 'init':
      return {
        banner: brandHeaderDataUri_(),
        flyer: portalImage_(PORTAL.FLYER_URL, 900)
      };
    case 'hub':
      return hubGet();
    case 'report':
      return MD19_afterPublicIntake_('Report', function () { return submitReport(d); });
    case 'support':
      return MD19_afterPublicIntake_('Support request', function () { return submitSupport(d); });
    case 'donate':
      return MD19_afterPublicIntake_('Donation request', function () { return donStart(d); });
    case 'verify':
      return apiVerify_(d.ref);
    case 'subscribe':
      return MD19_afterPublicIntake_('Newsletter signup', function () { return subscribeNewsletter(d); });
    case 'unsubscribe':
      return unsubscribeNewsletter(d.token);
    case 'sponsoredBanner':
      return getSponsoredBannerForWeb_();
    case 'sponsoredBanners':
      return getSponsoredBannersForWeb_();
    case 'recommendOpportunity':
      return MD19_afterPublicIntake_('Opportunity recommendation', function () { return MD17_recommendOpportunity_(d); });
    case 'publicOpportunities':
      // Opportunity Studio: returns approved listings plus templateId and summary.
      return MD23_getPublicCards_();
    case 'volunteerApplication':
      return MD19_afterPublicIntake_('Volunteer application', function () { return MD18_submitVolunteerApplication_(d); });
    default:
      // iBuild print access (MD_25): admin price/order actions first, then public ones.
      if (String(action).indexOf('adminPrint') === 0) return MD25_adminRoute_(action, d);
      if (String(action).indexOf('print') === 0) return MD25_route_(action, d);
      // This forwards adminRequestCode, adminVerifyCode, adminDashboard,
      // newsletter actions, queue actions, and future admin actions.
      if (String(action).indexOf('admin') === 0) return MD19_route_(action, d);
      throw new Error('Unknown request.');
  }
}

function apiVerify_(ref) {
  ref = String(ref || '').trim();
  if (!/^DON-\d{4}-\d{4}$/.test(ref)) return { status: 'unknown' };
  try {
    return donVerify_(ref);
  } catch (e) {
    Logger.log(e);
    return { status: 'pending', ref: ref };
  }
}

function getSponsoredBannerForWeb_() {
  var props = PropertiesService.getScriptProperties();
  var spreadsheetId = props.getProperty('SPONSORED_BANNER_SPREADSHEET_ID');
  var sheetName = props.getProperty('SPONSORED_BANNER_SHEET_NAME');
  if (!spreadsheetId || !sheetName) {
    throw new Error('Sponsored banner spreadsheet settings are not configured.');
  }
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sponsored banner sheet not found: ' + sheetName);
  if (sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return null;
  var values = sheet.getDataRange().getDisplayValues();
  var headers = values[0].map(function (value) { return String(value).trim().toLowerCase(); });
  function columnIndex(name) {
    var index = headers.indexOf(name.toLowerCase());
    if (index < 0) throw new Error('Sponsored banner sheet is missing header: ' + name);
    return index;
  }
  var columns = {
    active: columnIndex('Active'), image: columnIndex('Image URL'), target: columnIndex('Target URL'),
    alt: columnIndex('Alt Text'), sponsor: columnIndex('Sponsor Name')
  };
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var active = String(row[columns.active] || '').trim().toLowerCase();
    if (['true', 'yes', 'active', '1'].indexOf(active) < 0) continue;
    var imageUrl = String(row[columns.image] || '').trim();
    var targetUrl = String(row[columns.target] || '').trim();
    if (!isPublicHttpsUrl_(imageUrl) || !isPublicHttpsUrl_(targetUrl)) continue;
    return {
      imageUrl: imageUrl,
      targetUrl: targetUrl,
      altText: String(row[columns.alt] || '').trim().substring(0, 250) || 'Sponsored message',
      sponsorName: String(row[columns.sponsor] || '').trim().substring(0, 120)
    };
  }
  return null;
}

function getSponsoredBannersForWeb_() {
  var props = PropertiesService.getScriptProperties();
  var spreadsheetId = props.getProperty('SPONSORED_BANNER_SPREADSHEET_ID');
  var sheetName = props.getProperty('SPONSORED_BANNER_SHEET_NAME');
  if (!spreadsheetId || !sheetName) throw new Error('Sponsored banner spreadsheet settings are not configured.');
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  if (!sheet) throw new Error('Sponsored banner sheet not found: ' + sheetName);
  if (sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return [];
  var values = sheet.getDataRange().getDisplayValues();
  var headers = values[0].map(function (value) { return String(value).trim().toLowerCase(); });
  function columnIndex(name) {
    var index = headers.indexOf(name.toLowerCase());
    if (index < 0) throw new Error('Sponsored banner sheet is missing header: ' + name);
    return index;
  }
  var columns = {
    active: columnIndex('Active'), image: columnIndex('Image URL'), target: columnIndex('Target URL'),
    alt: columnIndex('Alt Text'), sponsor: columnIndex('Sponsor Name')
  };
  var output = [];
  for (var i = 1; i < values.length && output.length < 20; i++) {
    var row = values[i];
    var active = String(row[columns.active] || '').trim().toLowerCase();
    if (['true', 'yes', 'active', '1'].indexOf(active) < 0) continue;
    var imageUrl = String(row[columns.image] || '').trim();
    var targetUrl = String(row[columns.target] || '').trim();
    if (!isPublicHttpsUrl_(imageUrl) || !isPublicHttpsUrl_(targetUrl)) continue;
    output.push({
      imageUrl: imageUrl,
      targetUrl: targetUrl,
      altText: String(row[columns.alt] || '').trim().substring(0, 250) || 'Sponsored message',
      sponsorName: String(row[columns.sponsor] || '').trim().substring(0, 120)
    });
  }
  return output;
}

function isPublicHttpsUrl_(value) {
  return /^https:\/\/[A-Za-z0-9.-]+(?::\d{1,5})?(?:[/?#][^\s]*)?$/i.test(String(value || '').trim());
}

function testApi() {
  var r = JSON.parse(doPost({
    postData: { contents: JSON.stringify({ action: 'hub', data: {} }) }
  }).getContent());
  Logger.log(r.ok ? 'API OK: hub returned ' + Object.keys(r.result).join(', ') : 'API ERROR: ' + r.error);
  var bad = JSON.parse(doPost({
    postData: { contents: JSON.stringify({ action: 'nope' }) }
  }).getContent());
  Logger.log(bad.ok ? 'FIX: unknown actions should be refused' : 'Unknown actions refused: OK');
}

function testSponsoredBanner() {
  var banner = getSponsoredBannerForWeb_();
  if (banner === null) Logger.log('Connection OK. No active banner row yet.');
  else Logger.log('Connection OK. Banner data returned: ' + JSON.stringify(banner));
}
