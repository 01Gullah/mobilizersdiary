/**
 * The Mobilizer's Diary — sponsored banner data helper
 *
 * This file deliberately does not define doPost(e). Add the action branch shown
 * in INTEGRATION.md to the EXISTING doPost dispatcher to avoid duplicate entry
 * points. The browser receives only the selected public banner fields.
 *
 * Required Script Properties:
 *   SPONSORED_BANNER_SPREADSHEET_ID
 *   SPONSORED_BANNER_SHEET_NAME       (e.g. SponsoredBanners)
 *
 * Required row-1 headers:
 *   Active | Image URL | Target URL | Alt Text | Sponsor Name
 */

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
  var headers = values[0].map(function (v) { return String(v).trim().toLowerCase(); });
  function col(name) {
    var index = headers.indexOf(name.toLowerCase());
    if (index < 0) throw new Error('Sponsored banner sheet is missing header: ' + name);
    return index;
  }
  var c = {
    active: col('Active'),
    image: col('Image URL'),
    target: col('Target URL'),
    alt: col('Alt Text'),
    sponsor: col('Sponsor Name')
  };

  // First active valid row wins. Keep ordering in the sheet under editorial control.
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var active = String(row[c.active] || '').trim().toLowerCase();
    if (['true', 'yes', 'active', '1'].indexOf(active) < 0) continue;

    var imageUrl = String(row[c.image] || '').trim();
    var targetUrl = String(row[c.target] || '').trim();
    if (!isPublicHttpsUrl_(imageUrl) || !isPublicHttpsUrl_(targetUrl)) {
      Logger.log('Skipping sponsored banner row ' + (i + 1) + ': image/target URL must use HTTPS.');
      continue;
    }

    return {
      imageUrl: imageUrl,
      targetUrl: targetUrl,
      altText: String(row[c.alt] || '').trim().substring(0, 250) || 'Sponsored message',
      sponsorName: String(row[c.sponsor] || '').trim().substring(0, 120)
    };
  }
  return null;
}

function isPublicHttpsUrl_(value) {
  return /^https:\/\/[A-Za-z0-9.-]+(?::\d{1,5})?(?:[/?#][^\s]*)?$/i.test(String(value || '').trim());
}
