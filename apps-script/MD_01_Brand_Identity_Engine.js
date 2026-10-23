const BRAND = {
  BANNER_URL: 'https://drive.google.com/file/d/1dmQOK8zeDM2A4m-mHLMaftktWDlU0VhO/view?usp=drivesdk',
  GREEN: '#123F3B', GOLD: '#C59B35', IVORY: '#F7F4EE', EDITORIAL: '#163D39'
};

var BANNER_CACHE_ = null;

function bannerId_() {
  const u = BRAND.BANNER_URL;
  const m = u.match(/\/d\/([-\w]+)/) || u.match(/[?&]id=([-\w]+)/);
  return m ? m[1] : u.trim();
}

// Always returns a banner small enough for Sheets, Docs, Forms and email
function getBannerBlob_() {
  if (BANNER_CACHE_) return BANNER_CACHE_;
  const id = bannerId_();
  if (!id) return null;
  try {
    const res = UrlFetchApp.fetch('https://drive.google.com/thumbnail?id=' + id + '&sz=w1600', {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    const b = res.getBlob();
    if (res.getResponseCode() === 200 && b.getContentType().indexOf('image') === 0) {
      BANNER_CACHE_ = b.setName('md_header.png');
      return BANNER_CACHE_;
    }
  } catch (e) { Logger.log('Resize failed: ' + e); }
  try {
    BANNER_CACHE_ = DriveApp.getFileById(id).getBlob().setName('md_header.png');
    return BANNER_CACHE_;
  } catch (e) { Logger.log('Drive lookup failed: ' + e); return null; }
}

function brandInline_() {
  const b = getBannerBlob_();
  return b ? { mdHeader: b } : {};
}

function brandHeaderDataUri_() {
  const b = getBannerBlob_();
  return b ? 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()) : '';
}

function brandEmailWrap_(innerHtml) {
  const head = getBannerBlob_()
    ? '<img src="cid:mdHeader" alt="The Mobilizer\'s Diary" style="display:block;width:100%;max-width:600px;height:auto;border:0">'
    : '<div style="background:' + BRAND.GREEN + ';padding:24px;border:2px solid ' + BRAND.GOLD +
      ';color:#fff;font:800 24px Arial,sans-serif;letter-spacing:2px">THE MOBILIZER\'S <span style="color:' +
      BRAND.GOLD + '">DIARY</span></div>';
  return '<div style="background:' + BRAND.IVORY + ';max-width:600px;margin:auto;font-family:Georgia,serif;color:' +
    BRAND.EDITORIAL + '">' + head +
    '<div style="padding:24px;font-size:16px;line-height:1.6">' + innerHtml + '</div>' +
    '<div style="border-top:3px solid ' + BRAND.GOLD + ';padding:12px 24px;font:12px Arial,sans-serif;color:#607D8B">' +
    'The Mobilizer\'s Diary. Amplifying student voices, preserving institutional memory, promoting accountability.</div></div>';
}

function applyDocHeader_(doc) {
  const blob = getBannerBlob_();
  const header = doc.getHeader() || doc.addHeader();
  header.clear();
  const p = header.getNumChildren() ? header.getChild(0).asParagraph() : header.appendParagraph('');
  p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  if (blob) {
    const body = doc.getBody();
    const w = body.getPageWidth() - body.getMarginLeft() - body.getMarginRight();
    const img = p.appendInlineImage(blob);
    const r = img.getHeight() / img.getWidth();
    img.setWidth(w).setHeight(Math.round(w * r));
  } else {
    p.appendText("THE MOBILIZER'S DIARY").setBold(true).setForegroundColor(BRAND.GREEN);
  }
}

function addFormBanner_(form) {
  const blob = getBannerBlob_();
  if (!blob) return;
  form.getItems(FormApp.ItemType.IMAGE).forEach(function (it) {
    if (it.getTitle() === "The Mobilizer's Diary") form.deleteItem(it);
  });
  form.addImageItem().setImage(blob).setTitle("The Mobilizer's Diary")
    .setAlignment(FormApp.Alignment.CENTER).setWidth(600);
  form.moveItem(form.getItems().length - 1, 0);
}


function addSheetBanner_(ss) {
  const blob = getBannerBlob_();
  if (!blob) return;
  const sh = ss.getSheetByName('README') || ss.insertSheet('README', 0);
  sh.setHiddenGridlines(true);
  sh.getImages().forEach(function (i) { i.remove(); });

  const img = sh.insertImage(blob, 1, 1, 0, 0);
  const w = 800;
  const h = Math.round(w * img.getHeight() / img.getWidth());
  img.setWidth(w).setHeight(h);

  sh.setColumnWidth(1, w + 4);
  sh.setRowHeight(1, h + 4);
  sh.getRange('A1').setBackground(BRAND.GREEN);
  sh.setRowHeight(2, 20);
  sh.getRange('A3').setValue('MD_CMS. Do not rename tabs or columns. The scripts depend on them.')
    .setWrap(true).setFontColor(BRAND.EDITORIAL);
}

function applyBrandToCms() {
  const id = PropertiesService.getScriptProperties().getProperty('CMS_ID');
  addSheetBanner_(SpreadsheetApp.openById(id));
}

function testBrand() {
  const doc = DocumentApp.create('MD_Brand_Test');
  applyDocHeader_(doc);
  doc.getBody().appendParagraph('Header test.');
  doc.saveAndClose();
  Logger.log(doc.getUrl());
}

function checkBanner() {
  const b = getBannerBlob_();
  Logger.log(b ? 'OK: ' + b.getContentType() + ', ' + b.getBytes().length + ' bytes' : 'NOT FOUND');
}

