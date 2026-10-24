const PORTAL = {
  TITLE: 'Mobilizers Diary Student Reporting and Feedback Portal',
  FLYER_URL: 'https://drive.google.com/file/d/1ABdzVT8jNRJHtkLXAWHcCqB64hoLTCIf/view?usp=drivesdk',
  YEAR: 2027,
  NOMINATION: 'Campus Spotlight Nomination',
  ALERTS: true   // true = email you whenever a report, interview request or enquiry arrives
};

const HUB = {
  ENQUIRY: ['Advertise', 'Sponsor or Partner', 'Donor Update Request', 'General'],
  PLATFORM: ['Google Meet', 'Zoom', 'WhatsApp call', 'Phone call']
};

function doGet(e) {
  if (e && e.parameter && e.parameter.issue) return newsletterWebPage_(e.parameter.issue);
  const t = HtmlService.createTemplateFromFile('MD_02_Portal');
  t.flyer = portalImage_(PORTAL.FLYER_URL, 900);
  t.banner = brandHeaderDataUri_();
  return t.evaluate().setTitle(PORTAL.TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- helpers ---------- */
function portalSs_() { return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('CMS_ID')); }
function portalClean_(v, max) { return String(v == null ? '' : v).replace(/^\s+|\s+$/g, '').slice(0, max); }
function portalSafe_(s) { return /^[=+\-@]/.test(s) ? "'" + s : s; }
function portalEmailOk_(e) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e); }
function portalUrl_(u) { u = String(u || '').trim(); return /^https?:\/\/\S+$/i.test(u) ? u : ''; }
function portalEsc_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function portalNext_(prop) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const p = PropertiesService.getScriptProperties();
    const n = Number(p.getProperty(prop) || 0) + 1;
    p.setProperty(prop, String(n));
    return n;
  } finally { lock.releaseLock(); }
}

function portalNotify_(subject, body) {
  if (!PORTAL.ALERTS) return;
  try {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(), '[MD] ' + subject, body, { name: "The Mobilizer's Diary" });
  } catch (e) { Logger.log('Notify failed: ' + e); }
}

function portalImage_(url, width) {
  const m = String(url).match(/\/d\/([-\w]+)/) || String(url).match(/[?&]id=([-\w]+)/);
  if (!m) return '';
  try {
    const res = UrlFetchApp.fetch('https://drive.google.com/thumbnail?id=' + m[1] + '&sz=w' + width, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
    });
    if (res.getResponseCode() === 200) {
      const b = res.getBlob();
      return 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
    }
  } catch (e) { Logger.log('Image failed: ' + e); }
  return '';
}

function portalSendAck_(to, named, name, ref) {
  const html = named
    ? 'Dear ' + portalEsc_(name) + ',<br><br>Thank you for contacting Mobilizers Diary.<br><br><b>Reference Number: ' + ref +
      '</b><br><br>Your report has been received and will undergo review.<br><br>Thank you.'
    : 'Thank you for your submission.<br><br>Anonymous reports require verification before publication or action.<br><br><b>Reference Number: ' + ref + '</b>';
  MailApp.sendEmail({
    to: to,
    subject: 'Mobilizers Diary: Submission Received (' + ref + ')',
    body: html.replace(/<br>/g, '\n').replace(/<[^>]+>/g, ''),
    htmlBody: brandEmailWrap_(html),
    inlineImages: brandInline_(),
    name: "The Mobilizer's Diary"
  });
}

/* ---------- reports (with optional interview) ---------- */
function submitReport(d) {
  if (!d || d.hp) return { ok: true, ref: '', emailed: false, interview: false };
  const named = d.type === 'Named submission';
  const anon = d.type === 'Anonymous submission';
  if (!named && !anon) throw new Error('Please choose a submission type.');
  if (!d.consent || !d.declaration || (anon && !d.ack)) throw new Error('Consent and declaration are required.');

  const f = {};
  ['name', 'email', 'phone', 'role', 'publish', 'institution', 'category', 'title', 'evidence']
    .forEach(function (k) { f[k] = portalClean_(d[k], 300); });
  ['challenge', 'why', 'solutions'].forEach(function (k) { f[k] = portalClean_(d[k], 5000); });

  if (!f.category || !f.title || !f.challenge || !f.why || !f.solutions) throw new Error('Please complete all required fields.');
  if (named && (!f.name || !portalEmailOk_(f.email))) throw new Error('Please enter your name and a valid email address.');
  if (f.email && !portalEmailOk_(f.email)) throw new Error('Please enter a valid email address.');
  const nomination = f.category === PORTAL.NOMINATION;
  if (nomination && !d.nomConsent) throw new Error('Please confirm the nominee has agreed to be featured.');

  const iv = (d.iv && d.iv.consent) ? d.iv : null;
  let ivContact = '', ivPlat = 'Google Meet';
  if (iv) {
    ivContact = portalClean_(iv.contact, 200);
    if (!ivContact) throw new Error('Please add a contact for the interview, or untick the interview box.');
    ivPlat = HUB.PLATFORM.indexOf(iv.platform) >= 0 ? iv.platform : 'Google Meet';
  }

  const ss = portalSs_();
  const log = ss.getSheetByName('Submission_Log');
  const sv = ss.getSheetByName('Student_Voices');
  const ivs = iv ? ss.getSheetByName('Interview_Requests') : null;
  if (!log || !sv) throw new Error('CMS tabs not found.');
  if (iv && !ivs) throw new Error('Interview tab missing. Run setupCmsValidation first.');

  const num = portalNext_('REF_COUNTER');
  const ref = 'MD-' + PORTAL.YEAR + '-' + Utilities.formatString('%04d', num);
  const now = new Date();

  let emailed = 'No email provided';
  if (f.email) {
    try { portalSendAck_(f.email, named, f.name, ref); emailed = 'Yes'; }
    catch (e) { emailed = 'Failed: ' + e.message; Logger.log(emailed); }
  }

  const display = named ? (f.publish.indexOf('Yes') === 0 ? f.name : 'Name withheld') : 'Anonymous';
  const notes = named ? 'Role: ' + f.role + ' | Publish name: ' + f.publish + ' | Phone: ' + f.phone : '';
  const editorNote = nomination ? 'Spotlight nomination. Submitter confirmed nominee consent. Copy to Campus_Spotlight after verification.' : '';
  const S = portalSafe_;

  sv.appendRow([
    'SV-' + Utilities.formatString('%04d', num), ref, now, named ? 'Named' : 'Anonymous', S(display),
    S(f.institution), S(f.category), S(f.challenge), S(f.why), S(f.solutions), S(f.evidence),
    'Unverified', 'Draft', '', editorNote
  ]);
  log.appendRow([
    ref, now, d.type, named ? S(f.name) : 'Anonymous', S(f.email), S(f.institution), S(f.category), S(f.title),
    'Yes', anon ? 'Yes' : 'N/A', 'Unverified', '', emailed, S(notes)
  ]);
  if (iv) {
    ivs.appendRow([ref, now, named ? S(f.name) : 'Anonymous', S(ivContact), ivPlat,
      S(portalClean_(iv.availability, 200)), 'Yes', 'New', '', '']);
  }

  portalNotify_('New report ' + ref,
    'Type: ' + d.type + '\nCategory: ' + f.category + '\nTitle: ' + f.title + '\nInstitution: ' + f.institution +
    '\nInterview requested: ' + (iv ? 'Yes' : 'No') + '\n\nOpen Student_Voices in MD_CMS to review.');
  return { ok: true, ref: ref, emailed: emailed === 'Yes', interview: !!iv };
}

/* ---------- About / Support / Advertise data ---------- */
function hubSettingsMap_(ss) {
  const sh = ss.getSheetByName('Site_Settings');
  const m = {};
  if (!sh || sh.getLastRow() < 2) return m;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) { if (r[0]) m[String(r[0])] = String(r[1]); });
  return m;
}

function hubImage_(url) {
  const m = String(url || '').match(/\/d\/([-\w]+)/) || String(url || '').match(/[?&]id=([-\w]+)/);
  if (!m) return '';
  const cache = CacheService.getScriptCache();
  const key = 'hubimg_' + m[1];
  const hit = cache.get(key);
  if (hit) return hit;
  try {
    const res = UrlFetchApp.fetch('https://drive.google.com/thumbnail?id=' + m[1] + '&sz=w320', {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
    });
    if (res.getResponseCode() === 200) {
      const b = res.getBlob();
      const uri = 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
      if (uri.length < 95000) cache.put(key, uri, 21600);
      return uri;
    }
  } catch (e) { Logger.log('Hub image failed: ' + e); }
  return '';
}

function hubGet() {
  const ss = portalSs_();
  const all = hubSettingsMap_(ss);
  const pub = {};

  Object.keys(all).forEach(function (k) {
    if (/^(ABOUT_|CONTACT_|DONATE_|AD_)/.test(k)) {
      pub[k] = all[k];
    }
  });

  pub.DONATE_LINK = portalUrl_(pub.DONATE_LINK);

  let partners = [];
  const sh = ss.getSheetByName('Partners');

  if (sh && sh.getLastRow() > 1) {
    const v = sh.getDataRange().getValues();
    const h = v[0];

    partners = v.slice(1)
      .map(function (r) {
        const o = {};
        h.forEach(function (k, i) {
          o[k] = r[i];
        });
        return o;
      })
      .filter(function (o) {
        return String(o.Status).toLowerCase() === 'active' && o.Name;
      })
      .sort(function (a, b) {
        return (Number(a.Display_Order) || 999) -
          (Number(b.Display_Order) || 999);
      })
      .slice(0, 24)
      .map(function (o) {
        return {
          name: String(o.Name),
          role: String(o.Role_Title || ''),
          category: String(o.Category || 'Partner'),
          bio: String(o.Bio || ''),
          link: portalUrl_(o.Link_URL),
          image: hubImage_(o.Image_URL)
        };
      });
  }

  const secretKey = String(
    PropertiesService.getScriptProperties()
      .getProperty('PAYSTACK_SECRET_KEY') || ''
  ).trim();

  const pay = {
    ready: !!secretKey,
    min: 100,
    max: 1000000,
    presets: [500, 1000, 2000, 5000, 10000]
  };

  return {
    s: pub,
    p: partners,
    pay: pay
  };
}

/* ---------- Advertise, sponsor, donor-update and general enquiries ---------- */
function hubAck_(type, name, ref) {
  const who = name ? 'Dear ' + portalEsc_(name) + ',<br><br>' : 'Dear Supporter,<br><br>';
  if (type === 'Donor Update Request') {
    return who + "Thank you for supporting The Mobilizer's Diary. You will receive an update on how resources are being used.<br><br><b>Reference Number: " + ref + '</b><br><br>Thank you.';
  }
  if (type === 'Advertise' || type === 'Sponsor or Partner') {
    return who + "Thank you for your interest in partnering with The Mobilizer's Diary.<br><br><b>Reference Number: " + ref + '</b><br><br>Our team will review your enquiry and respond by email.<br><br>Thank you.';
  }
  return who + "Thank you for contacting The Mobilizer's Diary.<br><br><b>Reference Number: " + ref + '</b><br><br>We will respond by email.<br><br>Thank you.';
}

function submitSupport(d) {
  if (!d || d.hp) return { ok: true, ref: '', emailed: false };
  const type = portalClean_(d.type, 40);
  if (HUB.ENQUIRY.indexOf(type) < 0) throw new Error('Please choose an enquiry type.');
  const donor = type === 'Donor Update Request';
  const name = portalClean_(d.name, 150), org = portalClean_(d.org, 150), email = portalClean_(d.email, 150), phone = portalClean_(d.phone, 40);
  const msg = portalClean_(d.message, 3000), amount = portalClean_(d.amount, 60), date = portalClean_(d.date, 60);
  if (!portalEmailOk_(email)) throw new Error('Please enter a valid email address.');
  if (!d.consent) throw new Error('Please tick the box to agree to be contacted.');
  if (!donor && (!name || !msg)) throw new Error('Please enter your name and a short message.');

  const sh = portalSs_().getSheetByName('Support_Enquiries');
  if (!sh) throw new Error('Support tab missing. Run setupCmsValidation first.');
  const ref = 'SUP-' + PORTAL.YEAR + '-' + Utilities.formatString('%04d', portalNext_('SUP_COUNTER'));
  const S = portalSafe_;
  sh.appendRow([ref, new Date(), type, S(name), S(org), email, phone, S(msg), S(amount), S(date), 'Yes', 'New', '', '']);

  let emailed = false;
  try {
    const html = hubAck_(type, name, ref);
    MailApp.sendEmail({
      to: email, subject: "The Mobilizer's Diary: Enquiry Received (" + ref + ')',
      body: html.replace(/<br>/g, '\n').replace(/<[^>]+>/g, ''),
      htmlBody: brandEmailWrap_(html), inlineImages: brandInline_(), name: "The Mobilizer's Diary"
    });
    emailed = true;
  } catch (e) { Logger.log('Ack failed: ' + e); }
  portalNotify_('New ' + type + ' ' + ref, 'Name: ' + name + '\nOrganisation: ' + org + '\nEmail: ' + email + '\n\n' + msg +
    '\n\nOpen Support_Enquiries in MD_CMS.');
  return { ok: true, ref: ref, emailed: emailed };
}

/* ---------- Donor updates: edit IMPACT_SUBJECT and IMPACT_BODY in Site_Settings first ---------- */
function hubImpact_(testOnly) {
  const ss = portalSs_();
  const s = hubSettingsMap_(ss);
  const subject = s.IMPACT_SUBJECT || 'How your support is being used';
  const body = String(s.IMPACT_BODY || '');
  if (!body || body.indexOf('Replace this') === 0) throw new Error('Write your update in Site_Settings, key IMPACT_BODY, first.');
  let to = [];
  if (testOnly) {
    to = [Session.getEffectiveUser().getEmail()];
  } else {
    const sh = ss.getSheetByName('Support_Enquiries');
    if (sh && sh.getLastRow() > 1) {
      const v = sh.getDataRange().getValues();
      const h = v[0];
      const iT = h.indexOf('Enquiry_Type'), iE = h.indexOf('Email'), iC = h.indexOf('Consent_To_Contact');
      v.slice(1).forEach(function (r) {
        const e = String(r[iE]).toLowerCase();
        if (r[iT] === 'Donor Update Request' && r[iC] === 'Yes' && portalEmailOk_(e) && to.indexOf(e) < 0) to.push(e);
      });
    }
  }
  const paras = body.split(/\n+/).filter(function (x) { return x.trim(); });
  const html = paras.map(function (p) { return '<p style="margin:0 0 12px">' + portalEsc_(p) + '</p>'; }).join('') +
    '<p style="font:12px Arial,sans-serif;color:#607D8B">You receive this because you asked for updates. Reply STOP to be removed.</p>';
  to.forEach(function (addr) {
    MailApp.sendEmail({
      to: addr, subject: (testOnly ? '[TEST] ' : '') + subject, body: paras.join('\n\n'),
      htmlBody: brandEmailWrap_(html), inlineImages: brandInline_(), name: "The Mobilizer's Diary"
    });
  });
  Logger.log((testOnly ? 'Test sent to ' : 'Update sent to ') + to.length + ' recipient(s)');
}
function sendImpactUpdateTest() { hubImpact_(true); }
function sendImpactUpdate() { hubImpact_(false); }

// RUN ONCE to approve permissions
function authorizePortal() {
  MailApp.getRemainingDailyQuota();
  portalSs_();
  CacheService.getScriptCache().put('md_test', '1', 60);
  Logger.log(portalImage_(PORTAL.FLYER_URL, 300) ? 'Flyer OK' : 'Flyer NOT FOUND');
}
