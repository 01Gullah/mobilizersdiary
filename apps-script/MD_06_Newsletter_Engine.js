const NL = {
  PORTAL_URL: '',   // optional: paste your Web app URL to show a "Submit a Report" button
  OK_STATUS: ['Approved', 'Published']
};

function nlPad_(n) { return Utilities.formatString('%03d', Number(n)); }
function nlEsc_(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function nlDate_(v) {
  if (!v) return '';
  const d = (v instanceof Date) ? v : new Date(v);
  return isNaN(d.getTime()) ? String(v) : Utilities.formatDate(d, Session.getScriptTimeZone(), 'd MMMM yyyy');
}
function nlUrl_(u) { u = String(u || '').trim(); return /^https?:\/\/\S+$/i.test(u) ? u : ''; }
function nlParas_(t) {
  return String(t || '').split(/\n+/).filter(function (x) { return x.trim(); }).map(function (x) {
    return '<p style="margin:0 0 12px">' + nlEsc_(x) + '</p>';
  }).join('');
}
function nlLink_(label, url) {
  url = nlUrl_(url);
  return url ? '<a href="' + nlEsc_(url) + '" style="color:#123F3B;font-weight:700;text-decoration:underline">' + nlEsc_(label) + '</a>' : '';
}
function nlLab_(label, text) {
  return text ? '<p style="margin:0 0 8px"><b style="color:#123F3B">' + nlEsc_(label) + '</b> ' + nlEsc_(text) + '</p>' : '';
}

function nlRows_(ss, tab) {
  const sh = ss.getSheetByName(tab);
  if (!sh || sh.getLastRow() < 2) return [];
  const v = sh.getDataRange().getValues();
  const h = v[0];
  return v.slice(1).filter(function (r) { return r.join('') !== ''; }).map(function (r) {
    const o = {}; h.forEach(function (k, i) { o[k] = r[i]; }); return o;
  });
}

function nlData_(issue) {
  const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('CMS_ID'));
  const ok = function (r, needVerified) {
    if (Number(r.Issue_Number) !== Number(issue)) return false;
    if (NL.OK_STATUS.indexOf(String(r.Editorial_Status)) < 0) return false;
    return needVerified ? String(r.Verification_Status) === 'Verified' : true;
  };
  const s = nlRows_(ss, 'Issue_Settings').filter(function (r) { return Number(r.Issue_Number) === Number(issue); })[0];
  if (!s) throw new Error('No row for Issue ' + issue + ' in Issue_Settings.');
  return {
    issue: Number(issue), s: s,
    voices: nlRows_(ss, 'Student_Voices').filter(function (r) { return ok(r, true); }),
    history: nlRows_(ss, 'History_Corner').filter(function (r) { return ok(r, true); }),
    podcast: nlRows_(ss, 'Podcast').filter(function (r) { return ok(r, false); }),
    opps: nlRows_(ss, 'Opportunities').filter(function (r) { return ok(r, true); })
  };
}

function nlHead_(kicker, title) {
  return '<div style="margin:36px 0 14px">' +
    '<div style="font:800 12px Montserrat,Arial,sans-serif;letter-spacing:.18em;color:#C59B35;text-transform:uppercase">' + nlEsc_(kicker) + '</div>' +
    '<div style="font:700 24px/1.25 \'Playfair Display\',Georgia,serif;color:#123F3B;margin-top:4px">' + nlEsc_(title) + '</div>' +
    '<div style="width:56px;height:4px;background:#C59B35;margin-top:10px"></div></div>';
}
function nlCard_(inner) {
  return '<div style="background:#ffffff;border:1px solid #ddd5c2;border-left:5px solid #C59B35;padding:16px;margin:0 0 18px">' + inner + '</div>';
}
function nlItemTitle_(t) {
  return '<div style="font:700 18px/1.3 \'Playfair Display\',Georgia,serif;color:#123F3B;margin:0 0 6px">' + nlEsc_(t) + '</div>';
}
function nlMeta_(parts) {
  const t = parts.filter(function (x) { return x; }).join('  |  ');
  return t ? '<div style="font:12px Arial,sans-serif;color:#607D8B;margin:0 0 10px">' + nlEsc_(t) + '</div>' : '';
}

function nlContent_(d) {
  const s = d.s;
  const dt = nlDate_(s.Publish_Date);
  let h = '<div style="padding:26px 0 6px">' +
    '<div style="font:800 12px Montserrat,Arial,sans-serif;letter-spacing:.18em;color:#C59B35;text-transform:uppercase">' + nlEsc_(s.Cover_Kicker) + '</div>' +
    '<div style="font:700 32px/1.2 \'Playfair Display\',Georgia,serif;color:#123F3B;margin:6px 0 8px">' + nlEsc_(s.Cover_Headline) + '</div>' +
    '<div style="font:13px Arial,sans-serif;color:#607D8B">' + nlEsc_(s.Issue_Title) + (dt ? '  |  ' + dt : '') + '</div>' +
    '<div style="width:56px;height:4px;background:#C59B35;margin-top:12px"></div></div>';

  if (s.Founder_Note) {
    h += nlHead_("From the Founder's Desk", 'A Note from the Mobilizer') + nlParas_(s.Founder_Note) +
      '<p style="margin:0;font-weight:700;color:#123F3B">' + nlEsc_(s.Editor_Name || 'Comr. (Engr.) Lukman Aminu Gullah') +
      '<br><span style="font:12px Arial,sans-serif;color:#607D8B">The Mobilizer</span></p>';
  }

  if (d.history.length) {
    h += nlHead_('History Corner', 'Lessons from Student Leadership');
    d.history.forEach(function (r) {
      h += nlCard_(nlItemTitle_(r.Title) + nlMeta_([r.Person_or_Event, r.Period, r.Institution]) +
        nlLab_('Achievements.', r.Achievements) + nlLab_('Challenges.', r.Challenges) +
        nlLab_('Lessons learned.', r.Lessons_Learned) + nlLab_('Leadership model.', r.Leadership_Model) +
        (nlLink_('Source', r.Source_URL) ? '<p style="margin:6px 0 0">' + nlLink_('Source', r.Source_URL) + '</p>' : ''));
    });
  }

  if (d.voices.length) {
    h += nlHead_('Student Voices', 'Challenges and Solutions');
    d.voices.forEach(function (r) {
      h += nlCard_(nlItemTitle_(r.Category) + nlMeta_([r.Display_Name, r.Institution]) +
        nlLab_('The challenge.', r.Challenge) + nlLab_('Why it matters.', r.Why_It_Matters) +
        (r.Suggested_Solutions ? '<div style="background:#163D39;border-left:6px solid #C59B35;color:#ffffff;padding:12px 14px;margin:10px 0 0">' +
          '<div style="font:800 11px Montserrat,Arial,sans-serif;letter-spacing:.14em;color:#E0BB62;text-transform:uppercase;margin-bottom:4px">Suggested solutions</div>' +
          nlEsc_(r.Suggested_Solutions) + '</div>' : '') +
        (nlLink_('Evidence', r.Source_URL) ? '<p style="margin:10px 0 0">' + nlLink_('Evidence', r.Source_URL) + '</p>' : ''));
    });
  }

  if (d.podcast.length) {
    h += nlHead_('Podcast', 'On the Podcast');
    d.podcast.forEach(function (r) {
      h += nlCard_(nlItemTitle_(r.Episode_Title) + nlMeta_([r.Guest ? 'Guest: ' + r.Guest : '', nlDate_(r.Release_Date), r.Duration]) +
        nlParas_(r.Summary) + (nlLink_('Listen', r.Audio_URL) ? '<p style="margin:6px 0 0">' + nlLink_('Listen', r.Audio_URL) + '</p>' : ''));
    });
  }

  if (d.opps.length) {
    h += nlHead_('Opportunities', 'Opportunities for Students');
    d.opps.forEach(function (r) {
      h += nlCard_(nlItemTitle_(r.Title) + nlMeta_([r.Organisation, r.Category, r.Deadline ? 'Deadline: ' + nlDate_(r.Deadline) : '']) +
        nlLab_('Eligibility.', r.Eligibility) + (nlLink_('Apply', r.Apply_URL) ? '<p style="margin:6px 0 0">' + nlLink_('Apply', r.Apply_URL) + '</p>' : ''));
    });
  }

  if (NL.PORTAL_URL) {
    h += '<div style="margin:30px 0 8px;text-align:center"><a href="' + nlEsc_(NL.PORTAL_URL) +
      '" style="display:inline-block;background:#123F3B;color:#ffffff;border-bottom:4px solid #C59B35;padding:14px 24px;font:800 13px Arial,sans-serif;letter-spacing:.1em;text-decoration:none;text-transform:uppercase">Submit a Report</a></div>';
  }
  h += '<p style="margin:30px 0 0;font-style:italic;color:#607D8B">Good leadership is important, but active citizenship is equally important.</p>';
  return h;
}

function nlPage_(d, inner) {
  const banner = brandHeaderDataUri_();
  const title = "The Mobilizer's Diary: " + (d.s.Issue_Title || 'Issue ' + nlPad_(d.issue));
  return '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>' + nlEsc_(title) + '</title>' +
    '<link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@800&family=Playfair+Display:wght@700&family=Source+Serif+4:wght@400;600&display=swap" rel="stylesheet"></head>' +
    '<body style="margin:0;background:#EDE8DC"><div style="max-width:680px;margin:0 auto;background:#F7F4EE;font:17px/1.65 \'Source Serif 4\',\'Source Serif Pro\',Georgia,serif;color:#163D39">' +
    (banner ? '<img src="' + banner + '" alt="The Mobilizer\'s Diary" style="display:block;width:100%;height:auto">' : '') +
    '<div style="padding:8px 24px 28px">' + inner + '</div>' +
    '<div style="border-top:3px solid #C59B35;padding:14px 24px;font:12px Arial,sans-serif;color:#607D8B">' +
    "The Mobilizer's Diary. Amplifying student voices, preserving institutional memory, promoting accountability.</div></div></body></html>";
}

function nlFolder_(issue) {
  const root = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('ROOT_ID'));
  const sub = function (parent, name) {
    const it = parent.getFoldersByName(name);
    return it.hasNext() ? it.next() : parent.createFolder(name);
  };
  const f = sub(sub(root, '01_Issues'), 'Issue_' + nlPad_(issue));
  sub(f, 'Drafts'); sub(f, 'Assets');
  return sub(f, 'Final');
}

function nlSave_(folder, name, content, mime) {
  const it = folder.getFilesByName(name);
  if (it.hasNext()) { const f = it.next(); f.setContent(content); return f; }
  return folder.createFile(name, content, mime);
}

function nlDoc_(d, folder) {
  const name = 'MD_Issue_' + nlPad_(d.issue) + '_Newsletter';
  const old = folder.getFilesByName(name);
  while (old.hasNext()) old.next().setTrashed(true);

  const doc = DocumentApp.create(name);
  applyDocHeader_(doc);
  const b = doc.getBody();
  b.clear();
  const G = '#123F3B', GOLD = '#C59B35', ED = '#163D39', GREY = '#607D8B';

  const para = function (text, o) {
    o = o || {};
    const p = b.appendParagraph(text);
    p.setHeading(DocumentApp.ParagraphHeading.NORMAL);
    p.setSpacingBefore(o.before || 0).setSpacingAfter(o.after == null ? 8 : o.after);
    p.editAsText().setFontFamily(o.font || 'Georgia').setFontSize(o.size || 11)
      .setForegroundColor(o.color || ED).setBold(!!o.bold).setItalic(!!o.italic);
    return p;
  };
  const heading = function (kicker, title) {
    para(String(kicker).toUpperCase(), { font: 'Montserrat', size: 9, color: GOLD, bold: true, before: 22, after: 2 });
    para(title, { font: 'Playfair Display', size: 18, color: G, bold: true, after: 8 });
  };
  const item = function (title, meta) {
    para(title || 'Untitled', { font: 'Playfair Display', size: 14, color: G, bold: true, before: 8, after: 2 });
    const m = meta.filter(function (x) { return x; }).join('  |  ');
    if (m) para(m, { size: 9, color: GREY, after: 6 });
  };
  const labelled = function (label, text) {
    if (!text) return;
    const p = b.appendParagraph('');
    p.setSpacingAfter(6);
    p.appendText(label + ' ').setFontFamily('Georgia').setFontSize(11).setBold(true).setForegroundColor(G);
    p.appendText(String(text)).setFontFamily('Georgia').setFontSize(11).setBold(false).setForegroundColor(ED);
  };
  const link = function (label, url) {
    url = nlUrl_(url);
    if (!url) return;
    const p = b.appendParagraph('');
    p.setSpacingAfter(10);
    p.appendText(label).setLinkUrl(url).setFontFamily('Georgia').setFontSize(11).setBold(true);
  };
  const s = d.s;

  para(String(s.Cover_Kicker || '').toUpperCase(), { font: 'Montserrat', size: 9, color: GOLD, bold: true, before: 10, after: 4 });
  para(String(s.Cover_Headline || ''), { font: 'Playfair Display', size: 28, color: G, bold: true, after: 4 });
  const dt = nlDate_(s.Publish_Date);
  para(String(s.Issue_Title || '') + (dt ? '  |  ' + dt : ''), { size: 10, color: GREY, after: 6 });

  if (s.Founder_Note) {
    heading("From the Founder's Desk", 'A Note from the Mobilizer');
    String(s.Founder_Note).split(/\n+/).forEach(function (t) { if (t.trim()) para(t); });
    para(String(s.Editor_Name || 'Comr. (Engr.) Lukman Aminu Gullah'), { bold: true, color: G, before: 6, after: 0 });
    para('The Mobilizer', { size: 9, color: GREY });
  }
  if (d.history.length) {
    heading('History Corner', 'Lessons from Student Leadership');
    d.history.forEach(function (r) {
      item(r.Title, [r.Person_or_Event, r.Period, r.Institution]);
      labelled('Achievements.', r.Achievements); labelled('Challenges.', r.Challenges);
      labelled('Lessons learned.', r.Lessons_Learned); labelled('Leadership model.', r.Leadership_Model);
      link('Source', r.Source_URL);
    });
  }
  if (d.voices.length) {
    heading('Student Voices', 'Challenges and Solutions');
    d.voices.forEach(function (r) {
      item(r.Category, [r.Display_Name, r.Institution]);
      labelled('The challenge.', r.Challenge); labelled('Why it matters.', r.Why_It_Matters);
      labelled('Suggested solutions.', r.Suggested_Solutions);
      link('Evidence', r.Source_URL);
    });
  }
  if (d.podcast.length) {
    heading('Podcast', 'On the Podcast');
    d.podcast.forEach(function (r) {
      item(r.Episode_Title, [r.Guest ? 'Guest: ' + r.Guest : '', nlDate_(r.Release_Date), r.Duration]);
      if (r.Summary) para(String(r.Summary));
      link('Listen', r.Audio_URL);
    });
  }
  if (d.opps.length) {
    heading('Opportunities', 'Opportunities for Students');
    d.opps.forEach(function (r) {
      item(r.Title, [r.Organisation, r.Category, r.Deadline ? 'Deadline: ' + nlDate_(r.Deadline) : '']);
      labelled('Eligibility.', r.Eligibility);
      link('Apply', r.Apply_URL);
    });
  }
  para('Good leadership is important, but active citizenship is equally important.', { italic: true, color: GREY, before: 24 });

  try {
    const first = b.getChild(0);
    if (b.getNumChildren() > 1 && first.getType() === DocumentApp.ElementType.PARAGRAPH && first.asParagraph().getText() === '') b.removeChild(first);
  } catch (e) {}

  doc.saveAndClose();
  const file = DriveApp.getFileById(doc.getId());
  file.moveTo(folder);
  return file.getUrl();
}

function nlStrip_(html) {
  return html.replace(/<\/(p|div|h\d)>/g, '\n').replace(/<br\s*\/?>/g, '\n').replace(/<[^>]+>/g, '')
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/\n{3,}/g, '\n\n').trim();
}

function buildIssue_(issue, sendTest) {
  const d = nlData_(issue);
  const folder = nlFolder_(issue);
  const inner = nlContent_(d);
  const pad = nlPad_(issue);

  const htmlFile = nlSave_(folder, 'MD_Issue_' + pad + '_Newsletter.html', nlPage_(d, inner), MimeType.HTML);
  const docUrl = nlDoc_(d, folder);

  let emailed = 'skipped';
  if (sendTest) {
    const me = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
    MailApp.sendEmail({
      to: me,
      subject: "[DRAFT] The Mobilizer's Diary: " + (d.s.Issue_Title || 'Issue ' + pad) + ' | ' + (d.s.Cover_Headline || ''),
      body: nlStrip_(inner),
      htmlBody: brandEmailWrap_(inner),
      inlineImages: brandInline_(),
      name: "The Mobilizer's Diary"
    });
    emailed = 'sent to ' + me;
  }
  Logger.log('Included: ' + d.history.length + ' history, ' + d.voices.length + ' voices, ' + d.podcast.length + ' podcast, ' + d.opps.length + ' opportunities');
  Logger.log('HTML: ' + htmlFile.getUrl());
  Logger.log('DOC: ' + docUrl);
  Logger.log('EMAIL: ' + emailed);
}

// RUN THIS: builds Issue 001 (HTML, Google Doc, test email)
function buildIssue001() { buildIssue_(1, true); }

// Files only, no email
function buildIssue001FilesOnly() { buildIssue_(1, false); }

// Used by the web app to show a PUBLISHED issue at YOUR_WEB_APP_URL?issue=1
function newsletterWebPage_(issue) {
  try {
    const d = nlData_(Number(issue));
    if (String(d.s.Status) !== 'Published') {
      return HtmlService.createHtmlOutput('<p style="font-family:Georgia,serif;padding:24px">This issue has not been published yet.</p>')
        .setTitle("The Mobilizer's Diary");
    }
    return HtmlService.createHtmlOutput(nlPage_(d, nlContent_(d))).setTitle("The Mobilizer's Diary")
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (e) {
    return HtmlService.createHtmlOutput('<p style="font-family:Georgia,serif;padding:24px">Issue not found.</p>');
  }
}

