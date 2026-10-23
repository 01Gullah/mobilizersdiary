/*
 * THE MOBILIZER'S DIARY — ARCHIVE + COMPOSER RELEASE
 * COMPLETE REPLACEMENT: MD_21_Newsletter_Studio.gs (not a new parallel file).
 * Keep the existing MD_16_Newsletter_Controller.gs and MD_20_Editorial_Workflow.gs.
 * Replace MD_19_Admin_Portal.gs, admin.html and MD_Launch_Edition_Email.html together.
 * Run MD21_setupNewsletterStudio() once, then redeploy the existing web app.
 * No provider, credentials, triggers, signup, or send implementation is replaced.
 * Archive = immutable snapshots + hidden original rows, NEVER deletion. Original
 * Sent queue rows stay in place to preserve MD16 campaign/email deduplication.
 */
var MD21_CONFIG_ = {
  ITEMS: 'Newsletter_Editorial_Items', ISSUES: 'Newsletter_Issues',
  ARCHIVE: 'Newsletter_Archive', REVISIONS: 'Newsletter_Item_Revisions',
  ITEM_HEADERS: ['Item_ID','Section','Title','Body','Source_URL','Image_URL','Status','Created_By','Created_At','Approved_By','Approved_At','Display_Order','Revision','Updated_By','Updated_At','Approval_ID','Archived_At','Archived_By'],
  ISSUE_HEADERS: ['Issue_ID','Issue_Title','Subject','Cover_Headline','Status','Created_By','Created_At','Approved_At','Campaign_ID','HTML_Body','Display_Order','Item_Order_JSON','Footer_JSON','Revision','Compiled_Revision','Item_Revisions_JSON','Approval_ID','Updated_By','Updated_At','Archived_At','Archived_By'],
  ARCHIVE_HEADERS: ['Archive_ID','Record_Type','Record_ID','Archived_At','Archived_By','Reason','Summary','Part','Total_Parts','Snapshot_JSON'],
  REVISION_HEADERS: ['Item_ID','Revision','Saved_At','Saved_By','Snapshot_JSON'],
  ARCHIVE_META: ['Archived_At','Archived_By'],
  SECTIONS: ['History Corner','Campus Spotlight','Student Voices','Opportunities','Podcast','Other','Founder’s Note']
};
function MD21_withLock_(fn) {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function MD21_setupNewsletterStudio() {
  return MD21_withLock_(function() {
    var ss = MD19_getSpreadsheet_();
    MD21_sheet_(ss, MD21_CONFIG_.ITEMS, MD21_CONFIG_.ITEM_HEADERS);
    MD21_sheet_(ss, MD21_CONFIG_.ISSUES, MD21_CONFIG_.ISSUE_HEADERS);
    MD21_sheet_(ss, MD21_CONFIG_.ARCHIVE, MD21_CONFIG_.ARCHIVE_HEADERS);
    MD21_sheet_(ss, MD21_CONFIG_.REVISIONS, MD21_CONFIG_.REVISION_HEADERS);
    [MD16_CONFIG_.CAMPAIGNS_SHEET, MD16_CONFIG_.QUEUE_SHEET].forEach(function(name) {
      var sh = ss.getSheetByName(name);
      if (sh) MD21_sheet_(ss, name, MD21_CONFIG_.ARCHIVE_META);
    });
    return {ok:true,message:'Archive and composer ready. Existing records and delivery history were preserved.'};
  });
}
/* Add missing columns by name, never insert into/move an existing data column. */
function MD21_sheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name); if (!sh) sh = ss.insertSheet(name);
  if (!sh.getLastRow()) sh.getRange(1,1,1,headers.length).setValues([headers]);
  else {
    var actual = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
    headers.forEach(function(h) {
      if (actual.indexOf(h) < 0) {
        var col = actual.length + 1;
        if (col > sh.getMaxColumns()) sh.insertColumnsAfter(sh.getMaxColumns(), col - sh.getMaxColumns());
        sh.getRange(1,col).setValue(h); actual.push(h);
      }
    });
  }
  sh.setFrozenRows(1); return sh;
}
function MD21_map_(sh, headers) {
  var a = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0], m = {};
  (headers || a).forEach(function(h) { var i = a.indexOf(h); if (i < 0) throw new Error('Missing column '+h+' in '+sh.getName()); m[h] = i; });
  return m;
}
function MD21_rows_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  var h = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
  return sh.getRange(2,1,sh.getLastRow()-1,h.length).getValues().map(function(a,i) {
    var o = {_row:i+2}; h.forEach(function(k,n) { if (k) o[k] = a[n]; }); return o;
  }).filter(function(o) { return Object.keys(o).some(function(k) { return k !== '_row' && String(o[k] || '').trim(); }); });
}
function MD21_append_(sh, o) {
  var h = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
  sh.appendRow(h.map(function(k) { return o[k] == null ? '' : MD21_cell_(o[k]); }));
}
function MD21_cell_(v) { return typeof v === 'string' && /^[=+@'-]/.test(v) ? "'" + v : v; }
function MD21_write_(sh, row, changes) {
  var m = MD21_map_(sh);
  Object.keys(changes).forEach(function(k) { if (m[k] == null) throw new Error('Missing column '+k); sh.getRange(row,m[k]+1).setValue(MD21_cell_(changes[k])); });
}
function MD21_find_(sh, key, id) {
  var r = MD21_rows_(sh).filter(function(o) { return String(o[key]) === String(id); })[0];
  if (!r) throw new Error('Newsletter record not found: '+id); return r;
}
function MD21_active_(r) { return !r.Archived_At; }
function MD21_requireActive_(r) { if (!MD21_active_(r)) throw new Error('Archived records are read-only. Create a new item or issue instead.'); }
function MD21_id_(prefix) { return prefix+'-'+Utilities.getUuid().replace(/-/g,'').substring(0,14).toUpperCase(); }
function MD21_json_(v, fallback) { try { return v ? JSON.parse(String(v)) : fallback; } catch(e) { throw new Error('Invalid saved composer JSON; review the source row before proceeding.'); } }
function MD21_order_(v, fallback) {
  if (v == null || String(v).trim() === '') return fallback;
  var n = Number(v); if (!isFinite(n) || n < -100000 || n > 100000) throw new Error('Display order must be a number between -100000 and 100000.'); return n;
}
function MD21_defaultOrder_(section) { var n = MD21_CONFIG_.SECTIONS.indexOf(section); return n < 0 ? 60 : (n+1)*10; }
function MD21_sort_(a, orderKey, idKey, sectionKey) {
  return a.sort(function(x,y) {
    var d = MD21_order_(x[orderKey], sectionKey ? MD21_defaultOrder_(x[sectionKey]) : 100) - MD21_order_(y[orderKey], sectionKey ? MD21_defaultOrder_(y[sectionKey]) : 100);
    if (d) return d; return String(x[idKey] || '').localeCompare(String(y[idKey] || ''));
  });
}
function MD21_url_(v, max) {
  var s = MD19_clean_(v, max || 1000);
  if (s && !/^https?:\/\/[^\s<>"']+$/i.test(s)) throw new Error('Links must be complete http:// or https:// URLs.'); return s;
}
function MD21_itemData_(d) {
  var section = MD19_clean_(d.section,100), title = MD19_clean_(d.title,240);
  var body = String(d.body == null ? '' : d.body).replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g,' ').trim().substring(0,8000);
  if (MD21_CONFIG_.SECTIONS.indexOf(section) < 0 || !title || !body) throw new Error('Choose a section and enter a title and body.');
  return {Section:section,Title:title,Body:body,Source_URL:MD21_url_(d.sourceUrl,600),Image_URL:MD21_url_(d.imageUrl,1000),Display_Order:MD21_order_(d.displayOrder,MD21_defaultOrder_(section))};
}
function MD21_createEditorialItem_(u,d) {
  var sh = MD21_sheet_(MD19_getSpreadsheet_(),MD21_CONFIG_.ITEMS,MD21_CONFIG_.ITEM_HEADERS), o = MD21_itemData_(d || {});
  o.Item_ID = MD21_id_('NLI'); o.Status = 'Draft'; o.Revision = 1; o.Created_By = u.email; o.Created_At = new Date();
  MD21_append_(sh,o); return {ok:true,itemId:o.Item_ID,message:'Editorial item created as Draft.'};
}
function MD21_getItem_(u,d) {
  var sh = MD21_sheet_(MD19_getSpreadsheet_(),MD21_CONFIG_.ITEMS,MD21_CONFIG_.ITEM_HEADERS), r = MD21_find_(sh,'Item_ID',d.itemId);
  MD21_requireActive_(r); return {ok:true,item:r};
}
function MD21_updateEditorialItem_(u,d) {
  var ss = MD19_getSpreadsheet_(), sh = MD21_sheet_(ss,MD21_CONFIG_.ITEMS,MD21_CONFIG_.ITEM_HEADERS), r = MD21_find_(sh,'Item_ID',d.itemId);
  MD21_requireActive_(r);
  if (Number(d.revision) !== Number(r.Revision || 1)) throw new Error('This item changed since it was loaded. Reload it before saving.');
  var o = MD21_itemData_(d), history = MD21_sheet_(ss,MD21_CONFIG_.REVISIONS,MD21_CONFIG_.REVISION_HEADERS);
  MD21_append_(history,{Item_ID:r.Item_ID,Revision:r.Revision || 1,Saved_At:new Date(),Saved_By:u.email,Snapshot_JSON:JSON.stringify(r)});
  MD21_supersede_(r.Item_ID,u,'Item revised');
  MD21_invalidateItemIssues_(u,r.Item_ID);
  o.Status = 'Draft'; o.Revision = Number(r.Revision || 1)+1; o.Approved_By = ''; o.Approved_At = ''; o.Approval_ID = ''; o.Updated_By = u.email; o.Updated_At = new Date();
  MD21_write_(sh,r._row,o);
  MD19_audit_(u.email,'REVISE_NEWSLETTER_ITEM','Newsletter',r.Item_ID,{revision:o.Revision},'Draft; reapproval required');
  return {ok:true,itemId:r.Item_ID,revision:o.Revision,message:'Revision saved as Draft. Request a new approval. Unsent compiled issues using this item also need recompilation and approval; delivered snapshots are unchanged.'};
}
function MD21_listItems_(u,status) {
  var sh = MD21_sheet_(MD19_getSpreadsheet_(),MD21_CONFIG_.ITEMS,MD21_CONFIG_.ITEM_HEADERS), f = String(status || 'all').toLowerCase();
  return {ok:true,items:MD21_sort_(MD21_rows_(sh).filter(function(r) { return MD21_active_(r) && (f === 'all' || String(r.Status).toLowerCase() === f); }),'Display_Order','Item_ID','Section')};
}
function MD21_supersede_(ref,u,note) {
  var sh = MD19_getSpreadsheet_().getSheetByName(MD20_CONFIG_.APPROVALS_SHEET); if (!sh) return;
  MD21_rows_(sh).forEach(function(r) {
    if (r.Item_Type === 'Newsletter' && String(r.Item_Reference) === String(ref) && (r.Status === 'Pending' || r.Status === 'Approved')) {
      MD21_write_(sh,r._row,{Status:'Superseded',Decision_Note:note || 'Reapproval required',Last_Updated_At:new Date()});
    }
  });
}
function MD21_validateApproval_(r,id,ref,action) {
  var a = MD20_find_(id);
  if (!id || String(r.Approval_ID) !== String(id) || !a || a.itemType !== 'Newsletter' || a.itemReference !== ref || a.action !== action || a.status !== 'Pending') throw new Error('Use the current Pending approval ID for this exact record. Older revisions cannot be approved.');
}
function MD21_requestItemApproval_(u,d) {
  var sh = MD21_sheet_(MD19_getSpreadsheet_(),MD21_CONFIG_.ITEMS,MD21_CONFIG_.ITEM_HEADERS), r = MD21_find_(sh,'Item_ID',d.itemId);
  MD21_requireActive_(r); if (r.Status === 'Approved') throw new Error('This revision is already Approved. Edit it to create a new Draft revision.');
  var out = MD20_requestApproval_(u,{itemType:'Newsletter',itemReference:r.Item_ID,action:'Approve newsletter item',note:'Revision '+(r.Revision || 1)+'. '+(d.note || '')});
  MD21_write_(sh,r._row,{Status:'Review Submitted',Approval_ID:out.approvalId}); return out;
}
function MD21_approveItem_(u,d) {
  var sh = MD21_sheet_(MD19_getSpreadsheet_(),MD21_CONFIG_.ITEMS,MD21_CONFIG_.ITEM_HEADERS), r = MD21_find_(sh,'Item_ID',d.itemId); MD21_requireActive_(r);
  MD21_validateApproval_(r,d.approvalId,r.Item_ID,'Approve newsletter item');
  var out = MD20_approve_(u,d);
  if (out.status === 'Approved') MD21_write_(sh,r._row,{Status:'Approved',Approved_By:u.email,Approved_At:new Date()}); return out;
}
function MD21_footer_(d) {
  d = d || {}; return {websiteUrl:MD21_url_(d.websiteUrl),shareUrl:MD21_url_(d.shareUrl),recommendUrl:MD21_url_(d.recommendUrl),volunteerUrl:MD21_url_(d.volunteerUrl),adTitle:MD19_clean_(d.adTitle,160),adText:MD19_clean_(d.adText,1200),adImageUrl:MD21_url_(d.adImageUrl),adUrl:MD21_url_(d.adUrl)};
}
function MD21_itemOrder_(a) {
  if (!Array.isArray(a)) throw new Error('Issue items must be an array.'); if (a.length > 100) throw new Error('Use at most 100 editorial items in one issue.');
  var seen = {}; return a.map(function(x) {
    var id = MD19_clean_(x.itemId,120); if (!id || seen[id]) throw new Error('Select each editorial item only once.'); seen[id] = true;
    return {itemId:id,order:MD21_order_(x.order,100)};
  });
}
function MD21_issueData_(d) {
  var title = MD19_clean_(d.issueTitle,240), subject = MD19_clean_(d.subject,240);
  if (!title || !subject) throw new Error('Enter an issue title and email subject.');
  return {Issue_Title:title,Subject:subject,Cover_Headline:MD19_clean_(d.coverHeadline,240),Display_Order:MD21_order_(d.displayOrder,100),Item_Order_JSON:JSON.stringify(MD21_itemOrder_(d.itemOrder || [])),Footer_JSON:JSON.stringify(MD21_footer_(d.footer))};
}
function MD21_createIssue_(u,d) {
  var sh = MD21_sheet_(MD19_getSpreadsheet_(),MD21_CONFIG_.ISSUES,MD21_CONFIG_.ISSUE_HEADERS), o = MD21_issueData_(d || {});
  o.Issue_ID = MD21_id_('ISSUE'); o.Status = 'Draft'; o.Created_By = u.email; o.Created_At = new Date(); o.Revision = 1;
  MD21_append_(sh,o); return {ok:true,issueId:o.Issue_ID,revision:1,message:'Issue created as Draft. Review item selection and footer, then compile.'};
}
function MD21_listIssues_(u) {
  var sh = MD21_sheet_(MD19_getSpreadsheet_(),MD21_CONFIG_.ISSUES,MD21_CONFIG_.ISSUE_HEADERS);
  var items = MD21_sort_(MD21_rows_(sh).filter(MD21_active_),'Display_Order','Issue_ID');
  return {ok:true,items:items.map(function(r) { var o = {}; Object.keys(r).forEach(function(k) { if (k !== 'HTML_Body') o[k] = r[k]; }); return o; })};
}
function MD21_campaignRows_(id) {
  return MD21_rows_(MD19_getSpreadsheet_().getSheetByName(MD16_CONFIG_.QUEUE_SHEET)).filter(function(r) { return String(r.Campaign_ID) === String(id); });
}
function MD21_frozen_(issue) {
  return !!issue.Campaign_ID && MD21_campaignRows_(issue.Campaign_ID).some(function(r) { return ['Sent','Sending','Needs Review'].indexOf(String(r.Status)) >= 0; });
}
function MD21_retireCampaign_(u,id) {
  if (!id) return;
  var ss = MD19_getSpreadsheet_(), sh = ss.getSheetByName(MD16_CONFIG_.CAMPAIGNS_SHEET); if (!sh) return;
  var c = MD21_find_(sh,'Campaign_ID',id);
  var qsh = ss.getSheetByName(MD16_CONFIG_.QUEUE_SHEET), rows = MD21_campaignRows_(id);
  if (rows.some(function(r) { return ['Sent','Sending','Needs Review'].indexOf(String(r.Status)) >= 0; })) throw new Error('This campaign has delivery activity. Its snapshot is frozen; create a new issue instead.');
  MD21_write_(sh,c._row,{Approval_Status:'Draft',Queue_Status:'Complete'});
  rows.forEach(function(r) { if (r.Status !== 'Cancelled') MD21_write_(qsh,r._row,{Status:'Cancelled',Error:'Superseded by a revised issue. '+String(r.Error || '')}); });
}
function MD21_invalidateIssue_(u,sh,r,note) {
  MD21_retireCampaign_(u,r.Campaign_ID); MD21_supersede_(r.Issue_ID,u,note);
  MD21_write_(sh,r._row,{Status:'Draft',Approved_At:'',Approval_ID:'',Compiled_Revision:'',HTML_Body:''});
}
function MD21_invalidateItemIssues_(u,id) {
  var sh = MD21_sheet_(MD19_getSpreadsheet_(),MD21_CONFIG_.ISSUES,MD21_CONFIG_.ISSUE_HEADERS);
  MD21_rows_(sh).filter(MD21_active_).forEach(function(r) {
    var versions = MD21_json_(r.Item_Revisions_JSON,{});
    if (r.Campaign_ID && (!r.Item_Revisions_JSON || versions[id]) && !MD21_frozen_(r)) MD21_invalidateIssue_(u,sh,r,'Included item revised or archived; recompile and reapprove');
  });
}
function MD21_updateIssue_(u,d) {
  var sh = MD21_sheet_(MD19_getSpreadsheet_(),MD21_CONFIG_.ISSUES,MD21_CONFIG_.ISSUE_HEADERS), r = MD21_find_(sh,'Issue_ID',d.issueId); MD21_requireActive_(r);
  if (Number(d.revision) !== Number(r.Revision || 1)) throw new Error('This issue changed since it was loaded. Reload before saving.');
  var o = MD21_issueData_(d), changed = Object.keys(o).some(function(k) { return String(o[k]) !== String(r[k] == null ? '' : r[k]); });
  if (!changed) return {ok:true,issueId:r.Issue_ID,revision:Number(r.Revision || 1),message:'Issue settings are unchanged.'};
  if (MD21_frozen_(r)) throw new Error('This issue has delivery activity. Create a new issue to change content, order or footer.');
  MD21_invalidateIssue_(u,sh,r,'Issue settings revised'); o.Revision = Number(r.Revision || 1)+1; o.Updated_By = u.email; o.Updated_At = new Date();
  MD21_write_(sh,r._row,o); return {ok:true,issueId:r.Issue_ID,revision:o.Revision,message:'Issue settings saved as Draft. Recompile and request new approval.'};
}
function MD21_selectedItems_(r) {
  var all = MD21_listItems_({},'all').items, order = MD21_json_(r.Item_Order_JSON,[]);
  if (!r.Item_Order_JSON) order = all.filter(function(i) { return i.Status === 'Approved'; }).map(function(i) { return {itemId:i.Item_ID,order:MD21_order_(i.Display_Order,MD21_defaultOrder_(i.Section))}; });
  var items = order.map(function(x) {
    var i = all.filter(function(a) { return a.Item_ID === x.itemId; })[0];
    if (!i || i.Status !== 'Approved') throw new Error('Selected item '+x.itemId+' must be active and Approved. Revise the selection or approve it first.');
    var copy = {}; Object.keys(i).forEach(function(k) { copy[k] = i[k]; }); copy.Display_Order = x.order; return copy;
  });
  if (!items.length) throw new Error('Select at least one Approved editorial item before compiling.');
  return MD21_sort_(items,'Display_Order','Item_ID','Section');
}
function MD21_render_(title,cover,items,footer) {
  var t = HtmlService.createTemplateFromFile('MD_Launch_Edition_Email');
  t.edition = {issueTitle:title,headline:cover,items:items,footer:footer || {},founderGreeting:'Welcome to The Mobilizer’s Diary — student voices, shared memory and practical opportunities.',autoOpportunitiesHtml:'{{MD16_OPPORTUNITIES}}',unsubscribeUrl:''};
  return t.evaluate().getContent();
}
function MD21_compileIssue_(u,d) {
  var ss = MD19_getSpreadsheet_(), sh = MD21_sheet_(ss,MD21_CONFIG_.ISSUES,MD21_CONFIG_.ISSUE_HEADERS), r = MD21_find_(sh,'Issue_ID',d.issueId); MD21_requireActive_(r);
  if (MD21_frozen_(r)) throw new Error('This issue has delivery activity. Compile a new issue instead.');
  var items = MD21_selectedItems_(r), html = MD21_render_(r.Issue_Title,r.Cover_Headline,items,MD21_json_(r.Footer_JSON,{}));
  if (html.length > 49000) throw new Error('This issue exceeds the spreadsheet HTML cell limit. Shorten content or use fewer items.');
  MD21_invalidateIssue_(u,sh,r,'Issue recompiled; new campaign requires approval');
  var draft = MD16_createCampaignDraft(), cs = ss.getSheetByName(MD16_CONFIG_.CAMPAIGNS_SHEET), c = MD21_find_(cs,'Campaign_ID',draft.campaignId);
  MD21_write_(cs,c._row,{Issue_Title:r.Issue_Title,Subject:r.Subject,HTML_Body:html,Text_Body:MD16_htmlToText_(html)});
  var versions = {}; items.forEach(function(i) { versions[i.Item_ID] = Number(i.Revision || 1); });
  MD21_write_(sh,r._row,{Status:'Compiled',Campaign_ID:draft.campaignId,HTML_Body:html,Compiled_Revision:Number(r.Revision || 1),Item_Revisions_JSON:JSON.stringify(versions),Item_Order_JSON:JSON.stringify(items.map(function(i) { return {itemId:i.Item_ID,order:i.Display_Order}; }))});
  return {ok:true,issueId:r.Issue_ID,campaignId:draft.campaignId,message:'Issue compiled into a new Draft campaign. Request and obtain issue approval before queue preparation.'};
}
function MD21_checkCompiled_(r) {
  if (!r.Campaign_ID || !r.HTML_Body || Number(r.Compiled_Revision) !== Number(r.Revision || 1)) throw new Error('Compile this issue before requesting or granting approval.');
  var versions = MD21_json_(r.Item_Revisions_JSON,{}), items = MD21_selectedItems_(r);
  if (!Object.keys(versions).length || items.some(function(i) { return versions[i.Item_ID] !== Number(i.Revision || 1); })) throw new Error('Included items changed. Recompile and request a new approval.');
}
function MD21_requestIssueApproval_(u,d) {
  var sh = MD21_sheet_(MD19_getSpreadsheet_(),MD21_CONFIG_.ISSUES,MD21_CONFIG_.ISSUE_HEADERS), r = MD21_find_(sh,'Issue_ID',d.issueId); MD21_requireActive_(r); MD21_checkCompiled_(r);
  if (MD21_frozen_(r) || r.Status === 'Approved') throw new Error('This issue is already in delivery or Approved. Create a new issue for changes.');
  var out = MD20_requestApproval_(u,{itemType:'Newsletter',itemReference:r.Issue_ID,action:'Approve newsletter issue',note:'Compiled revision '+r.Compiled_Revision+'. '+(d.note || '')});
  MD21_write_(sh,r._row,{Status:'Review Submitted',Approval_ID:out.approvalId}); return out;
}
function MD21_approveIssue_(u,d) {
  var ss = MD19_getSpreadsheet_(), sh = MD21_sheet_(ss,MD21_CONFIG_.ISSUES,MD21_CONFIG_.ISSUE_HEADERS), r = MD21_find_(sh,'Issue_ID',d.issueId); MD21_requireActive_(r); MD21_checkCompiled_(r);
  MD21_validateApproval_(r,d.approvalId,r.Issue_ID,'Approve newsletter issue');
  var cs = ss.getSheetByName(MD16_CONFIG_.CAMPAIGNS_SHEET), c = MD21_find_(cs,'Campaign_ID',r.Campaign_ID); MD21_requireActive_(c);
  var out = MD20_approve_(u,d);
  if (out.status === 'Approved') { MD21_write_(sh,r._row,{Status:'Approved',Approved_At:new Date()}); MD21_write_(cs,c._row,{Approval_Status:'Approved',Approved_At:new Date()}); }
  return out;
}
/* Archive chunks stay below the Sheets per-cell limit. No restore/resend action. */
function MD21_snapshot_(u,type,id,r,reason) {
  var sh = MD21_sheet_(MD19_getSpreadsheet_(),MD21_CONFIG_.ARCHIVE,MD21_CONFIG_.ARCHIVE_HEADERS);
  var prior = MD21_rows_(sh).filter(function(a) { return a.Record_Type === type && String(a.Record_ID) === String(id); });
  var json = prior.length ? null : JSON.stringify(r), chunks = prior.length ? Number(prior[0].Total_Parts) : Math.max(1,Math.ceil(json.length/40000));
  if (prior.length) {
    if (prior.length !== chunks) throw new Error('Incomplete archive for '+id+'. Repair the archive before continuing. Original row has not been removed.');
    return prior[0].Archived_At;
  }
  var now = new Date(), aid = MD21_id_('ARC'), summary = String(r.Title || r.Issue_Title || r.Subject || r.Action || id)+' · '+String(r.Status || r.Approval_Status || '');
  var rows = [];
  for (var i=0;i<chunks;i++) rows.push([aid,type,id,now,u.email,reason,summary.substring(0,500),i+1,chunks,json.substring(i*40000,(i+1)*40000)]);
  sh.getRange(sh.getLastRow()+1,1,rows.length,MD21_CONFIG_.ARCHIVE_HEADERS.length).setValues(rows.map(function(a) { return a.map(MD21_cell_); }));
  SpreadsheetApp.flush(); return now;
}
function MD21_archiveRow_(u,type,key,sh,r,reason) {
  if (!MD21_active_(r)) return false;
  sh = MD21_sheet_(MD19_getSpreadsheet_(),sh.getName(),MD21_CONFIG_.ARCHIVE_META);
  var now = MD21_snapshot_(u,type,r[key],r,reason);
  MD21_write_(sh,r._row,{Archived_At:now,Archived_By:u.email}); return true;
}
function MD21_listArchive_(u,d) {
  var f = String(d.recordType || 'all'), sh = MD19_getSpreadsheet_().getSheetByName(MD21_CONFIG_.ARCHIVE);
  var items = MD21_rows_(sh).filter(function(r) { return Number(r.Part) === 1 && (f === 'all' || r.Record_Type === f); }).reverse();
  var page = Math.max(1,Math.floor(Number(d.page) || 1)), total = items.length, pages = Math.max(1,Math.ceil(total/30)); page = Math.min(page,pages);
  return {ok:true,total:total,page:page,pageCount:pages,items:items.slice((page-1)*30,page*30).map(function(r) { return {archiveId:r.Archive_ID,recordType:r.Record_Type,recordId:r.Record_ID,archivedAt:r.Archived_At,archivedBy:r.Archived_By,reason:r.Reason,summary:r.Summary}; })};
}
function MD21_getArchive_(u,d) {
  var a = MD21_rows_(MD19_getSpreadsheet_().getSheetByName(MD21_CONFIG_.ARCHIVE)).filter(function(r) { return r.Archive_ID === d.archiveId; }).sort(function(x,y) { return Number(x.Part)-Number(y.Part); });
  if (!a.length || a.length !== Number(a[0].Total_Parts)) throw new Error('Complete archive snapshot not found.');
  return {ok:true,recordType:a[0].Record_Type,recordId:a[0].Record_ID,snapshot:MD21_json_(a.map(function(r) { return r.Snapshot_JSON; }).join(''),{})};
}
function MD21_archiveSent_(u) {
  var sh = MD19_getSpreadsheet_().getSheetByName(MD16_CONFIG_.QUEUE_SHEET), count = 0;
  MD21_rows_(sh).filter(function(r) { return r.Status === 'Sent' && MD21_active_(r); }).forEach(function(r) { if (MD21_archiveRow_(u,'Queue','Queue_ID',sh,r,'Sent delivery archived')) count++; });
  MD19_audit_(u.email,'ARCHIVE_SENT_NEWSLETTER','Newsletter','',{count:count},'Success');
  return {ok:true,count:count,message:count+' Sent record(s) archived. Original Sent rows remain to prevent duplicate delivery.'};
}
function MD21_archiveCampaign_(u,id,reason) {
  var ss = MD19_getSpreadsheet_(), cs = ss.getSheetByName(MD16_CONFIG_.CAMPAIGNS_SHEET), c = MD21_find_(cs,'Campaign_ID',id), rows = MD21_campaignRows_(id);
  if (rows.some(function(r) { return ['Sent','Cancelled'].indexOf(String(r.Status)) < 0; })) throw new Error('Resolve or finish all Pending, Failed, Sending and Needs Review rows before archiving this campaign.');
  MD21_archiveRow_(u,'Campaign','Campaign_ID',cs,c,reason);
  MD21_write_(cs,c._row,{Queue_Status:'Complete'});
  var qs = ss.getSheetByName(MD16_CONFIG_.QUEUE_SHEET);
  rows.forEach(function(r) { MD21_archiveRow_(u,'Queue','Queue_ID',qs,r,reason); });
}
function MD21_archiveRecord_(u,d) {
  var ss = MD19_getSpreadsheet_(), reason = 'Archived from Newsletter Studio';
  if (d.recordType === 'Item') {
    var is = MD21_sheet_(ss,MD21_CONFIG_.ITEMS,MD21_CONFIG_.ITEM_HEADERS), item = MD21_find_(is,'Item_ID',d.recordId); MD21_requireActive_(item);
    MD21_invalidateItemIssues_(u,item.Item_ID); MD21_archiveRow_(u,'Item','Item_ID',is,item,reason); MD21_supersede_(item.Item_ID,u,'Item archived');
  } else if (d.recordType === 'Issue') {
    var sh = MD21_sheet_(ss,MD21_CONFIG_.ISSUES,MD21_CONFIG_.ISSUE_HEADERS), issue = MD21_find_(sh,'Issue_ID',d.recordId); MD21_requireActive_(issue);
    if (issue.Campaign_ID) MD21_archiveCampaign_(u,issue.Campaign_ID,reason);
    MD21_archiveRow_(u,'Issue','Issue_ID',sh,issue,reason); MD21_supersede_(issue.Issue_ID,u,'Issue archived');
  } else if (d.recordType === 'Campaign') {
    MD21_archiveCampaign_(u,d.recordId,reason);
    var ish = MD21_sheet_(ss,MD21_CONFIG_.ISSUES,MD21_CONFIG_.ISSUE_HEADERS);
    MD21_rows_(ish).filter(function(r) { return MD21_active_(r) && r.Campaign_ID === d.recordId; }).forEach(function(r) { MD21_archiveRow_(u,'Issue','Issue_ID',ish,r,reason); MD21_supersede_(r.Issue_ID,u,'Campaign archived'); });
  } else throw new Error('Choose an Item, Issue or Campaign to archive.');
  MD19_audit_(u.email,'ARCHIVE_NEWSLETTER_RECORD','Newsletter',d.recordId,{type:d.recordType},'Success');
  return {ok:true,message:'Record archived. History is preserved and archive snapshots are read-only.'};
}
function MD21_resetActiveWorkspace_(u,d) {
  if (u.role !== MD19_CONFIG_.ROLE_CHIEF) throw new Error('Only the Chief Administrator can reset the active newsletter workspace.');
  if (d.confirmation !== 'RESET ACTIVE WORKSPACE') throw new Error('Type RESET ACTIVE WORKSPACE to confirm.');
  var ss = MD19_getSpreadsheet_(), qs = ss.getSheetByName(MD16_CONFIG_.QUEUE_SHEET), qrows = MD21_rows_(qs).filter(MD21_active_);
  if (qrows.some(function(r) { return r.Status === 'Sending' || r.Status === 'Needs Review'; })) throw new Error('Resolve Sending / Needs Review records first. Reset cannot guess whether delivery occurred.');
  var counts = {}, reason = 'Reset Active Workspace';
  qrows.forEach(function(r) {
    MD21_archiveRow_(u,'Queue','Queue_ID',qs,r,reason); counts.Queue = (counts.Queue || 0)+1;
    if (r.Status !== 'Sent' && r.Status !== 'Cancelled') MD21_write_(qs,r._row,{Status:'Cancelled',Error:'Cancelled by workspace reset. '+String(r.Error || '')});
  });
  [[MD21_CONFIG_.ITEMS,'Item','Item_ID'],[MD21_CONFIG_.ISSUES,'Issue','Issue_ID'],[MD16_CONFIG_.CAMPAIGNS_SHEET,'Campaign','Campaign_ID']].forEach(function(spec) {
    var sh = ss.getSheetByName(spec[0]);
    MD21_rows_(sh).filter(MD21_active_).forEach(function(r) {
      MD21_archiveRow_(u,spec[1],spec[2],sh,r,reason); counts[spec[1]] = (counts[spec[1]] || 0)+1;
      if (spec[1] === 'Campaign') MD21_write_(sh,r._row,{Approval_Status:'Draft',Queue_Status:'Complete'});
    });
  });
  var ash = ss.getSheetByName(MD20_CONFIG_.APPROVALS_SHEET), archivedApprovalIds = {};
  MD21_rows_(ss.getSheetByName(MD21_CONFIG_.ARCHIVE)).forEach(function(r) { if (r.Record_Type === 'Approval') archivedApprovalIds[r.Record_ID] = true; });
  MD21_rows_(ash).filter(function(r) { return r.Item_Type === 'Newsletter' && !archivedApprovalIds[r.Approval_ID]; }).forEach(function(r) {
    MD21_snapshot_(u,'Approval',r.Approval_ID,r,reason); counts.Approval = (counts.Approval || 0)+1;
    if (r.Status === 'Pending' || r.Status === 'Approved') MD21_write_(ash,r._row,{Status:'Superseded',Decision_Note:'Newsletter workspace reset; historical decision is archived',Last_Updated_At:new Date()});
  });
  MD19_audit_(u.email,'RESET_ACTIVE_NEWSLETTER_WORKSPACE','Newsletter','',counts,'Archived; no rows deleted');
  return {ok:true,counts:counts,message:'Active newsletter workspace reset by archiving records and cancelling unsent delivery. Existing archive, subscribers, administrators, tasks, public submissions and opportunities were preserved.'};
}
/* Legacy editor entry point is intentionally secured, not a destructive shortcut. */
function MD21_resetNewsletterWorkspace() {
  throw new Error('Use Reset Active Workspace in admin.html while signed in as the Chief Administrator. No data was changed.');
}
function MD21_listQueue_(u,d) {
  var ss = MD19_getSpreadsheet_(), qs = ss.getSheetByName(MD16_CONFIG_.QUEUE_SHEET), archived = {};
  MD21_rows_(qs).forEach(function(r) { if (!MD21_active_(r)) archived[r.Queue_ID] = true; });
  var out = MD16_listQueueForPortal(d.status);
  out.items = out.items.filter(function(r) { return !archived[r.queueId]; }); return out;
}
function MD21_requireQueueActive_(d) {
  var sh = MD19_getSpreadsheet_().getSheetByName(MD16_CONFIG_.QUEUE_SHEET), n = Number(d.rowNumber);
  var r = MD21_rows_(sh).filter(function(x) { return x._row === n; })[0];
  if (!r || (d.queueId && d.queueId !== r.Queue_ID)) throw new Error('Queue row changed. Refresh the queue before taking action.');
  MD21_requireActive_(r);
  var c = MD21_find_(MD19_getSpreadsheet_().getSheetByName(MD16_CONFIG_.CAMPAIGNS_SHEET),'Campaign_ID',r.Campaign_ID); MD21_requireActive_(c);
}
function MD21_listCampaigns_(u) {
  return {ok:true,items:MD21_rows_(MD19_getSpreadsheet_().getSheetByName(MD16_CONFIG_.CAMPAIGNS_SHEET)).filter(MD21_active_).map(function(r) { return {Campaign_ID:r.Campaign_ID,Issue_Title:r.Issue_Title,Approval_Status:r.Approval_Status,Queue_Status:r.Queue_Status}; })};
}
function MD21_attention_(u,tasks,notices) {
  var ss = MD19_getSpreadsheet_(), now = Date.now();
  var openTasks = tasks.filter(function(r) { return ['Completed','Cancelled'].indexOf(r.Status) < 0; });
  var pendingApprovals = MD21_rows_(ss.getSheetByName(MD20_CONFIG_.APPROVALS_SHEET)).filter(function(r) { return r.Status === 'Pending'; }).length;
  var q = MD21_rows_(ss.getSheetByName(MD16_CONFIG_.QUEUE_SHEET)).filter(MD21_active_), failed = q.filter(function(r) { return r.Status === 'Failed' || r.Status === 'Needs Review' || r.Status === 'Sending'; }).length;
  return {openTasks:openTasks.length,overdueTasks:openTasks.filter(function(r) { return r.Due_At && new Date(r.Due_At).getTime() < now; }).length,unacknowledged48h:openTasks.filter(function(r) { return r.Status === 'Assigned' && MD19_hoursSince_(r.Created_At) >= 48; }).length,unreadNotifications:notices.filter(function(r) { return !r.Read_At; }).length,pendingEditorialApprovals:pendingApprovals,newsletterDeliveryReview:failed,pendingNewsletterMessages:q.filter(function(r) { return r.Status === 'Pending'; }).length};
}

function MD21_previewIssue_(u,d) {
  var sh = MD21_sheet_(MD19_getSpreadsheet_(),MD21_CONFIG_.ISSUES,MD21_CONFIG_.ISSUE_HEADERS), r = MD21_find_(sh,'Issue_ID',d.issueId);
  MD21_requireActive_(r); if (!r.HTML_Body) throw new Error('Compile the saved issue before previewing it.');
  return {ok:true,html:String(r.HTML_Body),text:MD16_htmlToText_(r.HTML_Body)};
}