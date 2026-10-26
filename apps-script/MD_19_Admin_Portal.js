/*
 * THE MOBILIZER'S DIARY — ADMIN PORTAL BACKEND
 *
 * COMPLETE REPLACEMENT for MD_19_Admin_Portal.gs; keep its existing active name.
 * Install with MD_21_Newsletter_Studio.gs, MD_23_Opportunity_Studio.gs, MD_Launch_Edition_Email.html and admin.html.
 * Keep MD16 and MD20 unchanged. Run MD21_setupNewsletterStudio() once and redeploy.
 * Run MD19_setupAdminPortal() once while signed in as the Chief Administrator.
 * The function automatically creates the administration tabs in the existing CMS_ID workbook.
 *
 * Security model:
 * - Public intake remains public. Administrative actions require a signed server-issued token.
 * - Sign-in is verified by a short-lived code sent to an active Admin_Users email.
 * - The browser never supplies a trusted actor identity; the server derives it from the signed token.
 * - Roles, active status, responsibility text, and per-user session version are checked server-side.
 * - Email notices contain only category and reference, not submission details.
 */
var MD19_CONFIG_ = {
  USERS: 'Admin_Users',
  TASKS: 'Admin_Tasks',
  AUDIT: 'Admin_Audit_Log',
  NOTIFICATIONS: 'Admin_Notifications',
  USERS_HEADERS: [
    'Admin_ID','Full_Name','Email','Role','Responsibilities','Access_Status',
    'Availability_Status','Away_Until','Away_Note','Last_Login_At','Last_Activity_At',
    'Session_Version','Created_At','Created_By','Updated_At'
  ],
  TASK_HEADERS: [
    'Task_ID','Title','Details','Category','Record_Reference','Assigned_To_Email',
    'Assigned_By_Email','Created_At','Due_At','Status','Acknowledged_At',
    'Completed_At','Reassignment_Requested','Reassignment_Note','Email_Status','Updated_At'
  ],
  DELIVERABLES: 'Admin_Task_Deliverables',
  DELIVERABLE_HEADERS: ['Deliverable_ID','Task_ID','Submitted_By_Email','Submitted_At','Work_Summary','Work_Product_URL','Note_To_Chief'],
  AUDIT_HEADERS: [
    'Audit_ID','Timestamp','Actor_Email','Action','Item_Type','Item_Reference',
    'Details_JSON','Outcome'
  ],
  NOTIFICATION_HEADERS: [
    'Notification_ID','Timestamp','Recipient_Email','Type','Category','Item_Reference',
    'Message','Status','Error','Read_At'
  ],
  ROLE_CHIEF: 'Chief Administrator',
  ROLE_ADMIN: 'Administrator',
  STATUS_ACTIVE: 'Active',
  STATUS_INACTIVE: 'Inactive',
  AVAILABILITY: ['Available','Temporarily unavailable'],
  TASK_STATUS: ['Assigned','Acknowledged','In Progress','Completed','Reassignment Requested','Cancelled'],
  OTP_TTL_SECONDS: 600,
  OTP_REQUEST_LIMIT: 3,
  OTP_REQUEST_WINDOW_SECONDS: 900,
  OTP_MAX_ATTEMPTS: 5,
  SESSION_DAYS: 7,
  EMAIL_SENDER_NAME: "The Mobilizer's Diary",
  PORTAL_URL: 'https://01gullah.github.io/mobilizersdiary/admin.html',
  // Legacy and current report locations; discovery uses header names and does not create report tabs.
  REPORT_SHEET_CANDIDATES: ['Reports','Student_Reports','Report_Responses','Submission_Log'],
  OPPORTUNITY_RECOMMENDATIONS_SHEET: 'Opportunity_Recommendations',
  OPPORTUNITIES_SHEET: 'Opportunities',
  VOLUNTEERS_SHEET: 'Volunteer_Applications',
  BRAND_GREEN: '#123F3B',
  BRAND_GOLD: '#C59B35'
};

/** Run once from the project owner's account to create admin tabs and bootstrap the Chief Administrator. */
function MD19_setupAdminPortal() {
  var ss = MD19_getSpreadsheet_();
  var users = MD19_ensureSheet_(ss, MD19_CONFIG_.USERS, MD19_CONFIG_.USERS_HEADERS);
  MD19_ensureSheet_(ss, MD19_CONFIG_.TASKS, MD19_CONFIG_.TASK_HEADERS);
  MD19_ensureDeliverablesSheet_(ss);
  MD19_ensureSheet_(ss, MD19_CONFIG_.AUDIT, MD19_CONFIG_.AUDIT_HEADERS);
  MD19_ensureSheet_(ss, MD19_CONFIG_.NOTIFICATIONS, MD19_CONFIG_.NOTIFICATION_HEADERS);
  MD19_styleAdminSheets_(ss);
  MD19_addAdminDropdowns_(ss);
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('MD19_SESSION_SIGNING_SECRET')) {
    props.setProperty('MD19_SESSION_SIGNING_SECRET', Utilities.getUuid() + Utilities.getUuid());
  }
  var map = MD19_headerMap_(users, MD19_CONFIG_.USERS_HEADERS);
  var chief = MD19_findChief_(users, map);
  var bootstrapped = false;
  if (!chief) {
    if (users.getLastRow() > 1) {
      throw new Error('Admin_Users already contains users but no Chief Administrator. Do not bootstrap automatically; review the admin records first.');
    }
    var email = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
    if (!MD19_isEmail_(email)) {
      throw new Error('Could not identify the Google account running setup. Sign in as the project owner and run MD19_setupAdminPortal() again.');
    }
    var now = new Date();
    var row = [
      'ADM-' + Utilities.getUuid().substring(0,8).toUpperCase(),
      'Chief Administrator', email, MD19_CONFIG_.ROLE_CHIEF,
      'Full access; assigns and reviews all administrative work.',
      MD19_CONFIG_.STATUS_ACTIVE, 'Available', '', '', '', '', 1, now, email, now
    ];
    users.appendRow(row);
    bootstrapped = true;
    MD19_audit_('SYSTEM', 'BOOTSTRAP_CHIEF_ADMIN', 'Admin', email, { email: email }, 'Success');
  }
  Logger.log(JSON.stringify({ok:true, workbook:ss.getName(), users:MD19_CONFIG_.USERS, tasks:MD19_CONFIG_.TASKS, audit:MD19_CONFIG_.AUDIT, notifications:MD19_CONFIG_.NOTIFICATIONS, deliverables:MD19_CONFIG_.DELIVERABLES, chiefBootstrapped:bootstrapped}));
  return {ok:true, chiefBootstrapped:bootstrapped};
}

/** Entry point called by MD_13_WEB_API.gs for actions beginning admin. */
function MD19_route_(action, data) {
  data = data || {};
  if (action === 'adminRequestCode') return MD19_requestCode_(data.email);
  if (action === 'adminVerifyCode') return MD19_verifyCode_(data.email, data.code);
  if (action === 'adminLogout') {
    var logoutUser = MD19_requireAdmin_(data.token);
    MD19_incrementSessionVersion_(logoutUser.email);
    MD19_audit_(logoutUser.email, 'LOGOUT_ALL_SESSIONS', 'Admin', logoutUser.email, {}, 'Success');
    return {ok:true, message:'Signed out. All active portal sessions for this account have been ended.'};
  }
  var user = MD19_requireAdmin_(data.token);
  if (action === 'adminSendNewsletterBatch') return MD16_sendPendingNewsletterBatch(data.maxRows);
  // All portal mutation/queue-preparation actions share MD16's existing script lock.
  // Sending deliberately remains on MD16's own per-message locking path.
  return MD21_withLock_(function () {
  switch (action) {
    case 'adminDashboard': return MD19_dashboard_(user);
    case 'adminListSubmissions': return MD19_listSubmissions_(user, data);
    case 'adminSetAvailability': return MD19_setAvailability_(user, data);
    case 'adminAcknowledgeTask': return MD19_updateTask_(user, data, 'Acknowledged');
    case 'adminUpdateTask': return MD19_updateTask_(user, data, data.status);
    case 'adminCreateTask': return MD19_createTask_(user, data);
    case 'adminReassignTask': return MD19_reassignTask_(user, data);
    case 'adminRetryTaskEmail': return MD19_retryTaskEmail_(user, data.taskId);
    case 'adminSubmitTaskDeliverable': return MD19_submitTaskDeliverable_(user, data);
    case 'adminManageUser': return MD19_manageUser_(user, data);
    case 'adminGetEntry': return MD19_getEntry_(user, data.category, data.ref);
    case 'adminMarkNotificationRead': return MD19_markNotificationRead_(user, data.notificationId);
    case 'adminRetryNotification': return MD19_retryNotification_(user, data.notificationId);
    case 'adminRequestEditorialApproval':
      if (data.itemType === 'Newsletter') throw new Error('Use Newsletter Studio to approve a specific current revision.');
      return MD20_requestApproval_(user, data);
    case 'adminApproveEditorialItem':
      var editorialRecord = MD20_find_(data.approvalId);
      if (editorialRecord && editorialRecord.itemType === 'Newsletter') throw new Error('Use Newsletter Studio to approve a specific current revision.');
      return MD20_approve_(user, data);
    case 'adminPublishOpportunity': return MD20_publishOpportunity_(user, data);
    // Opportunity Studio (MD_23)
    case 'adminStudioList':
    case 'adminStudioLoad':
    case 'adminStudioSave':
    case 'adminStudioShortlist':
    case 'adminStudioPublish':
    case 'adminStudioUnpublish':
      return MD23_route_(action, user, data);
    case 'adminListNewsletterQueue': return MD21_listQueue_(user, data);
    case 'adminRetryFailedNewsletterRow': MD21_requireQueueActive_(data); return MD16_retryFailedQueueRow(data.rowNumber, data.reviewNote, user.email);
    case 'adminResolveNewsletterRow': MD21_requireQueueActive_(data); return MD16_resolveNeedsReviewQueueRow(data.rowNumber, data.decision, data.reviewNote, user.email);
    case 'adminPrepareApprovedCampaigns': return MD16_prepareApprovedCampaigns();
    case 'adminCreateNewsletterItem': return MD21_createEditorialItem_(user, data);
    case 'adminListNewsletterItems': return MD21_listItems_(user, data.status);
    case 'adminRequestNewsletterItemApproval': return MD21_requestItemApproval_(user, data);
    case 'adminApproveNewsletterItem': return MD21_approveItem_(user, data);
    case 'adminCreateNewsletterIssue': return MD21_createIssue_(user, data);
    case 'adminCompileNewsletterIssue': return MD21_compileIssue_(user, data);
    case 'adminRequestNewsletterApproval': return MD21_requestIssueApproval_(user, data);
    case 'adminApproveNewsletterIssue': return MD21_approveIssue_(user, data);
    case 'adminGetNewsletterItem': return MD21_getItem_(user, data);
    case 'adminUpdateNewsletterItem': return MD21_updateEditorialItem_(user, data);
        // Paystack Financial Engine Integration
    case 'adminVerifyPaystackPayment':
      return MD22_verifyPaystackPayment_(data);

    case 'adminListNewsletterIssues': return MD21_listIssues_(user);
    case 'adminUpdateNewsletterIssue': return MD21_updateIssue_(user, data);
    case 'adminPreviewNewsletterIssue': return MD21_previewIssue_(user, data);
    case 'adminListNewsletterCampaigns': return MD21_listCampaigns_(user);
    case 'adminArchiveNewsletterRecord': return MD21_archiveRecord_(user, data);
    case 'adminArchiveSentNewsletterRecords': return MD21_archiveSent_(user);
    case 'adminListNewsletterArchive': return MD21_listArchive_(user, data);
    case 'adminGetNewsletterArchive': return MD21_getArchive_(user, data);
    case 'adminResetNewsletterWorkspace': return MD21_resetActiveWorkspace_(user, data);
    default: throw new Error('Unknown administrative request.');
  }
  });
}

/** Called after a successful public form action; failures here never undo the saved submission. */
function MD19_afterPublicIntake_(category, handler) {
  var result = typeof handler === 'function' ? handler() : handler;
  if (!MD19_intakeWasAccepted_(result)) return result;
  var ref = MD19_extractRef_(result);
  try {
    MD19_notifyAdmins_(category, ref);
  } catch (err) {
    Logger.log('Admin notification failure for ' + category + ' ' + ref + ': ' + MD19_errorText_(err));
  }
  return result;
}

function MD19_intakeWasAccepted_(result) {
  if (!result || typeof result !== 'object') return false;
  if (result.accepted === false || result.ok === false || result.status === 'already') return false;
  if (result.accepted === true || result.success === true || result.ref || result.reference || result.status === 'subscribed') return true;
  // Validation failures throw before returning. A successfully returned object from a form handler is a saved/accepted action.
  return Object.keys(result).length > 0;
}

function MD19_extractRef_(result) {
  if (!result || typeof result !== 'object') return '';
  return String(result.ref || result.reference || result.referenceNumber || result.id || '').trim().substring(0,120);
}

function MD19_notifyAdmins_(category, ref) {
  var ss = MD19_getSpreadsheet_();
  var usersSheet = ss.getSheetByName(MD19_CONFIG_.USERS);
  var notices = ss.getSheetByName(MD19_CONFIG_.NOTIFICATIONS);
  if (!usersSheet || !notices) throw new Error('Admin portal has not been set up. Run MD19_setupAdminPortal() first.');
  var map = MD19_headerMap_(usersSheet, MD19_CONFIG_.USERS_HEADERS);
  var recipients = MD19_activeAdmins_(usersSheet, map);
  var now = new Date();
  var safeCategory = String(category || 'New submission').substring(0,100);
  var safeRef = String(ref || 'Reference pending').substring(0,120);
  var message = 'New ' + safeCategory + ' received. Reference: ' + safeRef + '. Sign in to the Admin Portal to review.';
  recipients.forEach(function (admin) {
    var notificationId = 'NTF-' + Utilities.getUuid().substring(0,12).toUpperCase();
    var rowNumber = notices.getLastRow() + 1;
    notices.appendRow([notificationId, now, admin.email, 'New submission', safeCategory, safeRef, message, 'Queued', '', '']);
    try {
      MailApp.sendEmail({
        to: admin.email,
        subject: 'New ' + safeCategory + ' — The Mobilizer\'s Diary',
        body: message + '\n\nOpen the Admin Portal: ' + MD19_CONFIG_.PORTAL_URL,
        name: MD19_CONFIG_.EMAIL_SENDER_NAME
      });
      notices.getRange(rowNumber, 8).setValue('Sent');
    } catch (mailError) {
      notices.getRange(rowNumber, 8).setValue('Failed');
      notices.getRange(rowNumber, 9).setValue(MD19_errorText_(mailError).substring(0,500));
      Logger.log('Could not notify ' + admin.email + ' for ' + safeCategory + ' ' + safeRef + ': ' + MD19_errorText_(mailError));
    }
  });
}

function MD19_requestCode_(emailInput) {
  var email = String(emailInput || '').trim().toLowerCase();
  var generic = {ok:true, message:'If this email is an active administrator, a sign-in code has been sent.'};
  if (!MD19_isEmail_(email)) return generic;
  var cache = CacheService.getScriptCache();
  var hash = MD19_sha256_(email).substring(0,32);
  var rateKey = 'MD19_RATE_' + hash;
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var user = MD19_findUserByEmail_(email);
    if (!user || user.accessStatus !== MD19_CONFIG_.STATUS_ACTIVE) return generic;
    var count = Number(cache.get(rateKey) || 0);
    if (count >= MD19_CONFIG_.OTP_REQUEST_LIMIT) return generic;
    cache.put(rateKey, String(count + 1), MD19_CONFIG_.OTP_REQUEST_WINDOW_SECONDS);
    var digits = Utilities.getUuid().replace(/\D/g, '').substring(0,8);
    while (digits.length < 8) digits += String(Math.floor(Math.random() * 10));
    var salt = Utilities.getUuid();
    var otpHash = MD19_sha256_(email + '|' + digits + '|' + salt);
    cache.put('MD19_OTP_' + hash, JSON.stringify({salt:salt, hash:otpHash, attempts:0}), MD19_CONFIG_.OTP_TTL_SECONDS);
    try {
      MailApp.sendEmail({
        to: email,
        subject: 'Your Admin Portal sign-in code',
        body: 'Your sign-in code for The Mobilizer\'s Diary Admin Portal is:\n\n' + digits + '\n\nIt expires in 10 minutes and can be used once. If you did not request it, ignore this message.\n\nPortal: ' + MD19_CONFIG_.PORTAL_URL,
        name: MD19_CONFIG_.EMAIL_SENDER_NAME
      });
    } catch (sendError) {
      cache.remove('MD19_OTP_' + hash);
      Logger.log('Admin OTP email could not be sent for an active account: ' + MD19_errorText_(sendError));
    }
    return generic;
  } finally {
    lock.releaseLock();
  }
}

function MD19_verifyCode_(emailInput, codeInput) {
  var email = String(emailInput || '').trim().toLowerCase();
  var code = String(codeInput || '').replace(/\s/g, '');
  if (!MD19_isEmail_(email) || !/^\d{8}$/.test(code)) throw new Error('The email or code is invalid or expired. Request a new code and try again.');
  var hash = MD19_sha256_(email).substring(0,32);
  var cache = CacheService.getScriptCache();
  var key = 'MD19_OTP_' + hash;
  var raw = cache.get(key);
  var user = MD19_findUserByEmail_(email);
  if (!raw || !user || user.accessStatus !== MD19_CONFIG_.STATUS_ACTIVE) throw new Error('The email or code is invalid or expired. Request a new code and try again.');
  var record = JSON.parse(raw);
  var suppliedHash = MD19_sha256_(email + '|' + code + '|' + record.salt);
  if (!MD19_constantTimeEqual_(suppliedHash, record.hash)) {
    record.attempts = Number(record.attempts || 0) + 1;
    if (record.attempts >= MD19_CONFIG_.OTP_MAX_ATTEMPTS) cache.remove(key);
    else cache.put(key, JSON.stringify(record), MD19_CONFIG_.OTP_TTL_SECONDS);
    throw new Error('The email or code is invalid or expired. Request a new code and try again.');
  }
  cache.remove(key);
  var fresh = MD19_findUserByEmail_(email);
  if (!fresh || fresh.accessStatus !== MD19_CONFIG_.STATUS_ACTIVE) throw new Error('This administrator account is inactive. Contact the Chief Administrator.');
  var now = new Date();
  MD19_setUserField_(fresh.row, 'Last_Login_At', now);
  MD19_setUserField_(fresh.row, 'Last_Activity_At', now);
  var payload = {
    email: email,
    version: Number(fresh.sessionVersion || 1),
    exp: Math.floor(Date.now()/1000) + MD19_CONFIG_.SESSION_DAYS * 86400,
    nonce: Utilities.getUuid()
  };
  var token = MD19_signToken_(payload);
  MD19_audit_(email, 'LOGIN', 'Admin', email, {expiresAt:new Date(payload.exp*1000).toISOString()}, 'Success');
  return {ok:true, token:token, expiresAt:new Date(payload.exp*1000).toISOString(), user:{name:fresh.name, email:fresh.email, role:fresh.role, responsibilities:fresh.responsibilities, availability:fresh.availability}};
}

function MD19_signToken_(payload) {
  var encoded = Utilities.base64EncodeWebSafe(JSON.stringify(payload)).replace(/=+$/,'');
  var secret = PropertiesService.getScriptProperties().getProperty('MD19_SESSION_SIGNING_SECRET');
  if (!secret) throw new Error('Admin portal is not initialized. Run MD19_setupAdminPortal().');
  return encoded + '.' + MD19_hmac_(encoded, secret);
}

function MD19_requireAdmin_(token) {
  token = String(token || '').trim();
  var parts = token.split('.');
  if (parts.length !== 2) throw new Error('Please sign in again.');
  var secret = PropertiesService.getScriptProperties().getProperty('MD19_SESSION_SIGNING_SECRET');
  if (!secret || !MD19_constantTimeEqual_(parts[1], MD19_hmac_(parts[0], secret))) throw new Error('Please sign in again.');
  var payload;
  try { payload = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString()); }
  catch (e) { throw new Error('Please sign in again.'); }
  if (!payload.email || Number(payload.exp) < Math.floor(Date.now()/1000)) throw new Error('Your sign-in session has expired. Request a new code.');
  var user = MD19_findUserByEmail_(String(payload.email).toLowerCase());
  if (!user || user.accessStatus !== MD19_CONFIG_.STATUS_ACTIVE || Number(user.sessionVersion || 1) !== Number(payload.version)) {
    throw new Error('This administrator account is inactive or its session has been ended. Sign in again or contact the Chief Administrator.');
  }
  MD19_setUserField_(user.row, 'Last_Activity_At', new Date());
  return user;
}

function MD19_dashboard_(user) {
  MD19_expireAwayStatuses_();
  var ss = MD19_getSpreadsheet_();
  var usersSheet = ss.getSheetByName(MD19_CONFIG_.USERS);
  var tasksSheet = ss.getSheetByName(MD19_CONFIG_.TASKS);
  var noticesSheet = ss.getSheetByName(MD19_CONFIG_.NOTIFICATIONS);
  var taskRows = MD19_rowsAsObjects_(tasksSheet, MD19_CONFIG_.TASK_HEADERS);
  var noticeRows = MD19_rowsAsObjects_(noticesSheet, MD19_CONFIG_.NOTIFICATION_HEADERS);
  var isChief = user.role === MD19_CONFIG_.ROLE_CHIEF;
  // The whole team can see every assignment; mutation remains restricted to the assignee and Chief.
  var myTasks = taskRows.slice().reverse();
  var myNotices = noticeRows.filter(function (row) {
    return String(row.Recipient_Email || '').toLowerCase() === user.email;
  }).slice(-100).reverse();
  var deliverableRows = MD19_rowsAsObjects_(ss.getSheetByName(MD19_CONFIG_.DELIVERABLES), MD19_CONFIG_.DELIVERABLE_HEADERS);
  var visibleTaskIds = {};
  myTasks.forEach(function(row){visibleTaskIds[String(row.Task_ID||'')]=true;});
  var taskDeliverables = deliverableRows.filter(function(row){return visibleTaskIds[String(row.Task_ID||'')];}).slice(-500).reverse();
  var out = {
    ok:true,
    me:{name:user.name,email:user.email,role:user.role,responsibilities:user.responsibilities,availability:user.availability,awayUntil:user.awayUntil,awayNote:user.awayNote,lastLogin:user.lastLogin,lastActivity:user.lastActivity},
    isChief:isChief,
    tasks:myTasks,
    taskDeliverables:taskDeliverables,
    notifications:myNotices,
    inactivity:{daysSinceActivity:MD19_daysSince_(user.lastActivity), reviewAfterDays:14, taskAckHours:48},
    overdueAcknowledgements:isChief?taskRows.filter(function(r){return r.Status==='Assigned' && MD19_hoursSince_(r.Created_At)>=48;}):[]
  };
  out.attention = MD21_attention_(user, taskRows, myNotices);
  if (isChief) out.admins = MD19_rowsAsObjects_(usersSheet, MD19_CONFIG_.USERS_HEADERS).filter(function(r){return String(r.Email||'').trim();});
  MD19_audit_(user.email, 'VIEW_DASHBOARD', 'Admin', user.email, {}, 'Success');
  return out;
}

function MD19_setAvailability_(user, data) {
  var status = String(data.status || '').trim();
  if (MD19_CONFIG_.AVAILABILITY.indexOf(status) < 0) throw new Error('Choose Available or Temporarily unavailable.');
  var until = String(data.awayUntil || '').trim();
  var note = MD19_clean_(data.note, 300);
  if (status === 'Temporarily unavailable' && until && isNaN(new Date(until).getTime())) throw new Error('Enter a valid return date.');
  MD19_setUserField_(user.row, 'Availability_Status', status);
  MD19_setUserField_(user.row, 'Away_Until', status === 'Temporarily unavailable' ? until : '');
  MD19_setUserField_(user.row, 'Away_Note', status === 'Temporarily unavailable' ? note : '');
  MD19_setUserField_(user.row, 'Updated_At', new Date());
  MD19_audit_(user.email, 'CHANGE_AVAILABILITY', 'Admin', user.email, {status:status, awayUntil:until, note:note}, 'Success');
  return {ok:true,message:'Availability updated.'};
}

function MD19_expireAwayStatuses_() {
  var ss=MD19_getSpreadsheet_();
  var sheet=ss.getSheetByName(MD19_CONFIG_.USERS);
  if(!sheet||sheet.getLastRow()<2)return;
  var map=MD19_headerMap_(sheet,MD19_CONFIG_.USERS_HEADERS);
  var rows=sheet.getRange(2,1,sheet.getLastRow()-1,MD19_CONFIG_.USERS_HEADERS.length).getValues();
  var today=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');
  rows.forEach(function(row,i){
    var status=String(row[map.Availability_Status]||'');
    var until=row[map.Away_Until];
    if(status!=='Temporarily unavailable'||!until)return;
    var dateText=Utilities.formatDate(new Date(until),Session.getScriptTimeZone(),'yyyy-MM-dd');
    if(dateText>=today)return;
    var actualRow=i+2;
    sheet.getRange(actualRow,map.Availability_Status+1).setValue('Available');
    sheet.getRange(actualRow,map.Away_Until+1).clearContent();
    sheet.getRange(actualRow,map.Away_Note+1).clearContent();
    sheet.getRange(actualRow,map.Updated_At+1).setValue(new Date());
    MD19_audit_('SYSTEM','AUTO_RETURN_TO_AVAILABLE','Admin',String(row[map.Email]||''),{returnDate:dateText},'Success');
  });
}

function MD19_createTask_(user, data) {
  if (user.role !== MD19_CONFIG_.ROLE_CHIEF) throw new Error('Only the Chief Administrator can assign tasks.');
  var title = MD19_clean_(data.title, 180);
  var details = MD19_clean_(data.details, 3000);
  var category = MD19_clean_(data.category, 100) || 'General';
  var ref = MD19_clean_(data.ref, 120);
  var assigneeEmail = String(data.assigneeEmail || '').trim().toLowerCase();
  var assignee = MD19_findUserByEmail_(assigneeEmail);
  if (!title) throw new Error('Enter a task title.');
  if (!assignee || assignee.accessStatus !== MD19_CONFIG_.STATUS_ACTIVE) throw new Error('Choose an active administrator.');
  if (ref) MD19_getEntry_(user, category, ref);
  var due = String(data.dueAt || '').trim();
  if (due && isNaN(new Date(due).getTime())) throw new Error('Enter a valid due date.');
  var taskId = 'TASK-' + Utilities.getUuid().replace(/-/g,'').substring(0,12).toUpperCase();
  var now = new Date();
  var sheet = MD19_getSpreadsheet_().getSheetByName(MD19_CONFIG_.TASKS);
  sheet.appendRow([taskId,title,details,category,ref,assigneeEmail,user.email,now,due,'Assigned','','','No','', 'Queued',now]);
  var emailStatus = 'Sent';
  try {
    MailApp.sendEmail({
      to:assigneeEmail,
      subject:'New task assigned: ' + title,
      body:'A task has been assigned to you by the Chief Administrator.\n\nTask: ' + title + '\nCategory: ' + category + '\nReference: ' + (ref || 'Not linked to a form entry') + '\nDue: ' + (due || 'Not specified') + '\n\nSign in to review and acknowledge it: ' + MD19_CONFIG_.PORTAL_URL,
      name:MD19_CONFIG_.EMAIL_SENDER_NAME
    });
  } catch (e) {
    emailStatus = 'Failed: ' + MD19_errorText_(e).substring(0,300);
    sheet.getRange(sheet.getLastRow(),15).setValue(emailStatus);
  }
  MD19_audit_(user.email,'CREATE_AND_ASSIGN_TASK','Task',taskId,{title:title,category:category,ref:ref,assignee:assigneeEmail,dueAt:due},'Success; email '+(emailStatus==='Sent'?'sent':'failed'));
  return {ok:true,taskId:taskId,emailStatus:emailStatus};
}

function MD19_updateTask_(user, data, statusInput) {
  var taskId = MD19_clean_(data.taskId,80);
  var status = String(statusInput || '').trim();
  var ss = MD19_getSpreadsheet_();
  var sheet = ss.getSheetByName(MD19_CONFIG_.TASKS);
  var map = MD19_headerMap_(sheet, MD19_CONFIG_.TASK_HEADERS);
  var row = MD19_findRowByValue_(sheet,map.Task_ID,taskId);
  if (!row) throw new Error('Task not found.');
  var assigned = String(sheet.getRange(row,map.Assigned_To_Email+1).getDisplayValue()||'').trim().toLowerCase();
  var isChief = user.role === MD19_CONFIG_.ROLE_CHIEF;
  if (!isChief && assigned !== user.email) throw new Error('You can update only tasks assigned to you.');
  if (status === 'Request Reassignment') {
    var note = MD19_clean_(data.note,1000);
    sheet.getRange(row,map.Status+1).setValue('Reassignment Requested');
    sheet.getRange(row,map.Reassignment_Requested+1).setValue('Yes');
    sheet.getRange(row,map.Reassignment_Note+1).setValue(note);
    sheet.getRange(row,map.Updated_At+1).setValue(new Date());
    MD19_notifyChiefs_('Reassignment requested',taskId,'Task ' + taskId + ' has a reassignment request from ' + user.email + '.');
    MD19_audit_(user.email,'REQUEST_TASK_REASSIGNMENT','Task',taskId,{note:note},'Success');
    return {ok:true,message:'Your reassignment request has been sent to the Chief Administrator.'};
  }
  if (status === 'Completed') throw new Error('Submit the finished work using the task deliverable form; that records the work and marks the task complete.');
  if (MD19_CONFIG_.TASK_STATUS.indexOf(status)<0 || status === 'Assigned' || status === 'Reassignment Requested') throw new Error('Choose Acknowledged, In Progress, or Cancelled.');
  sheet.getRange(row,map.Status+1).setValue(status);
  sheet.getRange(row,map.Updated_At+1).setValue(new Date());
  if (status === 'Acknowledged' && !sheet.getRange(row,map.Acknowledged_At+1).getValue()) sheet.getRange(row,map.Acknowledged_At+1).setValue(new Date());
  if (status === 'Completed') sheet.getRange(row,map.Completed_At+1).setValue(new Date());
  MD19_audit_(user.email,'UPDATE_TASK_STATUS','Task',taskId,{status:status},'Success');
  return {ok:true,message:'Task updated.'};
}

function MD19_reassignTask_(user, data) {
  if (user.role !== MD19_CONFIG_.ROLE_CHIEF) throw new Error('Only the Chief Administrator can reassign tasks.');
  var taskId = MD19_clean_(data.taskId,80);
  var email = String(data.assigneeEmail||'').trim().toLowerCase();
  var assignee = MD19_findUserByEmail_(email);
  if (!assignee || assignee.accessStatus !== MD19_CONFIG_.STATUS_ACTIVE) throw new Error('Choose an active administrator.');
  var sheet = MD19_getSpreadsheet_().getSheetByName(MD19_CONFIG_.TASKS);
  var map = MD19_headerMap_(sheet,MD19_CONFIG_.TASK_HEADERS);
  var row = MD19_findRowByValue_(sheet,map.Task_ID,taskId);
  if (!row) throw new Error('Task not found.');
  var now = new Date();
  sheet.getRange(row,map.Assigned_To_Email+1).setValue(email);
  sheet.getRange(row,map.Assigned_By_Email+1).setValue(user.email);
  sheet.getRange(row,map.Status+1).setValue('Assigned');
  sheet.getRange(row,map.Acknowledged_At+1).clearContent();
  sheet.getRange(row,map.Reassignment_Requested+1).setValue('No');
  sheet.getRange(row,map.Reassignment_Note+1).setValue('');
  sheet.getRange(row,map.Email_Status+1).setValue('Queued');
  sheet.getRange(row,map.Updated_At+1).setValue(now);
  var title=String(sheet.getRange(row,map.Title+1).getDisplayValue()||'Assigned task');
  var category=String(sheet.getRange(row,map.Category+1).getDisplayValue()||'General');
  var ref=String(sheet.getRange(row,map.Record_Reference+1).getDisplayValue()||'');
  var due=String(sheet.getRange(row,map.Due_At+1).getDisplayValue()||'Not specified');
  var emailStatus='Sent';
  try {
    MailApp.sendEmail({to:email,subject:'Task assigned to you: '+title,body:'The Chief Administrator assigned this task to you.\n\nTask: '+title+'\nCategory: '+category+'\nReference: '+(ref||'Not linked to a form entry')+'\nDue: '+due+'\n\nSign in to review it: '+MD19_CONFIG_.PORTAL_URL,name:MD19_CONFIG_.EMAIL_SENDER_NAME});
  } catch(e) { emailStatus='Failed: '+MD19_errorText_(e).substring(0,300); }
  sheet.getRange(row,map.Email_Status+1).setValue(emailStatus);
  MD19_audit_(user.email,'REASSIGN_TASK','Task',taskId,{assignee:email},'Success; email '+(emailStatus==='Sent'?'sent':'failed'));
  return {ok:true,message:'Task reassigned.',emailStatus:emailStatus};
}

function MD19_retryTaskEmail_(user, taskIdInput) {
  if(user.role!==MD19_CONFIG_.ROLE_CHIEF)throw new Error('Only the Chief Administrator can retry assignment emails.');
  var taskId=MD19_clean_(taskIdInput,80);
  var sheet=MD19_getSpreadsheet_().getSheetByName(MD19_CONFIG_.TASKS);
  var map=MD19_headerMap_(sheet,MD19_CONFIG_.TASK_HEADERS);
  var row=MD19_findRowByValue_(sheet,map.Task_ID,taskId);
  if(!row)throw new Error('Task not found.');
  var email=String(sheet.getRange(row,map.Assigned_To_Email+1).getDisplayValue()||'').trim().toLowerCase();
  var assignee=MD19_findUserByEmail_(email);
  if(!assignee||assignee.accessStatus!==MD19_CONFIG_.STATUS_ACTIVE)throw new Error('The assigned administrator is inactive. Reassign the task first.');
  var title=String(sheet.getRange(row,map.Title+1).getDisplayValue()||'Assigned task');
  var category=String(sheet.getRange(row,map.Category+1).getDisplayValue()||'General');
  var ref=String(sheet.getRange(row,map.Record_Reference+1).getDisplayValue()||'');
  var due=String(sheet.getRange(row,map.Due_At+1).getDisplayValue()||'Not specified');
  try {
    MailApp.sendEmail({to:email,subject:'Task assigned to you: '+title,body:'A task has been assigned to you by the Chief Administrator.\n\nTask: '+title+'\nCategory: '+category+'\nReference: '+(ref||'Not linked to a form entry')+'\nDue: '+due+'\n\nSign in to review it: '+MD19_CONFIG_.PORTAL_URL,name:MD19_CONFIG_.EMAIL_SENDER_NAME});
    sheet.getRange(row,map.Email_Status+1).setValue('Sent');
    MD19_audit_(user.email,'RETRY_TASK_ASSIGNMENT_EMAIL','Task',taskId,{recipient:email},'Success');
    return {ok:true,message:'Assignment email sent.'};
  } catch(e) {
    var err=MD19_errorText_(e).substring(0,300);
    sheet.getRange(row,map.Email_Status+1).setValue('Failed: '+err);
    MD19_audit_(user.email,'RETRY_TASK_ASSIGNMENT_EMAIL','Task',taskId,{recipient:email},'Failed: '+err);
    return {ok:false,message:'Assignment email could not be sent: '+err};
  }
}

function MD19_manageUser_(user, data) {
  if (user.role !== MD19_CONFIG_.ROLE_CHIEF) throw new Error('Only the Chief Administrator can manage administrator accounts.');
  var email = String(data.email || '').trim().toLowerCase();
  var name = MD19_clean_(data.name,120);
  var responsibilities = MD19_clean_(data.responsibilities,1200);
  var access = String(data.accessStatus || MD19_CONFIG_.STATUS_ACTIVE).trim();
  var role = MD19_clean_(data.role,80) || MD19_CONFIG_.ROLE_ADMIN;
  if (!MD19_isEmail_(email)) throw new Error('Enter a valid administrator email address.');
  if (!name) throw new Error('Enter the administrator’s name.');
  if (access !== MD19_CONFIG_.STATUS_ACTIVE && access !== MD19_CONFIG_.STATUS_INACTIVE) throw new Error('Choose Active or Inactive access.');
  if (role === MD19_CONFIG_.ROLE_CHIEF && email !== user.email) throw new Error('A second Chief Administrator cannot be created in this version.');
  if (email === user.email && access === MD19_CONFIG_.STATUS_INACTIVE) throw new Error('You cannot deactivate your own Chief Administrator account.');
  var sheet = MD19_getSpreadsheet_().getSheetByName(MD19_CONFIG_.USERS);
  var map = MD19_headerMap_(sheet,MD19_CONFIG_.USERS_HEADERS);
  var existing = MD19_findUserByEmail_(email);
  var now = new Date();
  if (existing) {
    if (email === user.email) role = MD19_CONFIG_.ROLE_CHIEF;
    sheet.getRange(existing.row,map.Full_Name+1).setValue(name);
    sheet.getRange(existing.row,map.Role+1).setValue(role);
    sheet.getRange(existing.row,map.Responsibilities+1).setValue(responsibilities);
    sheet.getRange(existing.row,map.Access_Status+1).setValue(access);
    sheet.getRange(existing.row,map.Updated_At+1).setValue(now);
    if (access === MD19_CONFIG_.STATUS_INACTIVE || existing.role !== role) sheet.getRange(existing.row,map.Session_Version+1).setValue(Number(existing.sessionVersion||1)+1);
    MD19_audit_(user.email,'UPDATE_ADMIN','Admin',email,{name:name,role:role,responsibilities:responsibilities,accessStatus:access},'Success');
  } else {
    if (role === MD19_CONFIG_.ROLE_CHIEF) throw new Error('Only the bootstrapped Chief Administrator can have that role.');
    var row = ['ADM-' + Utilities.getUuid().substring(0,8).toUpperCase(),name,email,role,responsibilities,access,'Available','','','','',1,now,user.email,now];
    sheet.appendRow(row);
    MD19_audit_(user.email,'ADD_ADMIN','Admin',email,{name:name,role:role,responsibilities:responsibilities},'Success');
  }
  return {ok:true,message:'Administrator record saved.'};
}

function MD19_listSubmissions_(user, data) {
  data=data||{};
  var category=String(data.category||'all').trim().toLowerCase();
  var query=MD19_clean_(data.query,120).toLowerCase();
  var page=Math.max(1,Math.floor(Number(data.page)||1));
  var pageSize=40;
  var items=MD19_collectSubmissions_().filter(function(item){
    var group=MD19_submissionGroup_(item.category);
    if(category!=='all'&&group!==category)return false;
    if(!query)return true;
    return [item.ref,item.category,item.title,item.name,item.institution,item.summary,item.status].join(' ').toLowerCase().indexOf(query)>=0;
  });
  items.sort(function(a,b){return b.sortTime-a.sortTime;});
  var total=items.length;
  var pageCount=Math.max(1,Math.ceil(total/pageSize));
  if(page>pageCount)page=pageCount;
  var result=items.slice((page-1)*pageSize,page*pageSize).map(function(item){delete item.sortTime;return item;});
  MD19_audit_(user.email,'LIST_SUBMISSIONS','SubmissionInbox','Page '+page,{category:category,total:total},'Success');
  return {ok:true,items:result,total:total,page:page,pageSize:pageSize,pageCount:pageCount};
}

function MD19_collectSubmissions_() {
  var ss=MD19_getSpreadsheet_();
  var excluded={
    'Issue_Settings':true,'Student_Voices':true,'History_Corner':true,'Podcast':true,
    'Opportunities':true,'Opportunity_Studio':true,'NewsletterQueue':true,'Newsletter_Campaigns':true,'SponsoredBanners':true
  };
  var sheets=ss.getSheets().filter(function(sh){
    var name=sh.getName();
    return name.indexOf('Admin_')!==0&&(name.indexOf('Newsletter_')!==0||name==='Newsletter_Subscribers')&&!excluded[name];
  });
  sheets.sort(function(a,b){return (a.getName()==='Submission_Log'?1:0)-(b.getName()==='Submission_Log'?1:0);});
  var byRef={};
  var idAliases=['Reference_Number','Reference','Report_ID','Submission_ID','Application_ID','Recommendation_ID','Opp_ID','Donation_Reference','Subscriber_ID','ID'];
  sheets.forEach(function(sh){
    var lastRow=sh.getLastRow(),width=sh.getLastColumn();
    if(lastRow<2||width<1)return;
    var headers=sh.getRange(1,1,1,width).getDisplayValues()[0].map(function(v){return String(v||'').trim();});
    var idIndex=MD19_findHeaderIndex_(headers,idAliases);
    if(idIndex<0)return;
    var values=sh.getRange(2,1,lastRow-1,width).getValues();
    var display=sh.getRange(2,1,lastRow-1,width).getDisplayValues();
    function idx(names){return MD19_findHeaderIndex_(headers,names);}
    var dateIndex=idx(['Received_At','Timestamp','Date_Received','Submitted_At','Created_At','Date']);
    var typeIndex=idx(['Submission_Type','Category','Type']);
    var nameIndex=idx(['Name','Submitter_Name','Full_Name','Display_Name']);
    var titleIndex=idx(['Title','Report_Title','Subject','Headline']);
    var institutionIndex=idx(['Institution','Institution_Name']);
    var summaryIndex=idx(['Summary','Description','Challenge','Message','Additional_Details','Motivation','Why_It_Matters','Concern','Issue']);
    var statusIndex=idx(['Verification_Status','Editorial_Status','Application_Status','Review_Status','Status']);
    for(var r=0;r<values.length;r++){
      var ref=String(display[r][idIndex]||'').trim();
      if(!ref)continue;
      var sheetName=sh.getName();
      var rawType=typeIndex>=0?String(display[r][typeIndex]||'').trim():'';
      var lowerName=sheetName.toLowerCase();
      var categoryName=rawType||sheetName;
      if(lowerName==='opportunity_recommendations')categoryName='Opportunity recommendation';
      else if(lowerName==='volunteer_applications')categoryName='Volunteer application';
      else if(lowerName==='newsletter_subscribers')categoryName='Newsletter signup';
      else if(lowerName==='submission_log'&&!rawType)categoryName='Report';
      var received=dateIndex>=0?values[r][dateIndex]:'';
      var time=received instanceof Date?received.getTime():(received?new Date(received).getTime():0);
      var receivedText='';
      if(time&&!isNaN(time))receivedText=Utilities.formatDate(new Date(time),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm');
      else if(dateIndex>=0)receivedText=String(display[r][dateIndex]||'');
      var title=titleIndex>=0?String(display[r][titleIndex]||'').trim():'';
      var summary=summaryIndex>=0?String(display[r][summaryIndex]||'').trim():'';
      var item={
        ref:ref,category:categoryName,title:title||categoryName+' submission',
        name:nameIndex>=0?String(display[r][nameIndex]||'').trim():'',
        institution:institutionIndex>=0?String(display[r][institutionIndex]||'').trim():'',
        summary:summary.substring(0,1200),status:statusIndex>=0?String(display[r][statusIndex]||'').trim():'Received',
        submittedAt:receivedText,sourceTab:sheetName,sortTime:time||0
      };
      if(!byRef[ref]||sheetName!=='Submission_Log')byRef[ref]=item;
    }
  });
  return Object.keys(byRef).map(function(k){return byRef[k];});
}

function MD19_submissionGroup_(category) {
  var value=String(category||'').toLowerCase();
  if(value.indexOf('opportun')>=0)return 'opportunities';
  if(value.indexOf('volunteer')>=0)return 'volunteers';
  if(value.indexOf('support')>=0||value.indexOf('donat')>=0)return 'support';
  if(value.indexOf('newsletter')>=0||value.indexOf('subscriber')>=0)return 'newsletter';
  if(value.indexOf('report')>=0||value.indexOf('student voice')>=0)return 'reports';
  return 'other';
}

function MD19_getEntry_(user, categoryInput, refInput) {
  var category = String(categoryInput||'').trim().toLowerCase();
  var ref = MD19_clean_(refInput,120);
  if (!ref) throw new Error('Entry reference is missing.');
  var candidates = [];
  if (category.indexOf('volunteer')>=0 || ref.indexOf('VOL-')===0) candidates = [MD19_CONFIG_.VOLUNTEERS_SHEET];
  else if (category.indexOf('opportun')>=0 || ref.indexOf('ROP-')===0 || ref.indexOf('OP-')===0) candidates = [MD19_CONFIG_.OPPORTUNITY_RECOMMENDATIONS_SHEET,MD19_CONFIG_.OPPORTUNITIES_SHEET];
  else if (category.indexOf('newsletter')>=0 || ref.indexOf('NL-')===0) candidates = ['Newsletter_Subscribers'];
  else if (category.indexOf('report')>=0 || category.indexOf('support')>=0 || category.indexOf('donat')>=0) candidates = MD19_CONFIG_.REPORT_SHEET_CANDIDATES;
  else candidates = MD19_CONFIG_.REPORT_SHEET_CANDIDATES.concat([MD19_CONFIG_.OPPORTUNITY_RECOMMENDATIONS_SHEET,MD19_CONFIG_.VOLUNTEERS_SHEET]);
  var ss = MD19_getSpreadsheet_();
  if (category.indexOf('report')>=0 || category.indexOf('support')>=0 || category.indexOf('donat')>=0 || candidates.length===0) {
    ss.getSheets().forEach(function(sh){
      var name=sh.getName();
      if (name.indexOf('Admin_')===0 || (name.indexOf('Newsletter_')===0 && name!=='Newsletter_Subscribers') || name==='NewsletterQueue' || name==='SponsoredBanners' || name==='Newsletter_Campaigns' || name==='Opportunity_Studio') return;
      if (candidates.indexOf(name)<0) candidates.push(name);
    });
  }
  for (var i=0;i<candidates.length;i++) {
    var sheet = ss.getSheetByName(candidates[i]);
    if (!sheet || sheet.getLastRow()<2) continue;
    var width = sheet.getLastColumn();
    var headers = sheet.getRange(1,1,1,width).getDisplayValues()[0].map(function(v){return String(v||'').trim();});
    var idIndex = MD19_findHeaderIndex_(headers,['Reference_Number','Reference','Report_ID','Submission_ID','Application_ID','Recommendation_ID','Opp_ID','Donation_Reference','Subscriber_ID','ID']);
    if (idIndex<0) continue;
    var values = sheet.getRange(2,1,sheet.getLastRow()-1,width).getDisplayValues();
    for (var r=0;r<values.length;r++) {
      if (String(values[r][idIndex]||'').trim() !== ref) continue;
      var fields={};
      headers.forEach(function(h,j){if(h&&!(sheet.getName()==='Newsletter_Subscribers'&&String(h).toLowerCase().replace(/[\s_-]/g,'')==='unsubscribetoken'))fields[h]=values[r][j];});
      MD19_audit_(user.email,'VIEW_SUBMISSION','Submission',ref,{category:category,sheet:sheet.getName()},'Success');
      return {ok:true,category:category,ref:ref,fields:fields};
    }
  }
  throw new Error('The submission could not be found in the connected CMS. Check its reference and the report sheet name.');
}

function MD19_submitTaskDeliverable_(user, data) {
  data=data||{};
  var taskId=MD19_clean_(data.taskId,80);
  var summary=MD19_clean_(data.summary,6000);
  var workUrl=MD19_clean_(data.workUrl,1000);
  var noteToChief=MD19_clean_(data.noteToChief,2000);
  if(!summary&&!workUrl)throw new Error('Enter a work summary or provide a link to the completed work.');
  if(workUrl&&!/^https:\/\/\S+$/i.test(workUrl))throw new Error('The work link must start with https://.');
  var ss=MD19_getSpreadsheet_();
  var tasks=ss.getSheetByName(MD19_CONFIG_.TASKS);
  var taskMap=MD19_headerMap_(tasks,MD19_CONFIG_.TASK_HEADERS);
  var taskRow=MD19_findRowByValue_(tasks,taskMap.Task_ID,taskId);
  if(!taskRow)throw new Error('Task not found.');
  var assignee=String(tasks.getRange(taskRow,taskMap.Assigned_To_Email+1).getDisplayValue()||'').trim().toLowerCase();
  if(user.role!==MD19_CONFIG_.ROLE_CHIEF&&assignee!==user.email)throw new Error('Only the assigned administrator or Chief can submit work for this task.');
  var currentStatus=String(tasks.getRange(taskRow,taskMap.Status+1).getDisplayValue()||'');
  if(currentStatus==='Cancelled')throw new Error('This task is cancelled and cannot receive work.');
  var now=new Date();
  var id='WORK-'+Utilities.getUuid().replace(/-/g,'').substring(0,14).toUpperCase();
  var sheet=ss.getSheetByName(MD19_CONFIG_.DELIVERABLES);
  sheet.appendRow([id,taskId,user.email,now,summary,workUrl,noteToChief]);
  tasks.getRange(taskRow,taskMap.Status+1).setValue('Completed');
  tasks.getRange(taskRow,taskMap.Completed_At+1).setValue(now);
  tasks.getRange(taskRow,taskMap.Updated_At+1).setValue(now);
  MD19_audit_(user.email,'SUBMIT_TASK_DELIVERABLE','Task',taskId,{deliverableId:id,summaryLength:summary.length,hasWorkUrl:!!workUrl,noteLength:noteToChief.length},'Success');
  MD19_notifyChiefs_('Task work submitted',taskId,'Completed work was submitted for task '+taskId+' by '+user.email+'.');
  return {ok:true,message:'Work submitted and recorded. The task is marked Completed.',deliverableId:id};
}

function MD19_markNotificationRead_(user, idInput) {
  var id = MD19_clean_(idInput,100);
  var sheet = MD19_getSpreadsheet_().getSheetByName(MD19_CONFIG_.NOTIFICATIONS);
  var map = MD19_headerMap_(sheet,MD19_CONFIG_.NOTIFICATION_HEADERS);
  var row = MD19_findRowByValue_(sheet,map.Notification_ID,id);
  if (!row) throw new Error('Notification not found.');
  var recipient = String(sheet.getRange(row,map.Recipient_Email+1).getDisplayValue()||'').trim().toLowerCase();
  if (recipient !== user.email) throw new Error('You cannot change another administrator’s notification.');
  sheet.getRange(row,map.Read_At+1).setValue(new Date());
  MD19_audit_(user.email,'MARK_NOTIFICATION_READ','Notification',id,{},'Success');
  return {ok:true};
}

function MD19_retryNotification_(user, idInput) {
  if (user.role !== MD19_CONFIG_.ROLE_CHIEF) throw new Error('Only the Chief Administrator can retry administrative email.');
  var id=MD19_clean_(idInput,100);
  var sheet=MD19_getSpreadsheet_().getSheetByName(MD19_CONFIG_.NOTIFICATIONS);
  var map=MD19_headerMap_(sheet,MD19_CONFIG_.NOTIFICATION_HEADERS);
  var row=MD19_findRowByValue_(sheet,map.Notification_ID,id);
  if(!row)throw new Error('Notification not found.');
  var status=String(sheet.getRange(row,map.Status+1).getDisplayValue()||'');
  if(status!=='Failed')throw new Error('Only a failed notification can be retried.');
  var recipient=String(sheet.getRange(row,map.Recipient_Email+1).getDisplayValue()||'').trim().toLowerCase();
  var type=String(sheet.getRange(row,map.Type+1).getDisplayValue()||'Admin update');
  var message=String(sheet.getRange(row,map.Message+1).getDisplayValue()||'Please sign in to the Admin Portal.');
  if(!MD19_isEmail_(recipient))throw new Error('The notification recipient address is invalid.');
  try {
    MailApp.sendEmail({to:recipient,subject:type+' — The Mobilizer\'s Diary',body:message+'\n\nPortal: '+MD19_CONFIG_.PORTAL_URL,name:MD19_CONFIG_.EMAIL_SENDER_NAME});
    sheet.getRange(row,map.Status+1).setValue('Sent');
    sheet.getRange(row,map.Error+1).setValue('');
    MD19_audit_(user.email,'RETRY_NOTIFICATION','Notification',id,{recipient:recipient},'Success');
    return {ok:true,message:'Notification email sent.'};
  } catch(e) {
    sheet.getRange(row,map.Error+1).setValue(MD19_errorText_(e).substring(0,500));
    MD19_audit_(user.email,'RETRY_NOTIFICATION','Notification',id,{recipient:recipient},'Failed: '+MD19_errorText_(e));
    return {ok:false,message:'Email could not be sent: '+MD19_errorText_(e).substring(0,250)};
  }
}

function MD19_notifyChiefs_(type, ref, message) {
  var sheet = MD19_getSpreadsheet_().getSheetByName(MD19_CONFIG_.USERS);
  var map = MD19_headerMap_(sheet,MD19_CONFIG_.USERS_HEADERS);
  var chiefs = MD19_rowsAsObjects_(sheet,MD19_CONFIG_.USERS_HEADERS).filter(function(r){return r.Role===MD19_CONFIG_.ROLE_CHIEF && r.Access_Status===MD19_CONFIG_.STATUS_ACTIVE;});
  var notices = MD19_getSpreadsheet_().getSheetByName(MD19_CONFIG_.NOTIFICATIONS);
  chiefs.forEach(function(chief){
    var row=notices.getLastRow()+1; notices.appendRow(['NTF-'+Utilities.getUuid().substring(0,12).toUpperCase(),new Date(),chief.Email,type,'Task',ref,message,'Queued','','']);
    try { MailApp.sendEmail({to:chief.Email,subject:type+' — The Mobilizer\'s Diary',body:message+'\n\nPortal: '+MD19_CONFIG_.PORTAL_URL,name:MD19_CONFIG_.EMAIL_SENDER_NAME}); notices.getRange(row,8).setValue('Sent'); }
    catch(e){notices.getRange(row,8).setValue('Failed');notices.getRange(row,9).setValue(MD19_errorText_(e).substring(0,500));}
  });
}

function MD19_findUserByEmail_(email) {
  var ss=MD19_getSpreadsheet_(); var sheet=ss.getSheetByName(MD19_CONFIG_.USERS);
  if(!sheet||sheet.getLastRow()<2) return null;
  var map=MD19_headerMap_(sheet,MD19_CONFIG_.USERS_HEADERS);
  var values=sheet.getRange(2,1,sheet.getLastRow()-1,MD19_CONFIG_.USERS_HEADERS.length).getValues();
  for(var i=0;i<values.length;i++){
    if(String(values[i][map.Email]||'').trim().toLowerCase()!==String(email||'').toLowerCase()) continue;
    var v=values[i];
    return {row:i+2,id:v[map.Admin_ID],name:v[map.Full_Name],email:String(v[map.Email]).trim().toLowerCase(),role:v[map.Role],responsibilities:v[map.Responsibilities],accessStatus:v[map.Access_Status],availability:v[map.Availability_Status],awayUntil:v[map.Away_Until],awayNote:v[map.Away_Note],lastLogin:v[map.Last_Login_At],lastActivity:v[map.Last_Activity_At],sessionVersion:v[map.Session_Version]};
  }
  return null;
}

function MD19_activeAdmins_(sheet,map) {
  var out=[]; if(!sheet||sheet.getLastRow()<2)return out;
  var values=sheet.getRange(2,1,sheet.getLastRow()-1,MD19_CONFIG_.USERS_HEADERS.length).getValues();
  values.forEach(function(r){var email=String(r[map.Email]||'').trim().toLowerCase();if(email&&r[map.Access_Status]===MD19_CONFIG_.STATUS_ACTIVE)out.push({email:email,name:r[map.Full_Name],role:r[map.Role]});});
  return out;
}

function MD19_findChief_(sheet,map) {
  if(sheet.getLastRow()<2)return null;
  var values=sheet.getRange(2,1,sheet.getLastRow()-1,MD19_CONFIG_.USERS_HEADERS.length).getValues();
  for(var i=0;i<values.length;i++)if(values[i][map.Role]===MD19_CONFIG_.ROLE_CHIEF&&values[i][map.Access_Status]===MD19_CONFIG_.STATUS_ACTIVE)return {row:i+2,email:values[i][map.Email]};
  return null;
}

function MD19_incrementSessionVersion_(email) {
  var user=MD19_findUserByEmail_(email); if(!user)return;
  var sheet=MD19_getSpreadsheet_().getSheetByName(MD19_CONFIG_.USERS);
  var map=MD19_headerMap_(sheet,MD19_CONFIG_.USERS_HEADERS);
  sheet.getRange(user.row,map.Session_Version+1).setValue(Number(user.sessionVersion||1)+1);
}

function MD19_setUserField_(rowNumber, header, value) {
  var sheet=MD19_getSpreadsheet_().getSheetByName(MD19_CONFIG_.USERS);
  var map=MD19_headerMap_(sheet,MD19_CONFIG_.USERS_HEADERS);
  if(typeof map[header]!=='number')throw new Error('Missing Admin_Users field '+header+'.');
  sheet.getRange(rowNumber,map[header]+1).setValue(value);
}

function MD19_rowsAsObjects_(sheet, expectedHeaders) {
  if(!sheet||sheet.getLastRow()<2)return [];
  var map=MD19_headerMap_(sheet,expectedHeaders);
  var values=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  return values.map(function(row){var out={};expectedHeaders.forEach(function(h){out[h]=row[map[h]];});return out;}).filter(function(row){return Object.keys(row).some(function(k){return String(row[k]||'').trim();});});
}

function MD19_audit_(actor, action, itemType, itemRef, details, outcome) {
  try {
    var sheet=MD19_getSpreadsheet_().getSheetByName(MD19_CONFIG_.AUDIT);
    if(!sheet) return;
    sheet.appendRow(['AUD-'+Utilities.getUuid().replace(/-/g,'').substring(0,14).toUpperCase(),new Date(),String(actor||'SYSTEM'),String(action||''),String(itemType||''),String(itemRef||''),JSON.stringify(details||{}).substring(0,3000),String(outcome||'Success')]);
  } catch(e) { Logger.log('Admin audit write failed: '+MD19_errorText_(e)); }
}

function MD19_ensureSheet_(ss,name,headers) {
  var sheet=ss.getSheetByName(name);
  if(!sheet){sheet=ss.insertSheet(name);sheet.getRange(1,1,1,headers.length).setValues([headers]);}
  else if(sheet.getLastRow()===0){sheet.getRange(1,1,1,headers.length).setValues([headers]);}
  else {
    var actual=sheet.getRange(1,1,1,headers.length).getDisplayValues()[0].map(function(v){return String(v||'').trim();});
    for(var i=0;i<headers.length;i++)if(actual[i]!==headers[i])throw new Error('The '+name+' tab has unexpected headers in column '+(i+1)+'. No existing data was changed.');
  }
  return sheet;
}

function MD19_ensureDeliverablesSheet_(ss) {
  var name=MD19_CONFIG_.DELIVERABLES;
  var sheet=ss.getSheetByName(name);
  if(!sheet||sheet.getLastRow()===0)return MD19_ensureSheet_(ss,name,MD19_CONFIG_.DELIVERABLE_HEADERS);
  var legacy=MD19_CONFIG_.DELIVERABLE_HEADERS.slice(0,6);
  if(sheet.getLastColumn()<legacy.length)throw new Error('The '+name+' tab has unexpected headers. No existing data was changed.');
  var actual=sheet.getRange(1,1,1,legacy.length).getDisplayValues()[0].map(function(v){return String(v||'').trim();});
  for(var i=0;i<legacy.length;i++)if(actual[i]!==legacy[i])throw new Error('The '+name+' tab has unexpected headers in column '+(i+1)+'. No existing data was changed.');
  var noteHeader=sheet.getLastColumn()>=7?String(sheet.getRange(1,7).getDisplayValue()||'').trim():'';
  if(noteHeader&&noteHeader!=='Note_To_Chief')throw new Error('The '+name+' tab has an unexpected seventh column. No existing data was changed.');
  if(!noteHeader)sheet.getRange(1,7).setValue('Note_To_Chief');
  return sheet;
}

function MD19_styleAdminSheets_(ss) {
  [MD19_CONFIG_.USERS,MD19_CONFIG_.TASKS,MD19_CONFIG_.DELIVERABLES,MD19_CONFIG_.AUDIT,MD19_CONFIG_.NOTIFICATIONS].forEach(function(name){
    var sh=ss.getSheetByName(name); var width=sh.getLastColumn(); sh.setFrozenRows(1); sh.setTabColor(MD19_CONFIG_.BRAND_GOLD);
    sh.getRange(1,1,1,width).setBackground(MD19_CONFIG_.BRAND_GREEN).setFontColor('#FFFFFF').setFontWeight('bold').setWrap(true);
    sh.getRange(1,1,Math.max(sh.getMaxRows(),1),width).setVerticalAlignment('top');
  });
}

function MD19_addAdminDropdowns_(ss) {
  var users=ss.getSheetByName(MD19_CONFIG_.USERS); var rows=Math.max(users.getMaxRows()-1,1);
  users.getRange(2,6,rows,1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Active','Inactive'],true).setAllowInvalid(false).build());
  users.getRange(2,7,rows,1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(MD19_CONFIG_.AVAILABILITY,true).setAllowInvalid(false).build());
  var tasks=ss.getSheetByName(MD19_CONFIG_.TASKS); tasks.getRange(2,10,Math.max(tasks.getMaxRows()-1,1),1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(MD19_CONFIG_.TASK_STATUS,true).setAllowInvalid(false).build());
}

function MD19_getSpreadsheet_() {
  var id=PropertiesService.getScriptProperties().getProperty('CMS_ID');
  if(!id)throw new Error('CMS_ID is missing. Run the existing setupWorkspace() function first.');
  return SpreadsheetApp.openById(id);
}

function MD19_headerMap_(sheet,headers) {
  if(!sheet)throw new Error('Admin sheet is missing. Run MD19_setupAdminPortal() first.');
  var actual=sheet.getRange(1,1,1,Math.max(sheet.getLastColumn(),headers.length)).getDisplayValues()[0].map(function(v){return String(v||'').trim();});
  var map={}; headers.forEach(function(h){var i=actual.indexOf(h);if(i<0)throw new Error('Missing column '+h+' in '+sheet.getName()+'.');map[h]=i;}); return map;
}

function MD19_findRowByValue_(sheet,colIndex,value) {
  if(sheet.getLastRow()<2)return 0;
  var values=sheet.getRange(2,colIndex+1,sheet.getLastRow()-1,1).getDisplayValues();
  for(var i=0;i<values.length;i++)if(String(values[i][0]||'').trim()===String(value||''))return i+2;
  return 0;
}

function MD19_findHeaderIndex_(headers,aliases) {
  var normal=headers.map(function(h){return String(h||'').trim().toLowerCase().replace(/[\s_-]/g,'');});
  for(var i=0;i<aliases.length;i++){var k=String(aliases[i]).toLowerCase().replace(/[\s_-]/g,'');var n=normal.indexOf(k);if(n>=0)return n;}
  return -1;
}

function MD19_hmac_(text,key) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(text,key)).replace(/=+$/,''); }
function MD19_sha256_(text) { var bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(text));return bytes.map(function(b){var v=(b+256)%256;return ('0'+v.toString(16)).slice(-2);}).join(''); }
function MD19_constantTimeEqual_(a,b) { a=String(a||'');b=String(b||'');if(a.length!==b.length)return false;var diff=0;for(var i=0;i<a.length;i++)diff|=a.charCodeAt(i)^b.charCodeAt(i);return diff===0; }
function MD19_isEmail_(v) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(v||'').trim()); }
function MD19_clean_(v,max) { return String(v==null?'':v).replace(/[\x00-\x1F\x7F]/g,' ').trim().substring(0,max); }
function MD19_errorText_(e) { return String(e&&e.message?e.message:e||'Unknown error'); }
function MD19_daysSince_(value) { var t=value?new Date(value).getTime():0;return t?Math.max(0,Math.floor((Date.now()-t)/86400000)):null; }
function MD19_hoursSince_(value) { var t=value?new Date(value).getTime():0;return t?Math.max(0,(Date.now()-t)/3600000):0; }
