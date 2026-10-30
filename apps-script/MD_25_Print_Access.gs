/*
 * THE MOBILIZER'S DIARY — iBUILD PRINT ACCESS (backend, phase 1)
 *
 * New file. Needs MD_19 (admin portal), MD_13 (dispatcher) and the existing
 * PAYSTACK_SECRET_KEY script property used by MD_22.
 *
 * Run once: MD25_setup(), then MD25_configure(), then MD25_selfTest().
 *
 * Accounts: email + one-time code (6 digits, 10 minutes, one use).
 * Paid plans: an access code valid for the plan period (stored hashed, never readable).
 * Payments: created and verified on the server with Paystack; the browser never sets a price.
 */
var MD25_ = {
  PLANS: 'Print_Plans', ACCOUNTS: 'Print_Accounts', ACCESS: 'Print_Access', ORDERS: 'Print_Orders',
  PLAN_HEADERS: ['Plan_ID', 'Label', 'Days', 'Price_NGN', 'Active'],
  ACCOUNT_HEADERS: ['Email', 'Created_At', 'Free_Prints_Used', 'Last_Login_At'],
  ACCESS_HEADERS: ['Access_ID', 'Email', 'Plan_ID', 'Plan_Label', 'Code_Hash', 'Code_Salt', 'Starts_At', 'Expires_At', 'Order_Ref', 'Status'],
  ORDER_HEADERS: ['Order_Ref', 'Created_At', 'Email', 'Plan_ID', 'Plan_Label', 'Amount_NGN', 'Status', 'Channel', 'Paid_At', 'Receipt_No', 'Access_ID'],
  DEFAULT_PLANS: [
    ['DAY1', 'One-day access', 1, 0, 'No'],
    ['MONTH1', 'Monthly access', 30, 0, 'No'],
    ['MONTH6', '6-month access', 180, 0, 'No'],
    ['YEAR1', 'Yearly access', 365, 0, 'No'],
    ['LIFETIME', 'Lifetime access', 0, 0, 'No']
  ],
  FREE_PRINT_LIMIT: 5,
  MIN_PRICE: 100,
  SERVICE_NAME: 'iBuild Infra: Concrete Estimator Print Access',
  SENDER_NAME: 'iBuild by The Mobilizer\u2019s Diary',
  TIMEZONE: 'Africa/Lagos',
  OTP_TTL: 600,
  SESSION_DAYS: 7,
  CODE_ALPHABET: 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'
};

/* ---------- one-time setup ---------- */

function MD25_setup() {
  var ss = MD19_getSpreadsheet_();
  MD25_ensure_(ss, MD25_.PLANS, MD25_.PLAN_HEADERS, MD25_.DEFAULT_PLANS);
  MD25_ensure_(ss, MD25_.ACCOUNTS, MD25_.ACCOUNT_HEADERS);
  MD25_ensure_(ss, MD25_.ACCESS, MD25_.ACCESS_HEADERS);
  MD25_ensure_(ss, MD25_.ORDERS, MD25_.ORDER_HEADERS);
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('MD25_SIGNING_SECRET')) props.setProperty('MD25_SIGNING_SECRET', Utilities.getUuid() + Utilities.getUuid());
  Logger.log('Print access ready. Set Price_NGN and Active=Yes for each plan in the Print_Plans tab.');
}

/** Change the address below if ibuild.html is hosted somewhere else, then run this once. */
function MD25_configure() {
  PropertiesService.getScriptProperties().setProperty('IBUILD_PAGE_URL', 'https://mobilizersdiary.com.ng/ibuild.html');
  Logger.log('Return page saved: ' + PropertiesService.getScriptProperties().getProperty('IBUILD_PAGE_URL'));
}

function MD25_selfTest() {
  var props = PropertiesService.getScriptProperties();
  Logger.log('Paystack key set: ' + !!props.getProperty('PAYSTACK_SECRET_KEY'));
  Logger.log('Return page set: ' + (props.getProperty('IBUILD_PAGE_URL') || 'NO - run MD25_configure()'));
  Logger.log('Signing secret set: ' + !!props.getProperty('MD25_SIGNING_SECRET'));
  Logger.log('Plans on sale: ' + JSON.stringify(MD25_plans_(true)));
  Logger.log('Email quota left today: ' + MailApp.getRemainingDailyQuota());
}

/* ---------- routing ---------- */

function MD25_route_(action, d) {
  d = d || {};
  switch (action) {
    case 'printPlans': return { ok: true, plans: MD25_plans_(true) };
    case 'printRequestCode': return MD25_requestCode_(d);
    case 'printVerifyCode': return MD25_verifyCode_(d);
    case 'printAccessLogin': return MD25_accessLogin_(d);
    case 'printStatus': return MD25_status_(d);
    case 'printCheckout': return MD25_checkout_(d);
    case 'printVerifyOrder': return MD25_verifyOrder_(d);
    case 'printRecover': return MD25_recover_(d);
  }
  throw new Error('Unknown request.');
}

function MD25_adminRoute_(action, d) {
  d = d || {};
  var user = MD19_requireAdmin_(d.token);
  var isChief = user.role === MD19_CONFIG_.ROLE_CHIEF;
  if (action === 'adminPrintPlans') return { ok: true, plans: MD25_plans_(false), isChief: isChief };
  if (action === 'adminPrintOrders') {
    var orders = MD25_rows_(MD25_sheet_(MD25_.ORDERS), MD25_.ORDER_HEADERS).reverse().slice(0, 100).map(function (o) {
      return { ref: o.Order_Ref, createdAt: MD25_fmt_(o.Created_At), email: o.Email, plan: o.Plan_Label, amount: Number(o.Amount_NGN) || 0,
        status: o.Status, channel: o.Channel, paidAt: MD25_fmt_(o.Paid_At), receipt: o.Receipt_No };
    });
    return { ok: true, orders: orders };
  }
  if (action === 'adminPrintSavePlan') {
    if (!isChief) throw new Error('Only the Chief Administrator can change plan prices.');
    return MD25_withLock_(function () { return MD25_savePlan_(user, d); });
  }
  throw new Error('Unknown administrative request.');
}

function MD25_savePlan_(user, d) {
  var sh = MD25_sheet_(MD25_.PLANS), rows = MD25_rows_(sh, MD25_.PLAN_HEADERS);
  var id = MD19_clean_(d.planId, 40), row = null;
  rows.forEach(function (r) { if (String(r.Plan_ID) === id) row = r; });
  if (!row) throw new Error('Plan not found.');
  var price = Number(d.price), active = d.active === true || String(d.active).toLowerCase() === 'yes';
  if (!isFinite(price) || Math.floor(price) !== price || price < 0 || price > 100000000) throw new Error('Enter the price as a whole number of naira.');
  if (active && price < MD25_.MIN_PRICE) throw new Error('Set a price of at least \u20A6' + MD25_.MIN_PRICE + ' before putting this plan on sale.');
  var label = MD19_clean_(d.label, 60) || row.Label;
  sh.getRange(row._row, 2).setValue(label);
  sh.getRange(row._row, 4).setValue(price);
  sh.getRange(row._row, 5).setValue(active ? 'Yes' : 'No');
  MD19_audit_(user.email, 'PRINT_PLAN_SAVE', 'PrintPlan', id, { price: price, active: active, label: label }, 'Success');
  return { ok: true, message: 'Plan saved.', plans: MD25_plans_(false) };
}

/* ---------- sign-in ---------- */

function MD25_requestCode_(d) {
  var email = MD25_email_(d.email);
  var generic = { ok: true, message: 'If that address is valid, a sign-in code has been sent.' };
  if (!email) return generic;
  var cache = CacheService.getScriptCache(), h = MD19_sha256_(email).substring(0, 32);
  return MD25_withLock_(function () {
    var count = Number(cache.get('MD25_RATE_' + h) || 0);
    if (count >= 3) return generic;
    cache.put('MD25_RATE_' + h, String(count + 1), 900);
    var digits = Utilities.getUuid().replace(/\D/g, '').substring(0, 6);
    while (digits.length < 6) digits += String(Math.floor(Math.random() * 10));
    var salt = Utilities.getUuid();
    cache.put('MD25_OTP_' + h, JSON.stringify({ salt: salt, hash: MD19_sha256_(email + '|' + digits + '|' + salt), attempts: 0 }), MD25_.OTP_TTL);
    try {
      MailApp.sendEmail({ to: email, subject: 'Your iBuild sign-in code', name: MD25_.SENDER_NAME,
        body: 'Your iBuild sign-in code is:\n\n' + digits + '\n\nIt expires in 10 minutes and works once. If you did not ask for it, ignore this message.' });
    } catch (e) { cache.remove('MD25_OTP_' + h); Logger.log('Print OTP email failed: ' + MD19_errorText_(e)); }
    return generic;
  });
}

function MD25_verifyCode_(d) {
  var email = MD25_email_(d.email), code = String(d.code || '').replace(/\s/g, '');
  var bad = 'The email or code is invalid or expired. Request a new code and try again.';
  if (!email || !/^\d{6}$/.test(code)) throw new Error(bad);
  var cache = CacheService.getScriptCache(), key = 'MD25_OTP_' + MD19_sha256_(email).substring(0, 32);
  var raw = cache.get(key);
  if (!raw) throw new Error(bad);
  var rec = JSON.parse(raw);
  if (!MD19_constantTimeEqual_(MD19_sha256_(email + '|' + code + '|' + rec.salt), rec.hash)) {
    rec.attempts = Number(rec.attempts || 0) + 1;
    if (rec.attempts >= 5) cache.remove(key); else cache.put(key, JSON.stringify(rec), MD25_.OTP_TTL);
    throw new Error(bad);
  }
  cache.remove(key);
  MD25_withLock_(function () { MD25_touchAccount_(email); });
  return { ok: true, token: MD25_token_(email), status: MD25_entitlement_(email) };
}

/** Paid users can sign in with email + access code instead of waiting for an email. */
function MD25_accessLogin_(d) {
  var email = MD25_email_(d.email), code = MD25_normCode_(d.code);
  var bad = 'That email and access code do not match an active plan.';
  var cache = CacheService.getScriptCache(), rk = 'MD25_AL_' + MD19_sha256_(email || 'x').substring(0, 32);
  var tries = Number(cache.get(rk) || 0);
  if (tries >= 8) throw new Error('Too many attempts. Please wait 15 minutes, or use the email sign-in code.');
  cache.put(rk, String(tries + 1), 900);
  if (!email || code.length !== 12) throw new Error(bad);
  var now = new Date(), ok = false;
  MD25_rows_(MD25_sheet_(MD25_.ACCESS), MD25_.ACCESS_HEADERS).forEach(function (a) {
    if (String(a.Email).toLowerCase() !== email || a.Status !== 'Active' || !MD25_live_(a, now)) return;
    if (MD19_constantTimeEqual_(MD19_sha256_(email + '|' + code + '|' + a.Code_Salt), a.Code_Hash)) ok = true;
  });
  if (!ok) throw new Error(bad);
  cache.remove(rk);
  MD25_withLock_(function () { MD25_touchAccount_(email); });
  return { ok: true, token: MD25_token_(email), status: MD25_entitlement_(email) };
}

function MD25_status_(d) {
  var email = MD25_sessionEmail_(d.token);
  return { ok: true, email: email, status: MD25_entitlement_(email) };
}

/* ---------- entitlement (what this person may print) ---------- */

function MD25_entitlement_(email) {
  var now = new Date(), best = null;
  MD25_rows_(MD25_sheet_(MD25_.ACCESS), MD25_.ACCESS_HEADERS).forEach(function (a) {
    if (String(a.Email).toLowerCase() !== email || a.Status !== 'Active' || !MD25_live_(a, now)) return;
    var exp = a.Expires_At instanceof Date ? a.Expires_At.getTime() : Infinity;
    if (!best || exp > best.exp) best = { exp: exp, plan: a.Plan_Label, at: a.Expires_At };
  });
  if (best) return { tier: 'paid', watermark: false, planLabel: best.plan, expiresAt: best.at instanceof Date ? best.at.toISOString() : null };
  var used = MD25_accountRow_(email) ? Number(MD25_accountRow_(email).Free_Prints_Used) || 0 : 0;
  return { tier: 'free', watermark: true, freeLimit: MD25_.FREE_PRINT_LIMIT, freeRemaining: Math.max(0, MD25_.FREE_PRINT_LIMIT - used) };
}

/** Used by the print step (phase 2): checks the allowance and counts one print for free users. */
function MD25_consumePrint_(email) {
  return MD25_withLock_(function () {
    var ent = MD25_entitlement_(email);
    if (ent.tier === 'paid') return ent;
    var acc = MD25_touchAccount_(email), used = Number(acc.Free_Prints_Used) || 0;
    if (used >= MD25_.FREE_PRINT_LIMIT) throw new Error('Your ' + MD25_.FREE_PRINT_LIMIT + ' free prints have been used. Choose a plan to keep printing.');
    MD25_sheet_(MD25_.ACCOUNTS).getRange(acc._row, 3).setValue(used + 1);
    return MD25_entitlement_(email);
  });
}

/* ---------- payments ---------- */

function MD25_checkout_(d) {
  var email = MD25_sessionEmail_(d.token);
  var plan = MD25_plan_(String(d.planId || ''));
  if (!plan) throw new Error('That plan is not available right now.');
  var props = PropertiesService.getScriptProperties();
  var secret = String(props.getProperty('PAYSTACK_SECRET_KEY') || '').trim();
  var pageUrl = String(props.getProperty('IBUILD_PAGE_URL') || '').trim();
  if (!secret) throw new Error('Payments are not configured. Please contact the site administrator.');
  if (!pageUrl) throw new Error('The return page is not configured. Run MD25_configure().');

  var ref = MD25_withLock_(function () {
    var sh = MD25_sheet_(MD25_.ORDERS), year = new Date().getFullYear(), key = 'PRINT_REF_COUNTER_' + year;
    var counter = Number(props.getProperty(key) || 0), re = new RegExp('^IBP-' + year + '-(\\d{4})$');
    MD25_rows_(sh, MD25_.ORDER_HEADERS).forEach(function (o) { var m = String(o.Order_Ref).match(re); if (m) counter = Math.max(counter, Number(m[1])); });
    counter++;
    if (counter > 9999) throw new Error('Order reference limit reached for this year.');
    props.setProperty(key, String(counter));
    var r = 'IBP-' + year + '-' + ('0000' + counter).slice(-4);
    sh.appendRow([r, new Date(), email, plan.id, plan.label, plan.price, 'PENDING', '', '', '', '']);
    return r;
  });

  var res = UrlFetchApp.fetch('https://api.paystack.co/transaction/initialize', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + secret },
    payload: JSON.stringify({
      email: email, amount: plan.price * 100, currency: 'NGN', reference: ref, callback_url: pageUrl,
      metadata: { custom_fields: [
        { display_name: 'Service', variable_name: 'service', value: MD25_.SERVICE_NAME },
        { display_name: 'Plan', variable_name: 'plan', value: plan.label },
        { display_name: 'Order', variable_name: 'order_ref', value: ref }
      ] }
    })
  });
  var out; try { out = JSON.parse(res.getContentText()); } catch (e) { out = null; }
  if (!out || !out.status || !out.data || !out.data.authorization_url) {
    MD25_setOrder_(ref, { Status: 'FAILED', Channel: 'Could not start payment' });
    throw new Error('Payment could not be started. Please try again.');
  }
  return { ok: true, ref: ref, url: out.data.authorization_url, plan: plan.label, amount: plan.price };
}

function MD25_verifyOrder_(d) {
  var ref = String(d.reference || d.ref || '').trim();
  if (!/^IBP-\d{4}-\d{4}$/.test(ref)) return { status: 'unknown' };
  var order = MD25_order_(ref);
  if (!order) return { status: 'unknown' };
  if (order.Status === 'SUCCESS') return { status: 'success', ref: ref, planLabel: order.Plan_Label, receiptNo: order.Receipt_No, alreadyIssued: true };
  var secret = String(PropertiesService.getScriptProperties().getProperty('PAYSTACK_SECRET_KEY') || '').trim();
  if (!secret) throw new Error('Paystack secret key is not configured.');
  var res = UrlFetchApp.fetch('https://api.paystack.co/transaction/verify/' + encodeURIComponent(ref),
    { method: 'get', headers: { Authorization: 'Bearer ' + secret }, muteHttpExceptions: true });
  var out; try { out = JSON.parse(res.getContentText()); } catch (e) { throw new Error('Paystack returned an unreadable response.'); }
  if (res.getResponseCode() < 200 || res.getResponseCode() >= 300 || !out || !out.status) throw new Error('Paystack could not verify this payment.');
  var txn = out.data || {}, st = String(txn.status || '').toLowerCase();
  if (st === 'abandoned' || st === 'failed') { MD25_setOrder_(ref, { Status: st.toUpperCase() }); return { status: st, ref: ref }; }
  if (st !== 'success') return { status: 'pending', ref: ref };
  var valid = String(txn.reference || '') === ref && Number(txn.amount) === Math.round(Number(order.Amount_NGN) * 100) &&
    String(txn.currency || '').toUpperCase() === 'NGN' && String((txn.customer || {}).email || '').trim().toLowerCase() === String(order.Email).toLowerCase();
  if (!valid) { MD25_setOrder_(ref, { Status: 'REVIEW', Channel: 'Verification mismatch; manual review required' }); return { status: 'pending', ref: ref }; }

  return MD25_withLock_(function () {
    var fresh = MD25_order_(ref);
    if (fresh.Status === 'SUCCESS') return { status: 'success', ref: ref, planLabel: fresh.Plan_Label, receiptNo: fresh.Receipt_No, alreadyIssued: true };
    var plan = MD25_planAny_(fresh.Plan_ID) || { id: fresh.Plan_ID, label: fresh.Plan_Label, days: 0 };
    var acc = MD25_issueAccess_(String(fresh.Email).toLowerCase(), plan, ref);
    var receiptNo = 'RCT-' + ref, paidAt = new Date();
    MD25_setOrder_(ref, { Status: 'SUCCESS', Channel: txn.channel || '', Paid_At: paidAt, Receipt_No: receiptNo, Access_ID: acc.id });
    var sent = MD25_emailReceipt_(String(fresh.Email).toLowerCase(), fresh, receiptNo, paidAt, acc, txn.channel || '');
    return { status: 'success', ref: ref, planLabel: fresh.Plan_Label, receiptNo: receiptNo, expiresAt: acc.expires ? acc.expires.toISOString() : null, emailSent: sent };
  });
}

/** Lost code, or paid but never returned to the page: email + payment reference re-checks and re-sends. */
function MD25_recover_(d) {
  var generic = { ok: true, message: 'If those details match a payment, an email has been sent to the address used to pay.' };
  var email = MD25_email_(d.email), ref = String(d.reference || '').trim();
  if (!email || !/^IBP-\d{4}-\d{4}$/.test(ref)) return generic;
  var cache = CacheService.getScriptCache(), rk = 'MD25_RC_' + MD19_sha256_(email).substring(0, 32), n = Number(cache.get(rk) || 0);
  if (n >= 5) return generic;
  cache.put(rk, String(n + 1), 900);
  var order = MD25_order_(ref);
  if (!order || String(order.Email).toLowerCase() !== email) return generic;
  try {
    if (order.Status !== 'SUCCESS') { MD25_verifyOrder_({ reference: ref }); return generic; }
    MD25_withLock_(function () {
      var code = MD25_newCode_(), salt = Utilities.getUuid(), sh = MD25_sheet_(MD25_.ACCESS), found = false;
      MD25_rows_(sh, MD25_.ACCESS_HEADERS).forEach(function (a) {
        if (String(a.Access_ID) !== String(order.Access_ID)) return;
        found = true;
        sh.getRange(a._row, 5, 1, 2).setValues([[MD19_sha256_(email + '|' + code + '|' + salt), salt]]);
      });
      if (!found) return;
      MailApp.sendEmail({ to: email, subject: 'Your new iBuild access code', name: MD25_.SENDER_NAME,
        body: 'You asked for your iBuild access code again.\n\nNew access code: ' + MD25_display_(code) +
          '\n\nPlan: ' + order.Plan_Label + '\nPayment reference: ' + ref + '\n\nYour earlier code no longer works. If you did not ask for this, reply to this email.' });
      MD19_audit_('SYSTEM', 'PRINT_CODE_RECOVERED', 'PrintOrder', ref, { email: email }, 'Success');
    });
  } catch (e) { Logger.log('Print recover failed: ' + MD19_errorText_(e)); }
  return generic;
}

function MD25_issueAccess_(email, plan, orderRef) {
  var start = new Date(), expires = plan.days > 0 ? new Date(start.getTime() + plan.days * 86400000) : '';
  var code = MD25_newCode_(), salt = Utilities.getUuid(), id = 'ACC-' + Utilities.getUuid().replace(/-/g, '').substring(0, 10).toUpperCase();
  MD25_sheet_(MD25_.ACCESS).appendRow([id, email, plan.id, plan.label, MD19_sha256_(email + '|' + code + '|' + salt), salt, start, expires, orderRef, 'Active']);
  MD25_touchAccount_(email);
  return { id: id, code: code, start: start, expires: expires || null };
}

function MD25_emailReceipt_(email, order, receiptNo, paidAt, acc, channel) {
  var amount = '\u20A6' + Number(order.Amount_NGN).toLocaleString('en-US');
  var valid = acc.expires ? 'until ' + MD25_fmt_(acc.expires) : 'with no expiry (lifetime)';
  var lines = [
    'Receipt ' + receiptNo, '',
    'Service: ' + MD25_.SERVICE_NAME, 'Plan: ' + order.Plan_Label, 'Amount paid: ' + amount,
    'Paid on: ' + MD25_fmt_(paidAt), 'Payment reference: ' + order.Order_Ref, channel ? 'Paid by: ' + channel : '', '',
    'Your access code: ' + MD25_display_(acc.code), 'Valid from now ' + valid + '.', '',
    'Keep this email. The code is not stored in readable form, so we cannot read it back to you. If it is lost, use \u201CFind my code\u201D on the iBuild page with this email and the payment reference above.'
  ].filter(function (l, i, a) { return l !== '' || (a[i - 1] !== ''); });
  try {
    MailApp.sendEmail({ to: email, subject: 'iBuild receipt ' + receiptNo + ' \u2014 ' + order.Plan_Label, name: MD25_.SENDER_NAME, body: lines.join('\n') });
    return true;
  } catch (e) { Logger.log('Print receipt email failed: ' + MD19_errorText_(e)); return false; }
}

/* ---------- data helpers ---------- */

function MD25_ensure_(ss, name, headers, seed) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (seed) sh.getRange(2, 1, seed.length, headers.length).setValues(seed);
  } else {
    var actual = sh.getRange(1, 1, 1, headers.length).getDisplayValues()[0].map(function (v) { return String(v || '').trim(); });
    for (var i = 0; i < headers.length; i++) if (actual[i] !== headers[i]) throw new Error('The ' + name + ' tab has unexpected headers in column ' + (i + 1) + '. Nothing was changed.');
  }
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, headers.length).setBackground('#123F3B').setFontColor('#FFFFFF').setFontWeight('bold');
  return sh;
}

function MD25_sheet_(name) {
  var sh = MD19_getSpreadsheet_().getSheetByName(name);
  if (!sh) throw new Error('Run MD25_setup() first.');
  return sh;
}

function MD25_rows_(sh, headers) {
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, headers.length).getValues().map(function (v, i) {
    var o = { _row: i + 2 };
    headers.forEach(function (h, j) { o[h] = v[j]; });
    return o;
  });
}

function MD25_plans_(onlyOnSale) {
  return MD25_rows_(MD25_sheet_(MD25_.PLANS), MD25_.PLAN_HEADERS).map(function (r) {
    var on = /^(yes|true|active|1)$/i.test(String(r.Active).trim());
    return { id: String(r.Plan_ID), label: String(r.Label), days: Number(r.Days) || 0, price: Number(r.Price_NGN) || 0, active: on };
  }).filter(function (p) { return !onlyOnSale || (p.active && p.price >= MD25_.MIN_PRICE); });
}
function MD25_plan_(id) { var m = MD25_plans_(true).filter(function (p) { return p.id === id; }); return m.length ? m[0] : null; }
function MD25_planAny_(id) { var m = MD25_plans_(false).filter(function (p) { return p.id === id; }); return m.length ? m[0] : null; }

function MD25_order_(ref) {
  var m = MD25_rows_(MD25_sheet_(MD25_.ORDERS), MD25_.ORDER_HEADERS).filter(function (o) { return String(o.Order_Ref) === ref; });
  return m.length ? m[0] : null;
}
function MD25_setOrder_(ref, fields) {
  var o = MD25_order_(ref); if (!o) return;
  var sh = MD25_sheet_(MD25_.ORDERS);
  Object.keys(fields).forEach(function (k) { sh.getRange(o._row, MD25_.ORDER_HEADERS.indexOf(k) + 1).setValue(fields[k]); });
}

function MD25_accountRow_(email) {
  var m = MD25_rows_(MD25_sheet_(MD25_.ACCOUNTS), MD25_.ACCOUNT_HEADERS).filter(function (a) { return String(a.Email).toLowerCase() === email; });
  return m.length ? m[0] : null;
}
function MD25_touchAccount_(email) {
  var sh = MD25_sheet_(MD25_.ACCOUNTS), acc = MD25_accountRow_(email);
  if (!acc) { sh.appendRow([email, new Date(), 0, new Date()]); return MD25_accountRow_(email); }
  sh.getRange(acc._row, 4).setValue(new Date());
  return acc;
}

function MD25_live_(a, now) { return !(a.Expires_At instanceof Date) || a.Expires_At.getTime() > now.getTime(); }
function MD25_email_(v) { var e = String(v || '').trim().toLowerCase(); return MD19_isEmail_(e) && e.length <= 200 ? e : ''; }
function MD25_fmt_(v) { return v instanceof Date ? Utilities.formatDate(v, MD25_.TIMEZONE, 'd MMM yyyy, HH:mm') : String(v || ''); }
function MD25_withLock_(fn) { var l = LockService.getScriptLock(); l.waitLock(30000); try { return fn(); } finally { l.releaseLock(); } }

function MD25_newCode_() {
  var hex = Utilities.getUuid().replace(/-/g, ''), out = '';
  for (var i = 0; i < 12; i++) out += MD25_.CODE_ALPHABET.charAt(parseInt(hex.substr(i * 2, 2), 16) % 32);
  return out;
}
function MD25_normCode_(v) { return String(v || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); }
function MD25_display_(code) { return 'IB-' + code.substring(0, 4) + '-' + code.substring(4, 8) + '-' + code.substring(8, 12); }

function MD25_secret_() {
  var s = PropertiesService.getScriptProperties().getProperty('MD25_SIGNING_SECRET');
  if (!s) throw new Error('Run MD25_setup() first.');
  return s;
}
function MD25_token_(email) {
  var enc = Utilities.base64EncodeWebSafe(JSON.stringify({ e: email, exp: Math.floor(Date.now() / 1000) + MD25_.SESSION_DAYS * 86400, n: Utilities.getUuid() })).replace(/=+$/, '');
  return enc + '.' + MD19_hmac_(enc, MD25_secret_());
}
function MD25_sessionEmail_(token) {
  var parts = String(token || '').trim().split('.');
  if (parts.length !== 2 || !MD19_constantTimeEqual_(parts[1], MD19_hmac_(parts[0], MD25_secret_()))) throw new Error('Please sign in again.');
  var p;
  try { p = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString()); } catch (e) { throw new Error('Please sign in again.'); }
  if (!p.e || Number(p.exp) < Math.floor(Date.now() / 1000)) throw new Error('Your session has expired. Please sign in again.');
  return String(p.e).toLowerCase();
}
