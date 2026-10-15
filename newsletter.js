(function () {
  if (typeof window.show !== 'function' || typeof window.api !== 'function') { return; }
  var CONSENT = "I agree to receive The Mobilizer's Diary newsletter at this email address. I can unsubscribe at any time.";
  var ITEMS = ["Founder's Note", 'History Corner: lessons from student leadership', 'Student Voices: challenges and solutions', 'Campus Spotlight: standouts, innovators and creators', 'Podcast and Opportunities', 'An Open Invitation to dialogue'];
  var NOPILL = { 11: 1, 12: 1, 13: 1, 14: 1, 15: 1 };
  var TIMED = { 1: 1, 8: 1, 10: 1 };
  var SAVE = { report: 1, support: 1, donate: 1, subscribe: 1 };
  var cur = 1, pendingPay = '', secs = 0, lastClose = 0, focused = false;
  var testMode = /[?&]nlpop=1/.test(location.search);

  function g(i) { return document.getElementById(i); }
  function mk(t, c, x) { var e = document.createElement(t); if (c) e.className = c; if (x != null) e.textContent = x; return e; }
  function get(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }
  function set(k, x) { try { localStorage.setItem(k, x); } catch (e) {} }
  function del(k) { try { localStorage.removeItem(k); } catch (e) {} }
  function ok(e) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e); }
  function val(i) { return g(i).value.replace(/^\s+|\s+$/g, ''); }
  function subscribed() { return get('md_sub') === '1'; }
  function mask(e) { var p = e.split('@'); return p[0].charAt(0) + '***@' + p[1]; }
  function fillList(id) { var u = g(id); for (var i = 0; i < ITEMS.length; i++) { u.appendChild(mk('li', null, ITEMS[i])); } }

  var st = document.createElement('style');
  st.textContent = '.wrap{padding-bottom:96px!important}' +
    '#nlpill{position:fixed;right:12px;bottom:calc(12px + env(safe-area-inset-bottom,0px));z-index:40;min-height:44px;padding:10px 16px;border:0;border-bottom:3px solid #C59B35;border-radius:22px;background:#123F3B;color:#fff;font:800 12px Montserrat,Arial,sans-serif;letter-spacing:.08em;box-shadow:0 2px 8px rgba(0,0,0,.25);display:none}' +
    '.nllist{margin:8px 0 4px;padding-left:20px;font-size:15px}.nllist li{margin:2px 0}' +
    '.nlfree{display:inline-block;background:#DCEFDD;color:#1B5E20;font:800 11px Montserrat,Arial,sans-serif;letter-spacing:.1em;padding:4px 8px;border-radius:4px;margin-bottom:8px}' +
    '.nllink{background:none;border:0;color:#123F3B;text-decoration:underline;font:600 13px Arial,sans-serif;padding:10px 0}';
  document.head.appendChild(st);

  var wrap = document.querySelector('.wrap'), foot = document.querySelector('.foot');
  function addStep(el) { if (foot && foot.parentNode === wrap) { wrap.insertBefore(el, foot); } else { (wrap || document.body).appendChild(el); } }

  /* pill */
  var pill = mk('button', null, 'Get the Newsletter');
  pill.id = 'nlpill'; pill.type = 'button';
  pill.setAttribute('aria-label', 'Subscribe to the free newsletter');
  pill.onclick = function () { openSignup(); };
  document.body.appendChild(pill);

  /* pop-up (three modes) */
  var pop = mk('div', 'modal');
  pop.id = 'nlpop'; pop.setAttribute('role', 'dialog'); pop.setAttribute('aria-modal', 'true');
  pop.innerHTML = "<div class='mbox'><button class='mx' type='button' id='nl-x' aria-label='Close'>&times;</button><p class='kick' id='nl-k'></p><h2 id='nl-t'></h2><div class='rule'></div><span class='nlfree'>COMPLETELY FREE</span><p id='nl-l'></p><ul class='nllist' id='nl-list'></ul><button class='btn' type='button' id='nl-yes'>Subscribe, It's Free</button><button class='btn ghost' type='button' id='nl-alt'>Support Our Work</button><button class='btn ghost' type='button' id='nl-no'>Maybe Later</button></div>";
  document.body.appendChild(pop);
  fillList('nl-list');

  /* signup page */
  var s14 = mk('div', 'step'); s14.id = 's14';
  s14.innerHTML = "<p class='kick'>Newsletter</p><h2>Subscribe to the Newsletter</h2><div class='rule'></div><span class='nlfree'>COMPLETELY FREE</span><p>Short, clear and written for students. Each issue brings:</p><ul class='nllist' id='nl-list2'></ul>" +
    "<div id='nl-saved' style='display:none'><div class='note'>Use this email, or a different one?<br><b id='nl-mask'></b></div><button class='btn' type='button' id='nl-use'>Use This Email</button><button class='btn ghost' type='button' id='nl-other'>Use a Different Email</button><button class='nllink' type='button' id='nl-forget'>Forget the saved email on this device</button></div>" +
    "<div id='nl-form' style='display:none'><label>Email *</label><input type='email' id='nl-email' autocomplete='email'><label>Name (optional)</label><input type='text' id='nl-name' autocomplete='name'><input type='text' id='nl-hp' class='hp' tabindex='-1' autocomplete='off'><label class='chk'><input type='checkbox' id='nl-c'><span id='nl-ct'></span></label><small>Your subscription is kept separate from any report you have made and from any gift.</small><p class='err' id='nl-err'></p><button class='btn' type='button' id='nl-go'>Subscribe</button><button class='btn ghost' type='button' id='nl-back'>Back</button></div>" +
    "<div id='nl-ok' style='display:none'><h2 id='nl-okt'></h2><p>Look out for your welcome email. Every email carries an unsubscribe link.</p><button class='btn' type='button' id='nl-home'>Back to Home</button><button class='btn ghost' type='button' id='nl-again'>Submit Another Report</button></div>";
  addStep(s14);
  fillList('nl-list2');
  g('nl-ct').textContent = CONSENT;

  /* unsubscribe page */
  var s15 = mk('div', 'step'); s15.id = 's15';
  s15.innerHTML = "<p class='kick'>Newsletter</p><h2>Unsubscribe</h2><div class='rule'></div><p id='un-m'>Tap the button to stop receiving The Mobilizer's Diary newsletter.</p><p class='err' id='un-err'></p><button class='btn' type='button' id='un-go'>Unsubscribe Me</button><button class='btn ghost' type='button' id='un-home'>Back to Home</button>";
  addStep(s15);

  /* landing page button, gift pop-up button, privacy sentence */
  var hub = document.querySelector('#s1 .hubnav');
  if (hub) { var lb = mk('button', 'btn ghost', 'Subscribe to Newsletter'); lb.type = 'button'; lb.onclick = function () { openSignup(); }; hub.appendChild(lb); }
  var gf = g('gift'), gm = document.querySelector('#gift .mbox'), gb = null;
  if (gm) {
    gb = mk('button', 'btn ghost', 'Subscribe to Newsletter (Free)'); gb.type = 'button';
    gb.onclick = function () { window.closeGift(); openSignup(); };
    gm.appendChild(gb);
  }
  var doc = document.querySelector('#s2 .doc');
  if (doc) {
    var pp = mk('p'); pp.appendChild(mk('b', null, 'Newsletter: '));
    pp.appendChild(document.createTextNode('if you choose to subscribe, we keep your email, optional name and a record of your consent. Subscriptions are kept separate from reports and gifts, and every email carries an unsubscribe link.'));
    doc.appendChild(pp);
  }

  /* state helpers */
  function modalOpen() { return (gf && gf.style.display === 'flex') || pop.style.display === 'flex'; }
  function sync() { pill.style.display = (!NOPILL[cur] && !subscribed() && !modalOpen() && !focused) ? 'block' : 'none'; }
  function closeNl() { pop.style.display = 'none'; lastClose = Date.now(); sync(); }
  var MODES = {
    report: ['Free', 'Stay in the loop', 'Thank you for your report. Our newsletter is short, clear and completely free. Each issue brings:'],
    pay: ['Free', 'Thank you. One more thing.', 'Our newsletter is completely free. There is nothing to pay. Each issue brings:'],
    timed: ['Newsletter', "The Mobilizer's Diary Newsletter", 'Short, clear and free for every student. Each issue brings:']
  };
  function openNl(mode) {
    if (subscribed() || modalOpen()) { return; }
    var m = MODES[mode];
    g('nl-k').textContent = m[0]; g('nl-t').textContent = m[1]; g('nl-l').textContent = m[2];
    g('nl-alt').style.display = (mode === 'report') ? 'block' : 'none';
    pop.style.display = 'flex'; sync(); g('nl-yes').focus();
  }
  function showForm() { g('nl-saved').style.display = 'none'; g('nl-form').style.display = 'block'; g('nl-ok').style.display = 'none'; }
  function openSignup() {
    g('nl-err').textContent = ''; g('nl-ok').style.display = 'none'; g('nl-c').checked = false;
    var em = get('md_email');
    if (em && ok(em)) { g('nl-mask').textContent = mask(em); g('nl-saved').style.display = 'block'; g('nl-form').style.display = 'none'; }
    else { showForm(); }
    window.show(14);
  }
  function offerPay() { var r = pendingPay; pendingPay = ''; if (!r) { return; } set('md_nlpay_' + r, '1'); openNl('pay'); }

  /* handlers */
  g('nl-x').onclick = closeNl; g('nl-no').onclick = closeNl;
  g('nl-yes').onclick = function () { closeNl(); openSignup(); };
  g('nl-alt').onclick = function () { closeNl(); window.hubOpen('support', true); };
  pop.onclick = function (e) { if (e.target === pop) { closeNl(); } };
  document.addEventListener('keydown', function (e) { if (e.key === 'Escape') { closeNl(); } });
  g('nl-use').onclick = function () { g('nl-email').value = get('md_email') || ''; showForm(); };
  g('nl-other').onclick = function () { g('nl-email').value = ''; showForm(); };
  g('nl-forget').onclick = function () { del('md_email'); g('nl-email').value = ''; showForm(); };
  g('nl-back').onclick = function () { window.show(1); };
  g('nl-home').onclick = function () { window.show(1); };
  g('nl-again').onclick = function () { location.reload(); };
  g('nl-go').onclick = function () {
    var e = g('nl-err'), btn = g('nl-go'), em = val('nl-email');
    e.textContent = '';
    if (!ok(em)) { e.textContent = 'Please enter a valid email address.'; return; }
    if (!g('nl-c').checked) { e.textContent = 'Please tick the box to subscribe.'; return; }
    btn.disabled = true; btn.textContent = 'Subscribing...';
    window.api('subscribe', { email: em, name: val('nl-name'), consent: true, hp: val('nl-hp') }).then(function (r) {
      set('md_sub', '1'); set('md_email', em);
      g('nl-okt').textContent = (r && r.status === 'already') ? 'You are already subscribed.' : 'You are subscribed.';
      g('nl-form').style.display = 'none'; g('nl-ok').style.display = 'block';
      btn.disabled = false; btn.textContent = 'Subscribe'; sync();
    }, function (er) {
      btn.disabled = false; btn.textContent = 'Subscribe';
      e.textContent = (er && er.message) ? er.message : 'Something went wrong. Please try again.';
    });
  };

  /* wrappers around the existing page functions */
  var _show = window.show;
  window.show = function (n) { _show(n); cur = n; sync(); };

  var _api = window.api;
  window.api = function (a, p) {
    var pr = _api(a, p);
    if (p && p.email && SAVE[a]) {
      pr.then(function () { var m = String(p.email).replace(/^\s+|\s+$/g, ''); if (ok(m)) { set('md_email', m); } }, function () {});
    }
    if (a === 'support' && p && p.type === 'Donor Update Request') {
      pr.then(function () { if (cur === 13 && pendingPay) { setTimeout(offerPay, 1800); } }, function () {});
    }
    return pr;
  };

  var _cg = window.closeGift;
  if (typeof _cg === 'function') { window.closeGift = function () { _cg(); lastClose = Date.now(); sync(); }; }

  var _og = window.openGift;
  if (typeof _og === 'function') {
    window.openGift = function () {
      if (!g('s7') || !g('s7').classList.contains('on')) { return; }
      var n = parseInt(get('md_rot') || '0', 10) || 0;
      set('md_rot', String(n + 1));
      if (!subscribed() && n % 2 === 1) { openNl('report'); return; }
      if (gb) { gb.style.display = subscribed() ? 'none' : 'block'; }
      _og();
    };
  }

  var _ps = window.payShow;
  if (typeof _ps === 'function') {
    window.payShow = function (state, ref) {
      _ps(state, ref);
      if (state !== 'success' || !ref || subscribed() || get('md_nlpay_' + ref)) { return; }
      pendingPay = ref;
      if (get('md_upd_' + ref)) { setTimeout(function () { if (cur === 11) { offerPay(); } }, 2500); }
    };
  }

  var _us = window.updSkip;
  if (typeof _us === 'function') { window.updSkip = function () { _us(); if (pendingPay) { setTimeout(offerPay, 700); } }; }

  /* keep the pill out of the way of typing */
  document.addEventListener('focusin', function (e) {
    var t = e.target && e.target.tagName, ty = e.target && e.target.type;
    if ((t === 'INPUT' && ty !== 'checkbox' && ty !== 'radio') || t === 'TEXTAREA' || t === 'SELECT') { focused = true; sync(); }
  });
  document.addEventListener('focusout', function () { focused = false; setTimeout(sync, 200); });

  /* occasional "what you get" pop-up: 30s on home/about/advertise, once a week, 3 times at most */
  setInterval(function () {
    sync();
    if (!TIMED[cur] || subscribed() || modalOpen() || document.hidden) { return; }
    secs++;
    if (secs < 30 || Date.now() - lastClose < 5000) { return; }
    var n = parseInt(get('md_nlpop_n') || '0', 10) || 0, t = parseInt(get('md_nlpop_t') || '0', 10) || 0;
    if (!testMode && (n >= 3 || Date.now() - t < 604800000)) { return; }
    if (!testMode) { set('md_nlpop_n', String(n + 1)); set('md_nlpop_t', String(Date.now())); }
    secs = 0;
    openNl('timed');
  }, 1000);

  /* unsubscribe link from emails: ?unsubscribe=TOKEN */
  var um = /[?&]unsubscribe=([A-Za-z0-9]+)/.exec(location.search);
  if (um) {
    var tok = um[1];
    g('un-home').onclick = function () { window.show(1); };
    g('un-go').onclick = function () {
      var b = g('un-go'), e = g('un-err');
      e.textContent = ''; b.disabled = true; b.textContent = 'Please wait...';
      window.api('unsubscribe', { token: tok }).then(function () {
        g('un-m').textContent = 'You have been unsubscribed. You will not receive further newsletters.';
        b.style.display = 'none'; del('md_sub');
      }, function (er) {
        b.disabled = false; b.textContent = 'Unsubscribe Me';
        e.textContent = (er && er.message) ? er.message : 'Something went wrong. Please try again.';
      });
    };
    window.show(15);
  }
  sync();
})();
