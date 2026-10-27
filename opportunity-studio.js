/* Opportunity Studio UI for admin.html. Self-contained: reuses the portal's saved session token. */
(function () {
  'use strict';
  var KEY = 'md_admin_session_v1', cur = null;
  function $(id) { return document.getElementById(id); }
  function show(el, y) { el.classList.toggle('hidden', !y); }
  function msg(t, err) {
    var m = $('oppMessage'); m.textContent = t || ''; m.classList.remove('hidden', 'error');
    if (err) m.classList.add('error'); if (!t) m.classList.add('hidden');
  }
  function api(action, data) {
    if (typeof API_URL !== 'string') return Promise.reject(new Error('API URL is not configured.'));
    data = data || {}; data.token = localStorage.getItem(KEY) || '';
    return fetch(API_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ action: action, data: data }) })
      .then(function (r) { return r.text(); })
      .then(function (raw) {
        var o; try { o = JSON.parse(raw); } catch (e) { throw new Error('Unreadable response from the service.'); }
        if (!o.ok) throw new Error(o.error || 'Request failed.'); return o.result;
      });
  }
  function t(parent, tag, txt, cls) { var e = document.createElement(tag); if (cls) e.className = cls; e.textContent = txt; parent.appendChild(e); return e; }
  function btn(label, secondary, fn) { var b = document.createElement('button'); b.type = 'button'; b.className = 'btn' + (secondary ? ' secondary' : ''); b.textContent = label; b.addEventListener('click', fn); return b; }
  function fail(e) { msg(e.message, true); }

  var FIELDS = { title: 'oppTitle', organisation: 'oppOrg', category: 'oppCategory', deadline: 'oppDeadline', eligibility: 'oppEligibility', applyUrl: 'oppUrl', summary: 'oppSummary' };
  function collect() {
    var f = { templateId: $('oppTemplate').value };
    Object.keys(FIELDS).forEach(function (k) { f[k] = $(FIELDS[k]).value; });
    return f;
  }
  function fill(f) {
    Object.keys(FIELDS).forEach(function (k) { $(FIELDS[k]).value = f[k] || ''; });
    $('oppTemplate').value = f.templateId || 'classic';
  }
  function preview() {
    var p = $('oppPreview'); p.textContent = '';
    p.appendChild(window.MD_renderOpportunity(collect(), $('oppTemplate').value));
  }

  function loadList() {
    msg('Loading recommendations…');
    return api('adminStudioList', { filter: $('oppFilter').value }).then(function (r) {
      msg('');
      var root = $('oppList'); root.textContent = '';
      if (!r.items.length) { t(root, 'p', 'Nothing here for this filter.', 'muted'); return; }
      r.items.forEach(function (i) {
        var box = document.createElement('article'); box.className = 'item';
        t(box, 'strong', i.title || i.ref);
        t(box, 'p', [i.organisation, i.category, i.ref, i.receivedAt].filter(Boolean).join(' · '), 'muted');
        t(box, 'span', i.studioStatus + (i.publishedOppId ? ' · ' + i.publishedOppId : ''), 'pill');
        var a = document.createElement('div'); a.className = 'actions';
        a.appendChild(btn(i.studioStatus === 'New' ? 'Open in studio' : 'Continue', false, function () { openRec(i.ref); }));
        box.appendChild(a); root.appendChild(box);
      });
    }).catch(fail);
  }

  function openRec(ref) {
    return api('adminStudioLoad', { ref: ref }).then(function (r) {
      cur = r; fill(r.fields);
      $('oppEditorHeading').textContent = 'Editing ' + r.ref + ' · ' + r.studioStatus + (r.revision ? ' · revision ' + r.revision : '');
      $('oppOriginal').textContent = 'Boxes are pre-filled from the submission. Edits are saved as a draft; the original is never changed.';
      var published = r.studioStatus === 'Published';
      Object.keys(FIELDS).forEach(function (k) { $(FIELDS[k]).disabled = published; });
      $('oppTemplate').disabled = published;
      $('oppSave').disabled = published; $('oppShortlist').disabled = published;
      $('oppShortlist').textContent = r.studioStatus === 'Shortlisted' ? 'Save and keep shortlisted' : 'Save and shortlist';
      show($('oppPublish'), r.isChief && r.studioStatus === 'Shortlisted');
      show($('oppUnpublish'), r.isChief && published);
      show($('oppEditor'), true); preview();
      $('oppEditor').scrollIntoView({ behavior: 'smooth', block: 'start' });
    }).catch(fail);
  }

  function saveThen(next) {
    return api('adminStudioSave', { ref: cur.ref, fields: collect() }).then(function (r) { return next ? next(r) : r; });
  }
  function refreshAll(message) { msg(message); return openRec(cur.ref).then(loadList); }

  function init() {
    var tpl = $('oppTemplate');
    Object.keys(window.MD_OPP_TEMPLATES).forEach(function (id) {
      var o = document.createElement('option'); o.value = id; o.textContent = window.MD_OPP_TEMPLATES[id].label; tpl.appendChild(o);
    });
    $('oppEditor').addEventListener('input', preview); tpl.addEventListener('change', preview);
    $('oppRefresh').addEventListener('click', loadList); $('oppFilter').addEventListener('change', loadList);
    $('oppClose').addEventListener('click', function () { show($('oppEditor'), false); cur = null; });
    $('oppSave').addEventListener('click', function () { saveThen().then(function (r) { return refreshAll(r.message); }).catch(fail); });
    $('oppShortlist').addEventListener('click', function () {
      saveThen(function () { return api('adminStudioShortlist', { ref: cur.ref, on: true }); }).then(function (r) { return refreshAll(r.message); }).catch(fail);
    });
    $('oppPublish').addEventListener('click', function () {
      if (!window.confirm('Publish this listing to the public Opportunities page?')) return;
      api('adminStudioPublish', { ref: cur.ref }).then(function (r) { return refreshAll(r.message); }).catch(fail);
    });
    $('oppUnpublish').addEventListener('click', function () {
      if (!window.confirm('Hide this listing from the public page?')) return;
      api('adminStudioUnpublish', { ref: cur.ref }).then(function (r) { return refreshAll(r.message); }).catch(fail);
    });
    var app = $('appView');
    function maybe() { if (!app.classList.contains('hidden')) loadList(); }
    new MutationObserver(maybe).observe(app, { attributes: true, attributeFilter: ['class'] });
    maybe();
  }
  if ($('oppStudioPanel') && window.MD_OPP_TEMPLATES) init();
})();
