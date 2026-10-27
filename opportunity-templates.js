/* The Mobilizer's Diary — opportunity card templates.
 * Used by BOTH the admin Opportunity Studio preview and opportunities.html,
 * so what you design is exactly what visitors see.
 * To add a template: add an entry to T below, and its id to MD23_.TEMPLATES in MD_23.
 */
(function () {
  var CSS = '' +
    '.mdo{margin:0 0 14px;background:#FFFEFB;border:1px solid var(--line,#DED8C7);font-family:Georgia,serif;color:var(--ink,#163D39)}' +
    '.mdo h2,.mdo h3{margin:0;color:var(--green,#123F3B);font-family:"Playfair Display",Georgia,serif;line-height:1.25}' +
    '.mdo .m{margin:6px 0;color:var(--muted,#607D78);font:600 13px/1.5 Arial,sans-serif}' +
    '.mdo p{margin:8px 0 0;white-space:pre-line}' +
    '.mdo a.go{display:inline-block;margin-top:12px;padding:11px 15px;background:var(--green,#123F3B);border-bottom:3px solid #A77F20;color:#fff;text-decoration:none;font:800 11px/1.5 Arial,sans-serif;letter-spacing:.06em;text-transform:uppercase}' +
    '.mdo-classic{padding:18px;border-top:4px solid var(--gold,#C59B35)}.mdo-classic h2{font-size:21px}' +
    '.mdo-featured{border-top:0}.mdo-featured .hd{padding:16px 18px;background:var(--green,#123F3B);border-bottom:4px solid var(--gold,#C59B35)}' +
    '.mdo-featured .tag{display:inline-block;margin-bottom:6px;padding:3px 9px;background:var(--gold,#C59B35);color:#123F3B;font:800 10px/1.4 Arial,sans-serif;letter-spacing:.1em;text-transform:uppercase}' +
    '.mdo-featured h2{color:#fff;font-size:23px}.mdo-featured .bd{padding:16px 18px}.mdo-featured a.go{background:var(--gold,#C59B35);color:#123F3B;border-bottom-color:#8b6919}' +
    '.mdo-compact{display:flex;gap:12px;align-items:center;justify-content:space-between;padding:12px 14px;border-left:5px solid var(--gold,#C59B35)}' +
    '.mdo-compact h3{font-size:17px}.mdo-compact a.go{margin:0;padding:8px 12px;white-space:nowrap}';

  function injectStyles() {
    if (document.getElementById('mdo-styles')) return;
    var s = document.createElement('style'); s.id = 'mdo-styles'; s.textContent = CSS; document.head.appendChild(s);
  }
  function el(tag, cls, txt, parent) {
    var e = document.createElement(tag); if (cls) e.className = cls;
    if (txt != null) e.textContent = String(txt);
    if (parent) parent.appendChild(e); return e;
  }
  function safeUrl(v) {
    try { var u = new URL(String(v || '')); return (u.protocol === 'https:' || u.protocol === 'http:') ? u.href : ''; } catch (e) { return ''; }
  }
  function link(parent, item, label) {
    var u = safeUrl(item.applyUrl || item.url); if (!u) return;
    var a = el('a', 'go', label || 'View Official Opportunity', parent);
    a.href = u; a.target = '_blank'; a.rel = 'noopener noreferrer';
  }
  function meta(item) {
    return [item.organisation, item.category, item.deadline ? 'Deadline: ' + item.deadline : ''].filter(Boolean).join(' · ');
  }

  var T = {
    classic: { label: 'Classic card', render: function (i) {
      var c = el('article', 'mdo mdo-classic');
      el('h2', '', i.title || 'Untitled opportunity', c); el('p', 'm', meta(i), c);
      if (i.summary) el('p', '', i.summary, c);
      if (i.eligibility) el('p', '', 'Eligibility: ' + i.eligibility, c);
      link(c, i); return c; } },
    featured: { label: 'Featured (banner header)', render: function (i) {
      var c = el('article', 'mdo mdo-featured'), hd = el('div', 'hd', null, c), bd = el('div', 'bd', null, c);
      if (i.category) el('span', 'tag', i.category, hd);
      el('h2', '', i.title || 'Untitled opportunity', hd);
      el('p', 'm', [i.organisation, i.deadline ? 'Deadline: ' + i.deadline : ''].filter(Boolean).join(' · '), bd);
      if (i.summary) el('p', '', i.summary, bd);
      if (i.eligibility) el('p', '', 'Eligibility: ' + i.eligibility, bd);
      link(bd, i, 'Apply / Learn more'); return c; } },
    compact: { label: 'Compact row', render: function (i) {
      var c = el('article', 'mdo mdo-compact'), left = el('div', '', null, c);
      el('h3', '', i.title || 'Untitled opportunity', left); el('p', 'm', meta(i), left);
      link(c, i, 'View'); return c; } }
  };

  window.MD_OPP_TEMPLATES = T;
  window.MD_renderOpportunity = function (item, templateId) {
    injectStyles(); return (T[templateId] || T.classic).render(item || {});
  };
})();
