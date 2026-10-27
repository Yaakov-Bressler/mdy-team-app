/* MDY Team App — reads and writes the team SharePoint site through Microsoft Graph,
   as the signed-in person. Nothing is stored anywhere else. */
(function () {
  'use strict';

  var C = window.MDY_CONFIG;
  var SCOPES = ['User.Read', 'Sites.ReadWrite.All'];
  var LIGHT = { '#ffa040': 1, '#00bcd4': 1, '#f9e000': 1, '#a9db5e': 1 };
  var BASE = new URL('./', location.href).href;
  var TEST = window.MDY_TEST || null;

  var S = {
    me: null, site: null, drive: null,
    tracker: null, stages: null,
    deadlines: null, feedback: null,
    rootKids: null, folders: {}, docs: {}, users: {},
    trackerM: null, deadlineFilter: 'All', guideMode: 'Spelling', guideQ: '',
    graphicsM: null, graphicsStage: null,
    submit: null
  };
  var msalApp = null;

  /* ---------- small helpers ---------- */
  function $(s, r) { return (r || document).querySelector(s); }
  function $$(s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); }
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function hasHe(s) { return /[֐-׿]/.test(s || ''); }
  function norm(s) {
    return String(s || '').normalize('NFKD')
      .replace(/[֑-ׇ]/g, '').replace(/[̀-ͯ]/g, '')
      .replace(/[‘’ʼ`׳]/g, "'").replace(/[“”״]/g, '"')
      .toLowerCase();
  }
  function stageColor(s) { return C.stageColors[s] || '#8a8f98'; }
  /* the stages a masechta (or a set of masechtos) runs on: its own list from C.workflows, else every choice on the tracker */
  function stagesFor(ms) {
    var list = [].concat(ms || []);
    var wf = (C.workflows || []).filter(function (w) { return list.length && list.every(function (m) { return w.masechtos.indexOf(m) > -1; }); })[0];
    return wf ? wf.stages : (S.stages || Object.keys(C.stageColors));
  }
  function stageChip(s, label) {
    var c = stageColor(s);
    return '<span class="stage" style="background:' + c + ';color:' + (LIGHT[c] ? '#1a1814' : '#fff') + '">' + esc(label || s || '—') + '</span>';
  }
  function stageIn(text) {
    var names = (S.stages || Object.keys(C.stageColors)).slice().sort(function (a, b) { return b.length - a.length; });
    for (var i = 0; i < names.length; i++) {
      if (text === names[i] || text.indexOf(names[i] + ' ') === 0 || text.indexOf(names[i] + '(') === 0) return names[i];
    }
    return null;
  }
  var AMUD_RE = /\b(\d{1,3})\s*([ab])\b(?:\s*\((\d)\))?/gi;
  function amudList(text) {
    var out = [], m;
    AMUD_RE.lastIndex = 0;
    while ((m = AMUD_RE.exec(text))) out.push({ k: parseInt(m[1], 10) * 10 + (m[2].toLowerCase() === 'b' ? 5 : 0), part: m[3] ? parseInt(m[3], 10) : null });
    return out;
  }
  function rowAmud(r) {
    var m = /(\d{1,3})\s*([ab])\b(?:\s*\((\d)\))?/i.exec(r.title || '');
    return m ? { k: parseInt(m[1], 10) * 10 + (m[2].toLowerCase() === 'b' ? 5 : 0), part: m[3] ? parseInt(m[3], 10) : null } : null;
  }
  var ICON = {
    star: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3.5l2.6 5.4 5.9.8-4.3 4.1 1 5.8L12 16.9l-5.2 2.7 1-5.8-4.3-4.1 5.9-.8z"/></svg>',
    chev: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M9 6l6 6-6 6"/></svg>',
    back: '<svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true" style="fill:none;stroke:currentColor;stroke-width:2.4;stroke-linecap:round;stroke-linejoin:round"><path d="M15 6l-6 6 6 6"/></svg>',
    up: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 16V5M6.5 10.5L12 5l5.5 5.5M5 20h14"/></svg>',
    tick: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg>',
    cal: '<svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3.5" y="5" width="17" height="15" rx="2"/><path d="M3.5 10h17M8 3v4M16 3v4"/></svg>',
    pic: '<svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3.5" y="4.5" width="17" height="15" rx="2"/><circle cx="9" cy="10" r="1.8"/><path d="M20.5 16l-5-5-8 8.5"/></svg>',
    pen: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 20l1-4L16 5a2.1 2.1 0 0 1 3 3L8 19z"/></svg>',
    person: '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="8" r="4"/><path d="M4 21c1.5-4 4.5-6 8-6s6.5 2 8 6"/></svg>',
    chat: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 5h16v11H9l-5 4z"/></svg>',
    ask: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 5h16v11H9l-5 4z"/><path d="M12 7.5l1 2.1 2.2.9-2.2.9-1 2.1-1-2.1-2.2-.9 2.2-.9z"/></svg>',
    folder: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 6.5A1.5 1.5 0 0 1 4.5 5H9l2 2.5h8.5A1.5 1.5 0 0 1 21 9v9.5a1.5 1.5 0 0 1-1.5 1.5h-15A1.5 1.5 0 0 1 3 18.5z"/></svg>',
    ext: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M14 4h6v6M20 4l-9 9M18 14v5a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V7a1 1 0 0 1 1-1h5"/></svg>',
    out: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M15 4h4a1 1 0 0 1 1 1v14a1 1 0 0 1-1 1h-4M10 16l-4-4 4-4M6 12h10"/></svg>',
    file: '<svg viewBox="0 0 24 24" aria-hidden="true" style="width:22px;height:22px;fill:none;stroke:#1d3d6b;stroke-width:2;stroke-linecap:round;stroke-linejoin:round"><path d="M14 3H6.5A1.5 1.5 0 0 0 5 4.5v15A1.5 1.5 0 0 0 6.5 21h11a1.5 1.5 0 0 0 1.5-1.5V8z"/><path d="M14 3v5h5"/></svg>'
  };

  function toast(msg) {
    var t = $('#toast');
    t.textContent = msg; t.hidden = false;
    clearTimeout(toast._t);
    toast._t = setTimeout(function () { t.hidden = true; }, 3200);
  }
  function errBox(e, retry) {
    var msg = (e && e.message) || String(e);
    if (e && e.status === 403) msg = 'You don’t have access to this on the site. ' + msg;
    if (e && e.status === 404) msg = 'Not found on the site. ' + msg;
    return '<div class="err">' + esc(msg) + (retry ? '<br><button class="btn ghost" data-act="retry">Try again</button>' : '') + '</div>';
  }

  /* ---------- dates ---------- */
  function ymdIn(d) {
    var p = new Intl.DateTimeFormat('en-CA', { timeZone: C.timeZone, year: 'numeric', month: '2-digit', day: '2-digit' }).format(d);
    return p; // YYYY-MM-DD
  }
  function daysUntil(iso) {
    if (!iso) return null;
    var a = ymdIn(new Date()), b = ymdIn(new Date(iso));
    return Math.round((Date.parse(b + 'T00:00:00Z') - Date.parse(a + 'T00:00:00Z')) / 86400000);
  }
  function fmtDate(iso) {
    if (!iso) return '';
    return new Intl.DateTimeFormat('en-GB', { timeZone: C.timeZone, weekday: 'short', day: 'numeric', month: 'short' }).format(new Date(iso));
  }

  /* ---------- auth ---------- */
  function currentRoute() { return (location.hash || '#/home').slice(1); }
  async function initAuth() {
    if (TEST) return true;
    msalApp = new msal.PublicClientApplication({
      auth: { clientId: C.clientId, authority: 'https://login.microsoftonline.com/' + C.tenantId, redirectUri: BASE, postLogoutRedirectUri: BASE, navigateToLoginRequestUrl: false },
      cache: { cacheLocation: 'localStorage' }
    });
    await msalApp.initialize();
    var res = await msalApp.handleRedirectPromise();
    if (res && res.account) msalApp.setActiveAccount(res.account);
    var acc = msalApp.getActiveAccount() || msalApp.getAllAccounts()[0] || null;
    if (acc) msalApp.setActiveAccount(acc);
    var saved = sessionStorage.getItem('mdy-route');
    if (res && saved) { sessionStorage.removeItem('mdy-route'); history.replaceState(null, '', '#' + saved); }
    return !!acc;
  }
  function signIn() {
    sessionStorage.setItem('mdy-route', currentRoute());
    msalApp.loginRedirect({ scopes: SCOPES, prompt: 'select_account' });
  }
  function signOut() {
    if (TEST) return;
    localStorage.removeItem('mdy-stars');
    msalApp.logoutRedirect({ account: msalApp.getActiveAccount() });
  }
  async function getToken() {
    try {
      var r = await msalApp.acquireTokenSilent({ scopes: SCOPES, account: msalApp.getActiveAccount() });
      return r.accessToken;
    } catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError || /interaction_required|consent_required|login_required|no_tokens_found/.test(e.errorCode || '')) {
        sessionStorage.setItem('mdy-route', currentRoute());
        await msalApp.acquireTokenRedirect({ scopes: SCOPES });
      }
      throw e;
    }
  }

  /* ---------- Graph ---------- */
  async function g(path, opts) {
    opts = opts || {};
    if (TEST) return TEST.g(path, opts);
    var url = /^https:/.test(path) ? path : 'https://graph.microsoft.com/v1.0' + path;
    var headers = Object.assign({ Authorization: 'Bearer ' + (await getToken()) }, opts.headers || {});
    var body = opts.body;
    if (body && !(body instanceof Blob) && !(body instanceof ArrayBuffer) && typeof body !== 'string') {
      body = JSON.stringify(body); headers['Content-Type'] = 'application/json';
    }
    var res = await fetch(url, { method: opts.method || 'GET', headers: headers, body: body });
    if (res.status === 204) return null;
    var text = await res.text(), data = null;
    try { data = text ? JSON.parse(text) : null; } catch (x) { data = text; }
    if (!res.ok) {
      var err = new Error((data && data.error && data.error.message) || ('The site answered ' + res.status));
      err.status = res.status; err.code = data && data.error && data.error.code;
      throw err;
    }
    return data;
  }
  async function gAll(path) {
    var out = [], next = path;
    while (next) {
      var r = await g(next);
      out = out.concat(r.value || []);
      next = r['@odata.nextLink'] || null;
    }
    return out;
  }
  async function siteId() {
    if (!S.site) S.site = await g('/sites/' + C.siteHost + '?$select=id,webUrl');
    return S.site.id;
  }
  async function driveId() {
    if (!S.drive) S.drive = await g('/sites/' + (await siteId()) + '/drive?$select=id,webUrl');
    return S.drive.id;
  }
  async function me() {
    if (!S.me) S.me = await g('/me?$select=displayName,givenName,surname,mail,userPrincipalName');
    return S.me;
  }
  function nameParts() {
    var n = ((S.me && S.me.displayName) || '').replace(/^(R'|R’|Rabbi|Rav|Mrs\.?|Mr\.?|Dr\.?)\s+/i, '').trim();
    return n ? n.split(/\s+/) : [];
  }
  function surname() {
    var s = (S.me && S.me.surname) || nameParts().slice(-1)[0] || '';
    return s.toLowerCase();
  }
  function firstName() { return (S.me && S.me.givenName) || nameParts()[0] || ''; }
  function initials() {
    var p = nameParts();
    if (!p.length) return 'XX';
    return ((p[0][0] || '') + (p.length > 1 ? p[p.length - 1][0] : '')).toUpperCase();
  }

  /* ---------- the Amud Tracker ---------- */
  var MORD = {};
  C.masechtos.forEach(function (m, i) { MORD[m] = i; });
  function guessMasechta(t) { var m = String(t || '').split(/\s+/)[0]; return m || 'Other'; }
  function amudKey(t) {
    var m = /(\d+)\s*([ab])\s*(?:\((\d)\))?/i.exec(t || '');
    if (!m) return 9999;
    return parseInt(m[1], 10) * 10 + (m[2].toLowerCase() === 'b' ? 5 : 0) + (m[3] ? parseInt(m[3], 10) : 0);
  }
  function cmpRows(a, b) {
    var ma = MORD[a.masechta] != null ? MORD[a.masechta] : 99, mb = MORD[b.masechta] != null ? MORD[b.masechta] : 99;
    if (ma !== mb) return ma - mb;
    if (a.masechta !== b.masechta) return a.masechta < b.masechta ? -1 : 1;
    if (a.sort != null && b.sort != null && a.sort !== b.sort) return a.sort - b.sort;
    return amudKey(a.title) - amudKey(b.title);
  }
  async function personName(id) {
    if (!id) return '';
    if (S.users[id] != null) return S.users[id];
    try {
      var r = await g('/sites/' + (await siteId()) + '/lists/' + encodeURIComponent('User Information List') + '/items/' + id + '?$expand=fields($select=Title)');
      S.users[id] = (r.fields && r.fields.Title) || '';
    } catch (e) { S.users[id] = ''; }
    return S.users[id];
  }
  async function fieldText(f, key) {
    var v = f[key];
    if (typeof v === 'string') return v;
    if (Array.isArray(v)) return v.map(function (x) { return x.LookupValue || x.Email || ''; }).join(', ');
    if (v && typeof v === 'object') return v.LookupValue || v.Email || '';
    if (f[key + 'LookupId']) return personName(f[key + 'LookupId']);
    return '';
  }
  async function loadTracker(force) {
    if (S.tracker && !force) return S.tracker;
    var sid = await siteId();
    var res = await Promise.all([
      gAll('/sites/' + sid + '/lists/' + C.lists.tracker + '/items?$expand=fields&$top=500'),
      S.stages ? null : g('/sites/' + sid + '/lists/' + C.lists.tracker + '/columns')
    ]);
    if (res[1]) {
      var st = (res[1].value || []).filter(function (c) { return c.name === 'Stage'; })[0];
      S.stages = st && st.choice ? st.choice.choices : Object.keys(C.stageColors);
    }
    var rows = await Promise.all(res[0].map(async function (it) {
      var f = it.fields || {};
      return {
        id: it.id, f: f,
        title: f.Title || '', masechta: f.Masechta || guessMasechta(f.Title),
        stage: f.Stage || 'Open', workflow: f.Workflow || '',
        writer: await fieldText(f, 'Writer'), editor: await fieldText(f, 'Editor1'),
        notes: f.Notes || '', due: f.Due || null,
        sort: f.SortKey != null && f.SortKey !== '' && !isNaN(+f.SortKey) ? +f.SortKey : null
      };
    }));
    rows.sort(cmpRows);
    S.tracker = rows;
    return rows;
  }
  function stars() { try { return JSON.parse(localStorage.getItem('mdy-stars') || '[]'); } catch (e) { return []; } }
  function isStarred(id) { return stars().indexOf(String(id)) > -1; }
  function toggleStar(id) {
    var s = stars(), i = s.indexOf(String(id));
    if (i > -1) s.splice(i, 1); else s.push(String(id));
    try { localStorage.setItem('mdy-stars', JSON.stringify(s)); } catch (e) { /* ignore */ }
  }
  /* Which tracker rows a deadline covers, read from its Masechta, Amudim and "Done when" columns exactly as the site's
     deadlines core reads them (Claude/Deadlines/deadlines-core.js, VERSION 2026-10-26-01), so "My work" and the site's tally
     agree. A Masechta can be a volume ("Meilah volume" = Meilah, Kinnim, Tamid, Middos) or a list ("Meilah, Kinnim");
     Amudim takes "all", ranges, "all except ...", and terms that name a masechta ("Meilah 2a-10a, Kinnim all"). */
  var DL_VOLUMES = { 'meilah volume': ['Meilah', 'Kinnim', 'Tamid', 'Middos'] };
  var DL_STAGES = {
    New: ['Open', 'Generated', 'Edited', 'Reviewed', 'Mevukar', 'Kramer', 'Post Kramer', 'Edited 2', 'Edited 3', 'Review 2', 'Proofread', 'Check', 'Final'],
    Old: ['Open', 'Written', 'Run', 'Edited', 'Mevukar', 'Kramer', 'Post Kramer', 'Edited 2', 'Mevukar 2', 'Proofread', 'Closed']
  };
  var DL_TARGETS = {
    'Edited (first-round editing)': { New: 'Edited', Old: 'Edited' },
    'Mevukar (content editors)': { New: 'Mevukar', Old: 'Mevukar' },
    "Post Kramer (back from R' Kramer)": { New: 'Post Kramer', Old: 'Post Kramer' },
    'Proofread': { New: 'Proofread', Old: 'Proofread' },
    'Final': { New: 'Final', Old: 'Closed' }
  };
  function dlMasechtos(m) {
    var t = String(m == null ? '' : m).trim();
    if (!t) return [];
    var v = DL_VOLUMES[t.toLowerCase()];
    if (v) return v.slice();
    var out = [];
    t.split(/\s*(?:[,;+&\/]|\band\b)\s*/i).forEach(function (x) { x = x.trim(); if (x && out.indexOf(x) < 0) out.push(x); });
    return out;
  }
  function dlWf(w) { return w === 'Old' ? 'Old' : 'New'; }
  function dlRank(stage, w) {
    if (stage === 'Final' || stage === 'Closed') return 999;
    return DL_STAGES[dlWf(w)].indexOf(stage);
  }
  function dlAmud(s) {
    var m = String(s || '').toLowerCase().match(/(\d+)\s*([ab])(?:\s*\(\s*(\d+)\s*\))?/);
    if (!m) return null;
    return { part: m[3] ? +m[3] : 0, ord: +m[1] * 2 + (m[2] === 'b' ? 1 : 0) };
  }
  function dlSpec(spec) {
    var t = String(spec == null ? '' : spec).trim().toLowerCase().replace(/[\u2013\u2014]/g, '-');
    var parts = t.split(/\bexcept\b|\bbut not\b|\bwithout\b/);
    function terms(str) {
      var out = [];
      (str || '').split(/[,;\n]+/).forEach(function (tok) {
        tok = tok.trim().replace(/^and\s+/, '');
        if (!tok) return;
        var mas = null, mm = tok.match(/^([a-z][a-z'\u2019]*)\s+(?=\d|all\b|everything\b|whole\b|entire\b)/);
        if (mm && !/^(all|everything|whole|entire|and)$/.test(mm[1])) { mas = mm[1]; tok = tok.slice(mm[0].length); }
        function push(x) { if (mas) x.mas = mas; out.push(x); }
        if (/^(all|everything|whole|entire)/.test(tok)) { push({ all: true }); return; }
        var r = tok.split(/\s*(?:-|\bto\b|\bthrough\b)\s*/);
        if (r.length === 2) { var a = dlAmud(r[0]), b = dlAmud(r[1]); if (a && b) { push({ from: a, to: b }); return; } }
        var one = dlAmud(tok);
        if (one) push({ one: one });
      });
      return out;
    }
    var inc = terms(parts[0]);
    if (!inc.length && parts.length > 1) inc = [{ all: true }];
    return { inc: inc, exc: parts.length > 1 ? terms(parts.slice(1).join(',')) : [] };
  }
  function dlHit(term, a, mas) {
    if (term.mas && term.mas !== String(mas || '').toLowerCase()) return false;
    if (term.all) return true;
    if (term.one) return term.one.ord === a.ord && (!term.one.part || term.one.part === a.part);
    if (term.from) {
      var lo = term.from, hi = term.to;
      if (a.ord < lo.ord || a.ord > hi.ord) return false;
      if (a.ord === lo.ord && lo.part && a.part && a.part < lo.part) return false;
      if (a.ord === hi.ord && hi.part && a.part && a.part > hi.part) return false;
      return true;
    }
    return false;
  }
  /* the tracker rows a deadline covers that have not yet reached its stage */
  function dlRows(d, rows) {
    var f = d.f || {}, target = String(f.Target || '').trim();
    if (/^no tally$/i.test(target)) return [];
    var list = dlMasechtos(f.Masechta);
    if (!list.length) { /* older rows with no Masechta: the masechta named in the deadline's text */
      var txt = d.title + ' ' + d.lines.map(function (l) { return l.value; }).join(' ');
      list = C.masechtos.filter(function (m) { return txt.indexOf(m) > -1; }).slice(0, 1);
    }
    if (!list.length) return [];
    var spec = dlSpec(f.Amudim == null || String(f.Amudim).trim() === '' ? 'all' : f.Amudim), named = {};
    spec.inc.forEach(function (t) { if (t.mas) named[t.mas] = 1; });
    var anyNamed = Object.keys(named).length > 0;
    var want = list.map(function (m) { return m.toLowerCase(); }).filter(function (m) { return !anyNamed || named[m]; });
    var stage = DL_TARGETS[target] || null;
    return rows.filter(function (r) {
      var mas = String(r.masechta || '').toLowerCase();
      if (want.indexOf(mas) < 0) return false;
      var a = dlAmud(r.title);
      if (!a) return false;
      if (!spec.inc.some(function (t) { return dlHit(t, a, mas); }) || spec.exc.some(function (t) { return dlHit(t, a, mas); })) return false;
      if (stage) { var w = dlWf(r.workflow); if (dlRank(r.stage, w) >= dlRank(stage[w], w)) return false; }
      return true;
    });
  }
  function markFromDeadlines() {
    S.fromDl = {};
    if (!S.tracker || !S.deadlines) return;
    S.deadlines.filter(function (d) { return (d.days == null || d.days >= 0) && deadlineIsMine(d); }).forEach(function (d) {
      dlRows(d, S.tracker).forEach(function (r) {
        /* an amud under two of your deadlines shows the earlier one */
        if (!S.fromDl[r.id] || (d.due && S.fromDl[r.id].due && d.due < S.fromDl[r.id].due)) S.fromDl[r.id] = d;
      });
    });
  }
  function isMine(r) {
    var sn = surname();
    if (isStarred(r.id)) return true;
    if (S.fromDl && S.fromDl[r.id]) return true;
    if (!sn || sn.length < 3) return false;
    return (r.writer + ' ' + r.editor).toLowerCase().indexOf(sn) > -1;
  }
  function personLine(r) {
    var bits = [];
    if (S.fromDl && S.fromDl[r.id] && !isStarred(r.id)) bits.push('Your deadline: ' + fmtDate(S.fromDl[r.id].due));
    if (r.editor) bits.push('Editor: ' + r.editor);
    else if (r.writer) bits.push('Writer: ' + r.writer);
    if (r.workflow) bits.push(r.workflow + ' workflow');
    return bits.join(' · ');
  }
  function nextStage(s, m) {
    var list = stagesFor(m);
    var i = list.indexOf(s);
    return i > -1 && i < list.length - 1 ? list[i + 1] : s;
  }
  async function moveStage(row, stage) {
    var sid = await siteId();
    await g('/sites/' + sid + '/lists/' + C.lists.tracker + '/items/' + row.id + '/fields', { method: 'PATCH', body: { Stage: stage } });
    row.stage = stage; row.f.Stage = stage;
  }

  /* ---------- deadlines ---------- */
  async function loadDeadlines(force) {
    if (S.deadlines && !force) return S.deadlines;
    var sid = await siteId();
    var res = await Promise.all([
      gAll('/sites/' + sid + '/lists/' + C.lists.deadlines + '/items?$expand=fields&$top=200'),
      g('/sites/' + sid + '/lists/' + C.lists.deadlines + '/columns')
    ]);
    var cols = (res[1].value || []).filter(function (c) {
      return !c.hidden && !c.readOnly && c.name !== 'Title' && c.name !== 'ContentType' && c.name !== 'Attachments';
    });
    var dueCol = cols.filter(function (c) { return c.dateTime || /^due$/i.test(c.displayName); })[0];
    var extra = cols.filter(function (c) { return c !== dueCol && (c.text || c.choice || c.personOrGroup || c.number); });
    var items = await Promise.all(res[0].map(async function (it) {
      var f = it.fields || {}, lines = [];
      for (var i = 0; i < extra.length; i++) {
        var c = extra[i], v = await fieldText(f, c.name);
        if (v === '' && typeof f[c.name] === 'number') v = String(f[c.name]);
        if (!v) continue;
        v = String(v).trim();
        var line = { label: c.displayName, value: v, he: hasHe(v), job: /^job$/i.test(c.displayName) };
        if (/@/.test(v) || /^(on|off|yes|no|true|false)$/i.test(v)) continue;
        var am = amudList(v);
        if (am.length >= 2 && v.replace(AMUD_RE, '').replace(/[\s,;.&and]+/gi, '').length < 6) { line.amudim = am; }
        else if (!line.job && v.length < 40 && stageIn(v)) { line.stage = stageIn(v); }
        lines.push(line);
      }
      var due = dueCol ? f[dueCol.name] : null;
      return { id: it.id, f: f, title: f.Title || '', due: due, days: daysUntil(due), lines: lines };
    }));
    items.sort(function (a, b) { return (a.due || '9') < (b.due || '9') ? -1 : 1; });
    S.deadlines = items;
    return items;
  }
  /* yours when your name is in its Who (Owner) column or its title. Not the Details: they can name other people
     ("Perakim 1-3 have their own rows (R' Langer, R' Ochs)" on an unassigned row). */
  function deadlineIsMine(d) {
    var sn = surname();
    if (!sn) return false;
    var f = d.f || {};
    return (String(f.Owner || '') + ' ' + d.title).toLowerCase().indexOf(sn) > -1;
  }
  function deadlineCard(d) {
    var job = d.lines.filter(function (l) { return l.job; })[0];
    var heb = d.lines.filter(function (l) { return l.he && !l.job; });
    var am = d.lines.filter(function (l) { return l.amudim; })[0];
    var stg = d.lines.filter(function (l) { return l.stage; })[0];
    var seen = {};
    if (job) seen[job.value] = 1;
    var other = d.lines.filter(function (l) {
      if (l.he || l.job || l.amudim || l.stage || seen[l.value]) return false;
      seen[l.value] = 1; return true;
    });
    var facts = [];
    if (stg) facts.push(stageChip(stg.stage, 'Due at ' + stg.stage));
    if (am) facts.push('<span class="tag">' + am.amudim.length + ' amudim</span>');
    var cls = d.days == null ? '' : d.days < 0 ? 'late' : d.days <= 7 ? 'soon' : '';
    var daysTxt = d.days == null ? '–' : Math.abs(d.days);
    var daysLbl = d.days == null ? '' : d.days < 0 ? 'days late' : d.days === 1 ? 'day' : 'days';
    if (d.days === 0) { daysTxt = 'Today'; daysLbl = ''; }
    return '<article class="card dl"><div class="days ' + cls + '"><b>' + esc(daysTxt) + '</b><span>' + esc(daysLbl) + '</span></div><div class="dt">' +
      (job ? '<span class="tag">' + esc(job.value) + '</span>' : '') +
      '<b>' + esc(d.title) + '</b>' +
      other.map(function (l) { return '<span class="clamp">' + esc(l.value) + '</span>'; }).join('') +
      (facts.length ? '<span style="display:flex;gap:6px;flex-wrap:wrap;margin-top:2px">' + facts.join('') + '</span>' : '') +
      '<span style="color:var(--muted)">' + esc(fmtDate(d.due)) + heb.map(function (l) { return ' · <span lang="he" dir="rtl">' + esc(l.value) + '</span>'; }).join('') + '</span>' +
      '</div></article>';
  }

  /* ---------- library ---------- */
  async function rootKids() {
    if (!S.rootKids) S.rootKids = await gAll('/drives/' + (await driveId()) + '/root/children?$select=id,name,folder,webUrl&$top=200');
    return S.rootKids;
  }
  async function kids(itemId, extra) {
    var key = itemId + (extra || '');
    if (!S.folders[key]) {
      S.folders[key] = await gAll('/drives/' + (await driveId()) + '/items/' + itemId + '/children?$top=500' + (extra || '&$select=id,name,folder,file,size,webUrl,lastModifiedDateTime'));
    }
    return S.folders[key];
  }
  async function masechtaFolder(m) {
    var all = await rootKids();
    return all.filter(function (x) { return x.folder && x.name.toLowerCase().indexOf(m.toLowerCase()) === 0; })[0] || null;
  }
  async function subFolder(parent, name) {
    if (!parent) return null;
    var k = await kids(parent.id);
    return k.filter(function (x) { return x.folder && x.name.toLowerCase() === name.toLowerCase(); })[0] || null;
  }
  async function uploadFile(parentId, name, file, onProgress) {
    var d = await driveId();
    var path = '/drives/' + d + '/items/' + parentId + ':/' + encodeURIComponent(name) + ':';
    if (TEST) return TEST.g(path + '/content', { method: 'PUT', body: file });
    if (file.size < 4 * 1024 * 1024) {
      return g(path + '/content?@microsoft.graph.conflictBehavior=fail', { method: 'PUT', body: file, headers: { 'Content-Type': file.type || 'application/octet-stream' } });
    }
    var s = await g(path + '/createUploadSession', { method: 'POST', body: { item: { '@microsoft.graph.conflictBehavior': 'fail' } } });
    var CH = 327680 * 10, done = null;
    for (var start = 0; start < file.size; start += CH) {
      var end = Math.min(start + CH, file.size);
      var r = await fetch(s.uploadUrl, { method: 'PUT', headers: { 'Content-Range': 'bytes ' + start + '-' + (end - 1) + '/' + file.size }, body: file.slice(start, end) });
      if (!r.ok) throw new Error('The upload stopped (' + r.status + '). Try again.');
      if (r.status === 200 || r.status === 201) done = await r.json();
      if (onProgress) onProgress(end / file.size);
    }
    return done;
  }

  /* ---------- Guide documents ---------- */
  var W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  function textOf(node) {
    var s = '';
    (function walk(n) {
      for (var i = 0; i < n.childNodes.length; i++) {
        var c = n.childNodes[i];
        if (c.nodeType !== 1) continue;
        var ln = c.localName;
        if (ln === 't') s += c.textContent;
        else if (ln === 'tab') s += '\t';
        else if (ln === 'br' || ln === 'cr') s += ' ';
        else if (ln === 'del' || ln === 'delText' || ln === 'instrText') continue;
        else walk(c);
      }
    })(node);
    return s;
  }
  function parseDocx(xml) {
    var doc = new DOMParser().parseFromString(xml, 'application/xml');
    var body = doc.getElementsByTagNameNS(W, 'body')[0];
    var out = [], heading = '';
    function styleOf(p) {
      var ps = p.getElementsByTagNameNS(W, 'pStyle')[0];
      return ps ? (ps.getAttributeNS(W, 'val') || ps.getAttribute('w:val') || '') : '';
    }
    function each(container) {
      for (var i = 0; i < container.childNodes.length; i++) {
        var n = container.childNodes[i];
        if (n.nodeType !== 1) continue;
        if (n.localName === 'sdt') { var sc = n.getElementsByTagNameNS(W, 'sdtContent')[0]; if (sc) each(sc); continue; }
        if (n.localName === 'p') {
          var t = textOf(n).replace(/ /g, ' '), st = styleOf(n);
          if (!t.trim()) continue;
          if (/^(heading|title|berschrift)/i.test(st) || /^[1-6]$/.test(st)) { heading = t.trim(); out.push({ kind: 'h', text: heading }); continue; }
          var cells = t.indexOf('\t') > -1 ? t.split(/\t+/).map(function (x) { return x.trim(); }).filter(Boolean) : null;
          out.push({ kind: 'p', text: t.replace(/\t+/g, '   ').trim(), cells: cells && cells.length > 1 ? cells : null, heading: heading });
        } else if (n.localName === 'tbl') {
          var trs = n.getElementsByTagNameNS(W, 'tr');
          for (var r = 0; r < trs.length; r++) {
            var tcs = [];
            for (var k = 0; k < trs[r].childNodes.length; k++) {
              var tc = trs[r].childNodes[k];
              if (tc.nodeType === 1 && tc.localName === 'tc') tcs.push(textOf(tc).replace(/\s+/g, ' ').trim());
            }
            if (tcs.some(Boolean)) out.push({ kind: 'row', cells: tcs.filter(Boolean), text: tcs.filter(Boolean).join(' · '), heading: heading });
          }
        }
      }
    }
    if (body) each(body);
    out.forEach(function (e) { e.n = norm(e.text); });
    return out;
  }
  var GUIDE = [
    { key: 'Spelling', kind: 'doc', match: /translit/i, ph: 'Hebrew or English word', table: true },
    { key: 'Phrases', kind: 'corpus', ph: 'How was this phrase rendered?' },
    { key: 'Headers', kind: 'headers', ph: 'Which header was used?' },
    { key: 'Definitions', kind: 'doc', match: /definition/i, ph: 'Look up a term', table: false },   // plain paragraphs: "term (definition)"
    { key: 'Rulings', kind: 'doc', match: /rulings/i, ph: 'What’s our rule on…', table: false },
    { key: 'Style', kind: 'doc', match: /style guide/i, ph: 'Search the Style Guide', table: false }
  ];

  /* ---------- the Guide page's corpus and headers index (Site Assets/MDYLookup) ---------- */
  var IDB = null;
  function idb() {
    if (IDB) return IDB;
    IDB = new Promise(function (res) {
      try {
        var r = indexedDB.open('mdy-team', 1);
        r.onupgradeneeded = function () { r.result.createObjectStore('files'); };
        r.onsuccess = function () { res(r.result); };
        r.onerror = function () { res(null); };
      } catch (e) { res(null); }
    });
    return IDB;
  }
  async function idbGet(k) {
    var db = await idb(); if (!db) return null;
    return new Promise(function (res) {
      try { var q = db.transaction('files').objectStore('files').get(k); q.onsuccess = function () { res(q.result || null); }; q.onerror = function () { res(null); }; }
      catch (e) { res(null); }
    });
  }
  async function idbPut(k, v) { var db = await idb(); if (!db) return; try { db.transaction('files', 'readwrite').objectStore('files').put(v, k); } catch (e) { /* storage full: fine */ } }
  async function assetsDrive() {
    if (!S.assets) {
      var sid = await siteId();
      // Site Assets is a system library that /drives leaves out: try its known id, then the list by name
      try { S.assets = await g('/drives/' + C.siteAssetsDrive + '?$select=id,name,webUrl'); } catch (e) { S.assets = null; }
      if (!S.assets) {
        try { S.assets = await g('/sites/' + sid + '/lists/' + encodeURIComponent('Site Assets') + '/drive?$select=id,name,webUrl'); } catch (e) { S.assets = null; }
      }
      if (!S.assets) throw new Error('The Site Assets library was not found.');
    }
    return S.assets.id;
  }
  async function assetJson(name, onStatus) {
    if (TEST) return TEST.asset(name);
    var d = await assetsDrive();
    var item = await g('/drives/' + d + '/root:/MDYLookup/' + encodeURIComponent(name));
    var tag = item.eTag || item.lastModifiedDateTime;
    var hit = await idbGet(name);
    if (hit && hit.tag === tag) return hit.data;
    if (onStatus) onStatus('Downloading ' + name.replace(/\.json$/, '') + ' (' + (item.size / 1048576).toFixed(1) + ' MB, once)…');
    var data = await (await fetch(item['@microsoft.graph.downloadUrl'])).json();
    idbPut(name, { tag: tag, data: data });
    return data;
  }
  async function corpusIndex() {
    if (!S.cidx) S.cidx = await assetJson('corpus-index.json');
    return S.cidx;
  }
  async function loadCorpus(m, onStatus) {
    S.corp = S.corp || {};
    if (!S.corp[m]) {
      var e = (await corpusIndex()).masechtos.filter(function (x) { return x.name === m; })[0];
      if (!e) return [];
      var raw = await assetJson(e.file, onStatus);
      S.corp[m] = raw.map(function (r, i) { return { m: m, i: i, daf: r[0], heb: r[1] || '', eng: r[2] || '', rich: Array.isArray(r[3]) ? r[3] : null, part: r[4] || 1 }; });
    }
    return S.corp[m];
  }
  async function loadHeaders(m) {
    S.hdr = S.hdr || {};
    if (!S.hdr[m]) {
      var e = (await corpusIndex()).masechtos.filter(function (x) { return x.name === m; })[0];
      if (!e || !e.hfile) { S.hdr[m] = []; return []; }
      var raw = await assetJson(e.hfile);
      S.hdr[m] = raw.map(function (r) { return { m: m, daf: r[0], text: r[1] || '', rich: Array.isArray(r[2]) ? r[2] : null, kind: r[3], ord: r[4], part: r[5] || 1 }; });
    }
    return S.hdr[m];
  }
  /* English with no Hebrew before it (headings, the words that open a paragraph): lines-<m>.json, the same rows as headers
     ([daf, text, rich, kind, ord, part]); kind o opens its paragraph, s on its own line, t in a table. Site BUILD 2026-10-18-d32bd93. */
  async function loadLines(m) {
    S.lines = S.lines || {};
    if (!S.lines[m]) {
      var e = (await corpusIndex()).masechtos.filter(function (x) { return x.name === m; })[0];
      if (!e || !e.lfile) { S.lines[m] = []; return []; }
      var raw = await assetJson(e.lfile);
      S.lines[m] = raw.map(function (r) { return { m: m, daf: r[0], text: r[1] || '', rich: Array.isArray(r[2]) ? r[2] : null, kind: r[3], ord: r[4], part: r[5] || 1 }; });
    }
    return S.lines[m];
  }
  /* English only: grouped by wording; the exact wording first, then one starting with the query, then the query at a word start, then anywhere; most used first inside each */
  function lineSearch(sets, q, whole) {
    var nq = enNorm(q || '');
    if (hasHe(q) || nq.length < 2) return { groups: [], total: 0 };
    var nk = enKey(q), map = {}, list = [], total = 0;
    var wre = whole ? new RegExp('(^|[^a-z0-9])' + nq.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?![a-z0-9])') : null;
    sets.forEach(function (rows) {
      rows.forEach(function (h) {
        if (h.ne == null) h.ne = enNorm(h.text);
        var at;
        if (wre) { var mm = wre.exec(h.ne); at = mm ? mm.index + mm[1].length : -1; } else at = h.ne.indexOf(nq);
        if (at < 0) return;
        total++;
        var key = enKey(h.text), sc = key === nk ? 3 : at === 0 ? 2 : /[^a-z0-9']/.test(h.ne[at - 1]) ? 1 : 0;
        var gr = map[key];
        if (!gr) { gr = map[key] = { key: key, text: h.text, rich: h.rich, items: [], kinds: { o: 0, s: 0, t: 0 }, sc: sc }; list.push(gr); }
        gr.sc = Math.max(gr.sc, sc); gr.items.push(h);
        if (gr.kinds[h.kind] != null) gr.kinds[h.kind]++;
      });
    });
    list.sort(function (a, b) { return b.sc - a.sc || b.items.length - a.items.length; });
    return { groups: list, total: total };
  }
  /* a long line cut to about 240 characters around the match, at word breaks */
  function lineSnip(text, q) {
    if (text.length <= 260) return { text: text, cut: false };
    var at = norm(text).indexOf(enNorm(q)), a = 0, b = text.length, c;
    if (at < 0) at = 0;
    if (at > 110) { c = text.lastIndexOf(' ', at - 90); if (c > 0) a = c + 1; }
    if (b - at > 240) { c = text.indexOf(' ', Math.max(at + q.length, a + 240)); if (c > 0) b = c; }
    return { text: (a ? '… ' : '') + text.slice(a, b) + (b < text.length ? ' …' : ''), cut: a > 0 || b < text.length };
  }
  function heTok(s) { return norm(s).replace(/[^\u05d0-\u05ea\s]/g, ' ').split(/\s+/).filter(Boolean); }
  function enNorm(s) { return norm(s).replace(/\s+/g, ' ').trim(); }
  function enKey(s) { return enNorm(s).replace(/^[\s.,;:!?"'()\[\]\u2014\u2013-]+|[\s.,;:!?"'()\[\]\u2014\u2013-]+$/g, ''); }
  function seqAt(tokens, s, qt, partialLast) {
    for (var k = 0; k < qt.length; k++) {
      var t = tokens[s + k];
      if (t == null) return false;
      if (k === qt.length - 1 && partialLast) { if (t.indexOf(qt[k]) !== 0) return false; }
      else if (t !== qt[k]) return false;
    }
    return true;
  }
  function richHtml(rich, text, words) {
    if (!rich) return hl(text, words || []);
    return rich.map(function (seg) { var t = hl(seg[0], words || []); return seg[1] ? '<em>' + t + '</em>' : t; }).join('');
  }
  /* whole: whole words only (Ask Claude), rather than completing the last word as it is typed */
  function phraseSearch(sets, q, whole) {
    var groups = {}, list = [], total = 0, CAP = 6000;
    function add(rows, i, span, exact) {
      if (total >= CAP) return;
      total++;
      var parts = rows.slice(i, i + span);
      var eng = parts.map(function (r) { return r.eng; }).join(' ');
      var key = enKey(eng);
      var gr = groups[key];
      if (!gr) {
        gr = groups[key] = { key: key, eng: eng, rich: parts.every(function (r) { return r.rich; }) ? [].concat.apply([], parts.map(function (r, k) { return k ? [[' ', false]].concat(r.rich) : r.rich; })) : null,
          heb: parts.map(function (r) { return r.heb; }).join(' '), hits: [], exact: false };
        list.push(gr);
      }
      if (exact) gr.exact = true;
      gr.hits.push({ m: rows[i].m, i: i, span: span, daf: rows[i].daf, part: rows[i].part });
    }
    if (hasHe(q)) {
      var qt = heTok(q);
      if (!qt.length) return null;
      var n = qt.length, partialLast = whole ? false : n > 1 ? true : qt[0].length >= 2;
      sets.forEach(function (rows) {
        for (var i = 0; i < rows.length && total < CAP; i++) {
          var r = rows[i];
          if (!r.tok) r.tok = heTok(r.heb);
          var found = false;
          for (var st = 0; st <= r.tok.length - n; st++) { if (seqAt(r.tok, st, qt, partialLast)) { found = true; break; } }
          if (found) { add(rows, i, 1, n === r.tok.length && seqAt(r.tok, 0, qt, false)); continue; }
          if (n < 2 || !r.tok.length) continue;
          var joined = r.tok.slice(), span = 1;
          for (var j = i + 1; j < rows.length && span < 5; j++) {
            var nr = rows[j];
            if (nr.daf !== r.daf || nr.part !== r.part) break;
            if (!nr.tok) nr.tok = heTok(nr.heb);
            joined = joined.concat(nr.tok); span++;
            var hit = false;
            for (var s2 = Math.max(0, r.tok.length - n + 1); s2 < r.tok.length; s2++) { if (seqAt(joined, s2, qt, partialLast)) { hit = true; break; } }
            if (hit) { add(rows, i, span, false); break; }
            if (joined.length >= r.tok.length - 1 + n) break;
          }
        }
      });
    } else {
      var nq = enNorm(q);
      if (nq.length < 2) return null;
      var multi = nq.indexOf(' ') > -1;
      var wre = whole ? new RegExp('(^|[^a-z0-9])' + nq.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?![a-z0-9])') : null;
      var find = function (hay) { if (!wre) return hay.indexOf(nq); var m = wre.exec(hay); return m ? m.index + m[1].length : -1; };
      sets.forEach(function (rows) {
        for (var i = 0; i < rows.length && total < CAP; i++) {
          var r = rows[i];
          if (r.ne == null) r.ne = enNorm(r.eng);
          if (find(r.ne) > -1) { add(rows, i, 1, enKey(r.eng) === enKey(q)); continue; }
          if (!multi) continue;
          var nr = rows[i + 1];
          if (!nr || nr.daf !== r.daf || nr.part !== r.part) continue;
          if (nr.ne == null) nr.ne = enNorm(nr.eng);
          var both = r.ne + ' ' + nr.ne, at = find(both);
          if (at > -1 && at < r.ne.length) add(rows, i, 2, false);
        }
      });
    }
    list.sort(function (a, b) { return (b.exact - a.exact) || (b.hits.length - a.hits.length); });
    return { groups: list, total: total, capped: total >= CAP };
  }
  function headerSearch(sets, q) {
    var map = {}, list = [];
    sets.forEach(function (rows) {
      rows.forEach(function (h) {
        var key = enKey(h.text);
        if (!key) return;
        var gr = map[key];
        if (!gr) { gr = map[key] = { key: key, text: h.text, rich: h.rich, items: [], kinds: { o: 0, m: 0, s: 0 } }; list.push(gr); }
        gr.items.push(h);
        if (gr.kinds[h.kind] != null) gr.kinds[h.kind]++;
      });
    });
    var nq = enKey(q || '');
    if (!nq) { list.sort(function (a, b) { return b.items.length - a.items.length; }); return list; }
    var words = nq.split(' '), reWord = new RegExp('\\b' + nq.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\b');
    var scored = [];
    list.forEach(function (gr) {
      var k = gr.key, sc = -1;
      if (k === nq) sc = 0;
      else if (k.indexOf(nq) === 0) sc = 1;
      else if (reWord.test(k)) sc = 2;
      else if (words.every(function (w) { return k.indexOf(w) > -1; })) sc = k.indexOf(nq) > -1 ? 3 : 4;
      if (sc > -1) scored.push({ gr: gr, sc: sc });
    });
    scored.sort(function (a, b) { return a.sc - b.sc || b.gr.items.length - a.gr.items.length; });
    return scored.map(function (x) { return x.gr; });
  }
  function citeChips(hits, act, keyName) {
    var by = [], seen = {};
    hits.forEach(function (h) {
      var k = h.m + ' ' + h.daf + (h.part > 1 ? ' (' + h.part + ')' : '');
      if (!seen[k]) { seen[k] = { label: k, h: h, n: 0 }; by.push(seen[k]); }
      seen[k].n++;
    });
    var open = S.expanded && S.expanded[keyName];
    var shown = open ? by : by.slice(0, 8);
    var html = shown.map(function (c) {
      var h = c.h;
      return '<button class="cite" data-act="' + act + '" data-m="' + esc(h.m) + '" data-i="' + (h.i != null ? h.i : h.ord) + '" data-span="' + (h.span || 1) + '" data-daf="' + esc(h.daf) + '"' + (h.text ? ' data-h="' + esc(h.text) + '"' : '') + (h.kind && act === 'lctx' ? ' data-kind="' + esc(h.kind) + '"' : '') + '>' + esc(c.label) + (c.n > 1 ? ' ×' + c.n : '') + '</button>';
    }).join('');
    if (by.length > shown.length) html += '<button class="cite more" data-act="expand" data-k="' + esc(keyName) + '">+' + (by.length - shown.length) + ' more</button>';
    return '<div class="cites">' + html + '</div>';
  }
  async function resourceFiles() {
    var res = (await rootKids()).filter(function (x) { return x.folder && x.name.toLowerCase() === C.resourcesFolder.toLowerCase(); })[0];
    if (!res) throw new Error('The Resources folder was not found in the library.');
    return (await kids(res.id)).filter(function (x) { return x.file && /\.docx$/i.test(x.name); });
  }
  async function loadDoc(mode) {
    if (S.docs[mode.key]) return S.docs[mode.key];
    var files = (await resourceFiles()).filter(function (f) { return mode.match.test(f.name); });
    if (!files.length) throw new Error('No ' + mode.key.toLowerCase() + ' document in Resources.');
    files.sort(function (a, b) { return (b.lastModifiedDateTime || '') < (a.lastModifiedDateTime || '') ? -1 : 1; });
    var f = files[0], xml;
    if (TEST) xml = await TEST.docXml(f);
    else {
      var item = await g('/drives/' + (await driveId()) + '/items/' + f.id);
      var url = item['@microsoft.graph.downloadUrl'];
      var buf = await (await fetch(url)).arrayBuffer();
      var zip = await JSZip.loadAsync(buf);
      xml = await zip.file('word/document.xml').async('string');
    }
    S.docs[mode.key] = { file: f, entries: parseDocx(xml) };
    return S.docs[mode.key];
  }
  function hl(text, words) {
    var t = esc(text);
    if (!words.length) return t;
    // highlight on the plain (unpointed) form only when the text has no nikkud; otherwise leave as is
    if (/[֑-ׇ]/.test(text)) return t;
    words.forEach(function (w) {
      if (w.length < 2) return;
      var re = new RegExp('(' + w.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ')', 'ig');
      t = t.replace(re, '<mark>$1</mark>');
    });
    return t;
  }
  var SKIP_HEADS = /^(search words|notes to claude|open questions from claude|recent changes|start here)/i;
  function searchDoc(doc, q, mode) {
    var nq = norm(q).trim();
    if (!nq) return [];
    var words = nq.split(/\s+/).filter(Boolean);
    var hits = [];
    doc.entries.forEach(function (e) {
      if (e.kind === 'h') return;
      if (!mode.table && SKIP_HEADS.test(e.heading || '')) return;
      if (mode.table && !(e.cells && e.cells.length > 1) && e.kind !== 'row') return;
      if (!words.every(function (w) { return e.n.indexOf(w) > -1; })) return;
      var score = 3;
      var cells = (e.cells || []).map(norm);
      if (cells.some(function (c) { return c === nq; })) score = 0;
      else if (cells.some(function (c) { return c.indexOf(nq) === 0; })) score = 1;
      else if (e.n.indexOf(nq) > -1) score = 2;
      hits.push({ e: e, score: score });
    });
    hits.sort(function (a, b) { return a.score - b.score; });
    return hits.slice(0, 60);
  }

  /* ---------- the rulebooks search: the same engine as the site's "What's our rule on…?" (MDY Lookup) ----------
     Copied from mdy-lookup build 2026-10-18-45c8746, so Ask Claude finds the same passages in the app as on the site. */
  var RULES = (function () {

    var ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
    function unesc(s) {
      return s.replace(/&(#x[0-9a-f]+|#\d+|\w+);/gi, function (m, g) {
        if (g[0] === '#') return String.fromCodePoint(g[1] === 'x' || g[1] === 'X' ? parseInt(g.slice(2), 16) : parseInt(g.slice(1), 10));
        return ENT[g] != null ? ENT[g] : m;
      });
    }

    var RE_P = /<w:p[ >][\s\S]*?<\/w:p>/g;
    var RE_R = /<w:r[ >][\s\S]*?<\/w:r>/g;
    var RE_T = /<w:t(?:\s[^>]*)?\/>|<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;

    function paraText(p) {
      var s = '', m; RE_T.lastIndex = 0;
      while ((m = RE_T.exec(p))) s += m[1] || '';
      return unesc(s);
    }
    function isItalic(run) {
      var m = /<w:i(?:\s+w:val="([^"]*)")?\s*\/>/.exec(run);
      return !!m && !(m[1] === '0' || m[1] === 'false' || m[1] === 'off');
    }
    function richText(r) { return r.map(function (s) { return s[0]; }).join(''); }


    var FINALS = { 'ך': 'כ', 'ם': 'מ', 'ן': 'נ', 'ף': 'פ', 'ץ': 'צ' };
    function hebWord(w) {
      return w.replace(/[^א-ת]/g, '').replace(/[ךםןףץ]/g, function (c) { return FINALS[c]; });
    }
    function loose(w) { return w.length > 2 ? w[0] + w.slice(1).replace(/[וי]/g, '') : w; }
    function engNorm(s) {
      return String(s || '').normalize('NFKD').replace(/[̀-ͯ]/g, '')
        .replace(/[‘’ʼ`´]/g, "'").replace(/[“”]/g, '"').toLowerCase();
    }

    var RULE_SKIP = /^(start here|notes to claude|open questions from claude|recent changes|search words|introductory notes)$/i;
    var RULE_STOP = {};
    ('a an the of to in on for and or but nor is are was were be been being am do does did done doing we our ours us you your i me my mine ' +
     'it its this that these those what whats which who whom whose when whens where wheres why how hows should would could must shall will ' +
     'there theres their they them he him his she her with without from by as at about into onto than then so if any some all each every ' +
     'rule rules ruling rulings ruled house mdy our use used uses using write writes written wrote writing get gets got go goes going ' +
     'dont doesnt isnt arent wasnt cant wont shouldnt also just only very really please tell know want need way ok okay thing things one ones ' +
     'say ask') .split(' ').forEach(function (w) { RULE_STOP[w] = 1; });
    /* "say" and "ask" are stop words only as the frame of a question ("how do we say…"); as the subject
       they come back through the word groups (said/says, asked/challenged), which look at the raw word. */
    var RULE_FRAME = { say: 1, ask: 1 };

    var HEBC = '\\u0591-\\u05C7\\u05D0-\\u05EA\\u05F0-\\u05F4';
    var RULE_WORD = new RegExp('[A-Za-z\\u00C0-\\u024F0-9]+(?:[’\'\\-][A-Za-z\\u00C0-\\u024F0-9]+)*|[' + HEBC + ']+(?:["״׳\'][' + HEBC + ']+)*', 'g');

    function ruleStem(w) {
      if (w.length <= 3) return w;
      if (/ies$/.test(w) && w.length > 4) w = w.slice(0, -3) + 'y';
      else if (/(sses|shes|ches|xes|zes)$/.test(w)) w = w.slice(0, -2);
      else if (/s$/.test(w) && !/(ss|us|is|os)$/.test(w)) w = w.slice(0, -1);
      if (/ing$/.test(w) && w.length > 5) { w = w.slice(0, -3); if (/([^aeiouls])\1$/.test(w)) w = w.slice(0, -1); }
      else if (/ed$/.test(w) && w.length > 4) { w = w.slice(0, -2); if (/([^aeiouls])\1$/.test(w)) w = w.slice(0, -1); }
      if (/ization$/.test(w)) w = w.slice(0, -7);
      else if (/ation$/.test(w) && w.length > 8) w = w.slice(0, -5);
      else if (/i[sz]e?$/.test(w) && w.length > 7) w = w.replace(/i[sz]e?$/, 'i');
      else if (/ly$/.test(w) && w.length > 5) w = w.slice(0, -2);
      if (/e$/.test(w) && w.length > 4) w = w.slice(0, -1);
      if (/ic?i$/.test(w) && w.length > 6) w = w.replace(/i$/, '');
      return w;
    }
    function ruleNorm(word) {
      if (/[א-ת]/.test(word)) return hebWord(word);
      var w = engNorm(word).replace(/[’'\-]/g, '');
      return /^[0-9]+$/.test(w) ? w : ruleStem(w);
    }
    /* [{n: normalised, w: surface lower-case, s, e}] */
    function ruleToks(text) {
      var out = [], m; RULE_WORD.lastIndex = 0;
      while ((m = RULE_WORD.exec(text))) {
        var n = ruleNorm(m[0]);
        if (n) out.push({ n: n, w: m[0].toLowerCase(), s: m.index, e: m.index + m[0].length });
      }
      return out;
    }

    /* rich text with bold: [[text, italic, bold]] */
    function ruleRich(p) {
      var segs = [], m; RE_R.lastIndex = 0;
      while ((m = RE_R.exec(p))) {
        var run = m[0], t = paraText(run);
        if (/<w:tab\s*\/>/.test(run)) t = t + ' ';
        if (/<w:br\s*\/>/.test(run)) t = t + ' ';
        if (!t) continue;
        var it = isItalic(run), b = /<w:b(?:\s+w:val="(?!0|false|off)[^"]*")?\s*\/>/.test(run);
        var last = segs[segs.length - 1];
        if (last && ((last[1] === it && last[2] === b) || !t.trim())) last[0] += t;
        else segs.push([t, it, b]);
      }
      return segs;
    }

    /* docx -> blocks in order: {k:'p', style, rich, text} or {k:'tbl', rows:[[cellText]]} */
    function ruleBlocks(xml) {
      var body = (/<w:body>([\s\S]*)<\/w:body>/.exec(xml) || [0, xml])[1];
      var re = /<w:tbl[ >][\s\S]*?<\/w:tbl>|<w:p[ >][\s\S]*?<\/w:p>|<w:p\/>/g, m, out = [];
      while ((m = re.exec(body))) {
        var x = m[0];
        if (/^<w:tbl/.test(x)) {
          var rows = (x.match(/<w:tr[ >][\s\S]*?<\/w:tr>/g) || []).map(function (tr) {
            return (tr.match(/<w:tc[ >][\s\S]*?<\/w:tc>/g) || []).map(function (tc) {
              return (tc.match(RE_P) || []).map(function (p) { return richText(ruleRich(p)); }).join(' ').trim();
            });
          });
          out.push({ k: 'tbl', rows: rows });
        } else {
          var st = (/<w:pStyle w:val="([^"]+)"/.exec(x) || [0, ''])[1];
          var rich = ruleRich(x);
          out.push({ k: 'p', style: st, rich: rich, text: richText(rich).replace(/\s+/g, ' ').trim() });
        }
      }
      return out;
    }
    function headLevel(style) {
      if (/^title$/i.test(style)) return 0;
      var h = /heading\s*([1-6])/i.exec(style);
      return h ? +h[1] : -1;
    }

    /* One document -> passages, plus word groups when it carries a Search words table. */
    function rulePassages(xml, doc) {
      var blocks = ruleBlocks(xml), path = [], out = [], groups = [], lead = null, last = null;
      for (var i = 0; i < blocks.length; i++) {
        var b = blocks[i], top = path[0] || '';
        if (b.k === 'tbl') {
          last = null; lead = null;
          if (/^search words$/i.test(top)) {
            b.rows.slice(1).forEach(function (r) {
              var words = [];
              r.forEach(function (c) { c.split(/[,;\/]/).forEach(function (w) { w = w.replace(/^[\s"“”]+|[\s"“”.]+$/g, ''); if (w && w !== '-') words.push(w); }); });
              if (words.length > 1) groups.push(words);
            });
            continue;
          }
          if (RULE_SKIP.test(top)) continue;
          var head = b.rows[0] || [];
          b.rows.slice(1).forEach(function (r) {
            var cells = r.filter(function (c) { return c && c !== '-'; });
            if (!cells.length) return;
            var t = r.map(function (c, j) { return c && c !== '-' ? (head[j] ? head[j] + ': ' : '') + c : ''; }).filter(Boolean).join(' · ');
            out.push({ doc: doc, path: path.slice(), rich: [[t, false, false]], text: t });
          });
          continue;
        }
        var lv = headLevel(b.style);
        if (lv === 0) continue;
        if (lv > 0) { path = path.slice(0, lv - 1); while (path.length < lv - 1) path.push(''); path.push(b.text); lead = null; last = null; continue; }
        if (!b.text || /claudestamp/i.test(b.style)) continue;
        if (RULE_SKIP.test(path[0] || '')) continue;
        if (/example/i.test(b.style) && last) { last.ex = (last.ex || []).concat([b.rich]); continue; }
        var bullet = /list/i.test(b.style);
        var ps = { doc: doc, path: path.slice(), rich: b.rich, text: b.text, lead: bullet ? lead : null };
        out.push(ps); last = ps;
        if (!bullet) lead = /:\s*$/.test(b.text) ? b.text : null;
      }
      return { passages: out, groups: groups };
    }

    function ruleStatus(p) {
      var h = (p.path[0] || '').toLowerCase();
      if (/not settled|open question/.test(h)) return 'open';
      if (/superseded/.test(h)) return 'superseded';
      return 'settled';
    }

    /* Build the index. */
    function ruleIndex(docs) {
      var passages = [], groupsRaw = [];
      docs.forEach(function (d) { passages = passages.concat(d.passages); groupsRaw = groupsRaw.concat(d.groups || []); });
      /* word groups: single words -> group ids; several-word entries -> phrase keys */
      var gOf = {}, phrases = {}, groups = [];
      groupsRaw.forEach(function (words) {
        var gi = groups.length, mem = {};
        words.forEach(function (w) {
          var ts = ruleToks(w).map(function (t) { return t.n; }).filter(function (n) { return n.length > 1 || /[א-ת]/.test(n); });
          if (!ts.length) return;
          var key = ts.join('_');
          if (ts.length > 1) (phrases[ts[0]] = phrases[ts[0]] || []).push(ts);
          mem[key] = 1;
        });
        var keys = Object.keys(mem);
        if (keys.length < 2) return;
        groups.push(keys);
        keys.forEach(function (k) { (gOf[k] = gOf[k] || []).push(gi); });
      });
      var df = {}, total = 0, lzMap = {}, byFirst = {}, surf = {};
      passages.forEach(function (p, i) {
        p.i = i; p.status = ruleStatus(p);
        var all = (p.lead ? p.lead + ' ' : '') + p.text + (p.ex ? ' ' + p.ex.map(richText).join(' ') : '');
        p.all = all;
        var toks = ruleToks(all), seq = toks.map(function (t) { return t.n; });
        toks.forEach(function (t) { var c = surf[t.n] || (surf[t.n] = {}); c[t.w] = (c[t.w] || 0) + 1; });
        /* add phrase keys where a several-word entry occurs */
        var extra = [];
        for (var k = 0; k < seq.length; k++) {
          var cand = phrases[seq[k]]; if (!cand) continue;
          cand.forEach(function (ph) {
            for (var j = 0; j < ph.length; j++) if (seq[k + j] !== ph[j]) return;
            extra.push(ph.join('_'));
          });
        }
        p.seq = seq;
        p.tf = {};
        seq.concat(extra).forEach(function (n) { p.tf[n] = (p.tf[n] || 0) + 1; });
        p.len = seq.length || 1; total += p.len;
        p.head = {};
        ruleToks(p.path.join(' ')).forEach(function (t) { p.head[t.n] = 1; });
        var bold = (p.rich || []).filter(function (s) { return s[2]; }).map(function (s) { return s[0]; }).join(' ');
        ruleToks(bold).forEach(function (t) { p.head[t.n] = 1; });
        Object.keys(p.tf).forEach(function (n) { df[n] = (df[n] || 0) + 1; });
        Object.keys(p.head).forEach(function (n) { if (!p.tf[n]) df[n] = (df[n] || 0) + 1; });
      });
      Object.keys(df).forEach(function (n) {
        if (/[א-ת]/.test(n)) { var lk = loose(n); (lzMap[lk] = lzMap[lk] || []).push(n); }
        else (byFirst[n[0]] = byFirst[n[0]] || []).push(n);
      });
      return { passages: passages, df: df, N: passages.length || 1, avg: total / (passages.length || 1),
        groups: groups, gOf: gOf, phrases: phrases, lz: lzMap, byFirst: byFirst,
        surf: function (n) { return n.split('_').map(function (x) { var c = surf[x]; if (!c) return x; return Object.keys(c).sort(function (a, b) { return c[b] - c[a]; })[0]; }).join(' '); } };
    }

    function osa(a, b, max) {
      if (Math.abs(a.length - b.length) > max) return max + 1;
      var d = [], i, j;
      for (i = 0; i <= a.length; i++) { d[i] = [i]; }
      for (j = 0; j <= b.length; j++) d[0][j] = j;
      for (i = 1; i <= a.length; i++) {
        var rowMin = max + 1;
        for (j = 1; j <= b.length; j++) {
          var c = a[i - 1] === b[j - 1] ? 0 : 1;
          d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + c);
          if (i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) d[i][j] = Math.min(d[i][j], d[i - 2][j - 2] + 1);
          if (d[i][j] < rowMin) rowMin = d[i][j];
        }
        if (rowMin > max) return max + 1;
      }
      return d[a.length][b.length];
    }

    var PREFIX_HEB = /^[ובדהכלמש]/;
    /* Each query word -> {variant: weight}. */
    function ruleVariants(ix, n, raw, isLast) {
      var v = {}, heb = /[א-ת]/.test(n);
      function add(k, w, why) { if (!v[k] || v[k].w < w) v[k] = { w: w, why: why }; }
      if (ix.df[n]) add(n, 1, '');
      if (heb) {
        var s = n, cut = 0;
        while (PREFIX_HEB.test(s) && s.length > 2 && cut < 3) { s = s.slice(1); cut++; if (ix.df[s]) add(s, 0.85, ''); }
        (ix.lz[loose(n)] || []).forEach(function (k) { add(k, 0.85, ''); });
      } else if (!ix.df[n] && n.length >= 4) {
        var max = n.length >= 8 ? 2 : 1;
        (ix.byFirst[n[0]] || []).forEach(function (k) {
          var dd = osa(n, k, max);
          if (dd <= max) add(k, dd === 1 ? 0.7 : 0.55, 'spelling');
          else if (k.length >= 7 && n.indexOf(k) === 0 && n.length - k.length <= 5) add(k, 0.6, 'spelling');
        });
      }
      if (isLast && n.length >= 3 && !heb) {
        var pre = (ix.byFirst[n[0]] || []).filter(function (k) { return k.length > n.length && k.indexOf(n) === 0; });
        pre.sort(function (a, b) { return ix.df[b] - ix.df[a]; });
        pre.slice(0, 25).forEach(function (k) { add(k, 0.6, ''); });
      }
      /* word groups, from the word itself and from the raw form */
      var seeds = Object.keys(v).concat([n, ruleStem(engNorm(raw).replace(/[’'\-]/g, ''))]);
      seeds.forEach(function (k) {
        (ix.gOf[k] || []).forEach(function (gi) {
          ix.groups[gi].forEach(function (m) { if (m !== k && ix.df[m]) add(m, 0.8 * (v[k] ? v[k].w : 1), 'group'); });
        });
      });
      return v;
    }

    function ruleQuery(ix, q) {
      var quoted = [], m, re = /["“”]([^"“”]+)["“”]/g;
      while ((m = re.exec(q))) quoted.push(m[1]);
      var toks = ruleToks(q), endsInWord = /[A-Za-zא-ת]$/.test(q);
      var terms = [];
      /* several-word group entries typed in the query ("em dash", "Beis Hamikdash") */
      var seq = toks.map(function (t) { return t.n; }), used = {};
      for (var k = 0; k < seq.length; k++) {
        (ix.phrases[seq[k]] || []).forEach(function (ph) {
          for (var j = 0; j < ph.length; j++) if (seq[k + j] !== ph[j]) return;
          var key = ph.join('_');
          if (ix.gOf[key]) { terms.push({ n: key, raw: ph.join(' '), phrase: true }); for (j = 0; j < ph.length; j++) used[k + j] = 1; }
        });
      }
      var isQ = quoted.length > 0;
      toks.forEach(function (t, i) {
        if (used[i]) return;
        var inQuote = quoted.some(function (qq) { return qq.toLowerCase().indexOf(t.w) >= 0; });
        if (!inQuote && RULE_STOP[t.w.replace(/[’'\-]/g, '')] && !(RULE_FRAME[t.w] && toks.length === 1)) return;
        if (!inQuote && t.w.length < 2 && !/[א-ת]/.test(t.w)) return;
        terms.push({ n: t.n, raw: t.w, must: inQuote, last: i === toks.length - 1 && endsInWord && !isQ });
      });
      var seen = {};
      terms = terms.filter(function (t) { if (seen[t.n]) return false; seen[t.n] = 1; return true; });
      terms.forEach(function (t) {
        t.v = ruleVariants(ix, t.n, t.raw, t.last);
        if (t.must) { var keep = {}; if (t.v[t.n]) keep[t.n] = t.v[t.n]; t.v = keep; }
      });
      return terms;
    }

    function searchRules(ix, q, limit) {
      var terms = ruleQuery(ix, q);
      if (!terms.length) return { hits: [], terms: terms };
      var k1 = 1.2, b = 0.6, N = ix.N;
      function idf(n) { var d = ix.df[n] || 0; return Math.log(1 + (N - d + 0.5) / (d + 0.5)); }
      var STATUS = { settled: 1, open: 0.8, superseded: 0.5 };
      var own = {}; terms.forEach(function (t) { own[t.n] = 1; });
      var hits = [];
      ix.passages.forEach(function (p) {
        var sc = 0, matched = 0, marks = {}, mustFail = false, best = [];
        terms.forEach(function (t, ti) {
          var top = 0, topK = null;
          Object.keys(t.v).forEach(function (k) {
            if (k !== t.n && own[k]) return;          /* another word of the question is this word itself */
            var tf = p.tf[k] || 0, h = p.head[k] ? 1 : 0;
            if (!tf && !h) return;
            var w = (k !== t.n && ix.df[t.n]) ? Math.min(idf(k), idf(t.n)) : idf(k);   /* a synonym counts no more than the word typed */
            var s = (tf ? w * tf * (k1 + 1) / (tf + k1 * (1 - b + b * p.len / ix.avg)) : 0) + h * 0.45 * w;
            s *= t.v[k].w;
            if (s > top) { top = s; topK = k; }
            marks[k] = 1;
          });
          if (top > 0) { matched++; sc += top; best[ti] = topK; }
          else if (t.must) mustFail = true;
        });
        if (!matched || mustFail) return;
        /* adjacent query words side by side in the passage */
        for (var i = 0; i + 1 < terms.length; i++) {
          var a = best[i], c = best[i + 1];
          if (!a || !c || a.indexOf('_') >= 0 || c.indexOf('_') >= 0) continue;
          for (var j = 0; j + 1 < p.seq.length; j++) if (p.seq[j] === a && p.seq[j + 1] === c) { sc += 0.5 * (idf(a) + idf(c)) / 2; break; }
        }
        var cover = matched / terms.length;
        sc *= Math.pow(cover, terms.length > 1 ? 2 : 1) * (p.doc.weight || 1) * (STATUS[p.status] || 1);
        hits.push({ p: p, score: sc, marks: marks, cover: cover });
      });
      hits.sort(function (x, y) { return y.score - x.score; });
      /* when some passages match every word, drop the ones that match only part of the question */
      if (terms.length > 1 && hits.length && hits[0].cover === 1) {
        var full = hits.filter(function (h) { return h.cover === 1; }).length;
        if (full >= 3) hits = hits.filter(function (h) { return h.cover === 1 || h.score > hits[0].score * 0.5; });
      }
      var also = {};
      terms.forEach(function (t) { Object.keys(t.v).forEach(function (k) { if (own[k]) return; if (t.v[k].why === 'group' || t.v[k].why === 'spelling') also[ix.surf(k)] = t.v[k].why; }); });
      return { hits: hits.slice(0, limit || hits.length), total: hits.length, terms: terms, also: also };
    }
    return { passages: rulePassages, index: ruleIndex, search: searchRules, richText: richText };
  })();

  /* ---------- Ask Chaim Levy (was Ask Claude): the AI assistant ----------
     The same assistant as the site's Ask Claude page. Claude runs behind the team's relay (C.askRelay), which
     holds the API key and Claude's instructions, and lets in only people signed in with an Aleh Zayis account.
     The app runs the four searches Claude asks for, over what the Guide already reads: the rulebooks, the
     delivered corpus and its headers, and the transliteration and definitions lists. Nothing about people,
     contacts, the tracker or deadlines is ever sent. */
  var ASK_MAX_ROUNDS = 16;
  var ASK_GRACE_ROUNDS = 2;
  var ASK_BUDGET_NOTE = 'You have reached the search limit for this question. Do not call any more tools. Write your answer now from what you have already found. Where something could not be confirmed, say so briefly and say what the editor could check.';
  var ASK_LABEL = { search_rules: 'Searching the rulebooks', search_corpus: 'Searching the delivered volumes',
    search_headers: 'Searching the headers', lookup_terms: 'Checking the transliteration and definitions lists' };
  var ASK_TRIES = ['When do we write “may” and when “can”?', 'How was תא שמע rendered in each masechta?',
    'Which lead-ins do we use for a challenge, and how often?', 'Check this sentence against house style: …'];
  var RULE_SOURCES = [
    { key: 'rulings', label: 'House Rulings', res: /^house\s*rulings/i, weight: 1.15 },
    { key: 'style', label: 'Style Guide', res: /style\s*guide/i, weight: 1.1 },
    { key: 'proof', label: 'Proofreading Guidelines', path: 'Claude/Proofreading/Guidelines.docx', weight: 1.0 },
    { key: 'gfx', label: 'Graphics Guidelines', path: 'Claude/Graphics/Guidelines.docx', weight: 0.95 }
  ];

  /* the sign-in's ID token proves to the relay who is asking; it carries no access to anything */
  async function idToken(force) {
    if (TEST) return TEST.idToken ? TEST.idToken(force) : 'test-token';
    var r;
    try { r = await msalApp.acquireTokenSilent({ scopes: SCOPES, account: msalApp.getActiveAccount(), forceRefresh: !!force }); }
    catch (e) { await getToken(); throw e; }   // getToken sends the person back to sign in when that is what's needed
    var exp = (r.idTokenClaims && r.idTokenClaims.exp) || 0;
    if (!force && exp * 1000 < Date.now() + 5 * 60000) return idToken(true);
    return r.idToken;
  }
  /* A dropped connection (network blip, phone or laptop asleep mid-answer) throws "Failed to fetch".
     Each relay call is stateless (the whole turn is resent), so it is always safe to send it again. */
  var ASK_RETRY_WAITS = [1500, 4000, 9000];
  var ASK_RETRY_STATUS = [408, 425, 429, 500, 502, 503, 504, 520, 521, 522, 523, 524, 529];
  function askSleep(ms) { return new Promise(function (res) { setTimeout(res, ms); }); }
  function askAwake() {
    return new Promise(function (res) {
      if (!document.hidden) { res(); return; }
      function f() { if (!document.hidden) { document.removeEventListener('visibilitychange', f); res(); } }
      document.addEventListener('visibilitychange', f);
    });
  }
  async function askRelay(messages, retried) {
    var body = JSON.stringify({ messages: messages }), last = null;
    for (var a = 0; a <= ASK_RETRY_WAITS.length; a++) {
      if (a) {
        askStatus('Connection dropped. Trying again…');
        await askAwake();
        await askSleep(ASK_RETRY_WAITS[a - 1]);
      }
      var tok = await idToken(!!retried), res;
      try {
        res = TEST && TEST.relay ? await TEST.relay(messages, tok)
          : await fetch(C.askRelay, { method: 'POST', headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + tok }, body: body });
      } catch (e) {
        last = new Error('The connection to Chaim Levy dropped.'); last.resumable = true;
        continue;
      }
      var j = await res.json().catch(function () { return {}; });
      if (res.status === 401 && !retried) return askRelay(messages, true);
      if (res.ok) return j;
      last = new Error(j.error || ('Claude is not available right now (' + res.status + ').'));
      if (ASK_RETRY_STATUS.indexOf(res.status) < 0) throw last;
      last.resumable = true;
    }
    throw last;
  }

  async function wordXml(item) {
    if (TEST) return TEST.docXml(item);
    var full = item['@microsoft.graph.downloadUrl'] ? item : await g('/drives/' + (await driveId()) + '/items/' + item.id);
    var zip = await JSZip.loadAsync(await (await fetch(full['@microsoft.graph.downloadUrl'])).arrayBuffer());
    return zip.file('word/document.xml').async('string');
  }
  async function rulesIndex() {
    if (S.rulesIx) return S.rulesIx;
    var res = await resourceFiles(), d0 = await driveId();
    var parts = await Promise.all(RULE_SOURCES.map(async function (d) {
      try {
        var item;
        if (d.res) {
          var hits = res.filter(function (f) { return d.res.test(f.name) && !/^~/.test(f.name); });
          hits.sort(function (a, b) { return String(b.lastModifiedDateTime || '').localeCompare(String(a.lastModifiedDateTime || '')); });
          item = hits[0];
          if (!item) return null;
        } else item = await g('/drives/' + d0 + '/root:/' + d.path.split('/').map(encodeURIComponent).join('/'));
        return RULES.passages(await wordXml(item), { key: d.key, label: d.label, weight: d.weight });
      } catch (e) { return null; }
    }));
    var docs = parts.filter(Boolean);
    if (!docs.length) throw new Error('The rulebooks could not be opened.');
    S.rulesIx = RULES.index(docs);
    return S.rulesIx;
  }

  /* [[text, italic, bold?]] -> text with _italics_, so Claude sees what is italic */
  function askPlain(r) {
    return (r || []).map(function (s) {
      var t = s[0] || '';
      if (!s[1] || !t.trim()) return t;
      var m = t.match(/^(\s*)([\s\S]*?)(\s*)$/);
      return m[1] + '_' + m[2] + '_' + m[3];
    }).join('');
  }
  function askWhat(i) {
    i = i || {};
    if (i.query) return ' for “' + askCut(i.query, 60) + '”';
    if (i.amud) return ': ' + [i.masechta, i.amud, i.text].filter(Boolean).join(' ');
    if (i.ref || i.words) return ': ' + askCut(i.ref || i.words, 60);
    if (i.doc) return ': ' + askCut(i.doc, 60);
    return '';
  }
  function askCut(s, n) { s = String(s || ''); return s.length > n ? s.slice(0, n) + '…' : s; }
  function askNum(v, lo, hi, dflt) { v = parseInt(v, 10); return isNaN(v) ? dflt : Math.max(lo, Math.min(hi, v)); }
  function askPlaces(hits, n) {
    var out = [];
    hits.forEach(function (h) { var k = h.m + ' ' + h.daf + (h.part > 1 ? ' (' + h.part + ')' : ''); if (out.indexOf(k) < 0 && out.length < n) out.push(k); });
    return out;
  }
  /* after an answer, earlier search results are dropped from the history (Claude keeps its answer), so follow-ups stay cheap */
  function askSlim(turn) {
    return turn.map(function (m) {
      if (typeof m.content === 'string') return m;
      if (m.role === 'assistant') return { role: 'assistant', content: m.content.filter(function (b) { return b.type !== 'thinking' && b.type !== 'redacted_thinking'; }) };
      return { role: 'user', content: m.content.map(function (b) {
        return b.type === 'tool_result' ? { type: 'tool_result', tool_use_id: b.tool_use_id, content: '(earlier search results omitted)' } : b;
      }) };
    });
  }

  /* ---------- Ask Claude: the source texts, the Tanach and the skill instructions (shared by the site and the app) ----------
     io.json(name)   -> a JSON file from Claude/Ask Claude data (the sources-*.json files, tanach.json, sources-index.json)
     io.skills()     -> [{ name: 'mdy-writer/house-voice.md', id: <anything io.text understands> }] from Claude/skills
     io.text(id)     -> that file's text
     io.ledger()     -> the text of Claude/Ledger/rulings.jsonl (settled rulings and open questions)
     Same code in mdy-lookup (SharePoint) and the MDY Team app; keep the two copies identical. */
  var ASKX = (function () {
    var MARKS = /[֑-ׇֽֿׁׂׅׄ]/g;
    function hnorm(s) {
      return String(s || '').replace(MARKS, '').replace(/[־\-–—]/g, ' ').replace(/[״"“”]/g, '"').replace(/[׳'‘’`]/g, "'")
        .replace(/[.,;:!?()\[\]{}<>*_|]/g, ' ').toLowerCase().replace(/\s+/g, ' ').trim();
    }
    function cut(s, n) { s = String(s || ''); return s.length > n ? s.slice(0, n) + '…' : s; }
    function num(v, lo, hi, d) { v = parseInt(v, 10); return isNaN(v) ? d : Math.max(lo, Math.min(hi, v)); }
    var GEM = { 'א': 1, 'ב': 2, 'ג': 3, 'ד': 4, 'ה': 5, 'ו': 6, 'ז': 7, 'ח': 8, 'ט': 9, 'י': 10, 'כ': 20, 'ך': 20, 'ל': 30, 'מ': 40, 'ם': 40,
      'נ': 50, 'ן': 50, 'ס': 60, 'ע': 70, 'פ': 80, 'ף': 80, 'צ': 90, 'ץ': 90, 'ק': 100, 'ר': 200, 'ש': 300, 'ת': 400 };
    function gem(s) { var n = 0; String(s).replace(MARKS, '').split('').forEach(function (c) { n += GEM[c] || 0; }); return n; }
    function amudKey(a) {
      a = String(a || '').trim().replace(MARKS, '');
      var m = a.match(/(\d+)\s*([ab])\b/i) || a.match(/(\d+)\s*(?:amud\s*)?([ab])/i);
      if (m) return m[1] + m[2].toLowerCase();
      m = a.match(/([א-ת"״'׳]+)\s*([.:]|ע["״]?[אב])?/);
      if (m) { var n = gem(m[1]); if (n) return n + (/:|ע["״]?ב/.test(m[2] || '') ? 'b' : 'a'); }
      m = a.match(/(\d+)/);
      return m ? m[1] + 'a' : '';
    }
    var KIND = { gemara: 'gemara', text: 'gemara', gemoro: 'gemara', mishnah: 'mishnah', mishna: 'mishnah', mishnayos: 'mishnah',
      rashi: 'rashi', mesivta: 'mesivta', metivta: 'mesivta', biur: 'mesivta', hearos: 'mesivta_hearos', mesivta_hearos: 'mesivta_hearos', notes: 'mesivta_hearos',
      tosafos: 'tosafos', tosfos: 'tosafos', tosafot: 'tosafos', tosafes: 'tosafos', tos: 'tosafos',
      rosh: 'rosh', harosh: 'rosh', peirush_harosh: 'rosh', rabbeinu_asher: 'rosh',
      peirush: 'peirush', perush: 'peirush', pirush: 'peirush', vilna_peirush: 'peirush',
      rambam: 'rambam', peirush_hamishnayos: 'rambam', maimonides: 'rambam',
      bartenura: 'bartenura', bertinoro: 'bartenura', bartenuro: 'bartenura', rav: 'bartenura',
      delivered: 'delivered', english: 'delivered', translation: 'delivered', elucidation: 'delivered', final: 'delivered', corpus: 'delivered', volume: 'delivered',
      commentary: '@record', commentary_of_record: '@record', record: '@record', main_commentary: '@record' };
    var NOTE = {
      gemara: 'Menukad Gemara text of record (nikkud byte-for-byte from the source).',
      mishnah: 'The pointed Mishnah, divided by Vilna amud: the text of record for this masechta.',
      rashi: 'Vilna Rashi, unpointed.',
      mesivta: 'Mesivta (Oz VeHadar) biur column, OCR of the scan: expect occasional letter confusions; never use it as evidence for nikkud.',
      mesivta_hearos: "Mesivta he'aros (the numbered notes under the line), OCR. Reference only: never fold them into an elucidation.",
      delivered: 'The delivered English of this amud, chunk by chunk as "Hebrew ‖ English"; [line] marks an English line or header on its own line.'
    };
    var NAME = { rashi: 'Rashi', tosafos: 'Tosafos', rosh: 'the Rosh', peirush: 'the Vilna Peirush', rambam: 'the Rambam', bartenura: 'Bartenura' };
    function kindOf(v) { var k = String(v || 'gemara').toLowerCase().replace(/['’]/g, '').replace(/[^a-z_ ]/g, '').trim().replace(/\s+/g, '_'); return KIND[k] || k || 'gemara'; }
    /* where each instruction document sits in the authority order of rulings.md § Authority hierarchy */
    function standing(name) {
      var n = String(name).toLowerCase();
      if (/^ledger\//.test(n)) return 'Ledger ruling (level 3: settled by a senior editor, dated; a later ruling beats an earlier document)';
      if (/rulings\.md$|chat-rulings|handback-replies/.test(n)) return 'Level 3: senior editors\' settled decisions (rulings.md is canonical)';
      if (/style-guide|terminology|phrase-book|definitions\.md|pshat-checklist/.test(n)) return 'Level 6: the style guide and its satellites';
      if (/^mdy-proofreader\//.test(n)) return 'Proofreading rules (the proofreader\'s pass; the Transliteration List outranks her habits)';
      if (/^mdy-graphics\//.test(n)) return 'Graphics instructions';
      if (/\/handback-(?!replies)[^\/]*\.md$/.test(n)) return 'Handback: a write-up of one editor\'s pass on one early draft. A ruling it states or marks as one is level 3; a general rule it draws from a few edits is evidence only, ranked below the delivered Temurah and Kerisos volumes, and never grounds on its own for deleting a form the volumes use regularly (rulings.md, fourth qualification; R0053)';
      if (/^mdy-writer\//.test(n)) return 'Writer guidance, drawn from the rulings and the delivered Finals';
      return 'Editor guidance';
    }

    return function make(io) {
      var IX = null, FILES = {}, TAN = null, SK = null;
      function index() { return IX || (IX = io.json('sources-index.json')); }
      async function masechta(name) {
        var ix = await index(), w = hnorm(name).replace(/\s/g, '');
        if (!w) return null;
        var all = ix.masechtos;
        var hit = all.filter(function (m) { return [m.masechta].concat(m.aliases || []).some(function (x) { return hnorm(x).replace(/\s/g, '') === w; }); })[0] ||
          all.filter(function (m) { return [m.masechta].concat(m.aliases || []).some(function (x) { x = hnorm(x).replace(/\s/g, ''); return x.indexOf(w) === 0 || w.indexOf(x) === 0; }); })[0];
        return hit || null;
      }
      function file(m) { return FILES[m.file] || (FILES[m.file] = io.json(m.file)); }
      function load(name) { return FILES[name] || (FILES[name] = io.json(name)); }
      function have(ix) {
        return ix.masechtos.map(function (m) {
          return m.masechta + (m.producer ? ' (' + m.producer + ')' : '') + ': ' + Object.keys(m.texts).map(function (k) {
            var t = m.texts[k]; return k + ' ' + t.first + '–' + t.last + (t.role ? ' [' + t.role + ']' : '');
          }).join(', ');
        });
      }
      /* the text of one kind for one masechta: { note, amudim, role } from the masechta's sources file or from the text's own file */
      async function textOf(m, kind) {
        var T = m.texts[kind];
        if (!T) return null;
        if (T.file) { var f = await load(T.file); return { note: T.note || f.source || NOTE[kind] || '', amudim: f.amudim, role: T.role }; }
        if (!m.file) return null;
        var d = await file(m), X = d.texts[kind];
        return X ? { note: X.note || T.note || NOTE[kind] || '', amudim: X.amudim, role: T.role } : null;
      }
      /* which text answers a request: the commentary of record stands in the rashi slot where it is not Rashi */
      function pick(m, kind) {
        var rec = m.commentary_of_record || 'rashi';
        if (kind === '@record') kind = rec;
        if (!m.texts[kind]) {
          if (kind === 'gemara' && m.texts.mishnah) kind = 'mishnah';
          else if (kind === 'mishnah' && m.texts.gemara) kind = 'gemara';
          else if (kind === 'rashi' && m.texts[rec]) kind = rec;
          else if ((kind === 'gemara' || kind === 'mishnah') && m.texts.delivered && Object.keys(m.texts).length === 1) kind = 'delivered';
        }
        return kind;
      }
      function slice(body, input, out) {
        if (input.query) {
          var q = hnorm(input.query), paras = body.split(/\n+/).filter(function (p) { return hnorm(p).indexOf(q) >= 0; });
          out.query = input.query;
          out.matching_paragraphs = paras.length;
          out.content = cut(paras.join('\n\n'), 12000);
          if (!paras.length) out.note = 'No paragraph on this amud contains that; ask without a query for the whole amud.';
          return out;
        }
        var size = 9000, parts = Math.max(1, Math.ceil(body.length / size)), p = num(input.part, 1, parts, 1);
        out.part = p; out.parts = parts;
        out.content = body.slice((p - 1) * size, p * size);
        if (parts > p) out.note = 'Part ' + p + ' of ' + parts + '. Ask for part ' + (p + 1) + ' for more, or pass a query to get only the paragraphs you need.';
        return out;
      }

      async function readSource(input) {
        var ix = await index(), m = await masechta(input.masechta);
        if (!m) return { error: 'No source texts for "' + (input.masechta || '') + '".', available: have(ix) };
        var asked = kindOf(input.text), kind = pick(m, asked), rec = m.commentary_of_record || null;
        var T = await textOf(m, kind);
        if (!T) return { error: m.masechta + ' has no ' + kind + ' text in the library.', available: have({ masechtos: [m] })[0] };
        var a = amudKey(input.amud), body = T.amudim[a], keys = Object.keys(T.amudim);
        if (body == null) return { error: m.masechta + ' ' + kind + ' has no amud ' + (input.amud || '') + '.', range: keys[0] + '–' + keys[keys.length - 1] };
        var out = { masechta: m.masechta, amud: a, text: kind, about: T.note + (m.text_kind === 'derived' && kind === 'gemara' ? ' For this masechta the text is derived from the delivered English master, so it follows the delivered edition, not the Vilna daf.' : '') };
        if (T.role) out.role = T.role;
        if (rec && rec !== 'rashi') out.commentary_of_record = NAME[rec] + ' (text "' + rec + '")';
        if (kind !== asked && asked !== '@record') out.note_on_text = 'You asked for ' + asked + '; on ' + m.masechta + ' that place is held by ' + (NAME[kind] || kind) + '.';
        if (m.status) out.status = m.status;
        slice(body, input, out);
        /* a secondary commentary always comes with the commentary of record on the same amud, so the hierarchy is never lost */
        if (rec && kind !== rec && m.texts[rec] && T.role === 'secondary commentary') {
          var R = await textOf(m, rec), rb = R && R.amudim[a];
          if (rb != null) {
            var r2 = slice(rb, input, {});
            out.commentary_of_record_on_this_amud = { text: rec, about: R.note, content: cut(r2.content || '', 6000), note: r2.note };
          }
        }
        if (m.producer) out.warning = m.producer + '. Not our house style: never cite it as precedent for how we write.';
        return out;
      }

      async function searchSources(input) {
        var ix = await index(), q = hnorm(input.query);
        if (!q) return { error: 'empty query' };
        var ms = ix.masechtos;
        if (input.masechta) { var one = await masechta(input.masechta); if (!one) return { error: 'No source texts for "' + input.masechta + '".', available: have(ix) }; ms = [one]; }
        var want = input.text ? kindOf(input.text) : null;
        var lim = num(input.limit, 1, 40, 15), hits = [], counts = {}, total = 0;
        for (var i = 0; i < ms.length; i++) {
          var m = ms[i], w = want === '@record' ? (m.commentary_of_record || 'rashi') : want;
          var kinds = Object.keys(m.texts).filter(function (k) {
            if (w) return k === w || (w === 'gemara' && k === 'mishnah');
            return k !== 'delivered';      /* the delivered English is searched by search_corpus, or here with text "delivered" */
          });
          for (var j = 0; j < kinds.length; j++) {
            var k = kinds[j], T = await textOf(m, k);
            if (!T) continue;
            var am = T.amudim;
            Object.keys(am).forEach(function (a) {
              am[a].split(/\n+/).forEach(function (para) {
                var n = hnorm(para), at = n.indexOf(q);
                if (at < 0) return;
                total++;
                var c = m.masechta + ' ' + k; counts[c] = (counts[c] || 0) + 1;
                if (hits.length < lim) hits.push({ masechta: m.masechta, amud: a, text: k, role: T.role || undefined, snippet: '…' + n.slice(Math.max(0, at - 160), at + q.length + 160) + '…' });
              });
            });
          }
        }
        return { query: input.query, searched: ms.map(function (m) { return m.masechta; }), text: want || 'all but the delivered English', paragraphs_matched: total, per_masechta_and_text: counts, samples: hits,
          note: 'Matching ignores nikkud. Snippets are shown without nikkud; read_source gives the pointed text. Commentary of record: Tosafos on Meilah, the Rosh on Tamid, the Rambam on Kinnim and Middos, Rashi elsewhere; the others are secondary.' };
      }

      /* MDY's own Bechoros volume, English only: searched on request and beside the house corpus, always labelled as not house style */
      async function searchBechoros(input) {
        var ix = await index(), m = ix.masechtos.filter(function (x) { return x.masechta === 'Bechoros'; })[0];
        if (!m || !m.texts.delivered) return { error: 'Bechoros is not in the library.' };
        var warn = "MDY's own in-house Bechoros: NOT Aleh Zayis work and NOT our house style (different transliteration system, lower-case Mishnah/Baraisa, contractions, other spellings). Use it for content and to see how MDY put something; never as precedent for how we write.";
        var raw = String(input.query || '').trim(), q = hnorm(raw);
        if (!q) return { error: 'empty query' };
        if (/[א-ת]/.test(q)) return { masechta: 'Bechoros', warning: warn, matches: 0, note: 'Bechoros is English only (its Hebrew could not be recovered from the print PDF). Search it in English.' };
        var T = await textOf(m, 'delivered'), lim = num(input.limit, 1, 40, 15), hits = [], total = 0, per = {};
        Object.keys(T.amudim).forEach(function (a) {
          T.amudim[a].split(/\n+/).forEach(function (para) {
            var n = hnorm(para), at = n.indexOf(q);
            if (at < 0) return;
            total++; per[a] = (per[a] || 0) + 1;
            if (hits.length < lim) hits.push({ daf: a, english: (at > 300 ? '…' : '') + cut(n.length === para.length ? para.slice(Math.max(0, at - 300)) : para, 700) });
          });
        });
        return { masechta: 'Bechoros', producer: 'MDY (in-house), not Aleh Zayis', warning: warn, matches: total, amudim_with_matches: Object.keys(per).length, samples: hits,
          note: 'English only, recovered from the print PDF, so word order can be slightly disturbed where the Hebrew stood. read_source with masechta Bechoros (text "delivered") gives a whole amud.' };
      }

      function tanach() { return TAN || (TAN = io.json('tanach.json')); }
      function bookOf(t, s) {
        var k = hnorm(s).replace(/\s/g, '').replace(/^1/, 'i').replace(/^2/, 'ii');
        var al = t.aliases || {};
        if (al[k]) return al[k];
        var best = Object.keys(al).filter(function (x) { return x.length > 2 && k.indexOf(x) === 0; }).sort(function (a, b) { return b.length - a.length; })[0];
        return best ? al[best] : null;
      }
      async function lookupPasuk(input) {
        var t = await tanach(), books = t.books, out = [];
        if (input.ref) {
          var r = String(input.ref).trim().replace(MARKS, '');
          var m = r.match(/^(.*?)[\s,]*(\d+|[א-ת"״'׳]+)\s*[:.,\s]\s*(\d+|[א-ת"״'׳]+)(?:\s*[-–]\s*(\d+|[א-ת"״'׳]+))?\s*$/);
          if (!m) return { error: 'Give a reference like "Vayikra 5:17" or "ויקרא ה יז".' };
          var bn = bookOf(t, m[1]), b = books.filter(function (x) { return x.en === bn; })[0];
          if (!b) return { error: 'No book called "' + m[1] + '".' };
          var toN = function (x) { return /\d/.test(x) ? parseInt(x, 10) : gem(x); };
          var c = toN(m[2]), v1 = toN(m[3]), v2 = m[4] ? toN(m[4]) : v1;
          var ch = b.ch[c - 1];
          if (!ch) return { error: b.en + ' has ' + b.ch.length + ' chapters.' };
          for (var v = v1; v <= Math.min(v2, v1 + 9, ch.length); v++) out.push({ ref: b.en + ' ' + c + ':' + v, text: ch[v - 1] });
          if (!out.length) return { error: b.en + ' ' + c + ' has ' + ch.length + ' verses.' };
          return { source: t.source, pesukim: out };
        }
        var q = hnorm(input.words);
        if (!q) return { error: 'Give a ref or some words of the pasuk.' };
        var words = q.split(' '), lim = num(input.limit, 1, 15, 8), exact = [], loose = [];
        books.forEach(function (b) {
          b.ch.forEach(function (ch, ci) {
            ch.forEach(function (v, vi) {
              var n = hnorm(v);
              if (n.indexOf(q) >= 0) exact.push({ ref: b.en + ' ' + (ci + 1) + ':' + (vi + 1), text: v });
              else if (words.length > 1 && words.every(function (w) { return n.indexOf(w) >= 0; })) loose.push({ ref: b.en + ' ' + (ci + 1) + ':' + (vi + 1), text: v });
            });
          });
        });
        return { source: t.source, exact_matches: exact.length, all_words_matches: loose.length, pesukim: exact.concat(loose).slice(0, lim) };
      }

      async function skills() {
        if (SK) return SK;
        SK = (async function () {
          var list = await io.skills();
          var led = [];
          if (io.ledger) {
            try {
              String(await io.ledger()).split(/\n/).forEach(function (l) {
                if (!l.trim()) return;
                try { led.push(JSON.parse(l)); } catch (e) { /* skip a bad line */ }
              });
            } catch (e) { led = []; }
          }
          var docs = await Promise.all(list.map(async function (f) {
            try { return { name: f.name, text: String(await io.text(f.id)) }; } catch (e) { return null; }
          }));
          docs = docs.filter(Boolean);
          if (led.length) {
            docs.push({ name: 'ledger/rulings.jsonl', ledger: true,
              text: led.map(function (r) {
                return '## ' + [r.id, r.kind, r.status, r.authority, r.date, 'scope ' + r.scope].filter(Boolean).join(' · ') + '\n' +
                  [r.heb ? 'Hebrew: ' + r.heb : '', r.en ? 'Rendering: ' + r.en : '', r.was ? 'Retired: ' + r.was : '', r.rule ? 'Rule: ' + r.rule : '',
                   r.supersedes ? 'Supersedes: ' + r.supersedes : '', r.why ? 'Why: ' + r.why : '', r.source ? 'Source: ' + r.source : ''].filter(Boolean).join('\n');
              }).join('\n\n') });
          }
          docs.forEach(function (d) {
            var secs = [], cur = { path: [], start: 0, lines: [] }, stack = [];
            d.text.split(/\n/).forEach(function (line) {
              var h = line.match(/^(#{1,4})\s+(.*)$/);
              if (h) {
                if (cur.lines.join('').trim()) secs.push(cur);
                stack = stack.slice(0, h[1].length - 1); stack[h[1].length - 1] = h[2].trim();
                cur = { path: stack.filter(Boolean).slice(), lines: [] };
              } else cur.lines.push(line);
            });
            if (cur.lines.join('').trim()) secs.push(cur);
            d.sections = secs.map(function (s) { var body = s.lines.join('\n').trim(); return { path: s.path, body: body, n: hnorm(s.path.join(' ') + ' ' + body) }; });
          });
          return docs;
        })();
        SK.catch(function () { SK = null; });
        return SK;
      }
      function docMatch(docs, name) {
        var w = hnorm(name).replace(/\.md$/, '').replace(/\s/g, '-');
        return docs.filter(function (d) { return d.name.toLowerCase().replace(/\.md$/, '') === w; })[0] ||
          docs.filter(function (d) { return d.name.toLowerCase().replace(/\.md$/, '').split('/').pop() === w.split('/').pop(); })[0] ||
          docs.filter(function (d) { return d.name.toLowerCase().indexOf(w) >= 0; })[0];
      }
      async function searchInstructions(input) {
        var docs = await skills(), q = hnorm(input.query);
        if (!q) return { error: 'empty query' };
        var terms = q.split(' ').filter(function (w) { return w.length > 2 || /[א-ת]/.test(w); });
        if (!terms.length) terms = [q];
        var scored = [];
        docs.forEach(function (d) {
          d.sections.forEach(function (s) {
            var sc = 0;
            terms.forEach(function (w) { var i = -1, c = 0; while ((i = s.n.indexOf(w, i + 1)) >= 0 && c < 6) c++; sc += c ? 1 + Math.log(c) : 0; });
            if (!sc) return;
            if (terms.length > 1 && s.n.indexOf(q) >= 0) sc += 8;
            if (hnorm(s.path.join(' ')).indexOf(q) >= 0) sc += 2;
            if (d.ledger) sc *= 1.5;
            else if (/rulings\.md$/.test(d.name)) sc *= 1.25;
            else if (/\/handback-(?!replies)[^\/]*\.md$/.test(d.name)) sc *= 0.6;
            scored.push({ d: d, s: s, sc: sc });
          });
        });
        scored.sort(function (a, b) { return b.sc - a.sc; });
        var lim = num(input.limit, 1, 12, 6);
        return {
          searched: docs.length + ' documents (Claude/skills, plus the ruling ledger)', sections_matched: scored.length,
          results: scored.slice(0, lim).map(function (h) {
            var body = h.s.body, at = hnorm(body).indexOf(terms[0]);
            var start = 0;
            if (at > 600) { var raw = body.toLowerCase().indexOf(terms[0]); start = raw > 400 ? raw - 400 : 0; }
            return { doc: h.d.name, standing: standing(h.d.name), section: h.s.path.join(' › ') || '(top)', text: (start ? '…' : '') + cut(body.slice(start), 1600) };
          })
        };
      }
      async function readInstructions(input) {
        var docs = await skills();
        if (!input.doc) return { documents: docs.map(function (d) { return { doc: d.name, standing: standing(d.name), sections: d.sections.length }; }) };
        var d = docMatch(docs, input.doc);
        if (!d) return { error: 'No instruction document called "' + input.doc + '".', documents: docs.map(function (x) { return x.name; }) };
        if (input.section) {
          var w = hnorm(input.section);
          var hits = d.sections.filter(function (s) { return hnorm(s.path.join(' ')).indexOf(w) >= 0; });
          if (!hits.length) return { error: 'No section matching "' + input.section + '" in ' + d.name + '.', sections: d.sections.map(function (s) { return s.path.join(' › '); }).filter(Boolean) };
          return { doc: d.name, standing: standing(d.name), sections: hits.slice(0, 4).map(function (s) { return { section: s.path.join(' › '), text: cut(s.body, 8000) }; }) };
        }
        var size = 10000, parts = Math.max(1, Math.ceil(d.text.length / size)), p = num(input.part, 1, parts, 1);
        var out = { doc: d.name, standing: standing(d.name), part: p, parts: parts, content: d.text.slice((p - 1) * size, p * size) };
        if (p === 1) out.sections = d.sections.map(function (s) { return s.path.join(' › '); }).filter(Boolean).slice(0, 80);
        if (parts > p) out.note = 'Part ' + p + ' of ' + parts + '. Ask for a section by name, or part ' + (p + 1) + '.';
        return out;
      }

      return {
        extra: { bechoros: searchBechoros },
        labels: { read_source: 'Reading the source texts', search_sources: 'Searching the source texts and commentaries', lookup_pasuk: 'Looking up the pasuk',
          search_instructions: 'Searching the skill instructions and rulings', read_instructions: 'Reading the skill instructions' },
        tools: { read_source: readSource, search_sources: searchSources, lookup_pasuk: lookupPasuk, search_instructions: searchInstructions, read_instructions: readInstructions }
      };
    };
  })();

  /* the source texts and Tanach for Ask Claude live in Claude/Ask Claude data (documents library); cached by eTag like the corpus */
  var ASK_DATA = {};
  function askData(name) {
    if (TEST) return TEST.asset(name);
    return ASK_DATA[name] || (ASK_DATA[name] = (async function () {
      var item = await g('/drives/' + (await driveId()) + '/root:/Claude/Ask Claude data/' + encodeURIComponent(name));
      var tag = item.eTag || item.lastModifiedDateTime, key = 'askdata:' + name;
      var hit = await idbGet(key);
      if (hit && hit.tag === tag) return hit.data;
      var data = await (await fetch(item['@microsoft.graph.downloadUrl'])).json();
      idbPut(key, { tag: tag, data: data });
      return data;
    })().catch(function (e) { delete ASK_DATA[name]; throw e; }));
  }
  var ASKX_APP = ASKX({
    json: function (name) { return askData(name); },
    skills: async function () {
      var d = await driveId(), out = [];
      var dirs = await gAll('/drives/' + d + '/root:/Claude/skills:/children?$select=id,name,folder&$top=200');
      for (var i = 0; i < dirs.length; i++) {
        if (!dirs[i].folder) continue;
        var files = await gAll('/drives/' + d + '/items/' + dirs[i].id + '/children?$top=200');
        files.forEach(function (f) { if (f.file && /\.md$/i.test(f.name)) out.push({ name: dirs[i].name + '/' + f.name, id: f }); });
      }
      return out;
    },
    text: async function (f) {
      var url = f['@microsoft.graph.downloadUrl'] || (await g('/drives/' + (await driveId()) + '/items/' + f.id))['@microsoft.graph.downloadUrl'];
      return (await fetch(url)).text();
    },
    ledger: async function () {
      var it = await g('/drives/' + (await driveId()) + '/root:/Claude/Ledger/rulings.jsonl');
      return (await fetch(it['@microsoft.graph.downloadUrl'], { cache: 'no-store' })).text();
    }
  });
  Object.keys(ASKX_APP.labels).forEach(function (k) { ASK_LABEL[k] = ASKX_APP.labels[k]; });

  async function askTool(name, input, onStatus) {
    input = input || {};
    if (ASKX_APP.tools[name]) return ASKX_APP.tools[name](input);
    var q = String(input.query || '').trim();
    if (!q) return { error: 'empty query' };
    if (name === 'search_rules') {
      var r = RULES.search(await rulesIndex(), q, askNum(input.limit, 1, 15, 8));
      return {
        total: r.total,
        passages: r.hits.map(function (h) {
          var p = h.p;
          return {
            source: [p.doc.label].concat(p.path.filter(Boolean)).join(' › '),
            status: p.status || 'settled',
            lead: p.lead || undefined,
            text: askCut(askPlain(p.rich), 1500),
            examples: (p.ex || []).slice(0, 3).map(function (e) { return askCut(askPlain(e), 400); })
          };
        })
      };
    }
    if (name === 'search_corpus' && input.masechta && /^(bech|bchor|בכור)/i.test(String(input.masechta).trim())) return ASKX_APP.extra.bechoros(input);
    if (name === 'search_corpus' || name === 'search_headers') {
      var idx = await corpusIndex();
      var names = idx.masechtos.map(function (x) { return x.name; }), want = names;
      if (input.masechta) {
        var w = String(input.masechta).toLowerCase();
        want = names.filter(function (m) { var l = m.toLowerCase(); return l.indexOf(w) === 0 || w.indexOf(l) === 0; });
        if (!want.length) return { error: 'No masechta called ' + input.masechta + ' in the corpus.', masechtos: names };
      }
      if (name === 'search_corpus') {
        for (var i = 0; i < want.length; i++) await loadCorpus(want[i], onStatus);
        var res = phraseSearch(want.map(function (m) { return S.corp[m] || []; }), q, true);
        if (!res) return { error: 'The search needs at least two letters.' };
        for (var li = 0; li < want.length; li++) await loadLines(want[li]).catch(function () { return []; });
        var lres = lineSearch(want.map(function (m) { return (S.lines || {})[m] || []; }), q, true);
        var per = {}, exact = 0;
        res.groups.forEach(function (gr) { gr.hits.forEach(function (h) { per[h.m] = (per[h.m] || 0) + 1; if (gr.exact) exact++; }); });
        var out = {
          corpus_covers: names, searched: want, occurrences_per_masechta: per,
          total_occurrences: res.total, capped: res.capped || undefined, exact_phrase_chunks: exact, renderings_found: res.groups.length,
          renderings: res.groups.slice(0, askNum(input.limit, 1, 40, 15)).map(function (gr) {
            var pm = {};
            gr.hits.forEach(function (h) { pm[h.m] = (pm[h.m] || 0) + 1; });
            return { english: askCut(gr.rich ? askPlain(gr.rich) : gr.eng, 700), hebrew: askCut(gr.heb, 400), uses: gr.hits.length,
              per_masechta: pm, whole_chunk_is_the_phrase: gr.exact || undefined, places: askPlaces(gr.hits, 6) };
          }),
          english_lines_with_no_hebrew_before_them: lres.total,
          english_line_samples: lres.groups.slice(0, 10).map(function (gr) {
            return { english: askCut(gr.rich ? askPlain(gr.rich) : gr.text, 700), uses: gr.items.length, places: askPlaces(gr.items, 4) };
          }),
          also_available: 'Bechoros (MDY in-house, not house style): search_corpus with masechta "Bechoros". Whole delivered amudim: read_source with text "delivered".'
        };
        if (!input.masechta && !hasHe(q)) {
          try { var bx = await ASKX_APP.extra.bechoros({ query: q, limit: 3 }); if (bx.matches) out.mdy_bechoros_not_house_style = { matches: bx.matches, warning: bx.warning, samples: bx.samples }; } catch (e) { /* Bechoros is optional */ }
        }
        return out;
      }
      if (hasHe(q)) return { error: 'Headers are in English; search with English words.' };
      for (var k = 0; k < want.length; k++) await loadHeaders(want[k]);
      var groups = headerSearch(want.map(function (m) { return S.hdr[m] || []; }), q);
      var hper = {}, htotal = 0;
      groups.forEach(function (gr) { gr.items.forEach(function (h) { hper[h.m] = (hper[h.m] || 0) + 1; htotal++; }); });
      if (!htotal && !Object.keys(S.hdr).some(function (m) { return S.hdr[m].length; })) return { error: 'Headers have not been read into the corpus yet.' };
      return {
        corpus_covers: names, searched: want, total_headers_matched: htotal, per_masechta: hper, wordings_found: groups.length,
        wordings: groups.slice(0, askNum(input.limit, 1, 30, 12)).map(function (gr) {
          var pm = {};
          gr.items.forEach(function (h) { pm[h.m] = (pm[h.m] || 0) + 1; });
          return { wording: gr.rich ? askPlain(gr.rich) : gr.text, uses: gr.items.length, per_masechta: pm,
            kinds: { opens_a_paragraph: gr.kinds.o, mid_paragraph: gr.kinds.m, own_line: gr.kinds.s }, examples: askPlaces(gr.items, 4) };
        })
      };
    }
    if (name === 'lookup_terms') {
      var out = {};
      var lists = [['transliterations', 'Spelling', 10], ['definitions', 'Definitions', 6]];
      for (var j = 0; j < lists.length; j++) {
        var mode = GUIDE.filter(function (x) { return x.key === lists[j][1]; })[0];
        try {
          var doc = await loadDoc(mode);
          out[lists[j][0]] = searchDoc(doc, q, mode).slice(0, lists[j][2]).map(function (h) {
            return h.e.cells ? h.e.cells.map(function (c) { return askCut(c, 600); }).join(' | ') : askCut(h.e.text, 900);
          });
        } catch (e) { out[lists[j][0]] = { error: e.message }; }
      }
      return out;
    }
    return { error: 'unknown tool' };
  }

  /* Claude's answer: a small, safe subset of Markdown (paragraphs, lists, tables, bold, italics, code) */
  function askMd(s) {
    var out = [], list = null, table = null;
    function inline(t) {
      return esc(t)
        .replace(/`([^`]+)`/g, '<code>$1</code>')
        .replace(/\*\*\*([^*\n]+?)\*\*\*/g, '<strong><em>$1</em></strong>')
        .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
        .replace(/(^|[\s(\[“‘"';>])\*([^*\s][^*]*)\*(?=[\s)\].,;:!?”’"'s&<]|$)/g, '$1<em>$2</em>')
        .replace(/(^|[\s(\[“‘"';>])_([^_\s][^_]*)_(?=[\s)\].,;:!?”’"'s&<]|$)/g, '$1<em>$2</em>');
    }
    function endList() { if (list) { out.push('<ul>' + list.join('') + '</ul>'); list = null; } }
    function endTable() {
      if (!table) return;
      out.push('<div class="ask-tbl"><table>' + table.map(function (r, i) {
        var tag = i === 0 ? 'th' : 'td';
        return '<tr>' + r.map(function (c) { return '<' + tag + ' dir="auto">' + inline(c) + '</' + tag + '>'; }).join('') + '</tr>';
      }).join('') + '</table></div>');
      table = null;
    }
    String(s || '').split(/\n/).forEach(function (line) {
      if (/^\s*\|.*\|\s*$/.test(line)) {
        endList();
        if (/^\s*\|[\s:|-]+\|\s*$/.test(line)) return;
        (table = table || []).push(line.trim().slice(1, -1).split('|').map(function (c) { return c.trim(); }));
        return;
      }
      endTable();
      var m = line.match(/^\s*(?:[-*•]|\d+[.)])\s+(.*)$/);
      if (m) { (list = list || []).push('<li dir="auto">' + inline(m[1]) + '</li>'); return; }
      endList();
      var h = line.match(/^#{1,4}\s+(.*)$/);
      if (h) { out.push('<p><strong>' + inline(h[1]) + '</strong></p>'); return; }
      if (line.trim()) out.push('<p dir="auto">' + inline(line) + '</p>');
    });
    endTable(); endList();
    return out.join('');
  }

  /* ---- the question box: bold and italics (same code as the site's Ask Claude page, bundle 2026-10-18-821b188). A contenteditable box (not a textarea) so Ctrl+B / Ctrl+I, the B and I buttons,
     and italics pasted from Word all show; on Ask it is sent to Claude as Markdown (**bold**, *italic*). Only bold and italics
     survive a paste — fonts, colours, sizes, links and pictures are dropped. ---- */
  var ASK_SKIP = { STYLE: 1, SCRIPT: 1, META: 1, TITLE: 1, HEAD: 1, LINK: 1, XML: 1, IMG: 1, svg: 1, SVG: 1, TEMPLATE: 1 };
  var ASK_BLOCK = /^(P|DIV|LI|UL|OL|H[1-6]|TR|TABLE|BLOCKQUOTE|PRE|SECTION|ARTICLE|HEADER|FOOTER)$/;
  function askRuns(root, fromPaste) {
    var runs = [];
    function nl() { var l = runs[runs.length - 1]; if (l && !/\n$/.test(l[0])) runs.push(['\n', false, false]); }
    function walk(n, b, i) {
      if (n.nodeType === 3) {
        var t = n.nodeValue.replace(/\u00a0/g, ' ');
        if (fromPaste) t = t.replace(/[\r\n\t ]+/g, ' ');
        if (t) runs.push([t, b, i]);
        return;
      }
      if (n.nodeType !== 1 && n.nodeType !== 11) return;
      var tag = n.nodeName;
      if (ASK_SKIP[tag] || /^o:/i.test(tag) && !n.textContent) return;
      if (tag === 'BR') { runs.push(['\n', false, false]); return; }
      if (tag === 'B' || tag === 'STRONG') b = true;
      if (tag === 'I' || tag === 'EM') i = true;
      if (n.style) {
        var fw = n.style.fontWeight, fs = n.style.fontStyle;
        if (fw) b = fw === 'bold' || fw === 'bolder' || parseInt(fw, 10) >= 600;
        if (fs) i = fs === 'italic' || fs === 'oblique';
      }
      var blk = ASK_BLOCK.test(tag);
      if (blk) nl();
      for (var c = n.firstChild; c; c = c.nextSibling) walk(c, b, i);
      if (blk) nl();
    }
    walk(root, false, false);
    var out = [];
    runs.forEach(function (r) {
      var l = out[out.length - 1];
      if (r[0] === '\n') { r = ['\n', false, false]; }
      if (l && (l[1] === r[1] && l[2] === r[2])) l[0] += r[0]; else out.push(r.slice());
    });
    if (fromPaste) out.forEach(function (r) { r[0] = r[0].replace(/ *\n */g, '\n'); });
    var s = out.map(function (r) { return r[0]; }).join('');
    var lead = s.length - s.replace(/^\s+/, '').length, tail = s.length - s.replace(/\s+$/, '').length;
    while (out.length && lead > 0) { var f = out[0]; if (f[0].length <= lead) { lead -= f[0].length; out.shift(); } else { f[0] = f[0].slice(lead); lead = 0; } }
    while (out.length && tail > 0) { var z = out[out.length - 1]; if (z[0].length <= tail) { tail -= z[0].length; out.pop(); } else { z[0] = z[0].slice(0, z[0].length - tail); tail = 0; } }
    out.forEach(function (r) { r[0] = r[0].replace(/\n{3,}/g, '\n\n'); });
    return out;
  }
  function askRunsMd(runs) {
    return runs.map(function (r) {
      if (!r[1] && !r[2]) return r[0];
      var mk = r[1] && r[2] ? '***' : r[1] ? '**' : '*';
      return r[0].split('\n').map(function (line) {
        var m = line.match(/^(\s*)([\s\S]*?)(\s*)$/);
        return m[2] ? m[1] + mk + m[2] + mk + m[3] : line;
      }).join('\n');
    }).join('');
  }
  function askRunsHtml(runs, br) {
    return runs.map(function (r) {
      var h = esc(r[0]);
      if (br) h = h.replace(/\n/g, '<br>');
      if (r[2]) h = (br ? '<i>' : '<em>') + h + (br ? '</i>' : '</em>');
      if (r[1]) h = (br ? '<b>' : '<strong>') + h + (br ? '</b>' : '</strong>');
      return h;
    }).join('');
  }

  function askLogHtml(A) {
    var h = A.items.map(function (it) {
      if (it.k === 'u') return '<div class="ask-m u" dir="auto">' + (it.h || esc(it.t)) + '</div>';
      if (it.k === 'e') return '<div class="ask-m e">' + esc(it.t) + (it.resume ? ' This happens when the network blips or the phone or computer goes to sleep while he is working. Nothing he found is lost: <strong>Try again</strong> picks up where he stopped.' : '') + '</div>' +
        (it.resume && A.pending && !A.busy ? '<div class="ask-retry"><button type="button" class="chipbtn" data-act="askresume">Try again</button></div>' : '');
      return '<div class="ask-who"><span class="ask-av sm" aria-hidden="true"><img src="icons/chaim-levy.jpg" alt=""></span>Chaim Levy <span class="ask-ai">AI</span></div><div class="ask-m a">' + it.h + '</div>';
    }).join('');
    if (!A.items.length) h = '<div class="ask-m a"><p>What would you like to know? A few things to try:</p><div class="ask-tries">' +
      ASK_TRIES.map(function (t) { return '<button type="button" class="chipbtn" data-act="asktry">' + esc(t) + '</button>'; }).join('') + '</div></div>';
    if (A.busy) h += '<div class="ask-m s" id="askStatus">' + esc(A.status || 'Thinking…') + '</div>';
    return h;
  }
  function renderAsk(scroll) {
    var log = $('#askLog'), A = S.ask;
    if (!log || !A) return;
    log.innerHTML = askLogHtml(A);
    var go = $('#askGo'); if (go) go.disabled = A.busy;
    if (scroll) { var last = log.lastElementChild; if (last) last.scrollIntoView({ block: 'nearest' }); }
  }
  // A fresh conversation: the New conversation button, and a return to the app after it sat in the background
  // (a reload or a cold start already begins empty, since the conversation is kept only in memory)
  function askReset() {
    if (!S.ask || S.ask.busy) return false;
    S.ask.convo = []; S.ask.items = []; S.ask.draft = ''; S.ask.pending = null;
    renderAsk(false);
    var aq = $('#askQ'); if (aq) aq.innerHTML = '';
    return true;
  }
  var ASK_IDLE_MS = 30 * 60 * 1000, askHiddenAt = 0;
  document.addEventListener('visibilitychange', function () {
    if (document.visibilityState === 'hidden') { askHiddenAt = Date.now(); return; }
    if (askHiddenAt && Date.now() - askHiddenAt >= ASK_IDLE_MS) askReset();
    askHiddenAt = 0;
  });
  var askSelHooked = false, askSyncBtns = null;
  function askStatus(t) { S.ask.status = t; var el = $('#askStatus'); if (el) el.textContent = t; }
  function screenAsk() {
    frame({ title: 'Ask Chaim Levy', tab: 'more', kicker: 'AI assistant · style, the texts and the corpus' });
    var A = S.ask || (S.ask = { convo: [], items: [], busy: false });
    setView('<div class="ask-id"><span class="ask-av" aria-hidden="true"><img src="icons/chaim-levy.jpg" alt=""><b>AI</b></span><div><p class="ask-idt">Chaim Levy <span class="ask-ai">AI assistant</span></p>' +
      '<p class="note">Chaim Levy is an AI, not a person on the team. It runs on Claude, searches House Rulings, the Style Guide, the Guidelines and the corpus before it answers, and says where each answer comes from.</p></div></div>' +
      '<div class="ask-log" id="askLog" aria-live="polite"></div>' +
      '<form class="ask-form" id="askForm"><div class="ask-box"><div class="ask-tb" role="toolbar" aria-label="Formatting">' +
      '<button type="button" data-c="bold" title="Bold (Ctrl+B)" aria-label="Bold" aria-pressed="false"><b>B</b></button>' +
      '<button type="button" data-c="italic" title="Italic (Ctrl+I)" aria-label="Italic" aria-pressed="false"><i>I</i></button></div>' +
      '<div class="ask-ed" id="askQ" contenteditable="true" role="textbox" aria-multiline="true" aria-label="Your question" dir="auto" spellcheck="true" ' +
      'data-ph="Ask a question, or paste a passage to check…"></div></div>' +
      '<button class="btn" type="submit" id="askGo">Ask</button></form>' +
      '<div class="ask-foot"><p class="note">AI-generated answers on style, the texts, the corpus and the work itself. Bold and italics (the B and I buttons, Ctrl+B, Ctrl+I, or pasted from Word) are kept. Answers can be wrong; the rulebooks and the overseeing editor have the final word.</p>' +
      '<button type="button" class="btn ghost" data-act="asknew">New conversation</button></div>');
    renderAsk(false);
    var ta = $('#askQ'), fmtBtns = document.querySelectorAll('#askForm .ask-tb button');
    if (A.draft) ta.innerHTML = A.draft;
    try { document.execCommand('styleWithCSS', false, false); } catch (e) {}
    function syncBtns() {
      var sel = window.getSelection(), inside = sel && sel.rangeCount && ta.contains(sel.anchorNode);
      Array.prototype.forEach.call(fmtBtns, function (b) {
        var on = false;
        if (inside) { try { on = document.queryCommandState(b.getAttribute('data-c')); } catch (e) {} }
        b.classList.toggle('on', !!on); b.setAttribute('aria-pressed', on ? 'true' : 'false');
      });
    }
    function fmt(cmd) { ta.focus(); try { document.execCommand(cmd, false, null); } catch (e) {} A.draft = ta.innerHTML; syncBtns(); }
    function insertClean(dt) {
      var html = dt.getData('text/html'), text = dt.getData('text/plain');
      if (html) {
        var runs = askRuns(new DOMParser().parseFromString(html, 'text/html').body, true);
        if (runs.length) { document.execCommand('insertHTML', false, askRunsHtml(runs, true)); return; }
      }
      if (text) document.execCommand('insertText', false, text.replace(/\r\n?/g, '\n'));
    }
    Array.prototype.forEach.call(fmtBtns, function (b) {
      b.addEventListener('mousedown', function (e) { e.preventDefault(); });
      b.addEventListener('click', function () { fmt(b.getAttribute('data-c')); });
    });
    if (!askSelHooked) { askSelHooked = true; document.addEventListener('selectionchange', function () { if (askSyncBtns) askSyncBtns(); }); }
    askSyncBtns = function () { if (ta.isConnected) syncBtns(); };
    ta.addEventListener('paste', function (e) { if (!e.clipboardData) return; e.preventDefault(); insertClean(e.clipboardData); });
    ta.addEventListener('drop', function (e) {
      if (!e.dataTransfer) return;
      e.preventDefault(); ta.focus();
      var rg = document.caretRangeFromPoint ? document.caretRangeFromPoint(e.clientX, e.clientY) : null;
      if (rg && ta.contains(rg.startContainer)) { var sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(rg); }
      insertClean(e.dataTransfer);
    });
    ta.addEventListener('input', function () {
      if (!ta.textContent.replace(/\s/g, '') && ta.innerHTML) ta.innerHTML = '';
      A.draft = ta.innerHTML;
    });
    ta.addEventListener('keydown', function (e) {
      if (e.key === 'Enter' && !e.shiftKey && !e.isComposing && window.matchMedia('(min-width: 900px)').matches) { e.preventDefault(); askSubmit(); return; }
      if ((e.ctrlKey || e.metaKey) && !e.altKey && !e.shiftKey) {
        var k = (e.key || '').toLowerCase();
        if (k === 'b') { e.preventDefault(); fmt('bold'); }
        else if (k === 'i') { e.preventDefault(); fmt('italic'); }
        else if (k === 'u') { e.preventDefault(); }
      }
    });
    $('#askForm').addEventListener('submit', function (e) { e.preventDefault(); askSubmit(); });
  }
  async function askSubmit(text) {
    var A = S.ask, ta = $('#askQ'), q, uh;
    if (text != null) { q = String(text).trim(); uh = esc(q); }
    else { var runs = ta ? askRuns(ta, false) : []; q = askRunsMd(runs).trim(); uh = askRunsHtml(runs, false); }
    if (!q || !A || A.busy) return;
    A.busy = true; A.status = 'Thinking…'; A.draft = ''; A.pending = null;
    A.items.forEach(function (it) { if (it.k === 'e') it.resume = false; });
    if (ta) { ta.innerHTML = ''; if (askSyncBtns) askSyncBtns(); }
    A.items.push({ k: 'u', t: q, h: uh });
    renderAsk(true);
    await askWork(A.convo.concat([{ role: 'user', content: q }]), 0, q);
  }
  async function askResume() {
    var A = S.ask, P = A && A.pending;
    if (!P || A.busy) return;
    A.pending = null; A.busy = true; A.status = 'Picking up where he stopped…';
    for (var i = A.items.length - 1; i >= 0; i--) { if (A.items[i].k === 'e' && A.items[i].resume) { A.items.splice(i, 1); break; } }
    renderAsk(true);
    await askWork(P.turn, P.round, P.q);
  }
  async function askWork(turn, round0, q) {
    var A = S.ask, round = round0;
    try {
      for (; round < ASK_MAX_ROUNDS + ASK_GRACE_ROUNDS; round++) {
        var j = await askRelay(turn);
        var content = j.content || [];
        turn.push({ role: 'assistant', content: content });
        var uses = content.filter(function (b) { return b.type === 'tool_use'; });
        if (j.stop_reason !== 'tool_use' || !uses.length) {
          var answer = content.filter(function (b) { return b.type === 'text'; }).map(function (b) { return b.text; }).join('\n\n').trim();
          A.items.push({ k: 'a', h: askMd(answer || 'I could not find an answer to that.') });
          A.convo = askSlim(turn);
          return;
        }
        var results = [];
        var spent = round >= ASK_MAX_ROUNDS - 1;
        for (var k = 0; k < uses.length; k++) {
          var u = uses[k], out;
          if (round >= ASK_MAX_ROUNDS) {
            out = { error: 'Search limit reached. No more searches are available for this question; answer now from what you already found.' };
          } else {
            askStatus((ASK_LABEL[u.name] || 'Searching') + askWhat(u.input) + '…');
            try { out = await askTool(u.name, u.input, askStatus); } catch (e) { out = { error: String(e && e.message || e) }; }
          }
          var s = JSON.stringify(out);
          results.push({ type: 'tool_result', tool_use_id: u.id, content: s.length > 14000 ? s.slice(0, 14000) + '…(cut)' : s });
        }
        if (spent) results.push({ type: 'text', text: ASK_BUDGET_NOTE });
        turn.push({ role: 'user', content: results });
        askStatus(spent ? 'Writing the answer…' : 'Reading what it found…');
      }
      var said = turn.slice(A.convo.length + 1).filter(function (m) { return m.role === 'assistant' && Array.isArray(m.content); })
        .map(function (m) { return m.content.filter(function (b) { return b.type === 'text'; }).map(function (b) { return b.text; }).join('\n\n'); })
        .join('\n\n').trim();
      var fallback = (said ? said + '\n\n' : '') + '_I used up my searches before I could finish checking this. The notes above are what I found so far; ask me a follow-up and I will pick up from here._';
      A.items.push({ k: 'a', h: askMd(fallback) });
      A.convo = A.convo.concat([{ role: 'user', content: q }, { role: 'assistant', content: fallback }]);
    } catch (e) {
      var resumable = !!(e && e.resumable);
      A.items.push({ k: 'e', t: (e && e.message) || 'Something went wrong.', resume: resumable });
      if (resumable) A.pending = { turn: turn, round: round, q: q };
    } finally {
      A.busy = false;
      renderAsk(true);
      var t2 = $('#askQ'); if (t2 && window.matchMedia('(min-width: 900px)').matches) t2.focus();
    }
  }

  /* ---------- feedback ---------- */
  async function feedbackList() {
    if (S.feedback) return S.feedback;
    var sid = await siteId();
    var lists = await gAll('/sites/' + sid + '/lists?$select=id,displayName');
    var l = lists.filter(function (x) { return x.displayName === C.lists.feedbackTitle; })[0];
    if (!l) throw new Error('The ' + C.lists.feedbackTitle + ' list was not found.');
    var cols = (await g('/sites/' + sid + '/lists/' + l.id + '/columns')).value || [];
    var cat = cols.filter(function (c) { return c.choice && /categor/i.test(c.name + c.displayName); })[0];
    var body = cols.filter(function (c) { return c.text && c.text.allowMultipleLines && !c.readOnly && !c.hidden; })[0];
    S.feedback = { id: l.id, cat: cat, body: body };
    return S.feedback;
  }

  /* ---------- rendering frame ---------- */
  var view = function () { return $('#view'); };
  function frame(o) {
    $('#kicker').innerHTML = o.back ? '<a href="' + o.back + '">' + ICON.back + esc(o.backLabel || 'Back') + '</a>' : esc(o.kicker || 'MDY Aleh Zayis Team');
    $('#title').textContent = o.title;
    document.title = o.title === 'MDY Team' ? 'MDY Team' : o.title + ' · MDY Team';
    $('#subhead').innerHTML = o.sub || '';
    $('#refresh').hidden = !o.refresh;
    var base = currentRoute().replace(/^\/?/, '').split('/')[0] || 'home';
    var links = $$('#tabs a');
    links.forEach(function (a) {
      var t = a.getAttribute('data-tab');
      // the rail link for this screen (wide screens) and its bottom tab (phones, e.g. More)
      if (t === base || t === o.tab) a.setAttribute('aria-current', 'page'); else a.removeAttribute('aria-current');
    });
    view().setAttribute('data-screen', base);
  }
  function setView(html) { view().innerHTML = html; }
  function loading(msg) { setView('<div class="loading">' + esc(msg || 'Loading…') + '</div>'); }
  var retryFn = null;
  function fail(e, again) { retryFn = again || null; setView(errBox(e, !!again)); }

  /* ---------- screens ---------- */
  async function screenHome() {
    frame({ title: 'My work', tab: 'home', refresh: true, kicker: 'MDY Aleh Zayis Team' });
    loading();
    var res = await Promise.all([loadTracker(), loadDeadlines().catch(function () { return []; })]);
    var rows = res[0], dls = res[1];
    markFromDeadlines();
    var mine = rows.filter(isMine);
    var upcoming = dls.filter(function (d) { return d.days != null && d.days >= 0; });
    var myNext = upcoming.filter(deadlineIsMine)[0] || upcoming[0];
    frame({ title: 'My work', tab: 'home', refresh: true, kicker: (S.me && S.me.displayName) || 'MDY Aleh Zayis Team' });
    var h = '<div class="homegrid"><div class="hg-top">';
    if (myNext) h += '<a class="card dl" href="#/deadlines" style="text-decoration:none">' + deadlineCard(myNext).replace(/^<article class="card dl">|<\/article>$/g, '') + '</a>';
    h += '<a class="btn block" href="#/submit">' + ICON.up + 'Submit work</a></div><div class="hg-list">';
    h += '<h2 class="sec">Your amudim <small>' + mine.length + '</small></h2>';
    if (mine.length) h += '<div class="rows">' + mine.map(rowHtml).join('') + '</div>';
    else h += '<div class="card empty">Nothing is under your name yet. Tap the star on any amud in the Tracker to keep it here.</div>';
    if (mine.length) h += '<p class="note">Amudim from your deadlines that haven’t reached the deadline’s stage, plus any you’ve starred.</p>';
    h += '</div><div class="hg-links"><h2 class="sec">Quick links</h2><div class="card pad0 navlist">' +
      '<a href="#/guide">' + ICON.pen.replace('M4 20l1-4L16 5a2.1 2.1 0 0 1 3 3L8 19z', 'M11 4.5a6.5 6.5 0 1 0 0 13 6.5 6.5 0 0 0 0-13zM16 16l4.5 4.5') + '<span class="grow">Look-ups, phrases and headers</span>' + ICON.chev + '</a>' +
      '<a href="#/ask">' + ICON.ask + '<span class="grow">Ask Chaim Levy</span><span class="aside">AI assistant</span>' + ICON.chev + '</a>' +
      '<a href="#/deadlines">' + ICON.cal + '<span class="grow">All deadlines</span>' + ICON.chev + '</a>' +
      '<a href="#/files">' + ICON.folder + '<span class="grow">Files</span>' + ICON.chev + '</a>' +
      '<a href="' + esc(C.chatUrl) + '" target="_blank" rel="noopener">' + ICON.chat + '<span class="grow">Team chat</span>' + ICON.ext + '</a></div></div></div>';
    setView(h);
  }
  function rowHtml(r) {
    var c = stageColor(r.stage);
    return '<div class="row" style="border-left-color:' + c + '">' +
      '<button class="star" data-act="star" data-id="' + esc(r.id) + '" aria-pressed="' + isStarred(r.id) + '" aria-label="Keep ' + esc(r.title) + ' in My work">' + ICON.star + '</button>' +
      '<div class="rt"><b>' + esc(r.title) + '</b><span>' + esc(personLine(r) || r.notes || '') + '</span></div>' +
      '<button class="chipbtn" style="border:none;padding:0;background:transparent" data-act="move" data-id="' + esc(r.id) + '" aria-label="Move ' + esc(r.title) + ', now ' + esc(r.stage) + '">' + stageChip(r.stage) + '</button>' +
      '<a class="star" href="#/submit/' + esc(r.id) + '" aria-label="Submit work for ' + esc(r.title) + '" style="color:var(--navy)"><svg viewBox="0 0 24 24" aria-hidden="true" style="fill:none;stroke:currentColor;stroke-width:2;stroke-linecap:round;stroke-linejoin:round"><path d="M12 16V5M6.5 10.5L12 5l5.5 5.5M5 20h14"/></svg></a>' +
      '</div>';
  }

  /* The site's Workflow page banner: one bar per volume in production, every stage, Final and Closed shown together as Final.
     A bar shows from its 'from' date (Israel time) until every amud in it is Final, then hides itself (YB, 18 Oct 2026).
     A one-masechta volume is left out while that masechta is the one in view, since its own card shows the same. */
  function volumeBanner(rows, inView) {
    var today = ymdIn(new Date());
    return (C.volumes || []).map(function (v) {
      if (v.from && today < v.from) return '';
      if (v.masechtos.length === 1 && v.masechtos[0] === inView) return ''; /* the masechta's own card below already shows it */
      var stg = {}, total = 0, done = 0;
      rows.forEach(function (r) {
        if (v.masechtos.indexOf(r.masechta) < 0) return;
        var st = r.stage === 'Closed' ? 'Final' : (r.stage || 'Open');
        stg[st] = (stg[st] || 0) + 1; total++;
        if (st === 'Final') done++;
      });
      if (!total || done >= total) return '';
      var vlist = stagesFor(v.masechtos);
      var order = vlist.filter(function (s) { return stg[s]; }).concat(Object.keys(stg).filter(function (s) { return vlist.indexOf(s) < 0; }));
      return '<section class="card vol" aria-label="' + esc(v.name) + ' by stage"><div class="volhead"><b>' + esc(v.name) + '</b>' + (v.sub ? '<span class="note">' + esc(v.sub) + '</span>' : '') +
        '<span class="note volcount">' + done + ' of ' + total + ' in Final</span></div>' +
        '<div class="bar">' + order.map(function (s) { return '<i style="flex:' + stg[s] + ';background:' + stageColor(s) + '" title="' + esc(s + ': ' + stg[s]) + '"></i>'; }).join('') + '</div>' +
        '<div class="legend">' + order.map(function (s) { return stageChip(s, s + ' ' + stg[s]); }).join('') + '</div></section>';
    }).join('');
  }

  async function screenTracker(m) {
    var rows = await loadTracker();
    var ms = [];
    rows.forEach(function (r) { if (ms.indexOf(r.masechta) < 0) ms.push(r.masechta); });
    m = m || S.trackerM || ms[0];
    if (ms.indexOf(m) < 0) m = ms[0];
    S.trackerM = m;
    var chips = '<div class="chips scroll" role="group" aria-label="Masechta">' + ms.map(function (x) {
      return '<button class="chipbtn" data-act="tm" data-m="' + esc(x) + '" aria-pressed="' + (x === m) + '">' + esc(x) + '</button>';
    }).join('') + '</div>';
    frame({ title: 'Tracker', tab: 'tracker', refresh: true, kicker: 'Amud Tracker', sub: chips });
    var shown = rows.filter(function (r) { return r.masechta === m; });
    var counts = {};
    shown.forEach(function (r) { var st = r.stage === 'Closed' ? 'Final' : r.stage; counts[st] = (counts[st] || 0) + 1; });
    var mlist = stagesFor(m);
    var order = mlist.filter(function (s) { return counts[s]; }).concat(Object.keys(counts).filter(function (s) { return mlist.indexOf(s) < 0; }));
    var done = counts.Final || 0;
    var h = volumeBanner(rows, m) + '<section class="card" aria-label="Progress"><div style="display:flex;justify-content:space-between;align-items:baseline;margin-bottom:10px"><b>' + esc(m) + ' · ' + shown.length + ' rows</b><span class="note">' + done + ' final</span></div>' +
      '<div class="bar">' + order.map(function (s) { return '<i style="flex:' + counts[s] + ';background:' + stageColor(s) + '" title="' + esc(s) + '"></i>'; }).join('') + '</div>' +
      '<div class="legend">' + order.map(function (s) { return stageChip(s, s + ' ' + counts[s]); }).join('') + '</div></section>';
    h += '<p class="note">Tap an amud to move it to another stage. The star keeps it in My work.</p>';
    h += '<div class="grid2">' + shown.map(function (r) {
      return '<div style="position:relative"><button class="card-btn" style="border-top-color:' + stageColor(r.stage) + ';width:100%" data-act="move" data-id="' + esc(r.id) + '">' +
        '<b>' + esc(r.title.replace(new RegExp('^' + m + '\\s+'), '')) + '</b>' + stageChip(r.stage) + '<span class="who">' + esc(personLine(r)) + '</span></button>' +
        '<button class="star" style="position:absolute;top:4px;right:2px" data-act="star" data-id="' + esc(r.id) + '" aria-pressed="' + isStarred(r.id) + '" aria-label="Keep ' + esc(r.title) + ' in My work">' + ICON.star + '</button></div>';
    }).join('') + '</div>';
    if (!shown.length) h = '<div class="card empty">No rows for ' + esc(m) + '.</div>';
    setView(h);
  }

  function openMoveSheet(row) {
    var list = stagesFor(row.masechta);
    var w = $('#sheet');
    w.innerHTML = '<div class="sheet" role="dialog" aria-modal="true" aria-label="Move ' + esc(row.title) + '">' +
      '<h2>Move ' + esc(row.title) + '</h2><p class="note">Now at ' + esc(row.stage) + '. Pick its new stage.</p>' +
      '<div class="opts">' + list.map(function (s) {
        var c = stageColor(s);
        return '<button data-act="pickstage" data-s="' + esc(s) + '" aria-current="' + (s === row.stage) + '" style="background:' + c + ';color:' + (LIGHT[c] ? '#1a1814' : '#fff') + '">' + esc(s) + '</button>';
      }).join('') + '</div>' +
      '<button class="btn ghost block" data-act="closesheet">Cancel</button></div>';
    w.hidden = false;
    w._row = row;
    var first = $('.opts button[aria-current="true"]', w) || $('.opts button', w);
    if (first) first.focus();
  }
  function closeSheet() { var w = $('#sheet'); w.hidden = true; w.innerHTML = ''; w._row = null; }

  /* --- Submit --- */
  function suggestName(row, file) {
    var ini = initials();
    var base = file ? file.name.replace(/\.[^.]+$/, '') : '';
    var ext = file && /\.[^.]+$/.test(file.name) ? file.name.match(/\.[^.]+$/)[0] : '.docx';
    if (base && /^[A-Z]{3}\d{1,3}[ab]\d?(_[A-Za-z]{1,4})*$/i.test(base)) {
      var parts = base.split('_');
      if (parts[parts.length - 1].toUpperCase() !== ini) base = base + '_' + ini;
      return base + ext;
    }
    var pre = C.prefixes[row.masechta] || row.masechta.slice(0, 3).toUpperCase();
    var m = /(\d+)\s*([ab])\s*(?:\((\d)\))?/i.exec(row.title) || [];
    var daf = m[1] ? (m[1].length < 2 ? '0' + m[1] : m[1]) + m[2].toLowerCase() + (m[3] || '') : '';
    return pre + daf + '_' + ini + ext;
  }
  async function screenSubmit(rowId) {
    frame({ title: 'Submit work', tab: 'submit', kicker: 'Hand in an amud' });
    loading();
    var rows = await loadTracker();
    await loadDeadlines().catch(function () { return []; });
    markFromDeadlines();
    var s = S.submit;
    if (!s || (rowId && s.rowId !== rowId) || s.sent) {
      var mine = rows.filter(isMine);
      var row0 = rows.filter(function (r) { return r.id === rowId; })[0] || mine[0] || null;
      s = S.submit = { rowId: row0 ? row0.id : null, showAll: !!(row0 && !isMine(row0)), stage: row0 ? nextStage(row0.stage, row0.masechta) : '', file: null, name: '', nameTouched: false, role: 'Editors', person: null, note: '', sent: false, busy: false };
    }
    await renderSubmit();
  }
  async function renderSubmit() {
    var s = S.submit, rows = S.tracker;
    var row = rows.filter(function (r) { return r.id === s.rowId; })[0] || null;
    if (s.sent) {
      setView('<div class="done"><div class="tick">' + ICON.tick + '</div><h2>Sent</h2><p>' + esc(s.sent.name) + ' is in ' + esc(s.sent.path) + (s.sent.stage ? ', and ' + esc(s.sent.title) + ' is now marked ' + esc(s.sent.stage) : '') + '.</p>' +
        (s.sent.url ? '<a class="btn ghost block" href="' + esc(s.sent.url) + '" target="_blank" rel="noopener">Open the file</a>' : '') +
        '<button class="btn ghost block" data-act="again">Submit another</button><a class="btn block" href="#/home">Back to My work</a></div>');
      return;
    }
    var pool = s.showAll ? rows : rows.filter(isMine);
    if (row && pool.indexOf(row) < 0) pool = [row].concat(pool);
    var opts = '<option value="">Choose an amud…</option>';
    var lastM = null;
    pool.forEach(function (r) {
      if (r.masechta !== lastM) { if (lastM) opts += '</optgroup>'; opts += '<optgroup label="' + esc(r.masechta) + '">'; lastM = r.masechta; }
      opts += '<option value="' + esc(r.id) + '"' + (r.id === s.rowId ? ' selected' : '') + '>' + esc(r.title) + ' (now ' + esc(r.stage) + ')</option>';
    });
    if (lastM) opts += '</optgroup>';

    // folders
    var folderHtml = '<p class="note">Choose the amud first.</p>', path = '';
    if (row) {
      try {
        var mf = await masechtaFolder(row.masechta);
        if (!mf) folderHtml = '<div class="err">No folder for ' + esc(row.masechta) + ' in the library.</div>';
        else {
          var roles = (await kids(mf.id)).filter(function (x) { return x.folder; });
          if (!roles.some(function (x) { return x.name === s.role; })) s.role = (roles.filter(function (x) { return /editor/i.test(x.name); })[0] || roles[0] || {}).name;
          var roleF = roles.filter(function (x) { return x.name === s.role; })[0];
          var people = roleF ? (await kids(roleF.id)).filter(function (x) { return x.folder; }) : [];
          if (!s.person || !people.some(function (p) { return p.id === s.person; })) {
            var sn = surname();
            var auto = sn ? people.filter(function (p) { return p.name.toLowerCase().indexOf(sn) > -1; })[0] : null;
            s.person = auto ? auto.id : '';
          }
          var personF = people.filter(function (p) { return p.id === s.person; })[0];
          path = row.masechta + ' › ' + (roleF ? roleF.name : '') + (personF ? ' › ' + personF.name : '');
          s.parentId = personF ? personF.id : (roleF ? roleF.id : null);
          s.path = mf.name + ' › ' + (roleF ? roleF.name : '') + (personF ? ' › ' + personF.name : '');
          s.mfName = mf.name;
          folderHtml = '<div class="grid2" style="grid-template-columns:repeat(2,minmax(0,1fr))">' +
            '<div><label class="lbl" for="sRole">Folder</label><select id="sRole" data-act="srole">' + roles.map(function (x) { return '<option' + (x.name === s.role ? ' selected' : '') + '>' + esc(x.name) + '</option>'; }).join('') + '</select></div>' +
            '<div><label class="lbl" for="sPerson">Inside</label><select id="sPerson" data-act="sperson"><option value="">(the folder itself)</option>' + people.map(function (p) { return '<option value="' + esc(p.id) + '"' + (p.id === s.person ? ' selected' : '') + '>' + esc(p.name) + '</option>'; }).join('') + '</select></div></div>';
        }
      } catch (e) { folderHtml = errBox(e); }
    }
    if (row && !s.nameTouched) s.name = suggestName(row, s.file);
    var stages = row ? stagesFor(row.masechta) : (S.stages || []);
    var ready = row && s.file && s.name && s.parentId && !s.busy;
    var h = '<section class="field"><label class="lbl" for="sAmud">1. Which amud?</label><select id="sAmud" data-act="samud">' + opts + '</select>' +
      '<label style="display:flex;gap:8px;align-items:center;font-size:14.5px;min-height:36px"><input type="checkbox" data-act="sall"' + (s.showAll ? ' checked' : '') + ' style="width:20px;height:20px"> Show every amud, not just mine</label></section>';
    h += '<section class="field"><label class="lbl" for="sStage">2. What did you finish?</label><select id="sStage" data-act="sstage"><option value="">Don’t change the tracker</option>' +
      stages.map(function (x) { return '<option' + (x === s.stage ? ' selected' : '') + '>' + esc(x) + '</option>'; }).join('') + '</select></section>';
    h += '<section class="field"><span class="lbl">3. Your file</span>' +
      (s.file ? '<div class="filebox">' + ICON.file + '<span class="fn">' + esc(s.file.name) + '</span><button class="btn ghost" style="min-height:40px" data-act="pickfile">Change</button></div>'
        : '<button class="btn dashed block" style="min-height:64px" data-act="pickfile">Choose a file</button>') +
      '<input type="file" id="sFile" hidden accept=".docx,.doc,.pdf,.xlsx,.png,.jpg,.jpeg,.svg">' +
      '<label class="lbl" for="sName" style="margin:6px 0 0">Saved as</label><input type="text" id="sName" data-act="sname" value="' + esc(s.name) + '" autocomplete="off" spellcheck="false">' +
      '<label class="lbl" for="sNote" style="margin:6px 0 0">Note for the next person (optional)</label><textarea id="sNote" data-act="snote" rows="2">' + esc(s.note) + '</textarea></section>';
    h += '<section class="field"><span class="lbl">4. Where it goes</span>' + folderHtml + '</section>';
    h += '<section class="card summary" aria-label="What will happen"><span class="eyebrow">When you send</span>' +
      '<div><span class="k">Saved as</span><strong class="v">' + esc(s.name || '—') + '</strong></div>' +
      '<div><span class="k">Into</span><span class="v">' + esc(path || '—') + '</span></div>' +
      '<div><span class="k">Tracker</span><span class="v">' + (row ? (s.stage && s.stage !== row.stage ? esc(row.title) + ': ' + esc(row.stage) + ' → <strong>' + esc(s.stage) + '</strong>' : 'No change') : '—') + '</span></div></section>';
    h += '<button class="btn block" data-act="send"' + (ready ? '' : ' disabled') + '>' + (s.busy ? esc(s.busy) : row ? (s.file ? 'Send' : 'Choose a file to send') : 'Choose an amud') + '</button>';
    setView(h);
  }
  async function doSend() {
    var s = S.submit, row = S.tracker.filter(function (r) { return r.id === s.rowId; })[0];
    if (!row || !s.file || !s.parentId) return;
    s.busy = 'Uploading…'; await renderSubmit();
    try {
      var item = await uploadFile(s.parentId, s.name, s.file, function (p) { s.busy = 'Uploading… ' + Math.round(p * 100) + '%'; var b = $('[data-act="send"]'); if (b) b.textContent = s.busy; });
      var changed = '';
      if (s.stage && s.stage !== row.stage) {
        s.busy = 'Marking the tracker…'; var b2 = $('[data-act="send"]'); if (b2) b2.textContent = s.busy;
        var fields = { Stage: s.stage };
        if (s.note.trim()) {
          var stamp = ymdIn(new Date()) + ' ' + ((S.me && S.me.displayName) || '');
          fields.Notes = (row.notes ? row.notes + '\n' : '') + '[' + stamp + '] ' + s.note.trim();
        }
        await g('/sites/' + (await siteId()) + '/lists/' + C.lists.tracker + '/items/' + row.id + '/fields', { method: 'PATCH', body: fields });
        row.stage = s.stage; if (fields.Notes) row.notes = fields.Notes;
        changed = s.stage;
      }
      delete S.folders[s.parentId];
      s.sent = { name: (item && item.name) || s.name, url: item && item.webUrl, path: s.path || row.masechta, title: row.title, stage: changed };
      s.busy = false;
      await renderSubmit();
    } catch (e) {
      s.busy = false;
      if (e.status === 409 || e.code === 'nameAlreadyExists') e = new Error('A file called ' + s.name + ' is already in that folder. Change the name and send again.');
      await renderSubmit();
      view().insertAdjacentHTML('afterbegin', errBox(e));
    }
  }

  function guideMode() { return GUIDE.filter(function (x) { return x.key === S.guideMode; })[0] || GUIDE[0]; }
  async function screenGuide() {
    var mode = guideMode();
    S.cFilter = S.cFilter || 'All';
    var sub = '<div class="chips scroll" role="group" aria-label="Look-up">' + GUIDE.map(function (x) {
      return '<button class="chipbtn" data-act="gmode" data-k="' + x.key + '" aria-pressed="' + (x.key === mode.key) + '">' + x.key + '</button>';
    }).join('') + '<a class="chipbtn chip-ask" href="#/ask">' + ICON.ask + 'Ask Chaim Levy</a></div><label class="sr" for="gq">' + esc(mode.ph) + '</label>' +
      '<input id="gq" type="search" dir="auto" placeholder="' + esc(mode.ph) + '" value="' + esc(S.guideQ) + '" autocomplete="off" style="margin-top:10px">' +
      '<div id="gfilter"></div>';
    frame({ title: 'Guide', tab: 'guide', kicker: 'Style and look-ups', sub: sub });
    var inp = $('#gq');
    inp.addEventListener('input', function () {
      S.guideQ = inp.value; S.expanded = {}; S.limit = 25; S.llimit = 6;
      clearTimeout(screenGuide._t);
      screenGuide._t = setTimeout(mode.kind === 'doc' ? renderGuideResults : renderIndexResults, mode.kind === 'corpus' ? 280 : 150);
    });
    if (mode.kind === 'doc') {
      loading('Opening the ' + mode.key.toLowerCase() + ' document…');
      try { await loadDoc(mode); } catch (e) { fail(e, screenGuide); return; }
      renderGuideResults();
      return;
    }
    loading('Opening the corpus index…');
    try {
      var idx = await corpusIndex();
      var names = idx.masechtos.map(function (x) { return x.name; });
      if (S.cFilter !== 'All' && names.indexOf(S.cFilter) < 0) S.cFilter = 'All';
      $('#gfilter').innerHTML = '<div class="chips scroll" role="group" aria-label="Masechta" style="margin-top:10px">' + ['All'].concat(names).map(function (n) {
        return '<button class="chipbtn" data-act="cfilter" data-m="' + esc(n) + '" aria-pressed="' + (n === S.cFilter) + '">' + esc(n) + '</button>';
      }).join('') + '</div>';
      var want = S.cFilter === 'All' ? names : [S.cFilter];
      for (var i = 0; i < want.length; i++) {
        if (mode.kind === 'corpus') { await loadCorpus(want[i], function (msg) { loading(msg); }); await loadLines(want[i]).catch(function () { return []; }); }
        else await loadHeaders(want[i]);
      }
    } catch (e) { fail(e, screenGuide); return; }
    renderIndexResults();
  }
  function renderIndexResults() {
    var mode = guideMode();
    if (mode.kind === 'doc') return;
    var idx = S.cidx;
    if (!idx) return;
    var names = S.cFilter === 'All' ? idx.masechtos.map(function (x) { return x.name; }) : [S.cFilter];
    var q = S.guideQ.trim(), limit = S.limit || 25, h = '';
    if (mode.kind === 'corpus') {
      var sets = names.map(function (n) { return (S.corp || {})[n] || []; });
      var lsets = names.map(function (n) { return (S.lines || {})[n] || []; });
      var phrases = sets.reduce(function (a, b) { return a + b.length; }, 0), nlines = lsets.reduce(function (a, b) { return a + b.length; }, 0);
      if (!q) {
        setView('<div class="card empty">Type a Hebrew phrase, or English words, to see how it was rendered in the delivered amudim. English with no Hebrew before it (a heading, the words that open a paragraph) is searched too.</div><p class="note">' + phrases.toLocaleString() + ' phrases' + (nlines ? ' and ' + nlines.toLocaleString() + ' English-only lines' : '') + ' from ' + esc(names.join(', ')) + '. The same index as the site’s Guide page.</p>');
        return;
      }
      var res = phraseSearch(sets, q) || { groups: [], total: 0 };
      var lr = lineSearch(lsets, q);
      if (!res.groups.length && !lr.groups.length) { setView('<div class="card empty">Nothing found for “' + esc(q) + '”.</div>'); return; }
      var words = hasHe(q) ? [] : [enNorm(q)];
      var ex = res.groups.filter(function (g) { return g.exact; });
      var rest = res.groups.filter(function (g) { return !g.exact; });
      h += '<p class="note">' + (res.groups.length ? res.total.toLocaleString() + (res.capped ? '+' : '') + ' places in ' + res.groups.length.toLocaleString() + ' renderings' : 'No Hebrew–English phrase') +
        (lr.total ? ', and ' + lr.total.toLocaleString() + ' line' + (lr.total === 1 ? '' : 's') + ' with no Hebrew before ' + (lr.total === 1 ? 'it' : 'them') : '') + '.</p>';
      if (lr.groups.length) {
        var ll = S.llimit || 6, lmax = lr.groups.reduce(function (a, g) { return Math.max(a, g.items.length); }, 1);
        h += '<h2 class="sec">Lines with no Hebrew before them <small>' + lr.total.toLocaleString() + '</small></h2><div class="card pad0 hits">' + lr.groups.slice(0, ll).map(function (gr) {
          var kinds = [];
          if (gr.kinds.o) kinds.push('opens a paragraph ' + gr.kinds.o);
          if (gr.kinds.s) kinds.push('own line ' + gr.kinds.s);
          if (gr.kinds.t) kinds.push('in a table ' + gr.kinds.t);
          var sn = lineSnip(gr.text, q);
          return '<div class="hit"><div class="txt" style="color:var(--ink);font-size:16px">' + (sn.cut ? hl(sn.text, words) : richHtml(gr.rich, gr.text, words)) + '</div>' +
            (lr.groups.length > 1 ? '<div style="display:flex;align-items:center;gap:8px"><div class="bar" style="flex:1;height:8px"><i style="width:' + Math.max(3, Math.round(gr.items.length / lmax * 100)) + '%;background:var(--gold)"></i></div><span class="note">' + gr.items.length + '×</span></div>' : '<div class="note">' + gr.items.length + '×</div>') +
            '<div class="note">' + esc(kinds.join(' · ')) + '</div>' + citeChips(gr.items, 'lctx', 'l:' + gr.key) + '</div>';
        }).join('') + '</div>';
        if (lr.groups.length > ll) h += '<button class="btn ghost block" data-act="morelines">Show more lines (' + (lr.groups.length - ll).toLocaleString() + ' more wordings)</button>';
        if (res.groups.length) h += '<h2 class="sec">Phrases with this wording</h2>';
      }
      var card = function (gr) {
        return '<div class="hit"><div class="txt" style="color:var(--ink);font-size:16px">' + richHtml(gr.rich, gr.eng, words) + '</div>' +
          '<div class="he" lang="he" dir="rtl" style="font-size:16px;color:var(--ink-2)">' + esc(gr.heb) + '</div>' +
          '<div class="note">' + gr.hits.length + '×</div>' + citeChips(gr.hits, 'ctx', 'p:' + gr.key) + '</div>';
      };
      if (ex.length) h += '<h2 class="sec">Renderings of exactly this phrase</h2><div class="card pad0 hits">' + ex.slice(0, limit).map(card).join('') + '</div>';
      if (rest.length) h += (ex.length ? '<h2 class="sec">Inside longer phrases</h2>' : '') + '<div class="card pad0 hits">' + rest.slice(0, limit).map(card).join('') + '</div>';
      if (ex.length > limit || rest.length > limit) h += '<button class="btn ghost block" data-act="morelimit">Show more</button>';
      setView(h);
      return;
    }
    // headers
    if (hasHe(q)) { setView('<div class="card empty">Headers are in English. Type English words.</div>'); return; }
    var hsets = names.map(function (n) { return (S.hdr || {})[n] || []; });
    var groups = headerSearch(hsets, q);
    if (!groups.length) { setView('<div class="card empty">No header like “' + esc(q) + '”.</div>'); return; }
    var max = groups.reduce(function (a, g) { return Math.max(a, g.items.length); }, 1);
    var nwords = q ? enNorm(q).split(' ') : [];
    h += '<p class="note">' + (q ? groups.length.toLocaleString() + ' wordings match.' : 'All headers, most used first.') + '</p><div class="card pad0 hits">' + groups.slice(0, limit).map(function (gr) {
      var kinds = [];
      if (gr.kinds.o) kinds.push('opens a paragraph ' + gr.kinds.o);
      if (gr.kinds.m) kinds.push('mid-paragraph ' + gr.kinds.m);
      if (gr.kinds.s) kinds.push('own line ' + gr.kinds.s);
      return '<div class="hit"><div class="txt" style="color:var(--ink);font-size:16px;font-weight:600">' + richHtml(gr.rich, gr.text, nwords) + '</div>' +
        '<div style="display:flex;align-items:center;gap:8px"><div class="bar" style="flex:1;height:8px"><i style="width:' + Math.max(3, Math.round(gr.items.length / max * 100)) + '%;background:var(--gold)"></i></div><span class="note">' + gr.items.length + '×</span></div>' +
        '<div class="note">' + esc(kinds.join(' · ')) + '</div>' + citeChips(gr.items, 'hctx', 'h:' + gr.key) + '</div>';
    }).join('') + '</div>';
    if (groups.length > limit) h += '<button class="btn ghost block" data-act="morelimit">Show more</button>';
    setView(h);
  }
  async function amudLink(m, daf) {
    S.amudUrl = S.amudUrl || {};
    var k = m + ' ' + daf;
    if (S.amudUrl[k] !== undefined) return S.amudUrl[k];
    try {
      var it = await g('/drives/' + (await driveId()) + '/root:/' + encodeURIComponent(C.resourcesFolder) + '/Corpus/Amudim/' + encodeURIComponent(m) + '/' + encodeURIComponent(k + '.docx') + '?$select=webUrl');
      S.amudUrl[k] = it.webUrl;
    } catch (e) { S.amudUrl[k] = null; }
    return S.amudUrl[k];
  }
  async function openContext(m, i, span, headerText, dafHint, label) {
    var w = $('#sheet');
    w.innerHTML = '<div class="sheet" role="dialog" aria-modal="true" aria-label="In context"><p class="note">Loading…</p></div>';
    w.hidden = false;
    var rows;
    try { rows = await loadCorpus(m); } catch (e) { w.innerHTML = '<div class="sheet">' + errBox(e) + '<button class="btn ghost block" data-act="closesheet">Close</button></div>'; return; }
    /* a header or line at the very end of an amud points at the next amud's first chunk: step back into its own amud */
    var atEnd = false;
    if (dafHint && rows[i] && rows[i].daf !== dafHint && rows[i - 1] && rows[i - 1].daf === dafHint) { i--; atEnd = true; }
    var r = rows[i];
    if (!r) { closeSheet(); toast('That place is no longer in the corpus.'); return; }
    var from = i, to = i + span - 1;
    for (var a = 0; a < 3 && from > 0 && rows[from - 1].daf === r.daf && rows[from - 1].part === r.part; a++) from--;
    for (var b = 0; b < 3 && to < rows.length - 1 && rows[to + 1].daf === r.daf && rows[to + 1].part === r.part; b++) to++;
    var body = '';
    var hrow = headerText ? '<div class="ctxrow hl"><span class="eyebrow">' + esc(label || 'Header') + '</span><div>' + esc(headerText) + '</div></div>' : '';
    for (var k = from; k <= to; k++) {
      var x = rows[k], on = !headerText && k >= i && k < i + span;
      if (!atEnd && k === i) body += hrow; /* the header or line sits just before the chunk it opens */
      body += '<div class="ctxrow' + (on ? ' hl' : '') + '"><div class="he" lang="he" dir="rtl">' + esc(x.heb) + '</div><div>' + richHtml(x.rich, x.eng) + '</div></div>';
      if (atEnd && k === i) { body += hrow; break; }
    }
    var title = m + ' ' + r.daf + (r.part > 1 ? ' (' + r.part + ')' : '');
    w.innerHTML = '<div class="sheet" role="dialog" aria-modal="true" aria-label="' + esc(title) + ' in context"><h2>' + esc(title) + '</h2>' +
      (from > 0 && rows[from - 1].daf === r.daf ? '<p class="note">…</p>' : '') + '<div class="ctx">' + body + '</div>' +
      (to < rows.length - 1 && rows[to + 1].daf === r.daf ? '<p class="note">…</p>' : '') +
      '<div id="amudlink"></div><button class="btn ghost block" data-act="closesheet">Close</button></div>';
    var url = await amudLink(m, r.daf);
    var slot = $('#amudlink');
    if (slot && url) slot.innerHTML = '<a class="btn block" href="' + esc(url) + '" target="_blank" rel="noopener">Open ' + esc(m + ' ' + r.daf) + '</a>';
  }
  function renderGuideResults() {
    var mode = guideMode();
    if (mode.kind !== 'doc') return;
    var doc = S.docs[mode.key];
    if (!doc) return;
    var q = S.guideQ.trim();
    var foot = '<p class="note">From <a href="' + esc(doc.file.webUrl) + '" target="_blank" rel="noopener">' + esc(doc.file.name) + '</a>, read live from Resources.</p>';
    if (!q) {
      var heads = doc.entries.filter(function (e) { return e.kind === 'h' && !SKIP_HEADS.test(e.text); }).slice(0, 40);
      var intro = mode.table ? '<div class="card empty">Type a word in Hebrew or English.</div>'
        : '<h2 class="sec">Sections</h2><div class="card pad0 navlist">' + (heads.length ? heads.map(function (e) { return '<button data-act="ghead" data-t="' + esc(e.text) + '"><span class="grow">' + esc(e.text) + '</span>' + ICON.chev + '</button>'; }).join('') : '<div class="empty">Type to search.</div>') + '</div>';
      setView(intro + foot);
      return;
    }
    var hits = searchDoc(doc, q, mode);
    var words = norm(q).split(/\s+/).filter(Boolean);
    var h = '<div class="card pad0 hits">' + (hits.length ? hits.map(function (x) {
      var e = x.e;
      if (mode.table && e.cells) {
        return '<div class="hit"><div class="cells">' + e.cells.map(function (c) { return '<span class="c' + (hasHe(c) ? ' he" lang="he' : '') + '">' + hl(c, words) + '</span>'; }).join('') + '</div></div>';
      }
      return '<div class="hit">' + (e.heading ? '<span class="ctx">' + esc(e.heading) + '</span>' : '') + '<span class="txt"' + (hasHe(e.text) && !/[a-z]/i.test(e.text) ? ' dir="rtl" lang="he"' : '') + '>' + hl(e.text, words) + '</span></div>';
    }).join('') : '<div class="empty">Nothing found for “' + esc(q) + '”.</div>') + '</div>';
    setView(h + foot);
  }

  function screenMore() {
    frame({ title: 'More', tab: 'more', kicker: 'MDY Aleh Zayis Team' });
    setView('<div class="card" style="display:flex;align-items:center;gap:12px"><div style="width:46px;height:46px;border-radius:999px;background:var(--navy);color:var(--gold);display:flex;align-items:center;justify-content:center;font-weight:700">' + esc(initials()) + '</div><div><b>' + esc((S.me && S.me.displayName) || '') + '</b><div class="note">Signed in with your Aleh Zayis account</div></div></div>' +
      '<nav class="card pad0 navlist" aria-label="More pages">' +
      '<a href="#/ask">' + ICON.ask + '<span class="grow">Ask Chaim Levy</span><span class="aside">AI assistant</span>' + ICON.chev + '</a>' +
      '<a href="#/deadlines">' + ICON.cal + '<span class="grow">Deadlines</span>' + ICON.chev + '</a>' +
      '<a href="#/graphics">' + ICON.pic + '<span class="grow">Graphics</span>' + ICON.chev + '</a>' +
      '<a href="#/files">' + ICON.folder + '<span class="grow">Files</span><span class="aside">the team library</span>' + ICON.chev + '</a>' +
      '<a href="#/feedback">' + ICON.pen + '<span class="grow">Feedback</span>' + ICON.chev + '</a>' +
      '<a href="#/feedback/app">' + ICON.pen + '<span class="grow">Feedback about this app</span>' + ICON.chev + '</a>' +
      '<a href="#/contacts">' + ICON.person + '<span class="grow">Contacts</span>' + ICON.chev + '</a>' +
      '<a href="' + esc(C.chatUrl) + '" target="_blank" rel="noopener">' + ICON.chat + '<span class="grow">Team chat</span><span class="aside">Google Chat</span></a>' +
      '<a href="' + esc(C.siteUrl) + '" target="_blank" rel="noopener">' + ICON.ext + '<span class="grow">Open the full SharePoint site</span></a>' +
      '</nav><button class="btn ghost block" data-act="signout">' + ICON.out.replace('<svg ', '<svg style="width:20px;height:20px;fill:none;stroke:currentColor;stroke-width:2;stroke-linecap:round;stroke-linejoin:round" ') + 'Sign out</button>' +
      '<p class="slogan">It’s not about the Daf, it’s about the Yomi!</p>');
  }

  async function screenDeadlines() {
    var sub = '<div class="chips" role="group" aria-label="Show">' + ['All', 'Mine'].map(function (x) {
      return '<button class="chipbtn" data-act="dfilter" data-k="' + x + '" aria-pressed="' + (S.deadlineFilter === x) + '">' + x + '</button>';
    }).join('') + '</div>';
    frame({ title: 'Deadlines', tab: 'more', back: '#/more', backLabel: 'More', refresh: true, sub: sub });
    loading();
    var all = await loadDeadlines();
    var list = S.deadlineFilter === 'Mine' ? all.filter(deadlineIsMine) : all;
    var up = list.filter(function (d) { return d.days == null || d.days >= 0; });
    var past = list.filter(function (d) { return d.days != null && d.days < 0; });
    var h = up.length ? up.map(deadlineCard).join('') : '<div class="card empty">No upcoming deadlines' + (S.deadlineFilter === 'Mine' ? ' with your name on them' : '') + '.</div>';
    if (past.length) h += '<details class="card"><summary style="cursor:pointer;font-weight:700;min-height:32px">Past deadlines (' + past.length + ')</summary><div style="display:flex;flex-direction:column;gap:10px;margin-top:10px">' + past.reverse().map(deadlineCard).join('') + '</div></details>';
    h += '<p class="note">Deadlines are added on the site’s Deadlines list.</p>';
    setView(h);
  }

  async function screenGraphics(stageName) {
    loading();
    var ms = [];
    for (var i = 0; i < C.masechtos.length; i++) {
      var mf = await masechtaFolder(C.masechtos[i]);
      if (mf && await subFolder(mf, 'Graphics')) ms.push(C.masechtos[i]);
    }
    var m = S.graphicsM && ms.indexOf(S.graphicsM) > -1 ? S.graphicsM : ms[0];
    S.graphicsM = m;
    var sub = '<div class="chips scroll" role="group" aria-label="Masechta">' + ms.map(function (x) {
      return '<button class="chipbtn" data-act="gm" data-m="' + esc(x) + '" aria-pressed="' + (x === m) + '">' + esc(x) + '</button>';
    }).join('') + '</div>';
    if (!m) { frame({ title: 'Graphics', tab: 'more', back: '#/more', backLabel: 'More' }); setView('<div class="card empty">No Graphics folders found.</div>'); return; }
    var gf = await subFolder(await masechtaFolder(m), 'Graphics');
    var items = await kids(gf.id);
    if (stageName) {
      var sf = items.filter(function (x) { return x.folder && x.name === stageName; })[0];
      frame({ title: stageName, tab: 'more', back: '#/graphics', backLabel: m + ' graphics', refresh: true });
      if (!sf) { setView('<div class="card empty">That folder is gone.</div>'); return; }
      var files = (await kids(sf.id, '&$expand=thumbnails&$select=id,name,file,folder,webUrl')).filter(function (x) { return x.file; });
      setView(files.length ? '<div class="thumbs">' + files.map(function (f) {
        var t = f.thumbnails && f.thumbnails[0] && (f.thumbnails[0].large || f.thumbnails[0].medium);
        return '<a class="thumb" href="' + esc(f.webUrl) + '" target="_blank" rel="noopener"><div class="img"' + (t ? ' style="background-image:url(\'' + esc(t.url) + '\')"' : '') + '>' + (t ? '' : esc((f.name.split('.').pop() || '').toUpperCase())) + '</div><span>' + esc(f.name) + '</span></a>';
      }).join('') + '</div>' : '<div class="card empty">This folder is empty.</div>');
      return;
    }
    frame({ title: 'Graphics', tab: 'more', back: '#/more', backLabel: 'More', refresh: true, sub: sub });
    var folders = items.filter(function (x) { return x.folder && x.name[0] !== '.' && !/^archive$/i.test(x.name); });
    folders.sort(function (a, b) {
      var ia = C.graphicsStages.indexOf(a.name), ib = C.graphicsStages.indexOf(b.name);
      return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib) || (a.name < b.name ? -1 : 1);
    });
    var sheets = items.filter(function (x) { return x.file && /\.(xlsx|docx)$/i.test(x.name) && x.name[0] !== '.'; });
    var h = '<p class="note">' + esc(m) + ' pictures by stage. Each stage is a folder in ' + esc(gf.name) + '.</p><div class="card pad0 navlist">' +
      folders.map(function (f) {
        return '<a href="#/graphics/' + encodeURIComponent(f.name) + '"><span class="grow">' + esc(f.name) + '</span><span class="aside">' + (f.folder.childCount || 0) + '</span>' + ICON.chev + '</a>';
      }).join('') + '</div>';
    if (sheets.length) h += '<div class="card pad0 navlist">' + sheets.map(function (f) { return '<a href="' + esc(f.webUrl) + '" target="_blank" rel="noopener">' + ICON.ext + '<span class="grow">' + esc(f.name) + '</span></a>'; }).join('') + '</div>';
    setView(h);
  }

  var APP_CAT = 'Team App';
  function appChoice(cats) { return cats.filter(function (c) { return /team app|the app|\bapp\b/i.test(c); })[0] || null; }
  async function screenFeedback(arg) {
    frame({ title: 'Feedback', tab: 'more', back: '#/more', backLabel: 'More' });
    loading();
    var fb = await feedbackList();
    var siteCats = fb.cat ? fb.cat.choice.choices : [];
    var cats = siteCats.slice();
    if (!appChoice(cats)) cats.unshift(APP_CAT);
    if (arg === 'app') S.fbCat = appChoice(cats);
    S.fbCat = S.fbCat && cats.indexOf(S.fbCat) > -1 ? S.fbCat : (cats.filter(function (c) { return /general/i.test(c); })[0] || cats[0] || '');
    var h = '<p class="note">Signed with your name. Other team members can’t see your entries.</p>';
    if (cats.length) h += '<span class="lbl">What is it about?</span><div class="chips" role="group" aria-label="Category">' + cats.map(function (c) {
      return '<button class="chipbtn" data-act="fbcat" data-c="' + esc(c) + '" aria-pressed="' + (c === S.fbCat) + '">' + esc(c) + '</button>';
    }).join('') + '</div>';
    h += '<div class="field"><label class="lbl" for="fbText">Your feedback</label><textarea id="fbText" rows="7">' + esc(S.fbDraft || '') + '</textarea></div>' +
      '<button class="btn block" data-act="fbsend">Send feedback</button>';
    setView(h);
    $('#fbText').addEventListener('input', function (e) { S.fbDraft = e.target.value; });
  }
  async function sendFeedback() {
    var text = (S.fbDraft || '').trim();
    if (!text) { toast('Write something first.'); return; }
    var fb = await feedbackList(), fields = {};
    fields.Title = text.replace(/\s+/g, ' ').slice(0, 80) + (text.length > 80 ? '…' : '');
    var siteCats = fb.cat ? fb.cat.choice.choices : [];
    var cat = S.fbCat;
    if (cat === APP_CAT && siteCats.indexOf(APP_CAT) < 0) {
      // the list has no app category yet: file it under the website category, marked
      cat = siteCats.filter(function (c) { return /website/i.test(c); })[0] || siteCats.filter(function (c) { return /general/i.test(c); })[0] || siteCats[0];
      fields.Title = '[Team App] ' + fields.Title;
      text = '[Team App] ' + text;
    }
    if (fb.cat && cat) fields[fb.cat.name] = cat;
    if (fb.body) fields[fb.body.name] = text;
    else fields.Title = text.slice(0, 255);
    var b = $('[data-act="fbsend"]'); if (b) { b.disabled = true; b.textContent = 'Sending…'; }
    try {
      await g('/sites/' + (await siteId()) + '/lists/' + fb.id + '/items', { method: 'POST', body: { fields: fields } });
      S.fbDraft = '';
      setView('<div class="done"><div class="tick">' + ICON.tick + '</div><h2>Thank you</h2><p>Logged under ' + esc(S.fbCat || 'Feedback') + '.</p><button class="btn ghost block" data-act="fbmore">Send more</button></div>');
    } catch (e) {
      if (b) { b.disabled = false; b.textContent = 'Send feedback'; }
      view().insertAdjacentHTML('afterbegin', errBox(e));
    }
  }

  /* ---------- files: the team library itself ---------- */
  function fmtSize(n) {
    if (n == null) return '';
    if (n < 1024) return n + ' B';
    if (n < 1048576) return Math.round(n / 1024) + ' KB';
    return (n / 1048576).toFixed(1) + ' MB';
  }
  function fileIcon(name, isFolder) {
    if (isFolder) return '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 6.5A1.5 1.5 0 0 1 4.5 5H9l2 2.5h8.5A1.5 1.5 0 0 1 21 9v9.5a1.5 1.5 0 0 1-1.5 1.5h-15A1.5 1.5 0 0 1 3 18.5z" fill="#f6e6c6" stroke="#7a5c14"/></svg>';
    var ext = (name.split('.').pop() || '').toLowerCase();
    var col = /docx?$/.test(ext) ? '#1d3d6b' : /xlsx?|csv/.test(ext) ? '#24552a' : /pdf/.test(ext) ? '#a3271b' : /png|jpe?g|svg|gif|webp/.test(ext) ? '#7b1fa2' : '#5b6470';
    return '<svg viewBox="0 0 24 24" aria-hidden="true" style="stroke:' + col + '"><path d="M14 3H6.5A1.5 1.5 0 0 0 5 4.5v15A1.5 1.5 0 0 0 6.5 21h11a1.5 1.5 0 0 0 1.5-1.5V8z"/><path d="M14 3v5h5"/></svg>';
  }
  function fileRow(x, showPath) {
    var isF = !!x.folder;
    var when = x.lastModifiedDateTime ? fmtDate(x.lastModifiedDateTime) : '';
    var who = x.lastModifiedBy && x.lastModifiedBy.user ? x.lastModifiedBy.user.displayName : '';
    var meta = isF ? ((x.folder.childCount || 0) + ' items') : [fmtSize(x.size), when, who].filter(Boolean).join(' · ');
    var path = showPath && x.parentReference && x.parentReference.path ? decodeURIComponent(x.parentReference.path.replace(/^.*root:\/?/, '')) : '';
    var href = isF ? '#/files/' + encodeURIComponent(((path ? path + '/' : (S.filePath ? S.filePath + '/' : '')) + x.name)) : x.webUrl;
    return '<a href="' + esc(href) + '"' + (isF ? '' : ' target="_blank" rel="noopener"') + '>' + fileIcon(x.name, isF) +
      '<span class="grow" style="min-width:0"><span style="display:block;overflow:hidden;text-overflow:ellipsis">' + esc(x.name) + '</span>' +
      '<span class="aside" style="display:block;font-weight:400">' + esc(path ? path + ' · ' : '') + esc(meta) + '</span></span>' + (isF ? ICON.chev : '') + '</a>';
  }
  async function screenFiles(path) {
    path = (path || '').replace(/^\/+|\/+$/g, '');
    S.filePath = path;
    var parts = path ? path.split('/') : [];
    var crumbs = '<div class="crumbs"><a href="#/files">Library</a>' + parts.map(function (p, k) {
      return ' › <a href="#/files/' + encodeURIComponent(parts.slice(0, k + 1).join('/')) + '">' + esc(p) + '</a>';
    }).join('') + '</div>';
    var sub = '<label class="sr" for="fq">Search the library</label><input id="fq" type="search" placeholder="Search the whole library" value="' + esc(S.fileQ || '') + '" autocomplete="off">';
    frame({ title: parts.length ? parts[parts.length - 1] : 'Files', tab: 'more', back: parts.length ? '#/files/' + encodeURIComponent(parts.slice(0, -1).join('/')) : '#/more', backLabel: parts.length > 1 ? parts[parts.length - 2] : parts.length ? 'Library' : 'More', refresh: true, sub: sub });
    var fq = $('#fq');
    fq.addEventListener('keydown', function (e) { if (e.key === 'Enter') { S.fileQ = fq.value.trim(); runFileSearch(); } });
    fq.addEventListener('input', function () { S.fileQ = fq.value; clearTimeout(screenFiles._t); screenFiles._t = setTimeout(runFileSearch, 500); });
    if (S.fileQ && S.fileQ.trim().length > 1) { runFileSearch(); return; }
    loading();
    var d = await driveId();
    var url = path ? '/drives/' + d + '/root:/' + path.split('/').map(encodeURIComponent).join('/') + ':/children' : '/drives/' + d + '/root/children';
    var items = await gAll(url + '?$top=500&$select=id,name,folder,file,size,webUrl,lastModifiedDateTime,lastModifiedBy');
    var folders = items.filter(function (x) { return x.folder; }), files = items.filter(function (x) { return !x.folder; });
    var byName = function (a, b) { return a.name.localeCompare(b.name, undefined, { numeric: true }); };
    var byNew = function (a, b) { return (b.lastModifiedDateTime || '') < (a.lastModifiedDateTime || '') ? -1 : 1; };
    folders.sort(byName);
    files.sort(S.fileSort === 'new' ? byNew : byName);
    var here = path ? null : S.drive;
    var h = crumbs + '<div style="display:flex;justify-content:space-between;align-items:center"><span class="note">' + folders.length + ' folders · ' + files.length + ' files</span>' +
      '<button class="chipbtn" data-act="fsort">' + (S.fileSort === 'new' ? 'Newest first' : 'A–Z') + '</button></div>';
    h += items.length ? '<div class="card pad0 navlist files">' + folders.concat(files).map(function (x) { return fileRow(x, false); }).join('') + '</div>' : '<div class="card empty">This folder is empty.</div>';
    var webUrl = here ? here.webUrl : (C.siteUrl + '/Shared Documents/' + path);
    h += '<p class="note"><a href="' + esc(webUrl) + '" target="_blank" rel="noopener">Open this folder in SharePoint</a> to upload, move or rename.</p>';
    setView(h);
  }
  async function runFileSearch() {
    var q = (S.fileQ || '').trim();
    if (currentRoute().indexOf('/files') !== 0) return;
    if (q.length < 2) { if (!q) screenFiles(S.filePath); return; }
    loading('Searching…');
    try {
      var d = await driveId();
      var r = await g('/drives/' + d + "/root/search(q='" + encodeURIComponent(q.replace(/'/g, "''")) + "')?$top=60&$select=id,name,folder,file,size,webUrl,lastModifiedDateTime,lastModifiedBy,parentReference");
      var items = (r.value || []).filter(function (x) { return !/^Archive\b/.test(((x.parentReference && x.parentReference.path) || '').replace(/^.*root:\/?/, '')) || /archive/i.test(q); });
      setView('<p class="note">' + items.length + ' found for “' + esc(q) + '”.</p>' + (items.length ? '<div class="card pad0 navlist files">' + items.map(function (x) { return fileRow(x, true); }).join('') + '</div>' : ''));
    } catch (e) { fail(e, runFileSearch); }
  }

  /* contacts: the site's own Contacts page, read live */
  var ALLOWED = { TABLE: 1, TBODY: 1, THEAD: 1, TR: 1, TD: 1, TH: 1, P: 1, STRONG: 1, B: 1, EM: 1, I: 1, A: 1, BR: 1, H2: 1, H3: 1, H4: 1, UL: 1, OL: 1, LI: 1, SPAN: 1 };
  function cleanHtml(html) {
    var tpl = document.createElement('template');
    tpl.innerHTML = html;
    (function walk(n) {
      Array.prototype.slice.call(n.childNodes).forEach(function (c) {
        if (c.nodeType === 8) { c.remove(); return; }
        if (c.nodeType !== 1) return;
        if (!ALLOWED[c.tagName]) {
          if (/^(SCRIPT|STYLE|IFRAME|OBJECT|EMBED|svg)$/i.test(c.tagName)) { c.remove(); return; }
          var frag = document.createDocumentFragment();
          while (c.firstChild) frag.appendChild(c.firstChild);
          c.replaceWith(frag); walk(n); return;
        }
        Array.prototype.slice.call(c.attributes).forEach(function (a) {
          if (c.tagName === 'A' && a.name === 'href' && /^(https?:|mailto:|tel:)/i.test(a.value)) return;
          c.removeAttribute(a.name);
        });
        if (c.tagName === 'A') { c.setAttribute('target', '_blank'); c.setAttribute('rel', 'noopener'); }
        walk(c);
      });
    })(tpl.content);
    $$('table', tpl.content).forEach(function (t) { var w = document.createElement('div'); w.className = 'tbl'; t.replaceWith(w); w.appendChild(t); });
    return tpl.innerHTML;
  }
  async function screenContacts() {
    frame({ title: 'Contacts', tab: 'more', back: '#/more', backLabel: 'More', refresh: true });
    loading();
    var sid = await siteId();
    var pages = await gAll('/sites/' + sid + '/pages/microsoft.graph.sitePage?$select=id,name,title,webUrl');
    var p = pages.filter(function (x) { return x.name === C.contactsPage; })[0];
    if (!p) throw new Error('The Contacts page was not found on the site.');
    var full = await g('/sites/' + sid + '/pages/' + p.id + '/microsoft.graph.sitePage?$expand=canvasLayout');
    var parts = [];
    var lay = full.canvasLayout || {};
    (lay.horizontalSections || []).forEach(function (sec) {
      (sec.columns || []).forEach(function (col) {
        (col.webparts || []).forEach(function (wp) { if (wp.innerHtml) parts.push(wp.innerHtml); });
      });
    });
    if (lay.verticalSection) (lay.verticalSection.webparts || []).forEach(function (wp) { if (wp.innerHtml) parts.push(wp.innerHtml); });
    setView('<div class="card page-html">' + (parts.length ? parts.map(cleanHtml).join('') : '<div class="empty">The page has no text to show.</div>') + '</div>' +
      '<p class="note"><a href="' + esc(p.webUrl) + '" target="_blank" rel="noopener">Open the Contacts page on the site</a></p>');
  }

  /* ---------- router ---------- */
  var routing = 0;
  async function route() {
    var id = ++routing;
    closeSheet();
    closeMenu();
    var r = currentRoute().replace(/^\/?/, '').split('/');
    var name = r[0] || 'home', arg = r[1] ? decodeURIComponent(r[1]) : null;
    window.scrollTo(0, 0);
    try {
      if (name === 'home') await screenHome();
      else if (name === 'tracker') await screenTracker(arg);
      else if (name === 'submit') await screenSubmit(arg);
      else if (name === 'guide') await screenGuide();
      else if (name === 'more') screenMore();
      else if (name === 'deadlines') await screenDeadlines();
      else if (name === 'graphics') await screenGraphics(arg);
      else if (name === 'feedback') await screenFeedback(arg);
      else if (name === 'files') { if (!arg) S.fileQ = ''; await screenFiles(r.slice(1).map(decodeURIComponent).join('/')); }
      else if (name === 'contacts') await screenContacts();
      else if (name === 'ask') screenAsk();
      else { location.hash = '#/home'; return; }
    } catch (e) {
      if (id === routing) fail(e, route);
      if (window.console) console.error(e);
    }
  }
  async function refresh() {
    S.tracker = null; S.deadlines = null; S.rootKids = null; S.folders = {}; S.cidx = null; S.corp = {}; S.hdr = {}; S.feedback = null; S.rulesIx = null;
    await route();
    toast('Up to date');
  }

  /* ---------- events ---------- */
  document.addEventListener('click', async function (ev) {
    var el = ev.target.closest('[data-act]');
    if (!el) return;
    var act = el.getAttribute('data-act');
    var rowById = function (id) { return (S.tracker || []).filter(function (r) { return r.id === id; })[0]; };
    if (act === 'menu') { if (window.matchMedia('(max-width: 899px)').matches) { ev.preventDefault(); openMenu(); } return; }
    if (act === 'closemenu') { closeMenu(); return; }
    if (act === 'retry') { if (retryFn) retryFn(); return; }
    if (act === 'signin') { signIn(); return; }
    if (act === 'signout') { signOut(); return; }
    if (act === 'star') {
      ev.preventDefault();
      toggleStar(el.getAttribute('data-id'));
      el.setAttribute('aria-pressed', String(isStarred(el.getAttribute('data-id'))));
      toast(isStarred(el.getAttribute('data-id')) ? 'Kept in My work' : 'Removed from My work');
      return;
    }
    if (act === 'move') { var row = rowById(el.getAttribute('data-id')); if (row) openMoveSheet(row); return; }
    if (act === 'closesheet') { closeSheet(); return; }
    if (act === 'pickstage') {
      var w = $('#sheet'), rw = w._row, st = el.getAttribute('data-s');
      if (!rw || st === rw.stage) { closeSheet(); return; }
      $$('.opts button', w).forEach(function (b) { b.disabled = true; });
      el.textContent = 'Saving…';
      try { await moveStage(rw, st); closeSheet(); toast(rw.title + ' moved to ' + st); route(); }
      catch (e) { closeSheet(); toast('Could not move it: ' + e.message); }
      return;
    }
    if (act === 'tm') { location.hash = '#/tracker/' + encodeURIComponent(el.getAttribute('data-m')); return; }
    if (act === 'gm') { S.graphicsM = el.getAttribute('data-m'); route(); return; }
    if (act === 'dfilter') { S.deadlineFilter = el.getAttribute('data-k'); route(); return; }
    if (act === 'gmode') { S.guideMode = el.getAttribute('data-k'); S.expanded = {}; S.limit = 25; route(); return; }
    if (act === 'cfilter') { S.cFilter = el.getAttribute('data-m'); S.expanded = {}; S.llimit = 6; route(); return; }
    if (act === 'expand') { S.expanded = S.expanded || {}; S.expanded[el.getAttribute('data-k')] = true; renderIndexResults(); return; }
    if (act === 'morelimit') { S.limit = (S.limit || 25) + 25; renderIndexResults(); return; }
    if (act === 'ctx') { openContext(el.getAttribute('data-m'), +el.getAttribute('data-i'), +el.getAttribute('data-span') || 1); return; }
    if (act === 'hctx') { openContext(el.getAttribute('data-m'), +el.getAttribute('data-i'), 1, el.getAttribute('data-h'), el.getAttribute('data-daf')); return; }
    if (act === 'lctx') { openContext(el.getAttribute('data-m'), +el.getAttribute('data-i'), 1, el.getAttribute('data-h'), el.getAttribute('data-daf'), { o: 'Opens its paragraph', s: 'On its own line', t: 'In a table' }[el.getAttribute('data-kind')] || 'Line'); return; }
    if (act === 'morelines') { S.llimit = (S.llimit || 6) + 30; renderIndexResults(); return; }
    if (act === 'fsort') { S.fileSort = S.fileSort === 'new' ? 'name' : 'new'; route(); return; }
    if (act === 'ghead') { S.guideQ = el.getAttribute('data-t'); var gi = $('#gq'); if (gi) gi.value = S.guideQ; renderGuideResults(); return; }
    if (act === 'fbcat') { S.fbCat = el.getAttribute('data-c'); $$('[data-act="fbcat"]').forEach(function (b) { b.setAttribute('aria-pressed', String(b === el)); }); return; }
    if (act === 'fbsend') { sendFeedback(); return; }
    if (act === 'fbmore') { route(); return; }
    if (act === 'pickfile') { $('#sFile').click(); return; }
    if (act === 'send') { doSend(); return; }
    if (act === 'askresume') { askResume(); return; }
    if (act === 'asknew') { if (askReset()) { var aq = $('#askQ'); if (aq) aq.focus(); } return; }
    if (act === 'asktry') { var tq = el.textContent; if (/…$/.test(tq)) { var at = $('#askQ'); if (at) { at.textContent = tq.replace(/\s*…$/, ' '); S.ask.draft = at.innerHTML; at.focus(); var sl = window.getSelection(), rg = document.createRange(); rg.selectNodeContents(at); rg.collapse(false); sl.removeAllRanges(); sl.addRange(rg); } } else askSubmit(tq); return; }
    if (act === 'again') { S.submit = null; location.hash = '#/submit'; route(); return; }
  });
  document.addEventListener('change', async function (ev) {
    var el = ev.target, s = S.submit;
    if (el.id === 'sFile' && s) { s.file = el.files && el.files[0] || null; s.nameTouched = false; renderSubmit(); return; }
    var act = el.getAttribute && el.getAttribute('data-act');
    if (!s || !act) return;
    if (act === 'samud') { s.rowId = el.value || null; var r = (S.tracker || []).filter(function (x) { return x.id === s.rowId; })[0]; s.stage = r ? nextStage(r.stage, r.masechta) : ''; s.nameTouched = false; s.person = null; renderSubmit(); }
    else if (act === 'sall') { s.showAll = el.checked; renderSubmit(); }
    else if (act === 'sstage') { s.stage = el.value; renderSubmit(); }
    else if (act === 'srole') { s.role = el.value; s.person = null; renderSubmit(); }
    else if (act === 'sperson') { s.person = el.value; renderSubmit(); }
  });
  document.addEventListener('input', function (ev) {
    var s = S.submit, el = ev.target;
    if (!s) return;
    if (el.id === 'sName') { s.name = el.value.trim(); s.nameTouched = true; var v = $('.summary strong.v'); if (v) v.textContent = s.name || '—'; var b = $('[data-act="send"]'); if (b) b.disabled = !(s.file && s.name && s.parentId); }
    if (el.id === 'sNote') s.note = el.value;
  });
  document.addEventListener('keydown', function (ev) {
    if (ev.key !== 'Escape') return;
    if (!$('#drawer-wrap').hidden) closeMenu();
    else if (!$('#sheet').hidden) closeSheet();
  });
  function openMenu() {
    var w = $('#drawer-wrap');
    $('#drawer-chat').href = C.chatUrl;
    $('#drawer-site').href = C.siteUrl;
    $('#drawer-who').textContent = (S.me && S.me.displayName) || '';
    var base = currentRoute().replace(/^\/?/, '').split('/');
    $$('.drawer-links a').forEach(function (a) {
      var h = (a.getAttribute('href') || '').replace(/^#\/?/, '');
      if (h && (h === base.join('/') || (h === base[0] && h !== 'feedback'))) a.setAttribute('aria-current', 'page'); else a.removeAttribute('aria-current');
    });
    w.hidden = false;
    requestAnimationFrame(function () { w.classList.add('open'); });
    var first = $('.drawer-links a', w); if (first) first.focus();
  }
  function closeMenu() {
    var w = $('#drawer-wrap');
    if (w.hidden) return;
    w.classList.remove('open');
    setTimeout(function () { w.hidden = true; }, 200);
  }
  $('#refresh').addEventListener('click', refresh);
  window.addEventListener('hashchange', route);

  /* ---------- start ---------- */
  function signInScreen(msg) {
    frame({ title: 'MDY Team', tab: '' });
    $('#tabs').hidden = true;
    setView('<div class="signin"><img src="icons/icon-192.png" alt=""><h2>MDY Aleh Zayis Team</h2><p>Sign in with your Aleh Zayis account to see your amudim, the tracker, the Guide and deadlines.</p>' +
      (msg ? '<div class="err">' + esc(msg) + '</div>' : '') +
      '<button class="btn block" data-act="signin" style="max-width:320px">Sign in</button></div>');
  }
  async function start() {
    /* Microsoft's silent sign-in renewal loads this page in a hidden frame and reads the answer from its address.
       The app must not run there: it would consume or overwrite that answer, and renewal (needed daily) would fail. */
    if (window.self !== window.top) return;
    if ('serviceWorker' in navigator && !TEST && location.protocol === 'https:') {
      navigator.serviceWorker.register('sw.js').catch(function () { /* offline support is optional */ });
    }
    var signedIn;
    try { signedIn = await initAuth(); }
    catch (e) { signInScreen('Sign-in did not finish: ' + (e.errorMessage || e.message)); return; }
    if (!signedIn) { signInScreen(); return; }
    $('#tabs').hidden = false;
    try { await me(); } catch (e) { /* shown by the first screen */ }
    if (!location.hash || /code=|state=/.test(location.hash)) history.replaceState(null, '', '#/home');
    route();
  }
  start();
})();
