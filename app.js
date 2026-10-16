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
  function stageChip(s) {
    var c = stageColor(s);
    return '<span class="stage" style="background:' + c + ';color:' + (LIGHT[c] ? '#1a1814' : '#fff') + '">' + esc(s || '—') + '</span>';
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
  function isMine(r) {
    var sn = surname();
    if (isStarred(r.id)) return true;
    if (!sn || sn.length < 3) return false;
    return (r.writer + ' ' + r.editor).toLowerCase().indexOf(sn) > -1;
  }
  function personLine(r) {
    var bits = [];
    if (r.editor) bits.push('Editor: ' + r.editor);
    else if (r.writer) bits.push('Writer: ' + r.writer);
    if (r.workflow) bits.push(r.workflow + ' workflow');
    return bits.join(' · ');
  }
  function nextStage(s) {
    var list = S.stages || [];
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
        if (v) lines.push({ label: c.displayName, value: v, he: hasHe(v), job: /^job$/i.test(c.displayName) });
      }
      var due = dueCol ? f[dueCol.name] : null;
      return { id: it.id, title: f.Title || '', due: due, days: daysUntil(due), lines: lines };
    }));
    items.sort(function (a, b) { return (a.due || '9') < (b.due || '9') ? -1 : 1; });
    S.deadlines = items;
    return items;
  }
  function deadlineIsMine(d) {
    var sn = surname();
    if (!sn) return false;
    return (d.title + ' ' + d.lines.map(function (l) { return l.value; }).join(' ')).toLowerCase().indexOf(sn) > -1;
  }
  function deadlineCard(d) {
    var job = d.lines.filter(function (l) { return l.job; })[0];
    var heb = d.lines.filter(function (l) { return l.he && !l.job; });
    var other = d.lines.filter(function (l) { return !l.he && !l.job; });
    var cls = d.days == null ? '' : d.days < 0 ? 'late' : d.days <= 7 ? 'soon' : '';
    var daysTxt = d.days == null ? '–' : Math.abs(d.days);
    var daysLbl = d.days == null ? '' : d.days < 0 ? 'days late' : d.days === 1 ? 'day' : 'days';
    if (d.days === 0) { daysTxt = 'Today'; daysLbl = ''; }
    return '<article class="card dl"><div class="days ' + cls + '"><b>' + esc(daysTxt) + '</b><span>' + esc(daysLbl) + '</span></div><div class="dt">' +
      (job ? '<span class="tag">' + esc(job.value) + '</span>' : '') +
      '<b>' + esc(d.title) + '</b>' +
      other.map(function (l) { return '<span>' + esc(l.value) + '</span>'; }).join('') +
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
    { key: 'Spelling', match: /translit/i, ph: 'Hebrew or English word', table: true },
    { key: 'Definitions', match: /definition/i, ph: 'Look up a term', table: true },
    { key: 'Rulings', match: /rulings/i, ph: 'What’s our rule on…', table: false },
    { key: 'Style', match: /style guide/i, ph: 'Search the Style Guide', table: false }
  ];
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
  function searchDoc(doc, q, mode) {
    var nq = norm(q).trim();
    if (!nq) return [];
    var words = nq.split(/\s+/).filter(Boolean);
    var hits = [];
    doc.entries.forEach(function (e) {
      if (e.kind === 'h') return;
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
    document.title = o.title + ' · MDY Team';
    $('#subhead').innerHTML = o.sub || '';
    $('#refresh').hidden = !o.refresh;
    $$('#tabs a').forEach(function (a) {
      if (a.getAttribute('data-tab') === o.tab) a.setAttribute('aria-current', 'page');
      else a.removeAttribute('aria-current');
    });
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
    var mine = rows.filter(isMine);
    var upcoming = dls.filter(function (d) { return d.days != null && d.days >= 0; });
    var myNext = upcoming.filter(deadlineIsMine)[0] || upcoming[0];
    frame({ title: 'My work', tab: 'home', refresh: true, kicker: (S.me && S.me.displayName) || 'MDY Aleh Zayis Team' });
    var h = '';
    if (myNext) h += '<a class="card dl" href="#/deadlines" style="text-decoration:none">' + deadlineCard(myNext).replace(/^<article class="card dl">|<\/article>$/g, '') + '</a>';
    h += '<a class="btn block" href="#/submit">' + ICON.up + 'Submit work</a>';
    h += '<h2 class="sec">Your amudim <small>' + mine.length + '</small></h2>';
    if (mine.length) h += '<div class="rows">' + mine.map(rowHtml).join('') + '</div>';
    else h += '<div class="card empty">Nothing is under your name on the tracker yet. Tap the star on any amud in the Tracker to keep it here.</div>';
    h += '<h2 class="sec">Quick links</h2><div class="card pad0 navlist">' +
      '<a href="#/guide">' + ICON.pen.replace('M4 20l1-4L16 5a2.1 2.1 0 0 1 3 3L8 19z', 'M11 4.5a6.5 6.5 0 1 0 0 13 6.5 6.5 0 0 0 0-13zM16 16l4.5 4.5') + '<span class="grow">Look-ups and style guide</span>' + ICON.chev + '</a>' +
      '<a href="#/deadlines">' + ICON.cal + '<span class="grow">All deadlines</span>' + ICON.chev + '</a>' +
      '<a href="' + esc(C.chatUrl) + '" target="_blank" rel="noopener">' + ICON.chat + '<span class="grow">Team chat</span>' + ICON.ext + '</a></div>';
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
    shown.forEach(function (r) { counts[r.stage] = (counts[r.stage] || 0) + 1; });
    var order = (S.stages || []).filter(function (s) { return counts[s]; }).concat(Object.keys(counts).filter(function (s) { return (S.stages || []).indexOf(s) < 0; }));
    var done = (counts.Final || 0) + (counts.Closed || 0);
    var h = '<section class="card" aria-label="Progress"><div style="display:flex;justify-content:space-between;align-items:baseline;margin-bottom:10px"><b>' + esc(m) + ' · ' + shown.length + ' rows</b><span class="note">' + done + ' final</span></div>' +
      '<div class="bar">' + order.map(function (s) { return '<i style="flex:' + counts[s] + ';background:' + stageColor(s) + '" title="' + esc(s) + '"></i>'; }).join('') + '</div>' +
      '<div class="legend">' + order.map(function (s) { return stageChip(s + ' ' + counts[s]).replace(esc(s + ' ' + counts[s]), esc(s) + ' ' + counts[s]).replace('background:' + stageColor(s + ' ' + counts[s]), 'background:' + stageColor(s)).replace(/color:#[0-9a-f]+"/, 'color:' + (LIGHT[stageColor(s)] ? '#1a1814' : '#fff') + '"'); }).join('') + '</div></section>';
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
    var list = S.stages || Object.keys(C.stageColors);
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
    var s = S.submit;
    if (!s || (rowId && s.rowId !== rowId) || s.sent) {
      var mine = rows.filter(isMine);
      var row0 = rows.filter(function (r) { return r.id === rowId; })[0] || mine[0] || null;
      s = S.submit = { rowId: row0 ? row0.id : null, showAll: !!(row0 && !isMine(row0)), stage: row0 ? nextStage(row0.stage) : '', file: null, name: '', nameTouched: false, role: 'Editors', person: null, note: '', sent: false, busy: false };
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
    var stages = S.stages || [];
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

  async function screenGuide() {
    var mode = GUIDE.filter(function (x) { return x.key === S.guideMode; })[0] || GUIDE[0];
    var sub = '<div class="chips" role="group" aria-label="Look-up">' + GUIDE.map(function (x) {
      return '<button class="chipbtn" data-act="gmode" data-k="' + x.key + '" aria-pressed="' + (x.key === mode.key) + '">' + x.key + '</button>';
    }).join('') + '</div><label class="sr" for="gq">' + esc(mode.ph) + '</label>' +
      '<input id="gq" type="search" dir="auto" placeholder="' + esc(mode.ph) + '" value="' + esc(S.guideQ) + '" autocomplete="off" style="margin-top:10px">';
    frame({ title: 'Guide', tab: 'guide', kicker: 'Style and look-ups', sub: sub });
    var inp = $('#gq');
    inp.addEventListener('input', function () { S.guideQ = inp.value; clearTimeout(screenGuide._t); screenGuide._t = setTimeout(renderGuideResults, 150); });
    loading('Opening the ' + mode.key.toLowerCase() + ' document…');
    try { await loadDoc(mode); } catch (e) { fail(e, screenGuide); return; }
    renderGuideResults();
  }
  function renderGuideResults() {
    var mode = GUIDE.filter(function (x) { return x.key === S.guideMode; })[0] || GUIDE[0];
    var doc = S.docs[mode.key];
    if (!doc) return;
    var q = S.guideQ.trim();
    var foot = '<p class="note">From <a href="' + esc(doc.file.webUrl) + '" target="_blank" rel="noopener">' + esc(doc.file.name) + '</a>, read live from Resources.</p>';
    if (!q) {
      var heads = doc.entries.filter(function (e) { return e.kind === 'h'; }).slice(0, 40);
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
      '<a href="#/deadlines">' + ICON.cal + '<span class="grow">Deadlines</span>' + ICON.chev + '</a>' +
      '<a href="#/graphics">' + ICON.pic + '<span class="grow">Graphics</span>' + ICON.chev + '</a>' +
      '<a href="#/feedback">' + ICON.pen + '<span class="grow">Feedback</span>' + ICON.chev + '</a>' +
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
    var folders = items.filter(function (x) { return x.folder; });
    folders.sort(function (a, b) {
      var ia = C.graphicsStages.indexOf(a.name), ib = C.graphicsStages.indexOf(b.name);
      return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib) || (a.name < b.name ? -1 : 1);
    });
    var sheets = items.filter(function (x) { return x.file; });
    var h = '<p class="note">' + esc(m) + ' pictures by stage. Each stage is a folder in ' + esc(gf.name) + '.</p><div class="card pad0 navlist">' +
      folders.map(function (f) {
        return '<a href="#/graphics/' + encodeURIComponent(f.name) + '"><span class="grow">' + esc(f.name) + '</span><span class="aside">' + (f.folder.childCount || 0) + '</span>' + ICON.chev + '</a>';
      }).join('') + '</div>';
    if (sheets.length) h += '<div class="card pad0 navlist">' + sheets.map(function (f) { return '<a href="' + esc(f.webUrl) + '" target="_blank" rel="noopener">' + ICON.ext + '<span class="grow">' + esc(f.name) + '</span></a>'; }).join('') + '</div>';
    setView(h);
  }

  async function screenFeedback() {
    frame({ title: 'Feedback', tab: 'more', back: '#/more', backLabel: 'More' });
    loading();
    var fb = await feedbackList();
    var cats = fb.cat ? fb.cat.choice.choices : [];
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
    if (fb.cat) fields[fb.cat.name] = S.fbCat;
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
      else if (name === 'feedback') await screenFeedback();
      else if (name === 'contacts') await screenContacts();
      else { location.hash = '#/home'; return; }
    } catch (e) {
      if (id === routing) fail(e, route);
      if (window.console) console.error(e);
    }
  }
  async function refresh() {
    S.tracker = null; S.deadlines = null; S.rootKids = null; S.folders = {};
    await route();
    toast('Up to date');
  }

  /* ---------- events ---------- */
  document.addEventListener('click', async function (ev) {
    var el = ev.target.closest('[data-act]');
    if (!el) return;
    var act = el.getAttribute('data-act');
    var rowById = function (id) { return (S.tracker || []).filter(function (r) { return r.id === id; })[0]; };
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
    if (act === 'gmode') { S.guideMode = el.getAttribute('data-k'); route(); return; }
    if (act === 'ghead') { S.guideQ = el.getAttribute('data-t'); var gi = $('#gq'); if (gi) gi.value = S.guideQ; renderGuideResults(); return; }
    if (act === 'fbcat') { S.fbCat = el.getAttribute('data-c'); $$('[data-act="fbcat"]').forEach(function (b) { b.setAttribute('aria-pressed', String(b === el)); }); return; }
    if (act === 'fbsend') { sendFeedback(); return; }
    if (act === 'fbmore') { route(); return; }
    if (act === 'pickfile') { $('#sFile').click(); return; }
    if (act === 'send') { doSend(); return; }
    if (act === 'again') { S.submit = null; location.hash = '#/submit'; route(); return; }
  });
  document.addEventListener('change', async function (ev) {
    var el = ev.target, s = S.submit;
    if (el.id === 'sFile' && s) { s.file = el.files && el.files[0] || null; s.nameTouched = false; renderSubmit(); return; }
    var act = el.getAttribute && el.getAttribute('data-act');
    if (!s || !act) return;
    if (act === 'samud') { s.rowId = el.value || null; var r = (S.tracker || []).filter(function (x) { return x.id === s.rowId; })[0]; s.stage = r ? nextStage(r.stage) : ''; s.nameTouched = false; s.person = null; renderSubmit(); }
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
  document.addEventListener('keydown', function (ev) { if (ev.key === 'Escape' && !$('#sheet').hidden) closeSheet(); });
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
