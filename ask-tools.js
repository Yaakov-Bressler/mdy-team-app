/* Ask Chaim Levy: the second set of tools (28 Oct 2026).
   find_files, read_file, read_attachment, make_word_file, search_discussions, plus attachments in the question box.
   One module for the team app and the site's Ask Chaim Levy page: everything that touches SharePoint goes through
   env.g (Microsoft Graph AS THE PERSON ASKING), so it can only open what that person can already open.

   Cost guard (YB, 28 Oct 2026: no more than 8x the current API usage). Every relay round resends the whole turn, so
   anything big pulled in early is paid for again on every later round. So:
   - files and attachments are read in parts of PART_CHARS, never whole;
   - an attachment goes in the question as a short preview only; the rest is read_attachment;
   - at most MAX_IMAGES pictures per question, shrunk to IMAGE_PX, and dropped from the history after the answer;
   - make_word_file takes a short list of changed paragraphs, never a rewritten amud;
   - trimTurn() keeps the tool results in one turn under TURN_RESULT_CHARS by stubbing the oldest. */
var ASKY = (function () {
  'use strict';
  var W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  var XMLNS = 'http://www.w3.org/XML/1998/namespace';
  var PART_CHARS = 9000, PREVIEW_CHARS = 2500, MAX_IMAGES = 2, IMAGE_PX = 1200, TURN_RESULT_CHARS = 80000;
  var MAX_EDITS = 60, MAX_EDIT_CHARS = 4000, MAX_ATTACH_BYTES = 15 * 1024 * 1024;
  var AUTHOR = 'Chaim Levy', INITIALS = 'CL';

  function cut(s, n) { s = String(s || ''); return s.length > n ? s.slice(0, n) + '…' : s; }
  function num(v, lo, hi, d) { v = parseInt(v, 10); return isNaN(v) ? d : Math.max(lo, Math.min(hi, v)); }
  var MARKS = /[֑-ׇ]/g;
  function norm(s) {
    return String(s || '').replace(MARKS, '').replace(/[־\-–—]/g, ' ').replace(/[״"“”]/g, '"').replace(/[׳'‘’`]/g, "'")
      .replace(/[.,;:!?()\[\]{}<>*_|]/g, ' ').toLowerCase().replace(/\s+/g, ' ').trim();
  }
  /* letters, nikkud, cantillation, maqaf, geresh: any change at all to the Hebrew counts */
  function hebOnly(s) { return String(s || '').replace(/[^֑-״יִ-ﭏ]/g, ''); }
  function isoNow() { return new Date().toISOString().replace(/\.\d+Z$/, 'Z'); }

  /* ---------- docx reading ---------- */
  function kids(el, local) { var o = []; for (var c = el.firstChild; c; c = c.nextSibling) if (c.nodeType === 1 && c.localName === local && c.namespaceURI === W) o.push(c); return o; }
  function kid(el, local) { return kids(el, local)[0] || null; }
  function within(el, local, stop) { for (var p = el.parentNode; p && p !== stop; p = p.parentNode) if (p.localName === local && p.namespaceURI === W) return true; return false; }
  function onOff(el) { if (!el) return false; var v = el.getAttributeNS(W, 'val'); return !(v === '0' || v === 'false' || v === 'off'); }
  function runItalic(r) { var pr = kid(r, 'rPr'); return !!(pr && onOff(kid(pr, 'i'))); }
  function runText(r, del) {
    var t = '';
    for (var c = r.firstChild; c; c = c.nextSibling) {
      if (c.nodeType !== 1 || c.namespaceURI !== W) continue;
      var n = c.localName;
      if (n === 't' || (del && n === 'delText')) t += c.textContent;
      else if (n === 'tab') t += '\t';
      else if (n === 'br' || n === 'cr') t += '\n';
      else if (n === 'noBreakHyphen') t += '-';
    }
    return t;
  }
  var REV = { ins: 1, del: 1, moveFrom: 1, moveTo: 1, rPrChange: 1, pPrChange: 1 };
  function hasRevisions(p) {
    var all = p.getElementsByTagNameNS(W, '*');
    for (var i = 0; i < all.length; i++) if (REV[all[i].localName]) return true;
    return false;
  }
  function paraModel(p) {
    var runs = p.getElementsByTagNameNS(W, 'r'), out = '', segs = [];
    for (var i = 0; i < runs.length; i++) {
      var r = runs[i];
      if (within(r, 'del', p) || within(r, 'moveFrom', p)) continue;
      var t = runText(r, false);
      if (!t) continue;
      var it = runItalic(r), last = segs[segs.length - 1];
      if (last && last.i === it) last.t += t; else segs.push({ t: t, i: it });
      out += t;
    }
    var marked = segs.map(function (s) {
      if (!s.i || !s.t.trim()) return s.t;
      var m = s.t.match(/^(\s*)([\s\S]*?)(\s*)$/);
      return m[1] + '_' + m[2] + '_' + m[3];
    }).join('');
    return { text: out, marked: marked, rev: hasRevisions(p) };
  }
  function docParas(doc) {
    var body = doc.getElementsByTagNameNS(W, 'body')[0];
    var ps = body ? body.getElementsByTagNameNS(W, 'p') : [];
    var out = [];
    for (var i = 0; i < ps.length; i++) out.push(ps[i]);
    return out;
  }
  async function loadDocx(env, buf) {
    var zip = await env.JSZip.loadAsync(buf);
    var f = zip.file('word/document.xml');
    if (!f) throw new Error('That is not a Word document I can read.');
    var doc = new env.DOMParser().parseFromString(await f.async('string'), 'application/xml');
    var ps = docParas(doc), paras = ps.map(paraModel);
    var comments = [];
    var cf = zip.file('word/comments.xml');
    if (cf) {
      var cdoc = new env.DOMParser().parseFromString(await cf.async('string'), 'application/xml');
      var where = {};
      ps.forEach(function (p, i) {
        var st = p.getElementsByTagNameNS(W, 'commentRangeStart');
        for (var k = 0; k < st.length; k++) { var id = st[k].getAttributeNS(W, 'id'); if (!(id in where)) where[id] = i; }
        var rf = p.getElementsByTagNameNS(W, 'commentReference');
        for (var k2 = 0; k2 < rf.length; k2++) { var id2 = rf[k2].getAttributeNS(W, 'id'); if (!(id2 in where)) where[id2] = i; }
      });
      var cs = cdoc.getElementsByTagNameNS(W, 'comment');
      for (var j = 0; j < cs.length; j++) {
        var c = cs[j], cid = c.getAttributeNS(W, 'id');
        var txt = [], cps = c.getElementsByTagNameNS(W, 'p');
        for (var q = 0; q < cps.length; q++) txt.push(paraModel(cps[q]).text);
        comments.push({ para: where[cid] != null ? where[cid] + 1 : null, by: c.getAttributeNS(W, 'author') || '', text: cut(txt.join(' ').trim(), 600) });
      }
    }
    return { kind: 'docx', zip: zip, paras: paras, comments: comments };
  }
  /* the numbered text Claude reads: ¶N, _italics_ marked, [tracked] where earlier revisions sit (shown as if accepted) */
  function docxLines(d) {
    return d.paras.map(function (p, i) { return p.text.trim() ? '¶' + (i + 1) + (p.rev ? ' [tracked]' : '') + ' ' + p.marked.replace(/\n/g, ' / ') : null; })
      .filter(function (x) { return x != null; });
  }
  function partsOf(lines, size) {
    var parts = [], cur = '';
    lines.forEach(function (l) {
      if (cur && cur.length + l.length + 1 > size) { parts.push(cur); cur = ''; }
      cur += (cur ? '\n' : '') + l;
    });
    if (cur || !parts.length) parts.push(cur);
    return parts;
  }
  function rangeOf(spec, max) {
    var m = String(spec || '').match(/(\d+)\s*(?:[-–]\s*(\d+))?/);
    if (!m) return null;
    var a = num(m[1], 1, max, 1), b = m[2] ? num(m[2], a, max, a) : a;
    return [a, b];
  }

  /* ---------- xlsx reading (first sheet, header row) ---------- */
  async function xlsxRows(env, buf) {
    var zip = await env.JSZip.loadAsync(buf), P = new env.DOMParser();
    var ss = [], sf = zip.file('xl/sharedStrings.xml');
    if (sf) {
      var sdoc = P.parseFromString(await sf.async('string'), 'application/xml');
      var sis = sdoc.getElementsByTagName('si');
      for (var i = 0; i < sis.length; i++) { var ts = sis[i].getElementsByTagName('t'), s = ''; for (var k = 0; k < ts.length; k++) s += ts[k].textContent; ss.push(s); }
    }
    var names = Object.keys(zip.files).filter(function (n) { return /^xl\/worksheets\/sheet\d+\.xml$/.test(n); }).sort();
    if (!names.length) return [];
    var doc = P.parseFromString(await zip.file(names[0]).async('string'), 'application/xml');
    var rows = doc.getElementsByTagName('row'), out = [];
    function col(ref) { var m = String(ref).match(/^([A-Z]+)/); var n = 0; if (!m) return 0; for (var i = 0; i < m[1].length; i++) n = n * 26 + m[1].charCodeAt(i) - 64; return n - 1; }
    for (var r = 0; r < rows.length; r++) {
      var cells = rows[r].getElementsByTagName('c'), row = [];
      for (var c = 0; c < cells.length; c++) {
        var cell = cells[c], t = cell.getAttribute('t'), v = cell.getElementsByTagName('v')[0], val = '';
        if (t === 's' && v) val = ss[parseInt(v.textContent, 10)] || '';
        else if (t === 'inlineStr') { var is = cell.getElementsByTagName('t'); for (var x = 0; x < is.length; x++) val += is[x].textContent; }
        else if (v) val = v.textContent;
        row[col(cell.getAttribute('r'))] = val;
      }
      out.push({ n: parseInt(rows[r].getAttribute('r'), 10) || r + 1, cells: row });
    }
    return out;
  }

  /* ---------- tracked changes ---------- */
  function tokens(s) { return String(s).match(/[A-Za-zÀ-ɏא-ת֑-ׇ'’]+|\d+|\s+|[^\sA-Za-z\dÀ-ɏא-ת֑-ׇ]/g) || []; }
  function parseMarked(s) {
    /* _x_ = italic; anything else roman. Returns [{c, i}] */
    var out = [], it = false, str = String(s || '');
    for (var k = 0; k < str.length; k++) {
      var ch = str[k];
      if (ch === '_') {
        var prev = str[k - 1] || ' ', next = str[k + 1] || ' ';
        var opens = !it && /\S/.test(next) && !/[A-Za-z0-9א-ת]/.test(prev);
        var closes = it && /\S/.test(prev) && !/[A-Za-z0-9א-ת]/.test(next);
        if (opens || closes) { it = !it; continue; }
      }
      out.push({ c: ch, i: it });
    }
    return out;
  }
  function lcs(a, b) {
    var n = a.length, m = b.length, dp = [];
    for (var i = 0; i <= n; i++) dp.push(new Int32Array(m + 1));
    for (var x = n - 1; x >= 0; x--) for (var y = m - 1; y >= 0; y--) dp[x][y] = a[x] === b[y] ? dp[x + 1][y + 1] + 1 : Math.max(dp[x + 1][y], dp[x][y + 1]);
    var ops = [], p = 0, q = 0;
    while (p < n && q < m) {
      if (a[p] === b[q]) { ops.push(['=', p, q]); p++; q++; }
      else if (dp[p + 1][q] >= dp[p][q + 1]) { ops.push(['-', p]); p++; }
      else { ops.push(['+', q]); q++; }
    }
    while (p < n) ops.push(['-', p++]);
    while (q < m) ops.push(['+', q++]);
    return ops;
  }
  var RPR_FIRST = { rStyle: 1, rFonts: 1, b: 1, bCs: 1 };
  function setItalic(doc, rPr, on) {
    ['i', 'iCs'].forEach(function (n) { var e = kid(rPr, n); if (e) rPr.removeChild(e); });
    if (!on) return;
    var after = null;
    for (var c = rPr.firstChild; c; c = c.nextSibling) if (c.nodeType === 1 && RPR_FIRST[c.localName]) after = c;
    var ref = after ? after.nextSibling : rPr.firstChild;
    var i1 = doc.createElementNS(W, 'w:i'), i2 = doc.createElementNS(W, 'w:iCs');
    rPr.insertBefore(i2, ref); rPr.insertBefore(i1, i2);
  }
  var SIMPLE_RUN = { rPr: 1, t: 1, tab: 1, lastRenderedPageBreak: 1 };
  var KEEP_AS_ATOM = { bookmarkStart: 1, bookmarkEnd: 1, commentRangeStart: 1, commentRangeEnd: 1, proofErr: 1, permStart: 1, permEnd: 1 };
  /* one paragraph's content as chars (each with its run) and atoms (bookmarks, comment marks, comment-reference runs) */
  function paraStream(p) {
    var items = [], why = null;
    for (var c = p.firstChild; c; c = c.nextSibling) {
      if (c.nodeType !== 1) continue;
      var n = c.localName;
      if (n === 'pPr') continue;
      if (KEEP_AS_ATOM[n]) { items.push({ atom: c }); continue; }
      if (n !== 'r') { why = 'it has ' + n + ' content'; break; }
      var simple = true;
      for (var e = c.firstChild; e; e = e.nextSibling) if (e.nodeType === 1 && !SIMPLE_RUN[e.localName]) simple = false;
      if (!simple) {
        if (!runText(c, false)) { items.push({ atom: c }); continue; }
        why = 'it has a field, footnote or picture inside the text'; break;
      }
      var t = runText(c, false);
      for (var k = 0; k < t.length; k++) items.push({ c: t[k], run: c });
    }
    return { items: items, why: why };
  }
  function trackedPara(doc, p, newMarked, ids, date) {
    var st = paraStream(p);
    if (st.why) return 'left alone: ' + st.why;
    var chars = st.items.filter(function (x) { return x.c != null; });
    var oldText = chars.map(function (x) { return x.c; }).join('');
    var nw = parseMarked(newMarked), newText = nw.map(function (x) { return x.c; }).join('');
    if (hebOnly(oldText) !== hebOnly(newText)) return 'left alone: the edit changes the Hebrew, and I never touch the Hebrew in a Word file';
    if (/\t/.test(newText) && !/\t/.test(oldText)) return 'left alone: tabs cannot be inserted';
    var oldIt = chars.map(function (x) { return runItalic(x.run); });
    var oT = tokens(oldText), nT = tokens(newText);
    function keyed(toks, itals) {
      var pos = 0;
      return toks.map(function (t) { var it = itals[pos]; pos += t.length; return (/\S/.test(t) && it ? '\u0001' : '') + t; });
    }
    var oK = keyed(oT, oldIt), nK = keyed(nT, nw.map(function (x) { return x.i; }));
    if (oK.join('\u0002') === nK.join('\u0002')) return 'unchanged';
    var oStart = [], acc = 0; oT.forEach(function (t) { oStart.push(acc); acc += t.length; });
    var nStart = [], acc2 = 0; nT.forEach(function (t) { nStart.push(acc2); acc2 += t.length; });
    var ops = lcs(oK, nK);
    // atoms keyed by the char index they sit before
    var atomsAt = {}, ci = 0;
    st.items.forEach(function (x) { if (x.c != null) ci++; else (atomsAt[ci] = atomsAt[ci] || []).push(x.atom); });
    var out = [], pending = [];
    function flushAtoms(upto, into) { for (var k in atomsAt) if (+k <= upto) { atomsAt[k].forEach(function (a) { into.push(a); }); delete atomsAt[k]; } }
    function cloneRun(run, text, tag) {
      var r = doc.createElementNS(W, 'w:r'), pr = kid(run, 'rPr');
      if (pr) r.appendChild(pr.cloneNode(true));
      var t = doc.createElementNS(W, tag); t.setAttributeNS(XMLNS, 'xml:space', 'preserve'); t.textContent = text; r.appendChild(t);
      return r;
    }
    function groups(from, to) { // old char range -> [{run, text}]
      var g = [];
      for (var k = from; k < to; k++) { var x = chars[k], l = g[g.length - 1]; if (l && l.run === x.run) l.text += x.c; else g.push({ run: x.run, text: x.c }); }
      return g;
    }
    function emitEqual(from, to) {
      for (var k = from; k < to; k++) { flushAtoms(k, out); var l = out[out.length - 1], x = chars[k]; if (l && l.__run === x.run && l.__eq) l.__t.textContent += x.c; else { var r = cloneRun(x.run, x.c, 'w:t'); r.__run = x.run; r.__eq = true; r.__t = kid(r, 't'); out.push(r); } }
    }
    function emitDel(from, to) {
      if (from >= to) return;
      var d = doc.createElementNS(W, 'w:del'); d.setAttributeNS(W, 'w:id', String(ids.next++)); d.setAttributeNS(W, 'w:author', AUTHOR); d.setAttributeNS(W, 'w:date', date);
      flushAtoms(from, out);
      groups(from, to).forEach(function (gr) { d.appendChild(cloneRun(gr.run, gr.text, 'w:delText')); });
      out.push(d);
      flushAtoms(to - 1, pending);
    }
    function emitIns(nFrom, nTo, tplCharIdx) {
      if (nFrom >= nTo) return;
      var tpl = chars[Math.max(0, Math.min(chars.length - 1, tplCharIdx))];
      var ins = doc.createElementNS(W, 'w:ins'); ins.setAttributeNS(W, 'w:id', String(ids.next++)); ins.setAttributeNS(W, 'w:author', AUTHOR); ins.setAttributeNS(W, 'w:date', date);
      var seg = null;
      for (var k = nFrom; k < nTo; k++) {
        var x = nw[k];
        if (!seg || seg.i !== x.i) {
          seg = { i: x.i, r: doc.createElementNS(W, 'w:r') };
          var pr = tpl ? kid(tpl.run, 'rPr') : null;
          pr = pr ? pr.cloneNode(true) : doc.createElementNS(W, 'w:rPr');
          setItalic(doc, pr, x.i);
          seg.r.appendChild(pr);
          seg.t = doc.createElementNS(W, 'w:t'); seg.t.setAttributeNS(XMLNS, 'xml:space', 'preserve'); seg.r.appendChild(seg.t);
          ins.appendChild(seg.r);
        }
        seg.t.textContent += x.c;
      }
      out.push(ins);
    }
    var k2 = 0;
    while (k2 < ops.length) {
      var op = ops[k2];
      if (op[0] === '=') {
        pending.forEach(function (a) { out.push(a); }); pending = [];
        emitEqual(oStart[op[1]], oStart[op[1]] + oT[op[1]].length); k2++; continue;
      }
      var dFrom = null, dTo = null, iFrom = null, iTo = null;
      while (k2 < ops.length && ops[k2][0] !== '=') {
        var o = ops[k2];
        if (o[0] === '-') { var s0 = oStart[o[1]], e0 = s0 + oT[o[1]].length; if (dFrom == null) dFrom = s0; dTo = e0; }
        else { var s1 = nStart[o[1]], e1 = s1 + nT[o[1]].length; if (iFrom == null) iFrom = s1; iTo = e1; }
        k2++;
      }
      if (dFrom != null) emitDel(dFrom, dTo);
      if (iFrom != null) emitIns(iFrom, iTo, dFrom != null ? dFrom : (dTo != null ? dTo : nearestOld(iFrom)));
    }
    function nearestOld(nIdx) { // template for a pure insertion: the old char just before it
      var o = 0; for (var z = 0; z < ops.length; z++) { if (ops[z][0] === '=' && nStart[ops[z][2]] < nIdx) o = oStart[ops[z][1]] + oT[ops[z][1]].length - 1; }
      return o;
    }
    pending.forEach(function (a) { out.push(a); });
    flushAtoms(chars.length + 1, out);
    // replace the paragraph's content, keeping pPr
    var pPr = kid(p, 'pPr');
    while (p.firstChild) p.removeChild(p.firstChild);
    if (pPr) p.appendChild(pPr);
    out.forEach(function (n) { delete n.__run; delete n.__eq; delete n.__t; p.appendChild(n); });
    return 'changed';
  }
  async function ensureComments(env, zip) {
    var P = new env.DOMParser();
    var cf = zip.file('word/comments.xml'), cdoc;
    if (cf) cdoc = P.parseFromString(await cf.async('string'), 'application/xml');
    else {
      cdoc = P.parseFromString('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="' + W + '"/>', 'application/xml');
      var relsF = zip.file('word/_rels/document.xml.rels');
      var rels = P.parseFromString(await relsF.async('string'), 'application/xml');
      var R = rels.documentElement, used = {};
      var rs = R.getElementsByTagName('Relationship');
      for (var i = 0; i < rs.length; i++) used[rs[i].getAttribute('Id')] = 1;
      var rid = 'rIdCL1', n = 1; while (used[rid]) rid = 'rIdCL' + (++n);
      var rel = rels.createElementNS(R.namespaceURI, 'Relationship');
      rel.setAttribute('Id', rid); rel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'); rel.setAttribute('Target', 'comments.xml');
      R.appendChild(rel);
      zip.file('word/_rels/document.xml.rels', ser(env, rels));
      var ctF = zip.file('[Content_Types].xml'), ct = P.parseFromString(await ctF.async('string'), 'application/xml');
      var ov = ct.createElementNS(ct.documentElement.namespaceURI, 'Override');
      ov.setAttribute('PartName', '/word/comments.xml'); ov.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml');
      ct.documentElement.appendChild(ov);
      zip.file('[Content_Types].xml', ser(env, ct));
    }
    return cdoc;
  }
  function ser(env, doc) {
    var s = new env.XMLSerializer().serializeToString(doc);
    return /^<\?xml/.test(s) ? s : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + s;
  }
  function addComment(doc, cdoc, p, id, text, date) {
    var c = cdoc.createElementNS(W, 'w:comment');
    c.setAttributeNS(W, 'w:id', String(id)); c.setAttributeNS(W, 'w:author', AUTHOR); c.setAttributeNS(W, 'w:initials', INITIALS); c.setAttributeNS(W, 'w:date', date);
    String(text).split(/\n+/).forEach(function (line) {
      var cp = cdoc.createElementNS(W, 'w:p'), r = cdoc.createElementNS(W, 'w:r'), t = cdoc.createElementNS(W, 'w:t');
      t.setAttributeNS(XMLNS, 'xml:space', 'preserve'); t.textContent = line; r.appendChild(t); cp.appendChild(r); c.appendChild(cp);
    });
    cdoc.documentElement.appendChild(c);
    var s = doc.createElementNS(W, 'w:commentRangeStart'); s.setAttributeNS(W, 'w:id', String(id));
    var e = doc.createElementNS(W, 'w:commentRangeEnd'); e.setAttributeNS(W, 'w:id', String(id));
    var r2 = doc.createElementNS(W, 'w:r');
    var ref = doc.createElementNS(W, 'w:commentReference'); ref.setAttributeNS(W, 'w:id', String(id)); r2.appendChild(ref);
    var pPr = kid(p, 'pPr');
    p.insertBefore(s, pPr ? pPr.nextSibling : p.firstChild);
    p.appendChild(e); p.appendChild(r2);
  }
  function maxId(doc) {
    var m = 0, all = doc.getElementsByTagNameNS(W, '*');
    for (var i = 0; i < all.length; i++) { var v = parseInt(all[i].getAttributeNS(W, 'id'), 10); if (!isNaN(v) && v > m) m = v; }
    return m;
  }
  async function buildTracked(env, buf, edits, date) {
    var zip = await env.JSZip.loadAsync(buf);
    var doc = new env.DOMParser().parseFromString(await zip.file('word/document.xml').async('string'), 'application/xml');
    var ps = docParas(doc), report = [], ids = { next: maxId(doc) + 1 }, cdoc = null;
    for (var i = 0; i < edits.length; i++) {
      var e = edits[i], n = parseInt(e.para, 10), p = ps[n - 1], why;
      if (!p) { report.push({ para: e.para, result: 'no such paragraph' }); continue; }
      if (e.text != null) {
        if (hasRevisions(p)) why = 'left alone: this paragraph already has tracked changes; accept or reject them first';
        else why = trackedPara(doc, p, String(e.text).slice(0, MAX_EDIT_CHARS), ids, date);
      } else why = 'comment only';
      if (e.comment && !/^left alone/.test(why)) {
        if (!cdoc) { cdoc = await ensureComments(env, zip); ids.next = Math.max(ids.next, maxId(cdoc) + 1); }
        addComment(doc, cdoc, p, ids.next++, String(e.comment).slice(0, 1500), date);
      }
      report.push({ para: n, result: why });
    }
    zip.file('word/document.xml', ser(env, doc));
    if (cdoc) zip.file('word/comments.xml', ser(env, cdoc));
    var blob = await zip.generateAsync({ type: env.blobType || 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', compression: 'DEFLATE' });
    return { blob: blob, report: report };
  }

  /* ---------- discussions ---------- */
  function scoreText(text, terms, phrase) {
    var t = norm(text), s = 0;
    if (phrase && t.indexOf(phrase) >= 0) s += 5;
    terms.forEach(function (w) { if (t.indexOf(w) >= 0) s += w.length > 3 ? 2 : 1; });
    return s;
  }
  function qTerms(q) {
    var stop = { the: 1, and: 1, of: 1, to: 1, a: 1, an: 1, is: 1, in: 1, for: 1, on: 1, we: 1, do: 1, how: 1, what: 1, our: 1 };
    return norm(q).split(' ').filter(function (w) { return w && !stop[w] && (w.length > 1 || /[א-ת]/.test(w)); });
  }

  /* ---------- the factory ---------- */
  function create(env) {
    var S = { att: [], files: {}, made: [] };
    function libPath(p) { return String(p || '').replace(/^\/+/, '').replace(/^Shared Documents\//i, ''); }
    async function libItem(path) {
      var d = await env.driveId();
      return env.g('/drives/' + d + '/root:/' + libPath(path).split('/').map(encodeURIComponent).join('/') + '?$select=id,name,size,lastModifiedDateTime,file,eTag,@microsoft.graph.downloadUrl');
    }
    async function libBuf(item) { return env.download(item); }
    async function openSource(src) {
      var m = String(src || '').match(/^\s*attachment\s*#?\s*(\d+)\s*$/i);
      if (m) {
        var a = S.att[parseInt(m[1], 10) - 1];
        if (!a) throw new Error('There is no attachment ' + m[1] + ' in this conversation.');
        return a;
      }
      var key = libPath(src);
      var item = await libItem(key);
      var hit = S.files[key];
      if (hit && hit.eTag === item.eTag) return hit;
      if (item.size > MAX_ATTACH_BYTES) throw new Error('That file is too large to open here (' + Math.round(item.size / 1048576) + ' MB).');
      var name = item.name, buf = await libBuf(item), out;
      if (/\.docx$/i.test(name)) { out = await loadDocx(env, buf); out.lines = docxLines(out); }
      else if (/\.(md|txt|csv|json)$/i.test(name)) { var t = env.decode(buf); out = { kind: 'text', lines: t.split(/\n/) }; }
      else if (/\.xlsx$/i.test(name)) { var rows = await xlsxRows(env, buf); out = { kind: 'xlsx', lines: rows.map(function (r) { return 'row ' + r.n + ': ' + r.cells.map(function (c) { return c == null ? '' : String(c).replace(/\s+/g, ' '); }).join(' | '); }) }; }
      else throw new Error('I can read Word, Excel and text files from the library, not ' + (name.split('.').pop() || 'this kind of file') + ' files.');
      out.name = name; out.path = key; out.eTag = item.eTag; out.buf = buf; out.modified = item.lastModifiedDateTime;
      S.files[key] = out;
      return out;
    }
    function readOut(f, input, label) {
      var total = f.lines.length;
      if (input.paragraphs && f.kind === 'docx') {
        var rg = rangeOf(input.paragraphs, f.paras.length);
        if (!rg) return { error: 'paragraphs should look like "12" or "12-20".' };
        var pick = f.lines.filter(function (l) { var n = parseInt(l.slice(1), 10); return n >= rg[0] && n <= rg[1]; });
        return { file: label, paragraphs: rg[0] + '-' + rg[1], content: cut(pick.join('\n'), PART_CHARS) };
      }
      var parts = partsOf(f.lines, PART_CHARS), p = num(input.part, 1, parts.length, 1);
      var out = { file: label, part: p, parts: parts.length, content: parts[p - 1] };
      if (p === 1) {
        if (f.kind === 'docx') {
          out.paragraphs = f.paras.length;
          out.tracked_paragraphs = f.paras.filter(function (x) { return x.rev; }).length;
          out.how_to_read = '¶N is paragraph N (the number make_word_file uses). _x_ is italic. [tracked] means the paragraph already carries tracked changes, shown here as if accepted.';
          if (f.comments && f.comments.length) out.comments = f.comments.slice(0, 40);
        }
        if (f.modified) out.last_saved = f.modified;
        if (f.kind === 'xlsx') out.rows = total;
      }
      if (parts.length > p) out.note = 'Part ' + p + ' of ' + parts.length + '. Ask for part ' + (p + 1) + ', or for a paragraph range.';
      return out;
    }

    async function find_files(input) {
      var q = String(input.query || '').trim();
      if (!q) return { error: 'empty query' };
      var d = await env.driveId();
      var res = await env.g('/drives/' + d + "/root/search(q='" + encodeURIComponent(q.replace(/'/g, "''")) + "')?$select=name,size,lastModifiedDateTime,parentReference,file&$top=60");
      var items = (res && res.value || []).filter(function (x) { return x.file && /\.(docx|xlsx|md|txt)$/i.test(x.name); });
      if (input.folder) { var f = norm(input.folder); items = items.filter(function (x) { return norm((x.parentReference && x.parentReference.path) || '').indexOf(f) >= 0; }); }
      items = items.filter(function (x) { return !/\/(Archive|_Archive|\.graphics-backups|skill-backups)\b/i.test((x.parentReference && x.parentReference.path) || ''); });
      items.sort(function (a, b) { return String(b.lastModifiedDateTime).localeCompare(String(a.lastModifiedDateTime)); });
      return {
        found: items.length,
        files: items.slice(0, num(input.limit, 1, 20, 12)).map(function (x) {
          var par = ((x.parentReference && x.parentReference.path) || '').replace(/^\/drives\/[^/]+\/root:\/?/, '');
          return { path: decodeURIComponent(par ? par + '/' + x.name : x.name), saved: String(x.lastModifiedDateTime || '').slice(0, 16).replace('T', ' '), size_kb: Math.round((x.size || 0) / 1024) };
        }),
        note: 'Newest first. Open one with read_file and its path.'
      };
    }
    async function read_file(input) {
      if (!input.path) return { error: 'Give the file\'s path (find_files gives it).' };
      var f = await openSource(input.path);
      return readOut(f, input, f.path || f.name);
    }
    async function read_attachment(input) {
      var a = S.att[num(input.n, 1, 99, 1) - 1];
      if (!a) return { error: 'No attachment ' + (input.n || 1) + ' in this conversation.', attachments: S.att.map(function (x, i) { return (i + 1) + '. ' + x.name; }) };
      if (!a.lines) return { error: a.name + ' is a picture; it was shown with the question.' };
      return readOut(a, input, 'attachment ' + (S.att.indexOf(a) + 1) + ': ' + a.name);
    }
    async function make_word_file(input) {
      var edits = Array.isArray(input.edits) ? input.edits.slice(0, MAX_EDITS) : [];
      if (!input.source) return { error: 'Say which file to mark up: a library path, or "attachment 1".' };
      if (!edits.length) return { error: 'No edits given.' };
      var f = await openSource(input.source);
      if (f.kind !== 'docx') return { error: 'Only a Word file can be marked up.' };
      var built = await buildTracked(env, f.buf, edits, isoNow());
      var base = (input.name || f.name || 'edited.docx').replace(/\.docx$/i, '');
      if (!input.name && !/_CL$/.test(base)) base += '_CL';
      var name = base.replace(/[\\/:*?"<>|]+/g, '-') + '.docx';
      S.made.push({ name: name, blob: built.blob });
      var changed = built.report.filter(function (r) { return r.result === 'changed' || r.result === 'comment only'; }).length;
      return {
        file: name, ready_to_download: changed > 0, paragraphs_changed: built.report.filter(function (r) { return r.result === 'changed'; }).length,
        results: built.report, dropped_over_limit: Array.isArray(input.edits) && input.edits.length > MAX_EDITS ? input.edits.length - MAX_EDITS : undefined,
        note: 'The editor gets a download button for this file under your answer. Nothing was saved to SharePoint. Tell them which paragraphs were left alone and why.'
      };
    }
    async function search_discussions(input) {
      var q = String(input.query || '').trim();
      if (!q) return { error: 'empty query' };
      var terms = qTerms(q), phrase = norm(q), want = String(input.source || 'all').toLowerCase(), lim = num(input.limit, 1, 15, 8), out = { query: q };
      if (want === 'all' || /chat/.test(want)) {
        try {
          var log = await env.json('chat.json');
          var rows = log.rows || [], hits = [];
          for (var i = 0; i < rows.length; i++) { var sc = scoreText(rows[i][3], terms, phrase); if (sc >= Math.min(2, terms.length * 2)) hits.push([sc, i]); }
          hits.sort(function (a, b) { return b[0] - a[0] || b[1] - a[1]; });
          var seen = {}, res = [];
          hits.forEach(function (h) {
            if (res.length >= lim) return;
            var r = rows[h[1]], th = r[2];
            if (seen[th + h[1]]) return; seen[th + h[1]] = 1;
            var ctx = [];
            for (var k = Math.max(0, h[1] - 3); k <= Math.min(rows.length - 1, h[1] + 3); k++) if (rows[k][2] === th && k !== h[1]) ctx.push(rows[k][1] + ': ' + cut(rows[k][3], 220));
            res.push({ date: r[0], who: r[1], text: cut(r[3], 500), same_thread: ctx.length ? ctx : undefined });
          });
          out.team_chat = { covers: log.covers, matches: hits.length, messages: res };
        } catch (e) { out.team_chat = { error: 'The team chat log could not be opened.' }; }
      }
      async function sheet(path, label) {
        try {
          var f = await openSource(path), rows = f.lines.map(function (l, i) { return { l: l, i: i }; });
          var head = f.lines[0] || '';
          var hits = rows.slice(1).map(function (r) { return [scoreText(r.l, terms, phrase), r]; }).filter(function (x) { return x[0] >= Math.min(2, terms.length * 2); });
          hits.sort(function (a, b) { return b[0] - a[0]; });
          return { sheet: label, columns: cut(head.replace(/^row \d+: /, ''), 400), matches: hits.length, rows: hits.slice(0, lim).map(function (x) { return cut(x[1].l, 700); }) };
        } catch (e) { return { sheet: label, error: 'could not be opened' }; }
      }
      if (want === 'all' || /proof/.test(want)) out.proofreading_feedback = await sheet('Claude/Proofreading/Feedback.xlsx', 'Proofreading feedback (Claude/Proofreading/Feedback.xlsx)');
      if (want === 'all' || /graph/.test(want)) {
        var folders = (env.masechtaFolders && await env.masechtaFolders()) || [];
        out.graphics_feedback = [];
        for (var k = 0; k < folders.length; k++) {
          var r = await sheet(folders[k] + '/Graphics/feedback.xlsx', folders[k] + ' graphics feedback');
          if (!r.error && r.matches) out.graphics_feedback.push(r);
        }
      }
      out.reminder = 'Discussion is evidence of how the team got to a decision, not a ruling: cite the ruling where one exists. Never summarise or judge a person from these messages.';
      return out;
    }

    /* attachments: returns the content blocks for the question */
    async function attach(file) {
      if (file.size > MAX_ATTACH_BYTES) throw new Error(file.name + ' is too large (over 15 MB).');
      var name = file.name || 'attachment', ext = (name.split('.').pop() || '').toLowerCase(), a = { name: name };
      if (/^image\//.test(file.type) || /^(png|jpe?g|gif|webp|heic)$/.test(ext)) {
        if (S.att.filter(function (x) { return x.img; }).length >= MAX_IMAGES * 4) throw new Error('That is enough pictures for one conversation; start a new one.');
        a.img = await env.shrinkImage(file, IMAGE_PX);
        a.kind = 'image';
      } else {
        var buf = await env.fileBuf(file);
        if (ext === 'docx') { var d = await loadDocx(env, buf); a.kind = 'docx'; a.zip = d.zip; a.paras = d.paras; a.comments = d.comments; a.lines = docxLines(d); a.buf = buf; }
        else if (ext === 'pdf') {
          var pt = env.pdfText ? await env.pdfText(buf) : null;
          if (!pt || pt.replace(/\s/g, '').length < 40) throw new Error(name + ' has no text I can read (it looks like a scan). Attach a screenshot of the part you mean instead.');
          a.kind = 'pdf'; a.lines = pt.split(/\n/);
        }
        else if (/^(txt|md|csv)$/.test(ext)) { a.kind = 'text'; a.lines = env.decode(buf).split(/\n/); }
        else if (ext === 'xlsx') { var rows = await xlsxRows(env, buf); a.kind = 'xlsx'; a.lines = rows.map(function (r) { return 'row ' + r.n + ': ' + r.cells.map(function (c) { return c == null ? '' : c; }).join(' | '); }); }
        else throw new Error('I can take Word, PDF, Excel, text files and pictures, not .' + ext + ' files.');
      }
      S.att.push(a);
      return a;
    }
    /* the blocks one question carries for its new attachments: a short preview each, never the whole file */
    function attachBlocks(list) {
      var blocks = [], imgs = 0;
      list.forEach(function (a) {
        var n = S.att.indexOf(a) + 1;
        if (a.img) {
          if (imgs++ >= MAX_IMAGES) { blocks.push({ type: 'text', text: '[Attachment ' + n + ': ' + a.name + ' was not sent: at most ' + MAX_IMAGES + ' pictures per question.]' }); return; }
          blocks.push({ type: 'text', text: '[Attachment ' + n + ': ' + a.name + ' (picture)]' });
          blocks.push({ type: 'image', source: { type: 'base64', media_type: a.img.type, data: a.img.data } });
          return;
        }
        var all = a.lines.join('\n'), info = a.kind === 'docx' ? a.paras.length + ' paragraphs' + (a.comments && a.comments.length ? ', ' + a.comments.length + ' comments' : '') : a.lines.length + ' lines';
        blocks.push({ type: 'text', text: '[Attachment ' + n + ': ' + a.name + ' (' + a.kind + ', ' + info + ', ' + all.length.toLocaleString('en-US') + ' characters). ' +
          (all.length > PREVIEW_CHARS ? 'Beginning:]\n' + all.slice(0, PREVIEW_CHARS) + '\n[…read the rest with read_attachment n=' + n + (a.kind === 'docx' ? '. To return it marked up, make_word_file with source "attachment ' + n + '".]' : '.]') : 'Whole text:]\n' + all + (a.kind === 'docx' ? '\n[To return it marked up, make_word_file with source "attachment ' + n + '".]' : '')) });
      });
      return blocks;
    }
    /* after an answer: pictures leave the history (a stub stays), so follow-ups don't pay for them again */
    function slimBlocks(content) {
      if (!Array.isArray(content)) return content;
      return content.map(function (b) { return b.type === 'image' ? { type: 'text', text: '[picture omitted after the answer; ask the editor to attach it again if you need to look at it again]' } : b; });
    }
    /* inside one question: keep the tool results under TURN_RESULT_CHARS, stubbing the oldest first */
    function trimTurn(turn, from) {
      var res = [];
      for (var i = from || 0; i < turn.length; i++) {
        var m = turn[i];
        if (m.role === 'user' && Array.isArray(m.content)) m.content.forEach(function (b) { if (b.type === 'tool_result' && typeof b.content === 'string') res.push(b); });
      }
      var total = res.reduce(function (s, b) { return s + b.content.length; }, 0), k = 0;
      while (total > TURN_RESULT_CHARS && k < res.length - 2) {
        var b = res[k++];
        if (b.content.length < 200) continue;
        total -= b.content.length;
        b.content = '(this earlier result was dropped to keep the question affordable; search or read again if you still need it)';
        total += b.content.length;
      }
      return turn;
    }
    function takeMade() { var m = S.made; S.made = []; return m; }
    function reset() { S.att = []; S.files = {}; S.made = []; }

    return {
      tools: { find_files: find_files, read_file: read_file, read_attachment: read_attachment, make_word_file: make_word_file, search_discussions: search_discussions },
      labels: { find_files: 'Looking for the file', read_file: 'Reading the file', read_attachment: 'Reading your attachment', make_word_file: 'Marking up the Word file', search_discussions: 'Searching the team chat and the feedback sheets' },
      attach: attach, attachBlocks: attachBlocks, slimBlocks: slimBlocks, trimTurn: trimTurn, takeMade: takeMade, reset: reset,
      attachments: function () { return S.att.slice(); }
    };
  }

  return { create: create, _test: { loadDocx: loadDocx, docxLines: docxLines, buildTracked: buildTracked, parseMarked: parseMarked, xlsxRows: xlsxRows, tokens: tokens }, LIMITS: { PART_CHARS: PART_CHARS, PREVIEW_CHARS: PREVIEW_CHARS, MAX_IMAGES: MAX_IMAGES, TURN_RESULT_CHARS: TURN_RESULT_CHARS, MAX_EDITS: MAX_EDITS } };
})();
if (typeof module !== 'undefined') module.exports = ASKY;
