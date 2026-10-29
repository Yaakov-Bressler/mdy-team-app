/* ---------- MDYWiden: widened search (variant forms) ----------
   One engine for the site's Guide (MDY Lookup bundle) and the Team App, so both widen a search the same way.
   Asked for 29 Oct 2026: searches should also find variant verb forms, plurals and spellings.

   English (and transliteration):
     verb forms     say → says, said, saying; bring → brought; permit → permitted; teach → taught
     plurals        animal → animals, animal's; korban → korbanos; mitzvah → mitzvos; tanna → tannaim; kohen → kohanim
     spellings      only for transliterated words and names (set in italics, or always capitalised, in the delivered
                    volumes): apostrophes and hyphens, doubled letters, a closing h after a vowel, ts/tz.
                    So shemitah finds shemittah, kappara finds kapparah, bedieved finds bedi'eved, Yanai finds Yannai,
                    but "to" never finds "too" and "al" never finds "all".
   Hebrew / Aramaic:
     prefixes       ו ב ד ה כ ל מ ש (up to three) before the word
     malei/chaser   with or without ו and י
     endings        ־ה ↔ ־ות (מצוה, מצות); ־ים ↔ ־ין ↔ ־יא ↔ ־א ↔ none (כהן, כהנים; עבד, עבדא)

   Pure functions, no page code. ES5 so the SharePoint bundle and the app can both carry it verbatim.
   A form is only offered as a variant: callers run their exact search first and list what this adds after it. */
var MDYWiden = (function () {
  'use strict';

  /* ---------- English ---------- */

  function low(s) {
    return String(s || '').normalize('NFKD').replace(/[̀-ͯ]/g, '')
      .replace(/[‘’ʼ`´]/g, "'").replace(/[‐-―]/g, '-').toLowerCase();
  }
  var WORD = /[a-z0-9]+(?:['\-][a-z0-9]+)*/g;
  /* words of a text with their place in it: [{w, s, e}] (w lower case, apostrophes straightened) */
  function engTokens(text) {
    var t = low(text), out = [], m; WORD.lastIndex = 0;
    while ((m = WORD.exec(t))) out.push({ w: m[0], s: m.index, e: m.index + m[0].length });
    return out;
  }
  /* the transliteration spelling key: no apostrophes or hyphens, no doubled letters, no closing h after a vowel, ts = tz */
  function fold(w) {
    return w.replace(/['\-]/g, '').replace(/(.)\1+/g, '$1').replace(/([aeiou])h$/, '$1').replace(/ts/g, 'tz');
  }

  /* irregular forms: every word in a group finds every other */
  var IRR_GROUPS = [
    'be is am are was were been being', 'have has had having', 'do does did done doing', 'say says said saying',
    'go goes went gone going', 'come comes came coming', 'become becomes became becoming',
    'bring brings brought bringing', 'take takes took taken taking', 'give gives gave given giving',
    'see sees saw seen seeing', 'eat eats ate eaten eating', 'teach teaches taught teaching',
    'buy buys bought buying', 'sell sells sold selling', 'tell tells told telling', 'think thinks thought thinking',
    'find finds found finding', 'hold holds held holding', 'know knows knew known knowing',
    'leave leaves left leaving', 'keep keeps kept keeping', 'make makes made making', 'mean means meant meaning',
    'send sends sent sending', 'spend spends spent spending', 'lose loses lost losing', 'pay pays paid paying',
    'lay lays laid laying', 'lie lies lay lain lying', 'feed feeds fed feeding', 'lead leads led leading',
    'hear hears heard hearing', 'seek seeks sought seeking', 'catch catches caught catching',
    'steal steals stole stolen stealing', 'swear swears swore sworn swearing', 'wear wears wore worn wearing',
    'tear tears tore torn tearing', 'grow grows grew grown growing', 'throw throws threw thrown throwing',
    'draw draws drew drawn drawing', 'slay slays slew slain slaying', 'bear bears bore borne born bearing',
    'strike strikes struck stricken striking', 'hang hangs hung hanged hanging', 'bind binds bound binding',
    'shake shakes shook shaken shaking', 'write writes wrote written writing', 'ride rides rode ridden riding',
    'hide hides hid hidden hiding', 'bite bites bit bitten biting', 'feel feels felt feeling', 'deal deals dealt dealing',
    'flee flees fled fleeing', 'fly flies flew flown flying', 'forget forgets forgot forgotten forgetting',
    'forgive forgives forgave forgiven forgiving', 'understand understands understood understanding',
    'fight fights fought fighting', 'run runs ran running', 'begin begins began begun beginning',
    'choose chooses chose chosen choosing', 'break breaks broke broken breaking', 'speak speaks spoke spoken speaking',
    'forbid forbids forbade forbidden forbidding', 'get gets got gotten getting', 'stand stands stood standing',
    'sit sits sat sitting', 'rise rises rose risen rising', 'arise arises arose arisen arising', 'fall falls fell fallen falling',
    'drink drinks drank drunk drinking', 'set sets setting', 'put puts putting', 'cut cuts cutting', 'let lets letting',
    'burn burns burned burnt burning', 'dwell dwells dwelt dwelled dwelling', 'sleep sleeps slept sleeping',
    'weep weeps wept weeping', 'sweep sweeps swept sweeping', 'build builds built building', 'lend lends lent lending',
    'shoot shoots shot shooting', 'win wins won winning', 'wake wakes woke woken waking', 'freeze freezes froze frozen freezing',
    'blow blows blew blown blowing', 'sink sinks sank sunk sinking', 'spring springs sprang sprung springing',
    'breed breeds bred breeding', 'grind grinds ground grinding', 'stick sticks stuck sticking', 'cling clings clung clinging',
    'beat beats beaten beating', 'hit hits hitting', 'shut shuts shutting', 'spread spreads spreading',
    'cast casts casting', 'cost costs costing', 'hurt hurts hurting', 'quit quits quitting', 'read reads reading',
    /* nouns */
    'man men man\'s men\'s', 'woman women woman\'s women\'s', 'child children child\'s children\'s', 'ox oxen ox\'s oxen\'s',
    'foot feet', 'tooth teeth', 'mouse mice', 'goose geese', 'calf calves calf\'s', 'half halves', 'knife knives',
    'life lives', 'wife wives wife\'s', 'loaf loaves', 'leaf leaves', 'thief thieves', 'shelf shelves', 'wolf wolves',
    'hoof hooves hoofs', 'person people persons person\'s people\'s',
    /* transliterated nouns whose plural is not regular */
    'kohen kohanim kohen\'s kohanim\'s', 'levi leviim levi\'im levi\'s', 'kli keilim kelim', 'minchah menachos minchas minchos',
    'ben bnei banim', 'bas banos bnos', 'ir arim', 'yom yamim', 'shanah shanim', 'ish anashim', 'ishah nashim',
    'av avos', 'em imahos', 'beis batei', 'eid eidim', 'eidah eidos'
  ];
  var IRR = {};
  IRR_GROUPS.forEach(function (g, i) {
    g.split(' ').forEach(function (w) { (IRR[w] = IRR[w] || []).push(i); });
  });
  function irrForms(w) {
    var out = [];
    (IRR[w] || []).forEach(function (i) { IRR_GROUPS[i].split(' ').forEach(function (f) { out.push(f); }); });
    return out;
  }

  var V = /[aeiouy]/;
  function cvc(b) { return /[^aeiou][aeiou][^aeiouwxy]$/.test(b); }
  /* bases a word may be an inflection of (the word itself first) */
  function bases(w) {
    var out = [w], add = function (b) { if (b && b.length >= 3 && V.test(b) && out.indexOf(b) < 0) out.push(b); };
    var x = w.replace(/'s$/, '').replace(/s'$/, 's');
    if (x !== w) { out.push(x); w = x; }
    var n = w.length;
    if (/ies$/.test(w) && n > 4) add(w.slice(0, -3) + 'y');
    if (/ied$/.test(w) && n > 4) add(w.slice(0, -3) + 'y');
    if (/(ss|sh|ch|x|z|o)es$/.test(w)) add(w.slice(0, -2));
    if (/[^su'i]s$/.test(w) && n > 3) add(w.slice(0, -1));
    if (/ed$/.test(w) && n >= 4) {
      add(w.slice(0, -2)); add(w.slice(0, -1));
      if (/(.)\1ed$/.test(w)) add(w.slice(0, -3));
    }
    if (/ing$/.test(w) && n > 5) {
      add(w.slice(0, -3)); add(w.slice(0, -3) + 'e');
      if (/(.)\1ing$/.test(w)) add(w.slice(0, -4));
    }
    /* transliterated plurals: -os/-oth/-ot (korbanos, mitzvos, chata'os), -im (tannaim) */
    var h = w.replace(/'(?=(os|oth|ot|im)$)/, '');
    var m = /^(.+?)(os|oth|ot)$/.exec(h);
    if (m && m[1].length >= 2) { add(m[1]); add(m[1] + 'ah'); add(m[1] + 'a'); add(m[1] + 's'); }
    m = /^(.+?)im$/.exec(h);
    if (m && m[1].length >= 3) { add(m[1]); add(m[1] + 'a'); }
    return out;
  }
  /* the inflected forms of a base */
  function inflect(b, push) {
    push(b); push(b + "'s"); push(b + 's');
    if (/(s|sh|ch|x|z|o)$/.test(b)) push(b + 'es');
    if (/[^aeiou]y$/.test(b)) { var st = b.slice(0, -1); push(st + 'ies'); push(st + 'ied'); }
    if (/e$/.test(b)) { push(b + 'd'); push(b.slice(0, -1) + 'ing'); }
    else { push(b + 'ed'); push(b + 'ing'); }
    if (cvc(b)) { var l = b[b.length - 1]; push(b + l + 'ed'); push(b + l + 'ing'); }
    /* transliterated plurals */
    if (/ah$/.test(b)) { push(b.slice(0, -2) + 'os'); push(b.slice(0, -2) + 'as'); }
    else if (/a$/.test(b)) { push(b.slice(0, -1) + 'os'); push(b + 'im'); push(b + "'im"); }
    else if (/[^aeiou]$/.test(b)) { push(b + 'os'); push(b + 'im'); }
    if (/[aeiou]s$/.test(b) && b.length > 3) { push(b.slice(0, -1) + 'os'); push(b.slice(0, -1) + "'os"); }
  }
  /* every form a query word may take: { w: the word, forms: {form: 1}, folds: {fold: 1} } */
  function engWord(w) {
    w = low(w);
    var forms = {}, push = function (f) { forms[f] = 1; };
    var irr = irrForms(w);
    if (irr.length) irr.forEach(push);
    else bases(w).forEach(function (b) { irrForms(b).forEach(push); inflect(b, push); });
    push(w);
    var folds = {};
    Object.keys(forms).forEach(function (f) { if (f.replace(/['\-]/g, '').length >= 3) folds[fold(f)] = 1; });
    return { w: w, forms: forms, folds: folds };
  }

  /* which words in a body of text are transliterations or names: italic in at least 30% of their uses, or capitalised in 90%.
     Built once per list of texts; items are [plainText, richSegments|null] */
  function vocab(items) {
    var st = {};
    items.forEach(function (it) {
      var segs = it[1] && it[1].length ? it[1] : [[it[0], false]];
      segs.forEach(function (sg) {
        var t = String(sg[0] || '').normalize('NFKD').replace(/[̀-ͯ]/g, '').replace(/[‘’ʼ`´]/g, "'"), m, re = /[A-Za-z0-9]+(?:['\-‐-―][A-Za-z0-9]+)*/g;
        while ((m = re.exec(t))) {
          var w = low(m[0]), s = st[w] || (st[w] = [0, 0, 0]);
          s[0]++; if (sg[1]) s[1]++; if (/^[A-Z]/.test(m[0])) s[2]++;
        }
      });
    });
    var ok = {};
    Object.keys(st).forEach(function (w) { var s = st[w]; if (w.length >= 3 && (s[1] >= 0.3 * s[0] || s[2] >= 0.9 * s[0])) ok[w] = 1; });
    return { translit: ok, all: true };
  }
  var ALL_TRANSLIT = { all: true, any: true };

  /* a query: its words, each with its forms */
  function engQuery(q) {
    var toks = engTokens(q);
    return toks.length ? toks.map(function (t) { return engWord(t.w); }) : null;
  }
  /* 2 = the word itself, 1 = one of its forms, 0 = no */
  function engTok(tw, qw, vb) {
    if (tw === qw.w) return 2;
    if (qw.forms[tw]) return 1;
    if ((vb === ALL_TRANSLIT || (vb && vb.translit[tw])) && qw.folds[fold(tw)]) return 1;
    return 0;
  }
  /* the query's words, in order, among a text's words: { at, len } in the lower-case text, exact: true when no word needed a variant */
  function engFind(text, Q, vb, toks) {
    toks = toks || engTokens(text);
    var n = Q.length, best = null;
    outer: for (var i = 0; i + n <= toks.length; i++) {
      var exact = true, vw = [];
      for (var k = 0; k < n; k++) {
        var s = engTok(toks[i + k].w, Q[k], vb);
        if (!s) continue outer;
        if (s === 1) { exact = false; vw.push(toks[i + k].w); }
      }
      var hit = { at: toks[i].s, len: toks[i + n - 1].e - toks[i].s, exact: exact, vw: vw };
      if (!exact) return hit;
      best = best || hit;
    }
    return best;
  }

  /* ---------- Hebrew and Aramaic ---------- */

  var FIN = { 'ך': 'כ', 'ם': 'מ', 'ן': 'נ', 'ף': 'פ', 'ץ': 'צ' };
  function hebWord(w) { return String(w || '').replace(/[^א-ת]/g, '').replace(/[ךםןףץ]/g, function (c) { return FIN[c]; }); }
  function hebTokens(s) {
    return String(s || '').replace(/[֑-ֽֿ-ׇ]/g, '').split(/[\s־\/]+/).map(hebWord).filter(Boolean);
  }
  /* with or without ו and י inside the word; the first and last letters stay, so רב is not רבי and קרבנו is not קרבן */
  function loose(w) { return w.length > 2 ? w[0] + w.slice(1, -1).replace(/[וי]/g, '') + w[w.length - 1] : w; }
  var PFX = /^[ובדהכלמש]{1,3}$/;
  /* before a variant, מ is not read as a prefix: מזבח is not a form of זבח */
  var PFX_V = /^[ובדהכלש]{1,3}$/;
  /* the forms of one Hebrew/Aramaic word (normalised: no nikkud, final letters as medial), as loose spellings */
  function hebWordForms(w) {
    var out = {}, add = function (f) { if (f.length >= 2) out[loose(f)] = 1; };
    add(w);
    var m, X;
    if ((m = /^(.{2,})ה$/.exec(w))) { X = m[1]; add(X + 'ות'); add(X + 'ת'); }                 /* ־ה → ־ות */
    else if ((m = /^(.{2,})ות$/.exec(w))) { X = m[1]; add(X + 'ה'); add(X + 'תא'); }        /* ־ות → ־ה, ־תא */
    else if ((m = /^(.{2,})י[מנא]$/.exec(w))) { X = m[1]; plural(X); }                       /* ־ים ־ין ־יא */
    else if ((m = /^(.{3,})א$/.exec(w))) { X = m[1]; plural(X); }                                           /* ־א */
    else if (w.length >= 3) plural(w);
    function plural(X) {
      if (X.length >= 3) add(X);
      if (/[וי]$/.test(X)) return;      /* רבי: not רביא, which reads as רבא */
      add(X + 'ימ'); add(X + 'ינ'); add(X + 'יא');
      if (X.length >= 3) add(X + 'א');
    }
    return out;
  }
  function hebQuery(q) {
    var toks = hebTokens(q);
    return toks.length ? toks.map(function (w) { return { w: w, l: loose(w), forms: hebWordForms(w) }; }) : null;
  }
  /* 2 = the word itself (with or without prefix letters), 1 = a variant, 0 = no */
  function hebTok(t, qw) {
    if (t === qw.w) return 2;
    for (var k = 0; k <= 3 && t.length - k >= 2; k++) {
      var p = t.slice(0, k);
      if (k && !PFX.test(p)) break;
      var r = t.slice(k);
      if (k && r === qw.w) return 2;
      if ((!k || PFX_V.test(p)) && qw.forms[loose(r)]) return 1;
    }
    return 0;
  }
  /* the query's words, in order, in a token list: { at, len, exact } or null */
  function hebFind(toks, Q) {
    var n = Q.length, best = null;
    outer: for (var i = 0; i + n <= toks.length; i++) {
      var exact = true, vw = [], vi = [];
      for (var k = 0; k < n; k++) {
        var s = hebTok(toks[i + k], Q[k]);
        if (!s) continue outer;
        if (s === 1) { exact = false; vw.push(toks[i + k]); vi.push(i + k); }
      }
      if (!exact) return { at: i, len: n, exact: false, vw: vw, vi: vi };
      best = best || { at: i, len: n, exact: true, vw: [], vi: [] };
    }
    return best;
  }

  function hasHeb(s) { return /[א-ת]/.test(s || ''); }
  return {
    VERSION: '1.0.0',
    low: low, fold: fold, engTokens: engTokens, engWord: engWord, engQuery: engQuery, engTok: engTok, engFind: engFind, vocab: vocab,
    ALL_TRANSLIT: ALL_TRANSLIT, hebTokens: hebTokens, hebWord: hebWord, hebWordForms: hebWordForms, hebQuery: hebQuery,
    hebTok: hebTok, hebFind: hebFind, hasHeb: hasHeb, bases: bases
  };
})();
if (typeof module !== 'undefined' && module.exports) module.exports = MDYWiden;
