/* Quire — Find, Replace and Go To: Word's special characters (^p, ^t, ^?, ^# …), wildcards
 * (?, *, [ ], {n,m}, @, < >, groups and \1 back-references), match case / whole words / sounds like /
 * all word forms, formatting criteria, Highlight All, and the Go To tab.
 */
(function () {
  'use strict';
  const L = window.L, D = L.D, O = L.O, E = L.ed, LY = L.layout;
  const { h } = L;
  const ui = L.ui;
  const G = (L.dlg = L.dlg || {});
  const FD = (L.find = {});
  const doc = () => D.doc;
  const A = () => L.app;
  /* what the Find and Replace dialog shows: it stays open and keeps the keyboard (Enter is Find Next again) */
  const KEEP = { focus: false };

  /* ---------- story text with paragraph marks ---------- */
  function storyText(story) {
    D.reindex(doc());
    const paras = story.paras || [];
    const starts = [];
    let s = '';
    for (const p of paras) { starts.push(s.length); s += D.ptext(p) + '\r'; }
    return { text: s, paras, starts };
  }
  function toPos(st, off) {
    let lo = 0, hi = st.starts.length - 1;
    while (lo < hi) { const mid = (lo + hi + 1) >> 1; if (st.starts[mid] <= off) lo = mid; else hi = mid - 1; }
    const p = st.paras[lo];
    return D.pos(p, Math.min(off - st.starts[lo], D.plen(p)));
  }
  function fromPos(st, pos) { const i = st.paras.indexOf(pos.p); return i < 0 ? 0 : st.starts[i] + pos.o; }

  /* ---------- pattern compilation ---------- */
  const esc = (s) => s.replace(/[.*+?^${}()|[\]\\\/]/g, '\\$&');
  const SPECIAL = { p: '\\r', t: '\\t', l: '\\n', m: '\\f', n: '\\v', '?': '[^\\r]', '#': '\\d', $: '\\p{L}', w: '[ \\t\\u00a0]+', s: '\\u00a0', '~': '\\u2011', '-': '\\u00ad', '+': '\\u2014', '=': '\\u2013', '^': '\\^', g: '\\uFFFC', f: '\\uFFFC', e: '\\uFFFC', a: '(?:)', d: '(?:)', b: '\\r', '%': '§', v: '¶', '&': '' };
  function specials(src, wildcards) {
    let out = '';
    for (let i = 0; i < src.length; i++) {
      const c = src[i];
      if (c === '^' && i + 1 < src.length) {
        const n = src[i + 1];
        if (n === 'u' && /^\d+/.test(src.slice(i + 2))) { const m = /^\d+/.exec(src.slice(i + 2))[0]; out += esc(String.fromCharCode(+m)); i += 1 + m.length; continue; }
        if (n === '0' && /^0\d+/.test(src.slice(i + 1))) { const m = /^\d+/.exec(src.slice(i + 1))[0]; out += esc(String.fromCharCode(+m)); i += m.length; continue; }
        const lower = n.toLowerCase();
        if (SPECIAL[lower] != null || SPECIAL[n] != null) { out += SPECIAL[n] != null ? SPECIAL[n] : SPECIAL[lower]; i++; continue; }
      }
      out += wildcards ? c : esc(c);
    }
    return out;
  }
  /** translate Word wildcard syntax to a JS regexp source */
  function wildcard(src) {
    let out = '';
    for (let i = 0; i < src.length; i++) {
      const c = src[i];
      if (c === '\\' && i + 1 < src.length) { out += esc(src[i + 1]); i++; continue; }
      if (c === '^') { const m = specials(src.slice(i, i + 2), false); out += m; i++; continue; }
      if (c === '?') { out += '[^\\r]'; continue; }
      if (c === '*') { out += '[^\\r]*?'; continue; }
      if (c === '@') { out += '+'; continue; }
      if (c === '<') { out += '(?<![\\p{L}\\p{N}_])(?=[\\p{L}\\p{N}_])'; continue; }
      if (c === '>') { out += '(?<=[\\p{L}\\p{N}_])(?![\\p{L}\\p{N}_])'; continue; }
      if (c === '[') {
        const j = src.indexOf(']', i + 1);
        if (j < 0) throw new Error('“查找内容”中的模式匹配表达式无效。');
        let body = src.slice(i + 1, j);
        const neg = body[0] === '!';
        if (neg) body = body.slice(1);
        out += '[' + (neg ? '^' : '') + body.replace(/\\/g, '\\\\').replace(/\]/g, '\\]') + ']';
        i = j;
        continue;
      }
      if (c === '{') { const j = src.indexOf('}', i); if (j < 0) throw new Error(L.N('无效的 {} 表达式。')); out += src.slice(i, j + 1).replace(/;/g, ','); i = j; continue; }
      if (c === '(' || c === ')' || c === '|') { out += c; continue; }
      out += esc(c);
    }
    return out;
  }
  const FORMS = ['s', 'es', 'ed', 'd', 'ing', 'er', 'ers', 'ly', "'s"];
  function soundex(w) { const a = w.toUpperCase().replace(/[^A-Z]/g, ''); if (!a) return ''; const map = { B: 1, F: 1, P: 1, V: 1, C: 2, G: 2, J: 2, K: 2, Q: 2, S: 2, X: 2, Z: 2, D: 3, T: 3, L: 4, M: 5, N: 5, R: 6 }; let out = a[0], last = map[a[0]] || 0; for (const ch of a.slice(1)) { const c = map[ch] || 0; if (c && c !== last) out += c; if (ch !== 'H' && ch !== 'W') last = c; } return (out + '000').slice(0, 4); }
  /** compile options into a matcher: returns RegExp (global, unicode) and a post-filter */
  FD.compile = function (o) {
    let src;
    if (!o.text && !o.format) throw new Error('empty');
    if (!o.text) src = '[^\\r]+';
    else if (o.wildcards) src = wildcard(o.text);
    else if (o.allForms) { const w = o.text.trim(); const stem = w.replace(/(ing|ed|es|s)$/i, ''); src = '(?<![\\p{L}\\p{N}_])' + esc(stem.length > 2 ? stem : w) + '(?:' + FORMS.map(esc).join('|') + ')?(?![\\p{L}\\p{N}_])'; }
    else if (o.soundsLike) src = '[\\p{L}]+';
    else src = specials(o.text, false);
    if (o.wholeWord && !o.wildcards && !o.allForms && !o.soundsLike && /^[\p{L}\p{N}_]/u.test(o.text) && /[\p{L}\p{N}_]$/u.test(o.text)) src = '(?<![\\p{L}\\p{N}_])' + src + '(?![\\p{L}\\p{N}_])';
    const flags = 'gu' + (o.matchCase || o.wildcards ? '' : 'i');
    let re;
    try { re = new RegExp(src, flags); } catch (e) { throw new Error('“查找内容”中的模式匹配表达式无效。'); }
    const sx = o.soundsLike ? soundex(o.text) : null;
    return { re, filter: (m) => (!sx || soundex(m[0]) === sx) };
  };
  /** does the matched range satisfy formatting criteria */
  function formatOK(o, a, b) {
    if (!o.format) return true;
    const f = o.format;
    const d = doc();
    if (f.style) { const st = d.styles[f.style]; if (st && st.type === 'paragraph' && (a.p.pPr.style || 'Normal') !== f.style) return false; }
    if (f.highlight != null) { const it = D.itemAfter(a.p, a.o); if (!!(it && it.rPr && it.rPr.hl) !== f.highlight) return false; }
    if (f.rPr) {
      for (const p of D.parasBetween(d, a.p, b.p)) {
        const s = p === a.p ? a.o : 0, e = p === b.p ? b.o : D.plen(p);
        let o2 = 0;
        for (const it of p.runs) { const n = D.ilen(it); if (n && o2 < e && o2 + n > s) { const r = D.rProps(d, p, it, LY.tblCtx(p)); for (const k in f.rPr) if (JSON.stringify(r[k] || undefined) !== JSON.stringify(f.rPr[k] || undefined) && !(f.rPr[k] === false && !r[k])) return false; } o2 += n; }
      }
    }
    return true;
  }
  /** all matches in a story: [{a, b, m}] */
  FD.all = function (o, story) {
    const st = storyText(story || doc().main);
    const { re, filter } = FD.compile(o);
    const out = [];
    re.lastIndex = 0;
    let m;
    while ((m = re.exec(st.text))) {
      if (m[0].length === 0) { re.lastIndex++; continue; }
      if (!filter(m)) continue;
      const a = toPos(st, m.index), b = toPos(st, m.index + m[0].length);
      if (!formatOK(o, a, b)) continue;
      out.push({ a, b, m, s: m.index, e: m.index + m[0].length });
      if (out.length > 20000) break;
    }
    return { matches: out, st };
  };
  const storiesFor = (o) => (o.scope === 'all' ? D.stories(doc()).filter((s) => s.kind !== 'tb') : [o.scope === 'current' && E.sel ? E.story() : doc().main]);
  /** next match after (or before) the selection, wrapping around */
  FD.next = function (o) {
    const d = doc();
    const stories = storiesFor(o);
    const curStory = E.sel ? E.story() : d.main;
    const up = o.dir === 'up';
    const order = stories.includes(curStory) ? stories.slice(stories.indexOf(curStory)).concat(stories.slice(0, stories.indexOf(curStory))) : stories;
    for (let pass = 0; pass < order.length + 1; pass++) {
      const story = order[pass % order.length];
      const { matches, st } = FD.all(o, story);
      if (!matches.length) continue;
      if (pass === 0 && E.sel && story === curStory) {
        const [a, b] = E.range();
        const from = fromPos(st, up ? a : b), cur = fromPos(st, a);
        const hit = up ? matches.slice().reverse().find((x) => x.e <= from && x.s < cur) : matches.find((x) => x.s >= from && !(x.s === cur && x.e === from && !E.collapsed() && false));
        if (hit) return hit;
        if (o.dir !== 'all' && pass === 0 && order.length === 1) { FD.wrapped = true; return up ? matches[matches.length - 1] : matches[0]; }
        continue;
      }
      return up ? matches[matches.length - 1] : matches[0];
    }
    return null;
  };
  /** replacement text for a match (wildcard groups, ^&, ^c, specials) */
  function replacementText(o, hit) {
    let r = o.replace || '';
    let out = '';
    for (let i = 0; i < r.length; i++) {
      const c = r[i];
      if (c === '\\' && o.wildcards && /\d/.test(r[i + 1] || '')) { out += hit.m[+r[i + 1]] || ''; i++; continue; }
      if (c === '^' && i + 1 < r.length) {
        const n = r[i + 1];
        i++;
        if (n === '&') { out += hit.m[0]; continue; }
        if (n === 'c' || n === 'C') { out += E.clip ? E.clip.text : ''; continue; }
        const map = { p: '\n', t: '\t', l: '\u000b', m: '\f', n: '\u000c', s: ' ', '~': '‑', '-': '­', '+': '—', '=': '–', '^': '^' };
        if (map[n] != null) { out += map[n]; continue; }
        out += '^' + n;
        continue;
      }
      out += c;
    }
    return out;
  }
  /** replace one match (inside a transaction); returns the position after the inserted text */
  function replaceHit(o, hit) {
    const text = replacementText(o, hit);
    const keepR = (() => { const it = D.itemAfter(hit.a.p, hit.a.o); return it && it.rPr ? L.clone(it.rPr) : O.inheritRPr(hit.a); })();
    const pos = O.deleteRange(hit.a, hit.b);
    if (!text && !o.replaceFormat) return pos;
    let end = pos;
    if (text) {
      /* manual breaks: \u000b line break, \f page break; \n paragraph */
      const parts = text.split(/(\u000b|\f|\u000c)/);
      for (const part of parts) {
        if (part === '\u000b') end = O.insertItem(end, D.item('br', { type: 'line' }, L.clone(keepR)));
        else if (part === '\f') end = O.insertItem(end, D.item('br', { type: 'page' }, L.clone(keepR)));
        else if (part === '\u000c') end = O.insertItem(end, D.item('br', { type: 'column' }, L.clone(keepR)));
        else if (part) end = O.insertPlain(end, part, L.clone(keepR));
      }
    }
    if (o.replaceFormat) {
      const f = o.replaceFormat;
      if (f.rPr && D.cmp(doc(), pos, end) < 0) O.setRunProps(pos, end, f.rPr);
      if (f.style && doc().styles[f.style]) O.applyStyle(pos, D.cmp(doc(), pos, end) < 0 ? end : pos, f.style);
      if (f.highlight != null && D.cmp(doc(), pos, end) < 0) O.setRunProps(pos, end, { hl: f.highlight ? A().lastHighlight || 'yellow' : undefined });
      if (!text) return pos;
    }
    return end;
  }
  FD.replaceAll = function (o) {
    let n = 0;
    E.edit('全部替换', () => {
      for (const story of storiesFor(o)) {
        const { matches } = FD.all(o, story);
        for (let i = matches.length - 1; i >= 0; i--) { replaceHit(o, matches[i]); n++; }
        O.repairFields(story);
      }
      return E.sel && D.info(doc(), E.sel.f.p) ? E.sel : D.pos(D.firstPara(doc().main.blocks), 0);
    });
    return n;
  };
  FD.replaceOne = function (o) {
    /* if the selection is exactly a match, replace it; then find the next */
    if (E.sel && !E.collapsed()) {
      const story = E.story();
      const { matches, st } = FD.all(o, story);
      const [a, b] = E.range();
      const s = fromPos(st, a), e = fromPos(st, b);
      const hit = matches.find((x) => x.s === s && x.e === e);
      if (hit) E.edit(L.N('替换'), () => replaceHit(o, hit));
    }
    return FD.next(Object.assign({}, o, { dir: o.dir === 'up' ? 'up' : 'down' }));
  };
  /* ---------- Highlight All (CSS Custom Highlight API; no document change) ---------- */
  FD.highlightAll = function (o) {
    if (!(window.CSS && CSS.highlights && window.Highlight)) return -1;
    const ranges = [];
    let n = 0;
    for (const story of storiesFor(o)) {
      const { matches } = FD.all(o, story);
      n += matches.length;
      for (const m of matches.slice(0, 3000)) {
        const a = LY.posToDom(m.a), b = LY.posToDom(m.b);
        if (!a || !b) continue;
        try { const r = new Range(); r.setStart(a.node, a.offset); r.setEnd(b.node, b.offset); ranges.push(r); } catch (e) { /* skip */ }
      }
    }
    CSS.highlights.set('quire-find', new Highlight(...ranges));
    FD.hlOpts = o;
    return n;
  };
  FD.clearHighlight = () => { if (window.CSS && CSS.highlights) CSS.highlights.delete('quire-find'); FD.hlOpts = null; };
  L.bus.on('layout-done', () => { if (FD.hlOpts) setTimeout(() => { try { FD.highlightAll(FD.hlOpts); } catch (e) { FD.clearHighlight(); } }, 0); });

  /* ---------- Go To ---------- */
  /* [kind, name, the box's label] */
  const GOTO = [['page', L.N('页面'), L.N('输入页码：')], ['section', L.N('节'), L.N('输入节号：')], ['line', L.N('线条'), L.N('输入行号：')], ['bookmark', L.N('书签'), L.N('输入书签名：')], ['comment', L.N('批注'), L.N('输入批注编号：')],
    ['footnote', L.N('脚注'), L.N('输入脚注编号：')], ['endnote', L.N('尾注'), L.N('输入尾注编号：')], ['field', L.N('域'), L.N('输入域编号：')], ['table', L.N('表格'), L.N('输入表格编号：')], ['graphic', L.N('图形'), L.N('输入图形编号：')],
    ['equation', L.N('公式'), L.N('输入公式编号：')], ['heading', L.N('标题'), L.N('输入标题编号：')]];
  FD.goto = function (kind, val, dir) {
    const d = doc();
    const rel = /^[+-]/.test(val);
    const n = parseInt(val, 10);
    if (kind === 'page') {
      if (!LY.pages.length) { A().setView('print'); }
      const cur = Math.max(0, LY.pageOf(E.sel.f));
      let i = !val ? cur + (dir || 1) : rel ? cur + n : (() => { const k = LY.pages.findIndex((p) => String(D.fmtNum(p.numVal, p.numFmt)) === String(val).trim()); return k >= 0 ? k : n - 1; })();
      i = L.clamp(i, 0, LY.pages.length - 1);
      const pg = LY.pages[i];
      const f = pg && pg.body.querySelector('.p');
      if (f) { const p = D.byId(d, +f.dataset.pid); A().gotoPos(D.pos(p, +f.dataset.from || 0), null, KEEP); pg.el.scrollIntoView({ block: 'start' }); }
      return true;
    }
    if (kind === 'line') {
      /* line numbers counted in the main story */
      const target = !val ? null : rel ? null : n;
      if (target) {
        let count = 0;
        for (const pg of LY.pages) for (const f of pg.el.querySelectorAll('.pg-body .p')) {
          const lines = LY.lines(f, f.getBoundingClientRect().top, LY.zoom);
          if (count + lines.length >= target) { const ln = lines[target - count - 1]; const p = D.byId(d, +f.dataset.pid); A().gotoPos(D.pos(p, ln ? ln.o : 0), null, KEEP); return true; }
          count += lines.length;
        }
      }
      return false;
    }
    if (kind === 'bookmark') { if (val) return A().gotoBookmark(val); return false; }
    const map = { section: 'section', comment: 'comment', footnote: 'footnote', endnote: 'endnote', field: 'field', table: 'table', graphic: 'graphic', heading: '标题', equation: 'field' };
    if (val && !rel && n > 0 && kind !== 'section') {
      /* nth item */
      E.setSel(D.pos(D.firstPara(d.main.blocks), 0));
      for (let k = 0; k < n; k++) L.panes.browseTo(map[kind], 1);
      return true;
    }
    if (kind === 'section' && val) { const secs = D.sections(d); const s = secs[L.clamp((rel ? 0 : n) - 1, 0, secs.length - 1)]; const b0 = d.main.blocks[s.from]; const p = b0.t === 'p' ? b0 : D.firstPara([b0]); A().gotoPos(D.pos(p, 0), null, KEEP); return true; }
    L.panes.browseTo(map[kind], dir || (rel && n < 0 ? -1 : 1));
    return true;
  };

  /* ================= the dialog ================= */
  let open = null;
  const state = { text: '', replace: '', matchCase: false, wholeWord: false, wildcards: false, soundsLike: false, allForms: false, dir: 'all', scope: 'main', format: null, replaceFormat: null, highlight: false };
  const hist = { find: [], repl: [] };
  G.find = function (mode) {
    if (open) { open.select(mode); return; }
    if (E.sel && !E.collapsed()) { const [a, b] = E.range(); if (a.p === b.p && b.o - a.o < 255) state.text = O.textRange(a, b, '\n'); }
    const findIn = h('input', { type: 'text', id: 'fr-find', list: 'fr-fhist', value: state.text, style: 'flex:1;min-width:0' });
    const replIn = h('input', { type: 'text', id: 'fr-repl', list: 'fr-rhist', value: state.replace, style: 'flex:1;min-width:0' });
    const fhist = h('datalist', { id: 'fr-fhist' }), rhist = h('datalist', { id: 'fr-rhist' });
    const fillHist = () => { L.clear(fhist); L.clear(rhist); for (const x of hist.find) fhist.appendChild(h('option', { value: x })); for (const x of hist.repl) rhist.appendChild(h('option', { value: x })); };
    fillHist();
    const fmtNote = h('div', { class: 'tp-note fr-fmt' });
    const rfmtNote = h('div', { class: 'tp-note fr-fmt' });
    const showFmt = () => { fmtNote.textContent = state.format ? L.t('格式：{0}', describe(state.format)) : ''; rfmtNote.textContent = state.replaceFormat ? L.t('格式：{0}', describe(state.replaceFormat)) : ''; };
    const hlChk = ui.check('突出显示在以下范围中找到的所有项目(&H)：', state.highlight, (v) => { state.highlight = v; });
    const scope = ui.select([['main', '主文档'], ['all', '所有正文区域（页眉、注释、批注）']], state.scope, (v) => { state.scope = v; }, { id: 'fr-scope' });
    const opt = (k, l) => ui.check(l, state[k], (v) => { state[k] = v; if (k === 'wildcards' && v) { soundsChk.input.checked = state.soundsLike = false; formsChk.input.checked = state.allForms = false; } });
    const caseChk = opt('matchCase', '区分大小写(&h)'), wholeChk = opt('wholeWord', '全字匹配(&y)'), wildChk = opt('wildcards', '使用通配符(&d)'), soundsChk = opt('soundsLike', '同音（英语）(&k)'), formsChk = opt('allForms', '查找单词的所有形式（英语）(&w)');
    const dir = ui.select([['all', '全部'], ['down', '向下'], ['up', '向上']], state.dir, (v) => { state.dir = v; }, { id: 'fr-dir' });
    const more = h('div', { class: 'col', hidden: !state.more }, ui.group('搜索选项', G.f('搜索：', dir), caseChk, wholeChk, wildChk, soundsChk, formsChk), h('div', { class: 'tp-btns' }, ui.button('无格式(&o)', () => { if (curTab === 1 && document.activeElement === replIn) state.replaceFormat = null; else { state.format = null; state.replaceFormat = null; } showFmt(); }, { class: 'btn small' }), ui.button('格式 ▾(&o)', (e) => formatMenu(e.currentTarget, document.activeElement === replIn ? 'replaceFormat' : curTab === 1 && state.lastFocus === 'repl' ? 'replaceFormat' : 'format'), { class: 'btn small' }), ui.button('特殊 ▾(&e)', (e) => specialMenu(e.currentTarget, state.lastFocus === 'repl' ? replIn : findIn), { class: 'btn small' })));
    findIn.addEventListener('focus', () => { state.lastFocus = 'find'; });
    replIn.addEventListener('focus', () => { state.lastFocus = 'repl'; });
    const moreBtn = ui.button(state.more ? '<< 简略(&L)' : '更多 >>(&M)', () => { state.more = !state.more; more.hidden = !state.more; moreBtn.innerHTML = ui.ampHTML(state.more ? '<< 简略(&L)' : '更多 >>(&M)'); });
    /* Go To */
    const gotoKind = G.listBox(GOTO.map(([v, l]) => ({ value: v, label: l })), 'page', { height: 150, onChange: (v) => { gotoLbl.textContent = L.t(GOTO.find((x) => x[0] === v)[2]); gotoBtn.innerHTML = ui.ampHTML(gotoIn.value ? '定位(&T)' : '下一步(&t)'); } });
    const gotoLbl = h('label', { for: 'fr-goto', text: '输入页码：' });
    const gotoIn = h('input', { type: 'text', id: 'fr-goto', style: 'width:100%' });
    const gotoBtn = ui.button('下一步(&t)', () => FD.goto(gotoKind.get(), gotoIn.value.trim(), 1));
    gotoIn.addEventListener('input', () => { gotoBtn.innerHTML = ui.ampHTML(gotoIn.value ? '定位(&T)' : '下一步(&t)'); });
    const prevBtn = ui.button('上一个(&v)', () => FD.goto(gotoKind.get(), '', -1));
    const findBody = h('div', { class: 'col' }, h('div', { class: 'row', style: 'align-items:center;flex-wrap:nowrap' }, h('label', { for: 'fr-find', html: ui.ampHTML('查找内容(&n)：'), style: 'width:80px' }), findIn), fmtNote);
    const findTabBody = h('div', { class: 'col' });
    const replTabBody = h('div', { class: 'col' });
    const gotoTabBody = h('div', { class: 'row' }, h('div', { class: 'col', style: 'width:140px' }, h('label', { html: ui.ampHTML('定位目标(&o)：') }), gotoKind), h('div', { class: 'col', style: 'flex:1;min-width:150px' }, gotoLbl, gotoIn, h('div', { class: 'tp-note', text: '输入 + 或 − 可相对当前位置移动。例如，+4 表示向后移动四个项目。' })));
    const replRow = h('div', { class: 'col' }, h('div', { class: 'row', style: 'align-items:center;flex-wrap:nowrap' }, h('label', { for: 'fr-repl', html: ui.ampHTML('替换为(&p)：'), style: 'width:80px' }), replIn), rfmtNote);
    let curTab = mode === 'replace' ? 1 : mode === 'goto' ? 2 : 0;
    let dlg = null;
    const foot = h('div', { class: 'tp-btns', style: 'justify-content:flex-end;margin-top:8px' });
    const bFind = ui.button('查找下一处(&F)', () => findNext(), { class: 'btn primary' });
    const bRepl = ui.button('替换(&R)', () => replace());
    const bAll = ui.button('全部替换(&A)', () => replaceAll());
    const bClose = ui.button('Cancel', () => dlg && dlg.close(null));
    function setButtons(t) {
      L.clear(foot);
      if (t === 2) foot.append(prevBtn, gotoBtn, ui.button('Close', () => dlg && dlg.close(null)));
      else { if (t === 1) foot.append(moreBtn, bRepl, bAll, bFind, bClose); else foot.append(moreBtn, bFind, bClose); }
      more.hidden = t === 2 || !state.more;
    }
    const tabs = ui.tabs([
      { label: '查找(&d)', body: findTabBody, onShow: () => { findTabBody.append(findBody, h('div', { class: 'row', style: 'align-items:center' }, hlChk, scope)); setButtons(0); curTab = 0; setTimeout(() => findIn.focus(), 0); } },
      { label: '替换(&p)', body: replTabBody, onShow: () => { replTabBody.append(findBody, replRow); setButtons(1); curTab = 1; setTimeout(() => findIn.focus(), 0); } },
      { label: '定位(&G)', body: gotoTabBody, onShow: () => { setButtons(2); curTab = 2; setTimeout(() => gotoIn.focus(), 0); } },
    ], mode === 'replace' ? 1 : mode === 'goto' ? 2 : 0);
    const opts = () => Object.assign({}, state, { text: findIn.value, replace: replIn.value });
    const remember = () => { const t = findIn.value; if (t && !hist.find.includes(t)) hist.find.unshift(t); hist.find.length = Math.min(hist.find.length, 12); const r = replIn.value; if (r && !hist.repl.includes(r)) hist.repl.unshift(r); fillHist(); state.text = findIn.value; state.replace = replIn.value; };
    const findNext = () => {
      remember();
      try {
        if (state.highlight && curTab === 0) { const n = FD.highlightAll(opts()); A().status(n < 0 ? '此浏览器不支持突出显示。' : n === 1 ? L.t('Quire 找到 {0} 个符合此条件的项目。', n) : L.t('Quire 找到 {0} 个符合此条件的项目。', n)); if (n > 0) { const hit = FD.next(opts()); if (hit) A().gotoPos(hit.a, hit.b, KEEP); } return; }
        FD.wrapped = false;
        const hit = FD.next(opts());
        if (!hit) { ui.msg(`Quire 已完成文档搜索，未找到搜索项。`, { icon: 'info' }); return; }
        A().gotoPos(hit.a, hit.b, KEEP);
        if (FD.wrapped) A().status('Quire 已到达文档结尾，并从开头继续搜索。');
      } catch (e) { if (e.message !== 'empty') ui.msg(e.message, { icon: 'warn' }); }
    };
    /* the dialog keeps the keyboard where it was (an edit in the page takes it there) */
    const keepKeys = (fn) => () => { const back = document.activeElement; fn(); if (back && dlg && dlg.el.contains(back) && document.activeElement !== back) back.focus(); };
    const replace = keepKeys(() => { remember(); try { const hit = FD.replaceOne(opts()); if (hit) A().gotoPos(hit.a, hit.b, KEEP); else ui.msg('Quire 已完成文档搜索。', { icon: 'info' }); } catch (e) { if (e.message !== 'empty') ui.msg(e.message, { icon: 'warn' }); } });
    const replaceAll = () => { remember(); try { const n = FD.replaceAll(opts()); ui.msg(n === 1 ? L.t('Quire 已完成文档搜索，共替换 {0} 处。', n) : n ? L.t('Quire 已完成文档搜索，共替换 {0} 处。', n) : 'Quire 已完成文档搜索，未找到搜索项。', { icon: 'info' }); } catch (e) { if (e.message !== 'empty') ui.msg(e.message, { icon: 'warn' }); } };
    const body = h('div', { class: 'col' }, tabs, more, foot);
    dlg = ui.dialog({ title: '查找和替换', width: 470, body, modeless: true, buttons: [], focus: mode === 'goto' ? gotoIn : findIn, onClose: () => { open = null; FD.clearHighlight(); E.refocus(); } });   /* the found text is the page's selection again */
    dlg.el.querySelector('.dlg-foot').remove();
    dlg.el.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && e.target.tagName === 'INPUT') { e.preventDefault(); if (curTab === 2) gotoBtn.click(); else bFind.click(); }
    });
    findIn.addEventListener('input', () => { if (FD.hlOpts) FD.clearHighlight(); });
    showFmt();
    open = { select: (m) => { tabs.select(m === 'replace' ? 1 : m === 'goto' ? 2 : 0); }, dlg };
    function describe(f) {
      const parts = [];
      if (f.rPr) for (const [k, v] of Object.entries(f.rPr)) parts.push(k === 'font' ? L.t('字体：{0}', v) : k === 'sz' ? v + ' pt' : k === 'b' ? (v ? L.t('加粗') : L.t('非加粗')) : k === 'i' ? (v ? L.t('倾斜') : L.t('非倾斜')) : k === 'u' ? (v && v !== 'none' ? L.t('下划线') : L.t('无下划线')) : k === 'color' ? L.t('字体颜色：{0}', v) : k);
      if (f.style) parts.push(L.t('样式：{0}', D.styleDisplayName(doc().styles[f.style])));
      if (f.highlight != null) parts.push(f.highlight ? L.t('突出显示') : L.t('非突出显示'));
      return parts.join(', ');
    }
    function formatMenu(anchor, key) {
      const cur = state[key] || {};
      ui.openMenu([
        { label: '字体(&F)...', run: () => G.font({ target: Object.assign({}, cur.rPr || {}), sample: L.t('示例'), onOK: (r) => { const rr = Object.assign({}, cur.rPr || {}, r); for (const k in rr) if (rr[k] === undefined) delete rr[k]; state[key] = Object.assign({}, cur, { rPr: rr }); showFmt(); } }) },
        { label: '样式(&S)...', run: () => { const ids = Object.keys(doc().styles).filter((k) => ['paragraph', 'character'].includes(doc().styles[k].type)).sort(); const lb = G.listBox(ids.map((k) => ({ value: k, label: D.styleDisplayName(doc().styles[k]) })), cur.style || 'Normal', { height: 180, translate: 'no' }); ui.dialog({ title: key === 'format' ? '查找样式' : '替换样式', width: 280, body: lb, buttons: [{ label: '确定', primary: true, onClick: () => { state[key] = Object.assign({}, cur, { style: lb.get() }); showFmt(); } }, { label: '取消' }] }); } },
        { label: '突出显示(&H)', run: () => { state[key] = Object.assign({}, cur, { highlight: cur.highlight === true ? false : true }); showFmt(); } },
      ], anchor.getBoundingClientRect());
    }
    function specialMenu(anchor, input) {
      const ins = (code) => { const s = input.selectionStart ?? input.value.length, e = input.selectionEnd ?? s; input.value = input.value.slice(0, s) + code + input.value.slice(e); input.focus(); input.setSelectionRange(s + code.length, s + code.length); };
      const isRepl = input === replIn;
      const wild = state.wildcards;
      const items = isRepl ? [['^p', '段落标记'], ['^t', '制表符'], ['^^', '脱字号'], ['^&', '查找内容文本'], ['^c', '剪贴板内容'], ['^+', '长划线'], ['^=', '短划线'], ['^l', '手动换行符'], ['^m', '手动分页符'], ['^n', '分栏符'], ['^~', '不间断连字符'], ['^s', '不间断空格'], ['^-', '可选连字符']].concat(wild ? [['\\1', '查找内容表达式']] : [])
        : wild ? [['?', '任意字符'], ['[-]', '范围内的字符'], ['<', '单词开头'], ['>', '单词结尾'], ['()', L.N('表达式')], ['[!]', L.N('非')], ['{,}', '出现次数'], ['@', '前一个或多个'], ['*', L.N('零个或多个字符')], ['^t', '制表符'], ['^p', '段落标记'], ['^l', '手动换行符'], ['^m', '手动分页符']]
          : [['^p', '段落标记'], ['^t', '制表符'], ['^?', '任意字符'], ['^#', '任意数字'], ['^$', '任意字母'], ['^^', '脱字号'], ['^+', '长划线'], ['^=', '短划线'], ['^g', L.N('图形')], ['^l', '手动换行符'], ['^m', '手动分页符'], ['^n', '分栏符'], ['^~', '不间断连字符'], ['^s', '不间断空格'], ['^-', '可选连字符'], ['^w', '空白字符']];
      ui.openMenu(items.map(([c, l]) => ({ label: l, run: () => ins(c) })), anchor.getBoundingClientRect());
    }
  };
})();
