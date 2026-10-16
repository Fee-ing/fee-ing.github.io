/* Ledger — dialogs, part 1: Format Cells, rows / columns / cells / sheets, Find and Replace, Go To,
 * Paste Special, Series, Save As, Properties, Zoom, Symbol, comments, Options, Protection, Hyperlink, About. */
(function (root) {
  'use strict';
  const L = root.L;
  const M = L.model, O = L.ops, F = L.formula, NF = L.numfmt, LY = L.layout, C = L.calc;
  const { h } = L;
  const ui = L.ui;
  const D = (L.dlg = L.dlg || {});
  const G = () => L.grid;
  const A = () => L.app;
  const sh = () => G().sheet();
  const wb = () => G().wb;
  const tryRun = (fn) => { try { fn(); } catch (e) { A().error(e); } G().paint(); G().syncObjects(); ui.refresh(); };
  const num = (v, d) => { const x = parseFloat(String(v).replace(/,/g, '')); return isNaN(x) ? d : x; };
  const list = (items, value, onPick, o) => {
    /* a Windows-style list box; o.translate 'no': its items are data (names, codes), shown as they are */
    const el = h('div', Object.assign({ class: 'lbox', role: 'listbox', tabindex: '0' }, o || {}));
    const raw = !!o && o.translate === 'no';
    let cur = -1;
    const rows = items.map((it, i) => {
      const [v, label] = Array.isArray(it) ? it : [it, it];
      const r = h('div', { class: 'lbox-i', role: 'option', text: raw ? null : label });
      if (raw) r.textContent = label;
      r._v = v;
      r.addEventListener('click', () => { set(i); if (onPick) onPick(v, i); });
      r.addEventListener('dblclick', () => { if (el.ondbl) el.ondbl(v, i); });
      el.appendChild(r);
      return r;
    });
    const set = (i, scroll) => { if (rows[cur]) rows[cur].classList.remove('on'); cur = i; if (rows[cur]) { rows[cur].classList.add('on'); if (scroll !== false) rows[cur].scrollIntoView({ block: 'nearest' }); } };
    el.addEventListener('keydown', (e) => {
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); e.stopPropagation(); const n = Math.max(0, Math.min(rows.length - 1, cur + (e.key === 'ArrowDown' ? 1 : -1))); set(n); if (onPick) onPick(rows[n]._v, n); }
    });
    el.select = (v) => { const i = rows.findIndex((r) => String(r._v) === String(v)); set(i); return i; };
    el.value = () => (rows[cur] ? rows[cur]._v : null);
    el.index = () => cur;
    if (value != null) setTimeout(() => el.select(value), 0);
    return el;
  };
  D.list = list;

  /* ================================================================ Format Cells */
  /* the categories are also NF.category's names: English here, shown translated by list() */
  /* [category, label]: the categories are also NF.category's names; three need their own words in some languages */
  const CAT_ZH = { General: '常规', Number: '数值', Currency: '货币', Accounting: '会计专用', Date: '日期', Time: '时间', Percentage: '百分比', Fraction: '分数', Scientific: '科学记数', Text: '文本', Special: '特殊', Custom: '自定义' };
  const CATS = ['General', 'Number', 'Currency', 'Accounting', 'Date', 'Time', 'Percentage', 'Fraction', 'Scientific', 'Text', 'Special', '自定义'];
  const SYMBOLS = [['$', '"$"'], ['€', '"€"'], ['£', '"£"'], ['¥', '"¥"'], [L.N('无'), '']];
  const DATE_TYPES = ['m/d/yyyy', 'dddd, mmmm dd, yyyy', 'm/d', 'm/d/yy', 'mm/dd/yy', 'd-mmm', 'd-mmm-yy', 'dd-mmm-yy', 'mmm-yy', 'mmmm-yy', 'mmmm d, yyyy', 'm/d/yy h:mm AM/PM', 'm/d/yy h:mm', 'mmmmm', 'mmmmm-yy', 'd-mmm-yyyy'];
  const TIME_TYPES = ['h:mm:ss AM/PM', 'h:mm', 'h:mm AM/PM', 'h:mm:ss', 'mm:ss.0', '[h]:mm:ss', 'm/d/yy h:mm AM/PM', 'm/d/yy h:mm'];
  const FRACTIONS = [['# ?/?', '一位分母 (1/4)'], ['# ??/??', '两位分母 (21/25)'], ['# ???/???', '三位分母 (312/943)'], ['# ?/2', '分母为 2（1/2）'], ['# ?/4', '分母为 4（2/4）'], ['# ?/8', '分母为 8（4/8）'], ['# ??/16', '分母为 16（8/16）'], ['# ?/10', '分母为 10（3/10）'], ['# ??/100', '分母为 100（30/100）']];
  const SPECIAL = [['00000', '邮政编码'], ['00000-0000', '美国邮政编码+4'], ['[<=9999999]###-####;(###) ###-####', '电话号码'], ['000-00-0000', '社会保险号码']];
  function numCode(dec, sep, neg) {
    const base = (sep ? '#,##0' : '0') + (dec ? '.' + '0'.repeat(dec) : '');
    return [base, base + ';[Red]' + base, base + '_);\\(' + base + '\\)', base + '_);[Red]\\(' + base + '\\)'][neg || 0];
  }
  function curCode(dec, sym, neg) {
    const b = sym + '#,##0' + (dec ? '.' + '0'.repeat(dec) : '');
    return [b, b + ';[Red]' + b, b + '_);\\(' + b + '\\)', b + '_);[Red]\\(' + b + '\\)'][neg || 0];
  }
  function accCode(dec, sym) {
    const d = dec ? '.' + '0'.repeat(dec) : '';
    const s = sym ? sym + '* ' : '* ';
    return `_(${s}#,##0${d}_);_(${s}\\(#,##0${d}\\);_(${s}"-"${dec ? '?'.repeat(dec) : ''}_);_(@_)`;
  }
  /** what the current code looks like in terms of the dialog's options */
  function analyse(code) {
    const cat = NF.category(code);
    const o = { cat, dec: Math.min(30, NF.decimals(code)), sep: /#,##0/.test(code), neg: 0, sym: '"$"' };
    const sec = code.split(';');
    if (sec.length > 1) { const red = /\[Red\]/i.test(sec[1]), paren = /\\\(|\(/.test(sec[1]); o.neg = paren ? (red ? 3 : 2) : red ? 1 : 0; }
    const sm = /"([$€£¥])"/.exec(code);
    if (sm) o.sym = '"' + sm[1] + '"'; else if (cat === 'Accounting' && !/\$/.test(code)) o.sym = '';
    return o;
  }
  D.formatCells = function (tab, target) {
    const s = sh(), w = wb(), sel = G().sel();
    const cell = target ? null : s.get(sel.r, sel.c);
    const st = target ? target.style : LY.styleOf(s, sel.r, sel.c, cell);
    const v = target ? 1234.5 : cell ? (cell.dirty ? C.cellValue(s, sel.r, sel.c) : cell.v) : null;
    const changed = {};
    const out = { nf: st.nf || 'General', align: Object.assign({}, st.align || {}), font: Object.assign({}, st.font || {}), fill: st.fill ? JSON.parse(JSON.stringify(st.fill)) : null, prot: Object.assign({ locked: true }, st.prot || {}) };
    /* ---------- Number ---------- */
    const an = analyse(out.nf);
    const numBody = h('div', { class: 'fc-num' });
    const sample = h('div', { class: 'fc-sample', translate: 'no' });
    const opts = h('div', { class: 'fc-opts' });
    const desc = h('div', { class: 'fc-desc' });
    const showSample = () => { let t = ''; try { t = v == null ? '' : NF.format(out.nf, v, { date1904: w.date1904 }).text; } catch (e) { t = ''; } sample.textContent = t; };
    const setCode = (code) => { out.nf = code; changed.nf = true; showSample(); };
    let catState = an.cat;
    const catList = list(CATS.map((c) => [c, CAT_ZH[c]]), an.cat, (c) => { catState = c; drawOpts(true); });
    function drawOpts(fromUser) {
      opts.textContent = '';
      const c = catState;
      const decSpin = (d, fn) => ui.spin({ value: d, min: 0, max: 30, step: 1, dec: 0, onChange: (x) => fn(Math.round(x)) });
      if (c === 'General') { desc.textContent = L.t('常规格式的单元格没有特定的数字格式。'); if (fromUser) setCode('General'); }
      else if (c === 'Number') {
        let dec = an.cat === 'Number' ? an.dec : 2, sep = an.cat === 'Number' ? an.sep : false, neg = an.cat === 'Number' ? an.neg : 0;
        const negList = list([0, 1, 2, 3].map((k) => [k, k]), neg, (k) => { neg = k; upd(); }, { class: 'lbox short' });
        const drawNeg = () => { Array.from(negList.children).forEach((r, k) => { const code = numCode(dec, sep, k); r.textContent = NF.text(code, -1234.1); r.style.color = /Red/.test(code) ? '#FF0000' : ''; }); };
        const upd = () => { setCode(numCode(dec, sep, neg)); drawNeg(); };
        opts.append(ui.field('小数位数(&D)：', decSpin(dec, (x) => { dec = x; upd(); })), ui.check('使用千位分隔符 (,)(&U)', sep, (x) => { sep = x; upd(); }), h('label', { text: '负数：' }), negList);
        drawNeg();
        desc.textContent = L.t('数字格式用于一般数值显示。货币和会计专用格式则用于货币值。');
        if (fromUser) upd();
      } else if (c === 'Currency') {
        let dec = an.cat === 'Currency' ? an.dec : 2, sym = an.cat === 'Currency' ? an.sym : '"$"', neg = an.cat === 'Currency' ? an.neg : 3;
        const negList = list([0, 1, 2, 3].map((k) => [k, k]), neg, (k) => { neg = k; upd(); }, { class: 'lbox short' });
        const drawNeg = () => { Array.from(negList.children).forEach((r, k) => { const code = curCode(dec, sym, k); r.textContent = NF.text(code, -1234.1); r.style.color = /Red/.test(code) ? '#FF0000' : ''; }); };
        const upd = () => { setCode(curCode(dec, sym, neg)); drawNeg(); };
        opts.append(ui.field('小数位数(&D)：', decSpin(dec, (x) => { dec = x; upd(); })), ui.field('符号(&S)：', ui.select(SYMBOLS.map(([l, cde]) => [cde, l]), sym, (x) => { sym = x; upd(); })), h('label', { text: '负数：' }), negList);
        drawNeg();
        desc.textContent = L.t('货币格式用于一般货币数值。若要对齐一列中的小数点，请使用会计专用格式。');
        if (fromUser) upd();
      } else if (c === 'Accounting') {
        let dec = an.cat === 'Accounting' ? an.dec : 2, sym = an.cat === 'Accounting' ? an.sym : '"$"';
        const upd = () => setCode(accCode(dec, sym));
        opts.append(ui.field('小数位数(&D)：', decSpin(dec, (x) => { dec = x; upd(); })), ui.field('符号(&S)：', ui.select(SYMBOLS.map(([l, cde]) => [cde, l]), sym, (x) => { sym = x; upd(); })));
        desc.textContent = L.t('会计专用格式可对齐一列中的货币符号和小数点。');
        if (fromUser) upd();
      } else if (c === 'Date' || c === 'Time') {
        const types = c === 'Date' ? DATE_TYPES : TIME_TYPES;
        const sv = typeof v === 'number' ? v : 37329.5649;
        const lb = list(types.map((t) => [t, (t === 'm/d/yyyy' || t === 'h:mm:ss AM/PM' ? '*' : '') + NF.text(t, sv, { date1904: w.date1904 })]), types.includes(out.nf) ? out.nf : null, (t) => setCode(t), { translate: 'no' });
        opts.append(h('label', { text: '类型：' }), lb, ui.field('区域设置（位置）(&L)：', ui.select([['en-US', '英语（美国）']], 'en-US')));
        desc.textContent = c === 'Date' ? L.t('日期格式将日期和时间序列号显示为日期值。以星号 (*) 开头的格式会随区域日期和时间设置的更改而变化。') : L.t('时间格式将日期和时间序列号显示为时间值。');
        if (fromUser && !types.includes(out.nf)) { setCode(types[0]); lb.select(types[0]); }
      } else if (c === 'Percentage' || c === 'Scientific') {
        let dec = an.cat === c ? an.dec : 2;
        const upd = () => setCode(c === 'Percentage' ? '0' + (dec ? '.' + '0'.repeat(dec) : '') + '%' : '0' + (dec ? '.' + '0'.repeat(dec) : '') + 'E+00');
        opts.append(ui.field('小数位数(&D)：', decSpin(dec, (x) => { dec = x; upd(); })));
        desc.textContent = c === 'Percentage' ? L.t('百分比格式将单元格值乘以 100，并在结果后显示百分号。') : '';
        if (fromUser) upd();
      } else if (c === 'Fraction') {
        const lb = list(FRACTIONS, FRACTIONS.some((f) => f[0] === out.nf) ? out.nf : null, (t) => setCode(t));
        opts.append(h('label', { text: '类型：' }), lb);
        if (fromUser && !FRACTIONS.some((f) => f[0] === out.nf)) { setCode(FRACTIONS[0][0]); lb.select(FRACTIONS[0][0]); }
      } else if (c === 'Text') { desc.textContent = L.t('文本格式的单元格始终按文本处理，即使输入的是数字。单元格会按输入的原样显示。'); if (fromUser) setCode('@'); }
      else if (c === 'Special') {
        const lb = list(SPECIAL, SPECIAL.some((f) => f[0] === out.nf) ? out.nf : null, (t) => setCode(t));
        opts.append(h('label', { text: '类型：' }), lb, ui.field('区域设置（位置）(&L)：', ui.select([['en-US', '英语（美国）']], 'en-US')));
        desc.textContent = L.t('特殊格式便于跟踪列表和数据库中的值。');
        if (fromUser && !SPECIAL.some((f) => f[0] === out.nf)) { setCode(SPECIAL[0][0]); lb.select(SPECIAL[0][0]); }
      } else {
        const used = new Set();
        for (const x of w.styles.list) if (x.nf && x.nf !== 'General') used.add(x.nf);
        const codes = ['General', '0', '0.00', '#,##0', '#,##0.00', '#,##0_);(#,##0)', '#,##0_);[Red](#,##0)', '#,##0.00_);(#,##0.00)', '#,##0.00_);[Red](#,##0.00)', '"$"#,##0_);\\("$"#,##0\\)', '"$"#,##0_);[Red]\\("$"#,##0\\)', '"$"#,##0.00_);\\("$"#,##0.00\\)', '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)', '0%', '0.00%', '0.00E+00', '##0.0E+0', '# ?/?', '# ??/??', 'm/d/yyyy', 'd-mmm-yy', 'd-mmm', 'mmm-yy', 'h:mm AM/PM', 'h:mm:ss AM/PM', 'h:mm', 'h:mm:ss', 'm/d/yyyy h:mm', 'mm:ss', 'mm:ss.0', '@', '[h]:mm:ss', NF.BUILTIN[41], NF.BUILTIN[42], NF.BUILTIN[43], NF.BUILTIN[44]];
        for (const u of used) if (!codes.includes(u)) codes.push(u);
        const inp = h('input', { type: 'text', value: out.nf, style: 'width:100%' });
        const lb = list(codes, codes.includes(out.nf) ? out.nf : null, (t) => { inp.value = t; setCode(t); }, { translate: 'no' });
        inp.addEventListener('input', () => { try { NF.parse(inp.value); setCode(inp.value || 'General'); } catch (e) { /* keep */ } });
        const del = ui.button('删除(&D)', () => {
          const code = lb.value();
          if (!code || NF.builtinId(code) >= 0 || code === 'General') return;
          /* cells using the deleted format return to General */
          O.tx(w, '删除数字格式', () => { for (const s2 of w.sheets) s2.rows.forEach((row, r) => { if (row) row.cells.forEach((cl, c) => { if (cl && cl.s && w.styles.get(cl.s).nf === code) { const n = Object.assign({}, cl); n.s = w.styles.derive(cl.s, { nf: 'General' }); O.put(s2, r, c, n); } }); }); });
          drawOpts(false);
        });
        opts.append(ui.field('类型(&T)：', inp), lb, h('div', { class: 'row' }, del));
        desc.textContent = L.t('以现有代码为基础，输入数字格式代码。');
        if (fromUser && !changed.nf) setCode(out.nf);
      }
    }
    numBody.append(h('div', { class: 'fc-cats' }, h('label', { text: '类别：' }), catList), h('div', { class: 'fc-right' }, ui.group('示例', sample), opts, desc));
    drawOpts(false);
    showSample();
    /* ---------- Alignment ---------- */
    const al = out.align;
    const hSel = ui.select([['general', '常规'], ['left', '靠左（缩进）'], ['center', '居中'], ['right', '靠右（缩进）'], ['fill', '填充'], ['justify', '两端对齐'], ['centerContinuous', '跨列居中'], ['distributed', '分散对齐（缩进）']], al.h || 'general', (x) => { al.h = x === 'general' ? undefined : x; changed.align = true; indent.setDisabled(!/left|right|distributed/.test(x)); });
    const vSel = ui.select([['top', '顶端'], ['center', '居中'], ['bottom', '底端'], ['justify', '两端对齐'], ['distributed', '分散对齐']], al.v || 'bottom', (x) => { al.v = x === 'bottom' ? undefined : x; changed.align = true; });
    const indent = ui.spin({ value: al.indent || 0, min: 0, max: 15, step: 1, dec: 0, onChange: (x) => { al.indent = Math.round(x) || undefined; changed.align = true; } });
    indent.setDisabled(!/left|right|distributed/.test(al.h || ''));
    const rot0 = al.rot === 255 ? 0 : al.rot > 90 ? 90 - al.rot : al.rot || 0;
    const rotSpin = ui.spin({ value: rot0, min: -90, max: 90, step: 1, dec: 0, fmt: (x) => L.t('{0} 度', Math.round(x)), onChange: (x) => { const d = Math.round(x); al.rot = d < 0 ? 90 - d : d || undefined; vert.classList.remove('on'); changed.align = true; drawDial(); } });
    const vert = h('button', { type: 'button', class: 'fc-vert' + (al.rot === 255 ? ' on' : ''), html: Array.from(L.t('文本')).map(L.esc).join('<br>') });
    vert.addEventListener('click', () => { const on = !vert.classList.contains('on'); vert.classList.toggle('on', on); al.rot = on ? 255 : undefined; changed.align = true; drawDial(); });
    const dial = h('div', { class: 'fc-dial' });
    const drawDial = () => {
      const deg = al.rot === 255 ? 0 : al.rot > 90 ? 90 - al.rot : al.rot || 0;
      dial.innerHTML = `<svg width="110" height="130" viewBox="0 0 110 130">${Array.from({ length: 13 }, (_, i) => { const a = (-90 + i * 15) * Math.PI / 180; return `<circle cx="${20 + 70 * Math.cos(a)}" cy="${65 - 70 * Math.sin(a)}" r="${i % 6 === 0 ? 2.5 : 1.5}" fill="${i % 6 === 0 ? '#000' : '#888'}"/>`; }).join('')}<line x1="20" y1="65" x2="${20 + 64 * Math.cos(deg * Math.PI / 180)}" y2="${65 - 64 * Math.sin(deg * Math.PI / 180)}" stroke="#c00" stroke-width="1.5"/><text x="26" y="69" font-size="11" transform="rotate(${-deg} 20 65)" fill="${al.rot === 255 ? '#aaa' : '#000'}">${L.esc(L.t('Text'))} —</text></svg>`;
    };
    dial.addEventListener('click', (e) => { const r = dial.getBoundingClientRect(); const x = e.clientX - r.left - 20, y = 65 - (e.clientY - r.top); let d = Math.round(Math.atan2(y, Math.max(0, x)) * 180 / Math.PI); d = Math.max(-90, Math.min(90, d)); rotSpin.set(d); al.rot = d < 0 ? 90 - d : d || undefined; vert.classList.remove('on'); changed.align = true; drawDial(); });
    drawDial();
    const merged = !!LY.mergeAt(s, sel.r, sel.c);
    let mergeState = null;
    const alBody = h('div', { class: 'fc-align' },
      h('div', { class: 'col', style: 'flex:1' },
        ui.group('文本对齐方式', h('div', { class: 'row', style: 'gap:8px;align-items:flex-end' }, ui.field('水平(&H)：', hSel), ui.field('缩进(&I)：', indent)), ui.field('垂直(&V)：', vSel), ui.check('分散对齐(&J)', !!al.justLast, (x) => { al.justLast = x || undefined; changed.align = true; })),
        ui.group('文本控件', ui.check('自动换行(&W)', !!al.wrap, (x) => { al.wrap = x || undefined; changed.align = true; }), ui.check('缩小字体填充(&k)', !!al.shrink, (x) => { al.shrink = x || undefined; changed.align = true; }), ui.check('合并单元格(&M)', merged, (x) => { mergeState = x; })),
        ui.group('从右到左', ui.field('文字方向(&T)：', ui.select([['ctx', '上下文'], ['ltr', '从左到右'], ['rtl', '从右到左']], al.rtl ? 'rtl' : al.ltr ? 'ltr' : 'ctx', (x) => { al.rtl = x === 'rtl' || undefined; al.ltr = x === 'ltr' || undefined; changed.align = true; })))),
      ui.group('方向', h('div', { class: 'row' }, vert, dial), ui.field('', rotSpin)));
    /* ---------- Font ---------- */
    const f = out.font;
    const fontName = h('input', { type: 'text', value: LY.fontName(w, f), style: 'width:100%' });
    const fontList = list(L.FONT_LIST, LY.fontName(w, f), (x) => { fontName.value = x; f.name = x; delete f.scheme; changed.font = true; prev(); }, { class: 'lbox short', translate: 'no' });
    fontName.addEventListener('input', () => { f.name = fontName.value; delete f.scheme; changed.font = true; prev(); });
    const styles = [['r', L.N('常规')], ['i', L.N('倾斜')], ['b', L.N('加粗')], ['bi', '加粗、倾斜']];
    const curStyle = (f.b ? 'b' : '') + (f.i ? 'i' : '') || 'r';
    const styleList = list(styles, curStyle, (x) => { f.b = /b/.test(x) || undefined; f.i = /i/.test(x) || undefined; changed.font = true; prev(); }, { class: 'lbox short' });
    const sizeIn = h('input', { type: 'text', value: String(f.sz || 10), style: 'width:100%' });
    const sizeList = list(L.SIZE_LIST, f.sz || 10, (x) => { sizeIn.value = x; f.sz = +x; changed.font = true; prev(); }, { class: 'lbox short' });
    sizeIn.addEventListener('input', () => { const x = num(sizeIn.value, null); if (x >= 1 && x <= 409) { f.sz = x; changed.font = true; prev(); } });
    const uSel = ui.select([['none', '无'], ['single', '单线'], ['double', '双线'], ['singleAccounting', '会计用单下划线'], ['doubleAccounting', '会计用双下划线']], f.u || 'none', (x) => { f.u = x === 'none' ? undefined : x; changed.font = true; prev(); });
    const fcol = h('button', { type: 'button', class: 'swatch-btn wide', style: `background:${M.colorHex(w, f.color, '#000000')}` }, h('span', { text: f.color && !f.color.auto ? '' : '自动' }), h('span', { class: 'dd-arrow' }));
    fcol.addEventListener('click', () => ui.colorMenu(fcol, { mode: 'font', grid: true }, (c) => { if (c.auto) { f.color = undefined; fcol.style.background = '#000'; fcol.firstChild.textContent = L.t('自动'); } else if (typeof c === 'string') { f.color = M.rgb(c); fcol.style.background = c; fcol.firstChild.textContent = ''; } changed.font = true; prev(); }));
    const normal = ui.check('默认字体(&N)', false, (x) => { if (x) { Object.keys(f).forEach((k) => delete f[k]); Object.assign(f, JSON.parse(JSON.stringify(w.styles.get(0).font))); fontName.value = LY.fontName(w, f); sizeIn.value = f.sz; changed.font = true; prev(); } });
    const fprev = h('div', { class: 'fc-fprev', text: 'AaBbCcYyZz' });
    function prev() {
      fprev.style.font = LY.cssFont(Object.assign({}, f, { name: LY.fontName(w, f) }), 1.3);
      fprev.style.color = M.colorHex(w, f.color, '#000000');
      fprev.style.textDecoration = [f.u ? 'underline' : '', f.strike ? 'line-through' : ''].join(' ').trim() || 'none';
      fprev.style.textDecorationStyle = /double/.test(f.u || '') ? 'double' : 'solid';
      fprev.style.verticalAlign = f.vert === 'superscript' ? 'super' : f.vert === 'subscript' ? 'sub' : '';
    }
    prev();
    const fontBody = h('div', { class: 'fc-font' },
      h('div', { class: 'fc-f3' }, h('div', { class: 'col' }, h('label', { text: '字体：' }), fontName, fontList), h('div', { class: 'col' }, h('label', { text: '字形：' }), styleList), h('div', { class: 'col' }, h('label', { text: '大小：' }), sizeIn, sizeList)),
      h('div', { class: 'row', style: 'gap:12px;align-items:flex-end' }, ui.field('下划线(&U)：', uSel), ui.field('颜色(&C)：', fcol), normal),
      h('div', { class: 'row', style: 'gap:12px;align-items:stretch' },
        ui.group('效果', ui.check('删除线(&k)', !!f.strike, (x) => { f.strike = x || undefined; changed.font = true; prev(); }), ui.check('上标(&e)', f.vert === 'superscript', (x) => { f.vert = x ? 'superscript' : undefined; changed.font = true; prev(); }), ui.check('下标(&b)', f.vert === 'subscript', (x) => { f.vert = x ? 'subscript' : undefined; changed.font = true; prev(); })),
        ui.group('预览', fprev)));
    /* ---------- Border ---------- */
    const rg = G().range();
    const multiR = rg.r1 !== rg.r2, multiC = rg.c1 !== rg.c2;
    const b0 = st.border || {};
    const edges = { top: b0.t || null, bottom: b0.b || null, left: b0.l || null, right: b0.r || null, insideH: null, insideV: null, diagUp: b0.du ? b0.d : null, diagDown: b0.dd ? b0.d : null };
    if (multiR || multiC) { const right = LY.styleOf(s, rg.r1, rg.c2, s.get(rg.r1, rg.c2)).border || {}, bottom = LY.styleOf(s, rg.r2, rg.c1, s.get(rg.r2, rg.c1)).border || {}; edges.right = right.r || null; edges.bottom = bottom.b || null; if (multiR) edges.insideH = b0.b || null; if (multiC) edges.insideV = b0.r || null; }
    const touched = {};
    let line = { style: 'thin' };
    const LINES = ['hair', 'dotted', 'dashDotDot', 'dashDot', 'dashed', 'thin', 'mediumDashDotDot', 'slantDashDot', 'mediumDashDot', 'mediumDashed', 'medium', 'thick', 'double'];
    const lineBox = h('div', { class: 'fc-lines' });
    const lineSvg = (k) => { const w2 = 52, y = 7; const dash = { hair: '1 1', dotted: '2 2', dashDotDot: '6 2 2 2 2 2', dashDot: '6 2 2 2', dashed: '4 2', mediumDashDotDot: '6 2 2 2 2 2', slantDashDot: '8 2 3 2', mediumDashDot: '6 2 2 2', mediumDashed: '6 3' }[k] || ''; const sw = /medium|slant/.test(k) ? 2 : k === 'thick' ? 3 : 1; return k === 'double' ? `<svg width="${w2}" height="14"><line x1="2" y1="5" x2="${w2 - 2}" y2="5" stroke="#000"/><line x1="2" y1="9" x2="${w2 - 2}" y2="9" stroke="#000"/></svg>` : `<svg width="${w2}" height="14"><line x1="2" y1="${y}" x2="${w2 - 2}" y2="${y}" stroke="#000" stroke-width="${sw}" ${dash ? `stroke-dasharray="${dash}"` : ''}/></svg>`; };
    const LINE_NAMES = { none: L.N('无'), hair: L.N('细线'), dotted: L.N('点线'), dashDotDot: '双点划线', dashDot: '点划线', dashed: L.N('虚线'), thin: L.N('细'), mediumDashDotDot: '中粗双点划线', slantDashDot: '倾斜点划线', mediumDashDot: '中粗点划线', mediumDashed: '中粗虚线', medium: L.N('中等'), thick: L.N('粗'), double: L.N('双线') };
    for (const k of ['none'].concat(LINES)) { const b = h('button', { type: 'button', class: 'fc-line' + (k === 'thin' ? ' on' : ''), html: k === 'none' ? `<span style="font-size:10px">${L.esc(L.t('None'))}</span>` : lineSvg(k), 'aria-label': LINE_NAMES[k] }); b.addEventListener('click', () => { L.$$('.fc-line', lineBox).forEach((x) => x.classList.remove('on')); b.classList.add('on'); line = k === 'none' ? null : Object.assign({}, line || {}, { style: k }); }); lineBox.appendChild(b); }
    let lineColor = null;
    const lcol = h('button', { type: 'button', class: 'swatch-btn wide', style: 'background:#000' }, h('span', { text: '自动' }), h('span', { class: 'dd-arrow' }));
    lcol.addEventListener('click', () => ui.colorMenu(lcol, { mode: 'font', grid: true }, (c) => { if (c.auto) { lineColor = null; lcol.style.background = '#000'; lcol.firstChild.textContent = L.t('自动'); } else if (typeof c === 'string') { lineColor = M.rgb(c); lcol.style.background = c; lcol.firstChild.textContent = ''; } }));
    const prevBox = h('div', { class: 'fc-bprev' });
    const curLine = () => (line ? Object.assign({ style: line.style }, lineColor ? { color: lineColor } : {}) : null);
    const toggle = (k) => { const cur = edges[k]; const ln = curLine(); edges[k] = cur && ln && cur.style === ln.style && JSON.stringify(cur.color || null) === JSON.stringify(ln.color || null) ? null : ln; touched[k] = true; drawB(); };
    function drawB() {
      const W2 = 150, H2 = 110, x0 = 15, y0 = 10, x1 = W2 - 15, y1 = H2 - 10, mx = (x0 + x1) / 2, my = (y0 + y1) / 2, word = L.esc(L.t('文本'));
      const ln = (e, xa, ya, xb, yb) => { if (!e) return ''; const c = M.colorHex(w, e.color, '#000000'); const sw = /medium|slant/.test(e.style) ? 2 : e.style === 'thick' ? 3 : 1; const dash = /dash|dot|hair/i.test(e.style) ? 'stroke-dasharray="3 2"' : ''; return e.style === 'double' ? `<line x1="${xa}" y1="${ya}" x2="${xb}" y2="${yb}" stroke="${c}" stroke-width="3"/><line x1="${xa}" y1="${ya}" x2="${xb}" y2="${yb}" stroke="#fff" stroke-width="1"/>` : `<line x1="${xa}" y1="${ya}" x2="${xb}" y2="${yb}" stroke="${c}" stroke-width="${sw}" ${dash}/>`; };
      prevBox.innerHTML = `<svg width="${W2}" height="${H2}" viewBox="0 0 ${W2} ${H2}"><rect x="${x0}" y="${y0}" width="${x1 - x0}" height="${y1 - y0}" fill="#fff"/>` +
        ['M' + x0 + ',' + (y0 - 6) + 'v4', 'M' + (x0 - 6) + ',' + y0 + 'h4', 'M' + x1 + ',' + (y0 - 6) + 'v4', 'M' + (x1 + 2) + ',' + y0 + 'h4', 'M' + x0 + ',' + (y1 + 2) + 'v4', 'M' + (x0 - 6) + ',' + y1 + 'h4', 'M' + x1 + ',' + (y1 + 2) + 'v4', 'M' + (x1 + 2) + ',' + y1 + 'h4'].map((d) => `<path d="${d}" stroke="#888"/>`).join('') +
        `<text x="${multiC ? (x0 + mx) / 2 : mx}" y="${multiR ? (y0 + my) / 2 + 4 : my + 4}" font-size="11" text-anchor="middle" fill="#666">${word}</text>` + (multiC ? `<text x="${(mx + x1) / 2}" y="${multiR ? (y0 + my) / 2 + 4 : my + 4}" font-size="11" text-anchor="middle" fill="#666">${word}</text>` : '') + (multiR ? `<text x="${multiC ? (x0 + mx) / 2 : mx}" y="${(my + y1) / 2 + 4}" font-size="11" text-anchor="middle" fill="#666">${word}</text>` : '') +
        ln(edges.top, x0, y0, x1, y0) + ln(edges.bottom, x0, y1, x1, y1) + ln(edges.left, x0, y0, x0, y1) + ln(edges.right, x1, y0, x1, y1) + (multiR ? ln(edges.insideH, x0, my, x1, my) : '') + (multiC ? ln(edges.insideV, mx, y0, mx, y1) : '') + ln(edges.diagUp, x0, y1, x1, y0) + ln(edges.diagDown, x0, y0, x1, y1) + '</svg>';
    }
    prevBox.addEventListener('click', (e) => {
      const r = prevBox.getBoundingClientRect(); const x = e.clientX - r.left, y = e.clientY - r.top;
      const W2 = 150, H2 = 110;
      const d = { top: Math.abs(y - 10), bottom: Math.abs(y - (H2 - 10)), left: Math.abs(x - 15), right: Math.abs(x - (W2 - 15)) };
      if (multiR) d.insideH = Math.abs(y - H2 / 2); if (multiC) d.insideV = Math.abs(x - W2 / 2);
      const k = Object.keys(d).sort((a, b) => d[a] - d[b])[0];
      if (d[k] < 12) toggle(k);
    });
    drawB();
    const bbtn = (k, icon, tip, en) => { const b = h('button', { type: 'button', class: 'fc-bb', html: L.icons.get(icon), 'data-tip': tip, 'aria-label': tip }); b.disabled = en === false; b.addEventListener('click', () => toggle(k)); return b; };
    const preset = (label, icon, fn, en) => { const b = h('button', { type: 'button', class: 'fc-preset', html: L.icons.get(icon) + `<span>${L.esc(L.t(label))}</span>` }); b.disabled = en === false; b.addEventListener('click', fn); return b; };
    const borderBody = h('div', { class: 'fc-border' },
      h('div', { class: 'col' }, h('label', { text: '预设' }), h('div', { class: 'row', style: 'gap:16px' },
        preset(L.N('无'), 'bordersNone', () => { for (const k in edges) { edges[k] = null; touched[k] = true; } drawB(); }),
        preset(L.N('外边框'), 'bordersOut', () => { for (const k of ['top', 'bottom', 'left', 'right']) { edges[k] = curLine(); touched[k] = true; } drawB(); }),
        preset(L.N('内部'), 'bordersInside', () => { if (multiR) { edges.insideH = curLine(); touched.insideH = true; } if (multiC) { edges.insideV = curLine(); touched.insideV = true; } drawB(); }, multiR || multiC)),
      h('label', { text: '边框' }),
      h('div', { class: 'fc-bgrid' }, h('div', { class: 'col' }, bbtn('top', 'bordersTop', L.N('顶端')), bbtn('insideH', 'bordersInsideH', '内部水平线', multiR), bbtn('bottom', 'bordersBottom', L.N('底端'))), prevBox),
      h('div', { class: 'row', style: 'gap:4px;margin-left:34px' }, bbtn('diagUp', 'bordersDiagUp', '上对角线'), bbtn('left', 'bordersLeft', L.N('左')), bbtn('insideV', 'bordersInsideV', '内部垂直线', multiC), bbtn('right', 'bordersRight', L.N('右')), bbtn('diagDown', 'bordersDiagDown', '下对角线')),
      h('div', { class: 'hint', text: '单击预设、预览示意图或上方按钮可应用所选边框样式。' })),
      ui.group('线条', h('label', { text: '样式：' }), lineBox, ui.field('颜色(&C)：', lcol)));
    /* ---------- Patterns ---------- */
    const fill = out.fill;
    let fillColor = fill && fill.pattern === 'solid' ? M.colorHex(w, fill.fg) : fill && fill.bg ? M.colorHex(w, fill.bg) : null;
    let pattern = fill && fill.pattern && fill.pattern !== 'solid' ? fill.pattern : null;
    let patColor = fill && fill.pattern && fill.pattern !== 'solid' ? M.colorHex(w, fill.fg, '#000000') : '#000000';
    const sampleP = h('div', { class: 'fc-psample' });
    const showP = () => { sampleP.style.background = fillColor || '#fff'; sampleP.style.backgroundImage = pattern ? `url("${patternURL(pattern, patColor, fillColor || '#FFFFFF')}")` : 'none'; };
    const pal = h('div', { class: 'fc-pal' });
    pal.appendChild(h('button', { type: 'button', class: 'cp-wide', text: '无颜色', onclick: () => { fillColor = null; changed.fill = true; showP(); } }));
    const grid = h('div', { class: 'cp-grid' });
    for (const c of L.color.STANDARD) grid.appendChild(h('button', { type: 'button', class: 'cp-sw', style: `background:${c}`, 'aria-label': ui.colorName(c), 'data-tip': ui.colorName(c), onclick: () => { fillColor = c; changed.fill = true; showP(); } }));
    pal.appendChild(grid);
    const PATTERNS = ['solid', 'darkGray', 'mediumGray', 'lightGray', 'gray125', 'gray0625', 'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis', 'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', 'lightGrid', 'lightTrellis'];
    /* each pattern's name, as Excel's tips name them */
    const PAT_NAMES = { solid: L.N('实心'), darkGray: L.N('75% 灰度'), mediumGray: L.N('50% 灰度'), lightGray: L.N('25% 灰度'), gray125: L.N('12.5% 灰度'), gray0625: L.N('6.25% 灰度'), darkHorizontal: '水平条纹', darkVertical: '垂直条纹', darkDown: '反向斜条纹', darkUp: '斜条纹', darkGrid: '斜向交叉线', darkTrellis: '粗斜向交叉线',
      lightHorizontal: '细水平条纹', lightVertical: '细垂直条纹', lightDown: '细反向斜条纹', lightUp: '细斜条纹', lightGrid: '细水平交叉线', lightTrellis: '细斜向交叉线' };
    const patBtn = h('button', { type: 'button', class: 'swatch-btn wide', text: '图案' }, h('span', { class: 'dd-arrow' }));
    patBtn.addEventListener('click', () => {
      const r = patBtn.getBoundingClientRect();
      ui.openMenu([{ custom: (close) => {
        const g2 = h('div', { class: 'pat-grid' });
        for (const p of PATTERNS) { const b = h('button', { type: 'button', style: `background:#fff url("${patternURL(p, '#000', '#fff')}")`, 'aria-label': PAT_NAMES[p], 'data-tip': PAT_NAMES[p], 'data-pattern': p }); b.addEventListener('click', () => { close(); pattern = p === 'solid' ? null : p; changed.fill = true; showP(); }); g2.appendChild(b); }
        const g3 = h('div', { class: 'cp-grid' });
        for (const c of L.color.STANDARD) g3.appendChild(h('button', { type: 'button', class: 'cp-sw', style: `background:${c}`, 'aria-label': ui.colorName(c), 'data-tip': ui.colorName(c), onclick: () => { close(); patColor = c; changed.fill = true; showP(); } }));
        return h('div', null, g2, h('div', { class: 'cp-sep' }), h('div', { class: 'tp-note', text: '图案颜色：' }), g3);
      } }], { left: r.left, bottom: r.bottom });
    });
    showP();
    const patBody = h('div', { class: 'fc-pat' }, ui.group('单元格底纹', h('label', { text: '颜色：' }), pal, h('div', { class: 'row', style: 'gap:8px;align-items:center' }, h('span', { text: '图案：' }), patBtn)), ui.group('示例', sampleP));
    /* ---------- Protection ---------- */
    const pr = out.prot;
    const protBody = h('div', { class: 'col' }, ui.check('锁定(&L)', pr.locked !== false, (x) => { pr.locked = x; changed.prot = true; }), ui.check('隐藏(&i)', !!pr.hidden, (x) => { pr.hidden = x || undefined; changed.prot = true; }),
      h('p', { class: 'hint', text: '只有保护工作表后，锁定单元格或隐藏公式才会生效。要保护工作表，请在“工具”菜单中选择“保护”，再选择“保护工作表”。密码为可选项。' }));
    const tabs = ui.tabs([{ label: '数字', body: numBody }, { label: '对齐方式', body: alBody }, { label: '字体', body: fontBody }, { label: '边框', body: borderBody }, { label: '图案', body: patBody }, { label: '保护', body: protBody }], tab || 0);
    ui.dialog({
      title: target && target.title ? target.title : '设置单元格格式', body: tabs, width: 520,
      buttons: [{ label: '确定', primary: true, onClick: () => {
        if (target) {
          /* modifying a named style: build the whole style */
          const d = { nf: out.nf, align: Object.keys(al).some((k) => al[k] !== undefined) ? Object.assign({}, al) : null, font: Object.assign({}, f), fill: fillColor || pattern ? (pattern ? { pattern, fg: M.rgb(patColor), bg: fillColor ? M.rgb(fillColor) : M.rgb('#FFFFFF') } : { pattern: 'solid', fg: M.rgb(fillColor), bg: { indexed: 64 } }) : null, prot: pr.locked === false || pr.hidden ? { locked: pr.locked === false ? false : undefined, hidden: pr.hidden || undefined } : null };
          const b = {};
          if (edges.top) b.t = edges.top; if (edges.bottom) b.b = edges.bottom; if (edges.left) b.l = edges.left; if (edges.right) b.r = edges.right;
          if (edges.diagUp || edges.diagDown) { b.d = edges.diagUp || edges.diagDown; if (edges.diagUp) b.du = true; if (edges.diagDown) b.dd = true; }
          d.border = Object.keys(b).length ? b : null;
          const nst = M.mergeStyle(st, d);
          nst.border = d.border; nst.align = d.align; nst.fill = d.fill; nst.prot = d.prot;
          for (const k of Object.keys(nst.font || {})) if (nst.font[k] === undefined) delete nst.font[k];
          delete nst.xs;
          target.apply(nst);
          return;
        }
        tryRun(() => {
          O.tx(w, '设置单元格格式', () => {
            const delta = {};
            if (changed.nf) delta.nf = out.nf;
            if (changed.align) { const a = {}; for (const k of ['h', 'v', 'wrap', 'indent', 'rot', 'shrink', 'rtl', 'ltr', 'justLast']) a[k] = al[k]; delta.align = a; }
            if (changed.font) { const fo = Object.assign({}, f); for (const k of ['b', 'i', 'u', 'strike', 'vert', 'color']) if (fo[k] === undefined) fo[k] = undefined; delta.font = fo; }
            if (changed.fill) delta.fill = fillColor || pattern ? (pattern ? { pattern, fg: M.rgb(patColor), bg: fillColor ? M.rgb(fillColor) : M.rgb('#FFFFFF') } : { pattern: 'solid', fg: M.rgb(fillColor), bg: { indexed: 64 } }) : null;
            if (changed.prot) delta.prot = { locked: pr.locked === false ? false : undefined, hidden: pr.hidden || undefined };
            if (Object.keys(delta).length) O.format(s, G().ranges(), (stl) => {
              const d2 = Object.assign({}, delta);
              if (d2.font) d2.font = Object.assign({}, d2.font);
              if (d2.align) { d2.align = Object.assign({}, d2.align); }
              void stl;
              return d2;
            }, '设置单元格格式');
            const spec = {};
            for (const k in touched) spec[k] = edges[k];
            if (Object.keys(spec).length) O.borders(s, G().ranges(), spec, '设置单元格格式');
            if (mergeState === true && !merged) O.merge(s, G().ranges(), {});
            if (mergeState === false && merged) O.unmerge(s, G().ranges());
            if (changed.align && (al.wrap || al.rot) || changed.font) for (const r of A().selRows()) O.autoRow(s, r);
          });
        });
      } }, { label: '取消' }],
    });
  };
  /** a small tile showing a fill pattern (solid: the pattern colour) */
  function patternURL(p, fg, bg) {
    if (!L.render.CELL_PATTERNS[p]) return `data:image/svg+xml,${encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" width="4" height="4"><rect width="4" height="4" fill="${fg}"/></svg>`)}`;
    return L.render.patternURL(p, fg, bg);
  }
  D.patternURL = patternURL;

  /* ================================================================ rows, columns, cells */
  const sizeDialog = (title, label, value, min, max, apply) => {
    const inp = h('input', { type: 'text', value: String(value), style: 'width:80px' });
    ui.dialog({ title, width: 260, body: ui.field(label, inp), buttons: [{ label: '确定', primary: true, onClick: () => { const v = num(inp.value, NaN); if (!(v >= min && v <= max)) { ui.msg(L.t('请输入 {0} 到 {1} 之间的数字。', min, max), { icon: 'warn' }); return false; } tryRun(() => apply(v)); } }, { label: '取消' }] });
  };
  D.rowHeight = function () {
    const s = sh(), rows = A().selRows();
    const pts = rows.map((r) => M.rowPt(s, r));
    sizeDialog('行高', '行高(&h)：', pts.every((p) => p === pts[0]) ? pts[0] : '', 0, 409, (v) => O.setRowHeight(s, rows, v));
  };
  D.colWidth = function () {
    const s = sh(), cols = A().selCols();
    const mdw = M.mdw(s.wb);
    const ch = cols.map((c) => M.pxToChars(M.colPx(s, c), mdw));
    sizeDialog('列宽', '列宽(&C)：', ch.every((x) => x === ch[0]) ? ch[0] : '', 0, 255, (v) => O.setColWidth(s, cols, v === 0 ? 0 : M.pxToWidth(M.charsToPx(v, mdw), mdw)));
  };
  D.standardWidth = function () {
    const s = sh(), mdw = M.mdw(s.wb);
    sizeDialog('标准列宽', '标准列宽(&S)：', M.pxToChars(M.defaultColPx(s), mdw), 0, 255, (v) => {
      const old = { d: s.defColW, b: s.baseColW };
      O.tx(s.wb, '标准列宽', () => { s.defColW = M.pxToWidth(M.charsToPx(v, mdw), mdw); s.baseColW = Math.round(v); LY.invalidate(s); s.wb.undo.op(() => { s.defColW = old.d; s.baseColW = old.b; LY.invalidate(s); }, () => { s.defColW = M.pxToWidth(M.charsToPx(v, mdw), mdw); s.baseColW = Math.round(v); LY.invalidate(s); }); });
    });
  };
  const shiftDialog = (title, items, apply, def) => {
    let pick = def || items[0][0];
    const radios = items.map(([k, label]) => ui.radio('shift', label, k === pick, () => { pick = k; }));
    ui.dialog({ title, width: 240, body: ui.group(title.replace(/\.\.\.$/, ''), ...radios), buttons: [{ label: '确定', primary: true, onClick: () => tryRun(() => apply(pick)) }, { label: '取消' }] });
  };
  D.insertCells = function () {
    if (G().guardProtect('insertRows')) return;
    shiftDialog(L.N('插入'), [['right', '右移单元格(&i)'], ['down', '下移单元格(&d)'], ['row', '整行(&r)'], ['col', '整列(&c)']], (k) => {
      if (k === 'row') A().insertLines('r'); else if (k === 'col') A().insertLines('c');
      else O.shiftCells(sh(), O.clip(sh(), G().range()), k, true);
    }, 'down');
  };
  D.deleteCells = function () {
    if (G().guardProtect('deleteRows')) return;
    shiftDialog(L.N('删除'), [['left', '左移单元格(&l)'], ['up', '上移单元格(&u)'], ['row', '整行(&r)'], ['col', '整列(&c)']], (k) => {
      if (k === 'row') A().deleteLines('r'); else if (k === 'col') A().deleteLines('c');
      else O.shiftCells(sh(), O.clip(sh(), G().range()), k, false);
    });
  };
  D.groupRowsCols = function (d) {
    return new Promise((res) => {
      let pick = 'r';
      ui.dialog({ title: d > 0 ? '组合' : '取消组合', width: 220, body: ui.group(d > 0 ? '组合' : '取消组合', ui.radio('grp', '行(&R)', true, () => { pick = 'r'; }), ui.radio('grp', '列(&C)', false, () => { pick = 'c'; })), buttons: [{ label: '确定', primary: true, onClick: () => res(pick) }, { label: '取消', onClick: () => res(null) }] }).done.then(() => res(null));
    });
  };

  /* ================================================================ sheets */
  D.insertSheetDlg = function () {
    const items = [['ws', L.N('工作表')], ['chart', L.N('图表')]].concat(Object.keys(L.panes.TEMPLATES).map((k) => ['tpl:' + k, L.panes.TEMPLATES[k].name]));
    let pick = 'ws';
    const lb = list(items, 'ws', (v) => { pick = v; });
    const d = ui.dialog({ title: '插入', width: 360, body: ui.tabs([{ label: '常规', body: h('div', { class: 'col' }, lb) }, { label: '电子表格解决方案', body: h('div', { class: 'tp-note', text: '模板列在“常规”选项卡中。' }) }]), buttons: [{ label: '确定', primary: true, onClick: () => go() }, { label: '取消' }] });
    lb.ondbl = () => { d.close(0); go(); };
    function go() {
      if (pick === 'ws') A().insertSheet();
      else if (pick === 'chart') L.chartWizard.open({ sheet: true });
      else {
        /* copy the template's sheet into this workbook */
        const t = L.panes.buildTemplate(pick.slice(4));
        const src = t.sheets[0];
        const w = wb();
        tryRun(() => {
          const s = O.addSheet(w, w.sheetByName(src.name) ? w.nextSheetName(src.name) : src.name, w.active);
          O.structural(s, L.N('插入'), () => {
            const remap = new Map();
            const sid = (i) => { if (!i) return i; if (!remap.has(i)) remap.set(i, w.styles.add(JSON.parse(JSON.stringify(t.styles.get(i))))); return remap.get(i); };
            src.rows.forEach((row, r) => { if (row) row.cells.forEach((cl, c) => { if (cl) { const n = Object.assign({}, cl, { s: sid(cl.s) }); if (n.f != null) n.dirty = true; s.put(r, c, n); } }); });
            s.cols = src.cols.map((x) => (x ? Object.assign({}, x) : x));
            s.merges = src.merges.slice(); s.dv = JSON.parse(JSON.stringify(src.dv)); s.view.freeze = src.view.freeze; s.view.grid = src.view.grid;
          });
          w.active = w.sheets.indexOf(s);
          A().renderTabs();
        });
      }
    }
  };
  D.moveCopySheet = function () {
    const w = wb();
    const sheets = A().selectedSheets();
    const books = A().books.map((b, i) => [i, b.name]);
    let book = A().cur, before = w.sheets.indexOf(sheets[0]) + 1, copy = false;
    const lb = h('div');
    const drawList = () => { lb.textContent = ''; const tw = A().books[book].wb; lb.appendChild(list(tw.sheets.map((s, i) => [i, s.name]).concat([[tw.sheets.length, L.t('（移至最后）')]]), Math.min(before, tw.sheets.length), (v) => { before = v; }, { translate: 'no' })); };
    drawList();
    ui.dialog({ title: '移动或复制', width: 320, body: h('div', { class: 'col' }, h('div', { text: '移动选定工作表' }), ui.field('目标工作簿(&T)：', ui.select(books, book, (v) => { book = +v; before = A().books[book].wb.sheets.length; drawList(); }, { translate: 'no' })), h('label', { text: '置于工作表之前(&B)：' }), lb, ui.check('建立副本(&C)', false, (v) => { copy = v; })),
      buttons: [{ label: '确定', primary: true, onClick: () => tryRun(() => {
        const tw = A().books[book].wb;
        if (tw === w) {
          let at = before;
          for (const s of sheets) { if (copy) { const c = O.copySheet(w, s, at); at = w.sheets.indexOf(c) + 1; w.active = w.sheets.indexOf(c); } else { O.moveSheet(w, s, at); at = w.sheets.indexOf(s) + 1; } }
        } else {
          /* another open workbook: copy the cells, styles and settings across */
          for (const s of sheets) {
            const ns = O.addSheet(tw, tw.sheetByName(s.name) ? tw.nextSheetName(s.name) : s.name, before);
            const remap = new Map();
            const sid = (i) => { if (!i) return i; if (!remap.has(i)) remap.set(i, tw.styles.add(JSON.parse(JSON.stringify(w.styles.get(i))))); return remap.get(i); };
            s.rows.forEach((row, r) => { if (row) { const nr = ns.rowObj(r); Object.assign(nr, row, { cells: [] }); if (row.s) nr.s = sid(row.s); row.cells.forEach((cl, c) => { if (cl) { const n = Object.assign({}, cl, { s: sid(cl.s), _gnode: null }); if (n.f != null) n.dirty = true; ns.put(r, c, n); } }); } });
            ns.cols = s.cols.map((x) => (x ? Object.assign({}, x, x.s ? { s: sid(x.s) } : {}) : x));
            ns.merges = s.merges.map((m) => Object.assign({}, m));
            ns.print = JSON.parse(JSON.stringify(s.print)); ns.view = JSON.parse(JSON.stringify(s.view)); ns.tabColor = s.tabColor;
            O.recalc(tw, null, true);
            if (!copy) O.deleteSheet(w, s);
          }
        }
        A().renderTabs();
      }) }, { label: '取消' }] });
  };
  D.unhideSheet = function () {
    const w = wb();
    const hidden = w.sheets.map((s, i) => [i, s]).filter(([, s]) => s.state === 'hidden');
    const lb = list(hidden.map(([i, s]) => [i, s.name]), hidden.length ? hidden[0][0] : null, null, { translate: 'no' });
    const d = ui.dialog({ title: '取消隐藏', width: 280, body: h('div', { class: 'col' }, h('label', { text: '取消隐藏工作表(&U)：' }), lb), buttons: [{ label: '确定', primary: true, onClick: () => go() }, { label: '取消' }] });
    lb.ondbl = () => { d.close(0); go(); };
    function go() { const i = lb.value(); if (i == null) return; tryRun(() => { O.sheetProp(w, w.sheets[i], 'state', 'visible', '取消隐藏工作表'); w.active = i; A().renderTabs(); }); }
  };
  D.tabColor = function () {
    const w = wb(), sheets = A().selectedSheets();
    let pick = sheets[0].tabColor ? M.colorHex(w, sheets[0].tabColor) : null;
    const grid = h('div', { class: 'cp-grid' });
    const none = h('button', { type: 'button', class: 'cp-wide', text: '无颜色' });
    const mark = () => { L.$$('.cp-sw', grid).forEach((b) => b.classList.toggle('on', b.dataset.c === pick)); none.classList.toggle('on', !pick); };
    none.addEventListener('click', () => { pick = null; mark(); });
    for (const c of L.color.STANDARD) { const b = h('button', { type: 'button', class: 'cp-sw', style: `background:${c}`, 'data-tip': ui.colorName(c), 'aria-label': ui.colorName(c) }); b.dataset.c = c; b.addEventListener('click', () => { pick = c; mark(); }); grid.appendChild(b); }
    mark();
    ui.dialog({ title: '设置工作表标签颜色', width: 260, body: h('div', { class: 'col' }, ui.group('工作表标签颜色', none, grid)), buttons: [{ label: '确定', primary: true, onClick: () => tryRun(() => { O.tx(w, '工作表标签颜色', () => { for (const s of sheets) O.sheetProp(w, s, 'tabColor', pick ? M.rgb(pick) : null); }); A().renderTabs(); }) }, { label: '取消' }] });
  };

  /* ================================================================ Find and Replace */
  let findState = { what: '', with: '', within: 'sheet', by: 'rows', lookIn: 'formulas', matchCase: false, whole: false, more: false };
  D.find = function (mode) {
    const w = wb();
    const what = h('input', { type: 'text', value: findState.what, style: 'width:100%' });
    const withIn = h('input', { type: 'text', value: findState.with, style: 'width:100%' });
    const withRow = ui.field('替换为(&e)：', withIn);
    const optBox = h('div', { class: 'fr-opts' },
      h('div', { class: 'col' }, ui.field('范围(&h)：', ui.select([['sheet', '工作表'], ['workbook', '工作簿']], findState.within, (v) => { findState.within = v; })), ui.field('搜索(&S)：', ui.select([['rows', '按行'], ['cols', '按列']], findState.by, (v) => { findState.by = v; })), ui.field('查找范围(&L)：', ui.select([['formulas', '公式'], ['数值', '数值'], ['comments', '批注']], findState.lookIn, (v) => { findState.lookIn = v; }))),
      h('div', { class: 'col' }, ui.check('区分大小写(&c)', findState.matchCase, (v) => { findState.matchCase = v; }), ui.check('单元格内容完全匹配(&o)', findState.whole, (v) => { findState.whole = v; })));
    optBox.hidden = !findState.more;
    const results = h('div', { class: 'fr-results', hidden: true });
    const status = h('div', { class: 'hint' });
    const moreBtn = ui.button(findState.more ? '选项 <<' : '选项 >>(&t)', () => { findState.more = !findState.more; optBox.hidden = !findState.more; moreBtn.innerHTML = ui.ampHTML(findState.more ? '选项 <<' : '选项 >>(&t)'); });
    const sync = () => { findState.what = what.value; findState.with = withIn.value; };
    const query = () => { sync(); const selRg = G().range(); const multi = G().ranges().length > 1 || selRg.r1 !== selRg.r2 || selRg.c1 !== selRg.c2; return Object.assign({}, findState, { sheet: sh(), range: findState.within === 'sheet' && multi ? selRg : null }); };
    const goTo = (hit) => { const i = w.sheets.indexOf(hit.sh); if (i !== w.active) A().activateSheet(i); G().select(hit.r, hit.c); };
    const findNext = (back) => {
      const q = query();
      if (!q.what) return null;
      const hits = O.findAll(w, q);
      if (!hits.length) { ui.msg("Ledger 找不到您要查找的数据。", { icon: 'info' }); return null; }
      const s = G().sel(), cur = { si: w.active, r: s.r, c: s.c };
      const key = (x) => [w.sheets.indexOf(x.sh), q.by === 'cols' ? x.c : x.r, q.by === 'cols' ? x.r : x.c];
      const ck = [cur.si, q.by === 'cols' ? cur.c : cur.r, q.by === 'cols' ? cur.r : cur.c];
      const cmp = (a, b) => a[0] - b[0] || a[1] - b[1] || a[2] - b[2];
      let hit = back ? hits.slice().reverse().find((x) => cmp(key(x), ck) < 0) || hits[hits.length - 1] : hits.find((x) => cmp(key(x), ck) > 0) || hits[0];
      goTo(hit);
      return hit;
    };
    const findAllBtn = () => {
      const q = query();
      const hits = O.findAll(w, q);
      results.hidden = false;
      results.textContent = '';
      const tbl = h('table', { class: 'fr-table' }, h('tr', null, ...[L.N('工作簿'), L.N('工作表'), L.N('名称'), L.N('单元格'), L.N('值'), L.N('公式')].map((t) => h('th', { text: t }))));
      for (const x of hits.slice(0, 2000)) {
        const cl = x.sh.get(x.r, x.c);
        const nm = w.names.find((n) => { try { const a = F.parseRange(String(n.ref).split('!').pop().replace(/\$/g, '')); return a && M.rangeContains(a, x.r, x.c); } catch (e) { return false; } });
        const tr = h('tr', null, h('td', { translate: 'no', text: A().books[A().cur].name }), h('td', { translate: 'no', text: x.sh.name }), h('td', { translate: 'no', text: nm ? nm.name : '' }), h('td', { translate: 'no', text: F.cellName(x.r, x.c, true, true) }), h('td', { translate: 'no', text: L.csv.cellText(x.sh, x.r, x.c) }), h('td', { translate: 'no', text: cl && cl.f != null ? '=' + F.display(cl.f) : '' }));
        tr.addEventListener('click', () => { L.$$('tr', tbl).forEach((t) => t.classList.remove('on')); tr.classList.add('on'); goTo(x); });
        tbl.appendChild(tr);
      }
      results.appendChild(tbl);
      status.textContent = L.t('找到 {0} 个单元格', hits.length);
    };
    let tabsEl, d = null;
    const findTab = h('div', { class: 'col' }), repTab = h('div', { class: 'col' });
    const form = h('div', { class: 'col' }, ui.field('查找内容(&n)：', what), withRow, h('div', { class: 'row', style: 'justify-content:flex-end' }, moreBtn), optBox);
    tabsEl = ui.tabs([{ label: '查找(&n)', body: findTab, onShow: () => { withRow.hidden = true; findTab.appendChild(form); setBtns(false); } }, { label: '替换(&p)', body: repTab, onShow: () => { withRow.hidden = false; repTab.appendChild(form); setBtns(true); } }], mode === 'replace' ? 1 : 0);
    d = ui.dialog({
      title: '查找和替换', width: 520, modeless: true, body: h('div', { class: 'col' }, tabsEl, results, status),
      buttons: [
        { label: '全部替换(&A)', onClick: () => { const q = query(); const hits = O.findAll(w, Object.assign({}, q, { lookIn: 'formulas' })); tryRun(() => { const n = O.replaceAll(w, q, hits); ui.msg(n === 1 ? L.t('Ledger 已完成搜索，共替换 {0} 处。', n) : L.t('Ledger 已完成搜索，共替换 {0} 处。', n), { icon: 'info' }); }); return false; } },
        { label: '替换(&R)', onClick: () => { const q = query(); const s = G().sel(); const here = O.findAll(w, Object.assign({}, q, { lookIn: 'formulas', within: 'sheet', range: { r1: s.r, c1: s.c, r2: s.r, c2: s.c } })); if (here.length) tryRun(() => O.replaceAll(w, q, here)); findNext(false); return false; } },
        { label: '查找全部(&i)', onClick: () => { findAllBtn(); return false; } },
        { label: '查找下一处(&F)', primary: true, onClick: () => { findNext(false); return false; } },
        { label: '关闭' },
      ],
    });
    function setBtns(rep) { if (!d) return; d.buttons[0].hidden = !rep; d.buttons[1].hidden = !rep; }
    setTimeout(() => { setBtns(mode === 'replace'); what.focus(); what.select(); }, 0);
    what.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); findNext(e.shiftKey); } });
  };

  /* ================================================================ Go To */
  D.goTo = function () {
    const w = wb();
    const names = w.names.filter((n) => !n.hidden && !/^_xlnm\./i.test(n.name) && (n.scope == null || n.scope === w.active)).map((n) => n.name).concat(w.tables.map((t) => t.name)).concat(A().lastGoTo || []);
    const ref = h('input', { type: 'text', style: 'width:100%' });
    const lb = list(Array.from(new Set(names)), null, (v) => { ref.value = v; }, { translate: 'no' });
    const d = ui.dialog({ title: '定位', width: 320, body: h('div', { class: 'col' }, h('label', { text: '定位：' }), lb, ui.field('引用(&R)：', ref)),
      buttons: [{ label: '特殊(&S)...', onClick: () => { d.close(null); D.goToSpecial(); return false; } }, { label: '确定', primary: true, onClick: () => { if (!A().goToRef(ref.value.trim())) { ui.msg('引用无效。', { icon: 'warn' }); return false; } } }, { label: '取消' }] });
    lb.ondbl = (v) => { d.close(0); A().goToRef(v); };
  };
  D.goToSpecial = function () {
    const st = { kind: 'constants', types: { n: true, t: true, l: true, e: true }, arr: 'row' };
    const sub = (k) => h('div', { class: 'gs-sub' }, ...[['n', L.N('数字')], ['t', L.N('文本')], ['l', L.N('逻辑值')], ['e', L.N('错误')]].map(([t, l]) => ui.check(l, true, (v) => { st.types[t] = v; })));
    const r = (k, label, extra) => h('div', null, ui.radio('gts', label, k === 'constants', () => { st.kind = k; }), extra || null);
    const body = h('div', { class: 'gs-grid' },
      h('div', { class: 'col' }, r('comments', '批注(&m)'), r('constants', '常量(&o)'), r('formulas', '公式(&F)', sub('f')), r('blanks', '空值(&B)'), r('region', '当前区域(&r)'), r('array', '当前数组(&a)'), r('objects', '对象(&b)')),
      h('div', { class: 'col' }, r('rowdiff', '行内容差异单元格(&w)'), r('coldiff', '列内容差异单元格(&u)'), r('prec', '引用单元格(&P)'), r('dep', '从属单元格(&D)'), r('last', '最后一个单元格(&s)'), r('visible', '仅可见单元格(&y)'), r('cf', '条件格式(&t)'), r('dv', '数据验证(&v)')));
    ui.dialog({ title: '定位条件', width: 400, body, buttons: [{ label: '确定', primary: true, onClick: () => {
      const s = sh(), sel = G().sel();
      const scope = (() => { const rg = G().range(); const single = G().ranges().length === 1 && rg.r1 === rg.r2 && rg.c1 === rg.c2; return single ? { r1: 0, c1: 0, r2: Math.max(0, s.maxR), c2: Math.max(0, s.maxC) } : O.clip(s, rg); })();
      const out = [];
      const add = (r, c) => out.push({ r1: r, c1: c, r2: r, c2: c });
      const typeOk = (v) => (typeof v === 'number' && st.types.n) || (typeof v === 'string' && st.types.t) || (typeof v === 'boolean' && st.types.l) || (M.isErr(v) && st.types.e);
      switch (st.kind) {
        case 'comments': for (const cm of s.comments.values()) if (M.rangeContains(scope, cm.r, cm.c)) add(cm.r, cm.c); break;
        case 'constants': s.each(scope.r1, scope.c1, scope.r2, scope.c2, (cl, rr, cc) => { if (cl.f == null && !cl.am && cl.v != null && cl.v !== '' && typeOk(cl.v)) add(rr, cc); }); break;
        case 'formulas': s.each(scope.r1, scope.c1, scope.r2, scope.c2, (cl, rr, cc) => { if ((cl.f != null || cl.am) && typeOk(cl.dirty ? C.cellValue(s, rr, cc) : cl.v)) add(rr, cc); }); break;
        case 'blanks': for (let rr = scope.r1; rr <= scope.r2; rr++) for (let cc = scope.c1; cc <= scope.c2 && out.length < 50000; cc++) { const cl = s.get(rr, cc); if (!cl || (cl.v == null && cl.f == null)) add(rr, cc); } break;
        case 'region': out.push(O.currentRegion(s, sel.r, sel.c)); break;
        case 'array': { const a = O.arrayAt(s, sel.r, sel.c); if (a) out.push(a); break; }
        case 'last': { const e = s.contentEnd(); add(e.r, e.c); break; }
        case 'visible': { const g = LY.geo(s); for (const rg of G().ranges()) { let start = null; const rr = O.clip(s, rg); for (let row = rr.r1; row <= rr.r2 + 1; row++) { const vis = row <= rr.r2 && g.rows.size(row) > 0; if (vis && start == null) start = row; if (!vis && start != null) { out.push({ r1: start, c1: rg.c1, r2: row - 1, c2: rg.c2 }); start = null; } } } break; }
        case 'cf': for (const cf of s.cf) out.push(...cf.ranges); break;
        case 'dv': for (const d of s.dv) out.push(...d.ranges); break;
        case 'objects': if (s.drawings.length) G().selectObject(s.drawings[0]); return;
        case 'prec': case 'dep': {
          const cell = s.get(sel.r, sel.c);
          if (st.kind === 'prec') { if (cell && cell.f != null) { const p = C.precedents(w, s, C.astOf(cell), sel.r, sel.c); for (const a of p.refs) if (a.sheet === s) out.push({ r1: a.r1, c1: a.c1, r2: a.r2, c2: a.c2 }); } }
          else { const deps = []; w.graph.dependents(s, sel.r, sel.c, deps); for (const n of deps) if (n.sh === s) add(n.r, n.c); }
          break;
        }
        case 'rowdiff': case 'coldiff': {
          const rg = O.clip(s, G().range());
          const txt = (rr, cc) => { const cl = s.get(rr, cc); return cl ? (cl.f != null ? F.display(cl.f) : String(cl.v)) : ''; };
          if (st.kind === 'rowdiff') for (let rr = rg.r1; rr <= rg.r2; rr++) { const base = s.get(rr, sel.c); for (let cc = rg.c1; cc <= rg.c2; cc++) { if (cc === sel.c) continue; const cl = s.get(rr, cc); if (base && base.f != null && cl && cl.f != null) { if (F.translate(F.display(base.f), 0, cc - sel.c) !== F.display(cl.f)) add(rr, cc); } else if (txt(rr, cc) !== txt(rr, sel.c)) add(rr, cc); } }
          else for (let cc = rg.c1; cc <= rg.c2; cc++) { const base = s.get(sel.r, cc); for (let rr = rg.r1; rr <= rg.r2; rr++) { if (rr === sel.r) continue; const cl = s.get(rr, cc); if (base && base.f != null && cl && cl.f != null) { if (F.translate(F.display(base.f), rr - sel.r, 0) !== F.display(cl.f)) add(rr, cc); } else if (txt(rr, cc) !== txt(sel.r, cc)) add(rr, cc); } }
          break;
        }
        default: break;
      }
      if (!out.length) { ui.msg('未找到单元格。', { icon: 'info' }); return false; }
      G().selectRange(out[0], { r: out[0].r1, c: out[0].c1 });
      const s2 = G().sel();
      s2.ranges = out.slice(0, 5000);
      s2.active = 0;
      G().changed();
    } }, { label: '取消' }] });
  };

  /* ================================================================ Paste Special */
  D.pasteSpecial = function () {
    const st = { what: 'all', op: 'none', skipBlanks: false, transpose: false };
    const what = [['all', '全部(&A)'], ['formulas', '公式(&F)'], ['values', '数值(&V)'], ['formats', '格式(&t)'], ['comments', '批注(&C)'], ['validation', '数据验证(&n)'], ['noBorders', '边框除外(&x)'], ['colWidths', '列宽(&w)'], ['formulasNF', '公式和数字格式(&o)'], ['valuesNF', '值和数字格式(&m)']];
    const ops = [['none', '无(&o)'], ['add', '添加(&d)'], ['sub', '减(&S)'], ['mul', '乘(&u)'], ['div', '除(&i)']];
    const half = (arr, name, key) => { const radios = arr.map(([k, l]) => ui.radio(name, l, k === st[key], () => { st[key] = k; })); const m = Math.ceil(radios.length / 2); return h('div', { class: 'ps-cols' }, h('div', { class: 'col' }, ...radios.slice(0, m)), h('div', { class: 'col' }, ...radios.slice(m))); };
    ui.dialog({ title: '选择性粘贴', width: 400, body: h('div', { class: 'col' }, ui.group('粘贴', half(what, 'ps-what', 'what')), ui.group('运算', half(ops, 'ps-op', 'op')), h('div', { class: 'row', style: 'gap:24px' }, ui.check('跳过空单元格(&b)', false, (v) => { st.skipBlanks = v; }), ui.check('转置(&e)', false, (v) => { st.transpose = v; }))),
      buttons: [{ label: '粘贴链接(&L)', onClick: () => tryRun(() => L.clip.pasteInternal({ link: true })) }, { label: '确定', primary: true, onClick: () => tryRun(() => L.clip.pasteInternal(st)) }, { label: '取消' }] });
  };

  /* ================================================================ Series, Fill Across */
  D.series = function () {
    const s = sh(), rg = O.clip(s, G().range());
    const st = { rows: rg.c2 - rg.c1 > rg.r2 - rg.r1, type: 'linear', unit: 'day', trend: false };
    const first = s.get(rg.r1, rg.c1);
    if (first && typeof first.v === 'number' && NF.isDate(LY.styleOf(s, rg.r1, rg.c1, first).nf || 'General')) st.type = 'date';
    const step = h('input', { type: 'text', value: '1', style: 'width:80px' }), stop = h('input', { type: 'text', value: '', style: 'width:80px' });
    const units = [['day', '日(&D)'], ['weekday', '工作日(&W)'], ['month', '月(&M)'], ['year', '年(&Y)']].map(([k, l]) => ui.radio('ser-u', l, k === 'day', () => { st.unit = k; }));
    ui.dialog({ title: '序列', width: 380, body: h('div', { class: 'col' },
      h('div', { class: 'row', style: 'gap:10px;align-items:flex-start' },
        ui.group('序列产生在', ui.radio('ser-in', '行(&R)', st.rows, () => { st.rows = true; }), ui.radio('ser-in', '列(&C)', !st.rows, () => { st.rows = false; })),
        ui.group('类型', ...[['linear', '等差序列(&L)'], ['growth', '等比序列(&G)'], ['date', '日期(&t)'], ['auto', '自动填充(&F)']].map(([k, l]) => ui.radio('ser-t', l, k === st.type, () => { st.type = k; }))),
        ui.group('日期单位', ...units)),
      ui.check('趋势(&r)', false, (v) => { st.trend = v; }),
      h('div', { class: 'row', style: 'gap:10px' }, ui.field('步长值(&S)：', step), ui.field('终止值(&o)：', stop))),
    buttons: [{ label: '确定', primary: true, onClick: () => tryRun(() => {
      if (st.type === 'auto') { const src = st.rows ? { r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c1 } : { r1: rg.r1, c1: rg.c1, r2: rg.r1, c2: rg.c2 }; O.autoFill(s, src, rg, 'auto'); return; }
      O.fillSeries(s, rg, { rows: st.rows, type: st.type, unit: st.unit, step: num(step.value, 1), stop: stop.value.trim() ? num(stop.value, null) : null, trend: st.trend });
    }) }, { label: '取消' }] });
  };
  D.fillAcross = function () {
    const sheets = A().selectedSheets();
    let what = 'all';
    ui.dialog({ title: '填充到其他工作表', width: 260, body: ui.group('填充', ...[['all', '全部(&A)'], ['contents', '内容(&C)'], ['formats', '格式(&F)']].map(([k, l], i) => ui.radio('fa', l, i === 0, () => { what = k; }))), buttons: [{ label: '确定', primary: true, onClick: () => tryRun(() => {
      const src = sh(), w = wb();
      O.tx(w, '填充到其他工作表', () => {
        for (const rg0 of G().ranges()) {
          const rg = O.clip(src, rg0);
          for (const t of sheets) {
            if (t === src) continue;
            for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) {
              const a = src.get(r, c), b = t.get(r, c);
              const n = Object.assign({}, b || { v: null });
              if (what !== 'formats') { n.v = a ? a.v : null; if (a && a.f != null) { n.f = a.f; n.dirty = true; } else delete n.f; }
              if (what !== 'contents') n.s = a ? a.s : undefined;
              O.put(t, r, c, n.v == null && n.f == null && !n.s ? null : n);
            }
          }
        }
      });
    }) }, { label: '取消' }] });
  };

  /* ================================================================ Save As */
  D.saveAs = function () {
    const b = A().books[A().cur];
    const base = b.name.replace(/\.[^.]+$/, '');
    const types = [['xlsx', L.N('Microsoft Office Excel 工作簿 (*.xlsx)')], ['xlsm', L.N('Excel 启用宏的工作簿 (*.xlsm)')], ['xltx', L.N('模板 (*.xltx)')], ['xltm', L.N('启用宏的模板 (*.xltm)')], ['xmlss', L.N('XML Spreadsheet 2003 (*.xml)')], ['csv', L.N('CSV（逗号分隔）(*.csv)')], ['txt', L.N('文本（制表符分隔）(*.txt)')], ['html', L.N('网页 (*.htm)')], ['pdf', L.N('PDF (*.pdf)')]];
    let type = b.type === 'xlsm' || b.type === 'xltx' || b.type === 'xltm' || b.type === 'csv' || b.type === 'txt' ? b.type : 'xlsx';
    const name = h('input', { type: 'text', value: base, style: 'width:100%' });
    const pw = { open: b.password || '', backup: false };
    /* the Tools button's menu, under the button */
    const tools = (button) => {
      const r = button.getBoundingClientRect();
      ui.openMenu([{ label: '常规选项(&G)...', run: () => {
        const p1 = h('input', { type: 'password', value: pw.open, autocomplete: 'new-password' });
        ui.dialog({ title: '保存选项', width: 320, body: h('div', { class: 'col' }, ui.group('文件共享', ui.field('打开文件的密码(&o)：', p1)), h('div', { class: 'hint', text: '使用密码保存的工作簿经过 AES-256 加密，可在 Excel 2010 及更高版本中打开。' })), buttons: [{ label: '确定', primary: true, onClick: () => {
          if (!p1.value) { pw.open = ''; return; }
          const p2 = h('input', { type: 'password', autocomplete: 'new-password' });
          ui.dialog({ title: '确认密码', width: 300, body: h('div', { class: 'col' }, ui.field('再次输入密码以继续。', p2), h('div', { class: 'hint', text: '注意：密码丢失或遗忘后将无法恢复。' })), buttons: [{ label: '确定', primary: true, onClick: () => { if (p2.value !== p1.value) { ui.msg("两次输入的密码不一致。", { icon: 'warn' }); return false; } pw.open = p1.value; } }, { label: '取消' }] });
        } }, { label: '取消' }] });
      } }], { left: r.left, bottom: r.bottom });
    };
    return ui.dialog({ title: '另存为', width: 460, body: h('div', { class: 'col' }, ui.field('文件名(&n)：', name), ui.field('保存类型(&t)：', ui.select(types, type, (v) => { type = v; })), h('div', { class: 'hint', text: '文件已保存到浏览器的下载位置。' })),
      buttons: [{ label: '工具', onClick: (close, button) => { tools(button); return false; } }, { label: '保存(&S)', primary: true, onClick: (close) => {
        const ext = { xmlss: 'xml', html: 'htm' }[type] || type;
        const n = (name.value.trim() || base).replace(/\.(xlsx|xlsm|xltx|xltm|xml|csv|txt|htm|html|pdf)$/i, '') + '.' + ext;
        b.password = pw.open || undefined;
        /* Embedded copy (tools-box): keep the dialog open until the workbook has been written.
           Upstream called A().saveAs() and closed the dialog at once (ui.dialog only treats a
           `false` return as "stay open", it does not await a promise), so `File ▸ Exit` carried
           on and navigated away while the .xlsx was still being produced — the download could be
           truncated and VO.saved() never ran, so a workbook saved from the exit prompt never
           reached Recent Files. Resolving after the save fixes both.
           ★ `close` must be a *parameter* of this onClick: ui.dialog hands it in (b.onClick(close, el)).
           Without it the bare name resolves to the browser's window.close, which silently does nothing
           here — the dialog stayed open, and the exit prompt that was waiting on it never went on to
           the Start Center. */
        Promise.resolve(A().saveAs(type, n)).then((ok) => close(ok ? 1 : 0), () => close(0));
        return false;
      } }, { label: '取消' }] }).done;
  };

  /* ================================================================ Properties */
  D.properties = function () {
    const w = wb(), p = w.props;
    const fld = (k, label) => { const i = h('input', { type: 'text', value: p[k] || '', style: 'width:100%' }); i.dataset.k = k; return ui.field(label, i); };
    const summary = h('div', { class: 'col props' }, fld('title', '标题(&T)：'), fld('subject', '主题(&S)：'), fld('creator', '作者(&A)：'), fld('manager', '经理(&M)：'), fld('company', '公司(&o)：'), fld('category', '类别(&t)：'), fld('keywords', '关键词(&K)：'), fld('description', '批注(&o)：'));
    const cells = w.sheets.reduce((n, s) => { let k = 0; s.rows.forEach((row) => { if (row) row.cells.forEach((cl) => { if (cl && (cl.v != null || cl.f != null)) k++; }); }); return n + k; }, 0);
    const formulas = w.sheets.reduce((n, s) => { let k = 0; s.rows.forEach((row) => { if (row) row.cells.forEach((cl) => { if (cl && cl.f != null) k++; }); }); return n + k; }, 0);
    const stats = h('table', { class: 'kv' }, ...[['创建时间：', p.created ? new Date(p.created).toLocaleString() : ''], ['修改时间：', p.modified ? new Date(p.modified).toLocaleString() : ''], ['上次保存者：', p.lastModifiedBy || ''], ['应用程序：', p.application || ''], ['工作表：', w.sheets.length], ['包含数据的单元格：', cells], ['公式：', formulas], ['定义的名称：', w.names.filter((n) => !n.hidden).length], ['列表（表格）：', w.tables.length]].map(([k, v]) => h('tr', null, h('td', { text: k }), h('td', { translate: 'no', text: String(v) }))));
    const contents = h('div', { class: 'col' }, h('b', { text: '工作表' }), ...w.sheets.map((s) => h('div', { translate: 'no', text: '   ' + (s.state !== 'visible' ? L.t('{0}（隐藏）', s.name) : s.name) })), h('b', { text: '命名区域' }), ...w.names.filter((n) => !n.hidden).map((n) => h('div', { translate: 'no', text: '   ' + n.name })));
    const custom = (p.custom || []).slice();
    const cbox = h('div');
    const drawCustom = () => { cbox.textContent = ''; cbox.appendChild(h('table', { class: 'kv' }, ...custom.map((c, i) => h('tr', null, h('td', { translate: 'no', text: c.name }), h('td', { translate: 'no', text: c.value }), h('td', null, ui.button('删除', () => { custom.splice(i, 1); drawCustom(); }, { class: 'btn small' })))))); };
    drawCustom();
    const cn = h('input', { type: 'text', placeholder: '名称' }), cv = h('input', { type: 'text', placeholder: '值' });
    const customTab = h('div', { class: 'col' }, h('div', { class: 'row', style: 'gap:6px' }, cn, cv, ui.button('添加(&A)', () => { if (!cn.value.trim()) return; custom.push({ name: cn.value.trim(), type: 'lpwstr', value: cv.value }); cn.value = cv.value = ''; drawCustom(); })), cbox);
    ui.dialog({ title: L.t('{0} 属性', A().books[A().cur].name), width: 420, body: ui.tabs([{ label: '摘要', body: summary }, { label: '统计信息', body: stats }, { label: '内容', body: contents }, { label: '自定义', body: customTab }]),
      buttons: [{ label: '确定', primary: true, onClick: () => { const old = Object.assign({}, p); const now = Object.assign({}, p); L.$$('input[data-k]', summary).forEach((i) => { now[i.dataset.k] = i.value; }); now.custom = custom; O.tx(w, L.N('属性'), () => { Object.assign(p, now); w.undo.op(() => Object.assign(p, old), () => Object.assign(p, now)); }); } }, { label: '取消' }] });
  };

  /* ================================================================ Zoom, Symbol */
  D.zoom = function () {
    const s = sh();
    let z = s.view.zoom || 100, fit = false;
    const custom = h('input', { type: 'text', value: String(z) + '%', style: 'width:60px' });
    const opts = [200, 100, 75, 50, 25].map((v) => ui.radio('zoom', v + '%', v === z, () => { z = v; fit = false; }));
    ui.dialog({ title: '缩放', width: 240, body: ui.group('缩放比例', ...opts, ui.radio('zoom', '缩放到选区(&F)', false, () => { fit = true; }), h('div', { class: 'row', style: 'gap:6px' }, ui.radio('zoom', '自定义(&C)：', ![200, 100, 75, 50, 25].includes(z), () => { fit = false; z = parseInt(custom.value, 10) || 100; }), custom)),
      buttons: [{ label: '确定', primary: true, onClick: () => { if (fit) { A().zoomToSelection(); return; } if (custom === document.activeElement || ![200, 100, 75, 50, 25].includes(z)) z = parseInt(custom.value, 10) || z; if (!(z >= 10 && z <= 400)) { ui.msg('请输入 10 到 400 之间的数字。', { icon: 'warn' }); return false; } G().setZoom(z); } }, { label: '取消' }] });
  };
  D.symbol = function () {
    const font = h('input', { type: 'text', value: 'Arial', style: 'width:140px' });
    const ranges = [['基本拉丁文', 0x20, 0x7e], ['拉丁文-1 补充', 0xa1, 0xff], ['拉丁文扩展-A', 0x100, 0x17f], [L.N('希腊文'), 0x391, 0x3c9], [L.N('西里尔文'), 0x410, 0x44f], ['常用标点', 0x2010, 0x205e], ['货币符号', 0x20a0, 0x20bf], ['类字母符号', 0x2100, 0x214f], ['数字形式', 0x2150, 0x218b], [L.N('箭头'), 0x2190, 0x21ff], ['数学运算符', 0x2200, 0x22ff], ['制表符号', 0x2500, 0x257f], ['几何形状', 0x25a0, 0x25ff], ['杂项符号', 0x2600, 0x26ff], [L.N('装饰符号'), 0x2700, 0x27bf]];
    let pick = null;
    const grid = h('div', { class: 'sym-grid', translate: 'no' });
    const info = h('div', { class: 'hint' });
    const draw = (i) => { grid.textContent = ''; const [, a, b] = ranges[i]; for (let cp = a; cp <= b; cp++) { const ch = String.fromCodePoint(cp); const btn = h('button', { type: 'button', translate: 'no', text: ch, 'aria-label': 'U+' + cp.toString(16).toUpperCase() }); btn.style.fontFamily = LY.fontStack(font.value); btn.addEventListener('click', () => { L.$$('button', grid).forEach((x) => x.classList.remove('on')); btn.classList.add('on'); pick = ch; info.textContent = L.t('字符代码：{0}（Unicode 十六进制）', cp.toString(16).toUpperCase().padStart(4, '0')); }); btn.addEventListener('dblclick', () => { pick = ch; insert(); }); grid.appendChild(btn); } };
    const sub = ui.select(ranges.map((r, i) => [i, r[0]]), 1, (v) => draw(+v));
    draw(1);
    const special = list([['—', L.N('— 长划线')], ['–', L.N('– 短划线')], ['©', L.N('© 版权')], ['®', L.N('® 注册商标')], ['™', L.N('™ 商标')], ['§', L.N('§ 节')], ['¶', L.N('¶ 段落')], ['…', L.N('… 省略号')], ['‘', L.N('‘ 左单引号')], ['’', L.N('’ 右单引号')], ['“', L.N('“ 左双引号')], ['”', L.N('” 右双引号')], [' ', '不间断空格']], null, (v) => { pick = v; });
    const insert = () => {
      if (!pick) return;
      if (L.editor.active) { L.editor.insertText(pick); return; }
      L.editor.begin('edit');
      L.editor.insertText(pick);
    };
    ui.dialog({ title: '符号', width: 520, body: ui.tabs([{ label: '符号(&S)', body: h('div', { class: 'col' }, h('div', { class: 'row', style: 'gap:10px' }, ui.field('字体(&F)：', font), ui.field('子集(&b)：', sub)), grid, info) }, { label: '特殊字符(&p)', body: special }]),
      buttons: [{ label: '插入(&I)', primary: true, onClick: () => { insert(); return false; } }, { label: '关闭' }] });
  };

  /* ================================================================ comments */
  /** Excel edits a comment in place: a yellow note beside the cell */
  D.comment = function (r, c, cm) {
    const s = sh(), w = wb();
    const q = G().cellRect(r, c);
    const wrapR = L.$('#gridwrap').getBoundingClientRect();
    const author = (A().opts.userName || 'Ledger 用户');
    const box = h('div', { class: 'cmt-edit' });
    const ta = h('textarea', { 'aria-label': '批注', spellcheck: 'true' });
    const initialText = cm ? L.threads.editText(cm) || '' : author + ':\n';
    ta.value = initialText;
    box.appendChild(ta);
    Object.assign(box.style, { left: (wrapR.left + q.x + q.w + 12) + 'px', top: Math.max(wrapR.top + 2, wrapR.top + q.y - 6) + 'px', width: (cm && cm.w ? cm.w * 96 / 72 : 150) + 'px', height: (cm && cm.h ? cm.h * 96 / 72 : 80) + 'px' });
    document.body.appendChild(box);
    ta.focus();
    if (!cm) ta.setSelectionRange(ta.value.length, ta.value.length);
    let done = false;
    const finish = (save) => {
      if (done) return; done = true;
      box.remove();
      document.removeEventListener('pointerdown', outside, true);
      if (save) {
        const text = ta.value;
        const same = cm && initialText === text;
        if (!same) tryRun(() => O.tx(w, cm ? '编辑批注' : '插入批注', () => O.setComment(s, r, c, L.threads.editNote(Object.assign({}, cm || { author, visible: false }, { text, runs: undefined, w: box.offsetWidth * 0.75 || undefined, h: box.offsetHeight * 0.75 || undefined }), text))));
      }
      G().paint(); G().syncObjects(); G().focus();
    };
    const outside = (e) => { if (!box.contains(e.target)) finish(true); };
    setTimeout(() => document.addEventListener('pointerdown', outside, true), 0);
    ta.addEventListener('keydown', (e) => { e.stopPropagation(); if (e.key === 'Escape') { e.preventDefault(); finish(true); } });
  };

  /* ================================================================ Options */
  D.options = function (tabName) {
    const o = A().opts, w = wb(), s = sh();
    const v = s.view;
    const ck = (label, get, set) => ui.check(label, get(), set);
    const viewTab = h('div', { class: 'col' },
      ui.group('显示', ck('启动任务窗格(&S)', () => o.taskOpen !== false, (x) => { o.taskOpen = x; }), ck('任务窗格跟随所选对象(&o)', () => o.paneFollows !== false, (x) => { o.paneFollows = x; }), ck('编辑栏(&F)', () => o.formulaBar !== false, (x) => { o.formulaBar = x; }), ck('状态栏(&a)', () => o.statusBar !== false, (x) => { o.statusBar = x; })),
      ui.group('批注', ui.radio('opt-cm', '无(&N)', !w._showComments, () => { w._showComments = false; }), ui.radio('opt-cm', '仅批注标识符(&i)', !w._showComments, () => { w._showComments = false; }), ui.radio('opt-cm', '批注和标识符(&c)', !!w._showComments, () => { w._showComments = true; })),
      ui.group('窗口选项', h('div', { class: 'opt-cols' },
        ck('分页符(&k)', () => !!s._pages, (x) => { if (x) L.print.computeBreaks(s); else s._pages = null; }), ck('公式(&r)', () => !!v.formulas, (x) => { v.formulas = x; }), ck('网格线(&G)', () => v.grid !== false, (x) => { v.grid = x; }), ck('行列标题(&e)', () => v.headings !== false, (x) => { v.headings = x; }),
        ck('零值(&Z)', () => v.zeros !== false, (x) => { v.zeros = x; }), ck('分级显示符号(&y)', () => v.outlineSymbols !== false, (x) => { v.outlineSymbols = x; }), ck('水平滚动条(&z)', () => !w.hideHScroll, (x) => { w.hideHScroll = !x; }), ck('垂直滚动条(&V)', () => !w.hideVScroll, (x) => { w.hideVScroll = !x; }), ck('工作表标签(&e)', () => !w.hideTabs, (x) => { w.hideTabs = !x; }))));
    const calc = w.calcPr;
    const iterN = h('input', { type: 'text', value: String(calc.iterateCount || 100), style: 'width:60px' }), iterD = h('input', { type: 'text', value: String(calc.iterateDelta || 0.001), style: 'width:60px' });
    const calcTab = h('div', { class: 'col' },
      ui.group('计算', ui.radio('opt-calc', '自动(&A)', calc.mode !== 'manual' && calc.mode !== 'autoNoTable', () => { calc.mode = 'auto'; }), ui.radio('opt-calc', '除模拟运算表外自动重算(&t)', calc.mode === 'autoNoTable', () => { calc.mode = 'autoNoTable'; }), ui.radio('opt-calc', '手动(&M)', calc.mode === 'manual', () => { calc.mode = 'manual'; }), ui.button('立即计算 (F9)(&N)', () => A().calculateNow())),
      ui.group('迭代', ck('迭代(&I)', () => !!calc.iterate, (x) => { calc.iterate = x; }), h('div', { class: 'row', style: 'gap:10px' }, ui.field('最多迭代次数(&n)：', iterN), ui.field('最大误差(&c)：', iterD))),
      ui.group('工作簿选项', ck('以显示精度为准(&P)', () => calc.fullPrecision === false, (x) => { calc.fullPrecision = !x; }), ck(L.N('1904 日期系统(&1)'), () => !!w.date1904, (x) => { w.date1904 = x; }), ck('接受公式中的标志(&b)', () => false, () => {})));
    const dir = ui.select([['down', '向下'], ['right', '右'], ['up', '向上'], ['left', '左']], o.enterDir || 'down', (x) => { o.enterDir = x; });
    const editTab = h('div', { class: 'col' },
      ck('直接在单元格内编辑(&e)', () => o.editInCell !== false, (x) => { o.editInCell = x; }), ck('允许拖放单元格(&d)', () => o.dragDrop !== false, (x) => { o.dragDrop = x; }),
      h('div', { class: 'row', style: 'gap:8px' }, ck('按 Enter 后移动选定单元格(&M)', () => o.moveAfterEnter !== false, (x) => { o.moveAfterEnter = x; }), ui.field('方向(&i)：', dir)),
      ck('为单元格值启用记忆式键入(&C)', () => o.autoComplete !== false, (x) => { o.autoComplete = x; }), ck('扩展数据区域的格式及公式(&f)', () => o.extendFormats !== false, (x) => { o.extendFormats = x; }), ck('启用自动输入百分比', () => o.autoPercent !== false, (x) => { o.autoPercent = x; }));
    const uname = h('input', { type: 'text', value: o.userName || '', style: 'width:200px' });
    const nSheets = ui.spin({ value: o.sheetsInNew || 3, min: 1, max: 255, step: 1, dec: 0 });
    const dfont = ui.select(L.FONT_LIST, o.defaultFont || 'Arial', null, { translate: 'no' }), dsize = ui.select(L.SIZE_LIST, o.defaultSize || 10);
    const look = ui.lookField();
    const genTab = h('div', { class: 'col' }, look && look.el, ck('R1C1 引用样式(&y)', () => !!w.r1c1, (x) => { w.r1c1 = x; }), ck('最近使用的文件列表(&R)', () => true, () => {}),
      ui.field('新工作簿内的工作表数(&b)：', nSheets), h('div', { class: 'row', style: 'gap:8px' }, ui.field('标准字体(&t)：', dfont), ui.field('大小(&z)：', dsize)), ui.field('用户名(&n)：', uname));
    const lists = O.customLists;
    const listBox = h('div');
    const listText = h('textarea', { rows: '8', style: 'width:100%', translate: 'no' });
    let curList = -1;
    const drawLists = () => { listBox.textContent = ''; listBox.appendChild(list([[-1, L.t('新列表')]].concat(lists.map((l, i) => [i, l.slice(0, 4).join(', ') + (l.length > 4 ? ', …' : '')])), curList, (i) => { curList = i; listText.value = i >= 0 ? lists[i].join('\n') : ''; listText.readOnly = i >= 0 && i < 4; }, { translate: 'no' })); };
    drawLists();
    const listsTab = h('div', { class: 'row', style: 'gap:10px;align-items:flex-start' }, h('div', { class: 'col', style: 'flex:1' }, h('label', { text: '自定义序列：' }), listBox), h('div', { class: 'col', style: 'flex:1' }, h('label', { text: '列表项：' }), listText),
      h('div', { class: 'col' }, ui.button('添加(&A)', () => { const items = listText.value.split(/[\n,]+/).map((x) => x.trim()).filter(Boolean); if (!items.length) return; if (curList >= 4) lists[curList] = items; else lists.push(items); L.store.set('customLists', lists.slice(4)); curList = -1; listText.value = ''; drawLists(); }), ui.button('删除(&D)', () => { if (curList >= 4) { lists.splice(curList, 1); L.store.set('customLists', lists.slice(4)); curList = -1; listText.value = ''; drawLists(); } })));
    const dictSel = ui.select([['en_US', '英语（美国）'], ['en_GB', '英语（英国）']], o.dictLang === 'en_GB' ? 'en_GB' : 'en_US', (v) => { o.dictLang = v; });
    const errTab = h('div', { class: 'col' }, ck('启用后台错误检查(&E)', () => o.bgErrors !== false, (x) => { o.bgErrors = x; }), ck('求值结果为错误值(&v)', () => true, () => {}), ck('以文本形式存储的数字(&t)', () => o.errNumText !== false, (x) => { o.errNumText = x; }), ck('区域内公式不一致(&I)', () => o.errInconsistent !== false, (x) => { o.errInconsistent = x; }));
    /* Security: the password to open this workbook (applied when it is next saved as .xlsx / .xlsm) */
    const book = A().books[A().cur];
    const pwOpen = h('input', { type: 'password', value: (book && book.password) || '', autocomplete: 'new-password', style: 'width:180px' });
    const securityTab = h('div', { class: 'col' },
      ui.group('此工作簿的文件加密设置', ui.field('打开文件的密码(&o)：', pwOpen)),
      h('div', { class: 'hint', text: '下次保存为 Excel 工作簿时，将采用 AES-256 加密，与 Excel 2007 及更高版本相同。将密码框留空可不设密码保存。密码丢失或遗忘后将无法恢复。' }),
      ui.group('隐私选项', ck('保存时从文件属性中删除个人信息(&R)', () => !!(w.extra && (w.extra.removePersonal || w.extra.filterPrivacy)), (x) => { w.extra = w.extra || {}; w.extra.removePersonal = x; w.extra.filterPrivacy = x; })));
    const tabs = [{ label: '视图', body: viewTab }, { label: '计算', body: calcTab }, { label: '编辑', body: editTab }, { label: '常规', body: genTab }, { label: '自定义序列', body: listsTab }, { label: '错误检查', body: errTab }, { label: '安全性', body: securityTab }];
    const idx = Math.max(0, tabs.findIndex((t) => t.label === tabName));
    ui.dialog({ title: '选项', width: 540, body: ui.tabs(tabs, idx), buttons: [{ label: '确定', primary: true, onClick: () => {
      calc.iterateCount = Math.max(1, Math.min(32767, parseInt(iterN.value, 10) || 100));
      calc.iterateDelta = num(iterD.value, 0.001);
      o.userName = uname.value; o.sheetsInNew = Math.round(nSheets.get()); o.defaultFont = dfont.value; o.defaultSize = +dsize.value;
      if (book && (pwOpen.value || '') !== (book.password || '')) { book.password = pwOpen.value || undefined; book.dirty = true; }
      A().applyOpts();
      if (look) look.apply();
      if (calc.mode !== 'manual') A().calculateNow();
      G().paint(true); G().syncObjects(true); A().renderTabs();
    } }, { label: '取消' }] });
  };
  D.autoCorrect = function () {
    const AC = L.autocorrect, o = A().opts;
    const ck = (label, key, def) => ui.check(label, o[key] !== undefined ? o[key] !== false : def, (x) => { o[key] = x; });
    const from = h('input', { type: 'text', style: 'width:120px' }), to = h('input', { type: 'text', style: 'width:220px' });
    const box = h('div');
    const draw = () => { box.textContent = ''; const all = Array.from(AC.list().entries()).sort((a, b) => a[0].localeCompare(b[0])); box.appendChild(list(all.map(([a, b]) => [a, a + '  →  ' + b]), null, (k) => { from.value = k; to.value = AC.list().get(k); }, { class: 'lbox ac-list', translate: 'no' })); };
    draw();
    ui.dialog({ title: '自动更正：英语（美国）', width: 460, body: ui.tabs([{ label: '自动更正', body: h('div', { class: 'col' },
      ck('更正连续两个大写字母', 'acTwoCaps', true), ck('句首字母大写(&s)', 'acSentence', false), ck('星期名称首字母大写(&n)', 'acDays', true), ck('更正意外使用大写锁定键造成的大小写错误(&L)', 'acCapsLock', true), ck('键入时自动替换(&t)', 'acReplace', true),
      h('div', { class: 'row', style: 'gap:8px' }, ui.field('替换(&R)：', from), ui.field('替换为(&W)：', to)), box,
      h('div', { class: 'row', style: 'gap:8px;justify-content:flex-end' }, ui.button('添加(&A)', () => { if (from.value) { AC.setEntry(from.value, to.value); draw(); } }), ui.button('删除(&D)', () => { if (from.value) { AC.setEntry(from.value, null); from.value = to.value = ''; draw(); } }))) }]),
    buttons: [{ label: '确定', primary: true, onClick: () => A().saveOpts() }, { label: '取消' }] });
  };
  D.autocorrect = D.autoCorrect;

  /* ================================================================ protection */
  /** Excel's legacy 16-bit password hash (sheetProtection/@password) */
  D.legacyHash = function (pw) {
    let hash = 0;
    const chars = Array.from(String(pw)).map((c) => c.charCodeAt(0) & 0x7fff);
    for (let i = chars.length - 1; i >= 0; i--) { hash = ((hash >> 14) & 1) | ((hash << 1) & 0x7fff); hash ^= chars[i]; }
    hash = ((hash >> 14) & 1) | ((hash << 1) & 0x7fff);
    hash ^= chars.length;
    hash ^= 0xce4b;
    return hash.toString(16).toUpperCase().padStart(4, '0');
  };
  /** check a password against stored protection attributes; true when it unlocks */
  async function checkPw(attrs, pw) {
    if (!attrs) return true;
    if (attrs.password) return D.legacyHash(pw) === String(attrs.password).toUpperCase();
    if (attrs.hashValue && attrs.algorithmName && root.crypto && root.crypto.subtle) {
      const alg = { 'SHA-512': 'SHA-512', 'SHA-256': 'SHA-256', 'SHA-384': 'SHA-384', 'SHA-1': 'SHA-1' }[attrs.algorithmName];
      if (!alg) return false;
      const b64 = (s) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));
      const salt = b64(attrs.saltValue || '');
      const enc = new Uint8Array(Array.from(String(pw)).flatMap((ch) => { const c = ch.charCodeAt(0); return [c & 255, c >> 8]; }));
      let hbuf = new Uint8Array(await crypto.subtle.digest(alg, new Uint8Array([...salt, ...enc])));
      const n = +attrs.spinCount || 0;
      for (let i = 0; i < n; i++) { const it = new Uint8Array(4); new DataView(it.buffer).setUint32(0, i, true); hbuf = new Uint8Array(await crypto.subtle.digest(alg, new Uint8Array([...hbuf, ...it]))); }
      return btoa(String.fromCharCode(...hbuf)) === attrs.hashValue;
    }
    return !attrs.password && !attrs.hashValue;
  }
  D.checkPw = checkPw;
  const askPw = (title) => new Promise((res) => {
    const p = h('input', { type: 'password', autocomplete: 'off' });
    let v = null;
    ui.dialog({ title, width: 280, body: ui.field('密码(&P)：', p), buttons: [{ label: '确定', primary: true, onClick: () => { v = p.value; } }, { label: '取消' }] }).done.then(() => res(v));
  });
  D.protectSheet = async function () {
    const s = sh(), w = wb();
    if (s.protection) {
      const a = s.protection.attrs || {};
      if (a.password || a.hashValue) {
        const pw = await askPw('撤消工作表保护');
        if (pw == null) return;
        if (!(await checkPw(a, pw))) { ui.msg('您输入的密码不正确。请检查 Caps Lock 是否已关闭，并确认字母大小写正确。', { icon: 'warn' }); return; }
      }
      tryRun(() => O.setProtection(s, null));
      return;
    }
    const allow = { selectLockedCells: true, selectUnlockedCells: true, formatCells: false, formatColumns: false, formatRows: false, insertColumns: false, insertRows: false, insertHyperlinks: false, deleteColumns: false, deleteRows: false, sort: false, autoFilter: false, pivotTables: false, objects: false, scenarios: false };
    const labels = { selectLockedCells: '选定锁定单元格', selectUnlockedCells: '选定未锁定的单元格', formatCells: '设置单元格格式', formatColumns: '设置列格式', formatRows: '设置行格式', insertColumns: '插入列', insertRows: '插入行', insertHyperlinks: '插入超链接', deleteColumns: '删除列', deleteRows: '删除行', sort: L.N('排序'), autoFilter: '使用自动筛选', pivotTables: '使用数据透视表', objects: '编辑对象', scenarios: '编辑方案' };
    const pw = h('input', { type: 'password', autocomplete: 'new-password', style: 'width:100%' });
    const box = h('div', { class: 'lbox checks' }, ...Object.keys(allow).map((k) => ui.check(labels[k], allow[k], (x) => { allow[k] = x; })));
    ui.dialog({ title: '保护工作表', width: 340, body: h('div', { class: 'col' }, ui.check('保护工作表及锁定单元格的内容(&P)', true, null), ui.field('撤消工作表保护的密码：', pw), h('label', { text: '允许此工作表的所有用户进行：' }), box),
      buttons: [{ label: '确定', primary: true, onClick: () => {
        const finish = () => {
          /* attributes in the file mean "this action is protected" (1), so allowed actions are written as 0 */
          const attrs = { sheet: '1', objects: allow.objects ? '0' : '1', scenarios: allow.scenarios ? '0' : '1' };
          for (const k of ['formatCells', 'formatColumns', 'formatRows', 'insertColumns', 'insertRows', 'insertHyperlinks', 'deleteColumns', 'deleteRows', 'sort', 'autoFilter', 'pivotTables']) attrs[k] = allow[k] ? '0' : '1';
          if (!allow.selectLockedCells) attrs.selectLockedCells = '1';
          if (!allow.selectUnlockedCells) attrs.selectUnlockedCells = '1';
          if (pw.value) attrs.password = D.legacyHash(pw.value);
          tryRun(() => O.setProtection(s, { attrs }));
        };
        if (pw.value) {
          const p2 = h('input', { type: 'password', autocomplete: 'new-password' });
          ui.dialog({ title: '确认密码', width: 300, body: h('div', { class: 'col' }, ui.field('再次输入密码以继续(&R)。', p2), h('div', { class: 'hint', text: '注意：密码丢失或遗忘后将无法恢复。' })), buttons: [{ label: '确定', primary: true, onClick: () => { if (p2.value !== pw.value) { ui.msg('两次输入的密码不一致。', { icon: 'warn' }); return false; } finish(); } }, { label: '取消' }] });
        } else finish();
        void w;
      } }, { label: '取消' }] });
  };
  D.protectWorkbook = async function () {
    const w = wb();
    if (w.protection) {
      const a = w.protection.attrs || {};
      const legacy = a.workbookPassword ? { password: a.workbookPassword } : a.workbookHashValue ? { hashValue: a.workbookHashValue, saltValue: a.workbookSaltValue, spinCount: a.workbookSpinCount, algorithmName: a.workbookAlgorithmName } : null;
      if (legacy) { const pw = await askPw('撤消工作簿保护'); if (pw == null) return; if (!(await checkPw(legacy, pw))) { ui.msg('您输入的密码不正确。', { icon: 'warn' }); return; } }
      const old = w.protection;
      O.tx(w, '撤消工作簿保护', () => { w.protection = null; w.undo.op(() => { w.protection = old; }, () => { w.protection = null; }); });
      ui.refresh();
      return;
    }
    let structure = true, windows = false;
    const pw = h('input', { type: 'password', autocomplete: 'new-password' });
    ui.dialog({ title: '保护工作簿', width: 300, body: h('div', { class: 'col' }, ui.group('保护工作簿的', ui.check('结构(&S)', true, (x) => { structure = x; }), ui.check('窗口(&W)', false, (x) => { windows = x; })), ui.field('密码（可选）(&P)：', pw)),
      buttons: [{ label: '确定', primary: true, onClick: () => { const attrs = {}; if (pw.value) attrs.workbookPassword = D.legacyHash(pw.value); const p = { attrs, structure, windows }; O.tx(w, '保护工作簿', () => { w.protection = p; w.undo.op(() => { w.protection = null; }, () => { w.protection = p; }); }); ui.refresh(); } }, { label: '取消' }] });
  };

  /* ================================================================ Hyperlink */
  D.hyperlink = function () {
    const s = sh(), w = wb(), sel = G().sel();
    if (G().guardProtect('insertHyperlinks')) return;
    const cur = s.links.find((l) => M.rangeContains(l.ref, sel.r, sel.c));
    const text = h('input', { type: 'text', value: cur && cur.display != null ? cur.display : L.csv.cellText(s, sel.r, sel.c), style: 'width:100%' });
    const addr = h('input', { type: 'text', value: cur && cur.target ? cur.target : '', style: 'width:100%', placeholder: 'https://', translate: 'no' });
    const cellRef = h('input', { type: 'text', value: cur && cur.location ? String(cur.location).split('!').pop() : 'A1', style: 'width:100%' });
    const sheetSel = ui.select(w.sheets.map((x) => [x.name, x.name]).concat(w.names.filter((n) => !n.hidden && !/^_xlnm/.test(n.name)).map((n) => ['name:' + n.name, L.t('{0}（定义的名称）', n.name)])), cur && cur.location && /!/.test(cur.location) ? cur.location.split('!')[0].replace(/^'|'$/g, '').replace(/''/g, "'") : s.name, null, { translate: 'no' });
    const mailTo = h('input', { type: 'text', value: cur && /^mailto:/i.test(cur.target || '') ? cur.target.replace(/^mailto:/i, '').split('?')[0] : '', style: 'width:100%' });
    const subj = h('input', { type: 'text', style: 'width:100%' });
    const tip = h('input', { type: 'text', value: cur && cur.tooltip || '', style: 'width:100%' });
    let kind = cur && cur.location ? 'place' : cur && /^mailto:/i.test(cur.target || '') ? 'mail' : 'web';
    const pages = { web: h('div', { class: 'col' }, ui.field('地址(&e)：', addr)), place: h('div', { class: 'col' }, ui.field('输入单元格引用(&n)：', cellRef), ui.field('或选择本文档中的位置：', sheetSel)), mail: h('div', { class: 'col' }, ui.field('电子邮件地址(&E)：', mailTo), ui.field('主题(&b)：', subj)) };
    const right = h('div', { class: 'col', style: 'flex:1' });
    const showKind = () => { right.textContent = ''; right.appendChild(pages[kind]); L.$$('.hl-kind', body).forEach((b) => b.classList.toggle('on', b.dataset.k === kind)); };
    const kbtn = (k, label) => { const b = h('button', { type: 'button', class: 'hl-kind', text: label }); b.dataset.k = k; b.addEventListener('click', () => { kind = k; showKind(); }); return b; };
    const body = h('div', { class: 'col' }, ui.field('要显示的文字(&T)：', text), h('div', { class: 'row', style: 'gap:10px;align-items:stretch' }, h('div', { class: 'hl-bar col' }, kbtn('web', '现有文件或网页'), kbtn('place', '本文档中的位置'), kbtn('mail', '电子邮件地址')), right), ui.field('屏幕提示(&p)：', tip));
    showKind();
    ui.dialog({ title: cur ? '编辑超链接' : '插入超链接', width: 520, body, buttons: [
      ...(cur ? [{ label: '删除链接(&R)', onClick: () => tryRun(() => O.clear(s, [{ r1: sel.r, c1: sel.c, r2: sel.r, c2: sel.c }], 'hyperlinks')) }] : []),
      { label: '确定', primary: true, onClick: () => tryRun(() => {
        const l = { ref: { r1: sel.r, c1: sel.c, r2: sel.r, c2: sel.c } };
        if (kind === 'web') { let t = addr.value.trim(); if (!t) return; if (!/^[a-z]+:/i.test(t) && /^www\.|\.[a-z]{2,}(\/|$)/i.test(t)) t = 'http://' + t; l.target = t; }
        else if (kind === 'mail') { if (!mailTo.value.trim()) return; l.target = 'mailto:' + mailTo.value.trim() + (subj.value ? '?subject=' + encodeURIComponent(subj.value) : ''); }
        else { const v = sheetSel.value; l.location = v.startsWith('name:') ? v.slice(5) : F.quoteSheet(v) + '!' + (cellRef.value.trim() || 'A1'); }
        const disp = text.value || l.target || l.location;
        l.display = disp;
        if (tip.value) l.tooltip = tip.value;
        O.tx(w, L.N('超链接'), () => {
          O.setLinks(s, s.links.filter((x) => !M.rangeContains(x.ref, sel.r, sel.c)).concat([l]));
          const cell = s.get(sel.r, sel.c);
          if (!cell || cell.f == null) {
            if (!cell || cell.v == null || String(cell.v) !== disp) O.setValue(s, sel.r, sel.c, disp, true);
            const c2 = s.get(sel.r, sel.c);
            const stl = w.cellStyles.find((x) => x.name === 'Hyperlink');
            const n = Object.assign({}, c2);
            n.s = w.styles.derive(c2.s || 0, stl ? { font: Object.assign({}, w.styles.get(stl.style).font) } : { font: { u: 'single', color: M.rgb('#0000FF') } });
            O.put(s, sel.r, sel.c, n);
          }
        });
      }) }, { label: '取消' }] });
  };

  /* ================================================================ misc */
  D.links = function () {
    const w = wb();
    const rows = (w.externalLinks || []).map((e, i) => h('tr', null, h('td', { translate: 'no', text: decodeURIComponent(String(e.relTarget || '').split(/[\\/]/).pop()) }), h('td', { text: '工作表' }), h('td', { text: 'A' }), h('td', { text: '此工作簿中缓存的值' })));
    ui.dialog({ title: '编辑链接', width: 520, body: h('div', { class: 'col' }, h('table', { class: 'fr-table' }, h('tr', null, h('th', { text: '来源' }), h('th', { text: '类型' }), h('th', { text: '更新' }), h('th', { text: '状态' })), ...rows), h('div', { class: 'hint', text: '引用其他工作簿的公式使用 Excel 随此文件保存的值。“断开链接”会将这些公式替换为其值。' })),
      buttons: [{ label: '断开链接(&B)', onClick: () => { tryRun(() => { O.tx(w, '断开链接', () => { for (const s of w.sheets) s.rows.forEach((row, r) => { if (row) row.cells.forEach((cl, c) => { if (cl && cl.f != null && /\[\d+\]/.test(cl.f)) O.setValue(s, r, c, cl.v, true); }); }); }); w.externalLinks = []; }); } }, { label: '关闭', primary: true }] });
  };
  D.customize = function () {
    const ids = [['standard', L.N('常用')], ['formatting', L.N('格式设置')], ['borders', L.N('边框')], ['chart', L.N('图表')], ['drawing', L.N('绘图')], ['auditing', '公式审核'], ['reviewing', L.N('审阅')], ['list', L.N('列表')]];
    const box = h('div', { class: 'lbox checks' }, ...ids.map(([k, l]) => ui.check(l, A().toolbarVisible(k), (x) => A().toggleToolbar(k, x))));
    const opts = h('div', { class: 'col' }, ui.check('分两行显示“常用”和“格式”工具栏(&t)', true, null), ui.check('在工具栏上显示屏幕提示(&T)', true, null), ui.check('在屏幕提示中显示快捷键(&k)', true, null));
    ui.dialog({ title: '自定义', width: 380, body: ui.tabs([{ label: '工具栏', body: box }, { label: '选项', body: opts }]), buttons: [{ label: '关闭', primary: true }] });
  };
  D.about = () => ui.about({ name: 'Ledger', office: 'Excel', text: [
    L.N('一个仿 Microsoft Office Excel 的电子表格程序，完全在浏览器中运行。工作簿以 Office Open XML（.xlsx、.xlsm）、CSV 和文本读写；不会上传任何内容。'),
    '计算引擎支持超过 450 个工作表函数、数组及动态数组公式、结构化引用和迭代计算。'] });
})(typeof window !== 'undefined' ? window : globalThis);
