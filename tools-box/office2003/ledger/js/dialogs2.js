/* Ledger — dialogs, part 2 (Data and Format): Sort, Validation, Subtotals, Text to Columns, Advanced Filter, Form,
 * Create List, Outline settings, Conditional Formatting, AutoFormat, Style, Goal Seek, Scenarios, Data Table,
 * Consolidate, PivotTable report. */
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
  const val = (s, r, c) => C.cellValue(s, r, c);
  const textOf = (s, r, c) => L.csv.cellText(s, r, c);
  /** a reference box: type a range or pick one on the sheet */
  function refBox(value, width) {
    const inp = h('input', { type: 'text', value: value || '', style: `width:${width || 200}px`, class: 'refbox' });
    const pick = h('button', { type: 'button', class: 'ref-pick', 'aria-label': '选择区域', 'data-tip': '使用当前选区', html: '<svg width="14" height="12" viewBox="0 0 14 12"><rect x=".5" y=".5" width="13" height="11" fill="#fff" stroke="#555"/><path d="M1 4h12M5 1v10" stroke="#999"/><rect x="5" y="4" width="8" height="7" fill="#c6d6f2"/></svg>' });
    pick.addEventListener('click', () => { const s = sh(); inp.value = F.quoteSheet(s.name) + '!' + F.absRangeName(G().range()); inp.dispatchEvent(new Event('change')); });
    const wrap = h('span', { class: 'refwrap' }, inp, pick);
    wrap.input = inp;
    return wrap;
  }
  D.refBox = refBox;
  /** "Sheet1!$A$1:$B$5" → {sheet, r1, c1, r2, c2} or null */
  function parseRef(text, single) {
    if (!text) return null;
    try {
      const ast = F.parse(String(text).trim().replace(/^=/, ''));
      const v = C.ev(ast, C.ctx(wb(), sh(), 0, 0));
      const a = v instanceof C.RefList ? v.list[0] : v;
      if (!(a instanceof C.Ref)) return null;
      if (single && (a.r1 !== a.r2 || a.c1 !== a.c2)) return null;
      return { sheet: a.sheet, r1: a.r1, c1: a.c1, r2: a.r2, c2: a.c2 };
    } catch (e) { return null; }
  }
  D.parseRef = parseRef;
  const listRange = () => {
    const s = sh(), sel = G().sel(), rg = G().range();
    const t = wb().tableAt(s, sel.r, sel.c);
    if (t) return { r1: t.ref.r1, c1: t.ref.c1, r2: t.ref.r2 - (t.totals ? 1 : 0), c2: t.ref.c2, table: t };
    const single = G().ranges().length === 1 && rg.r1 === rg.r2 && rg.c1 === rg.c2;
    return single ? O.currentRegion(s, sel.r, sel.c) : O.clip(s, rg);
  };
  const colLabel = (s, rg, header, c) => { if (header) { const v = s.val(rg.r1, c); if (v != null && v !== '') return String(v); } return L.t('列 {0}', F.colName(c)); };

  /* ================================================================ Sort */
  D.sort = function () {
    const s = sh(), sel = G().sel();
    const rg = listRange();
    if (rg.r1 === rg.r2 && rg.c1 === rg.c2 && s.get(rg.r1, rg.c1) == null) { ui.msg('排序前请先选择列表中的单元格。', { icon: 'warn' }); return; }
    let header = rg.table ? true : O.guessHeader(s, rg);
    const st = { keys: [{ c: Math.max(rg.c1, Math.min(rg.c2, sel.c)), desc: false }, { c: null, desc: false }, { c: null, desc: false }], caseSensitive: false, byCols: false, list: null };
    G().selectRange(rg, { r: sel.r, c: sel.c }, { noScroll: true });
    const body = h('div', { class: 'col' });
    const draw = () => {
      body.textContent = '';
      const opts = st.byCols ? Array.from({ length: rg.r2 - rg.r1 + 1 }, (_, i) => [rg.r1 + i, L.t('行 {0}', rg.r1 + i + 1)]) : Array.from({ length: rg.c2 - rg.c1 + 1 }, (_, i) => [rg.c1 + i, colLabel(s, rg, header, rg.c1 + i)]);
      st.keys.forEach((k, i) => {
        const selEl = ui.select((i ? [['', L.t('（无）')]] : []).concat(opts), k.c == null ? '' : k.c, (v) => { k.c = v === '' ? null : +v; }, { translate: 'no' });
        body.appendChild(ui.group(i === 0 ? '排序依据' : '次要关键字', h('div', { class: 'row', style: 'gap:12px;align-items:center' }, selEl, h('div', { class: 'col' }, ui.radio('sk' + i, '升序(&A)', !k.desc, () => { k.desc = false; }), ui.radio('sk' + i, '降序(&D)', k.desc, () => { k.desc = true; })))));
      });
      body.appendChild(ui.group('数据区域包含', h('div', { class: 'row', style: 'gap:16px' }, ui.radio('sk-h', '标题行(&r)', header, () => { header = true; draw(); }), ui.radio('sk-h', '无标题行(&w)', !header, () => { header = false; draw(); }))));
    };
    draw();
    const options = () => {
      const lists = [['', L.t('普通')]].concat(O.customLists.map((l, i) => [i, l.slice(0, 4).join(', ')]));
      ui.dialog({ title: '排序选项', width: 320, body: h('div', { class: 'col' }, ui.field('主关键字排序次序(&F)', ui.select(lists, st.list == null ? '' : st.list, (v) => { st.list = v === '' ? null : +v; }, { translate: 'no' })), ui.check('区分大小写(&C)', st.caseSensitive, (v) => { st.caseSensitive = v; }),
        ui.group('方向', ui.radio('so-o', '从上到下排序(&t)', !st.byCols, () => { st.byCols = false; st.keys.forEach((k) => { k.c = null; }); st.keys[0].c = sel.c; }), ui.radio('so-o', '从左到右排序(&l)', st.byCols, () => { st.byCols = true; st.keys.forEach((k) => { k.c = null; }); st.keys[0].c = sel.r; }))),
      buttons: [{ label: '确定', primary: true, onClick: () => draw() }, { label: '取消' }] });
    };
    ui.dialog({ title: '排序', width: 360, body, buttons: [{ label: '选项(&O)...', onClick: () => { options(); return false; } }, { label: '确定', primary: true, onClick: () => tryRun(() => {
      const keys = st.keys.filter((k) => k.c != null).map((k, i) => ({ index: st.byCols ? k.c - rg.r1 : k.c - rg.c1, desc: k.desc, list: i === 0 && st.list != null ? O.customLists[st.list] : undefined }));
      if (!keys.length) return;
      O.sort(s, { r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2 }, keys, { header, caseSensitive: st.caseSensitive, byCols: st.byCols });
    }) }, { label: '取消' }] });
  };

  /* ================================================================ Validation */
  D.validation = function () {
    const s = sh(), w = wb(), sel = G().sel();
    const cur = L.cf.dvAt(s, sel.r, sel.c);
    const ranges = G().ranges().map((r) => O.clip(s, r)).map((r, i) => (G().isWholeCols(G().ranges()[i]) || G().isWholeRows(G().ranges()[i]) ? G().ranges()[i] : r));
    if (cur && ranges.some((rg) => { for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 300); r++) for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 50); c++) if (L.cf.dvAt(s, r, c) !== cur) return true; return false; })) {
      /* mixed: Excel asks before erasing */
    }
    const st = Object.assign({ type: 'none', op: 'between', allowBlank: true, showDrop: true, showInput: true, showError: true, errorStyle: 'stop', f1: '', f2: '', prompt: '', promptTitle: '', error: '', errorTitle: '' }, cur ? JSON.parse(JSON.stringify(cur)) : {});
    delete st.ranges;
    const shows = (st.f1 || '').replace(/^"(.*)"$/, '$1');
    const allow = ui.select([['none', '任何值'], ['whole', '整数'], ['decimal', '小数'], ['list', '列表'], ['date', '日期'], ['time', '时间'], ['textLength', '文本长度'], ['custom', '自定义']], st.type, (v) => { st.type = v; draw(); });
    const dataSel = ui.select([['between', '介于'], ['notBetween', '不介于'], ['equal', '等于'], ['notEqual', '不等于'], ['greaterThan', '大于'], ['lessThan', '小于'], ['greaterThanOrEqual', '大于或等于'], ['lessThanOrEqual', '小于或等于']], st.op, (v) => { st.op = v; draw(); });
    const v1 = h('input', { type: 'text', style: 'width:100%' }), v2 = h('input', { type: 'text', style: 'width:100%' });
    const toInput = (f) => { if (f == null || f === '') return ''; if (/^"/.test(f)) return f.slice(1, -1).replace(/""/g, '"'); if (/^-?[\d.]+(E[+-]?\d+)?$/i.test(f)) { const n = +f; if (st.type === 'date') return NF.text('m/d/yyyy', n); if (st.type === 'time') return NF.text('h:mm:ss AM/PM', n); return String(n); } return '=' + F.display(f); };
    v1.value = st.type === 'list' && /^"/.test(st.f1 || '') ? shows : toInput(st.f1);
    v2.value = toInput(st.f2);
    const blank = ui.check('忽略空值(&b)', st.allowBlank, (v) => { st.allowBlank = v; });
    const drop = ui.check('提供下拉箭头(&I)', st.showDrop !== false, (v) => { st.showDrop = v; });
    const fields = h('div', { class: 'col' });
    function draw() {
      fields.textContent = '';
      const t = st.type;
      const ranged = ['whole', 'decimal', 'date', 'time', 'textLength'].includes(t);
      if (ranged) fields.appendChild(ui.field('数据(&D)：', dataSel));
      const two = ranged && (st.op === 'between' || st.op === 'notBetween');
      const lab1 = t === 'list' ? '来源(&S)：' : t === 'custom' ? '公式(&o)：' : two ? (t === 'date' ? '开始日期(&S)：' : t === 'time' ? '开始时间(&S)：' : '最小值(&M)：') : t === 'date' ? '日期(&D)：' : t === 'time' ? '时间(&T)：' : t === 'textLength' ? '长度(&L)：' : '值(&V)：';
      if (t !== 'none') fields.appendChild(ui.field(lab1, v1));
      if (two) fields.appendChild(ui.field(t === 'date' ? '结束日期(&E)：' : t === 'time' ? '结束时间(&E)：' : '最大值(&x)：', v2));
      blank.hidden = t === 'none';
      drop.hidden = t !== 'list';
    }
    draw();
    const settings = h('div', { class: 'col' }, ui.group('验证条件', h('div', { class: 'row', style: 'gap:10px;align-items:flex-start' }, h('div', { class: 'col', style: 'flex:1' }, ui.field('允许(&A)：', allow), fields), h('div', { class: 'col' }, blank, drop))),
      ui.check('将这些更改应用于所有设置相同的其他单元格', false, (v) => { st.applyAll = v; }));
    const pt = h('input', { type: 'text', value: st.promptTitle || '', style: 'width:100%' }), pm = h('textarea', { rows: '4', style: 'width:100%' });
    pm.value = st.prompt || '';
    const input = h('div', { class: 'col' }, ui.check('选定单元格时显示输入信息(&S)', st.showInput !== false, (v) => { st.showInput = v; }), h('div', { text: '选定单元格时显示以下输入信息：' }), ui.field('标题(&T)：', pt), ui.field('输入信息(&I)：', pm));
    const et = h('input', { type: 'text', value: st.errorTitle || '', style: 'width:100%' }), em = h('textarea', { rows: '4', style: 'width:100%' });
    em.value = st.error || '';
    const errSel = ui.select([['stop', '停止'], ['warning', '警告'], ['information', '信息']], st.errorStyle || 'stop', (v) => { st.errorStyle = v; });
    const error = h('div', { class: 'col' }, ui.check('输入无效数据时显示出错警告(&S)', st.showError !== false, (v) => { st.showError = v; }), h('div', { text: '用户输入无效数据时显示以下出错警告：' }), h('div', { class: 'row', style: 'gap:10px;align-items:flex-start' }, ui.field('样式(&y)：', errSel), h('div', { class: 'col', style: 'flex:1' }, ui.field('标题(&T)：', et), ui.field('错误信息(&E)：', em))));
    const ime = h('div', { class: 'col' }, ui.field('模式(&M)：', ui.select([['', '不控制'], ['on', '开启'], ['off', '关闭（英语模式）'], ['disabled', '禁用']], st.ime || '', (v) => { st.ime = v || undefined; })));
    const toFormula = (text) => {
      const t = String(text).trim();
      if (t === '') return null;
      if (t[0] === '=') { F.parse(t.slice(1)); return F.toStore(t.slice(1)); }
      if (st.type === 'list') return '"' + t.replace(/"/g, '""') + '"';
      const p = NF.parseInput(t, { date1904: w.date1904 });
      if (p && typeof p.v === 'number') return String(p.v);
      throw new Error('您输入的值不是有效的数字、日期或时间。');
    };
    ui.dialog({ title: '数据验证', width: 440, body: ui.tabs([{ label: '设置', body: settings }, { label: '输入信息', body: input }, { label: '出错警告', body: error }, { label: '输入法模式', body: ime }]),
      buttons: [{ label: '全部清除(&C)', onClick: () => { tryRun(() => O.tx(w, L.N('数据验证'), () => { for (const rg of ranges) O.removeRangeRules(s, rg, 'dv'); })); } },
        { label: '确定', primary: true, onClick: () => {
          let f1, f2;
          try {
            f1 = st.type === 'none' ? undefined : toFormula(v1.value);
            f2 = ['whole', 'decimal', 'date', 'time', 'textLength'].includes(st.type) && (st.op === 'between' || st.op === 'notBetween') ? toFormula(v2.value) : undefined;
            if (st.type !== 'none' && f1 == null) throw new Error(st.type === 'list' ? L.t('必须在“来源”框中输入值。') : L.t('必须在“最小值”框中输入值。'));
            if (f2 === null) throw new Error('必须在“最大值”框中输入值。');
          } catch (e) { ui.msg(e.message, { icon: 'warn' }); return false; }
          const rule = { type: st.type, op: ['whole', 'decimal', 'date', 'time', 'textLength'].includes(st.type) ? st.op : undefined, allowBlank: st.allowBlank, showDrop: st.showDrop !== false, showInput: st.showInput !== false, showError: st.showError !== false, errorStyle: st.errorStyle, f1: f1 == null ? undefined : f1, f2: f2 == null ? undefined : f2, promptTitle: pt.value || undefined, prompt: pm.value || undefined, errorTitle: et.value || undefined, error: em.value || undefined, ime: st.ime };
          if (st.x14) { rule.x14 = true; rule.extKeep = st.extKeep; }
          tryRun(() => O.tx(w, L.N('数据验证'), () => {
            let target = ranges;
            if (st.applyAll && cur) target = target.concat(cur.ranges);
            for (const rg of target) O.removeRangeRules(s, rg, 'dv');
            if (st.type !== 'none' || rule.prompt || rule.promptTitle) O.setDV(s, s.dv.concat([Object.assign({ ranges: target.map((x) => Object.assign({}, x)) }, rule)]));
          }));
        } }, { label: '取消' }] });
  };

  /* ================================================================ Subtotals */
  /* the functions' names, shown translated; the English goes into the subtotal rows */
  const SUBFN = [[9, L.N('求和')], [3, L.N('计数')], [1, L.N('平均值')], [4, L.N('最大值')], [5, L.N('最小值')], [6, L.N('乘积')], [2, '数值计数'], [7, L.N('样本标准偏差')], [8, L.N('总体标准偏差')], [10, L.N('样本方差')], [11, L.N('总体方差')]];
  D.subtotals = function () {
    const s = sh(), w = wb();
    const rg = listRange();
    if (rg.r2 <= rg.r1) { ui.msg('Ledger 无法确定列表或选区中哪一行包含此命令所需的列标签。', { icon: 'warn' }); return; }
    const cols = Array.from({ length: rg.c2 - rg.c1 + 1 }, (_, i) => rg.c1 + i);
    const st = { at: rg.c1, fn: 9, add: new Set([rg.c2]), replace: true, pageBreak: false, below: true };
    const addBox = h('div', { class: 'lbox checks', translate: 'no' }, ...cols.map((c) => ui.check(colLabel(s, rg, true, c), st.add.has(c), (v) => { if (v) st.add.add(c); else st.add.delete(c); })));
    const removeAll = () => tryRun(() => O.structural(s, '删除分类汇总', () => {
      for (let r = rg.r2; r > rg.r1; r--) { const row = s.rows[r]; if (row && row.cells.some((cl) => cl && cl.f != null && /^SUBTOTAL\(/i.test(cl.f))) O.insertLines(s, 'r', r, -1); }
      s.rows.forEach((row) => { if (row) { delete row.level; delete row.hidden; } });
      s.outline.levelRow = 0;
      LY.invalidate(s);
    }));
    ui.dialog({ title: '分类汇总', width: 340, body: h('div', { class: 'col' },
      ui.field('分类字段(&A)：', ui.select(cols.map((c) => [c, colLabel(s, rg, true, c)]), st.at, (v) => { st.at = +v; }, { translate: 'no' })),
      ui.field('使用函数(&U)：', ui.select(SUBFN, st.fn, (v) => { st.fn = +v; })),
      h('label', { text: '选定汇总项：' }), addBox,
      ui.check('替换当前分类汇总(&c)', true, (v) => { st.replace = v; }), ui.check('每组数据分页(&P)', false, (v) => { st.pageBreak = v; }), ui.check('汇总结果显示在数据下方(&S)', true, (v) => { st.below = v; })),
    buttons: [{ label: '全部移除(&R)', onClick: removeAll }, { label: '确定', primary: true, onClick: () => tryRun(() => {
      if (!st.add.size) return;
      const fnName = SUBFN.find((x) => x[0] === st.fn)[1];
      O.structural(s, L.N('分类汇总'), () => {
        let r2 = rg.r2;
        if (st.replace) { for (let r = r2; r > rg.r1; r--) { const row = s.rows[r]; if (row && row.cells.some((cl) => cl && cl.f != null && /^SUBTOTAL\(/i.test(cl.f))) { O.insertLines(s, 'r', r, -1); r2--; } } }
        /* groups: runs of equal values in the "at each change" column */
        const groups = [];
        let start = rg.r1 + 1;
        const key = (r) => textOf(s, r, st.at).toLowerCase();
        for (let r = rg.r1 + 2; r <= r2 + 1; r++) if (r > r2 || key(r) !== key(start)) { groups.push([start, r - 1]); start = r; }
        const bold = (r, c) => { const cl = s.get(r, c) || { v: null }; const n = Object.assign({}, cl); n.s = w.styles.derive(cl.s || 0, { font: { b: true } }); s.put(r, c, n); };
        /* insert from the bottom so earlier rows keep their numbers */
        const totalRows = [];
        for (let g = groups.length - 1; g >= 0; g--) {
          const [a, b] = groups[g];
          const at = st.below ? b + 1 : a;
          O.insertLines(s, 'r', at, 1);
          const ga = st.below ? a : a + 1, gb = st.below ? b : b + 1;
          s.put(at, st.at, { v: textOf(s, ga, st.at) + (st.fn === 9 ? ' Total' : ' ' + fnName) });
          bold(at, st.at);
          for (const c of st.add) { s.put(at, c, { f: `SUBTOTAL(${st.fn},${F.rangeName({ r1: ga, c1: c, r2: gb, c2: c })})`, v: null, dirty: true }); bold(at, c); }
          totalRows.push(at);
          for (let r = ga; r <= gb; r++) { const o = s.rowObj(r); o.level = 2; }
          s.rowObj(at).level = 1;
          if (st.pageBreak && g > 0) s.print.rowBreaks = (s.print.rowBreaks || []).concat([st.below ? at + 1 : at]).sort((x, y) => x - y);
        }
        /* grand total */
        const dataEnd = r2 + groups.length;
        const gt = st.below ? dataEnd + 1 : rg.r1 + 1;
        O.insertLines(s, 'r', gt, 1);
        const lastData = st.below ? dataEnd : dataEnd + 1;
        s.put(gt, st.at, { v: st.fn === 9 ? 'Grand Total' : 'Grand ' + fnName });
        bold(gt, st.at);
        for (const c of st.add) { s.put(gt, c, { f: `SUBTOTAL(${st.fn},${F.rangeName({ r1: st.below ? rg.r1 + 1 : rg.r1 + 2, c1: c, r2: lastData, c2: c })})`, v: null, dirty: true }); bold(gt, c); }
        s.outline.levelRow = 3;
        s.outline.below = st.below;
        if (s.maxR < gt) s.maxR = gt;
        LY.invalidate(s);
      });
    }) }, { label: '取消' }] });
  };

  /* ================================================================ Text to Columns */
  D.textToColumns = function () {
    const s = sh(), w = wb();
    const rg0 = O.clip(s, G().range());
    if (rg0.c1 !== rg0.c2) { ui.msg('“分列”一次只能转换一列。', { icon: 'warn' }); return; }
    const rows = [];
    for (let r = rg0.r1; r <= Math.min(rg0.r2, s.maxR); r++) { const v = s.val(r, rg0.c1); rows.push(v == null ? '' : typeof v === 'string' ? v : textOf(s, r, rg0.c1)); }
    const text = rows.join('\n');
    const st = { fixed: !/[\t,;]/.test(text) && / {2,}/.test(text), delims: { tab: /\t/.test(text), semi: false, comma: /,/.test(text) && !/\t/.test(text), space: false, other: '' }, merge: false, quote: '"', breaks: null, formats: [], dest: F.cellName(rg0.r1, rg0.c1, true, true) };
    if (!st.delims.tab && !st.delims.comma && !st.fixed) st.delims.space = true;
    const delimStr = () => (st.delims.tab ? '\t' : '') + (st.delims.semi ? ';' : '') + (st.delims.comma ? ',' : '') + (st.delims.space ? ' ' : '') + (st.delims.other ? st.delims.other[0] : '');
    const parse = () => (st.fixed ? L.csv.parseFixed(text, st.breaks || (st.breaks = L.csv.guessBreaks(text))) : rows.map((t) => (L.csv.parse(t, { delims: delimStr(), quote: st.quote, merge: st.merge })[0] || [''])));
    const preview = h('div', { class: 'tiw-preview' });
    let selCol = 0;
    const drawPrev = () => {
      preview.textContent = '';
      const data = parse().slice(0, 12);
      const n = Math.max(1, ...data.map((r) => r.length));
      const tbl = h('table', { class: 'tiw-table' });
      const hr = h('tr');
      for (let c = 0; c < n; c++) { const th = h('th', { class: c === selCol ? 'on' : '', text: { general: L.N('常规'), text: '文本', skip: L.N('跳过') }[st.formats[c] || 'general'] || st.formats[c] }); th.addEventListener('click', () => { selCol = c; drawPrev(); fmtSync(); }); hr.appendChild(th); }
      tbl.appendChild(hr);
      for (const r of data) { const tr = h('tr'); for (let c = 0; c < n; c++) tr.appendChild(h('td', { class: c === selCol ? 'on' : '', translate: 'no', text: r[c] == null ? '' : r[c] })); tbl.appendChild(tr); }
      preview.appendChild(tbl);
    };
    const d = st.delims;
    const delimBox = h('div', { class: 'tiw-delims' }, ui.check('制表符(&T)', d.tab, (v) => { d.tab = v; drawPrev(); }), ui.check('分号(&m)', d.semi, (v) => { d.semi = v; drawPrev(); }), ui.check('逗号(&C)', d.comma, (v) => { d.comma = v; drawPrev(); }), ui.check('空格(&S)', d.space, (v) => { d.space = v; drawPrev(); }), ui.field('其他(&O)：', (() => { const i = h('input', { type: 'text', maxlength: '1', style: 'width:24px' }); i.addEventListener('input', () => { d.other = i.value; drawPrev(); }); return i; })()));
    const breaksIn = h('input', { type: 'text', value: '', style: 'width:200px', placeholder: 'e.g. 10, 25, 40' });
    breaksIn.addEventListener('change', () => { st.breaks = breaksIn.value.split(/[,\s]+/).map(Number).filter((x) => x > 0).sort((a, b) => a - b); drawPrev(); });
    const fmtRadios = [['general', '常规(&G)'], ['text', '文本(&T)'], ['MDY', '日期（月日年）(&D)'], ['DMY', '日期（日月年）(&Y)'], ['YMD', '日期（年月日）(&Y)'], ['skip', '不导入此列（跳过）(&k)']].map(([k, l]) => ui.radio('ttc-f', l, k === 'general', () => { st.formats[selCol] = k; drawPrev(); }));
    const fmtSync = () => { const f = st.formats[selCol] || 'general'; ['general', 'text', 'MDY', 'DMY', 'YMD', 'skip'].forEach((k, i) => { fmtRadios[i].input.checked = f === k; }); };
    const dest = refBox(st.dest, 160);
    const typeBox = ui.group('原始数据类型', ui.radio('ttc-t', '分隔符号 - 用逗号或制表符等字符分隔各字段(&D)。', !st.fixed, () => { st.fixed = false; delimBox.hidden = false; brk.hidden = true; drawPrev(); }), ui.radio('ttc-t', '固定宽度 - 各字段按列对齐，字段之间以空格分隔(&w)。', st.fixed, () => { st.fixed = true; delimBox.hidden = true; brk.hidden = false; if (!st.breaks) st.breaks = L.csv.guessBreaks(text); breaksIn.value = st.breaks.join(', '); drawPrev(); }));
    const brk = ui.field('在以下字符处断列(&b)：', breaksIn);
    delimBox.hidden = st.fixed; brk.hidden = !st.fixed;
    if (st.fixed) { st.breaks = L.csv.guessBreaks(text); breaksIn.value = st.breaks.join(', '); }
    drawPrev();
    ui.dialog({ title: '文本分列向导', width: 600, body: h('div', { class: 'col' }, typeBox, delimBox, brk, h('div', { class: 'row', style: 'gap:12px' }, ui.check('将连续分隔符视为单个分隔符(&o)', false, (v) => { st.merge = v; drawPrev(); }), ui.field('文本识别符(&q)：', ui.select([['"', '"'], ["'", "'"], ['', '{无}']], '"', (v) => { st.quote = v; drawPrev(); }))), ui.group('列数据格式', h('div', { class: 'ttc-fmts' }, ...fmtRadios)), ui.field('目标(&e)：', dest), h('div', { class: 'tiw-cap', text: '数据预览：' }), preview),
      buttons: [{ label: '完成(&F)', primary: true, onClick: () => tryRun(() => {
        const to = parseRef(dest.input.value, false) || { sheet: s, r1: rg0.r1, c1: rg0.c1 };
        const data = parse();
        const dsh = to.sheet || s;
        const n = Math.max(...data.map((r) => r.length));
        /* overwriting other data needs a confirmation, as in Excel */
        let busy = false;
        for (let i = 0; i < data.length && !busy; i++) for (let c = 1; c < n; c++) { const cl = dsh.get(to.r1 + i, to.c1 + c); if (cl && cl.v != null && !(dsh === s && to.c1 + c === rg0.c1)) busy = true; }
        const go = () => O.tx(w, '分列', () => {
          data.forEach((row, i) => {
            let c = 0;
            row.forEach((t, j) => {
              const f = st.formats[j] || 'general';
              if (f === 'skip') return;
              const x = L.csv.convert(t, f, { date1904: w.date1904 });
              const old = dsh.get(to.r1 + i, to.c1 + c);
              const cl = Object.assign({}, old || {});
              delete cl.f; delete cl.rt;
              cl.v = x.v == null ? null : x.v;
              if (x.f != null) { cl.f = x.f; cl.dirty = true; }
              if (x.nf) cl.s = w.styles.derive(cl.s || 0, { nf: x.nf });
              O.put(dsh, to.r1 + i, to.c1 + c, cl.v == null && cl.f == null && !cl.s ? null : cl);
              c++;
            });
          });
        });
        if (busy) ui.msg('要替换目标单元格的内容吗？', { icon: 'question', buttons: ['确定', '取消'] }).then((r) => { if (r === 0) tryRun(go); });
        else go();
      }) }, { label: '取消' }] });
  };

  /* ================================================================ Advanced Filter */
  D.advancedFilter = function () {
    const s = sh();
    const rg = listRange();
    let copy = false, unique = false;
    const list = refBox(F.quoteSheet(s.name) + '!' + F.absRangeName(rg), 220);
    const crit = refBox('', 220), to = refBox('', 220);
    to.hidden = true;
    ui.dialog({ title: '高级筛选', width: 360, body: h('div', { class: 'col' }, ui.group('操作', ui.radio('af-a', '在原有区域显示筛选结果(&F)', true, () => { copy = false; to.input.disabled = true; }), ui.radio('af-a', '复制到其他位置(&o)', false, () => { copy = true; to.input.disabled = false; })),
      ui.field('列表区域(&L)：', list), ui.field('条件区域(&C)：', crit), ui.field('复制到(&t)：', (() => { to.hidden = false; to.input.disabled = true; return to; })()), ui.check('选择不重复的记录(&r)', false, (v) => { unique = v; })),
    buttons: [{ label: '确定', primary: true, onClick: () => {
      const lr = parseRef(list.input.value), cr = crit.input.value.trim() ? parseRef(crit.input.value) : null;
      if (!lr) { ui.msg('列表区域无效。', { icon: 'warn' }); return false; }
      if (crit.input.value.trim() && !cr) { ui.msg('条件区域无效。', { icon: 'warn' }); return false; }
      const dest = copy ? parseRef(to.input.value) : null;
      if (copy && !dest) { ui.msg('提取区域无效。', { icon: 'warn' }); return false; }
      tryRun(() => {
        const n = L.filter.advanced(lr.sheet, { r1: lr.r1, c1: lr.c1, r2: Math.min(lr.r2, lr.sheet.maxR), c2: lr.c2 }, cr ? cr.sheet : null, cr ? { r1: cr.r1, c1: cr.c1, r2: cr.r2, c2: cr.c2 } : null, { unique, copyTo: dest ? { sheet: dest.sheet, r: dest.r1, c: dest.c1 } : null });
        A().status(L.t('找到 {0} 条记录，共 {1} 条', n, lr.r2 - lr.r1));
      });
    } }, { label: '取消' }] });
  };

  /* ================================================================ Data Form */
  D.dataForm = function () {
    const s = sh(), w = wb();
    let rg = listRange();
    if (rg.r2 <= rg.r1 && !(s.get(rg.r1, rg.c1))) { ui.msg('未找到列表。请选择列表中的单元格。', { icon: 'warn' }); return; }
    const cols = Array.from({ length: rg.c2 - rg.c1 + 1 }, (_, i) => rg.c1 + i);
    let rec = 1, criteria = null;
    const inputs = cols.map((c) => { const i = h('input', { type: 'text', style: 'width:100%' }); i.dataset.c = c; return i; });
    const count = h('div', { class: 'df-count' });
    const fieldsBox = h('div', { class: 'df-fields', translate: 'no' }, ...cols.map((c, k) => ui.field(L.t('{0}:', colLabel(s, rg, true, c)), inputs[k])));
    const total = () => rg.r2 - rg.r1;
    const load = () => {
      if (criteria) return;
      const r = rg.r1 + rec;
      cols.forEach((c, k) => { const cl = s.get(r, c); inputs[k].value = rec > total() ? '' : cl && cl.f != null ? textOf(s, r, c) : cl && cl.v != null ? L.editor.cellText(s, r, c) : ''; inputs[k].readOnly = !!(cl && cl.f != null) && rec <= total(); });
      count.textContent = rec > total() ? L.t('新建记录') : L.t('{0}/{1}', rec, total());
    };
    const save = () => {
      if (criteria || rec > total() + 1) return;
      const r = rg.r1 + rec;
      const changes = inputs.filter((i, k) => !i.readOnly && i.value !== (s.get(r, cols[k]) && s.get(r, cols[k]).f == null ? L.editor.cellText(s, r, cols[k]) : textOf(s, r, cols[k])));
      if (!changes.length) return;
      if (rec > total() && !inputs.some((i) => i.value)) return;
      tryRun(() => O.tx(w, '记录单', () => { for (const i of changes) O.enter(s, r, +i.dataset.c, i.value); }));
      if (rec > total()) { rg = Object.assign({}, rg, { r2: rg.r2 + 1 }); if (rg.table) { const t = rg.table; O.tx(w, '记录单', () => { const old = t.ref; t.ref = Object.assign({}, t.ref, { r2: t.ref.r2 + 1 }); w.undo.op(() => { t.ref = old; }, () => { t.ref = Object.assign({}, old, { r2: old.r2 + 1 }); }); }); } }
    };
    const matches = (r) => cols.every((c, k) => { const q = criteria[k]; if (!q) return true; const t = textOf(s, r, c); const m = /^(<=|>=|<>|<|>|=)?(.*)$/.exec(q); const op = m[1] || '', x = m[2]; const n = num(x, null); const v = val(s, r, c); if (op && n != null && typeof v === 'number') return op === '<' ? v < n : op === '>' ? v > n : op === '<=' ? v <= n : op === '>=' ? v >= n : op === '<>' ? v !== n : v === n; return op === '<>' ? t.toLowerCase() !== x.toLowerCase() : t.toLowerCase().startsWith(x.toLowerCase()); });
    const find = (d) => { for (let k = rec + d; k >= 1 && k <= total(); k += d) if (!criteria || matches(rg.r1 + k)) return k; return null; };
    load();
    const btn = (label, fn) => ui.button(label, fn, { style: 'width:100%' });
    let critMode = null;
    const side = h('div', { class: 'col df-side' }, count,
      btn('新建(&w)', () => { save(); rec = total() + 1; load(); inputs[0].focus(); }),
      btn('删除(&D)', () => { if (rec > total()) return; ui.msg('当前显示的记录将被永久删除', { icon: 'warn', buttons: ['确定', '取消'] }).then((r) => { if (r !== 0) return; tryRun(() => O.shiftCells(s, { r1: rg.r1 + rec, c1: rg.c1, r2: rg.r1 + rec, c2: rg.c2 }, 'up', false)); rg = Object.assign({}, rg, { r2: rg.r2 - 1 }); rec = Math.min(rec, Math.max(1, total())); load(); }); }),
      btn('还原(&R)', () => load()),
      btn('查找上一处(&P)', () => { save(); const k = find(-1); if (k) { rec = k; load(); } }),
      btn('查找下一处(&N)', () => { save(); const k = find(1); if (k) { rec = k; load(); } }),
      critMode = btn('条件(&C)', () => {
        if (!criteria) { save(); criteria = inputs.map(() => ''); inputs.forEach((i) => { i.value = ''; i.readOnly = false; }); count.textContent = L.t('条件'); critMode.textContent = L.t('窗体'); }
        else { criteria = inputs.map((i) => i.value.trim()); const k = find(1) || find(-1) || rec; const crit = criteria; criteria = null; rec = k; load(); criteria = crit; criteria = crit.some(Boolean) ? crit : null; critMode.textContent = L.t('条件'); }
      }));
    inputs.forEach((i) => i.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); e.stopPropagation(); save(); if (rec <= total()) rec++; load(); } }));
    ui.dialog({ title: s.name, width: 460, body: h('div', { class: 'row df', style: 'gap:12px;align-items:flex-start' }, fieldsBox, side), buttons: [{ label: '关闭(&o)', primary: true, onClick: () => save() }] });
  };

  /* ================================================================ Create List */
  D.createList = function () {
    const s = sh(), w = wb();
    if (A().tableAt()) { ui.msg('列表不能与其他列表重叠。', { icon: 'warn' }); return; }
    const rg = listRange();
    const ref = refBox(F.absRangeName(rg), 200);
    let header = O.guessHeader(s, rg) || rg.r2 > rg.r1 && typeof s.val(rg.r1, rg.c1) === 'string';
    ui.dialog({ title: '创建列表', width: 320, body: h('div', { class: 'col' }, h('div', { text: '列表的数据在哪里？' }), ref, ui.check('列表有标题(&M)', header, (v) => { header = v; })),
      buttons: [{ label: '确定', primary: true, onClick: () => {
        const p = parseRef(ref.input.value);
        if (!p) { ui.msg('引用无效。', { icon: 'warn' }); return false; }
        tryRun(() => {
          const t = D.makeTable(p.sheet, { r1: p.r1, c1: p.c1, r2: Math.min(p.r2, Math.max(p.r1 + 1, p.sheet.maxR)), c2: p.c2 }, header);
          G().selectRange(t.ref, { r: t.ref.r1 + 1, c: t.ref.c1 });
        });
      } }, { label: '取消' }] });
  };
  /** turn a range into a list (table); adds a header row when the data has none */
  D.makeTable = function (s, rg, header, styleName) {
    const w = s.wb;
    let t;
    O.tx(w, '创建列表', () => {
      if (!header) { O.insertLines(s, 'r', rg.r1, 1); rg = Object.assign({}, rg, { r2: rg.r2 + 1 }); for (let c = rg.c1; c <= rg.c2; c++) O.put(s, rg.r1, c, { v: 'Column' + (c - rg.c1 + 1) }); }
      let n = 1; while (w.findTable('Table' + n) || w.findName('Table' + n)) n++;
      const id = Math.max(0, ...w.tables.map((x) => x.id || 0)) + 1;
      const used = new Set();
      const columns = [];
      for (let c = rg.c1; c <= rg.c2; c++) { let name = String(s.val(rg.r1, c) == null || s.val(rg.r1, c) === '' ? 'Column' + (c - rg.c1 + 1) : textOf(s, rg.r1, c)); let k = 2; const base = name; while (used.has(name.toLowerCase())) name = base + k++; used.add(name.toLowerCase()); columns.push({ name, id: c - rg.c1 + 1 }); if (s.val(rg.r1, c) !== name) O.put(s, rg.r1, c, Object.assign({}, s.get(rg.r1, c) || {}, { v: name, f: undefined })); }
      t = { id, name: 'Table' + n, sheet: s, ref: Object.assign({}, rg), header: true, totals: false, columns, style: { name: styleName || 'TableStyleMedium2', rowStripes: true, colStripes: false, first: false, last: false }, autoFilter: true, filter: { cols: [] } };
      w.tables.push(t);
      w.undo.op(() => { w.tables = w.tables.filter((x) => x !== t); }, () => { if (!w.tables.includes(t)) w.tables.push(t); });
      O.markStructural();
    });
    return t;
  };

  /* ================================================================ Outline settings */
  D.outlineSettings = function () {
    const s = sh();
    const o = Object.assign({}, s.outline);
    ui.dialog({ title: '设置', width: 300, body: h('div', { class: 'col' }, ui.group('方向', ui.check('汇总行位于明细数据下方(&b)', o.below !== false, (v) => { o.below = v; }), ui.check('汇总列位于明细数据右侧(&g)', o.right !== false, (v) => { o.right = v; })), ui.check('自动样式(&A)', !!o.applyStyles, (v) => { o.applyStyles = v; })),
      buttons: [{ label: '创建(&C)', onClick: () => A().autoOutline() }, { label: '确定', primary: true, onClick: () => tryRun(() => { const old = s.outline; O.tx(s.wb, '分级显示设置', () => { s.outline = Object.assign({}, s.outline, o); s.wb.undo.op(() => { s.outline = old; }, () => { s.outline = Object.assign({}, old, o); }); }); }) }, { label: '取消' }] });
  };

  /* ================================================================ Conditional Formatting (Excel 2003) */
  const CF_OPS = [['between', L.N('介于')], ['notBetween', L.N('不介于')], ['equal', L.N('等于')], ['notEqual', L.N('不等于')], ['greaterThan', L.N('大于')], ['lessThan', L.N('小于')], ['greaterThanOrEqual', L.N('大于或等于')], ['lessThanOrEqual', L.N('小于或等于')]];
  D.conditionalFormat = function () {
    const s = sh(), w = wb(), sel = G().sel();
    const ranges = G().ranges().map((r) => O.clip(s, r));
    const top = ranges[0];
    const editable = (r) => r.type === 'cellIs' || r.type === 'expression';
    /* the conditions that apply to the active cell */
    const existing = [];
    for (const cf of s.cf) if (cf.ranges.some((rg) => M.rangeContains(rg, sel.r, sel.c))) for (const r of cf.rules) if (editable(r)) existing.push({ rule: r, anchor: cf.ranges[0] });
    existing.sort((a, b) => (a.rule.priority || 0) - (b.rule.priority || 0));
    const fromStored = (f, anchor) => (f == null ? '' : /^"/.test(f) ? f.slice(1, -1).replace(/""/g, '"') : /^-?[\d.]+(E[+-]?\d+)?$/i.test(f) ? f : '=' + F.translate(F.display(f), sel.r - anchor.r1, sel.c - anchor.c1));
    const conds = existing.slice(0, 3).map(({ rule, anchor }) => ({ keep: JSON.parse(JSON.stringify(rule)), kind: rule.type === 'expression' ? 'formula' : 'value', op: rule.op || 'between', v1: fromStored(rule.f && rule.f[0], anchor), v2: fromStored(rule.f && rule.f[1], anchor), dxf: rule.dxf != null && w.dxfs[rule.dxf] ? JSON.parse(JSON.stringify(w.dxfs[rule.dxf])) : {}, stop: rule.stop }));
    if (!conds.length) conds.push({ kind: 'value', op: 'between', v1: '', v2: '', dxf: {} });
    const body = h('div', { class: 'col' });
    const prevCss = (d) => { const f = d.font || {}; const css = []; if (f.b) css.push('font-weight:bold'); if (f.i) css.push('font-style:italic'); if (f.u || f.strike) css.push('text-decoration:' + (f.u ? 'underline ' : '') + (f.strike ? 'line-through' : '')); if (f.color) css.push('color:' + M.colorHex(w, f.color)); if (d.fill && (d.fill.fg || d.fill.bg)) css.push('background:' + M.colorHex(w, d.fill.fg || d.fill.bg)); if (d.border) css.push('outline:1px solid #000;outline-offset:-3px'); return css.join(';'); };
    function draw() {
      body.textContent = '';
      conds.forEach((cd, i) => {
        const kindSel = ui.select([['value', '单元格数值'], ['formula', '公式']], cd.kind, (v) => { cd.kind = v; draw(); });
        const opSel = ui.select(CF_OPS, cd.op, (v) => { cd.op = v; draw(); });
        const a = h('input', { type: 'text', value: cd.v1, style: 'width:' + (cd.kind === 'formula' ? '300px' : cd.op === 'between' || cd.op === 'notBetween' ? '110px' : '240px') });
        a.addEventListener('input', () => { cd.v1 = a.value; });
        const b = h('input', { type: 'text', value: cd.v2, style: 'width:110px' });
        b.addEventListener('input', () => { cd.v2 = b.value; });
        const two = cd.kind === 'value' && (cd.op === 'between' || cd.op === 'notBetween');
        const prev = h('div', { class: 'cf-prev', style: prevCss(cd.dxf), text: Object.keys(cd.dxf).length ? 'AaBbCcYyZz' : '未设置格式' });
        const fmtBtn = ui.button('格式(&F)...', () => D.cfFormat(cd.dxf, (d2) => { cd.dxf = d2; draw(); }));
        body.appendChild(ui.group(L.t('条件 {0}', i + 1), h('div', { class: 'row', style: 'gap:6px;align-items:center;flex-wrap:wrap' }, kindSel, cd.kind === 'value' ? opSel : null, a, two ? h('span', { text: '与' }) : null, two ? b : null),
          h('div', { class: 'row', style: 'gap:8px;align-items:center;margin-top:4px' }, h('span', { class: 'hint', text: '条件成立时所用格式的预览：' }), prev, fmtBtn)));
      });
    }
    draw();
    const toStore = (t) => {
      t = String(t).trim();
      if (t === '') return null;
      if (t[0] === '=') { const rel = F.translate(t.slice(1), top.r1 - sel.r, top.c1 - sel.c); F.parse(rel); return F.toStore(rel); }
      const p = NF.parseInput(t, { date1904: w.date1904 });
      if (p && typeof p.v === 'number') return String(p.v);
      if (p && typeof p.v === 'boolean') return p.v ? 'TRUE' : 'FALSE';
      return '"' + t.replace(/"/g, '""') + '"';
    };
    ui.dialog({ title: '条件格式', width: 560, body,
      buttons: [{ label: '添加 >>(&A)', onClick: () => { if (conds.length < 3) { conds.push({ kind: 'value', op: 'between', v1: '', v2: '', dxf: {} }); draw(); } return false; } },
        { label: '删除(&D)...', onClick: () => { const checks = conds.map((c, i) => ui.check(L.t('条件 {0}', i + 1), false, null)); ui.dialog({ title: '删除条件格式', width: 260, body: ui.group('选择要删除的条件：', ...checks), buttons: [{ label: '确定', primary: true, onClick: () => { for (let i = checks.length - 1; i >= 0; i--) if (checks[i].input.checked) conds.splice(i, 1); if (!conds.length) conds.push({ kind: 'value', op: 'between', v1: '', v2: '', dxf: {} }); draw(); } }, { label: '取消' }] }); return false; } },
        { label: '确定', primary: true, onClick: () => {
          const rules = [];
          try {
            for (const cd of conds) {
              if (cd.kind === 'formula') { if (!String(cd.v1).trim()) continue; const f = toStore(cd.v1[0] === '=' ? cd.v1 : '=' + cd.v1); rules.push({ ...cd.keep, type: 'expression', op: undefined, f: [f], dxf: cd.dxf }); }
              else { const f1 = toStore(cd.v1); if (f1 == null) continue; const two = cd.op === 'between' || cd.op === 'notBetween'; const f2 = two ? toStore(cd.v2) : null; if (two && f2 == null) throw new Error('对于“介于”条件，两个框都必须输入值。'); rules.push({ ...cd.keep, type: 'cellIs', op: cd.op, f: two ? [f1, f2] : [f1], dxf: cd.dxf }); }
            }
          } catch (e) { ui.msg(e.message || '公式包含错误。', { icon: 'warn' }); return false; }
          tryRun(() => O.tx(w, '条件格式', () => {
            for (const rg of ranges) O.removeRangeRules(s, rg, 'cf', editable);
            if (!rules.length) return;
            let prio = Math.max(0, ...s.cf.flatMap((x) => x.rules.map((r) => r.priority || 0)));
            const out = rules.map((r) => { const dx = w.dxfs.push(r.dxf) - 1; return { ...r, dxf: dx, priority: ++prio }; });
            /* Excel 2003 conditions come before newer rules: renumber so these are first */
            const others = s.cf.flatMap((x) => x.rules).sort((a, b) => (a.priority || 0) - (b.priority || 0));
            out.forEach((r, i) => { r.priority = i + 1; });
            const shifted = s.cf.map((x) => Object.assign({}, x, { rules: x.rules.map((r) => Object.assign({}, r, { priority: out.length + 1 + others.indexOf(r) })) }));
            O.setCF(s, shifted.concat([{ ranges: ranges.map((r) => Object.assign({}, r)), rules: out }]));
          }));
        } }, { label: '取消' }] });
  };
  /** the reduced Format Cells used by conditional formats: font style / colour, borders, fill */
  D.cfFormat = function (dxf, done) {
    const w = wb();
    const d = JSON.parse(JSON.stringify(dxf || {}));
    const f = d.font || {};
    const styleSel = ui.select([['', '（不变）'], ['r', '常规'], ['i', '倾斜'], ['b', '加粗'], ['bi', '加粗、倾斜']], f.b == null && f.i == null ? '' : (f.b ? 'b' : '') + (f.i ? 'i' : '') || 'r', (v) => { if (!v) { delete f.b; delete f.i; } else { f.b = /b/.test(v); f.i = /i/.test(v); } });
    const uSel = ui.select([['', '（不变）'], ['none', '无'], ['single', '单线'], ['double', '双线']], f.u || '', (v) => { if (!v) delete f.u; else f.u = v === 'none' ? undefined : v; });
    const col = (get, set, mode) => { const b = h('button', { type: 'button', class: 'swatch-btn wide', style: `background:${get() || 'transparent'}` }, h('span', { text: get() ? '' : 'Automatic' }), h('span', { class: 'dd-arrow' })); b.addEventListener('click', () => ui.colorMenu(b, { mode, grid: true }, (c) => { if (c.auto || c.none) { set(null); b.style.background = 'transparent'; b.firstChild.textContent = c.none ? L.t('无颜色') : L.t('自动'); } else if (typeof c === 'string') { set(c); b.style.background = c; b.firstChild.textContent = ''; } })); return b; };
    const fontTab = h('div', { class: 'col' }, ui.field('字形(&y)：', styleSel), ui.field('下划线(&U)：', uSel), ui.field('颜色(&C)：', col(() => (f.color ? M.colorHex(w, f.color) : null), (c) => { if (c) f.color = M.rgb(c); else delete f.color; }, 'font')), ui.check('删除线(&k)', !!f.strike, (v) => { f.strike = v || undefined; }));
    const fill = d.fill || {};
    const patTab = h('div', { class: 'col' }, ui.field('单元格底纹颜色：', col(() => (fill.fg || fill.bg ? M.colorHex(w, fill.fg || fill.bg) : null), (c) => { if (c) { fill.pattern = 'solid'; fill.fg = M.rgb(c); fill.bg = M.rgb(c); } else { delete fill.pattern; delete fill.fg; delete fill.bg; } }, 'fill')));
    const bd = d.border || {};
    const bTab = h('div', { class: 'col' }, ui.check('外边框(&O)', !!(bd.t && bd.b && bd.l && bd.r), (v) => { for (const k of ['t', 'b', 'l', 'r']) if (v) bd[k] = { style: 'thin' }; else delete bd[k]; }), ...[['t', L.N('顶端')], ['b', L.N('底端')], ['l', L.N('左')], ['r', L.N('右')]].map(([k, l]) => ui.check(l, !!bd[k], (v) => { if (v) bd[k] = { style: 'thin' }; else delete bd[k]; })));
    ui.dialog({ title: '设置单元格格式', width: 340, body: ui.tabs([{ label: '字体', body: fontTab }, { label: '边框', body: bTab }, { label: '图案', body: patTab }]),
      buttons: [{ label: '清除(&C)', onClick: () => { done({}); } }, { label: '确定', primary: true, onClick: () => { const out = {}; const fo = {}; for (const k in f) if (f[k] !== undefined) fo[k] = f[k]; if (Object.keys(fo).length) out.font = fo; if (fill.pattern) out.fill = fill; if (Object.keys(bd).length) out.border = bd; done(out); } }, { label: '取消' }] });
  };

  /* ================================================================ AutoFormat */
  const DARK = M.rgb('#000080'), MID = M.rgb('#C0C0C0'), PALE = M.rgb('#FFFFCC');
  const AUTOFMT = {
    [L.N('简洁')]: { head: { font: { b: true }, border: { t: { style: 'thin' }, b: { style: 'thin' } } }, total: { font: { b: true }, border: { t: { style: 'thin' }, b: { style: 'thin' } } }, body: {} },
    [L.N('经典 1')]: { head: { font: { i: true }, border: { b: { style: 'thin' } } }, total: { border: { t: { style: 'thin' }, b: { style: 'double' } } }, first: { font: { b: true } }, body: {} },
    [L.N('经典 2')]: { head: { font: { b: true, color: M.rgb('#FFFFFF') }, fill: { pattern: 'solid', fg: M.rgb('#800080') } }, first: { fill: { pattern: 'solid', fg: MID }, font: { b: true } }, total: { font: { b: true }, border: { t: { style: 'thin' } } }, body: {} },
    [L.N('经典 3')]: { head: { font: { b: true, i: true, color: M.rgb('#FFFFFF') }, fill: { pattern: 'solid', fg: DARK } }, body: { fill: { pattern: 'solid', fg: MID } }, total: { font: { b: true }, fill: { pattern: 'solid', fg: MID }, border: { t: { style: 'thin' } } } },
    [L.N('会计专用 1')]: { head: { font: { b: true, i: true }, border: { b: { style: 'thin' } } }, body: { nf: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)' }, total: { font: { b: true }, border: { t: { style: 'thin' }, b: { style: 'double' } }, nf: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)' } },
    [L.N('会计专用 2')]: { head: { font: { b: true }, border: { t: { style: 'medium' }, b: { style: 'thin' } } }, body: { nf: '#,##0.00_);(#,##0.00)' }, total: { font: { b: true }, border: { t: { style: 'thin' }, b: { style: 'medium' } }, nf: '"$"#,##0.00_);\\("$"#,##0.00\\)' } },
    [L.N('会计专用 3')]: { head: { font: { b: true, i: true }, border: { b: { style: 'double' } } }, first: { font: { b: true } }, body: { nf: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)' }, total: { font: { b: true }, border: { t: { style: 'thin' } }, nf: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)' } },
    [L.N('会计专用 4')]: { head: { font: { b: true }, border: { b: { style: 'thin' } } }, body: { nf: '#,##0.00_);(#,##0.00)' }, total: { font: { b: true }, border: { t: { style: 'thin' }, b: { style: 'double' } }, nf: '"$"#,##0.00_);\\("$"#,##0.00\\)' } },
    [L.N('彩色 1')]: { head: { font: { b: true, color: M.rgb('#FFFFFF') }, fill: { pattern: 'solid', fg: M.rgb('#008080') } }, first: { font: { b: true, color: M.rgb('#FFFFFF') }, fill: { pattern: 'solid', fg: M.rgb('#008080') } }, body: { fill: { pattern: 'solid', fg: M.rgb('#CCFFFF') } }, total: { font: { b: true }, fill: { pattern: 'solid', fg: M.rgb('#99CCFF') } } },
    [L.N('彩色 2')]: { head: { font: { b: true, color: M.rgb('#FFFFFF') }, fill: { pattern: 'solid', fg: M.rgb('#993366') } }, body: { fill: { pattern: 'solid', fg: PALE } }, total: { font: { b: true }, fill: { pattern: 'solid', fg: M.rgb('#FFCC99') } } },
    [L.N('彩色 3')]: { head: { font: { b: true, color: M.rgb('#FFFFFF') }, fill: { pattern: 'solid', fg: M.rgb('#000000') } }, first: { font: { b: true, color: M.rgb('#FFFFFF') }, fill: { pattern: 'solid', fg: M.rgb('#000000') } }, body: { fill: { pattern: 'solid', fg: M.rgb('#FFCC00') } }, total: { font: { b: true, color: M.rgb('#FFFFFF') }, fill: { pattern: 'solid', fg: M.rgb('#000000') } } },
    [L.N('列表 1')]: { head: { font: { b: true }, border: { b: { style: 'medium' } } }, band: { fill: { pattern: 'solid', fg: M.rgb('#E3E3E3') } }, body: {}, total: { font: { b: true }, border: { t: { style: 'thin' } } } },
    [L.N('列表 2')]: { head: { font: { b: true, color: M.rgb('#FFFFFF') }, fill: { pattern: 'solid', fg: M.rgb('#333399') } }, band: { fill: { pattern: 'solid', fg: M.rgb('#CCCCFF') } }, body: {}, total: { font: { b: true }, border: { t: { style: 'double' } } } },
    [L.N('列表 3')]: { head: { font: { b: true, i: true }, border: { t: { style: 'thin' }, b: { style: 'thin' } } }, band: { fill: { pattern: 'solid', fg: M.rgb('#CCFFCC') } }, body: {}, total: { font: { b: true }, border: { t: { style: 'thin' }, b: { style: 'thin' } } } },
    [L.N('三维效果 1')]: { head: { font: { b: true }, fill: { pattern: 'solid', fg: MID }, border: { t: { style: 'thin', color: M.rgb('#FFFFFF') }, l: { style: 'thin', color: M.rgb('#FFFFFF') }, b: { style: 'thin', color: M.rgb('#808080') }, r: { style: 'thin', color: M.rgb('#808080') } } }, body: { fill: { pattern: 'solid', fg: MID }, border: { t: { style: 'thin', color: M.rgb('#FFFFFF') }, l: { style: 'thin', color: M.rgb('#FFFFFF') }, b: { style: 'thin', color: M.rgb('#808080') }, r: { style: 'thin', color: M.rgb('#808080') } } }, total: { font: { b: true }, fill: { pattern: 'solid', fg: MID } } },
    [L.N('三维效果 2')]: { head: { font: { b: true }, fill: { pattern: 'solid', fg: MID }, border: { t: { style: 'thin', color: M.rgb('#808080') }, l: { style: 'thin', color: M.rgb('#808080') }, b: { style: 'thin', color: M.rgb('#FFFFFF') }, r: { style: 'thin', color: M.rgb('#FFFFFF') } } }, body: { fill: { pattern: 'solid', fg: MID }, border: { t: { style: 'thin', color: M.rgb('#808080') }, l: { style: 'thin', color: M.rgb('#808080') }, b: { style: 'thin', color: M.rgb('#FFFFFF') }, r: { style: 'thin', color: M.rgb('#FFFFFF') } } }, total: { font: { b: true }, fill: { pattern: 'solid', fg: MID } } },
    [L.N('无')]: null,
  };
  D.autoFormat = function () {
    const s = sh(), w = wb();
    const rg = listRange();
    if (rg.r1 === rg.r2 && rg.c1 === rg.c2) { ui.msg('无法在活动单元格周围检测到表格。\n\n请选择要自动套用格式的区域。', { icon: 'warn' }); return; }
    let pick = 'Simple';
    const parts = { nf: true, border: true, font: true, fill: true, align: true, size: true };
    const grid = h('div', { class: 'afmt-grid' });
    const sample = (name) => {
      const spec = AUTOFMT[name];
      const tbl = h('table', { class: 'afmt-sample', translate: 'no' });
      const data = [['', 'Jan', 'Feb', 'Mar', 'Total'], ['East', 7, 7, 5, 19], ['West', 6, 4, 7, 17], ['South', 8, 7, 9, 24], ['Total', 21, 18, 21, 60]];
      data.forEach((row, r) => { const tr = h('tr'); row.forEach((x, c) => { const role = !spec ? {} : r === 0 ? spec.head : r === data.length - 1 ? spec.total : c === 0 && spec.first ? spec.first : spec.band && r % 2 === 0 ? Object.assign({}, spec.body, spec.band) : spec.body; const td = h('td', { translate: 'no', text: String(x) }); const css = []; const f = role.font || {}; if (f.b) css.push('font-weight:bold'); if (f.i) css.push('font-style:italic'); if (f.color) css.push('color:' + M.colorHex(w, f.color)); if (role.fill) css.push('background:' + M.colorHex(w, role.fill.fg)); const b = role.border || {}; for (const [k, side] of [['t', 'top'], ['b', 'bottom'], ['l', 'left'], ['r', 'right']]) if (b[k]) css.push(`border-${side}:${b[k].style === 'medium' ? 2 : b[k].style === 'double' ? 3 : 1}px ${b[k].style === 'double' ? 'double' : 'solid'} ${M.colorHex(w, b[k].color, '#000')}`); td.style.cssText = css.join(';'); if (typeof x === 'number') td.style.textAlign = 'right'; tr.appendChild(td); }); tbl.appendChild(tr); });
      return tbl;
    };
    for (const name of Object.keys(AUTOFMT)) { const b = h('button', { type: 'button', class: 'afmt-item' + (name === pick ? ' on' : '') }, sample(name), h('div', { text: name })); b.addEventListener('click', () => { pick = name; L.$$('.afmt-item', grid).forEach((x) => x.classList.toggle('on', x === b)); }); grid.appendChild(b); }
    const optBox = h('div', { class: 'row', style: 'gap:10px;flex-wrap:wrap' }, ...[['nf', '数字(&N)'], ['border', '边框(&B)'], ['font', '字体(&F)'], ['fill', '图案(&P)'], ['align', '对齐方式(&A)'], ['size', '宽度/高度(&W)']].map(([k, l]) => ui.check(l, true, (v) => { parts[k] = v; })));
    optBox.hidden = true;
    ui.dialog({ title: '自动套用格式', width: 560, body: h('div', { class: 'col' }, grid, optBox), buttons: [{ label: '选项(&O)...', onClick: () => { optBox.hidden = !optBox.hidden; return false; } }, { label: '确定', primary: true, onClick: () => tryRun(() => {
      const spec = AUTOFMT[pick];
      O.tx(w, L.N('自动套用格式'), () => {
        for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) {
          const cl = s.get(r, c) || { v: null };
          const isTotal = r === rg.r2 && rg.r2 > rg.r1 + 1 && /total/i.test(String(s.val(r, rg.c1) || ''));
          let role = !spec ? null : r === rg.r1 ? spec.head : isTotal ? spec.total : c === rg.c1 && spec.first ? spec.first : spec.band && (r - rg.r1) % 2 === 0 ? Object.assign({}, spec.body, spec.band) : spec.body;
          const base = w.styles.get(cl.s || 0);
          const d = { };
          if (parts.font) d.font = role && role.font ? Object.assign({ b: undefined, i: undefined, color: undefined }, role.font) : { b: undefined, i: undefined, color: undefined };
          if (parts.fill) d.fill = role && role.fill ? role.fill : null;
          if (parts.border) d.border = role && role.border ? Object.assign({ t: null, b: null, l: null, r: null }, role.border) : { t: null, b: null, l: null, r: null };
          if (parts.nf && role && role.nf && typeof s.val(r, c) === 'number') d.nf = role.nf;
          if (parts.align && r === rg.r1 && c > rg.c1 && role) d.align = { h: 'center' };
          void base;
          const ns = w.styles.derive(cl.s || 0, d);
          if (ns !== (cl.s || 0)) O.put(s, r, c, Object.assign({}, cl, { s: ns }));
        }
        if (parts.size) O.autofitCols(s, Array.from({ length: rg.c2 - rg.c1 + 1 }, (_, i) => rg.c1 + i), { r1: rg.r1, r2: rg.r2 });
      });
    }) }, { label: '取消' }] });
  };

  /* ================================================================ Style */
  /* the horizontal alignments, as Format Cells names them */
  const ALIGN_NAMES = { general: L.N('常规'), left: '靠左（缩进）', center: L.N('居中'), right: '靠右（缩进）', fill: L.N('填充'), justify: L.N('两端对齐'), centerContinuous: '跨列居中', distributed: '分散对齐（缩进）' };
  const BUILTIN_STYLES = { Normal: 0, Comma: 3, 'Comma [0]': 6, Currency: 4, 'Currency [0]': 7, Percent: 5, Hyperlink: 8, 'Followed Hyperlink': 9 };
  D.builtinStyle = function (w, name) {
    const base = w.styles.get(0);
    switch (name) {
      case 'Comma': return M.mergeStyle(base, { nf: NF.BUILTIN[43] });
      case 'Comma [0]': return M.mergeStyle(base, { nf: NF.BUILTIN[41] });
      case 'Currency': return M.mergeStyle(base, { nf: NF.BUILTIN[44] });
      case 'Currency [0]': return M.mergeStyle(base, { nf: NF.BUILTIN[42] });
      case 'Percent': return M.mergeStyle(base, { nf: '0%' });
      case 'Hyperlink': return M.mergeStyle(base, { font: { u: 'single', color: { theme: 10 } } });
      case 'Followed Hyperlink': return M.mergeStyle(base, { font: { u: 'single', color: { theme: 11 } } });
      default: return base;
    }
  };
  D.style = function () {
    const s = sh(), w = wb(), sel = G().sel();
    const cur = LY.styleOf(s, sel.r, sel.c, s.get(sel.r, sel.c));
    const names = Array.from(new Set(Object.keys(BUILTIN_STYLES).concat(w.cellStyles.filter((c) => !c.hidden).map((c) => c.name))));
    let name = cur.xs || 'Normal';
    const nameIn = h('input', { type: 'text', value: name, list: 'style-names', style: 'width:220px' });
    const dl = h('datalist', { id: 'style-names' }, ...names.map((n) => h('option', { value: n })));
    const styleOf = (n) => { const cs = w.cellStyles.find((c) => c.name === n); return cs ? w.styles.get(cs.style) : D.builtinStyle(w, n); };
    const inc = { nf: true, align: true, font: true, border: true, fill: true, prot: true };
    const descBox = h('div', { class: 'col' });
    const describe = () => {
      const st = styleOf(nameIn.value) || cur;
      descBox.textContent = '';
      const fo = st.font || {}, bi = fo.b && fo.i ? L.t('加粗、倾斜') : fo.b ? L.t('加粗') : fo.i ? L.t('倾斜') : '';
      const lines = [['nf', L.N('数字'), st.nf || 'General', 1], ['align', L.N('对齐方式'), st.align && st.align.h ? L.t(ALIGN_NAMES[st.align.h] || st.align.h) : L.t('常规、底端对齐')], ['font', L.N('字体'), `${LY.fontName(w, fo)} ${fo.sz || 10}${bi ? ' ' + bi : ''}`, 1], ['border', L.N('边框'), st.border ? L.t('边框') : L.t('无框线')], ['fill', L.N('图案'), st.fill ? L.t('底纹') : L.t('无底纹')], ['prot', L.N('保护'), st.prot && st.prot.locked === false ? L.t('未锁定') : L.t('锁定')]];
      for (const [k, l, v, data] of lines) descBox.appendChild(h('div', { class: 'row', style: 'gap:6px' }, ui.check(l, inc[k], (x) => { inc[k] = x; }), h('span', { class: 'hint', translate: data ? 'no' : null, text: v })));
    };
    nameIn.addEventListener('input', describe);
    describe();
    const fromCell = () => { const st = JSON.parse(JSON.stringify(cur)); delete st.xs; return st; };
    const save = (n, st) => {
      const idx = w.styles.add(st);
      const list = w.cellStyles.filter((c) => c.name !== n);
      const bi = BUILTIN_STYLES[n];
      list.push({ name: n, style: idx, builtinId: bi, customBuiltin: bi != null ? true : undefined });
      const old = w.cellStyles;
      O.tx(w, L.N('样式'), () => { w.cellStyles = list; w.undo.op(() => { w.cellStyles = old; }, () => { w.cellStyles = list; }); });
    };
    ui.dialog({ title: '样式', width: 420, body: h('div', { class: 'col' }, ui.field('样式名(&S)：', nameIn), dl, ui.group('样式包括', descBox)),
      buttons: [
        { label: '添加(&A)', onClick: () => { const n = nameIn.value.trim(); if (!n) return false; save(n, fromCell()); describe(); return false; } },
        { label: '删除(&D)', onClick: () => { const n = nameIn.value.trim(); if (n === 'Normal') { ui.msg('不能删除“正文”样式。', { icon: 'warn' }); return false; } const old = w.cellStyles; const list = old.filter((c) => c.name !== n); O.tx(w, L.N('样式'), () => { w.cellStyles = list; w.undo.op(() => { w.cellStyles = old; }, () => { w.cellStyles = list; }); }); nameIn.value = 'Normal'; describe(); return false; } },
        { label: '修改(&M)...', onClick: () => {
          const n = nameIn.value.trim() || 'Normal';
          const base = styleOf(n) || cur;
          D.formatCells(0, { style: base, title: L.t('设置单元格格式 - {0}', n), apply: (newSt) => {
            if (n === 'Normal') {
              /* the default style: every cell without a style of its own follows it */
              const old = w.styles.list[0];
              const st = Object.freeze(Object.assign({}, newSt));
              O.tx(w, '修改样式', () => { const set = (x) => { w.styles.map.delete(JSON.stringify(w.styles.list[0])); w.styles.list[0] = x; w.styles.map.set(JSON.stringify(x), 0); w._mdw = null; for (const s2 of w.sheets) LY.invalidate(s2); }; set(st); w.undo.op(() => set(old), () => set(st)); O.markStructural(); });
            } else save(n, newSt);
            describe();
          } });
          return false;
        } },
        { label: '确定', primary: true, onClick: () => tryRun(() => {
          const n = nameIn.value.trim() || 'Normal';
          let st = styleOf(n);
          if (!w.cellStyles.some((c) => c.name === n) && BUILTIN_STYLES[n] == null) { save(n, fromCell()); st = fromCell(); }
          O.format(s, G().ranges(), (cs) => {
            const d = {};
            if (inc.nf) d.nf = st.nf;
            if (inc.align) d.align = st.align;
            if (inc.font) d.font = Object.assign({ b: undefined, i: undefined, u: undefined, color: undefined }, st.font);
            if (inc.border) d.border = st.border ? Object.assign({ t: null, b: null, l: null, r: null }, st.border) : { t: null, b: null, l: null, r: null };
            if (inc.fill) d.fill = st.fill;
            if (inc.prot) d.prot = st.prot;
            d.xs = n === 'Normal' ? undefined : n;
            void cs;
            return d;
          }, L.N('样式'));
        }) }, { label: '取消' }] });
  };

  /* ================================================================ Goal Seek */
  D.goalSeek = function () {
    const s = sh(), w = wb(), sel = G().sel();
    const setC = refBox(F.cellName(sel.r, sel.c, true, true), 120), toV = h('input', { type: 'text', style: 'width:120px' }), byC = refBox('', 120);
    ui.dialog({ title: '单变量求解', width: 300, body: h('div', { class: 'col' }, ui.field('目标单元格(&e)：', setC), ui.field('目标值(&v)：', toV), ui.field('可变单元格(&c)：', byC)),
      buttons: [{ label: '确定', primary: true, onClick: () => {
        const t = parseRef(setC.input.value, true), ch = parseRef(byC.input.value, true), goal = num(toV.value, NaN);
        if (!t || !(t.sheet.get(t.r1, t.c1) && t.sheet.get(t.r1, t.c1).f != null)) { ui.msg('单元格必须包含公式。', { icon: 'warn' }); return false; }
        if (!ch || (ch.sheet.get(ch.r1, ch.c1) && ch.sheet.get(ch.r1, ch.c1).f != null)) { ui.msg('单元格必须包含值。', { icon: 'warn' }); return false; }
        if (isNaN(goal)) { ui.msg('无效值。', { icon: 'warn' }); return false; }
        const res = D.solveGoal(w, t, ch, goal);
        const before = res.start;
        tryRun(() => { const cell = ch.sheet.get(ch.r1, ch.c1); if (cell) cell.v = before; C.changed(w, [{ sh: ch.sheet, r: ch.r1, c: ch.c1 }]); O.tx(w, '单变量求解', () => O.setValue(ch.sheet, ch.r1, ch.c1, res.x, true)); });
        ui.dialog({ title: '单变量求解状态', width: 300, body: h('div', { class: 'col' }, h('div', { text: res.ok ? L.t('单元格 {0} 的单变量求解已找到解。', F.cellName(t.r1, t.c1)) : L.t('单元格 {0} 的单变量求解可能尚未找到解。', F.cellName(t.r1, t.c1)) }), h('table', { class: 'kv' }, h('tr', null, h('td', { text: '目标值：' }), h('td', { translate: 'no', text: String(goal) })), h('tr', null, h('td', { text: '当前值：' }), h('td', { translate: 'no', text: L.csv.cellText(t.sheet, t.r1, t.c1) })))),
          buttons: [{ label: '确定', primary: true }, { label: '取消', onClick: () => A().undo() }] });
      } }, { label: '取消' }] });
  };
  /** secant / bisection search for x with f(x) = goal; leaves the model as it was */
  D.solveGoal = function (w, t, ch, goal) {
    const cell = ch.sheet.get(ch.r1, ch.c1) || ch.sheet.cell(ch.r1, ch.c1);
    const start = typeof cell.v === 'number' ? cell.v : 0;
    const f = (x) => { cell.v = x; C.changed(w, [{ sh: ch.sheet, r: ch.r1, c: ch.c1 }]); const v = C.cellValue(t.sheet, t.r1, t.c1); return typeof v === 'number' ? v - goal : NaN; };
    let x0 = start, x1 = start === 0 ? 0.01 : start * 1.01;
    let f0 = f(x0), f1 = f(x1), best = { x: x0, e: Math.abs(f0) };
    let ok = false;
    for (let i = 0; i < 100; i++) {
      if (isFinite(f1) && Math.abs(f1) < best.e) best = { x: x1, e: Math.abs(f1) };
      if (Math.abs(f1) <= 0.001 * Math.max(1, Math.abs(goal)) * 1e-3 || Math.abs(f1) < 1e-9) { ok = true; break; }
      if (!isFinite(f0) || !isFinite(f1) || f1 === f0) { x1 = x1 * 2 + 1; f1 = f(x1); continue; }
      const x2 = x1 - f1 * (x1 - x0) / (f1 - f0);
      x0 = x1; f0 = f1; x1 = x2; f1 = f(x1);
    }
    if (!ok && best.e < 0.001 * Math.max(1, Math.abs(goal))) ok = true;
    cell.v = start;
    C.changed(w, [{ sh: ch.sheet, r: ch.r1, c: ch.c1 }]);
    return { x: ok ? x1 : best.x, ok, start };
  };

  /* ================================================================ Scenarios */
  D.scenarios = function () {
    const s = sh(), w = wb();
    if (!s.scenarios) s.scenarios = [];
    const box = h('div');
    let cur = 0;
    const draw = () => { box.textContent = ''; const lb = D.list(s.scenarios.map((x, i) => [i, x.name]), s.scenarios.length ? cur : null, (i) => { cur = i; info.textContent = s.scenarios[i] ? L.t('可变单元格：{0}\n批注：{1}', s.scenarios[i].cells.map((c) => F.cellName(c.r, c.c, true, true)).join(','), s.scenarios[i].comment || '') : ''; }, { translate: 'no' }); lb.ondbl = () => show(); box.appendChild(s.scenarios.length ? lb : h('div', { class: 'tp-note', text: '尚未定义方案。请选择“添加”来添加方案。' })); };
    const info = h('div', { class: 'hint', style: 'white-space:pre-line' });
    const show = () => { const sc = s.scenarios[cur]; if (!sc) return; tryRun(() => O.tx(w, '显示方案', () => { for (const c of sc.cells) O.enter(s, c.r, c.c, String(c.v)); })); };
    const edit = (sc) => {
      const nm = h('input', { type: 'text', value: sc ? sc.name : '', style: 'width:100%' });
      const cells = refBox(sc ? sc.cells.map((c) => F.cellName(c.r, c.c, true, true)).join(',') : F.absRangeName(G().range()), 220);
      const cm = h('textarea', { rows: '3', style: 'width:100%' });
      cm.value = sc ? sc.comment || '' : L.t('由 {0} 创建于 {1}', A().opts.userName || '', new Date().toLocaleDateString());
      ui.dialog({ title: sc ? '编辑方案' : '添加方案', width: 360, body: h('div', { class: 'col' }, ui.field('方案名(&n)：', nm), ui.field('可变单元格(&c)：', cells), ui.field('批注(&m)：', cm)),
        buttons: [{ label: '确定', primary: true, onClick: () => {
          const refs = [];
          for (const part of cells.input.value.split(',')) { const p = parseRef(part); if (!p) { ui.msg('引用无效。', { icon: 'warn' }); return false; } for (let r = p.r1; r <= p.r2; r++) for (let c = p.c1; c <= p.c2; c++) refs.push({ r, c }); }
          if (!nm.value.trim()) return false;
          /* Scenario Values: one box per changing cell */
          const ins = refs.map((x) => h('input', { type: 'text', value: (sc && (sc.cells.find((q) => q.r === x.r && q.c === x.c) || {}).v) != null ? String(sc.cells.find((q) => q.r === x.r && q.c === x.c).v) : L.editor.cellText(s, x.r, x.c), style: 'width:120px' }));
          ui.dialog({ title: '方案值', width: 300, body: h('div', { class: 'col' }, h('div', { text: '为各个可变单元格输入值。' }), ...refs.map((x, i) => ui.field(F.cellName(x.r, x.c, true, true) + ':', ins[i]))),
            buttons: [{ label: '确定', primary: true, onClick: () => {
              const item = { name: nm.value.trim(), comment: cm.value, cells: refs.map((x, i) => { const p = NF.parseInput(ins[i].value, {}); return { r: x.r, c: x.c, v: p && typeof p.v === 'number' ? p.v : ins[i].value }; }) };
              const old = s.scenarios.slice();
              const list = sc ? s.scenarios.map((q) => (q === sc ? item : q)) : s.scenarios.concat([item]);
              O.tx(w, L.N('方案'), () => { s.scenarios = list; w.undo.op(() => { s.scenarios = old; }, () => { s.scenarios = list; }); });
              cur = list.indexOf(item); draw();
            } }, { label: '取消' }] });
        } }, { label: '取消' }] });
    };
    const summary = () => {
      if (!s.scenarios.length) return;
      tryRun(() => {
        const name = w.sheetByName('Scenario Summary') ? w.nextSheetName('Scenario Summary ') : 'Scenario Summary';
        const ns = O.addSheet(w, name, w.sheets.indexOf(s) + 1);
        O.tx(w, '方案摘要', () => {
          const allCells = [];
          for (const sc of s.scenarios) for (const c of sc.cells) if (!allCells.some((x) => x.r === c.r && x.c === c.c)) allCells.push(c);
          const B = { font: { b: true } };
          const put = (r, c, v, st) => { const cell = { v }; if (st) cell.s = w.styles.derive(0, st); O.put(ns, r, c, cell); };
          put(0, 1, 'Scenario Summary', { font: { b: true, sz: 12 } });
          put(2, 2, 'Current Values:', B);
          s.scenarios.forEach((sc, i) => put(2, 3 + i, sc.name, B));
          put(3, 1, 'Changing Cells:', B);
          allCells.forEach((c, k) => { put(4 + k, 1, F.cellName(c.r, c.c, true, true)); put(4 + k, 2, s.val(c.r, c.c)); s.scenarios.forEach((sc, i) => { const x = sc.cells.find((q) => q.r === c.r && q.c === c.c); put(4 + k, 3 + i, x ? x.v : s.val(c.r, c.c)); }); });
          const base = 5 + allCells.length;
          put(base, 1, 'Result Cells:', B);
          /* result cells: the formulas that depend on the changing cells */
          const results = [];
          for (const c of allCells) { const deps = []; w.graph.dependents(s, c.r, c.c, deps); for (const d of deps) if (d.sh === s && !results.some((x) => x.r === d.r && x.c === d.c)) results.push({ r: d.r, c: d.c }); }
          results.slice(0, 30).forEach((rc, k) => {
            put(base + 1 + k, 1, F.cellName(rc.r, rc.c, true, true)); put(base + 1 + k, 2, val(s, rc.r, rc.c));
            s.scenarios.forEach((sc, i) => {
              const saved = sc.cells.map((x) => { const cl = s.get(x.r, x.c); return [cl, cl ? cl.v : null]; });
              sc.cells.forEach((x) => { const cl = s.get(x.r, x.c) || s.cell(x.r, x.c); cl.v = x.v; });
              C.changed(w, sc.cells.map((x) => ({ sh: s, r: x.r, c: x.c })));
              put(base + 1 + k, 3 + i, val(s, rc.r, rc.c));
              saved.forEach(([cl, v]) => { if (cl) cl.v = v; });
              C.changed(w, sc.cells.map((x) => ({ sh: s, r: x.r, c: x.c })));
            });
          });
          ns.cols[1] = { w: 16, custom: true }; ns.cols[2] = { w: 15, custom: true };
          for (let i = 0; i < s.scenarios.length; i++) ns.cols[3 + i] = { w: 13, custom: true };
        });
        w.active = w.sheets.indexOf(ns);
        A().renderTabs();
      });
    };
    draw();
    ui.dialog({ title: '方案管理器', width: 420, body: h('div', { class: 'col' }, h('label', { text: '方案(&e)：' }), box, info),
      buttons: [{ label: '显示(&S)', onClick: () => { show(); return false; } }, { label: '添加(&A)...', onClick: () => { edit(null); return false; } }, { label: '删除(&D)', onClick: () => { const old = s.scenarios.slice(); const list = old.filter((_, i) => i !== cur); O.tx(w, L.N('方案'), () => { s.scenarios = list; w.undo.op(() => { s.scenarios = old; }, () => { s.scenarios = list; }); }); cur = 0; draw(); return false; } }, { label: '编辑(&E)...', onClick: () => { edit(s.scenarios[cur]); return false; } }, { label: '摘要(&u)...', onClick: () => { summary(); } }, { label: '关闭', primary: true }] });
  };

  /* ================================================================ Data Table (what-if) */
  D.dataTable = function () {
    const s = sh();
    const rg = O.clip(s, G().range());
    if (rg.r2 <= rg.r1 && rg.c2 <= rg.c1) { ui.msg('选择数据表区域，须包含公式和输入值。', { icon: 'warn' }); return; }
    const rowIn = refBox('', 120), colIn = refBox('', 120);
    ui.dialog({ title: '表格', width: 300, body: h('div', { class: 'col' }, ui.field('输入引用行的单元格(&R)：', rowIn), ui.field('输入引用列的单元格(&C)：', colIn)),
      buttons: [{ label: '确定', primary: true, onClick: () => {
        const ri = rowIn.input.value.trim() ? parseRef(rowIn.input.value, true) : null, ci = colIn.input.value.trim() ? parseRef(colIn.input.value, true) : null;
        if (!ri && !ci) { ui.msg('输入单元格引用无效。', { icon: 'warn' }); return false; }
        tryRun(() => D.buildDataTable(s, rg, ri, ci));
      } }, { label: '取消' }] });
  };
  /** fill a one- or two-variable data table (values; the TABLE() master keeps it recalculable on save) */
  D.buildDataTable = function (s, rg, ri, ci) {
    const w = s.wb;
    const two = ri && ci;
    const res = { r1: rg.r1 + 1, c1: rg.c1 + 1, r2: rg.r2, c2: rg.c2 };
    const inputs = [ri, ci].filter(Boolean).map((p) => { const cl = p.sheet.get(p.r1, p.c1) || p.sheet.cell(p.r1, p.c1); return { p, cl, v: cl.v }; });
    const setIn = (p, v) => { const cl = p.sheet.get(p.r1, p.c1) || p.sheet.cell(p.r1, p.c1); cl.v = v; C.changed(w, [{ sh: p.sheet, r: p.r1, c: p.c1 }]); };
    const out = [];
    try {
      for (let r = res.r1; r <= res.r2; r++) for (let c = res.c1; c <= res.c2; c++) {
        let fr, fc;
        if (two) { setIn(ri, val(s, rg.r1, c)); setIn(ci, val(s, r, rg.c1)); fr = rg.r1; fc = rg.c1; }
        else if (ci) { setIn(ci, val(s, r, rg.c1)); fr = rg.r1; fc = c; }
        else { setIn(ri, val(s, rg.r1, c)); fr = r; fc = rg.c1; }
        out.push({ r, c, v: val(s, fr, fc) });
      }
    } finally { for (const x of inputs) { x.cl.v = x.v; C.changed(w, [{ sh: x.p.sheet, r: x.p.r1, c: x.p.c1 }]); } }
    O.tx(w, L.N('表格'), () => {
      for (const x of out) { const old = s.get(x.r, x.c); const n = Object.assign({}, old || {}); delete n.f; n.v = M.isErr(x.v) ? x.v : x.v; O.put(s, x.r, x.c, n); }
      const master = Object.assign({}, s.get(res.r1, res.c1) || {});
      master.dt = { ref: F.rangeName(res), dt2D: !!two, dtr: !!(ri && !ci), r1: ri ? F.cellName(ri.r1, ri.c1) : F.cellName(ci.r1, ci.c1), r2: two ? F.cellName(ci.r1, ci.c1) : undefined };
      O.put(s, res.r1, res.c1, master);
    });
  };

  /* ================================================================ Consolidate */
  const CONS_FN = [['SUM', L.N('求和')], ['COUNT', L.N('计数')], ['AVERAGE', L.N('平均值')], ['MAX', L.N('最大值')], ['MIN', L.N('最小值')], ['PRODUCT', L.N('乘积')], ['COUNTNUMS', '数值计数'], ['STDEV', L.N('样本标准偏差')], ['STDEVP', L.N('总体标准偏差')], ['VAR', L.N('样本方差')], ['VARP', L.N('总体方差')]];
  D.consolidate = function () {
    const s = sh(), w = wb(), sel = G().sel();
    let fn = 'SUM', top = false, left = false;
    const refs = (w._consRefs || []).slice();
    const refIn = refBox('', 220);
    const box = h('div');
    let cur = -1;
    const draw = () => { box.textContent = ''; box.appendChild(D.list(refs.map((x, i) => [i, x]), cur >= 0 ? cur : null, (i) => { cur = i; refIn.input.value = refs[i]; }, { class: 'lbox short', translate: 'no' })); };
    draw();
    ui.dialog({ title: '合并计算', width: 380, body: h('div', { class: 'col' }, ui.field('函数(&F)：', ui.select(CONS_FN, fn, (v) => { fn = v; })), ui.field('引用(&R)：', refIn), h('div', { class: 'row', style: 'gap:6px' }, ui.button('添加(&A)', () => { const t = refIn.input.value.trim(); if (!parseRef(t)) { ui.msg('引用无效。', { icon: 'warn' }); return; } if (!refs.includes(t)) refs.push(t); draw(); }), ui.button('删除(&D)', () => { if (cur >= 0) { refs.splice(cur, 1); cur = -1; draw(); } })), h('label', { text: '所有引用(&e)：' }), box, ui.group('标签位置', ui.check('最上行(&T)', false, (v) => { top = v; }), ui.check('最左列(&L)', false, (v) => { left = v; }))),
      buttons: [{ label: '确定', primary: true, onClick: () => tryRun(() => {
        w._consRefs = refs.slice();
        const srcs = refs.map((t) => parseRef(t)).filter(Boolean);
        if (!srcs.length) return;
        const rowKeys = [], colKeys = [];
        const data = new Map();
        for (const p of srcs) {
          const r0 = p.r1 + (top ? 1 : 0), c0 = p.c1 + (left ? 1 : 0);
          for (let r = r0; r <= Math.min(p.r2, p.sheet.maxR); r++) for (let c = c0; c <= Math.min(p.c2, p.sheet.maxC); c++) {
            const rk = left ? String(L.csv.cellText(p.sheet, r, p.c1)) : String(r - r0);
            const ck = top ? String(L.csv.cellText(p.sheet, p.r1, c)) : String(c - c0);
            if (!rowKeys.includes(rk)) rowKeys.push(rk);
            if (!colKeys.includes(ck)) colKeys.push(ck);
            const v = val(p.sheet, r, c);
            const k = rk + '\u0001' + ck;
            if (!data.has(k)) data.set(k, []);
            data.get(k).push(v);
          }
        }
        const agg = (vals) => { const n = vals.filter((x) => typeof x === 'number'); switch (fn) { case 'COUNT': return vals.filter((x) => x != null && x !== '').length; case 'COUNTNUMS': return n.length; case 'AVERAGE': return n.length ? n.reduce((a, b) => a + b, 0) / n.length : 0; case 'MAX': return n.length ? Math.max(...n) : 0; case 'MIN': return n.length ? Math.min(...n) : 0; case 'PRODUCT': return n.reduce((a, b) => a * b, 1); case 'STDEV': case 'STDEVP': case 'VAR': case 'VARP': { const m = n.reduce((a, b) => a + b, 0) / (n.length || 1); const ss = n.reduce((a, b) => a + (b - m) * (b - m), 0); const vv = ss / Math.max(1, fn.endsWith('P') ? n.length : n.length - 1); return /STDEV/.test(fn) ? Math.sqrt(vv) : vv; } default: return n.reduce((a, b) => a + b, 0); } };
        O.tx(w, L.N('合并计算'), () => {
          const r0 = sel.r + (top ? 1 : 0), c0 = sel.c + (left ? 1 : 0);
          if (top) colKeys.forEach((k, j) => O.put(s, sel.r, c0 + j, { v: k }));
          if (left) rowKeys.forEach((k, i) => O.put(s, r0 + i, sel.c, { v: k }));
          rowKeys.forEach((rk, i) => colKeys.forEach((ck, j) => { const vals = data.get(rk + '\u0001' + ck); if (vals) O.put(s, r0 + i, c0 + j, { v: agg(vals) }); }));
        });
      }) }, { label: '关闭' }] });
  };

  /* ================================================================ PivotTable report (static summary) */
  const PIV_FN = [['sum', L.N('求和')], ['count', L.N('计数')], ['average', L.N('平均值')], ['max', L.N('最大值')], ['min', L.N('最小值')], ['countNums', '数值计数']];
  D.pivotTable = function () {
    const s = sh(), w = wb();
    const rg = listRange();
    if (rg.r2 <= rg.r1) { ui.msg('数据透视表字段名无效。要创建数据透视表，必须使用按列表组织且各列有标签的数据。', { icon: 'warn' }); return; }
    const cols = Array.from({ length: rg.c2 - rg.c1 + 1 }, (_, i) => [rg.c1 + i, colLabel(s, rg, true, rg.c1 + i)]);
    const numericCol = cols.find(([c]) => typeof val(s, rg.r1 + 1, c) === 'number');
    const st = { row: cols[0][0], col: '', data: numericCol ? numericCol[0] : cols[cols.length - 1][0], fn: numericCol ? 'sum' : 'count', page: '', pageVal: '', where: 'new' };
    const src = refBox(F.quoteSheet(s.name) + '!' + F.absRangeName(rg), 220);
    const opt = (k, none) => ui.select((none ? [['', L.t('（无）')]] : []).concat(cols), st[k], (v) => { st[k] = v === '' ? '' : +v; }, { translate: 'no' });
    ui.dialog({ title: '数据透视表和数据透视图向导', width: 420, body: h('div', { class: 'col' },
      ui.field('区域(&R)：', src),
      ui.group('布局', ui.field('行字段(&R)：', opt('row')), ui.field('列字段(&C)：', opt('col', true)), h('div', { class: 'row', style: 'gap:8px' }, ui.field('数据字段(&D)：', opt('data')), ui.field('汇总方式：', ui.select(PIV_FN, st.fn, (v) => { st.fn = v; }))), ui.field('页码域(&P)：', opt('page', true))),
      ui.group('您希望将数据透视表放在哪里？', ui.radio('pv-w', '新建工作表(&N)', true, () => { st.where = 'new'; }), ui.radio('pv-w', '现有工作表（从活动单元格开始）(&E)', false, () => { st.where = 'here'; })),
      h('div', { class: 'hint', text: 'Ledger 将报表生成为带格式的数值汇总；源数据更改后需重新生成。' })),
    buttons: [{ label: '完成(&F)', primary: true, onClick: () => {
      const p = parseRef(src.input.value);
      if (!p) { ui.msg('引用无效。', { icon: 'warn' }); return false; }
      tryRun(() => D.buildPivot(p, st));
    } }, { label: '取消' }] });
  };
  D.buildPivot = function (p, st) {
    const w = wb(), s = p.sheet;
    const r2 = Math.min(p.r2, s.maxR);
    const rowKeys = [], colKeys = [];
    const cells = new Map(), rowTot = new Map(), colTot = new Map(), all = [];
    const keyOf = (r, c) => { const t = L.csv.cellText(s, r, c); return t === '' ? '(blank)' : t; };
    const sortKey = (r, c) => { const v = val(s, r, c); return v; };
    const rowSort = new Map(), colSort = new Map();
    for (let r = p.r1 + 1; r <= r2; r++) {
      if (st.page !== '' && st.pageVal && keyOf(r, st.page) !== st.pageVal) continue;
      const rk = keyOf(r, st.row), ck = st.col === '' ? '' : keyOf(r, st.col);
      if (!rowSort.has(rk)) { rowSort.set(rk, sortKey(r, st.row)); rowKeys.push(rk); }
      if (!colSort.has(ck)) { colSort.set(ck, st.col === '' ? '' : sortKey(r, st.col)); colKeys.push(ck); }
      const v = val(s, r, st.data);
      const k = rk + '\u0001' + ck;
      for (const [map, kk] of [[cells, k], [rowTot, rk], [colTot, ck]]) { if (!map.has(kk)) map.set(kk, []); map.get(kk).push(v); }
      all.push(v);
    }
    const cmp = (m) => (a, b) => O.compareValues(m.get(a), m.get(b), {});
    rowKeys.sort(cmp(rowSort)); colKeys.sort(cmp(colSort));
    const agg = (vals) => { const n = vals.filter((x) => typeof x === 'number'); switch (st.fn) { case 'count': return vals.filter((x) => x != null && x !== '').length; case 'countNums': return n.length; case 'average': return n.length ? n.reduce((a, b) => a + b, 0) / n.length : M.ERR.DIV0; case 'max': return n.length ? Math.max(...n) : 0; case 'min': return n.length ? Math.min(...n) : 0; default: return n.reduce((a, b) => a + b, 0); } };
    const fnLabel = PIV_FN.find((x) => x[0] === st.fn)[1];
    const head = (c) => String(s.val(p.r1, c) == null ? F.colName(c) : s.val(p.r1, c));
    let dst, r0, c0;
    O.tx(w, L.N('数据透视表'), () => {
      if (st.where === 'new') { dst = O.addSheet(w, w.nextSheetName('Sheet'), w.sheets.indexOf(s)); r0 = 2; c0 = 0; }
      else { dst = sh(); r0 = G().sel().r; c0 = G().sel().c; }
      const BOLD = w.styles.derive(0, { font: { b: true } });
      const HEAD = w.styles.derive(0, { font: { b: true }, fill: { pattern: 'solid', fg: M.rgb('#C0C0C0') }, border: { t: { style: 'thin' }, b: { style: 'thin' } } });
      const TOT = w.styles.derive(0, { font: { b: true }, fill: { pattern: 'solid', fg: M.rgb('#E3E3E3') }, border: { t: { style: 'thin' } } });
      const put = (r, c, v, sid) => O.put(dst, r, c, { v, s: sid || undefined });
      if (st.page !== '') { put(r0 - 2 < 0 ? r0 : r0 - 2, c0, head(st.page), BOLD); put(r0 - 2 < 0 ? r0 : r0 - 2, c0 + 1, st.pageVal || '(All)'); }
      put(r0, c0, fnLabel + ' of ' + head(st.data), HEAD);
      if (st.col !== '') put(r0, c0 + 1, head(st.col), HEAD);
      const hr = r0 + 1;
      put(hr, c0, head(st.row), HEAD);
      const showCols = st.col === '' ? [''] : colKeys;
      showCols.forEach((ck, j) => put(hr, c0 + 1 + j, st.col === '' ? 'Total' : ck, HEAD));
      if (st.col !== '') put(hr, c0 + 1 + showCols.length, 'Grand Total', HEAD);
      rowKeys.forEach((rk, i) => {
        put(hr + 1 + i, c0, rk);
        showCols.forEach((ck, j) => { const vals = cells.get(rk + '\u0001' + ck); if (vals) put(hr + 1 + i, c0 + 1 + j, agg(vals)); });
        if (st.col !== '') put(hr + 1 + i, c0 + 1 + showCols.length, agg(rowTot.get(rk)), BOLD);
      });
      const gr = hr + 1 + rowKeys.length;
      put(gr, c0, 'Grand Total', TOT);
      showCols.forEach((ck, j) => put(gr, c0 + 1 + j, agg(colTot.get(ck) || []), TOT));
      if (st.col !== '') put(gr, c0 + 1 + showCols.length, agg(all), TOT);
      O.autofitCols(dst, Array.from({ length: showCols.length + 2 }, (_, i) => c0 + i));
    });
    if (st.where === 'new') { w.active = w.sheets.indexOf(dst); A().renderTabs(); }
  };
})(typeof window !== 'undefined' ? window : globalThis);
