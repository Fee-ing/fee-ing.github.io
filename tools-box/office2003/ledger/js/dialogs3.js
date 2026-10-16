/* Ledger — dialogs, part 3 (formulas and printing): Insert Function, Function Arguments, names, Evaluate Formula,
 * Watch Window, Error Checking, Spelling, Page Setup with Header/Footer, Print. */
(function (root) {
  'use strict';
  const L = root.L;
  const M = L.model, O = L.ops, F = L.formula, NF = L.numfmt, LY = L.layout, C = L.calc;
  const { h } = L;
  const ui = L.ui;
  const D = (L.dlg = L.dlg || {});
  const G = () => L.grid;
  const A = () => L.app;
  const E = () => L.editor;
  const sh = () => G().sheet();
  const wb = () => G().wb;
  const tryRun = (fn) => { try { fn(); } catch (e) { A().error(e); } G().paint(); G().syncObjects(); ui.refresh(); };
  /** a value as Excel shows it in dialogs: "text", 12.5, TRUE, {1,2;3,4} */
  function show(v, depth) {
    if (v instanceof C.Ref) {
      if (v.single) return show(C.cellValue(v.sheet, v.r1, v.c1));
      const rows = [];
      for (let r = v.r1; r <= Math.min(v.r2, v.r1 + 5); r++) { const row = []; for (let c = v.c1; c <= Math.min(v.c2, v.c1 + 5); c++) row.push(show(C.cellValue(v.sheet, r, c), 1)); rows.push(row.join(',') + (v.c2 > v.c1 + 5 ? ',...' : '')); }
      return '{' + rows.join(';') + (v.r2 > v.r1 + 5 ? ';...' : '') + '}';
    }
    if (v instanceof C.RefList) return v.list.map((x) => show(x)).join(',');
    if (v && v.rows && Array.isArray(v.rows)) { const rows = v.rows.slice(0, 6).map((r) => r.slice(0, 6).map((x) => show(x, 1)).join(',')); return '{' + rows.join(';') + '}'; }
    if (v == null) return depth ? '0' : '';
    if (M.isErr(v)) return v.e;
    if (typeof v === 'string') return '"' + v + '"';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'number') return NF.general(v, 15, true);
    return String(v);
  }
  D.showValue = show;
  const evalText = (text) => {
    const s = sh(), sel = G().sel();
    try { const ast = F.parse(String(text).replace(/^=/, '')); return C.ev(ast, C.ctx(wb(), s, sel.r, sel.c)); } catch (e) { return null; }
  };

  /* ================================================================ Insert Function */
  const MRU = () => L.store.get('fnMRU', ['SUM', 'AVERAGE', 'IF', 'HYPERLINK', 'COUNT', 'MAX', 'SIN', 'SUMIF', 'PMT', 'STDEV']);
  D.insertFunction = function (prefer) {
    const FI = L.fninfo;
    let cat = prefer ? 'All' : '最近使用';
    let pick = prefer || null;
    const search = h('input', { type: 'text', placeholder: '简要描述您要执行的操作，然后单击“转到”', style: 'flex:1;min-width:0' });
    const listBox = h('div');
    const syntax = h('div', { class: 'if-syntax', translate: 'no' }), desc = h('div', { class: 'if-desc' });
    const describe = (name) => { const f = FI.get(name); syntax.textContent = f ? FI.syntax(f) : ''; desc.textContent = f ? L.t(f.desc) : ''; };
    let lb;
    const draw = (names) => {
      listBox.textContent = '';
      lb = D.list(names.map((n) => [n, n]), names.includes(pick) ? pick : names[0], (n) => { pick = n; describe(n); }, { class: 'lbox fnlist', translate: 'no' });
      lb.ondbl = (n) => { pick = n; d.close(1); go(); };
      listBox.appendChild(lb);
      if (!names.includes(pick)) pick = names[0] || null;
      describe(pick);
    };
    const byCat = () => {
      if (cat === 'Most Recently Used') return MRU().filter((n) => FI.get(n));
      const all = FI.all().filter((f) => cat === 'All' || f.cat === cat).map((f) => f.name);
      return all.sort();
    };
    const catSel = ui.select(FI.categories, cat, (v) => { cat = v; draw(byCat()); });
    const doSearch = () => {
      const words = search.value.toLowerCase().split(/[^\p{L}\p{N}_]+/u).filter((w) => w.length > 2 || /[^\x00-\x7f]/.test(w));
      if (!words.length) return;
      const scored = FI.all().map((f) => { const t = (f.name + ' ' + f.desc + ' ' + L.t(f.desc)).toLowerCase(); let sc = 0; for (const w of words) { if (f.name.toLowerCase() === w) sc += 10; if (t.includes(w)) sc += 2; if (t.includes(w.replace(/s$/, ''))) sc += 1; } return [f.name, sc]; }).filter((x) => x[1] > 0).sort((a, b) => b[1] - a[1]).slice(0, 30).map((x) => x[0]);
      catSel.value = 'Recommended';
      draw(scored.length ? scored : []);
    };
    search.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); e.stopPropagation(); doSearch(); } });
    draw(byCat());
    const go = () => {
      if (!pick) return;
      const list = MRU().filter((n) => n !== pick); list.unshift(pick); L.store.set('fnMRU', list.slice(0, 10));
      D.functionArgs(pick, true);
    };
    const d = ui.dialog({ title: '插入函数', width: 460, body: h('div', { class: 'col' }, h('label', { text: '搜索函数：' }), h('div', { class: 'row', style: 'gap:6px' }, search, ui.button('转到(&G)', doSearch)), ui.field('或选择类别(&c)：', catSel), h('label', { text: '选择函数(&n)：' }), listBox, syntax, desc),
      buttons: [{ label: '确定', primary: true, onClick: () => go() }, { label: '取消' }] });
  };

  /* ================================================================ Function Arguments */
  D.functionArgs = function (name, fresh) {
    const FI = L.fninfo;
    const ed = E();
    let existing = null;
    if (!name && ed.active) { const cf = ed.currentFunction(); if (cf) name = cf.name; }
    if (!name) { const s = sh(), sel = G().sel(); const cl = s.get(sel.r, sel.c); if (cl && cl.f != null) { const m = /^\s*(?:_xlfn\.)?([A-Z][A-Z0-9.]*)\(/i.exec(cl.f); if (m) { name = m[1].toUpperCase(); existing = F.display(cl.f); } } }
    if (!name) { D.insertFunction(); return; }
    const f = FI.get(name) || { name, args: ['value1', '[value2]', '...'], desc: '' };
    /* existing arguments of a formula consisting of this one call */
    let preset = [];
    if (existing && !fresh) { try { const ast = F.parse(existing); if (ast.t === 'fn') preset = ast.args.map((a) => F.toText(a)); } catch (e) { preset = []; } }
    const base = f.args.filter((a) => a !== '...');
    const repeating = f.args.includes('...');
    const rows = h('div', { class: 'fa-rows' });
    const inputs = [];
    const help = h('div', { class: 'fa-help' });
    const result = h('div', { class: 'fa-result' });
    const total = h('div', { class: 'fa-total' });
    const argName = (i) => { if (i < base.length) return base[i]; const last = base[base.length - 1] || 'value1'; const m = /^(\[?)([a-z_]+?)(\d+)(\]?)$/.exec(last); return m ? '[' + m[2] + (+m[3] + i - base.length + 1) + ']' : last; };
    const addRow = (i, value) => {
      const nm = argName(i);
      const inp = h('input', { type: 'text', value: value || '', class: 'fa-in' });
      const out = h('span', { class: 'fa-val' });
      const upd = () => { const t = inp.value.trim(); out.textContent = t ? '= ' + show(evalText(t)) : nm.startsWith('[') ? '' : '= '; refresh(); };
      inp.addEventListener('input', upd);
      inp.addEventListener('focus', () => { help.textContent = ''; help.append(h('b', { translate: 'no', text: nm.replace(/[[\]]/g, '').replace(/^./, (c) => c.toUpperCase()) + ' ' }), document.createTextNode(L.t(FI.argHelp(nm)) || '')); if (repeating && i === inputs.length - 1) addRow(i + 1, ''); });
      const label = h('label', { class: nm.startsWith('[') ? 'opt' : 'req', translate: 'no', text: nm.replace(/[[\]]/g, '').replace(/^./, (c) => c.toUpperCase()) });
      const rb = D.refBox ? null : null;
      void rb;
      rows.appendChild(h('div', { class: 'fa-row' }, label, inp, out));
      inputs.push(inp);
      upd();
    };
    const callText = () => { const args = inputs.map((i) => i.value.trim()); while (args.length && !args[args.length - 1] && args.length > base.filter((a) => !a.startsWith('[')).length) args.pop(); return f.name + '(' + args.join(',') + ')'; };
    const n0 = Math.max(base.length, preset.length, 1);
    for (let i = 0; i < n0; i++) addRow(i, preset[i]);
    function refresh() { if (!inputs.length) return; const v = evalText(callText()); total.textContent = L.t('公式结果 = {0}', v == null ? '' : show(v)); result.textContent = '= ' + (v == null ? '' : show(v)); }
    refresh();
    help.textContent = L.t(f.desc);
    ui.dialog({ title: '函数参数', width: 520, body: h('div', { class: 'col' }, h('b', { translate: 'no', text: f.name }), rows, result, h('div', { class: 'fa-desc', text: f.desc }), help, total),
      buttons: [{ label: '确定', primary: true, onClick: () => {
        const call = callText();
        if (ed.active) {
          /* replace a partly typed name before the caret, then insert the call */
          const t = ed.text(), surf = document.activeElement;
          void surf;
          const m = new RegExp('(?:_xlfn\\.)?' + f.name.replace('.', '\\.') + '\\($', 'i');
          const caret = (G().input.selectionStart != null ? G().input.selectionStart : t.length);
          const before = t.slice(0, caret), after = t.slice(caret);
          if (m.test(before)) ed.setText(before.replace(m, '') + call + after.replace(/^\)/, ''), before.replace(m, '').length + call.length);
          else ed.insertText(call);
          return;
        }
        const s = sh(), sel = G().sel();
        if (G().guardProtect(null, true)) return;
        ed.begin('enter', '=' + (existing && !fresh ? call : call));
        ed.commit();
        void s; void sel;
      } }, { label: '取消' }] });
  };

  /* ================================================================ names */
  const visibleNames = (w) => w.names.filter((n) => !n.hidden && !/^_xlnm\./i.test(n.name));
  const qualified = (w, n) => (n.scope != null && w.sheets[n.scope] ? F.quoteSheet(w.sheets[n.scope].name) + '!' : '') + n.name;
  D.validName = function (name) {
    if (!/^[A-Za-z_\\À-￿][\w.\\À-￿]*$/.test(name)) return '您输入的名称无效。';
    if (/^[A-Za-z]{1,3}\d+$/.test(name) || /^R\d*C\d*$/i.test(name) || /^[rRcC]$/.test(name)) return '您输入的名称无效。\n\n可能的原因包括：\n• 名称未以字母或下划线开头\n• 名称包含空格或其他无效字符\n• 名称与 Excel 内置名称或工作簿中其他对象的名称冲突';
    if (name.length > 255) return '名称过长。';
    return null;
  };
  D.defineName = function () {
    const w = wb(), s = sh();
    let names = w.names.map((n) => Object.assign({}, n));
    const nameIn = h('input', { type: 'text', style: 'width:100%' });
    const refIn = D.refBox(F.quoteSheet(s.name) + '!' + F.absRangeName(G().range()), 300);
    const comment = h('input', { type: 'text', style: 'width:100%' });
    const box = h('div');
    const draw = () => { box.textContent = ''; const lb = D.list(names.filter((n) => !n.hidden && !/^_xlnm\./i.test(n.name)).map((n) => [qualified(w, n), qualified(w, n)]), null, (q) => { const n = names.find((x) => qualified(w, x) === q); nameIn.value = q; refIn.input.value = '=' + String(n.ref).replace(/^=/, ''); comment.value = n.comment || ''; }, { class: 'lbox short', translate: 'no' }); box.appendChild(lb); };
    draw();
    /* a name for the active cell's label, as Excel proposes */
    const sel = G().sel();
    const left = s.val(sel.r, sel.c - 1), above = s.val(sel.r - 1, sel.c);
    const proposal = typeof left === 'string' ? left : typeof above === 'string' ? above : '';
    nameIn.value = proposal.trim().replace(/\s+/g, '_').replace(/[^\w.]/g, '');
    const add = () => {
      let nm = nameIn.value.trim();
      if (!nm) return false;
      let scope = null;
      const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(nm);
      if (m) { const sname = (m[1] || m[2]).replace(/''/g, "'"); const idx = w.sheets.findIndex((x) => x.name.toLowerCase() === sname.toLowerCase()); if (idx < 0) { ui.msg('引用无效。', { icon: 'warn' }); return false; } scope = idx; nm = m[3]; }
      const err = D.validName(nm);
      if (err) { ui.msg(err, { icon: 'warn' }); return false; }
      let ref = refIn.input.value.trim().replace(/^=/, '');
      try { F.parse(ref); } catch (e) { ui.msg('您输入的公式有错误。', { icon: 'warn' }); return false; }
      ref = F.toStore(ref);
      names = names.filter((x) => !(x.name.toLowerCase() === nm.toLowerCase() && (x.scope == null ? null : x.scope) === scope));
      names.push({ name: nm, ref, scope, comment: comment.value || undefined });
      draw();
      return true;
    };
    ui.dialog({ title: '定义名称', width: 380, body: h('div', { class: 'col' }, h('label', { text: '工作簿中的名称(&w)：' }), nameIn, box, ui.field('批注(&m)：', comment), ui.field('引用位置(&R)：', refIn)),
      buttons: [{ label: '添加(&A)', onClick: () => { add(); return false; } }, { label: '删除(&D)', onClick: () => { const q = nameIn.value.trim(); names = names.filter((x) => qualified(w, x) !== q && !(x.scope == null && x.name === q)); nameIn.value = ''; draw(); return false; } },
        { label: '确定', primary: true, onClick: () => { if (nameIn.value.trim() && !names.some((x) => qualified(w, x) === nameIn.value.trim())) { if (add() === false) return false; } tryRun(() => { O.tx(w, '定义名称', () => O.setNames(w, names)); }); } }, { label: '关闭' }] });
  };
  D.pasteName = function () {
    const w = wb();
    const names = visibleNames(w);
    if (!names.length) { ui.msg('此工作簿中没有定义名称。', { icon: 'info' }); return; }
    let pick = qualified(w, names[0]);
    const lb = D.list(names.map((n) => [qualified(w, n), qualified(w, n)]), pick, (v) => { pick = v; }, { translate: 'no' });
    const insert = () => { const ed = E(); if (ed.active) ed.insertText(pick); else { ed.begin('enter', '=' + pick); } };
    const d = ui.dialog({ title: '粘贴名称', width: 300, body: h('div', { class: 'col' }, h('label', { text: '粘贴名称(&n)' }), lb),
      buttons: [{ label: '粘贴列表(&L)', onClick: () => { tryRun(() => { const s = sh(), sel = G().sel(); O.tx(w, '粘贴列表', () => names.forEach((n, i) => { O.put(s, sel.r + i, sel.c, { v: qualified(w, n) }); O.put(s, sel.r + i, sel.c + 1, { v: '=' + F.display(String(n.ref).replace(/^=/, '')) }); })); }); } }, { label: '确定', primary: true, onClick: () => insert() }, { label: '取消' }] });
    lb.ondbl = () => { d.close(1); insert(); };
  };
  D.createNames = function () {
    const s = sh(), w = wb();
    const rg = O.clip(s, G().range());
    const st = { top: typeof s.val(rg.r1, rg.c1 + (rg.c2 > rg.c1 ? 1 : 0)) === 'string', left: typeof s.val(rg.r1 + (rg.r2 > rg.r1 ? 1 : 0), rg.c1) === 'string', bottom: false, right: false };
    ui.dialog({ title: '创建名称', width: 260, body: ui.group('名称来源', ui.check('最上行(&T)', st.top, (v) => { st.top = v; }), ui.check('最左列(&L)', st.left, (v) => { st.left = v; }), ui.check('最下行(&B)', false, (v) => { st.bottom = v; }), ui.check('最右列(&R)', false, (v) => { st.right = v; })),
      buttons: [{ label: '确定', primary: true, onClick: () => tryRun(() => {
        let names = w.names.slice();
        const mk = (label, ref) => { let nm = String(label == null ? '' : label).trim().replace(/\s+/g, '_').replace(/[^\w.\\]/g, ''); if (!nm) return; if (/^\d/.test(nm) || D.validName(nm)) nm = '_' + nm; names = names.filter((x) => !(x.name.toLowerCase() === nm.toLowerCase() && x.scope == null)); names.push({ name: nm, ref: F.quoteSheet(s.name) + '!' + F.absRangeName(ref), scope: null }); };
        const r1 = rg.r1 + (st.top ? 1 : 0), r2 = rg.r2 - (st.bottom ? 1 : 0), c1 = rg.c1 + (st.left ? 1 : 0), c2 = rg.c2 - (st.right ? 1 : 0);
        if (st.top) for (let c = c1; c <= c2; c++) mk(s.val(rg.r1, c), { r1, c1: c, r2, c2: c });
        if (st.bottom) for (let c = c1; c <= c2; c++) mk(s.val(rg.r2, c), { r1, c1: c, r2, c2: c });
        if (st.left) for (let r = r1; r <= r2; r++) mk(s.val(r, rg.c1), { r1: r, c1, r2: r, c2 });
        if (st.right) for (let r = r1; r <= r2; r++) mk(s.val(r, rg.c2), { r1: r, c1, r2: r, c2 });
        O.tx(w, '创建名称', () => O.setNames(w, names));
      }) }, { label: '取消' }] });
  };
  D.applyNames = function () {
    const s = sh(), w = wb();
    const names = visibleNames(w).filter((n) => { try { const a = F.parse(String(n.ref)); return a.t === 'ref' || a.t === 'area'; } catch (e) { return false; } });
    if (!names.length) { ui.msg('没有引用单元格的名称。', { icon: 'info' }); return; }
    const picked = new Set(names.map((n) => n.name));
    const box = h('div', { class: 'lbox checks', translate: 'no' }, ...names.map((n) => ui.check(n.name, true, (v) => { if (v) picked.add(n.name); else picked.delete(n.name); })));
    ui.dialog({ title: '应用名称', width: 300, body: h('div', { class: 'col' }, h('label', { text: '应用名称(&n)：' }), box), buttons: [{ label: '确定', primary: true, onClick: () => tryRun(() => {
      const targets = names.filter((n) => picked.has(n.name)).map((n) => { const a = F.parse(String(n.ref)); return { name: n.name, a, sheet: (a.sheet || s.name).toLowerCase() }; });
      let rg = O.clip(s, G().range());
      if (rg.r1 === rg.r2 && rg.c1 === rg.c2) rg = { r1: 0, c1: 0, r2: s.maxR, c2: s.maxC };
      O.tx(w, '应用名称', () => {
        s.each(rg.r1, rg.c1, rg.r2, rg.c2, (cl, r, c) => {
          if (cl.f == null) return;
          let changed = false;
          const ast = F.map(F.parse(cl.f), (n) => {
            if (n.t !== 'ref' && n.t !== 'area') return null;
            const sheet = (n.sheet || s.name).toLowerCase();
            const hit = targets.find((t) => t.sheet === sheet && t.a.t === n.t && (n.t === 'ref' ? t.a.r === n.r && t.a.c === n.c : t.a.r1 === n.r1 && t.a.c1 === n.c1 && t.a.r2 === n.r2 && t.a.c2 === n.c2));
            if (!hit) return null;
            changed = true;
            return { t: 'name', name: hit.name };
          });
          if (changed) O.setFormula(s, r, c, F.toText(ast, { store: true }));
        });
      });
    }) }, { label: '取消' }] });
  };

  /* ================================================================ Evaluate Formula */
  D.evaluateFormula = function () {
    const s = sh(), w = wb(), sel = G().sel();
    const cell = s.get(sel.r, sel.c);
    if (!cell || cell.f == null) { ui.msg('活动单元格不包含公式。', { icon: 'info' }); return; }
    const ctx = C.ctx(w, s, sel.r, sel.c, !!cell.af);
    let ast = F.parse(cell.f);
    const lit = (v) => {
      if (v instanceof C.Ref || v instanceof C.RefList) return null;
      if (v && v.rows && Array.isArray(v.rows)) return { t: 'arr', rows: v.rows.slice(0, 20).map((r) => r.slice(0, 20).map((x) => lit(x) || { t: 'str', v: '' })) };
      if (typeof v === 'number') return { t: 'num', v };
      if (typeof v === 'string') return { t: 'str', v };
      if (typeof v === 'boolean') return { t: 'bool', v };
      if (M.isErr(v)) return { t: 'err', v: v.e };
      if (v == null) return { t: 'num', v: 0 };
      return null;
    };
    const isLit = (n) => n && /^(num|str|bool|err|arr|missing)$/.test(n.t);
    /** the next node to evaluate: a single-cell reference, or an operator / call whose operands are values */
    const next = (n) => {
      if (!n) return null;
      if (n.t === 'ref' || n.t === 'name' || n.t === 'struct') return n;
      if (n.t === 'area') return null;
      const kidsOf = n.t === 'fn' ? n.args : n.t === 'op' ? [n.a, n.b] : n.t === 'un' || n.t === 'paren' || n.t === 'pct' ? [n.a] : [];
      for (const k of kidsOf) { const x = next(k); if (x) return x; }
      if (n.t === 'fn' && n.args.every((a) => isLit(a) || a.t === 'area' || a.t === 'ref')) return n;
      if ((n.t === 'op' || n.t === 'un' || n.t === 'pct' || n.t === 'paren') && kidsOf.every((k) => isLit(k) || k.t === 'area')) return n;
      return null;
    };
    const text = h('div', { class: 'ev-text', translate: 'no' });
    const render = (mark) => {
      text.textContent = '';
      const full = F.toText(ast);
      if (!mark) { text.textContent = '=' + full; return; }
      const part = F.toText(mark);
      const i = full.indexOf(part);
      text.append(document.createTextNode('=' + full.slice(0, Math.max(0, i))), h('u', { translate: 'no', text: part }), document.createTextNode(i >= 0 ? full.slice(i + part.length) : ''));
    };
    let cur = next(ast);
    render(cur);
    const evalBtn = { label: '求值(&E)', primary: true, onClick: () => {
      if (!cur) { ast = F.parse(cell.f); cur = next(ast); render(cur); return false; }
      let v;
      try { v = C.ev(cur, ctx); } catch (e) { v = M.ERR.VALUE; }
      if (v instanceof C.Ref && v.single) v = C.cellValue(v.sheet, v.r1, v.c1);
      const rep = lit(v) || { t: 'str', v: show(v) };
      const target = cur;
      ast = target === ast ? rep : F.map(ast, (n) => (n === target ? rep : null));
      cur = isLit(ast) ? null : next(ast);
      render(cur);
      if (!cur) d.buttons[0].textContent = L.t('重新开始');
      return false;
    } };
    const d = ui.dialog({ title: '公式求值', width: 520, body: h('div', { class: 'col' }, h('div', { class: 'row', style: 'gap:16px' }, h('div', null, h('div', { class: 'hint', text: '引用：' }), h('b', { translate: 'no', text: F.quoteSheet(s.name) + '!' + F.cellName(sel.r, sel.c, true, true) })), h('div', { style: 'flex:1' }, h('div', { class: 'hint', text: '求值：' }), text)), h('div', { class: 'hint', text: '单击“求值”可显示带下划线表达式的结果。最近的结果以斜体显示。' })),
      buttons: [evalBtn, { label: '关闭' }] });
  };

  /* ================================================================ Watch Window */
  let watchDlg = null;
  D.watchWindow = function (addCurrent) {
    const w = wb();
    if (!w._watches) w._watches = [];
    if (addCurrent) { const s = sh(); for (const rg of G().ranges()) { const r = O.clip(s, rg); for (let rr = r.r1; rr <= Math.min(r.r2, r.r1 + 50); rr++) for (let cc = r.c1; cc <= Math.min(r.c2, r.c1 + 20); cc++) if (!w._watches.some((x) => x.sh === s && x.r === rr && x.c === cc)) w._watches.push({ sh: s, r: rr, c: cc }); } }
    if (watchDlg && document.contains(watchDlg.el)) { watchDlg.draw(); return; }
    const tbl = h('table', { class: 'fr-table' });
    let sel = -1;
    const draw = () => {
      tbl.textContent = '';
      tbl.appendChild(h('tr', null, ...[L.N('工作簿'), L.N('工作表'), L.N('名称'), L.N('单元格'), L.N('值'), L.N('公式')].map((t) => h('th', { text: t }))));
      w._watches.forEach((x, i) => {
        if (!w.sheets.includes(x.sh)) return;
        const cl = x.sh.get(x.r, x.c);
        const tr = h('tr', { class: i === sel ? 'on' : '' }, h('td', { translate: 'no', text: A().books[A().cur].name }), h('td', { translate: 'no', text: x.sh.name }), h('td', { translate: 'no', text: '' }), h('td', { translate: 'no', text: F.cellName(x.r, x.c, true, true) }), h('td', { translate: 'no', text: L.csv.cellText(x.sh, x.r, x.c) }), h('td', { translate: 'no', text: cl && cl.f != null ? '=' + F.display(cl.f) : '' }));
        tr.addEventListener('click', () => { sel = i; draw(); });
        tr.addEventListener('dblclick', () => { const k = w.sheets.indexOf(x.sh); if (k !== w.active) A().activateSheet(k); G().select(x.r, x.c); });
        tbl.appendChild(tr);
      });
    };
    draw();
    const off = L.bus.on ? L.bus.on('changed', () => { if (watchDlg && document.contains(watchDlg.el)) draw(); }) : null;
    void off;
    const d = ui.dialog({ title: '监视窗口', width: 520, modeless: true, body: h('div', { class: 'col' }, h('div', { class: 'row', style: 'gap:6px' }, ui.button('添加监视(&A)...', () => { D.watchWindow(true); }), ui.button('删除监视(&D)', () => { if (sel >= 0) { w._watches.splice(sel, 1); sel = -1; draw(); } })), tbl), buttons: [{ label: '关闭', primary: true }] });
    watchDlg = { el: d.el, draw };
  };

  /* ================================================================ Error Checking */
  const ERRDESC = { '#DIV/0!': '除数为零错误', '#N/A': '值不可用错误', '#NAME?': '无效的名称错误', '#NULL!': '空交集错误', '#NUM!': '数值错误', '#REF!': '无效的单元格引用错误', '#VALUE!': '值错误', '#SPILL!': '溢出区域非空', '#CALC!': '计算错误' };
  const ERRHELP = { '#DIV/0!': '所用公式或函数的除数为零或空单元格。', '#N/A': '公式或函数所需的值不可用。', '#NAME?': '公式包含无法识别的文本。', '#NULL!': '两个区域的交集为空。', '#NUM!': '公式或函数包含无效的数值。', '#REF!': '单元格引用无效。', '#VALUE!': '公式中使用的值的数据类型错误。' };
  D.errorChecking = function () {
    const s = sh(), w = wb();
    const ignored = (w._ignoredErr || (w._ignoredErr = new Set()));
    const list = [];
    s.rows.forEach((row, r) => { if (row) row.cells.forEach((cl, c) => {
      if (!cl) return;
      const k = s.id + ':' + r + ':' + c;
      if (ignored.has(k)) return;
      const v = cl.dirty ? C.cellValue(s, r, c) : cl.v;
      if (cl.f != null && M.isErr(v)) list.push({ r, c, kind: v.e });
      else if (cl.f == null && typeof v === 'string' && /^\s*-?[\d,]+(\.\d+)?\s*$/.test(v) && !(LY.styleOf(s, r, c, cl).nf === '@')) list.push({ r, c, kind: 'numtext' });
    }); });
    const sel = G().sel();
    let i = Math.max(0, list.findIndex((x) => x.r > sel.r || (x.r === sel.r && x.c >= sel.c)));
    if (!list.length) { ui.msg('整个工作表的错误检查已完成。', { icon: 'info' }); return; }
    const head = h('b'), formula = h('div', { class: 'ev-text', translate: 'no' }), title = h('div', { style: 'font-weight:bold' }), desc = h('div');
    const showI = () => {
      const x = list[i];
      G().select(x.r, x.c);
      const cl = s.get(x.r, x.c);
      head.textContent = L.t('单元格 {0} 中的错误：', F.cellName(x.r, x.c));
      formula.textContent = cl && cl.f != null ? '=' + F.display(cl.f) : String(cl ? cl.v : '');
      title.textContent = x.kind === 'numtext' ? L.t('以文本形式存储的数字') : L.t(ERRDESC[x.kind] || L.N('错误'));
      desc.textContent = x.kind === 'numtext' ? L.t('此单元格中的数字设置为文本格式，或前面有单引号。') : L.t(ERRHELP[x.kind] || '');
      fixBtn.textContent = x.kind === 'numtext' ? L.t('转换为数字') : L.t('显示计算步骤...');
    };
    const fixBtn = ui.button('显示计算步骤...', () => {
      const x = list[i];
      if (x.kind === 'numtext') { tryRun(() => O.tx(w, '转换为数字', () => O.enter(s, x.r, x.c, String(s.val(x.r, x.c)).trim()))); list.splice(i, 1); if (!list.length) { d.close(0); ui.msg('整个工作表的错误检查已完成。', { icon: 'info' }); return; } i = Math.min(i, list.length - 1); showI(); }
      else D.evaluateFormula();
    }, { style: 'width:100%' });
    const btn = (label, fn) => ui.button(label, fn, { style: 'width:100%' });
    const d = ui.dialog({ title: '错误检查', width: 480, modeless: true, body: h('div', { class: 'row', style: 'gap:12px;align-items:flex-start' }, h('div', { class: 'col', style: 'flex:1' }, head, formula, title, desc),
      h('div', { class: 'col', style: 'width:170px' }, btn('有关此错误的帮助', () => { L.panes.helpQuery = 'formulas'; L.panes.task.show('help'); }), fixBtn, btn('忽略错误', () => { const x = list[i]; ignored.add(s.id + ':' + x.r + ':' + x.c); list.splice(i, 1); if (!list.length) { d.close(0); ui.msg('整个工作表的错误检查已完成。', { icon: 'info' }); return; } i = Math.min(i, list.length - 1); showI(); }), btn('在编辑栏中编辑', () => { d.close(0); E().begin('edit'); }))),
      buttons: [{ label: '上一个(&P)', onClick: () => { i = (i - 1 + list.length) % list.length; showI(); return false; } }, { label: '下一步(&N)', primary: true, onClick: () => { i = (i + 1) % list.length; showI(); return false; } }, { label: '关闭' }] });
    showI();
  };

  /* ================================================================ Page Setup */
  const HF_PRESETS = ['', '&P', '第 &P 页', '第 &P 页，共 &N 页', '&A', '&F', '&A，第 &P 页', '&F，第 &P 页', '&D', '机密，&D，第 &P 页', '&F, &A', '由 &[User] 编制，&D，第 &P 页'];
  const hfLabel = (code, s) => { if (!code) return L.t('（无）'); const p = L.print.parseHF(code.replace('&[User]', A().opts.userName || ''), { page: 1, pages: 1, sheet: s.name, file: A().books[A().cur].name }); return ['l', 'c', 'r'].map((k) => p[k].map((x) => x.text).join('')).filter(Boolean).join(', '); };
  D.pageSetup = function (tab) {
    const s = sh(), w = wb();
    const p = JSON.parse(JSON.stringify(s.print));
    p.margins = Object.assign({ l: 0.75, r: 0.75, t: 1, b: 1, header: 0.5, footer: 0.5 }, p.margins);
    const idx = w.sheets.indexOf(s);
    const getName = (key) => { const n = w.names.find((x) => x.name.toLowerCase() === key && x.scope === idx); return n ? String(n.ref).replace(/^=/, '').replace(new RegExp('(?:' + F.quoteSheet(s.name).replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ')!', 'g'), '') : ''; };
    let area = getName('_xlnm.print_area');
    const titles = L.print.titles(s);
    /* page */
    const adj = ui.spin({ value: p.scale || 100, min: 10, max: 400, step: 5, dec: 0, fmt: (v) => L.t('正常大小的 {0}%', Math.round(v)), onChange: (v) => { p.scale = Math.round(v); } });
    const fitW = ui.spin({ value: p.fitW == null ? 1 : p.fitW, min: 0, max: 32767, step: 1, dec: 0, onChange: (v) => { p.fitW = Math.round(v) || null; } });
    const fitH = ui.spin({ value: p.fitH == null ? 1 : p.fitH, min: 0, max: 32767, step: 1, dec: 0, onChange: (v) => { p.fitH = Math.round(v) || null; } });
    const first = h('input', { type: 'text', value: p.firstPage != null ? String(p.firstPage) : L.t('自动'), style: 'width:70px' });
    const pageTab = h('div', { class: 'col' },
      ui.group('方向', h('div', { class: 'row', style: 'gap:24px' }, ui.radio('ps-o', '纵向(&t)', p.orientation !== 'landscape', () => { p.orientation = 'portrait'; }), ui.radio('ps-o', '横向(&L)', p.orientation === 'landscape', () => { p.orientation = 'landscape'; }))),
      ui.group('缩放', h('div', { class: 'row', style: 'gap:8px' }, ui.radio('ps-s', '缩放比例(&A)：', !p.fit, () => { p.fit = false; }), adj), h('div', { class: 'row', style: 'gap:8px;flex-wrap:wrap' }, ui.radio('ps-s', '调整为(&F)：', !!p.fit, () => { p.fit = true; }), fitW, h('span', { text: '页宽，乘以' }), fitH, h('span', { text: '页高' }))),
      ui.field('纸张大小(&z)：', ui.select(Object.keys(L.print.PAPERS).map((k) => [k, L.t('{0}（{1} × {2} 英寸）', L.t(L.print.PAPERS[k][0]), L.print.PAPERS[k][1], L.print.PAPERS[k][2])]), p.paper || 1, (v) => { p.paper = +v; })),
      ui.field('打印质量(&q)：', ui.select(['600 dpi', '300 dpi', '150 dpi'], '600 dpi', null, { translate: 'no' })),
      ui.field('起始页码(&r)：', first));
    /* margins */
    const mspin = (k) => ui.spin({ value: p.margins[k], min: 0, max: 20, step: 0.25, unit: '', onChange: (v) => { p.margins[k] = v; } });
    const marTab = h('div', { class: 'col' }, h('div', { class: 'ps-margins' }, ui.field('顶端(&T)：', mspin('t')), ui.field('页眉(&e)：', mspin('header')), ui.field('左(&L)：', mspin('l')), ui.field('右(&R)：', mspin('r')), ui.field('底端(&B)：', mspin('b')), ui.field('页脚(&F)：', mspin('footer'))),
      ui.group('居中方式', ui.check('水平(&z)', !!p.hCenter, (v) => { p.hCenter = v; }), ui.check('垂直(&V)', !!p.vCenter, (v) => { p.vCenter = v; })), h('div', { class: 'hint', text: '页边距单位为英寸。' }));
    /* header / footer */
    /* the headers and footers are the sheet's own text: shown as they print */
    const hfSel = (key) => { const opts = HF_PRESETS.slice(); if (p[key] && !opts.includes(p[key])) opts.unshift(p[key]); const sel = ui.select(opts.map((c) => [c, hfLabel(c, s)]), p[key] || '', (v) => { p[key] = v; prevH.textContent = hfLabel(p.header, s); prevF.textContent = hfLabel(p.footer, s); }, { translate: 'no' }); return sel; };
    const prevH = h('div', { class: 'hf-prev', translate: 'no', text: hfLabel(p.header, s) }), prevF = h('div', { class: 'hf-prev', translate: 'no', text: hfLabel(p.footer, s) });
    const hSel = hfSel('header'), fSel = hfSel('footer');
    const custom = (key) => D.customHF(p[key] || '', key === 'header' ? L.N('页眉') : L.N('页脚'), (code) => { p[key] = code; const sel = key === 'header' ? hSel : fSel; if (!Array.from(sel.options).some((o) => o.value === code)) sel.insertBefore(h('option', { value: code, translate: 'no', text: hfLabel(code, s) }), sel.firstChild); sel.value = code; prevH.textContent = hfLabel(p.header, s); prevF.textContent = hfLabel(p.footer, s); });
    const hfTab = h('div', { class: 'col' }, prevH, ui.field('页眉(&a)：', hSel), h('div', { class: 'row', style: 'gap:8px' }, ui.button('自定义页眉(&C)...', () => custom('header')), ui.button('自定义页脚(&u)...', () => custom('footer'))), ui.field('页脚(&F)：', fSel), prevF,
      ui.check('奇偶页不同(&o)', !!p.diffOddEven, (v) => { p.diffOddEven = v; }), ui.check('首页不同(&p)', !!p.diffFirst, (v) => { p.diffFirst = v; }), ui.check('随文档缩放', p.hfScale !== false, (v) => { p.hfScale = v; }), ui.check('与页边距对齐', p.hfAlign !== false, (v) => { p.hfAlign = v; }));
    /* sheet */
    const areaIn = D.refBox(area, 220), rowsIn = D.refBox(titles.rows ? '$' + (titles.rows.r1 + 1) + ':$' + (titles.rows.r2 + 1) : '', 160), colsIn = D.refBox(titles.cols ? '$' + F.colName(titles.cols.c1) + ':$' + F.colName(titles.cols.c2) : '', 160);
    const sheetTab = h('div', { class: 'col' }, ui.field('打印区域(&a)：', areaIn), ui.group('打印标题', ui.field('顶端标题行(&R)：', rowsIn), ui.field('左端标题列(&C)：', colsIn)),
      ui.group('打印', h('div', { class: 'ps-print' }, ui.check('网格线(&G)', !!p.gridLines, (v) => { p.gridLines = v; }), ui.check('黑白(&B)', !!p.bw, (v) => { p.bw = v; }), ui.check('草稿品质(&q)', !!p.draft, (v) => { p.draft = v; }), ui.check('行列标题(&m)', !!p.headings, (v) => { p.headings = v; }),
        ui.field('批注(&m)：', ui.select([['none', '（无）'], ['atEnd', '工作表末尾'], ['asDisplayed', '如工作表所示']], p.comments || 'none', (v) => { p.comments = v; })), ui.field('单元格错误显示为(&e)：', ui.select([['displayed', '已显示'], ['blank', '<空白>'], ['dash', '--'], ['NA', '#N/A']], p.errors || '已显示', (v) => { p.errors = v; })))),
      ui.group('打印顺序', ui.radio('ps-po', '先列后行(&D)', p.pageOrder !== 'overThenDown', () => { p.pageOrder = 'downThenOver'; }), ui.radio('ps-po', '先行后列(&v)', p.pageOrder === 'overThenDown', () => { p.pageOrder = 'overThenDown'; })));
    const apply = () => {
      p.firstPage = /^\d+$/.test(first.value.trim()) ? +first.value.trim() : null;
      const names = w.names.filter((n) => !(n.scope === idx && /^_xlnm\.(print_area|print_titles)$/i.test(n.name)));
      const q = F.quoteSheet(s.name);
      const a = areaIn.input.value.trim().replace(/^=/, '');
      if (a) { const parts = a.split(',').map((x) => x.trim()).filter(Boolean).map((x) => (/!/.test(x) ? x : q + '!' + x)); names.push({ name: '_xlnm.Print_Area', ref: parts.join(','), scope: idx }); }
      const tparts = [];
      const rr = rowsIn.input.value.trim().replace(/^=/, ''), cc = colsIn.input.value.trim().replace(/^=/, '');
      if (cc) tparts.push(/!/.test(cc) ? cc : q + '!' + cc);
      if (rr) tparts.push(/!/.test(rr) ? rr : q + '!' + rr);
      if (tparts.length) names.push({ name: '_xlnm.Print_Titles', ref: tparts.join(','), scope: idx });
      try { for (const n of names) if (/^_xlnm/.test(n.name)) F.parse(n.ref); } catch (e) { ui.msg('引用无效。', { icon: 'warn' }); return false; }
      tryRun(() => O.tx(w, '页面设置', () => { O.setPrint(s, p); O.setNames(w, names); }));
      L.print.computeBreaks(s);
      return true;
    };
    return new Promise((res) => {
      ui.dialog({ title: '页面设置', width: 470, body: ui.tabs([{ label: '页面', body: pageTab }, { label: '页边距', body: marTab }, { label: '页眉/页脚', body: hfTab }, { label: '工作表', body: sheetTab }], tab || 0),
        buttons: [{ label: '打印(&P)...', onClick: () => { if (apply() === false) return false; setTimeout(() => D.print(), 0); } }, { label: '打印预览(&w)', onClick: () => { if (apply() === false) return false; setTimeout(() => L.print.preview(), 0); } }, { label: '确定', primary: true, onClick: () => apply() }, { label: '取消' }] }).done.then(res);
    });
  };
  /** Custom Header / Footer: three sections with the field buttons */
  D.customHF = function (code, kind, done) {
    const secs = { l: '', c: '', r: '' };
    /* split the code into its &L / &C / &R parts, keeping the codes */
    let cur = 'c';
    String(code).split(/(&[LCR])/).forEach((part) => { if (/^&[LCR]$/.test(part)) cur = part[1].toLowerCase(); else secs[cur] += part; });
    const areas = { l: h('textarea', { rows: '5' }), c: h('textarea', { rows: '5' }), r: h('textarea', { rows: '5' }) };
    for (const k in areas) areas[k].value = secs[k];
    let last = areas.c;
    for (const k in areas) areas[k].addEventListener('focus', () => { last = areas[k]; });
    const ins = (t) => { const a = last; const s0 = a.selectionStart, s1 = a.selectionEnd; a.value = a.value.slice(0, s0) + t + a.value.slice(s1); a.focus(); a.setSelectionRange(s0 + t.length, s0 + t.length); };
    const fb = (label, t) => h('button', { type: 'button', class: 'btn small', text: label, onclick: () => ins(t) });
    const font = () => { const f = h('input', { type: 'text', value: 'Arial', style: 'width:110px' }); const st = ui.select([['Regular', L.N('常规')], ['Bold', L.N('加粗')], ['Italic', L.N('倾斜')], ['Bold Italic', '加粗、倾斜']], 'Regular'); const sz = h('input', { type: 'text', value: '10', style: 'width:40px' }); ui.dialog({ title: '字体', width: 320, body: h('div', { class: 'row', style: 'gap:8px' }, ui.field('字体：', f), ui.field('样式：', st), ui.field('大小：', sz)), buttons: [{ label: '确定', primary: true, onClick: () => ins(`&"${f.value},${st.value}"&${parseInt(sz.value, 10) || 10}`) }, { label: '取消' }] }); };
    ui.dialog({ title: kind, width: 560, body: h('div', { class: 'col' }, h('div', { class: 'hint', text: '设置文字格式：选择文字后单击字体按钮。插入页码、日期、时间、文件路径、文件名或工作表标签名：在编辑框中定位插入点，再单击相应按钮。' }),
      h('div', { class: 'row', style: 'gap:4px;flex-wrap:wrap' }, h('button', { type: 'button', class: 'btn small', text: 'A（字体）', onclick: font }), fb('页码', '&P'), fb(L.N('页'), '&N'), fb(L.N('日期'), '&D'), fb(L.N('时间'), '&T'), fb('路径和文件名', '&Z&F'), fb(L.N('文件'), '&F'), fb(L.N('制表符'), '&A')),
      h('div', { class: 'hf-secs' }, ui.field('左侧部分(&L)：', areas.l), ui.field('中间部分(&C)：', areas.c), ui.field('右侧部分(&R)：', areas.r))),
    buttons: [{ label: '确定', primary: true, onClick: () => { const out = (areas.l.value ? '&L' + areas.l.value : '') + (areas.c.value ? '&C' + areas.c.value : '') + (areas.r.value ? '&R' + areas.r.value : ''); done(out); } }, { label: '取消' }] });
  };

  /* ================================================================ Print */
  D.print = function () {
    const st = { what: 'sheets', range: 'all', from: 1, to: 1, copies: 1 };
    const from = h('input', { type: 'text', value: '1', style: 'width:40px' }), to = h('input', { type: 'text', value: '1', style: 'width:40px' });
    const items = L.print.allPages(A().selectedSheets(), null);
    to.value = String(Math.max(1, items.length));
    return new Promise((res) => {
      ui.dialog({ title: '打印', width: 440, body: h('div', { class: 'col' },
        ui.group('打印机', ui.field('名称(&N)：', ui.select([['pdf', 'Ledger PDF 生成器']], 'pdf')), h('div', { class: 'hint', text: 'Ledger 通过创建 PDF 进行打印，您可以打开该 PDF 并使用任意打印机打印。' })),
        h('div', { class: 'row', style: 'gap:10px;align-items:stretch' },
          ui.group('打印范围', ui.radio('pr-r', '全部(&A)', true, () => { st.range = 'all'; }), h('div', { class: 'row', style: 'gap:6px' }, ui.radio('pr-r', '页(&g)', false, () => { st.range = 'pages'; }), h('span', { text: '从：' }), from, h('span', { text: '到：' }), to)),
          ui.group('打印内容', ui.radio('pr-w', '选区(&n)', false, () => { st.what = 'selection'; }), ui.radio('pr-w', '选定工作表(&h)', true, () => { st.what = 'sheets'; }), ui.radio('pr-w', '整个工作簿(&E)', false, () => { st.what = 'workbook'; }))),
        ui.field('打印份数(&c)：', ui.spin({ value: 1, min: 1, max: 99, step: 1, dec: 0 }))),
      buttons: [{ label: '预览(&v)', onClick: () => { setTimeout(() => L.print.preview(), 0); } }, { label: '确定', primary: true, onClick: () => { const o = { what: st.what }; if (st.range === 'pages') { o.from = parseInt(from.value, 10) || 1; o.to = parseInt(to.value, 10) || o.from; } L.print.exportPDF(o); } }, { label: '取消' }] }).done.then(res);
    });
  };
})(typeof window !== 'undefined' ? window : globalThis);
