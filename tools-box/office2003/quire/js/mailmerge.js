/* Quire — mail merge: data sources (.csv/.txt/.docx tables or a typed list), recipient filtering,
 * address block / greeting line / merge fields, preview, and merging to a new document or PDF.
 */
(function () {
  'use strict';
  const L = window.L, D = L.D, O = L.O, E = L.ed, LY = L.layout;
  const { h } = L;
  const ui = L.ui;
  const MM = (L.mailmerge = { data: null, rec: 0, preview: false, type: 'letters', wstep: 1 });
  const doc = () => D.doc;
  const A = () => L.app;
  const norm = (s) => String(s).toLowerCase().replace(/[^a-z0-9]/g, '');
  const ADDRESS_FIELDS = ['Title', 'First Name', 'Last Name', 'Company Name', 'Address Line 1', 'Address Line 2', 'City', 'State', 'ZIP Code', 'Country', 'Home Phone', 'Work Phone', 'E-mail Address'];

  /* ---------- data sources ---------- */
  function parseDelimited(text) {
    text = text.replace(/^﻿/, '');
    const first = text.split(/\r?\n/)[0] || '';
    const delim = (first.match(/\t/g) || []).length > (first.match(/,/g) || []).length ? '\t' : (first.match(/;/g) || []).length > (first.match(/,/g) || []).length ? ';' : ',';
    const rows = [];
    let row = [], cell = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (q) { if (c === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; } else cell += c; continue; }
      if (c === '"' && !cell) { q = true; continue; }
      if (c === delim) { row.push(cell); cell = ''; continue; }
      if (c === '\n' || c === '\r') { if (c === '\r' && text[i + 1] === '\n') i++; row.push(cell); cell = ''; if (row.some((x) => x !== '')) rows.push(row); row = []; continue; }
      cell += c;
    }
    row.push(cell);
    if (row.some((x) => x !== '')) rows.push(row);
    return rows;
  }
  function setData(name, rows) {
    if (!rows.length) { ui.msg('数据源不包含记录。', { icon: 'warn' }); return false; }
    const fields = rows[0].map((f, i) => String(f).trim() || 'Field' + (i + 1));
    MM.data = { name, fields, rows: rows.slice(1).map((r) => fields.map((_, i) => (r[i] == null ? '' : String(r[i])))), included: null };
    MM.data.included = new Set(MM.data.rows.map((_, i) => i));
    MM.rec = 0;
    A().opts.toolbars.mailMerge = true;
    A().updateToolbars();
    refreshPane();
    return true;
  }
  MM.openSource = async function () {
    const files = await L.pickFiles('.csv,.txt,.tsv,.docx,.htm,.html');
    const f = files[0];
    if (!f) return;
    try {
      if (/\.docx$/i.test(f.name)) {
        const res = await L.docx.read(await L.readAsArrayBuffer(f));
        const t = res.doc.main.blocks.find((b) => b.t === 'tbl');
        let rows;
        if (t) rows = t.rows.map((r) => r.cells.map((c) => c.blocks.filter((b) => b.t === 'p').map((p) => D.plainText(p)).join(' ').trim()));
        else rows = res.doc.main.blocks.filter((b) => b.t === 'p' && D.plainText(b).trim()).map((p) => D.plainText(p).split('\t'));
        setData(f.name, rows);
      } else if (/\.html?$/i.test(f.name)) {
        const dom = new DOMParser().parseFromString(await L.readAsText(f), 'text/html');
        const t = dom.querySelector('table');
        if (!t) { ui.msg('该文件中未找到表格。', { icon: 'warn' }); return; }
        setData(f.name, Array.from(t.rows).map((r) => Array.from(r.cells).map((c) => c.textContent.trim())));
      } else setData(f.name, parseDelimited(await L.readAsText(f)));
      MM.recipientsDialog();
    } catch (e) { ui.msg(L.t('无法打开数据源。{0}', e.message || ''), { icon: 'error' }); }
  };
  /** Type a New Address List */
  MM.typeList = function () {
    const fields = ADDRESS_FIELDS.slice();
    const rows = MM.data && MM.data.typed ? MM.data.rows.map((r) => r.slice()) : [fields.map(() => '')];
    let cur = 0;
    const inputs = fields.map((f, i) => { const inp = h('input', { type: 'text', id: 'mm-f' + i, style: 'width:100%' }); inp.addEventListener('input', () => { rows[cur][i] = inp.value; }); return inp; });
    /* the fields' names are the data's own (they become MERGEFIELD names): not translated */
    const grid = h('div', { style: 'display:grid;grid-template-columns:120px 1fr;gap:3px 8px;max-height:300px;overflow:auto;padding-right:4px', translate: 'no' });
    fields.forEach((f, i) => grid.append(h('label', { for: 'mm-f' + i }, document.createTextNode(f)), inputs[i]));
    const info = h('span');
    const load = () => { inputs.forEach((inp, i) => { inp.value = rows[cur][i] || ''; }); info.textContent = L.t('第 {0} 项，共 {1} 项', cur + 1, rows.length); };
    const nav = h('div', { class: 'tp-btns' },
      ui.button('新建词条(&N)', () => { rows.push(fields.map(() => '')); cur = rows.length - 1; load(); inputs[0].focus(); }, { class: 'btn small' }),
      ui.button('删除词条(&D)', () => { if (rows.length > 1) { rows.splice(cur, 1); cur = Math.min(cur, rows.length - 1); load(); } }, { class: 'btn small' }),
      ui.button('<', () => { if (cur > 0) { cur--; load(); } }, { class: 'btn small' }), ui.button('>', () => { if (cur < rows.length - 1) { cur++; load(); } }, { class: 'btn small' }), info);
    load();
    ui.dialog({
      title: '新建地址列表', width: 440, body: h('div', { class: 'col' }, grid, nav), buttons: [{ label: '确定', primary: true, onClick: () => { const data = [fields].concat(rows.filter((r) => r.some(Boolean))); if (setData(L.t('地址列表'), data)) MM.data.typed = true; } }, { label: '取消' }],
    });
  };
  MM.recipientsDialog = function () {
    const d = MM.data;
    if (!d) return;
    let sortCol = -1, desc = false;
    const box = h('div', { class: 'datasheet', style: 'max-height:280px' });
    const draw = () => {
      L.clear(box);
      const order = d.rows.map((_, i) => i);
      if (sortCol >= 0) order.sort((a, b) => { const r = String(d.rows[a][sortCol]).localeCompare(String(d.rows[b][sortCol]), undefined, { numeric: true }); return desc ? -r : r; });
      const t = h('table', { class: 'mm-table', translate: 'no' });   // the data's own fields and values
      const hr = h('tr', null, h('th', null, (() => { const c = h('input', { type: 'checkbox', 'aria-label': '全部包含' }); c.checked = d.included.size === d.rows.length; c.addEventListener('change', () => { d.included = c.checked ? new Set(d.rows.map((_, i) => i)) : new Set(); draw(); }); return c; })()));
      d.fields.forEach((f, i) => { const th = h('th', { style: 'cursor:pointer' }, document.createTextNode(f + (sortCol === i ? (desc ? ' ▼' : ' ▲') : ''))); th.addEventListener('click', () => { if (sortCol === i) desc = !desc; else { sortCol = i; desc = false; } draw(); }); hr.appendChild(th); });
      t.appendChild(hr);
      for (const i of order) {
        const c = h('input', { type: 'checkbox', 'aria-label': '包含' });
        c.checked = d.included.has(i);
        c.addEventListener('change', () => { if (c.checked) d.included.add(i); else d.included.delete(i); });
        t.appendChild(h('tr', null, h('td', null, c), ...d.rows[i].map((v) => h('td', null, document.createTextNode(v)))));
      }
      box.appendChild(t);
    };
    draw();
    ui.dialog({ title: '邮件合并收件人', width: 620, body: h('div', { class: 'col' }, h('div', { text: L.t('数据源：{0}。单击列标题可排序；清除复选框可排除相应收件人。', d.name) }), box), buttons: [{ label: '确定', primary: true, onClick: () => { MM.order = null; refreshPane(); } }] });
  };
  MM.records = () => (MM.data ? MM.data.rows.map((r, i) => i).filter((i) => MM.data.included.has(i)) : []);
  MM.value = function (name) {
    const d = MM.data;
    if (!d) return '';
    const recs = MM.records();
    const row = d.rows[recs[L.clamp(MM.rec, 0, recs.length - 1)]];
    if (!row) return '';
    const k = norm(name);
    let i = d.fields.findIndex((f) => norm(f) === k);
    if (i < 0) { const alias = { firstname: ['first', 'fname', 'givenname'], lastname: ['last', 'lname', 'surname', 'familyname'], addressline1: ['address', 'address1', 'street'], zipcode: ['zip', 'postalcode', 'postcode'], state: ['province', 'region'], companyname: ['company', 'organization'] }[k] || []; i = d.fields.findIndex((f) => alias.includes(norm(f))); }
    return i >= 0 ? row[i] : '';
  };
  const fieldFor = (want) => { const d = MM.data; if (!d) return null; const k = norm(want); const f = d.fields.find((x) => norm(x) === k) || d.fields.find((x) => norm(x).includes(k)); return f || null; };
  const mf = (name) => name.replace(/\s+/g, '_');

  /* ---------- inserting fields ---------- */
  function insertParts(parts, label) {
    E.edit(label, () => {
      let pos = E.deleteSelection();
      for (const part of parts) {
        if (part === '\n') { pos = O.splitPara(pos); continue; }
        if (typeof part === 'string') { pos = O.insertText(pos, part, O.inheritRPr(pos)); continue; }
        const instr = `MERGEFIELD ${mf(part.f)}`;
        pos = O.insertField(pos, instr, L.fields.evaluate(instr, pos), O.inheritRPr(pos));
      }
      return pos;
    });
  }
  MM.insertField = (name) => insertParts([{ f: name }], '插入合并域');
  MM.insertBlock = function (kind) {
    if (!MM.data) return;
    if (kind === 'greeting') {
      const parts = ['Dear '];
      const t = fieldFor('Title'), fn = fieldFor('First Name'), ln = fieldFor('Last Name');
      if (t) parts.push({ f: t }, ' ');
      if (fn) parts.push({ f: fn }, ' ');
      if (ln) parts.push({ f: ln });
      parts.push(',');
      insertParts(parts, '插入问候语');
      return;
    }
    const parts = [];
    const add = (names, sep) => { const fs = names.map(fieldFor).filter(Boolean); if (!fs.length) return false; fs.forEach((f, i) => { if (i) parts.push(sep || ' '); parts.push({ f }); }); return true; };
    if (add(['Title', 'First Name', 'Last Name'])) parts.push('\n');
    if (add(['Company Name'])) parts.push('\n');
    if (add(['Address Line 1'])) parts.push('\n');
    if (add(['Address Line 2'])) parts.push('\n');
    const city = fieldFor('City'), st = fieldFor('State'), zip = fieldFor('ZIP Code');
    if (city) { parts.push({ f: city }); if (st || zip) parts.push(', '); }
    if (st) { parts.push({ f: st }); if (zip) parts.push(' '); }
    if (zip) parts.push({ f: zip });
    if (!parts.length) { MM.fieldDialog(); return; }
    insertParts(parts, '插入地址块');
  };
  MM.fieldDialog = function () {
    if (!MM.data) return;
    const sel = h('select', { size: 10, id: 'mm-fields', style: 'width:100%', translate: 'no' });
    for (const f of MM.data.fields) sel.appendChild(h('option', { value: f }, document.createTextNode(f)));
    sel.selectedIndex = 0;
    sel.addEventListener('dblclick', () => { MM.insertField(sel.value); });
    ui.dialog({ title: '插入合并域', width: 300, body: h('div', { class: 'col' }, h('label', { for: 'mm-fields', text: '域：' }), sel), buttons: [{ label: '插入(&I)', primary: true, onClick: () => { MM.insertField(sel.value); return false; } }, { label: '关闭' }] });
  };

  /* ---------- preview ---------- */
  function refreshFields() {
    if (!L.fields) return;
    const d = doc();
    const list = [];
    for (const st of D.stories(d)) D.walk(st, (b, c, i, story) => { if (b.t === 'p') for (const it of b.runs) if (it.t === 'fb' && /^\s*MERGEFIELD\b/i.test(it.instr)) list.push({ fid: it.fid, it, p: b, story }); });
    if (!list.length) return;
    const dirty = d.dirty;
    const n0 = D.hist.undo.length;
    D.tx('预览结果', () => { for (const fr of list) L.fields.updateField(fr); });
    if (D.hist.undo.length > n0) D.hist.undo.pop();
    d.dirty = dirty;
  }
  MM.togglePreview = function () { MM.preview = !MM.preview; refreshFields(); refreshPane(); L.ui.refresh(); };
  MM.go = function (i) { const n = MM.records().length; if (!n) return; MM.rec = i < 0 ? n - 1 : L.clamp(i, 0, n - 1); if (!MM.preview) MM.preview = true; refreshFields(); refreshPane(); L.ui.refresh(); };
  MM.step = (dlt) => MM.go(MM.rec + dlt);

  /* ---------- merging ---------- */
  /** blocks for the current record with merge fields replaced by plain text */
  function mergedBlocks(srcBlocks) {
    const blocks = D.cloneBlocks(srcBlocks);
    D.walk({ blocks }, (b) => {
      if (b.t !== 'p') return;
      const out = [];
      let skip = null;
      for (const it of b.runs) {
        if (skip) { if (it.t === 'fe' && it.fid === skip) skip = null; continue; }
        if (it.t === 'fb' && /^\s*MERGEFIELD\b/i.test(it.instr)) {
          const f = L.fields.parse(it.instr);
          const v = MM.value(f.args[0] || '');
          const txt = (v && f.sw['\\b'] ? f.sw['\\b'] : '') + v + (v && f.sw['\\f'] ? f.sw['\\f'] : '');
          if (txt) out.push(D.text(L.fields.applyGeneral(txt, f.fmt.filter((x) => !/MERGEFORMAT/i.test(x))), L.clone(it.rPr || {})));
          skip = it.fid;
          continue;
        }
        if (it.t === 'fb' && /^\s*(NEXT|NEXTIF|SKIPIF|MERGEREC|MERGESEQ)\b/i.test(it.instr)) { skip = it.fid; if (/MERGEREC|MERGESEQ/i.test(it.instr)) out.push(D.text(String(MM.rec + 1), L.clone(it.rPr || {}))); continue; }
        out.push(it);
      }
      b.runs = out;
      D.normalize(b);
      /* address lines left empty by missing data are removed (Word's blank-line suppression) */
    });
    return blocks.filter((b) => !(b.t === 'p' && b._wasAddr && !D.plen(b)));
  }
  MM.buildMerged = function () {
    const src = doc();
    const recs = MM.records();
    if (!recs.length) return null;
    const out = D.newDoc();
    out.styles = L.clone(src.styles); out.numbering = L.clone(src.numbering); out.defaults = L.clone(src.defaults); out.settings = L.clone(src.settings); out.settings.track = false;
    out.hf = L.clone(src.hf); out.theme = L.clone(src.theme); out.sect = L.clone(src.sect); out.props = Object.assign({}, src.props, { title: (src.props.title || A().fileName) + ' (merged)' });
    out.main.blocks = [];
    const save = MM.rec;
    recs.forEach((ri, k) => {
      MM.rec = k;
      const bl = mergedBlocks(src.main.blocks);
      if (MM.type !== 'directory' && k < recs.length - 1) {
        /* each letter becomes its own section */
        const last = bl[bl.length - 1].t === 'p' ? bl[bl.length - 1] : (bl.push(D.para()), bl[bl.length - 1]);
        last.sect = Object.assign(L.opc.duplicate(src.sect), { type: 'nextPage' });
      }
      out.main.blocks.push(...bl);
    });
    MM.rec = save;
    return out;
  };
  MM.mergeToNew = async function () {
    const d = MM.buildMerged();
    if (!d) { ui.msg('未选择收件人。', { icon: 'info' }); return; }
    A().untitled++;
    MM.preview = false;
    refreshFields();
    A().loadDoc(d, 'Letters' + A().untitled, { newWindow: true });
    MM.preview = false;
    const n = MM.records().length;
    A().status(n === 1 ? L.t('已将 {0} 条记录合并到新文档。', n) : L.t('已将 {0} 条记录合并到新文档。', n));
  };
  MM.mergeToPDF = async function () {
    const merged = MM.buildMerged();
    if (!merged) return;
    const keep = D.doc, keepHist = { u: D.hist.undo.slice(), r: D.hist.redo.slice() };
    try {
      D.doc = merged;
      LY.render();
      const blob = await A().buildPDF();
      L.ui.busy(false);
      D.doc = keep;
      D.hist.undo = keepHist.u; D.hist.redo = keepHist.r;
      LY.render();
      E.restoreDom(true);
      await L.saveFile(A().fileName + ' (merged).pdf', blob);
    } catch (e) { D.doc = keep; LY.render(); L.ui.busy(false); ui.msg(L.t('无法打印合并结果。{0}', e.message || ''), { icon: 'error' }); }
  };

  /* ---------- wizard (task pane) ---------- */
  /* the wizard's six steps */
  const STEPS = [L.N('选择文档类型'), L.N('开始文档'), L.N('选择收件人'), L.N('撰写信函'), L.N('预览信函'), L.N('完成合并')];
  function refreshPane() { if (L.panes && L.panes.task.current === 'mailmerge') L.panes.task.render(); }
  MM.renderWizard = function (b) {
    const { sec, link } = L.panes;
    const step = MM.wstep;
    const nav = h('div', { class: 'tp-btns', style: 'margin-top:14px' }, h('div', { class: 'tp-note', text: L.t('第 {0} 步，共 6 步', step) }));
    if (step < 6) nav.appendChild(link('nextChange', L.t('下一步：{0}', L.t(STEPS[step])), () => { MM.wstep++; if (MM.wstep === 5 && MM.data) MM.go(MM.rec); refreshPane(); }, { disabled: step === 3 && !MM.data }));
    if (step > 1) nav.appendChild(link('prevChange', L.t('上一步：{0}', L.t(STEPS[step - 2])), () => { MM.wstep--; refreshPane(); }));
    if (step === 1) {
      const rad = (v, l) => ui.radio('mm-type', l, MM.type === v, () => { MM.type = v; });
      b.append(sec('选择文档类型', h('div', { class: 'tp-note', text: '您正在处理哪种类型的文档？' }), rad('letters', '信函(&L)'), rad('envelopes', '信封(&E)'), rad('labels', '标签(&b)'), rad('directory', '目录(&D)')), h('div', { class: 'tp-note', text: '信函：向一组收件人发送信函，可为每位收件人定制信函内容。' }));
    } else if (step === 2) {
      b.append(sec('选择开始文档', h('div', { class: 'tp-note', text: '您希望如何设置信函？' }), ui.radio('mm-start', '使用当前文档(&d)', true, () => {}), ui.radio('mm-start', '从模板开始(&t)', false, () => L.dlg.newFromTemplate()), ui.radio('mm-start', '从现有文档开始(&c)', false, () => ui.exec('open'))));
      if (MM.type === 'envelopes') b.appendChild(link('envelope', '信封选项...', () => L.dlg.envelopes()));
      if (MM.type === 'labels') b.appendChild(link('labels', '标签选项...', () => L.dlg.envelopes('labels')));
    } else if (step === 3) {
      b.append(sec('选择收件人', ui.radio('mm-src', '使用现有列表(&l)', !MM.data || !MM.data.typed, () => {}), ui.radio('mm-src', '键入新列表(&y)', !!(MM.data && MM.data.typed), () => MM.typeList()),
        MM.data ? h('div', { class: 'tp-note', text: L.t('当前收件人来源：{0}（已选择 {1} 位，共 {2} 位）', MM.data.name, MM.records().length, MM.data.rows.length) }) : h('div', { class: 'tp-note', text: '使用文件或数据库中的姓名和地址。' }),
        link('open', '浏览...', () => MM.openSource()), MM.data ? link('mmRecipients', '编辑收件人列表...', () => MM.recipientsDialog()) : null, link('new', '键入新列表...', () => MM.typeList())));
    } else if (step === 4) {
      b.append(sec('撰写信函', h('div', { class: 'tp-note', text: '请先撰写信函，再单击文档中的位置并选择以下项目之一。' }), link('mmAddress', '地址块...', () => MM.insertBlock('address')), link('mmGreeting', '问候语...', () => MM.insertBlock('greeting')), link('mmField', '更多项目...', () => MM.fieldDialog())));
    } else if (step === 5) {
      const n = MM.records().length;
      b.append(sec('预览信函', h('div', { class: 'tp-note', text: '此处预览一封合并后的信函。要预览其他信函，请单击以下按钮之一：' }), h('div', { class: 'tp-row' }, ui.button('<<', () => MM.step(-1), { class: 'btn small' }), h('span', { text: L.t('收件人：第 {0} 位，共 {1} 位', MM.rec + 1, n) }), ui.button('>>', () => MM.step(1), { class: 'btn small' })), link('find', '查找收件人...', async () => { const q = await ui.prompt('查找：', '', '查找条目'); if (!q) return; const recs = MM.records(); const k = recs.findIndex((ri) => MM.data.rows[ri].some((v) => String(v).toLowerCase().includes(q.toLowerCase()))); if (k >= 0) MM.go(k); }), link('mmRecipients', '编辑收件人列表...', () => MM.recipientsDialog()), ui.button('排除此收件人', () => { const recs = MM.records(); MM.data.included.delete(recs[MM.rec]); MM.go(MM.rec); }, { class: 'btn small' })));
    } else {
      b.append(sec('完成合并', h('div', { class: 'tp-note', text: 'Quire 已准备好生成信函。' }), link('mmMergePrint', '打印... (PDF)', () => MM.mergeToPDF()), link('mmMergeNew', '编辑单个信函...', () => MM.mergeToNew())));
    }
    b.appendChild(nav);
  };
})();
