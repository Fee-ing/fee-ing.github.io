/* Quire — dialogs, part 2: Insert menu (Break, Page Numbers, Date and Time, Field, Symbol,
 * Footnote, Caption, Cross-reference, Index and Tables, Bookmark, Hyperlink, Object, Diagram) and
 * the Table menu dialogs.
 */
(function () {
  'use strict';
  const L = window.L, D = L.D, O = L.O, E = L.ed, LY = L.layout, R = L.R;
  const { h } = L;
  const ui = L.ui;
  const G = L.dlg;
  const doc = () => D.doc;
  const A = () => L.app;
  const F = () => L.fields;

  /* ================= Break ================= */
  G.breakDlg = function () {
    const r = G.radios('br', [['page', '分页符(&P)'], ['column', '分栏符(&C)'], ['wrap', '自动换行符(&w)'], ['nextPage', '下一页(&N)'], ['continuous', '连续(&t)'], ['evenPage', '偶数页(&E)'], ['oddPage', '奇数页(&O)']], 'page');
    const els = r.querySelectorAll('.ck');
    const body = h('div', { class: 'col' }, ui.group('分隔符类型', els[0], els[1], els[2]), ui.group('分节符类型', els[3], els[4], els[5], els[6]));
    ui.dialog({ title: '分隔符', width: 260, body, buttons: [{ label: '确定', primary: true, onClick: () => {
      const v = r.get();
      if (v === 'page' || v === 'column') { E.enter(v); return; }
      if (v === 'wrap') { E.enter('line'); return; }
      A().insertSectionBreak(v);
    } }, { label: '取消' }] });
  };
  A().insertSectionBreak = function (type) {
    const d = doc();
    if (!E.sel || D.storyOf(d, E.sel.f.p) !== d.main || D.cellOf(d, E.sel.f.p)) { A().status('分节符只能插入正文中，且须位于表格之外。'); return; }
    E.edit('分节符', () => {
      const pos = E.deleteSelection();
      const cur = D.sectFor(d, pos.p);
      const np = O.splitPara(pos);
      D.touch(pos.p);
      pos.p.sect = L.opc.duplicate(cur);
      /* the new section (after the break) carries the requested start type */
      const after = A().curSect ? null : null; void after;
      const tb = D.topBlock(d, np.p);
      const sec = D.sectOfBlock(d, tb.i);
      const holder = sec.endPara || d;
      D.touchKey(holder, 'sect');
      (holder === d ? d.sect : holder.sect).type = type;
      return np;
    });
  };

  /* ================= Page Numbers ================= */
  G.pageNumbers = function (mode) {
    const d = doc();
    const s = A().curSect();
    if (mode === 'format') { formatPageNumbers(s); return; }
    const pos = ui.select([['bottom', '页面底端（页脚）'], ['top', '页面顶端（页眉）']], 'bottom', () => draw(), { id: 'pn-pos' });
    const al = ui.select([['right', '右'], ['left', '左'], ['center', '居中'], ['inside', '内部'], ['outside', '外部']], 'right', () => draw(), { id: 'pn-al' });
    const first = ui.check('首页显示页码(&S)', true, null);
    const pv = h('div', { class: 'ps-page', style: 'width:80px;height:104px;margin:4px auto' });
    const draw = () => { L.clear(pv); const y = pos.value === 'top' ? 'top:6px' : 'bottom:6px'; const x = { left: 'left:8px', right: 'right:8px', center: 'left:36px', inside: 'right:8px', outside: 'right:8px' }[al.value]; pv.appendChild(h('span', { style: `position:absolute;${y};${x};width:8px;height:8px;border:1px solid #000;background:#fff` })); };
    draw();
    ui.dialog({
      title: '页码', width: 360, body: G.row(h('div', { class: 'col' }, G.f('位置(&P)：', pos), G.f('对齐方式(&A)：', al), first), h('div', { class: 'col' }, h('div', { class: 'cd-cap', text: '预览' }), pv)),
      buttons: [{ label: '格式(&F)...', onClick: () => { formatPageNumbers(s); return false; } }, { label: '确定', primary: true, onClick: () => {
        const kind = pos.value === 'top' ? 'hdr' : 'ftr';
        E.edit('页码', () => {
          D.touchKey(d, 'hf');
          const holder = s.holder;
          D.touchKey(holder, 'sect');
          const sect = holder === d ? d.sect : holder.sect;
          sect.refs = sect.refs || { hdr: {}, ftr: {} };
          sect.refs[kind] = Object.assign({}, sect.refs[kind]);
          const add = (ty, jc) => {
            let story = sect.refs[kind][ty] && d.hf[sect.refs[kind][ty]];
            if (!story) { const id = (kind === 'hdr' ? 'h' : 'f') + D.nid(); story = d.hf[id] = { kind, id, blocks: [D.para([], { style: kind === 'hdr' ? 'Header' : 'Footer' })] }; sect.refs[kind][ty] = id; }
            const fid = D.newFid();
            /* Word puts the number in its own paragraph aligned as chosen */
            story.blocks.push(D.para([D.item('fb', { fid, instr: ' PAGE ' }, { style: 'PageNumber' }), D.item('fs', { fid }, { style: 'PageNumber' }), D.text('1', { style: 'PageNumber' }), D.item('fe', { fid }, { style: 'PageNumber' })], { style: kind === 'hdr' ? 'Header' : 'Footer', jc }));
            if (!d.styles.PageNumber) { D.touchKey(d, 'styles'); d.styles.PageNumber = D.builtinStyles().PageNumber; }
          };
          const jc = (v, even) => (v === 'inside' ? (even ? 'right' : 'left') : v === 'outside' ? (even ? 'left' : 'right') : v);
          add('default', jc(al.value, false));
          if (d.settings.evenOdd) add('even', jc(al.value, true));
          sect.titlePg = !first.input.checked || sect.titlePg;
          if (!first.input.checked) { /* first page gets an empty header/footer */ if (!sect.refs[kind].first) { const id = (kind === 'hdr' ? 'h' : 'f') + D.nid(); d.hf[id] = { kind, id, blocks: [D.para([], { style: kind === 'hdr' ? 'Header' : 'Footer' })] }; sect.refs[kind].first = id; } }
          else if (sect.titlePg && sect.refs[kind].first) add('first', jc(al.value, false));
          d._idxDirty = true;
          return E.sel;
        });
      } }, { label: '取消' }],
      onOpen: (el) => el.querySelector('.dlg-foot .btn').classList.add('left'),
    });
  };
  function formatPageNumbers(s) {
    const d = doc();
    const sect = s.sect;
    const pn = Object.assign({ fmt: 'decimal', start: null }, sect.pgNum || {});
    const fmt = ui.select([['decimal', '1, 2, 3, ...'], ['numberInDash', '- 1 -, - 2 -, - 3 -, ...'], ['lowerLetter', 'a, b, c, ...'], ['upperLetter', 'A, B, C, ...'], ['lowerRoman', 'I, II, III, ...'], ['upperRoman', 'I, II, III, ...']], pn.fmt, null, { id: 'pf-fmt' });
    const chap = ui.check('包含章节号(&n)', !!pn.chapStyle, null);
    const mode = G.radios('pf-m', [['cont', '续前节(&C)'], ['start', '起始值(&a)：']], pn.start != null ? 'start' : 'cont');
    const startAt = G.num(pn.start != null ? pn.start : 1, { min: 0, id: 'pf-start' });
    ui.dialog({ title: '页码格式', width: 330, body: h('div', { class: 'col' }, G.f('数字格式(&f)：', fmt, 'wide'), chap, ui.group('页码编号', mode, startAt)), buttons: [{ label: '确定', primary: true, onClick: () => {
      E.edit('页码格式', () => {
        D.touchKey(s.holder, 'sect');
        const sc = s.holder === d ? d.sect : s.holder.sect;
        sc.pgNum = Object.assign({}, sc.pgNum, { fmt: fmt.value, start: mode.get() === 'start' ? startAt.get() : null });
        if (chap.input.checked !== !!pn.chapStyle) {
          if (chap.input.checked) sc.pgNum.chapStyle = 1;
          else { delete sc.pgNum.chapStyle; delete sc.pgNum.chapSep; }
        }
        return E.sel;
      });
    } }, { label: '取消' }] });
  }

  /* ================= Date and Time ================= */
  G.dateTime = function () {
    const now = new Date();
    const pics = F().DATE_PICTURES;
    const def = L.store.get('defDatePic', pics[0]);
    const lb = G.listBox(pics.map((p) => ({ value: p, label: F().formatDate(now, p) })), def, { height: 220, translate: 'no', onDbl: () => ok.click() });
    const lang = ui.select([['en-US', '英语（美国）'], ['en-GB', '英语（英国）']], 'en-US', null, { id: 'dt-lang' });
    const auto = ui.check('自动更新(&U)', !!L.store.get('defDateAuto', false), null);
    const dlg = ui.dialog({ title: '日期和时间', width: 360, body: h('div', { class: 'col' }, h('label', { text: '可用格式：' }), lb, G.f('语言(&L)：', lang), auto), buttons: [
      { label: '默认(&D)...', onClick: () => { L.store.set('defDatePic', lb.get()); L.store.set('defDateAuto', auto.input.checked); A().status('已保存默认日期格式。'); return false; } },
      { label: '确定', primary: true, onClick: () => { const pic = lb.get(); if (auto.input.checked) A().insertFieldNow(`${/[hHms]/.test(pic) && !/[dMy]/.test(pic) ? 'TIME' : 'DATE'} \\@ "${pic}"`, '插入日期和时间'); else E.insertTextAt(F().formatDate(new Date(), pic), '插入日期和时间'); } },
      { label: '取消' }], onOpen: (el) => el.querySelector('.dlg-foot .btn').classList.add('left') });
    const ok = dlg.buttons[1];
  };

  /* ================= Field ================= */
  const FIELD_CATS = {
    [L.N('日期和时间')]: ['CREATEDATE', 'DATE', 'EDITTIME', 'PRINTDATE', 'SAVEDATE', 'TIME'],
    [L.N('文档信息')]: ['AUTHOR', 'COMMENTS', 'DOCPROPERTY', 'FILENAME', 'FILESIZE', 'INFO', 'KEYWORDS', 'LASTSAVEDBY', 'NUMCHARS', 'NUMPAGES', 'NUMWORDS', 'SUBJECT', 'TEMPLATE', 'TITLE'],
    [L.N('公式与方程式')]: ['=', 'ADVANCE', 'EQ', 'SYMBOL'],
    [L.N('索引和目录')]: ['INDEX', 'RD', 'TA', 'TC', 'TOA', 'TOC', 'XE'],
    [L.N('链接和引用')]: ['AUTOTEXT', 'HYPERLINK', 'INCLUDEPICTURE', 'INCLUDETEXT', 'NOTEREF', 'PAGEREF', 'QUOTE', 'REF', 'STYLEREF'],
    [L.N('邮件合并')]: ['ASK', 'COMPARE', 'DATABASE', 'FILLIN', 'IF', 'MERGEFIELD', 'MERGEREC', 'MERGESEQ', 'NEXT', 'NEXTIF', 'SET', 'SKIPIF'],
    [L.N('编号')]: ['AUTONUM', 'AUTONUMLGL', 'AUTONUMOUT', 'BARCODE', 'LISTNUM', 'PAGE', 'REVNUM', 'SECTION', 'SECTIONPAGES', 'SEQ'],
    [L.N('用户信息')]: ['USERADDRESS', 'USERINITIALS', 'USERNAME'],
  };
  const FIELD_DESC = { DATE: "今天的日期", TIME: '当前时间', CREATEDATE: '文档创建日期', SAVEDATE: '文档上次保存日期', PRINTDATE: '文档上次打印日期', EDITTIME: '文档总编辑时间', AUTHOR: '文档属性中的作者姓名', TITLE: "文档属性中的标题", SUBJECT: "文档主题", KEYWORDS: "文档关键词", COMMENTS: '文档属性中的备注', FILENAME: "文档名称", NUMPAGES: '文档页数', NUMWORDS: '文档字数', NUMCHARS: '文档字符数', PAGE: '当前页码', SECTION: '当前节号', SECTIONPAGES: '节中的页数', SEQ: '插入自动序列编号', REF: '插入书签标记的文字', PAGEREF: '插入书签所在页的页码', NOTEREF: '插入脚注或尾注标记', STYLEREF: '插入具有相同样式的段落文字', TOC: '生成目录', INDEX: '生成索引', XE: '标记索引项', TC: '标记目录项', '=': '计算表达式的结果', MERGEFIELD: '插入邮件合并域', USERNAME: '用户信息中的姓名', USERINITIALS: '用户信息中的姓名缩写', USERADDRESS: '用户信息中的邮政地址', LASTSAVEDBY: '上次保存此文档的人员姓名', HYPERLINK: '打开并跳转到指定文件', QUOTE: '插入文字', SYMBOL: '插入符号', IF: '按条件对参数求值', REVNUM: "文档修订号", TEMPLATE: "文档模板名称", DOCPROPERTY: '文档属性的值', MACROBUTTON: '运行宏' };
  G.field = function (fieldRef, opts) {
    opts = opts || {};
    const editing = !!fieldRef;
    const cat = ui.select([['all', '（全部）']].concat(Object.keys(FIELD_CATS).map((k) => [k, k])), 'all', () => fillNames(), { id: 'fl-cat' });
    const names = G.listBox([], null, { height: 200, translate: 'no', onChange: (v) => { desc.textContent = L.t(FIELD_DESC[v] || ''); code.value = v === '=' ? '= ' : v + ' '; preview(); } });
    const fillNames = () => { const list = cat.value === 'all' ? Array.from(new Set(Object.values(FIELD_CATS).flat())).sort() : FIELD_CATS[cat.value]; names.setItems(list); };
    fillNames();
    const code = h('input', { type: 'text', id: 'fl-code', style: 'width:100%' });
    const desc = h('div', { class: 'tp-note', style: 'min-height:28px' });
    const res = h('div', { class: 'tp-box', style: 'min-height:22px', translate: 'no' });   // the field's result: the document's text
    const fmtSel = ui.select([['', '（无）'], ['Upper', '大写'], ['Lower', '小写'], ['FirstCap', '首字母大写'], ['Caps', '每词首字母大写'], ['Arabic', '1, 2, 3, ...'], ['roman', 'I, II, III, ...'], ['ROMAN', 'I, II, III, ...'], ['alphabetic', 'a, b, c, ...'], ['ALPHABETIC', 'A, B, C, ...'], ['Ordinal', '第 1、第 2、第 3 ...'], ['CardText', '一、二、三'], ['OrdText', '第一、第二、第三']], '', () => preview(), { id: 'fl-fmt' });
    const datePic = ui.select([['', L.t('（默认）')]].concat(F().DATE_PICTURES.map((p) => [p, p])), '', () => preview(), { id: 'fl-date', translate: 'no' });
    const keep = ui.check('更新时保留格式', true, null);
    const fullCode = () => { let c = code.value.trim(); if (fmtSel.value) c += ' \\* ' + fmtSel.value; if (datePic.value && /^(DATE|TIME|CREATEDATE|SAVEDATE|PRINTDATE)\b/i.test(c) && !/\\@/.test(c)) c += ` \\@ "${datePic.value}"`; if (keep.input.checked && !/MERGEFORMAT/i.test(c) && !/^(TOC|INDEX|=)/i.test(c)) c += ' \\* MERGEFORMAT'; return c; };
    const preview = () => { try { const r = F().evaluate(fullCode(), E.sel ? E.sel.f : null); res.textContent = r == null ? L.t('（插入或更新域时生成结果）') : Array.isArray(r) ? L.t('（生成的段落）') : String(r); } catch (e) { res.textContent = L.t('错误！'); } };
    if (editing) { code.value = fieldRef.it.instr.trim().replace(/\s*\\\*\s*MERGEFORMAT/i, ''); const t = F().parse(fieldRef.it.instr).type; names.set(t); desc.textContent = L.t(FIELD_DESC[t] || ''); }
    else if (opts.codes) code.value = '';
    code.addEventListener('input', preview);
    preview();
    ui.dialog({
      title: '域', width: 560, focus: opts.codes ? code : null,
      body: G.row(h('div', { class: 'col', style: 'width:190px' }, G.f('类别(&C)：', cat), h('label', { text: '域名：' }), names), h('div', { class: 'col', style: 'flex:1;min-width:220px' }, h('div', { class: 'cd-cap', text: '域代码：' }), code, desc, ui.group('域选项', G.f('格式(&F)：', fmtSel), G.f('日期格式(&p)：', datePic)), h('div', { class: 'cd-cap', text: '结果预览' }), res, keep)),
      buttons: [{ label: '确定', primary: true, onClick: () => {
        const c = fullCode();
        if (!c) return false;
        const t = F().parse(c).type;
        if (editing) { F().setInstr(fieldRef, c); return; }
        if (t === 'TOC' || t === 'INDEX') { F().insertGenerated(c); return; }
        A().insertFieldNow(c);
      } }, { label: '取消' }],
    });
  };

  /* ================= Symbol ================= */
  const SPECIALS = [['—', L.N('长划线'), 'Alt+Ctrl+小键盘 -'], ['–', L.N('短划线'), 'Ctrl+数字键盘 -'], ['‑', L.N('不间断连字符'), 'Ctrl+_'], ['­', L.N('可选连字符'), 'Ctrl+-'], [' ', L.N('全角空格'), ''], [' ', L.N('半角空格'), ''], [' ', L.N('1/4 全角空格'), ''], [' ', L.N('不间断空格'), 'Ctrl+Shift+Space'], ['©', L.N('版权'), 'Alt+Ctrl+C'], ['®', L.N('注册商标'), 'Alt+Ctrl+R'], ['™', L.N('商标'), 'Alt+Ctrl+T'], ['§', L.N('节'), ''], ['¶', L.N('段落'), ''], ['…', L.N('省略号'), 'Alt+Ctrl+.'], ['‘', L.N('左单引号'), "Ctrl+`,`"], ['’', L.N('右单引号'), "Ctrl+','"], ['“', L.N('左双引号'), 'Ctrl+`,"'], ['”', L.N('右双引号'), "Ctrl+',\""], ['​', L.N('无宽可选分隔符'), ''], ['⁠', L.N('无宽不换行符'), '']];
  const SUBSETS = [[L.N('基本拉丁文'), 0x20, 0x7e], [L.N('拉丁文-1 补充'), 0xa0, 0xff], [L.N('拉丁文扩展-A'), 0x100, 0x17f], [L.N('希腊文和科普特文'), 0x370, 0x3ff], [L.N('西里尔文'), 0x400, 0x4ff], [L.N('常用标点'), 0x2010, 0x205e], [L.N('货币符号'), 0x20a0, 0x20bf], [L.N('类字母符号'), 0x2100, 0x214f], [L.N('数字形式'), 0x2150, 0x218b], [L.N('箭头'), 0x2190, 0x21ff], [L.N('数学运算符'), 0x2200, 0x22ff], [L.N('杂项技术符号'), 0x2300, 0x23ff], [L.N('制表符号'), 0x2500, 0x257f], [L.N('区块元素'), 0x2580, 0x259f], [L.N('几何形状'), 0x25a0, 0x25ff], [L.N('杂项符号'), 0x2600, 0x26ff], [L.N('装饰符号'), 0x2700, 0x27bf]];
  G.symbol = function (opts) {
    opts = opts || {};
    let cur = null;
    const recent = L.store.get('recentSymbols', ['€', '£', '¥', '©', '®', '™', '±', '≠', '≤', '≥', '÷', '×', '∞', 'µ', 'α', 'β']);
    const font = ui.select([['', L.t('（普通文本）')], ['Symbol', 'Symbol'], ['Wingdings', 'Wingdings'], ['Webdings', 'Webdings']].concat(G.FONTS().map((f) => [f, f])), '', () => fill(), { id: 'sy-font', translate: 'no' });
    const subset = ui.select(SUBSETS.map((s, i) => [i, s[0]]), 1, () => fill(), { id: 'sy-sub' });
    const grid = h('div', { class: 'sym-grid' });
    const codeIn = h('input', { type: 'text', id: 'sy-code', style: 'width:70px' });
    const name = h('div', { class: 'tp-note', style: 'min-height:16px' });
    const fill = () => {
      L.clear(grid);
      const f = font.value;
      const chars = [];
      if (f === 'Symbol' || f === 'Wingdings' || f === 'Webdings') for (let c = 0x21; c <= 0xff; c++) { if (c >= 0x7f && c < 0xa1) continue; const u = L.mapSymbolChar(String.fromCharCode(c), f); if (u) chars.push({ ch: u, raw: String.fromCharCode(c), font: f }); }
      else { const [, a, b] = SUBSETS[+subset.value]; for (let c = a; c <= b; c++) chars.push({ ch: String.fromCodePoint(c), font: f }); }
      for (const c of chars) {
        const b = h('button', { type: 'button', text: c.ch, style: c.font && !/Symbol|Wingdings|Webdings/.test(c.font) ? `font-family:${L.fontStack(c.font)}` : '' });
        b.addEventListener('click', () => { L.$$('button.on', grid).forEach((x) => x.classList.remove('on')); b.classList.add('on'); cur = c; codeIn.value = c.ch.codePointAt(0).toString(16).toUpperCase().padStart(4, '0'); name.textContent = L.t('Unicode 字符 {0}', codeIn.value); });
        b.addEventListener('dblclick', () => { cur = c; insertBtn.click(); });
        grid.appendChild(b);
      }
    };
    fill();
    codeIn.addEventListener('change', () => { const n = parseInt(codeIn.value, 16); if (!isNaN(n)) { cur = { ch: String.fromCodePoint(n), font: '' }; name.textContent = L.t('Unicode 字符 {0}', codeIn.value.toUpperCase()); } });
    const rec = h('div', { class: 'sym-recent' });
    for (const c of recent.slice(0, 16)) { const b = h('button', { type: 'button', text: c }); b.addEventListener('click', () => { cur = { ch: c, font: '' }; codeIn.value = c.codePointAt(0).toString(16).toUpperCase().padStart(4, '0'); }); b.addEventListener('dblclick', () => { cur = { ch: c, font: '' }; insertBtn.click(); }); rec.appendChild(b); }
    const spec = G.listBox(SPECIALS.map(([c, n, k]) => ({ value: c, label: L.t(n) + (k ? '   ' + ui.keyText(k) : '') })), SPECIALS[0][0], { height: 220, onDbl: () => insertBtn.click() });
    const tabs = ui.tabs([{ label: '符号(&S)', body: h('div', { class: 'col' }, G.row(G.f('字体(&F)：', font), G.f('子集：', subset)), grid, h('div', { class: 'cd-cap', text: '近期使用过的符号：' }), rec, G.row(name, G.f('字符代码：', codeIn))) }, { label: '特殊字符(&p)', body: spec }]);
    let tabIdx = 0;
    tabs.querySelectorAll('.tab').forEach((t, i) => t.addEventListener('click', () => { tabIdx = i; }));
    const doInsert = () => {
      const c = tabIdx === 1 ? { ch: spec.get(), font: '' } : cur;
      if (!c) return;
      if (opts.pick) { opts.pick(c.ch, c.font); return; }
      const list = [c.ch].concat(recent.filter((x) => x !== c.ch)).slice(0, 16);
      L.store.set('recentSymbols', list);
      if (c.raw && c.font) E.insertItem(D.item('sym', { font: c.font, char: c.raw, code: c.raw.charCodeAt(0).toString(16).toUpperCase() }, O.inheritRPr(E.sel.f)), '插入符号');
      else E.insertTextAt(c.ch, '插入符号');
    };
    const dlg = ui.dialog({ title: '符号', width: 520, body: tabs, buttons: [{ label: '插入(&I)', primary: true, onClick: () => { doInsert(); if (opts.pick) return; return false; } }, { label: opts.pick ? '取消' : '关闭' }] });
    const insertBtn = dlg.buttons[0];
  };

  /* ================= Footnote and Endnote ================= */
  G.footnote = function () {
    const d = doc();
    const kind = G.radios('fn-k', [['fn', '脚注(&F)：'], ['en', '尾注(&E)：']], 'fn', () => sync());
    const fnLoc = ui.select([['pageBottom', '页面底端'], ['beneathText', '文字下方']], (d.settings.fnPr && d.settings.fnPr.pos) || 'pageBottom', null, { id: 'fn-floc' });
    const enLoc = ui.select([['docEnd', '文档结尾'], ['sectEnd', '节的结尾']], (d.settings.enPr && d.settings.enPr.pos) || 'docEnd', null, { id: 'fn-eloc' });
    const fmt = ui.select([['decimal', '1, 2, 3, ...'], ['lowerLetter', 'a, b, c, ...'], ['upperLetter', 'A, B, C, ...'], ['lowerRoman', 'I, II, III, ...'], ['upperRoman', 'I, II, III, ...'], ['chicago', '*, †, ‡, §, ...']], 'decimal', null, { id: 'fn-fmt' });
    const custom = h('input', { type: 'text', id: 'fn-custom', style: 'width:60px' });
    const startAt = G.num(1, { min: 1, id: 'fn-start' });
    const numbering = ui.select([['continuous', '连续'], ['eachSect', '每节重新编号'], ['eachPage', '每页重新编号']], 'continuous', null, { id: 'fn-num' });
    const sync = () => { const k = kind.get(); const pr = (k === 'fn' ? d.settings.fnPr : d.settings.enPr) || {}; fmt.value = pr.fmt || (k === 'fn' ? 'decimal' : 'lowerRoman'); startAt.set(pr.start || 1); numbering.value = pr.restart || 'continuous'; };
    sync();
    const settings = () => ({ fmt: fmt.value, start: startAt.get(), restart: numbering.value, pos: kind.get() === 'fn' ? fnLoc.value : enLoc.value });
    const kEls = kind.querySelectorAll('.ck');
    ui.dialog({
      title: '脚注和尾注', width: 380,
      body: h('div', { class: 'col' }, ui.group('位置', G.row(kEls[0], fnLoc), G.row(kEls[1], enLoc), ui.button('转换(&o)...', () => convert(), { class: 'btn small' })), ui.group('格式', G.f('数字格式(&o)：', fmt, 'wide'), G.row(G.f('自定义标记(&m)：', custom, 'wide'), ui.button('符号(&y)...', () => G.symbol({ pick: (ch) => { custom.value = ch; } }), { class: 'btn small' })), G.f('起始值(&S)：', startAt, 'wide'), G.f('编号：', numbering, 'wide'))),
      buttons: [{ label: '插入(&I)', primary: true, onClick: () => { L.notes.setOptions(kind.get(), settings()); L.notes.insert(kind.get(), { mark: custom.value.trim() || null }); } }, { label: '取消' }, { label: '应用(&A)', onClick: () => { L.notes.setOptions(kind.get(), settings()); return false; } }],
    });
    function convert() {
      const r = G.radios('fn-conv', [['fn', '将所有脚注转换为尾注(&f)'], ['en', '将所有尾注转换为脚注(&e)']], 'fn');
      ui.dialog({ title: '转换注释', width: 300, body: r, buttons: [{ label: '确定', primary: true, onClick: () => L.notes.convertAll(r.get()) }, { label: '取消' }] });
    }
  };

  /* ================= Caption ================= */
  G.caption = function () {
    const labels = Array.from(new Set(F().CAPTION_LABELS.concat(L.store.get('captionLabels', []))));
    const ci = L.tables && L.tables.inTable();
    const lab = ui.select(labels.map((x) => [x, x]), ci ? 'Table' : 'Figure', () => upd(), { id: 'cp-lab', translate: 'no' });   // labels go into the document as they are
    const txt = h('input', { type: 'text', id: 'cp-text', style: 'width:100%' });
    const pos = ui.select([['below', '所选项目下方'], ['above', '所选项目上方']], ci ? 'above' : 'below', null, { id: 'cp-pos' });
    const excl = ui.check('题注中不包含标签(&E)', false, () => upd());
    let num = { fmt: 'ARABIC', chapter: false, chapterStyle: 1, sep: '-' };
    const upd = () => { const n = L.D.fmtNum(1, { ARABIC: 'decimal', alphabetic: 'lowerLetter', ALPHABETIC: 'upperLetter', roman: 'lowerRoman', ROMAN: 'upperRoman' }[num.fmt]); txt.value = (excl.input.checked ? '' : lab.value + ' ') + (num.chapter ? '1' + num.sep : '') + n; };
    upd();
    ui.dialog({
      title: '题注', width: 360, focus: txt,
      body: h('div', { class: 'col' }, h('label', { for: 'cp-text', html: ui.ampHTML('题注(&C)：') }), txt, ui.group('选项', G.f('标签(&L)：', lab), G.f('位置(&P)：', pos), excl), h('div', { class: 'tp-btns' }, ui.button('新建标签(&N)...', async () => { const n = await ui.prompt('标签：', '', '新建标签'); if (!n) return; const list = L.store.get('captionLabels', []); if (!list.includes(n)) { list.push(n); L.store.set('captionLabels', list); } lab.appendChild(G.raw('option', { value: n }, n)); lab.value = n; upd(); }, { class: 'btn small' }), ui.button('删除标签(&D)', () => { if (F().CAPTION_LABELS.includes(lab.value)) return; L.store.set('captionLabels', L.store.get('captionLabels', []).filter((x) => x !== lab.value)); lab.querySelector(`option[value="${lab.value}"]`).remove(); upd(); }, { class: 'btn small' }), ui.button('编号(&u)...', () => numbering(), { class: 'btn small' }))),
      buttons: [{ label: '确定', primary: true, onClick: () => {
        const prefix = (excl.input.checked ? '' : lab.value + ' ');
        let rest = txt.value;
        const m = new RegExp('^' + prefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?:[\\dA-Za-z]+' + (num.chapter ? '[-.:—–]?[\\dA-Za-z]+' : '') + ')').exec(rest);
        rest = m ? rest.slice(m[0].length) : '';
        F().insertCaption({ label: lab.value, position: pos.value, text: rest, fmt: num.fmt, chapter: num.chapter, chapterStyle: num.chapterStyle, chapterSep: num.sep, noLabel: excl.input.checked });
      } }, { label: '取消' }],
    });
    function numbering() {
      const fmt = ui.select([['ARABIC', '1, 2, 3, ...'], ['alphabetic', 'a, b, c, ...'], ['ALPHABETIC', 'A, B, C, ...'], ['roman', 'i, ii, iii, ...'], ['ROMAN', 'I, II, III, ...']], num.fmt, null, { id: 'cn-fmt' });
      const chap = ui.check('包含章节号(&c)', num.chapter, null);
      const style = ui.select([1, 2, 3, 4, 5, 6, 7, 8, 9].map((i) => [i, 'Heading ' + i]), num.chapterStyle, null, { id: 'cn-st', translate: 'no' });   // style names
      const sep = ui.select([['-', '-（连字符）'], ['.', '.（句点）'], [':', ':（冒号）'], ['—', '—（长划线）'], ['–', '–（短划线）']], num.sep, null, { id: 'cn-sep' });
      ui.dialog({ title: '题注编号', width: 320, body: h('div', { class: 'col' }, G.f('格式(&o)：', fmt), chap, G.f('章节起始样式：', style, 'wide'), G.f('使用分隔符：', sep, 'wide')), buttons: [{ label: '确定', primary: true, onClick: () => { num = { fmt: fmt.value, chapter: chap.input.checked, chapterStyle: +style.value, sep: sep.value }; upd(); } }, { label: '取消' }] });
    }
  };

  /* ================= Cross-reference ================= */
  G.crossRef = function () {
    const kinds = [L.N('编号项'), L.N('标题'), L.N('书签'), L.N('脚注'), L.N('尾注')];
    const types = kinds.concat(['Equation', 'Figure', 'Table'], L.store.get('captionLabels', []));   // then the caption labels, as the document has them
    const what = { 'Numbered item': [['number', '段落编号'], ['numberFull', '段落编号（完整上下文）'], ['text', '段落文字'], ['page', '页码'], ['abovebelow', L.N('上方/下方')]], Heading: [['text', '标题文字'], ['page', '页码'], ['number', '标题编号'], ['abovebelow', L.N('上方/下方')]], Bookmark: [['text', '书签文字'], ['page', '页码'], ['number', '段落编号'], ['abovebelow', L.N('上方/下方')]], Footnote: [['notenum', '脚注编号'], ['page', '页码'], ['abovebelow', L.N('上方/下方')]], Endnote: [['notenum', '尾注编号'], ['page', '页码'], ['abovebelow', L.N('上方/下方')]] };
    const capWhat = [['text', '整项题注'], ['labelnum', '仅标签和编号'], ['page', '页码'], ['abovebelow', L.N('上方/下方')]];
    const type = ui.select(types.map((t) => [t, kinds.includes(t) ? L.t(t) : t]), 'Heading', () => refill(), { id: 'xr-type', translate: 'no' });
    const ref = ui.select([], '', null, { id: 'xr-ref' });
    const link = ui.check('插入为超链接(&h)', true, null);
    const ab = ui.check('包含“上方/下方”(&b)', false, null);
    const list = G.listBox([], null, { height: 200, translate: 'no' });
    let items = [];
    const refill = () => {
      L.clear(ref);
      for (const [v, l] of what[type.value] || capWhat) ref.appendChild(h('option', { value: v, text: l }));
      items = F().xrefItems(type.value);
      list.setItems(items.map((x, i) => ({ value: String(i), label: x.text || L.t('（空）') })));
      if (items.length) list.set('0');
    };
    refill();
    const dlg = ui.dialog({
      title: '交叉引用', width: 460, modeless: true,
      body: h('div', { class: 'col' }, G.row(G.f('引用类型：', type), G.f('引用内容(&r)：', ref)), G.row(link, ab), h('div', { class: 'cd-cap', text: '引用哪一项：' }), list),
      buttons: [{ label: '插入(&I)', primary: true, onClick: () => { const it = items[+list.get()]; if (!it) return false; F().insertXref(type.value, it, ref.value, { link: link.input.checked, abovebelow: ab.input.checked }); return false; } }, { label: '关闭' }],
    });
    void dlg;
  };

  /* ================= Index and Tables ================= */
  G.indexTables = function (tab) {
    const leaders = [['dot', '.......'], ['hyphen', '-------'], ['underscore', '_______'], ['none', L.N('（无）')]];
    /* TOC */
    const tShow = ui.check('显示页码(&S)', true, () => pvToc());
    const tRight = ui.check('页码右对齐(&R)', true, () => pvToc());
    const tLead = ui.select(leaders, 'dot', () => pvToc(), { id: 'it-tl' });
    const tLinks = ui.check('使用超链接而不使用页码（Web）(&U)', true, null);
    const tLevels = G.num(3, { min: 1, max: 9, id: 'it-lv', onChange: () => pvToc() });
    const tFormat = ui.select([['template', '根据模板'], ['classic', '经典'], ['distinctive', '独特'], ['formal', '正式'], ['simple', '简洁']], 'template', null, { id: 'it-tf' });
    let tStyles = null;
    const tPv = h('div', { class: 'preview-box', style: 'height:150px;flex-direction:column;align-items:stretch;justify-content:flex-start;padding:6px;font:11px Times New Roman,Tinos,serif' });
    const pvToc = () => { L.clear(tPv); for (let i = 1; i <= Math.min(3, tLevels.get()); i++) tPv.appendChild(h('div', { style: `display:flex;padding-left:${(i - 1) * 12}px;${i === 1 ? 'font-weight:bold' : ''}` }, h('span', { text: L.t('标题 {0}', i) }), h('span', { style: `flex:1;${tRight.input.checked && tLead.value !== 'none' ? 'border-bottom:1px dotted #000;margin:0 2px 3px' : ''}` }), h('span', { text: tShow.input.checked ? String(i * 2 - 1) : '' }))); };
    pvToc();
    const tocOptions = () => {
      const d = doc();
      const rows = h('div', { style: 'max-height:240px;overflow:auto;border:1px solid #919b9c;background:#fff' });
      const ids = Object.keys(d.styles).filter((k) => d.styles[k].type === 'paragraph').sort((a, b) => D.styleDisplayName(d.styles[a]).localeCompare(D.styleDisplayName(d.styles[b])));
      const inputs = {};
      for (const k of ids) { const lv = D.styleProps(d, k).pPr.outline; const inp = h('input', { type: 'text', style: 'width:36px', value: tStyles ? tStyles[k] || '' : lv != null && lv < tLevels.get() && /^Heading\d$/.test(k) ? String(lv + 1) : '' }); inputs[k] = inp; rows.appendChild(h('div', { class: 'row', style: 'justify-content:space-between;padding:1px 4px' }, G.raw('span', null, D.styleDisplayName(d.styles[k])), inp)); }
      const outl = ui.check('大纲级别(&O)', true, null), fields = ui.check('目录项域(&e)', false, null);
      ui.dialog({ title: '目录选项', width: 340, body: h('div', { class: 'col' }, h('div', { text: '目录来源：样式 — 目录级别：' }), rows, outl, fields), buttons: [{ label: '确定', primary: true, onClick: () => { tStyles = {}; for (const k in inputs) if (inputs[k].value.trim()) tStyles[k] = +inputs[k].value; tStyles.__outline = outl.input.checked; tStyles.__fields = fields.input.checked; } }, { label: '取消' }] });
    };
    const tocBody = h('div', { class: 'col' }, G.row(h('div', { class: 'col', style: 'flex:1' }, h('div', { class: 'cd-cap', text: '打印预览' }), tPv, tShow, tRight, G.f('制表位前导符(&b)：', tLead)), h('div', { class: 'col', style: 'flex:1' }, h('div', { class: 'cd-cap', text: 'Web 预览' }), h('div', { class: 'preview-box', style: 'height:150px;flex-direction:column;align-items:flex-start;justify-content:flex-start;padding:6px;gap:3px' }, h('u', { text: '标题 1', style: 'color:#00f' }), h('u', { text: '标题 2', style: 'color:#00f;margin-left:10px' }), h('u', { text: '标题 3', style: 'color:#00f;margin-left:20px' })), tLinks)), ui.group('常规', G.row(G.f('格式(&t)：', tFormat), G.f('显示级别(&l)：', tLevels))), h('div', { class: 'tp-btns' }, ui.button('选项(&O)...', tocOptions, { class: 'btn small' }), ui.button('修改(&M)...', () => G.style('TOC1'), { class: 'btn small' })));
    /* Table of figures */
    const fLabel = ui.select(['Figure', 'Table', 'Equation'].concat(L.store.get('captionLabels', [])).map((x) => [x, x]), 'Figure', null, { id: 'it-fl', translate: 'no' });
    const fIncl = ui.check('包含标签和编号(&n)', true, null);
    const fShow = ui.check('显示页码(&S)', true, null);
    const fRight = ui.check('页码右对齐(&R)', true, null);
    const fLead = ui.select(leaders, 'dot', null, { id: 'it-fle' });
    const tofBody = h('div', { class: 'col' }, G.f('题注标签(&l)：', fLabel), fIncl, fShow, fRight, G.f('制表位前导符(&b)：', fLead), h('div', { class: 'tp-note', text: '通过“插入 ▸ 引用 ▸ 题注”插入的题注会列在这里。' }));
    /* Index */
    const iType = G.radios('it-it', [['indent', '缩进式(&n)'], ['runin', '接排式(&n)']], 'indent');
    const iCols = G.num(2, { min: 1, max: 4, id: 'it-ic' });
    const iRight = ui.check('页码右对齐(&R)', false, null);
    const iLead = ui.select(leaders, 'dot', null, { id: 'it-il' });
    const iHead = ui.check('显示信函标题(&h)', false, null);
    const idxBody = h('div', { class: 'col' }, G.row(ui.group('类型', iType), h('div', { class: 'col' }, G.f('列(&o)：', iCols), iRight, G.f('制表位前导符(&b)：', iLead), iHead)), h('div', { class: 'tp-btns' }, ui.button('标记索引项(&k)...', () => G.markEntry(), { class: 'btn small' })), h('div', { class: 'tp-note', text: '先标记索引项 (Alt+Shift+X)，再在所需位置插入索引。' }));
    const toaBody = h('div', { class: 'tp-note', text: '引文目录需要使用 TA 域标记引文。通过“插入 ▸ 域 ▸ TA”标记引文，再使用“域 ▸ TOA”生成目录。' });
    let cur = tab === 'toc' || tab == null ? 1 : tab === 'tof' ? 2 : tab === 'index' ? 0 : 1;
    const tabs = ui.tabs([{ label: '索引(&x)', body: idxBody, onShow: () => { cur = 0; } }, { label: '目录(&C)', body: tocBody, onShow: () => { cur = 1; } }, { label: '图表目录(&F)', body: tofBody, onShow: () => { cur = 2; } }, { label: '引文目录(&A)', body: toaBody, onShow: () => { cur = 3; } }], cur);
    ui.dialog({
      title: '索引和目录', width: 520, body: tabs,
      buttons: [{ label: '确定', primary: true, onClick: () => {
        if (cur === 1) {
          let instr = `TOC \\o "1-${tLevels.get()}"`;
          if (tStyles) { const st = Object.entries(tStyles).filter(([k]) => !k.startsWith('__') && !/^Heading\d$/.test(k)).map(([k, v]) => `${D.styleDisplayName(doc().styles[k])},${v}`).join(','); if (!tStyles.__outline) instr = 'TOC'; if (st) instr += ` \\t "${st}"`; if (tStyles.__fields) instr += ' \\f'; }
          if (tLinks.input.checked) instr += ' \\h';
          instr += ' \\z \\u';
          if (!tShow.input.checked) instr += ' \\n';
          else if (!tRight.input.checked) instr += ' \\p " "';
          F().insertGenerated(instr);
          if (tFormat.value !== 'template') applyTocFormat(tFormat.value, tLevels.get());
        } else if (cur === 2) {
          let instr = `TOC \\h \\z \\c "${fLabel.value}"`;
          if (!fIncl.input.checked) instr += ` \\a "${fLabel.value}"`;
          if (!fShow.input.checked) instr += ' \\n';
          F().insertGenerated(instr);
        } else if (cur === 0) {
          let instr = `INDEX \\c "${iCols.get()}"`;
          if (iType.get() === 'runin') instr += ' \\r';
          if (iHead.input.checked) instr += ' \\h "A"';
          if (iRight.input.checked) instr += ' \\e "\t"';
          F().insertGenerated(instr, { columns: iCols.get() });
        }
      } }, { label: '取消' }],
    });
  };
  function applyTocFormat(kind, levels) {
    const d = doc();
    const look = { classic: [{ b: true, caps: true }, { smallCaps: true }, { i: true }], distinctive: [{ b: true, font: 'Arial', sz: 12 }, { font: 'Arial', sz: 11 }, { sz: 10 }], formal: [{ caps: true, b: true }, { caps: true }, { i: true }], simple: [{}, {}, {}] }[kind];
    E.edit('目录格式', () => { D.touchKey(d, 'styles'); for (let i = 1; i <= levels; i++) { const id = 'TOC' + i; d.styles[id] = d.styles[id] || D.builtinStyles()[id]; d.styles[id].rPr = Object.assign({}, look[Math.min(i - 1, 2)]); } D.stylesChanged(); return E.sel; });
  }
  G.markEntry = function () {
    const sel = E.sel && !E.collapsed() ? O.textRange(...E.range(), ' ').trim() : '';
    const main = h('input', { type: 'text', id: 'me-main', value: sel, style: 'width:100%' });
    const sub = h('input', { type: 'text', id: 'me-sub', style: 'width:100%' });
    const opt = G.radios('me-o', [['cur', '当前页(&C)'], ['xref', '交叉引用(&r)：']], 'cur');
    const xref = h('input', { type: 'text', id: 'me-x', value: 'See ', style: 'width:100%' });
    const bold = ui.check('加粗(&B)', false, null), ital = ui.check('倾斜(&I)', false, null);
    const mark = (all) => {
      const m = main.value.trim();
      if (!m) return;
      const entry = (m + (sub.value.trim() ? ':' + sub.value.trim() : '')).replace(/"/g, '\\"');
      let instr = `XE "${entry}"`;
      if (opt.get() === 'xref') instr += ' \\t "' + xref.value.replace(/"/g, '\\"') + '"';
      if (bold.input.checked) instr += ' \\b';
      if (ital.input.checked) instr += ' \\i';
      if (!all) { E.edit('标记索引项', () => { const [, b] = E.range(); return O.insertField(b, instr, '', {}); }); return; }
      /* Mark All: every occurrence of the text in the main story */
      const d = doc();
      D.reindex(d);
      let n = 0;
      E.edit('标记索引项', () => {
        for (const p of d.main.paras.slice()) {
          const t = D.ptext(p);
          let i = t.toLowerCase().indexOf(m.toLowerCase());
          const hits = [];
          while (i >= 0) { hits.push(i + m.length); i = t.toLowerCase().indexOf(m.toLowerCase(), i + m.length); }
          for (const at of hits.reverse()) { O.insertField(D.pos(p, at), instr, '', {}); n++; }
        }
        return E.sel;
      });
      A().status(n === 1 ? L.t('已标记 {0} 项。', n) : L.t('已标记 {0} 项。', n));
    };
    ui.dialog({ title: '标记索引项', width: 360, modeless: true, body: h('div', { class: 'col' }, G.f('主索引项(&e)：', main, 'wide'), G.f('次索引项(&S)：', sub, 'wide'), ui.group('选项', opt, xref), ui.group('页码格式', bold, ital)), buttons: [{ label: '标记(&M)', primary: true, onClick: () => { mark(false); return false; } }, { label: '标记全部(&A)', onClick: () => { mark(true); return false; } }, { label: '关闭' }] });
  };

  /* ================= Bookmark ================= */
  G.bookmark = function () {
    const sortBy = G.radios('bm-s', [['name', '名称(&N)'], ['loc', '位置(&L)']], 'name', () => fill());
    const hidden = ui.check('隐藏书签(&H)', false, () => fill());
    const name = h('input', { type: 'text', id: 'bm-name', style: 'width:100%' });
    const list = G.listBox([], null, { height: 160, translate: 'no', onChange: (v) => { name.value = v; }, onDbl: () => go() });
    const fill = () => {
      const all = Array.from(F().bookmarks().values()).filter((b) => hidden.input.checked || !b.name.startsWith('_'));
      if (sortBy.get() === 'name') all.sort((a, b) => a.name.localeCompare(b.name)); else all.sort((a, b) => (a.p._o || 0) - (b.p._o || 0) || a.o - b.o);
      list.setItems(all.map((b) => b.name));
    };
    fill();
    if (E.sel && !E.collapsed()) { const t = O.textRange(...E.range(), ' ').trim().replace(/[^\p{L}\p{N}_]/gu, '_').replace(/^[^\p{L}]+/u, '').slice(0, 40); name.value = t; }
    const valid = (n) => /^[\p{L}][\p{L}\p{N}_]{0,39}$/u.test(n);
    const go = () => { const n = name.value.trim(); if (n) A().gotoBookmark(n); };
    ui.dialog({
      title: '书签', width: 320, focus: name,
      body: h('div', { class: 'col' }, h('label', { for: 'bm-name', html: ui.ampHTML('书签名(&B)：') }), name, list, ui.group('排序依据：', sortBy), hidden),
      buttons: [{ label: '添加(&A)', primary: true, onClick: () => { const n = name.value.trim(); if (!valid(n)) { ui.msg('书签名必须以字母开头，且只能包含字母、数字和下划线。', { icon: 'warn' }); return false; } const [a, b] = E.range(); E.edit(L.N('书签'), () => { F().addBookmark(n, a, b); return E.sel; }); } }, { label: '删除(&D)', onClick: () => { const bm = F().bookmarks().get(name.value.trim()); if (bm) E.edit('删除书签', () => { F().removeBookmarkItems(bm); return E.sel; }); fill(); name.value = ''; return false; } }, { label: '定位(&G)', onClick: () => { go(); return false; } }, { label: '取消' }],
    });
  };

  /* ================= Hyperlink ================= */
  G.hyperlink = function () {
    const d = doc();
    let cur = null;
    if (E.sel) {
      const it = D.itemAfter(E.sel.f.p, E.sel.f.o) || D.itemBefore(E.sel.f.p, E.sel.f.o);
      if (it && it.rPr && it.rPr.link) cur = it.rPr.link;
    }
    /* the whole link text when the caret sits in a link */
    let range = E.sel ? E.range() : null;
    if (cur && E.collapsed()) {
      const p = E.sel.f.p; let pos = 0, s = -1, e = -1;
      for (const it of p.runs) { const n = D.ilen(it); if (it.rPr && it.rPr.link && JSON.stringify(it.rPr.link) === JSON.stringify(cur) && pos <= E.sel.f.o + 1) { if (s < 0) s = pos; e = pos + n; } else if (s >= 0 && pos >= E.sel.f.o) break; pos += n; }
      if (s >= 0) range = [D.pos(p, s), D.pos(p, e)];
    }
    const selText = range && !D.eqPos(range[0], range[1]) ? O.textRange(range[0], range[1], ' ') : '';
    const disp = h('input', { type: 'text', id: 'hl-disp', value: selText || '<<Selection in Document>>', style: 'width:100%' });
    if (!selText) disp.value = '';
    const addr = h('input', { type: 'text', id: 'hl-addr', value: cur && cur.url && !/^mailto:/.test(cur.url) ? cur.url : '', style: 'width:100%' });
    const tip = cur && cur.tip ? cur.tip : '';
    const place = G.listBox([], cur && cur.anchor ? cur.anchor : null, { height: 170, translate: 'no' });   // headings and bookmarks: the document's own
    const places = [{ value: '', label: L.t('文档开头') }];
    D.reindex(d);
    places.push({ value: '', label: L.t('标题'), group: true });
    for (const p of d.main.paras) { const lv = D.headingLevel(d, p); if (lv) places.push({ value: '§' + p.id, label: '    '.repeat(lv - 1) + D.plainText(p).trim().slice(0, 60) }); }
    places.push({ value: '', label: L.t('书签'), group: true });
    for (const [n] of F().bookmarks()) if (!n.startsWith('_')) places.push({ value: n, label: n });
    place.setItems(places);
    const mail = h('input', { type: 'text', id: 'hl-mail', value: cur && /^mailto:/.test(cur.url || '') ? cur.url.slice(7).split('?')[0] : '', style: 'width:100%' });
    const subj = h('input', { type: 'text', id: 'hl-subj', style: 'width:100%' });
    let kind = cur && cur.anchor && !cur.url ? 1 : cur && /^mailto:/.test(cur.url || '') ? 2 : 0;
    const tabs = ui.tabs([
      { label: '现有文件或网页(&x)', body: h('div', { class: 'col' }, h('label', { for: 'hl-addr', html: ui.ampHTML('地址(&e)：') }), addr, h('div', { class: 'tp-note', text: L.t('输入网址 (https://…) 或文件名。') })), onShow: () => { kind = 0; } },
      { label: '本文档中的位置(&D)', body: h('div', { class: 'col' }, h('div', { text: '选择本文档中的位置：' }), place), onShow: () => { kind = 1; } },
      { label: '电子邮件地址(&m)', body: h('div', { class: 'col' }, h('label', { for: 'hl-mail', text: '电子邮件地址：' }), mail, h('label', { for: 'hl-subj', text: '主题：' }), subj), onShow: () => { kind = 2; } },
    ], kind);
    let screenTip = tip;
    ui.dialog({
      title: cur ? '编辑超链接' : '插入超链接', width: 520, focus: kind === 0 ? addr : null,
      body: h('div', { class: 'col' }, G.row(h('label', { for: 'hl-disp', html: ui.ampHTML('要显示的文字(&T)：'), style: 'width:110px' }), h('div', { style: 'flex:1' }, disp), ui.button('屏幕提示(&p)...', async () => { const t = await ui.prompt('屏幕提示文字：', screenTip, '设置超链接屏幕提示'); if (t != null) screenTip = t; }, { class: 'btn small' })), tabs),
      buttons: (cur ? [{ label: '删除链接(&R)', onClick: () => { A().removeHyperlink(); } }] : []).concat([{ label: '确定', primary: true, onClick: () => {
        let link;
        if (kind === 0) { let u = addr.value.trim(); if (!u) return false; if (/^www\./i.test(u)) u = 'http://' + u; link = { url: u }; }
        else if (kind === 1) {
          let a = place.get();
          if (a && a.startsWith('§')) { const p = D.byId(d, +a.slice(1)); E.edit(L.N('书签'), () => { a = F().ensureParaBookmark(p, '_Toc'); return E.sel; }); }
          link = { anchor: a || '_top' };
        } else { const m = mail.value.trim(); if (!m) return false; link = { url: 'mailto:' + m + (subj.value ? '?subject=' + encodeURIComponent(subj.value) : '') }; }
        if (screenTip) link.tip = screenTip;
        const text = disp.value || (kind === 0 ? addr.value : kind === 2 ? mail.value : place.get() || '文档开头');
        if (!d.styles.Hyperlink) E.edit(L.N('样式'), () => { D.touchKey(d, 'styles'); d.styles.Hyperlink = D.builtinStyles().Hyperlink; D.stylesChanged(); return E.sel; });
        E.edit(cur ? '编辑超链接' : '插入超链接', () => {
          if (range && !D.eqPos(range[0], range[1]) && text === selText) { O.formatRange(range[0], range[1], (r) => Object.assign(r, { link, style: 'Hyperlink' })); return { a: range[0], f: range[1] }; }
          E.sel = { a: range[0], f: range[1] };
          const pos = E.deleteSelection();
          const base = O.inheritRPr(pos);
          delete base.link;
          return O.insertText(pos, text, Object.assign(base, { link, style: 'Hyperlink' }));
        });
        E.pending = { link: undefined, style: undefined };
      } }, { label: '取消' }]),
    });
  };

  /* ================= Insert Object ================= */
  G.insertObject = function () {
    const types = [['chart', L.N('图表')], ['equation', L.N('公式')], ['wordart', L.N('LettersArt')], ['excel', L.N('电子表格（表格）')], ['document', L.N('文档（作为文本插入）')], ['picture', L.N('位图图像')]];
    const lb = G.listBox(types.map(([v, l]) => ({ value: v, label: l })), 'chart', { height: 160, onDbl: () => ok.click() });
    const dlg = ui.dialog({ title: '对象', width: 360, body: ui.tabs([{ label: '新建(&C)', body: h('div', { class: 'col' }, h('label', { text: '对象类型：' }), lb) }, { label: '由文件创建(&F)', body: h('div', { class: 'col' }, h('div', { class: 'tp-note', text: '插入 .docx、.txt 或 .htm 文件的内容，或插入图片。' }), ui.button('浏览(&B)...', () => { dlg.close(null); A().insertFileDialog(); })) }]), buttons: [{ label: '确定', primary: true, onClick: () => {
      const v = lb.get();
      if (v === 'chart') A().insertChart();
      else if (v === 'equation') G.equation();
      else if (v === 'wordart') G.wordart();
      else if (v === 'excel') L.tables.insert(4, 5, { worksheet: true });
      else if (v === 'document') A().insertFileDialog();
      else A().insertPictureDialog();
    } }, { label: '取消' }] });
    const ok = dlg.buttons[0];
  };
  /** a simple equation editor: linear format with Unicode math symbols */
  G.equation = function () {
    const inp = h('textarea', { rows: 3, id: 'eq-in', style: 'width:100%;font:16px "Cambria Math",Caladea,serif' });
    const pal = h('div', { class: 'sym-recent', style: 'grid-template-columns:repeat(16,26px)' });
    for (const c of '±×÷≠≈≤≥∞√∛∑∏∫∂∆∇πθαβγδλμσφωΩ→⇒∈∉⊂⊃∪∩∀∃¬∧∨²³⁴ⁿ₀₁₂ᵢ½'.split('')) { const b = h('button', { type: 'button', text: c }); b.addEventListener('click', () => { const s = inp.selectionStart; inp.value = inp.value.slice(0, s) + c + inp.value.slice(inp.selectionEnd); inp.focus(); inp.setSelectionRange(s + 1, s + 1); }); pal.appendChild(b); }
    ui.dialog({ title: '公式', width: 470, focus: inp, body: h('div', { class: 'col' }, pal, inp, h('div', { class: 'tp-note', text: '键入公式；将以 Cambria Math 斜体插入。' })), buttons: [{ label: '确定', primary: true, onClick: () => { if (!inp.value.trim()) return; E.edit('插入公式', () => { const pos = E.deleteSelection(); return O.insertText(pos, inp.value.trim(), Object.assign(O.inheritRPr(pos), { font: 'Cambria Math', i: true })); }); } }, { label: '取消' }] });
  };

  /* ================= Diagram Gallery ================= */
  G.diagram = function () {
    const kinds = [['org', '组织结构图'], ['cycle', '循环图'], ['radial', '射线图'], ['pyramid', '棱锥图'], ['venn', '维恩图'], ['target', '靶心图']];
    const svg = { org: '<svg width="56" height="44"><rect x="20" y="2" width="16" height="10" fill="#fdd" stroke="#333"/><path d="M28 12v6M10 18h36M10 18v4M28 18v4M46 18v4" stroke="#333" fill="none"/><rect x="2" y="22" width="16" height="10" fill="#dfd" stroke="#333"/><rect x="20" y="22" width="16" height="10" fill="#dfd" stroke="#333"/><rect x="38" y="22" width="16" height="10" fill="#dfd" stroke="#333"/></svg>', cycle: '<svg width="56" height="44"><circle cx="28" cy="22" r="16" fill="none" stroke="#888" stroke-dasharray="3 2"/><circle cx="28" cy="6" r="5" fill="#9cf" stroke="#333"/><circle cx="42" cy="30" r="5" fill="#fc9" stroke="#333"/><circle cx="14" cy="30" r="5" fill="#cf9" stroke="#333"/></svg>', radial: '<svg width="56" height="44"><path d="M28 22L28 6M28 22L44 32M28 22L12 32" stroke="#333"/><circle cx="28" cy="22" r="7" fill="#fc9" stroke="#333"/><circle cx="28" cy="6" r="5" fill="#9cf" stroke="#333"/><circle cx="44" cy="32" r="5" fill="#9cf" stroke="#333"/><circle cx="12" cy="32" r="5" fill="#9cf" stroke="#333"/></svg>', pyramid: '<svg width="56" height="44"><path d="M28 2L50 42H6z" fill="#fc9" stroke="#333"/><path d="M19 18h18M12 30h32" stroke="#333"/></svg>', venn: '<svg width="56" height="44"><circle cx="22" cy="18" r="12" fill="#9cf" fill-opacity=".6" stroke="#333"/><circle cx="34" cy="18" r="12" fill="#fc9" fill-opacity=".6" stroke="#333"/><circle cx="28" cy="28" r="12" fill="#cf9" fill-opacity=".6" stroke="#333"/></svg>', target: '<svg width="56" height="44"><circle cx="24" cy="24" r="18" fill="#9cf" stroke="#333"/><circle cx="24" cy="24" r="12" fill="#fc9" stroke="#333"/><circle cx="24" cy="24" r="6" fill="#f99" stroke="#333"/></svg>' };
    let pick = 'org';
    const grid = h('div', { class: 'preset-grid' });
    for (const [k, l] of kinds) { const b = h('button', { type: 'button', class: k === pick ? 'on' : '', 'data-tip': l, 'aria-label': l, html: svg[k] }); b.addEventListener('click', () => { pick = k; L.$$('button', grid).forEach((x) => x.classList.toggle('on', x === b)); name.textContent = L.t(l); }); b.addEventListener('dblclick', () => ok.click()); grid.appendChild(b); }
    const name = h('div', { class: 'cd-cap', text: kinds[0][1] });
    const dlg = ui.dialog({ title: '图示库', width: 340, body: h('div', { class: 'col' }, h('div', { text: '选择图示类型：' }), grid, name), buttons: [{ label: '确定', primary: true, onClick: () => insertDiagram(pick) }, { label: '取消' }] });
    const ok = dlg.buttons[0];
  };
  function insertDiagram(kind) {
    const S = L.drawing.newShape;
    const tb = (sh, text) => { sh.tb = { kind: 'tb', blocks: [D.para(text ? [D.text(text, { sz: 10 })] : [], { jc: 'center' })] }; sh.ins = { l: 3.6, t: 3.6, r: 3.6, b: 3.6 }; sh.anchor = 'ctr'; return sh; };
    const kid = (sh, x, y) => Object.assign(sh, { x, y });
    const W = 324, H = 216;
    const kids = [];
    const fills = ['#99CCFF', '#FFCC99', '#CCFF99', '#FF9999', '#CC99FF', '#FFFF99'];
    if (kind === 'org') {
      kids.push(kid(tb(S('rect', 90, 40), 'Manager'), 117, 8));
      for (let i = 0; i < 3; i++) kids.push(kid(tb(S('rect', 90, 40), 'Employee ' + (i + 1)), 9 + i * 108, 120));
      kids.push(kid(S('line', 0, 40), 162, 48), kid(S('line', 216, 0), 54, 88));
      for (let i = 0; i < 3; i++) kids.push(kid(S('line', 0, 32), 54 + i * 108, 88));
    } else if (kind === 'pyramid') {
      for (let i = 0; i < 3; i++) { const sh = S('trapezoid', 120 + i * 80, 56); sh.fill = { t: 'solid', c: fills[i], a: 1 }; tb(sh, 'Level ' + (i + 1)); kids.push(kid(sh, (W - (120 + i * 80)) / 2, 16 + i * 62)); }
      kids[0].geom = 'triangle'; kids[0].w = 120;
    } else {
      const n = kind === 'target' ? 3 : kind === 'venn' ? 3 : 4;
      for (let i = 0; i < n; i++) {
        const a = (i / n) * Math.PI * 2 - Math.PI / 2;
        let sh;
        if (kind === 'target') { const r = 100 - i * 30; sh = S('ellipse', r * 2, r * 2); kids.push(kid(Object.assign(sh, { fill: { t: 'solid', c: fills[i], a: 1 } }), W / 2 - r, H / 2 - r)); continue; }
        if (kind === 'venn') { sh = S('ellipse', 120, 120); sh.fill = { t: 'solid', c: fills[i], a: 0.5 }; kids.push(kid(sh, W / 2 - 60 + Math.cos(a) * 40, H / 2 - 60 + Math.sin(a) * 40)); continue; }
        sh = tb(S('ellipse', 72, 52), kind === 'cycle' ? 'Step ' + (i + 1) : 'Item ' + (i + 1));
        sh.fill = { t: 'solid', c: fills[i], a: 1 };
        kids.push(kid(sh, W / 2 - 36 + Math.cos(a) * 110, H / 2 - 26 + Math.sin(a) * 75));
      }
      if (kind === 'radial') { const c = tb(S('ellipse', 80, 56), 'Center'); c.fill = { t: 'solid', c: '#FFCC99', a: 1 }; kids.unshift(kid(c, W / 2 - 40, H / 2 - 28)); }
    }
    const g = { t: 'group', w: W, h: H, kids, name: 'Diagram ' + D.nid() };
    E.edit('插入图示', () => O.insertItem(E.deleteSelection(), g));
  }

  /* ================= WordArt ================= */
  const WA_STYLES = [
    { fill: { t: 'solid', c: '#000000' }, warp: 'textPlain' }, { fill: { t: 'grad', ang: 90, stops: [{ p: 0, c: '#FFFFFF' }, { p: 1, c: '#999999' }] }, line: { c: '#000000', w: 0.75 }, warp: 'textPlain' }, { fill: { t: 'solid', c: '#C0C0C0' }, line: { c: '#000000', w: 0.75 }, warp: 'textPlain', i: true }, { fill: { t: 'solid', c: '#000000' }, warp: 'textSlantUp' }, { fill: { t: 'grad', ang: 0, stops: [{ p: 0, c: '#FF9900' }, { p: 1, c: '#FFFF00' }] }, warp: 'textArchUp' },
    { fill: { t: 'solid', c: '#3366FF' }, line: { c: '#000080', w: 0.75 }, warp: 'textPlain', shadow: { dx: 3, dy: 3 } }, { fill: { t: 'grad', ang: 90, stops: [{ p: 0, c: '#FF0000' }, { p: 0.5, c: '#FFFF00' }, { p: 1, c: '#00CC00' }] }, warp: 'textWave1' }, { fill: { t: 'solid', c: '#FFFFFF' }, line: { c: '#3366CC', w: 1.5 }, warp: 'textPlain', b: true }, { fill: { t: 'solid', c: '#993300' }, warp: 'textChevron' }, { fill: { t: 'grad', ang: 90, stops: [{ p: 0, c: '#C0C0C0' }, { p: 1, c: '#333333' }] }, warp: 'textCanUp' },
    { fill: { t: 'solid', c: '#CC0000' }, line: { c: '#660000', w: 0.75 }, warp: 'textArchDown' }, { fill: { t: 'solid', c: '#336633' }, warp: 'textTriangle' }, { fill: { t: 'grad', ang: 45, stops: [{ p: 0, c: '#0000FF' }, { p: 1, c: '#00FFFF' }] }, warp: 'textPlain', b: true }, { fill: { t: 'solid', c: '#FFCC00' }, line: { c: '#996600', w: 0.75 }, warp: 'textSlantDown', shadow: { dx: 2, dy: 2 } }, { fill: { t: 'solid', c: '#660099' }, warp: 'textCircle' },
  ];
  G.wordart = function (mode) {
    const it = mode && L.drawing.selectedItem();
    if (mode === true && it && it.wordart) { editText(it, (wa) => L.drawing.modifySelected('编辑 LettersArt', (x) => { x.wordart = Object.assign({}, x.wordart, wa); })); return; }
    let pick = 0;
    const grid = h('div', { class: 'wa-grid' });
    WA_STYLES.forEach((s, i) => {
      const b = h('button', { type: 'button', class: i === 0 ? 'on' : '', 'aria-label': L.t('LettersArt 样式 {0}', i + 1) });
      const sh = { t: 'shape', w: 66, h: 40, fill: Object.assign({ a: 1 }, s.fill), line: s.line ? Object.assign({ dash: 'solid' }, s.line) : { t: 'none' }, warp: s.warp, shadow: s.shadow ? Object.assign({ c: '#808080', a: 0.6 }, s.shadow) : undefined, wordart: { text: 'LettersArt', font: 'Arial Black', b: s.b, i: s.i } };
      b.appendChild(L.drawing.wordartSVG(sh));
      b.addEventListener('click', () => { pick = i; L.$$('button', grid).forEach((x) => x.classList.toggle('on', x === b)); });
      b.addEventListener('dblclick', () => ok.click());
      grid.appendChild(b);
    });
    const dlg = ui.dialog({ title: 'LettersArt 库', width: 470, body: h('div', { class: 'col' }, h('div', { text: '选择 LettersArt 样式：' }), grid), buttons: [{ label: '确定', primary: true, onClick: () => {
      const s = WA_STYLES[pick];
      if (mode === 'gallery' && it) { L.drawing.modifySelected('LettersArt Gallery', (x) => { x.fill = Object.assign({ a: 1 }, s.fill); x.line = s.line ? Object.assign({ dash: 'solid' }, s.line) : { t: 'none' }; x.warp = s.warp; x.shadow = s.shadow ? Object.assign({ c: '#808080', a: 0.6 }, s.shadow) : undefined; }); return; }
      const init = E.sel && !E.collapsed() ? O.textRange(...E.range(), ' ').trim().slice(0, 200) : '在此输入文字';
      setTimeout(() => editText({ wordart: { text: init, font: 'Arial Black', sz: 36, b: s.b, i: s.i } }, (wa) => {
        const lines = wa.text.split('\n');
        const longest = Math.max(...lines.map((l) => l.length), 1);
        const w = Math.min(468, longest * wa.sz * 0.62 + 10), hh = lines.length * wa.sz * 1.15 + 6;
        const shape = { t: 'shape', geom: 'rect', w: L.round(w, 2), h: L.round(hh, 2), fill: Object.assign({ a: 1 }, s.fill), line: s.line ? Object.assign({ dash: 'solid' }, s.line) : { t: 'none' }, warp: s.warp, shadow: s.shadow ? Object.assign({ c: '#808080', a: 0.6 }, s.shadow) : undefined, wordart: wa, name: 'LettersArt ' + D.nid() };
        E.edit('插入 LettersArt', () => O.insertItem(E.deleteSelection(), shape));
      }), 0);
    } }, { label: '取消' }] });
    const ok = dlg.buttons[0];
  };
  function editText(it, done) {
    const wa = Object.assign({ text: '在此输入文字', font: 'Arial Black', sz: 36 }, it.wordart || {});
    const font = ui.select(G.FONTS().map((f) => [f, f]), wa.font, () => upd(), { id: 'wa-font', translate: 'no' });
    const size = ui.select([8, 10, 12, 14, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96].map((x) => [x, String(x)]), wa.sz || 36, null, { id: 'wa-size' });
    const b = ui.check('加粗(&B)', !!wa.b, () => upd()), i = ui.check('倾斜(&I)', !!wa.i, () => upd());
    const ta = h('textarea', { rows: 5, id: 'wa-text', style: 'width:100%;font-size:20px' });
    ta.value = wa.text;
    const upd = () => { ta.style.fontFamily = L.fontStack(font.value); ta.style.fontWeight = b.input.checked ? 'bold' : 'normal'; ta.style.fontStyle = i.input.checked ? 'italic' : 'normal'; };
    upd();
    ui.dialog({ title: '编辑 LettersArt 文字', width: 440, focus: ta, body: h('div', { class: 'col' }, G.row(G.f('字体(&F)：', font), G.f('大小(&S)：', size), b, i), h('label', { for: 'wa-text', html: ui.ampHTML('文本(&T)：') }), ta), buttons: [{ label: '确定', primary: true, onClick: () => done({ text: ta.value || ' ', font: font.value, sz: +size.value, b: b.input.checked || undefined, i: i.input.checked || undefined }) }, { label: '取消' }] });
  }

  /* ================= Table dialogs ================= */
  G.insertTable = function () {
    const last = L.store.get('tableDims', { c: 5, r: 2 });
    const cols = G.num(last.c, { min: 1, max: 63, id: 'it-c' }), rows = G.num(last.r, { min: 1, max: 32767, id: 'it-r' });
    const fit = G.radios('it-fit', [['fixed', '固定列宽(&x)：'], ['contents', '根据内容调整表格(&F)'], ['window', '根据窗口调整表格(&d)']], 'fixed');
    const cw = G.len(null, { id: 'it-cw', autoLabel: L.N('自动'), min: 9 });
    let style = null, look = null;
    const styleNote = h('div', { class: 'tp-note', text: L.t('表格样式：{0}', '网格型') });
    const remember = ui.check('记住新表格的尺寸(&n)', false, null);
    const fEls = fit.querySelectorAll('.ck');
    ui.dialog({
      title: '插入表格', width: 330, focus: cols.input,
      body: h('div', { class: 'col' }, ui.group('表格尺寸', G.f('列数(&c)：', cols, 'wide'), G.f('行数(&r)：', rows, 'wide')), ui.group('自动调整操作', G.row(fEls[0], cw), fEls[1], fEls[2]), G.row(styleNote, ui.button('自动套用格式(&A)...', () => G.tableAutoFormat({ pick: (s, lk) => { style = s; look = lk; styleNote.textContent = L.t('表格样式：{0}', s); } }), { class: 'btn small' })), remember),
      buttons: [{ label: '确定', primary: true, onClick: () => {
        if (remember.input.checked) L.store.set('tableDims', { c: cols.get(), r: rows.get() });
        const o = { autofit: fit.get() };
        const w = cw.get();
        if (fit.get() === 'fixed' && w) o.colWidth = w;
        if (style) { o.style = L.tables.styleId(style); o.look = look; }
        L.tables.insert(rows.get(), cols.get(), o);
        if (style) L.tables.applyStyle(style, look);
      } }, { label: '取消' }],
    });
  };
  G.insertCells = function () {
    const r = G.radios('ic', [['right', '右移单元格(&i)'], ['down', '下移单元格(&d)'], ['row', '插入整行(&r)'], ['col', '插入整列(&c)']], 'down');
    ui.dialog({ title: '插入单元格', width: 240, body: r, buttons: [{ label: '确定', primary: true, onClick: () => L.tables.insertCells(r.get()) }, { label: '取消' }] });
  };
  G.deleteCells = function () {
    const r = G.radios('dc', [['left', '左移单元格(&l)'], ['up', '上移单元格(&u)'], ['row', '删除整行(&r)'], ['col', '删除整列(&c)']], 'left');
    ui.dialog({ title: '删除单元格', width: 240, body: r, buttons: [{ label: '确定', primary: true, onClick: () => L.tables.deleteCells(r.get()) }, { label: '取消' }] });
  };
  G.splitCells = function () {
    const c = G.num(2, { min: 1, max: 63, id: 'sc-c' }), r = G.num(1, { min: 1, max: 100, id: 'sc-r' });
    const merge = ui.check('拆分前合并单元格(&M)', !!E.cellSel, null);
    ui.dialog({ title: '拆分单元格', width: 260, focus: c.input, body: h('div', { class: 'col' }, G.f('列数(&c)：', c, 'wide'), G.f('行数(&r)：', r, 'wide'), merge), buttons: [{ label: '确定', primary: true, onClick: () => L.tables.split(r.get(), c.get(), merge.input.checked) }, { label: '取消' }] });
  };
  G.tableAutoFormat = function (opts) {
    opts = opts || {};
    const names = ['Table Normal', 'Table Grid'].concat(Object.keys(L.tables.STYLES).filter((n) => n !== 'Table Grid'));
    const ci = L.tables.cur();
    const curStyle = ci ? (doc().styles[ci.tbl.tblPr.style] || {}).name || '网格型' : 'Table Grid';
    const look = Object.assign({ firstRow: true, firstCol: true, lastRow: false, lastCol: false }, ci ? ci.tbl.tblPr.look || {} : {});
    const lb = G.listBox(names, names.includes(curStyle) ? curStyle : '网格型', { height: 200, translate: 'no', onChange: () => draw() });
    const pv = h('div', { class: 'preview-box', style: 'width:220px;height:170px' });
    const chk = (k, l) => ui.check(l, !!look[k], (v) => { look[k] = v; draw(); });
    const checks = [chk('firstRow', '标题行(&H)'), chk('firstCol', '首列(&s)'), chk('lastRow', '末行(&L)'), chk('lastCol', '末列(&l)')];
    function draw() {
      L.clear(pv);
      const name = lb.get();
      const def = L.tables.STYLES[name];
      const grid = [['', L.N('一月'), L.N('二月'), L.N('三月'), L.N('合计')], [L.N('东'), '7', '7', '5', '19'], [L.N('西'), '6', '4', '7', '17'], [L.N('南'), '8', '7', '9', '24'], [L.N('合计'), '21', '18', '21', '60']];
      const t = h('table', { style: 'border-collapse:collapse;font:10px Arial,Arimo,sans-serif;width:200px' });
      const bd = (b) => (b && b.val !== 'nil' ? `${Math.max(1, b.sz || 0.5)}px ${b.val === 'double' ? 'double' : b.val === 'dotted' ? 'dotted' : 'solid'} ${b.color && b.color !== 'auto' ? '#' + b.color : '#000'}` : 'none');
      grid.forEach((row, ri) => {
        const tr = h('tr');
        row.forEach((v, cix) => {
          const conds = R.cellConds({ tblPr: {} }, ri, cix, grid.length, row.length, Object.assign({ noVBand: !/Columns/.test(name), noHBand: !/List|Contemporary|Subtle|3D/.test(name) }, look));
          let rr = Object.assign({}, def && def.rPr), tc = Object.assign({}, def && def.tcPr), tb = Object.assign({}, def && def.tblPr && def.tblPr.borders);
          if (name === 'Table Grid') tb = { top: { val: 'single' }, bottom: { val: 'single' }, left: { val: 'single' }, right: { val: 'single' }, insideH: { val: 'single' }, insideV: { val: 'single' } };
          let cb = {};
          for (const c of conds) { const cf = def && def.cond && def.cond[c]; if (!cf) continue; if (cf.rPr) rr = Object.assign(rr, cf.rPr); if (cf.tcPr) { tc = Object.assign(tc, cf.tcPr); if (cf.tcPr.borders) cb = Object.assign(cb, cf.tcPr.borders); } }
          const side = (k, outer, inner, isOuter) => cb[k] || (isOuter ? tb[outer] : tb[inner]);
          const css = [`border-top:${bd(side('top', 'top', 'insideH', ri === 0))}`, `border-bottom:${bd(side('bottom', 'bottom', 'insideH', ri === grid.length - 1))}`, `border-left:${bd(side('left', 'left', 'insideV', cix === 0))}`, `border-right:${bd(side('right', 'right', 'insideV', cix === row.length - 1))}`, 'padding:2px 4px'];
          if (tc.shd) css.push('background:' + R.shadeCSS(tc.shd));
          if (rr.b) css.push('font-weight:bold'); if (rr.i) css.push('font-style:italic'); if (rr.color) css.push('color:#' + rr.color); if (rr.caps) css.push('text-transform:uppercase');
          tr.appendChild(h('td', { text: v, style: css.join(';') }));
        });
        t.appendChild(tr);
      });
      pv.appendChild(t);
    }
    draw();
    ui.dialog({
      title: '表格自动套用格式', width: 480, body: h('div', { class: 'col' }, h('div', { text: '类别：所有表格样式' }), G.row(h('div', { class: 'col', style: 'width:200px' }, h('label', { text: '表格样式：' }), lb), h('div', { class: 'col' }, h('div', { class: 'cd-cap', text: '预览' }), pv)), ui.group('特殊格式应用于', G.row(...checks))),
      buttons: [{ label: '应用(&A)', primary: true, onClick: () => { if (opts.pick) { opts.pick(lb.get(), Object.assign({}, look, { noVBand: true })); return; } L.tables.applyStyle(lb.get(), Object.assign({}, look, { noVBand: !/Columns/.test(lb.get()) }), { clearDirect: true }); } }, { label: '取消' }],
    });
  };
  G.tableProps = function (tab) {
    const ci = L.tables.cur();
    if (!ci) return;
    const t = ci.tbl;
    const d = doc();
    const { tp } = R.tblProps(d, t);
    const rg = L.tables.range();
    const total = t.grid.reduce((a, b) => a + b, 0);
    /* Table tab */
    const tW = ui.check('指定宽度(&P)：', !!(t.tblPr.w && t.tblPr.w.type !== 'auto' && t.tblPr.w.v), null);
    const tWv = G.len(t.tblPr.w && t.tblPr.w.type === 'dxa' ? t.tblPr.w.v : total, { id: 'tp-w', min: 9 });
    const tWu = ui.select([['dxa', '英寸/厘米'], ['pct', '百分比']], t.tblPr.w && t.tblPr.w.type === 'pct' ? 'pct' : 'dxa', null, { id: 'tp-wu' });
    const align = G.radios('tp-al', [['left', '左(&L)'], ['center', '居中(&C)'], ['right', '右(&h)']], tp.jc === 'center' ? 'center' : tp.jc === 'right' || tp.jc === 'end' ? 'right' : 'left');
    const indent = G.len(tp.ind || 0, { id: 'tp-ind' });
    const wrap = G.radios('tp-wr', [['none', '无(&N)'], ['around', '四周型(&u)']], t.tblPr.float ? 'around' : 'none');
    const mar = { t: G.len(tp.cellMar.t || 0, { id: 'tp-mt', min: 0 }), b: G.len(tp.cellMar.b || 0, { id: 'tp-mb', min: 0 }), l: G.len(tp.cellMar.l || 0, { id: 'tp-ml', min: 0 }), r: G.len(tp.cellMar.r || 0, { id: 'tp-mr', min: 0 }) };
    const spacing = G.len(t.tblPr.cellSpacing || 0, { id: 'tp-sp', min: 0 });
    const autoResize = ui.check('自动调整尺寸以适应内容(&z)', t.tblPr.layout !== 'fixed', null);
    const optsBtn = ui.button('选项(&O)...', () => ui.dialog({ title: '表格选项', width: 320, body: h('div', { class: 'col' }, ui.group('默认单元格边距', G.row(G.f('顶端(&T)：', mar.t, 'narrow'), G.f('左(&L)：', mar.l, 'narrow')), G.row(G.f('底端(&B)：', mar.b, 'narrow'), G.f('右(&R)：', mar.r, 'narrow'))), ui.group('默认单元格间距', G.f('允许单元格间距', spacing, 'wide')), ui.group('选项', autoResize)), buttons: [{ label: '确定', primary: true }] }), { class: 'btn small' });
    /* Row tab */
    const row = t.rows[rg.r0];
    const rH = ui.check('指定高度(&S)：', !!row.trPr.h, null);
    const rHv = G.len(row.trPr.h || 14.4, { id: 'tp-rh', min: 1 });
    const rRule = ui.select([['atLeast', '最小值'], ['exact', '固定值']], row.trPr.hRule === 'exact' ? 'exact' : 'atLeast', null, { id: 'tp-rr' });
    const rSplit = ui.check('允许跨页断行(&k)', !row.trPr.cantSplit, null);
    const rHead = ui.check('在各页顶端以标题行形式重复出现(&h)', !!row.trPr.header, null);
    /* Column tab */
    const cW = ui.check('指定宽度(&w)：', true, null);
    const cWv = G.len(t.grid.slice(rg.c0, rg.c1 + 1).reduce((a, b) => a + b, 0) / (rg.c1 - rg.c0 + 1), { id: 'tp-cw', min: 6 });
    /* Cell tab */
    const cell = ci.cell;
    const ceW = ui.check('指定宽度(&w)：', true, null);
    const ceWv = G.len(t.grid.slice(rg.map[ci.ri].find((m) => m.cell === cell).c0, rg.map[ci.ri].find((m) => m.cell === cell).c0 + (cell.tcPr.span || 1)).reduce((a, b) => a + b, 0), { id: 'tp-cew', min: 6 });
    const va = G.radios('tp-va', [['top', '顶端(&p)'], ['center', '居中(&C)'], ['bottom', '底端(&B)']], cell.tcPr.vAlign || 'top');
    const cm = cell.tcPr.mar || {};
    const cmSame = ui.check('与整张表格相同(&t)', !cell.tcPr.mar, null);
    const cmar = { t: G.len(cm.t != null ? cm.t : tp.cellMar.t || 0, { id: 'tp-cmt', min: 0 }), b: G.len(cm.b != null ? cm.b : tp.cellMar.b || 0, { id: 'tp-cmb', min: 0 }), l: G.len(cm.l != null ? cm.l : tp.cellMar.l || 0, { id: 'tp-cml', min: 0 }), r: G.len(cm.r != null ? cm.r : tp.cellMar.r || 0, { id: 'tp-cmr', min: 0 }) };
    const noWrap = ui.check('自动换行(&W)', !cell.tcPr.noWrap, null);
    const fitText = ui.check('适应文字(&F)', !!cell.tcPr.fitText, null);
    const cellOpts = ui.button('选项(&O)...', () => ui.dialog({ title: '单元格选项', width: 320, body: h('div', { class: 'col' }, ui.group('单元格边距', cmSame, G.row(G.f('顶端：', cmar.t, 'narrow'), G.f('左：', cmar.l, 'narrow')), G.row(G.f('底端：', cmar.b, 'narrow'), G.f('右：', cmar.r, 'narrow'))), ui.group('选项', noWrap, fitText)), buttons: [{ label: '确定', primary: true }] }), { class: 'btn small' });
    const tabs = ui.tabs([
      { label: '表格(&T)', body: h('div', { class: 'col' }, ui.group('大小', G.row(tW, tWv, G.f('度量单位：', tWu))), ui.group('对齐方式', G.row(align, G.f('左缩进(&I)：', indent, 'wide'))), ui.group('文字环绕', wrap), h('div', { class: 'tp-btns' }, ui.button('边框和底纹(&B)...', () => G.borders(), { class: 'btn small' }), optsBtn)) },
      { label: '行(&R)', body: h('div', { class: 'col' }, h('div', { class: 'cd-cap', text: rg.r1 > rg.r0 ? L.t('第 {0}–{1} 行：', rg.r0 + 1, rg.r1 + 1) : L.t('行 {0}：', rg.r0 + 1) }), ui.group('大小', G.row(rH, rHv, G.f('行高值为(&i)：', rRule))), ui.group('选项', rSplit, rHead)) },
      { label: '列(&u)', body: h('div', { class: 'col' }, h('div', { class: 'cd-cap', text: rg.c1 > rg.c0 ? L.t('第 {0}–{1} 列：', rg.c0 + 1, rg.c1 + 1) : L.t('列 {0}：', rg.c0 + 1) }), ui.group('大小', G.row(cW, cWv))) },
      { label: '单元格(&e)', body: h('div', { class: 'col' }, ui.group('大小', G.row(ceW, ceWv)), ui.group('垂直对齐', va), h('div', { class: 'tp-btns' }, cellOpts)) },
    ], tab || 0);
    ui.dialog({
      title: '表格属性', width: 430, body: tabs,
      buttons: [{ label: '确定', primary: true, onClick: () => {
        E.edit('表格属性', () => {
          D.touchTbl(t);
          const tpp = t.tblPr;
          if (tW.input.checked) {
            if (tWu.value === 'pct') tpp.w = { type: 'pct', v: 5000 };
            else { const w = tWv.get(); tpp.w = { type: 'dxa', v: w }; const tot = t.grid.reduce((a, b) => a + b, 0) || 1; t.grid = t.grid.map((g) => L.round((g * w) / tot, 2)); }
          } else tpp.w = { type: 'auto', v: 0 };
          tpp.jc = align.get() === 'left' ? undefined : align.get();
          if (!tpp.jc) delete tpp.jc;
          const ind = indent.get();
          if (align.get() === 'left' && ind) tpp.ind = ind; else delete tpp.ind;
          tpp.cellMar = { t: mar.t.get(), b: mar.b.get(), l: mar.l.get(), r: mar.r.get() };
          const sp = spacing.get(); if (sp) tpp.cellSpacing = sp; else delete tpp.cellSpacing;
          tpp.layout = autoResize.input.checked ? 'autofit' : 'fixed';
          for (let r = rg.r0; r <= rg.r1; r++) {
            const tr = t.rows[r].trPr;
            if (rH.input.checked) { tr.h = rHv.get(); tr.hRule = rRule.value; } else { delete tr.h; delete tr.hRule; }
            if (rSplit.input.checked) delete tr.cantSplit; else tr.cantSplit = true;
          }
          if (rHead.input.checked) for (let r = 0; r <= rg.r1; r++) t.rows[r].trPr.header = true; else for (let r = rg.r0; r <= rg.r1; r++) delete t.rows[r].trPr.header;
          if (cW.input.checked) { const w = cWv.get(); for (let i = rg.c0; i <= rg.c1; i++) t.grid[i] = L.round(w, 2); }
          for (const c of rg.cells) {
            c.tcPr.vAlign = va.get() === 'top' ? undefined : va.get();
            if (!c.tcPr.vAlign) delete c.tcPr.vAlign;
            if (cmSame.input.checked) delete c.tcPr.mar; else c.tcPr.mar = { t: cmar.t.get(), b: cmar.b.get(), l: cmar.l.get(), r: cmar.r.get() };
            if (noWrap.input.checked) delete c.tcPr.noWrap; else c.tcPr.noWrap = true;
            if (fitText.input.checked) c.tcPr.fitText = true; else delete c.tcPr.fitText;
          }
          if (ceW.input.checked && rg.cells.length === 1 && !cW.input.checked) {
            const m = rg.map[ci.ri].find((x) => x.cell === cell);
            const want = ceWv.get(), have = t.grid.slice(m.c0, m.c0 + m.span).reduce((a, b) => a + b, 0);
            const k = want / (have || 1);
            for (let i = m.c0; i < m.c0 + m.span; i++) t.grid[i] = L.round(t.grid[i] * k, 2);
          }
          L.tables.fixWidths(t);
          return E.sel;
        });
      } }, { label: '取消' }],
    });
  };
  G.sort = function () {
    const inTbl = L.tables.inTable();
    const ci = inTbl ? L.tables.cur() : null;
    const fields = inTbl ? ci.tbl.grid.map((_, i) => ['' + i, L.t('列 {0}', i + 1)]) : [['0', L.t('段落')]].concat([1, 2, 3, 4, 5].map((i) => ['' + i, L.t('域 {0}', i)]));
    const mk = (n, def) => ({ by: ui.select(n ? [['', '（无）']].concat(fields) : fields, def, null, { id: 'so-b' + n }), type: ui.select([['text', '文本'], ['number', '数字'], ['date', '日期']], 'text', null, { id: 'so-t' + n }), dir: G.radios('so-d' + n, [['asc', '升序(&A)'], ['desc', '降序(&D)']], 'asc') });
    const k1 = mk(0, fields[0][0]), k2 = mk(1, ''), k3 = mk(2, '');
    const hdr = G.radios('so-h', [['yes', '标题行(&r)'], ['no', '无标题行(&w)']], inTbl && ci.tbl.rows[0].trPr.header ? 'yes' : 'no');
    const sepSel = ui.select([['\t', '制表位'], [',', '逗号'], ['other', '其他']], '\t', null, { id: 'so-sep' });
    const caseS = ui.check('区分大小写(&s)', false, null);
    const row = (lbl, k) => ui.group(lbl, G.row(k.by, G.f('类型：', k.type), k.dir));
    ui.dialog({
      title: inTbl ? '排序' : '文字排序', width: 470,
      body: h('div', { class: 'col' }, row('排序依据(&S)', k1), row('次要关键字(&T)', k2), row('次要关键字(&b)', k3), ui.group('列表包含', hdr), G.row(G.f('字段分隔符：', sepSel), caseS)),
      buttons: [{ label: '确定', primary: true, onClick: () => {
        const keys = [k1, k2, k3].filter((k) => k.by.value !== '').map((k) => ({ col: +k.by.value, field: +k.by.value, type: k.type.value, desc: k.dir.get() === 'desc' }));
        if (inTbl) L.tables.sortTable(keys, { header: hdr.get() === 'yes', caseSensitive: caseS.input.checked });
        else L.tables.sortParagraphs(keys, { header: hdr.get() === 'yes', sep: sepSel.value === 'other' ? ';' : sepSel.value, caseSensitive: caseS.input.checked });
      } }, { label: '取消' }],
    });
  };
  G.formula = function () {
    let guess = '=SUM(ABOVE)';
    try { if (L.tables.inTable()) { const v = L.tables.evalFormula('SUM(ABOVE)', E.sel.f.p); const l = L.tables.evalFormula('SUM(LEFT)', E.sel.f.p); if (!v && l) guess = '=SUM(LEFT)'; } else guess = '='; } catch (e) { /* default */ }
    const inp = h('input', { type: 'text', id: 'fo-f', value: guess, style: 'width:100%' });
    const fmt = h('input', { type: 'text', id: 'fo-fmt', list: 'fo-fmts', style: 'width:100%' });
    const dl = h('datalist', { id: 'fo-fmts' }, ...['#,##0', '#,##0.00', '$#,##0.00;($#,##0.00)', '0', '0%', '0.00', '0.00%'].map((x) => h('option', { value: x })));
    const fn = ui.select([['', L.t('粘贴函数')]].concat(['ABS', 'AND', 'AVERAGE', 'COUNT', 'DEFINED', 'FALSE', 'IF', 'INT', 'MAX', 'MIN', 'MOD', 'NOT', 'OR', 'PRODUCT', 'ROUND', 'SIGN', 'SUM', 'TRUE'].map((x) => [x, x])), '', (v) => { if (!v) return; inp.value += v + '()'; inp.focus(); inp.setSelectionRange(inp.value.length - 1, inp.value.length - 1); fn.value = ''; }, { id: 'fo-fn', translate: 'no' });   // function names
    const bms = ui.select([['', L.t('粘贴书签')]].concat(Array.from(F().bookmarks().keys()).filter((n) => !n.startsWith('_')).map((n) => [n, n])), '', (v) => { if (v) { inp.value += v; bms.value = ''; } }, { id: 'fo-bm', translate: 'no' });
    ui.dialog({ title: '公式', width: 340, focus: inp, body: h('div', { class: 'col' }, h('label', { for: 'fo-f', html: ui.ampHTML('公式(&F)：') }), inp, dl, h('label', { for: 'fo-fmt', html: ui.ampHTML('数字格式(&N)：') }), fmt, G.row(fn, bms)), buttons: [{ label: '确定', primary: true, onClick: () => {
      let c = inp.value.trim();
      if (!c.startsWith('=')) c = '=' + c;
      if (fmt.value.trim()) c += ` \\# "${fmt.value.trim()}"`;
      A().insertFieldNow(c, L.N('公式'));
    } }, { label: '取消' }] });
  };
  G.textToTable = function () {
    const paras = E.selectedParas();
    const sepR = G.radios('tt-s', [['para', '段落(&P)'], [',', '逗号(&C)'], ['\t', '制表位(&T)'], ['other', '其他(&t)：']], paras.some((p) => p.runs.some((it) => it.t === 'tab')) ? '\t' : paras.some((p) => D.plainText(p).includes(',')) ? ',' : 'para', () => upd());
    const other = h('input', { type: 'text', maxlength: 1, value: '-', style: 'width:30px' });
    const cols = G.num(1, { min: 1, max: 63, id: 'tt-c', onChange: () => upd(true) });
    const rowsN = h('span');
    const fit = G.radios('tt-fit', [['fixed', '固定列宽(&x)'], ['contents', '根据内容调整表格(&F)'], ['window', '根据窗口调整表格(&d)']], 'fixed');
    const upd = (fromCols) => {
      const sep = sepR.get() === 'other' ? other.value || '-' : sepR.get();
      if (!fromCols) cols.set(sep === 'para' ? 1 : Math.max(...paras.map((p) => (sep === '\t' ? p.runs.filter((it) => it.t === 'tab').length : (D.plainText(p).split(sep).length - 1)) + 1)));
      rowsN.textContent = String(sep === 'para' ? Math.ceil(paras.length / cols.get()) : paras.length);
    };
    upd();
    const sEls = sepR.querySelectorAll('.ck');
    ui.dialog({ title: '文本转换成表格', width: 330, body: h('div', { class: 'col' }, ui.group('表格尺寸', G.f('列数(&c)：', cols, 'wide'), G.f('行数：', rowsN, 'wide')), ui.group('自动调整操作', fit), ui.group('文字分隔符', G.row(sEls[0], sEls[1]), G.row(sEls[2], sEls[3], other))), buttons: [{ label: '确定', primary: true, onClick: () => L.tables.textToTable({ sep: sepR.get() === 'other' ? other.value || '-' : sepR.get(), cols: cols.get(), autofit: fit.get() }) }, { label: '取消' }] });
  };
  G.tableToText = function () {
    const r = G.radios('tx-s', [['para', '段落标记(&P)'], ['\t', '制表位(&T)'], [',', '逗号(&C)'], ['other', '其他(&O)：']], '\t');
    const other = h('input', { type: 'text', maxlength: 1, value: '-', style: 'width:30px' });
    ui.dialog({ title: '表格转换成文本', width: 280, body: ui.group('文字分隔符', r, other), buttons: [{ label: '确定', primary: true, onClick: () => L.tables.tableToText(r.get() === 'other' ? other.value || '-' : r.get()) }, { label: '取消' }] });
  };
})();
