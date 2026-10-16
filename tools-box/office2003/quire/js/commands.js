/* Quire — command registry. Every menu item, toolbar button and shortcut runs one of these. */
(function () {
  'use strict';
  const L = window.L, D = L.D, O = L.O, E = L.ed, LY = L.layout, ui = L.ui;
  const A = (L.app = L.app || {});
  const C = (id, def) => ui.cmd(id, def);
  const doc = () => D.doc;
  const dlg = (name, ...args) => () => { if (L.dlg && L.dlg[name]) return L.dlg[name](...args); ui.msg('此对话框不可用。'); };
  const pane = (name) => () => { if (L.panes) { A.toggleTask(true); L.panes.task.show(name); } };
  const viewing = () => A.view === 'reading' || A.view === 'preview';
  const canEdit = () => !!D.doc && !E.readOnly && !viewing();
  const hasSel = () => !!E.sel && !E.collapsed();
  const inTable = () => !!(L.tables && E.sel && L.tables.inTable());
  const objSel = () => !!E.objSel;
  const floatSel = () => !!(E.objSel && E.objSel.float);
  const run = () => E.curRun();
  const para = () => E.curPara();
  const tbl = (fn, ...args) => () => L.tables && L.tables[fn](...args);

  /* ================= lists ================= */
  const LS = (L.lists = {});
  /** the list a paragraph belongs to: {numId, lvl, fmt} or null */
  LS.info = (p) => {
    const d = doc();
    const pp = D.pProps(d, p);
    if (!pp.num || !pp.num.id || pp.num.id === '0') return null;
    const lv = D.numLevelDef(d, pp.num.id, pp.num.lvl || 0);
    if (!lv) return null;
    return { numId: pp.num.id, lvl: pp.num.lvl || 0, fmt: lv.fmt, lv };
  };
  /** toolbar Bullets / Numbering: toggle, continuing an adjacent list of the same kind */
  LS.toggle = function (kind) {
    if (!E.sel) return;
    const d = doc();
    const paras = E.selectedParas();
    const first = paras[0];
    const cur = LS.info(first);
    const isKind = (inf) => inf && (kind === 'bullet' ? inf.fmt === 'bullet' : inf.fmt !== 'bullet');
    if (isKind(cur)) { LS.remove(paras); return; }
    /* continue the list right above when it has the same kind */
    const prev = D.prevPara(d, first);
    const pinf = prev && O.containerOf(prev) === O.containerOf(first) ? LS.info(prev) : null;
    E.edit(kind === 'bullet' ? L.N('项目符号') : L.N('编号'), () => {
      let numId;
      if (isKind(pinf)) numId = pinf.numId;
      else {
        D.touchKey(d, 'numbering');
        const last = A.lastList && A.lastList[kind];
        const abs = last ? L.clone(last) : kind === 'bullet' ? D.makeBulletAbs('•', 'Symbol', '') : D.makeNumberAbs('decimal', '%1.');
        numId = D.addNum(d, abs);
      }
      const lvl = isKind(pinf) ? pinf.lvl : 0;
      O.setParaProps(paras, (pPr) => { pPr.num = { id: numId, lvl: pPr.num && pPr.num.id !== '0' ? pPr.num.lvl || 0 : lvl }; delete pPr.ind; });
      return E.sel;
    });
  };
  LS.remove = function (paras) {
    paras = paras || E.selectedParas();
    const d = doc();
    E.edit('取消编号', () => {
      O.setParaProps(paras, (pPr, p) => {
        const sp = D.styleProps(d, pPr.style || 'Normal').pPr;
        if (sp.num && sp.num.id && sp.num.id !== '0') pPr.num = { id: '0', lvl: 0 }; else delete pPr.num;
        if (/^List(Bullet|Number)/.test(pPr.style || '')) delete pPr.style;
        void p;
        delete pPr.ind;
      });
      return E.sel;
    });
  };
  /** apply a list definition (from the Bullets and Numbering dialog). opts: {restart, kind} */
  LS.apply = function (abs, opts) {
    opts = opts || {};
    abs = L.preserve.materializeNumbering(abs);
    const d = doc();
    const paras = E.selectedParas();
    const first = paras[0];
    const prev = D.prevPara(d, first);
    const pinf = prev ? LS.info(prev) : null;
    E.edit('项目符号和编号', () => {
      D.touchKey(d, 'numbering');
      let numId;
      const cur = LS.info(first);
      if (cur && !opts.restart && !opts.newList) {
        /* redefine the list the paragraphs are in */
        const nd = D.numDef(d, cur.numId);
        const absNew = L.clone(abs);
        absNew.id = nd.num.abs;
        d.numbering.abs[nd.num.abs] = absNew;
        numId = cur.numId;
      } else if (opts.continuePrev && pinf) numId = pinf.numId;
      else numId = D.addNum(d, L.clone(abs));
      if (opts.restart) { const n = d.numbering.nums[numId]; n.ov = { 0: { start: opts.start || 1 } }; }
      O.setParaProps(paras, (pPr) => { pPr.num = { id: numId, lvl: pPr.num && pPr.num.id !== '0' && pPr.num.id ? pPr.num.lvl || 0 : 0 }; delete pPr.ind; });
      return E.sel;
    });
    A.lastList = A.lastList || {};
    A.lastList[abs.levels && abs.levels[0] && abs.levels[0].fmt === 'bullet' ? 'bullet' : 'number'] = L.clone(abs);
  };
  /** change list level of selected list paragraphs */
  LS.shift = function (delta) {
    const d = doc();
    const paras = E.selectedParas().filter((p) => LS.info(p));
    if (!paras.length) return false;
    E.edit(L.N('编号'), () => {
      for (const p of paras) { const inf = LS.info(p); D.touch(p); p.pPr.num = { id: inf.numId, lvl: L.clamp(inf.lvl + delta, 0, 8) }; delete p.pPr.ind; }
      return E.sel;
    });
    void d;
    return true;
  };
  LS.restartAt = function (start) {
    const d = doc();
    const p = E.sel.f.p;
    const inf = LS.info(p);
    if (!inf) return;
    E.edit('重新开始编号', () => {
      D.touchKey(d, 'numbering');
      const absId = d.numbering.nums[inf.numId].abs;
      const nid = D.addNumInstance(d, absId, true);
      d.numbering.nums[nid].ov = { [inf.lvl]: { start: start || 1 } };
      /* this paragraph and the rest of the same list switch to the new instance */
      let q = p;
      while (q) { const qi = LS.info(q); if (qi && qi.numId === inf.numId) { D.touch(q); q.pPr.num = { id: nid, lvl: qi.lvl }; } else if (D.plen(q) && !qi) break; q = D.nextPara(d, q); }
      return E.sel;
    });
  };
  LS.continuePrev = function () {
    const d = doc();
    const p = E.sel.f.p;
    const inf = LS.info(p);
    if (!inf) return;
    let q = D.prevPara(d, p), target = null;
    while (q) { const qi = LS.info(q); if (qi && qi.numId !== inf.numId && qi.fmt !== 'bullet') { target = qi; break; } q = D.prevPara(d, q); }
    if (!target) { A.status('没有可接续的前一列表。'); return; }
    E.edit('继续编号', () => {
      let r = p;
      while (r) { const ri = LS.info(r); if (ri && ri.numId === inf.numId) { D.touch(r); r.pPr.num = { id: target.numId, lvl: ri.lvl }; } r = D.nextPara(d, r); }
      return E.sel;
    });
  };

  /* ================= helpers ================= */
  const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
  function stepSize(dir, one) {
    const cur = E.uniformRun('sz') || run().sz || 12;
    let v;
    if (one) v = cur + dir;
    else if (dir > 0) v = SIZES.find((s) => s > cur) || Math.round(cur / 10 + 1) * 10;
    else v = SIZES.slice().reverse().find((s) => s < cur) || Math.max(1, cur - 1);
    A.setFontSize(v);
  }
  function setAlign(jc) {
    E.formatPara((pPr) => { const cur = pPr.jc; pPr.jc = cur === jc && jc !== 'left' ? 'left' : jc; }, L.N('对齐'));
  }
  const alignIs = (jc) => { const v = para().jc || 'left'; return jc === 'left' ? v === 'left' || v === 'start' : jc === 'both' ? v === 'both' || v === 'distribute' : v === jc; };
  function indent(delta) {
    if (LS.info(E.sel.f.p) && LS.shift(delta > 0 ? 1 : -1)) return;
    const defTab = doc().settings.defTab || 36;
    E.formatPara((pPr, p) => {
      const cur = D.pProps(doc(), p).ind || {};
      const l = cur.l || 0;
      const nl = delta > 0 ? (Math.floor(l / defTab + 1e-6) + 1) * defTab : Math.max(0, (Math.ceil(l / defTab - 1e-6) - 1) * defTab);
      pPr.ind = Object.assign({}, pPr.ind, { l: nl });
    }, delta > 0 ? '增加缩进量' : '减少缩进量');
  }
  function hanging(delta) {
    const defTab = doc().settings.defTab || 36;
    E.formatPara((pPr, p) => {
      const cur = D.pProps(doc(), p).ind || {};
      const l = Math.max(0, (cur.l || 0) + delta * defTab);
      const fl = l > 0 ? -Math.min(l, delta > 0 ? Math.max(defTab, -(cur.fl || 0) + defTab) : Math.max(0, -(cur.fl || 0) - defTab)) : 0;
      pPr.ind = Object.assign({}, pPr.ind, { l, fl: delta > 0 ? -Math.max(defTab, -(cur.fl || 0)) : fl });
    }, '悬挂缩进');
  }
  function lineSpacing(mult) { E.formatPara((pPr) => { pPr.sp = Object.assign({}, pPr.sp, { line: mult, rule: 'auto' }); }, '行距'); }
  function applyStyle(id) {
    const d = doc();
    if (!d.styles[id]) {
      const b = D.builtinStyles()[id];
      if (!b) return;
      E.edit(L.N('样式'), () => { D.touchKey(d, 'styles'); d.styles[id] = b; D.stylesChanged(); return E.sel; });
    }
    const [a, b] = E.range();
    E.edit(L.N('样式'), () => { O.applyStyle(a, b, id); return E.sel; });
  }
  A.applyStyle = applyStyle;
  function insertFieldNow(instr, label) {
    const res = L.fields ? L.fields.evaluate(instr, E.sel.f) : '';
    E.edit(label || '插入域', () => {
      const pos = E.deleteSelection();
      return O.insertField(pos, instr, res == null ? '' : String(res), O.inheritRPr(pos));
    });
    if (L.fields && /^\s*(PAGE|NUMPAGES|SECTIONPAGES|SECTION)\b/i.test(instr)) LY.refreshDynamic();
  }
  A.insertFieldNow = insertFieldNow;
  const typeChar = (ch) => () => E.typeText(ch);
  function clip() {
    try {
      if (document.execCommand('copy')) return true;
    } catch (e) { /* not allowed */ }
    return false;
  }
  A.clipItems = [];
  L.bus.on('clip', (c) => { A.clipItems.unshift(c); A.clipItems.length = Math.min(A.clipItems.length, 24); if (L.panes && L.panes.task.current === 'clipboard') L.panes.task.show('clipboard'); });
  async function pasteCmd() {
    if (!E.sel) return;
    if (E.clip) { E.pasteBlocks(E.clip.blocks); return; }
    try {
      if (navigator.clipboard && navigator.clipboard.readText) {
        const t = await navigator.clipboard.readText();
        if (t) { E.pasteText(t); return; }
      }
    } catch (e) { /* reading the clipboard is often blocked */ }
    ui.msg('浏览器不允许网页通过菜单读取剪贴板。\n请按 Ctrl+V 粘贴。', { icon: 'info' });
  }
  A.paste = pasteCmd;
  const selectedObjIs = (pred) => { const it = L.drawing && L.drawing.selectedItem(); return !!it && pred(it); };

  /* ================= File ================= */
  C('new', { label: '新建(&N)...', menuLabel: '新建(&N)...', icon: 'new', key: 'Ctrl+N', tip: '新建空白文档', run: (arg, ev) => { if (ev && ev.currentTarget && ev.currentTarget.closest && ev.currentTarget.closest('.toolbar')) A.newDocument(); else if (window.VO) VO.newDialog('quire', (id) => (id ? L.templates.create(id) : A.newDocument())); else pane('new')(); } });
  C('newBlank', { label: '空白文档(&B)', icon: 'new', run: () => A.newDocument() });
  C('open', { label: '打开(&O)...', icon: 'open', key: 'Ctrl+O', run: () => A.openDialog() });
  C('close', { label: '关闭(&C)', icon: 'close', run: () => A.closeDocument() });
  C('save', { label: '保存(&S)', icon: 'save', key: 'Ctrl+S', enabled: () => !!D.doc, run: () => A.save() });
  C('saveAs', { label: '另存为(&A)...', key: 'F12', run: () => A.saveAs() });
  C('saveWeb', { label: '另存为网页(&g)...', icon: 'webPage', run: () => A.exportAs('html') });
  C('fileSearch', { label: '文件搜索(&h)...', icon: 'find', run: pane('search') });
  C('webPreview', { label: '网页预览(&b)', icon: 'webPage', run: () => A.setView('web') });
  C('pageSetup', { label: '页面设置(&u)...', icon: 'pageSetup', run: dlg('pageSetup') });
  C('printPreview', { label: '打印预览(&v)', icon: 'preview', key: 'Ctrl+F2', checked: () => A.view === 'preview', run: () => A.setView(A.view === 'preview' ? A.prevView || 'print' : 'preview') });
  C('print', { label: '打印(&P)...', icon: 'print', key: 'Ctrl+P', tip: '打印', run: (arg, ev) => { if (ev && ev.currentTarget && ev.currentTarget.closest && ev.currentTarget.closest('.toolbar')) A.printPDF(); else dlg('print')(); } });
  C('exportPDF', { label: 'PDF 文档(&P)...', icon: 'print', run: () => A.printPDF() });
  C('sendMail', { label: '邮件收件人（作为附件）(&M)...', icon: 'mail', tip: '电子邮件', run: async () => { await A.exportAs('docx'); A.status('已保存可用作电子邮件附件的副本。', 6000); } });
  C('permission', { label: '权限(&P)', icon: 'permission', tip: '权限（不限制访问）', run: pane('protect') });
  C('properties', { label: '属性(&i)', icon: 'properties', run: dlg('properties') });
  /* close every document (asking to save each), then return to the Start Center */
  C('exitApp', { label: '退出(&x)', run: async () => { for (;;) { const last = A.docs.length <= 1; if (!(await A.closeDocument())) return; if (last) break; } if (window.VO) VO.exit(); } });

  /* ================= Edit ================= */
  C('undo', { get label() { return D.canUndo() ? L.t('撤消 {0}(&U)', L.t(D.undoLabel())) : "无法撤消(&C)"; }, tip: '撤消', icon: 'undo', key: 'Ctrl+Z', enabled: () => D.canUndo() && !viewing(), run: () => { D.undo(); } });
  C('redo', { get label() { return D.canRedo() ? L.t('重做 {0}(&R)', L.t(D.redoLabel())) : "重复(&R)"; }, tip: '重做', icon: 'redo', key: 'Ctrl+Y', enabled: () => D.canRedo() && !viewing(), run: () => { D.redo(); } });
  C('cut', { label: '剪切(&t)', icon: 'cut', key: 'Ctrl+X', enabled: () => canEdit() && (hasSel() || objSel()), run: () => { E.focus(); if (!clip()) { const d = E.copyData(); if (d) { E.clip = d; L.bus.emit('clip', d); } } E.edit(L.N('剪切'), () => E.deleteSelection()); } });
  C('copy', { label: '复制(&C)', icon: 'copy', key: 'Ctrl+C', enabled: () => hasSel() || objSel(), run: () => { E.focus(); const d = E.copyData(); if (d) { E.clip = d; L.bus.emit('clip', d); } clip(); } });
  C('officeClipboard', { label: 'Office 剪贴板(&b)...', icon: 'clipboard', run: pane('clipboard') });
  C('paste', { label: '粘贴(&P)', icon: 'paste', key: 'Ctrl+V', enabled: () => canEdit(), run: () => pasteCmd() });
  C('pasteSpecial', { label: '选择性粘贴(&S)...', enabled: () => canEdit(), run: dlg('pasteSpecial') });
  C('pasteHyperlink', { label: '粘贴为超链接(&H)', enabled: () => canEdit() && !!E.clip, run: () => { const t = E.clip.text.trim(); if (!t) return; const url = /^(https?:|mailto:|www\.)/i.test(t) ? (/^www\./i.test(t) ? 'http://' + t : t) : null; E.edit('粘贴为超链接', () => { const pos = E.deleteSelection(); return O.insertText(pos, t, Object.assign(O.inheritRPr(pos), { style: 'Hyperlink', link: url ? { url } : { anchor: '' } })); }); } });
  C('clearFormats', { label: '格式(&F)', enabled: canEdit, run: () => { const [a, b] = E.range(); E.edit('清除格式', () => { if (!D.eqPos(a, b)) O.formatRange(a, b, (r) => { const k = {}; if (r.link) k.link = r.link; if (r.ins) k.ins = r.ins; if (r.del) k.del = r.del; return k; }); O.setParaProps(E.selectedParas(), (pPr) => { const k = {}; if (pPr.num) k.num = pPr.num; if (pPr.sect) k.sect = pPr.sect; return k; }); return E.sel; }); } });
  C('clearContents', { label: '内容(&C)', key: 'Del', enabled: () => canEdit() && hasSel(), run: () => E.edit(L.N('清除'), () => E.deleteSelection()) });
  C('selectAll', { label: '全选(&l)', key: 'Ctrl+A', icon: 'selectAll', run: () => { E.focus(); E.selectAll(); } });
  C('find', { label: '查找(&F)...', icon: 'find', key: 'Ctrl+F', run: dlg('find', 'find') });
  C('replace', { label: '替换(&e)...', icon: 'replace', key: 'Ctrl+H', enabled: () => !viewing(), run: dlg('find', 'replace') });
  C('goTo', { label: '定位(&G)...', key: 'Ctrl+G', run: dlg('find', 'goto') });
  C('links', { label: '链接(&k)...', enabled: () => false, run: () => {} });
  C('editObject', { get label() { const it = L.drawing && L.drawing.selectedItem(); return it && it.t === 'chart' ? '图表对象(&O)' : it && it.wordart ? 'LettersArt 对象(&O)' : '对象(&O)'; }, enabled: () => selectedObjIs((it) => it.t === 'chart' || !!it.wordart), run: () => A.formatObject(true) });

  /* ================= View ================= */
  C('viewNormal', { label: '普通(&N)', icon: 'normalView', key: 'Ctrl+Alt+N', radio: true, checked: () => A.view === 'normal', run: () => A.setView('normal') });
  C('viewWeb', { label: 'Web 版式(&W)', icon: 'webView', radio: true, checked: () => A.view === 'web', run: () => A.setView('web') });
  C('viewPrint', { label: '页面视图(&P)', icon: 'printView', key: 'Ctrl+Alt+P', radio: true, checked: () => A.view === 'print', run: () => A.setView('print') });
  C('viewReading', { label: '阅读版式(&e)', icon: 'read', tip: '阅读', radio: true, checked: () => A.view === 'reading', run: () => A.setView(A.view === 'reading' ? A.prevView || 'print' : 'reading') });
  C('viewOutline', { label: '大纲(&O)', icon: 'outlineView', key: 'Ctrl+Alt+O', radio: true, checked: () => A.view === 'outline', run: () => A.setView('outline') });
  C('taskPane', { label: '任务窗格(&k)', icon: 'taskpane', key: 'Ctrl+F1', checked: () => A.taskOpen(), run: () => A.toggleTask() });
  const TBNAMES = { standard: L.N('常用'), formatting: L.N('格式设置'), drawing: L.N('绘图'), tables: '表格和边框', reviewing: L.N('审阅'), picture: L.N('图片'), outlining: L.N('大纲'), headerFooter: '页眉和页脚', mailMerge: '邮件合并', wordart: L.N('LettersArt') };
  for (const k in TBNAMES) C('tb_' + k, { label: TBNAMES[k], tip: TBNAMES[k], checked: () => A.tbShown(k), run: () => A.toggleToolbar(k) });
  C('ruler', { label: '标尺(&R)', icon: 'ruler', checked: () => !!A.opts.ruler, run: () => { A.opts.ruler = !A.opts.ruler; A.saveOpts(); A.updateRulers(); } });
  C('docMap', { label: '文档结构图(&D)', icon: 'docMap', checked: () => !!A.opts.docMap, run: () => A.toggleDocMap() });
  C('thumbnails', { label: '缩略图(&m)', icon: 'thumbnails', checked: () => !!A.opts.thumbs, run: () => A.toggleThumbs() });
  C('headerFooter', { label: '页眉和页脚(&H)', icon: 'headerFooter', checked: () => !!LY.hfEdit, run: () => (LY.hfEdit ? A.closeHeaderFooter() : A.editHeaderFooter('hdr')) });
  C('footnotes', { label: '脚注(&n)', enabled: () => Object.keys(doc().fn).length + Object.keys(doc().en).length > 0, run: () => { const k = Object.keys(doc().fn)[0]; if (k) A.gotoNote('fn:' + k); else { const e = Object.keys(doc().en)[0]; if (e) A.gotoNote('en:' + e); } } });
  C('markup', { label: '标记(&M)', checked: () => LY.opts.markup !== 'final' && LY.opts.markup !== 'original' && LY.opts.showComments, run: () => { const on = !(LY.opts.markup !== 'final' && LY.opts.markup !== 'original' && LY.opts.showComments); LY.opts.markup = on ? 'markup' : 'final'; LY.opts.showComments = on; LY.render(); E.restoreDom(true); } });
  C('fullScreen', { label: '全屏显示(&u)', icon: 'fullScreen', checked: () => L.$('#app').classList.contains('fullscreen'), run: () => A.toggleFullScreen() });
  C('zoomDlg', { label: '缩放(&Z)...', icon: 'zoom', run: dlg('zoom') });
  C('showMarks', { label: '显示/隐藏 ¶', icon: 'showMarks', key: 'Ctrl+*', tip: '显示/隐藏 ¶', checked: () => !!LY.opts.showMarks, run: () => { LY.opts.showMarks = !LY.opts.showMarks; LY.root.classList.toggle('marks', LY.opts.showMarks); LY.render(); E.restoreDom(true); } });
  C('gridlines', { label: '显示网格线(&G)', icon: 'grid', checked: () => LY.root && LY.root.classList.contains('gridlines'), run: () => { LY.root.classList.toggle('gridlines'); A.opts.gridlines = LY.root.classList.contains('gridlines'); A.saveOpts(); } });

  /* ================= Insert ================= */
  C('insertBreak', { label: '分隔符(&B)...', enabled: canEdit, run: dlg('breakDlg') });
  C('pageNumbers', { label: '页码(&u)...', enabled: canEdit, run: dlg('pageNumbers') });
  C('dateTime', { label: '日期和时间(&T)...', icon: 'date', enabled: canEdit, run: dlg('dateTime') });
  C('insertField', { label: '域(&F)...', enabled: canEdit, run: dlg('field') });
  C('insertSymbol', { label: '符号(&S)...', icon: 'symbol', enabled: canEdit, run: dlg('symbol') });
  C('insertComment', { label: '批注(&m)', icon: 'comment', key: 'Ctrl+Alt+M', tip: '插入批注', enabled: () => !!D.doc && !viewing(), run: () => L.review && L.review.insertComment() });
  C('insertFootnote', { label: '脚注(&F)...', icon: 'footnote', enabled: canEdit, run: dlg('footnote') });
  C('insertFootnoteNow', { label: '插入脚注', enabled: canEdit, run: () => L.notes && L.notes.insert('fn') });
  C('insertEndnoteNow', { label: '插入尾注', enabled: canEdit, run: () => L.notes && L.notes.insert('en') });
  C('insertCaption', { label: '题注(&C)...', enabled: canEdit, run: dlg('caption') });
  C('crossReference', { label: '交叉引用(&r)...', enabled: canEdit, run: dlg('crossRef') });
  C('indexTables', { label: '索引和目录(&d)...', enabled: canEdit, run: dlg('indexTables') });
  C('insertClipArt', { label: '剪贴画(&C)...', icon: 'clipart', tip: '插入剪贴画', enabled: canEdit, run: pane('clipart') });
  C('insertPicture', { label: '来自文件(&F)...', icon: 'picture', tip: '插入图片', enabled: canEdit, run: () => A.insertPictureDialog() });
  C('newDrawing', { label: '新建绘图(&N)', icon: 'drawing', enabled: canEdit, run: () => A.openAutoShapes() });
  C('insertWordArt', { label: 'LettersArt(&L)...', icon: 'wordart', tip: '插入 LettersArt', keys: 'WordArt', enabled: canEdit, run: dlg('wordart') });
  C('insertChart', { label: '图表(&h)', icon: 'chart', enabled: canEdit, run: () => A.insertChart() });
  C('insertExcel', { label: '插入电子表格', icon: 'excel', tip: '插入电子表格', enabled: canEdit, run: () => L.tables.insert(4, 5, { worksheet: true }) });
  C('insertDiagram', { label: '图示(&g)...', icon: 'diagram', tip: '插入图示或组织结构图', enabled: canEdit, run: dlg('diagram') });
  C('textBox', { label: '文本框(&x)', icon: 'textbox', tip: '文本框', enabled: canEdit, run: () => L.drawtool && L.drawtool.start({ geom: 'rect', textbox: true }) });
  C('insertFile', { label: '文件(&e)...', enabled: canEdit, run: () => A.insertFileDialog() });
  C('insertObject', { label: '对象(&O)...', enabled: canEdit, run: dlg('insertObject') });
  C('bookmark', { label: '书签(&k)...', key: 'Ctrl+Shift+F5', run: dlg('bookmark') });
  C('hyperlink', { label: '超链接(&i)...', icon: 'hyperlink', key: 'Ctrl+K', tip: '插入超链接', enabled: canEdit, run: dlg('hyperlink') });
  C('removeHyperlink', { label: '删除超链接(&R)', enabled: canEdit, run: () => A.removeHyperlink() });

  /* ================= Format ================= */
  C('fontDlg', { label: '字体(&F)...', icon: 'font', key: 'Ctrl+D', enabled: canEdit, run: dlg('font') });
  C('paragraphDlg', { label: '段落(&P)...', icon: 'paragraph', enabled: canEdit, run: dlg('paragraph') });
  C('bulletsDlg', { label: '项目符号和编号(&N)...', icon: 'numbering', enabled: canEdit, run: dlg('bullets') });
  C('bordersDlg', { label: '边框和底纹(&B)...', icon: 'bordersOut', enabled: canEdit, run: dlg('borders') });
  C('columnsDlg', { label: '分栏(&C)...', icon: 'columns', enabled: canEdit, run: dlg('columns') });
  C('tabsDlg', { label: '制表位(&T)...', enabled: canEdit, run: dlg('tabs') });
  C('dropCap', { label: '首字下沉(&D)...', icon: 'dropCap', enabled: canEdit, run: dlg('dropCap') });
  C('textDirection', { label: '文字方向(&x)...', icon: 'textDir', tip: '更改文字方向', enabled: canEdit, run: () => (inTable() ? L.tables.cycleTextDirection() : dlg('textDirection')()) });
  C('changeCase', { label: '更改大小写(&e)...', icon: 'changeCase', enabled: canEdit, run: dlg('changeCase') });
  C('changeCaseCycle', { label: '更改大小写', key: 'Shift+F3', enabled: () => canEdit(), run: () => A.cycleCase() });
  C('printedWatermark', { label: '水印(&W)...', enabled: canEdit, run: dlg('watermark') });
  C('themeDlg', { label: '主题(&h)...', enabled: canEdit, run: dlg('theme') });
  C('autoFormat', { label: '自动套用格式(&A)...', enabled: canEdit, run: dlg('autoFormatDoc') });
  C('stylesPane', { label: '样式和格式(&S)...', icon: 'stylesPane', tip: '样式和格式', key: 'Ctrl+Alt+Shift+S', run: pane('styles') });
  C('revealFormatting', { label: '显示格式(&v)...', icon: 'reveal', key: 'Shift+F1', run: pane('reveal') });
  C('formatObject', { get label() { const it = L.drawing && L.drawing.selectedItem(); return it ? (it.t === 'img' ? '图片(&i)...' : it.tb ? '文本框(&o)...' : it.wordart ? 'LettersArt(&A)...' : it.t === 'chart' ? '图表(&a)...' : '自选图形(&p)...') : '对象(&e)...'; }, icon: 'formatPicture', tip: '设置对象格式', enabled: () => objSel(), run: () => A.formatObject() });
  /* character */
  C('bold', { label: '加粗', icon: 'bold', key: 'Ctrl+B', enabled: canEdit, checked: () => !!run().b, run: () => E.toggleRun('b', undefined, L.N('加粗')) });
  C('italic', { label: '倾斜', icon: 'italic', key: 'Ctrl+I', enabled: canEdit, checked: () => !!run().i, run: () => E.toggleRun('i', undefined, L.N('倾斜')) });
  C('underline', { label: '下划线', icon: 'underline', key: 'Ctrl+U', enabled: canEdit, checked: () => !!run().u && run().u !== 'none', run: () => { const u = run().u; E.formatRun({ u: u && u !== 'none' ? 'none' : 'single' }, L.N('下划线')); } });
  C('dblUnderline', { label: '双下划线', enabled: canEdit, run: () => { const u = run().u; E.formatRun({ u: u === 'double' ? 'none' : 'double' }, L.N('下划线')); } });
  C('wordUnderline', { label: '字下划线', enabled: canEdit, run: () => { const u = run().u; E.formatRun({ u: u === 'words' ? 'none' : 'words' }, L.N('下划线')); } });
  C('strike', { label: '删除线', icon: 'strike', enabled: canEdit, checked: () => !!run().strike, run: () => E.toggleRun('strike', undefined, L.N('删除线')) });
  C('subscript', { label: '下标', icon: 'subscript', key: 'Ctrl+=', enabled: canEdit, checked: () => run().vert === 'subscript', run: () => E.formatRun({ vert: run().vert === 'subscript' ? undefined : 'subscript' }, L.N('下标')) });
  C('superscript', { label: '上标', icon: 'superscript', key: 'Ctrl+Shift++', enabled: canEdit, checked: () => run().vert === 'superscript', run: () => E.formatRun({ vert: run().vert === 'superscript' ? undefined : 'superscript' }, L.N('上标')) });
  C('allCaps', { label: '全部大写', enabled: canEdit, checked: () => !!run().caps, run: () => E.toggleRun('caps', undefined, '全部大写') });
  C('smallCaps', { label: '小型大写字母', enabled: canEdit, checked: () => !!run().smallCaps, run: () => E.toggleRun('smallCaps', undefined, '小型大写字母') });
  C('hiddenText', { label: '隐藏', enabled: canEdit, checked: () => !!run().hidden, run: () => E.toggleRun('hidden', undefined, L.N('隐藏')) });
  C('resetChar', { label: '重置字符格式', key: 'Ctrl+Space', enabled: canEdit, run: () => { if (E.collapsed()) { E.pending = { b: false, i: false, u: undefined, strike: undefined, vert: undefined, caps: undefined, smallCaps: undefined, color: undefined, hl: undefined, font: undefined, sz: undefined }; L.bus.emit('sel', E.sel); return; } const [a, b] = E.range(); E.edit('重置字符', () => { O.formatRange(a, b, (r) => { const k = {}; for (const x of ['link', 'ins', 'del', 'style']) if (r[x]) k[x] = r[x]; if (k.style && !/^(Hyperlink|FootnoteReference|EndnoteReference|CommentReference)$/.test(k.style)) delete k.style; return k; }); return E.sel; }); } });
  C('resetPara', { label: '重置段落格式', key: 'Ctrl+Q', enabled: canEdit, run: () => E.formatPara((pPr) => { const k = {}; for (const x of ['style', 'num', 'sect', 'frame']) if (pPr[x]) k[x] = pPr[x]; return k; }, '重置段落') });
  C('growFont', { label: '增大字号', icon: 'growFont', key: 'Ctrl+Shift+>', enabled: canEdit, run: () => stepSize(1) });
  C('shrinkFont', { label: '减小字号', icon: 'shrinkFont', key: 'Ctrl+Shift+<', enabled: canEdit, run: () => stepSize(-1) });
  C('growFont1', { label: '增大字号 1 磅', key: 'Ctrl+]', enabled: canEdit, run: () => stepSize(1, true) });
  C('shrinkFont1', { label: '减小字号 1 磅', key: 'Ctrl+[', enabled: canEdit, run: () => stepSize(-1, true) });
  C('focusFont', { label: '字体', run: () => { const i = L.$('#tb-font'); if (i) i.focus(); } });
  C('focusSize', { label: '字号', run: () => { const i = L.$('#tb-size'); if (i) i.focus(); } });
  C('focusStyle', { label: '样式', run: () => { const i = L.$('#tb-style'); if (i) i.focus(); } });
  C('copyFormat', { label: '复制格式', run: () => { A.copiedFormat = A.pickFormat(); } });
  C('pasteFormat', { label: '粘贴格式', enabled: () => canEdit() && !!A.copiedFormat, run: () => A.paintFormat(A.copiedFormat) });
  C('painter', { label: '格式刷', icon: 'painter', tip: '格式刷', enabled: canEdit, checked: () => A.painter.on(), run: (arg, ev) => A.painter.click(ev) });
  C('fontColorBtn', { label: '字体颜色', icon: 'fontColor', tip: '字体颜色', enabled: canEdit, swatch: () => '#' + A.lastFontColor, run: () => A.applyFontColor(A.lastFontColor) });
  C('highlightBtn', { label: '突出显示', icon: 'highlight', tip: '突出显示', enabled: canEdit, checked: () => A.highlighter.on(), swatch: () => L.R.HIGHLIGHT[A.lastHighlight] || '#ffff00', run: () => { if (A.highlighter.on()) A.highlighter.stop(); else A.applyHighlight(A.lastHighlight); } });
  C('borderBtn', { get label() { const b = A.BORDER_ITEMS.find((x) => x[0] === A.lastBorder); return b ? b[2] : '外侧框线'; }, get icon() { const b = A.BORDER_ITEMS.find((x) => x[0] === A.lastBorder); return b ? b[1] : 'bordersOut'; }, enabled: canEdit, run: () => A.applyBorder(A.lastBorder) });
  /* paragraph */
  C('alignLeft', { label: '左对齐', icon: 'alignLeft', key: 'Ctrl+L', enabled: canEdit, checked: () => alignIs('left'), run: () => setAlign('left') });
  C('alignCenter', { label: '居中', icon: 'alignCenter', key: 'Ctrl+E', enabled: canEdit, checked: () => alignIs('center'), run: () => setAlign('center') });
  C('alignRight', { label: '右对齐', icon: 'alignRight', key: 'Ctrl+R', enabled: canEdit, checked: () => alignIs('right'), run: () => setAlign('right') });
  C('justify', { label: '两端对齐', icon: 'justify', key: 'Ctrl+J', enabled: canEdit, checked: () => alignIs('both'), run: () => setAlign('both') });
  C('single', { label: '单倍行距', key: 'Ctrl+1', enabled: canEdit, run: () => lineSpacing(1) });
  C('onePointFive', { label: '1.5 倍行距', key: 'Ctrl+5', enabled: canEdit, run: () => lineSpacing(1.5) });
  C('double', { label: '双倍行距', key: 'Ctrl+2', enabled: canEdit, run: () => lineSpacing(2) });
  C('spaceBefore', { label: '段前间距', key: 'Ctrl+0', enabled: canEdit, run: () => E.formatPara((pPr, p) => { const cur = (D.pProps(doc(), p).sp || {}).b || 0; pPr.sp = Object.assign({}, pPr.sp, { b: cur >= 12 ? 0 : 12, bAuto: undefined }); }, '段前间距') });
  C('numbering', { label: '编号', icon: 'numbering', tip: '编号', enabled: canEdit, checked: () => { const i = E.sel && LS.info(E.sel.f.p); return !!i && i.fmt !== 'bullet'; }, run: () => LS.toggle('number') });
  C('bullets', { label: '项目符号', icon: 'bullets', tip: '项目符号', enabled: canEdit, checked: () => { const i = E.sel && LS.info(E.sel.f.p); return !!i && i.fmt === 'bullet'; }, run: () => LS.toggle('bullet') });
  C('incIndent', { label: '增加缩进量', icon: 'incIndent', key: 'Ctrl+M', enabled: canEdit, run: () => indent(1) });
  C('decIndent', { label: '减少缩进量', icon: 'decIndent', key: 'Ctrl+Shift+M', enabled: canEdit, run: () => indent(-1) });
  C('hangIndent', { label: '悬挂缩进', key: 'Ctrl+T', enabled: canEdit, run: () => hanging(1) });
  C('unhangIndent', { label: '减少悬挂缩进', key: 'Ctrl+Shift+T', enabled: canEdit, run: () => hanging(-1) });
  C('restartNumbering', { label: '重新开始编号(&R)', enabled: () => !!(E.sel && LS.info(E.sel.f.p)), run: () => LS.restartAt(1) });
  C('continueNumbering', { label: '继续编号(&C)', enabled: () => !!(E.sel && LS.info(E.sel.f.p)), run: () => LS.continuePrev() });
  /* styles */
  C('styleNormal', { label: '普通', key: 'Ctrl+Shift+N', enabled: canEdit, run: () => applyStyle('Normal') });
  C('styleH1', { label: '标题 1', key: 'Ctrl+Alt+1', enabled: canEdit, run: () => applyStyle('Heading1') });
  C('styleH2', { label: '标题 2', key: 'Ctrl+Alt+2', enabled: canEdit, run: () => applyStyle('Heading2') });
  C('styleH3', { label: '标题 3', key: 'Ctrl+Alt+3', enabled: canEdit, run: () => applyStyle('Heading3') });
  C('styleListBullet', { label: '项目符号列表', key: 'Ctrl+Shift+L', enabled: canEdit, run: () => applyStyle('ListBullet') });

  /* ================= Insert shortcuts ================= */
  C('insertDateField', { label: '日期域', key: 'Alt+Shift+D', enabled: canEdit, run: () => insertFieldNow('DATE \\@ "M/d/yyyy"', '插入日期') });
  C('insertTimeField', { label: '时间域', key: 'Alt+Shift+T', enabled: canEdit, run: () => insertFieldNow('TIME \\@ "h:mm am/pm"', '插入时间') });
  C('insertPageField', { label: '页码域', key: 'Alt+Shift+P', enabled: canEdit, run: () => insertFieldNow('PAGE', '插入页码') });
  C('symCopyright', { label: '©', enabled: canEdit, run: typeChar('©') });
  C('symRegistered', { label: '®', enabled: canEdit, run: typeChar('®') });
  C('symTrademark', { label: '™', enabled: canEdit, run: typeChar('™') });
  C('symEllipsis', { label: '…', enabled: canEdit, run: typeChar('…') });
  C('optHyphen', { label: '可选连字符', enabled: canEdit, run: typeChar('­') });
  C('nbHyphen', { label: '不间断连字符', enabled: canEdit, run: typeChar('‑') });
  C('emDash', { label: '长划线', enabled: canEdit, run: typeChar('—') });
  C('enDash', { label: '短划线', enabled: canEdit, run: typeChar('–') });

  /* ================= fields ================= */
  C('updateField', { label: '更新域(&U)', key: 'F9', enabled: canEdit, run: () => L.fields && L.fields.updateSelection() });
  C('editField', { label: '编辑域(&E)...', enabled: () => !!(L.fields && L.fields.at(E.sel && E.sel.f)), run: () => L.dlg.field(L.fields.at(E.sel.f)) });
  C('toggleCodes', { label: '切换域代码(&T)', key: 'Shift+F9', checked: () => !!LY.opts.fieldCodes, run: () => { LY.opts.fieldCodes = !LY.opts.fieldCodes; LY.render(); E.restoreDom(true); } });
  C('insertFieldBraces', { label: '插入域', key: 'Ctrl+F9', enabled: canEdit, run: () => L.dlg.field(null, { codes: true }) });
  C('unlinkField', { label: '断开域链接', key: 'Ctrl+Shift+F9', enabled: canEdit, run: () => L.fields && L.fields.unlinkSelection() });
  C('lockField', { label: '锁定域', key: 'Ctrl+F11', enabled: canEdit, run: () => L.fields && L.fields.lockSelection(true) });
  C('unlockField', { label: '解除域锁定', key: 'Ctrl+Shift+F11', enabled: canEdit, run: () => L.fields && L.fields.lockSelection(false) });

  /* ================= Tools ================= */
  C('research', { label: '信息检索(&R)...', icon: 'research', tip: '信息检索', run: () => { if (L.panes) { L.panes.researchQuery = E.sel && !E.collapsed() ? O.textRange(...E.range(), ' ').trim().slice(0, 80) : (E.sel && O.wordAt(E.sel.f) ? O.textRange(...O.wordAt(E.sel.f)) : ''); } pane('research')(); } });
  C('setLanguage', { label: '设置语言(&S)...', run: dlg('language') });
  C('hyphenation', { label: '断字(&H)...', run: dlg('hyphenation') });
  C('wordCount', { label: '字数统计(&W)...', icon: 'wordCount', run: dlg('wordCount') });
  C('autoSummarize', { label: '自动摘要(&u)...', icon: 'summary', enabled: canEdit, run: dlg('autoSummarize') });
  C('trackChanges', { label: '修订(&T)', icon: 'trackChanges', key: 'Ctrl+Shift+E', tip: '修订', enabled: () => !!D.doc && !E.readOnly, checked: () => !!(D.doc && doc().settings.track), run: () => { if (L.review) L.review.toggleTracking(); } });
  C('compareDocs', { label: '比较并合并文档(&D)...', enabled: canEdit, run: () => L.review && L.review.compareDialog() });
  C('protectDoc', { label: '保护文档(&P)...', icon: 'permission', run: pane('protect') });
  C('mailMerge', { label: '邮件合并(&M)...', icon: 'mailMerge', run: pane('mailmerge') });
  C('envelopes', { label: '信封和标签(&E)...', icon: 'envelope', run: dlg('envelopes') });
  C('letterWizard', { label: '信函向导(&L)...', run: dlg('letterWizard') });
  C('macros', { label: '宏(&M)...', key: 'Alt+F8', run: dlg('macros') });
  C('templates', { label: '模板和加载项(&I)...', run: dlg('templates') });
  C('autocorrectDlg', { label: '自动更正选项(&A)...', run: dlg('autocorrect') });
  C('customize', { label: '自定义(&C)...', run: dlg('customize') });
  C('optionsDlg', { label: '选项(&O)...', run: dlg('options') });

  /* ================= Table ================= */
  C('drawTable', { label: '绘制表格(&w)', icon: 'pen', tip: '绘制表格', enabled: canEdit, checked: () => !!(L.tables && L.tables.tool === 'draw'), run: () => L.tables.setTool(L.tables.tool === 'draw' ? null : 'draw') });
  C('eraser', { label: '橡皮擦', icon: 'eraser', tip: '橡皮擦', enabled: canEdit, checked: () => !!(L.tables && L.tables.tool === 'erase'), run: () => L.tables.setTool(L.tables.tool === 'erase' ? null : 'erase') });
  C('insertTable', { label: '表格(&T)...', icon: 'table', enabled: canEdit, run: dlg('insertTable') });
  C('colLeft', { label: '左侧列(&L)', icon: 'colLeft', enabled: () => canEdit() && inTable(), run: tbl('insertCols', true) });
  C('colRight', { label: '右侧列(&R)', icon: 'colRight', enabled: () => canEdit() && inTable(), run: tbl('insertCols', false) });
  C('rowAbove', { label: '上方行(&A)', icon: 'rowAbove', enabled: () => canEdit() && inTable(), run: tbl('insertRows', true) });
  C('rowBelow', { label: '下方行(&B)', icon: 'rowBelow', enabled: () => canEdit() && inTable(), run: tbl('insertRows', false) });
  C('insertCells', { label: '单元格(&e)...', enabled: () => canEdit() && inTable(), run: dlg('insertCells') });
  C('deleteTable', { label: '表格(&T)', icon: 'deleteTable', enabled: () => canEdit() && inTable(), run: tbl('deleteTable') });
  C('deleteCols', { label: '列(&C)', icon: 'deleteCol', enabled: () => canEdit() && inTable(), run: tbl('deleteCols') });
  C('deleteRows', { label: '行(&R)', icon: 'deleteRow', enabled: () => canEdit() && inTable(), run: tbl('deleteRows') });
  C('deleteCells', { label: '单元格(&e)...', enabled: () => canEdit() && inTable(), run: dlg('deleteCells') });
  C('selectTable', { label: '表格(&T)', enabled: inTable, run: tbl('select', 'table') });
  C('selectCol', { label: '列(&C)', enabled: inTable, run: tbl('select', 'col') });
  C('selectRow', { label: '行(&R)', enabled: inTable, run: tbl('select', 'row') });
  C('selectCell', { label: '单元格(&e)', enabled: inTable, run: tbl('select', 'cell') });
  C('mergeCells', { label: '合并单元格(&g)', icon: 'mergeCells', tip: '合并单元格', enabled: () => canEdit() && inTable() && !!E.cellSel, run: tbl('merge') });
  C('splitCells', { label: '拆分单元格(&p)...', icon: 'splitCells', tip: '拆分单元格', enabled: () => canEdit() && inTable(), run: dlg('splitCells') });
  C('splitTable', { label: '拆分表格(&t)', enabled: () => canEdit() && inTable(), run: tbl('splitTable') });
  C('tableAutoFormat', { label: '表格自动套用格式(&F)...', icon: 'tableAutoFormat', tip: '表格自动套用格式', enabled: () => canEdit() && inTable(), run: dlg('tableAutoFormat') });
  C('autofitContents', { label: '根据内容调整表格(&C)', enabled: () => canEdit() && inTable(), run: tbl('autofit', 'contents') });
  C('autofitWindow', { label: '根据窗口调整表格(&W)', enabled: () => canEdit() && inTable(), run: tbl('autofit', 'window') });
  C('fixedWidth', { label: '固定列宽(&F)', enabled: () => canEdit() && inTable(), run: tbl('autofit', 'fixed') });
  C('distRows', { label: '平均分布各行(&y)', icon: 'distRows', tip: '平均分布各行', enabled: () => canEdit() && inTable(), run: tbl('distribute', 'rows') });
  C('distCols', { label: '平均分布各列(&y)', icon: 'distCols', tip: '平均分布各列', enabled: () => canEdit() && inTable(), run: tbl('distribute', 'cols') });
  C('headingRows', { label: '标题行重复(&H)', enabled: () => canEdit() && inTable(), checked: () => inTable() && L.tables.isHeaderRow(), run: tbl('toggleHeaderRow') });
  C('textToTable', { label: '文本转换成表格(&x)...', enabled: () => canEdit() && hasSel() && !inTable(), run: dlg('textToTable') });
  C('tableToText', { label: '表格转换成文本(&b)...', enabled: () => canEdit() && inTable(), run: dlg('tableToText') });
  C('sortDlg', { label: '排序(&S)...', icon: 'sortAsc', enabled: canEdit, run: dlg('sort') });
  C('sortAsc', { label: '升序排序', icon: 'sortAsc', enabled: canEdit, run: () => L.tables.quickSort(false) });
  C('sortDesc', { label: '降序排序', icon: 'sortDesc', enabled: canEdit, run: () => L.tables.quickSort(true) });
  C('autoSum', { label: '自动求和', icon: 'autoSum', enabled: () => canEdit() && inTable(), run: () => L.tables.autoSum() });
  C('formulaDlg', { label: '公式(&o)...', enabled: canEdit, run: dlg('formula') });
  C('tableProps', { label: '表格属性(&r)...', icon: 'tableProps', enabled: () => inTable(), run: dlg('tableProps') });
  C('cellTop', { label: '顶端对齐', icon: 'alignTop', enabled: () => canEdit() && inTable(), checked: () => inTable() && L.tables.cellVAlign() === 'top', run: tbl('setVAlign', 'top') });
  C('cellMiddle', { label: '水平居中', icon: 'alignMiddle', enabled: () => canEdit() && inTable(), checked: () => inTable() && L.tables.cellVAlign() === 'center', run: tbl('setVAlign', 'center') });
  C('cellBottom', { label: '底端对齐', icon: 'alignBottom', enabled: () => canEdit() && inTable(), checked: () => inTable() && L.tables.cellVAlign() === 'bottom', run: tbl('setVAlign', 'bottom') });

  /* ================= Window & Help ================= */
  C('newWindow', { label: '新建窗口(&N)', enabled: () => false, run: () => {} });
  C('arrangeAll', { label: '全部重排(&A)', enabled: () => false, run: () => {} });
  C('splitWindow', { label: '拆分(&S)', checked: () => !!A.splitOn, run: () => A.toggleSplit() });
  C('help', { label: 'Quire 帮助(&H)', menuLabel: 'Quire 帮助(&H)', icon: 'help', key: 'F1', tip: '帮助', run: pane('help') });
  C('gettingStarted', { label: '入门(&G)', run: pane('getting-started') });
  C('keyboardHelp', { label: '键盘快捷键(&K)', run: () => { if (L.panes) { L.panes.helpTopic = 'keys'; pane('help')(); } } });
  C('about', { label: '关于 Quire(&A)', run: dlg('about') });

  /* ================= Reviewing ================= */
  const MK = { markupFinalMarkup: ['显示标记的最终状态', 'markup'], markupFinal: [L.N('最终状态'), 'final'], markupOriginalMarkup: ['显示标记的原始状态', 'origMarkup'], markupOriginal: [L.N('原始状态'), 'original'] };
  for (const k in MK) C(k, { label: MK[k][0], radio: true, checked: () => LY.opts.markup === MK[k][1], run: () => { LY.opts.markup = MK[k][1]; LY.render(); E.restoreDom(true); } });
  C('showComments', { label: '批注(&C)', checked: () => LY.opts.showComments !== false, run: () => { LY.opts.showComments = !LY.opts.showComments; LY.render(); E.restoreDom(true); } });
  C('showInsDel', { label: '插入和删除(&I)', checked: () => LY.opts.markup !== 'final', run: () => { LY.opts.markup = LY.opts.markup === 'final' ? 'markup' : 'final'; LY.render(); E.restoreDom(true); } });
  C('showFormatting', { label: '格式设置(&F)', checked: () => !!LY.opts.showFmtChanges, run: () => { LY.opts.showFmtChanges = !LY.opts.showFmtChanges; LY.render(); E.restoreDom(true); } });
  C('reviewersMenu', { label: '审阅者(&R)', run: () => L.review && L.review.reviewersDialog() });
  C('balloonsToggle', { label: '批注框(&B)', checked: () => LY.opts.balloons !== false, run: () => { LY.opts.balloons = !LY.opts.balloons; LY.render(); E.restoreDom(true); } });
  C('reviewingPane', { label: '审阅窗格(&P)', icon: 'reviewingPane', tip: '审阅窗格', checked: () => !L.$('#revpane').hidden, run: () => L.review && L.review.togglePane() });
  C('prevChange', { label: '上一个', icon: 'prevChange', tip: '上一个', run: () => L.review && L.review.step(-1) });
  C('nextChange', { label: '下一步', icon: 'nextChange', tip: '下一步', run: () => L.review && L.review.step(1) });
  C('acceptChange', { label: '接受修订(&A)', icon: 'accept', tip: '接受修订', enabled: canEdit, run: () => L.review && L.review.acceptReject(true, 'selection') });
  C('acceptShown', { label: '接受所有显示的修订(&S)', enabled: canEdit, run: () => L.review && L.review.acceptReject(true, 'all') });
  C('acceptAll', { label: '接受对文档的所有修订(&D)', enabled: canEdit, run: () => L.review && L.review.acceptReject(true, 'all') });
  C('rejectChange', { label: '拒绝修订/删除批注(&R)', icon: 'reject', tip: '拒绝修订/删除批注', enabled: canEdit, run: () => L.review && L.review.acceptReject(false, 'selection') });
  C('rejectShown', { label: '拒绝所有显示的修订', enabled: canEdit, run: () => L.review && L.review.acceptReject(false, 'all') });
  C('rejectAll', { label: '拒绝对文档的所有修订', enabled: canEdit, run: () => L.review && L.review.acceptReject(false, 'all') });
  C('deleteAllComments', { label: '删除所有显示的批注', enabled: () => canEdit() && Object.keys(doc().comments).length > 0, run: () => L.review && L.review.deleteAllComments() });

  /* ================= Picture & drawing ================= */
  const imgSel = () => canEdit() && selectedObjIs((it) => it.t === 'img');
  C('picMoreContrast', { label: '增加对比度', icon: 'moreContrast', enabled: imgSel, run: () => A.picAdjust('contrast', 0.1) });
  C('picLessContrast', { label: '降低对比度', icon: 'lessContrast', enabled: imgSel, run: () => A.picAdjust('contrast', -0.1) });
  C('picMoreBright', { label: '增加亮度', icon: 'moreBright', enabled: imgSel, run: () => A.picAdjust('bright', 0.1) });
  C('picLessBright', { label: '降低亮度', icon: 'lessBright', enabled: imgSel, run: () => A.picAdjust('bright', -0.1) });
  C('picCrop', { label: '裁剪', icon: 'crop', enabled: imgSel, run: () => L.dlg.formatObject('picture') });
  C('picRotate', { label: '向左旋转 90°', icon: 'rotateL', enabled: () => canEdit() && objSel(), run: () => L.drawing.modifySelected(L.N('旋转'), (it) => { it.rot = (((it.rot || 0) - 90) % 360 + 360) % 360 || undefined; }) });
  C('picReset', { label: '重置图片', icon: 'resetPicture', enabled: imgSel, run: () => L.drawing.modifySelected('重置图片', (it) => { delete it.bright; delete it.contrast; delete it.gray; delete it.bw; delete it.washout; delete it.crop; delete it.rot; if (it.natW) { it.w = it.natW; it.h = it.natH; } }) });
  C('selectObjects', { label: '选择对象', icon: 'select', checked: () => !!A.selectObjectsMode, run: () => { A.selectObjectsMode = !A.selectObjectsMode; L.$('#scroller').classList.toggle('selobj', A.selectObjectsMode); } });
  C('drawLine', { label: '线条(&L)', icon: 'line', tip: '线条', enabled: canEdit, run: () => L.drawtool && L.drawtool.start({ geom: 'line', line: true }) });
  C('drawArrow', { label: '箭头(&A)', icon: 'arrowTool', tip: '箭头', enabled: canEdit, run: () => L.drawtool && L.drawtool.start({ geom: 'line', line: true, arrow: true }) });
  C('drawDblArrow', { label: '双箭头(&D)', icon: 'doubleArrow', enabled: canEdit, run: () => L.drawtool && L.drawtool.start({ geom: 'line', line: true, dblArrow: true }) });
  C('drawRect', { label: '矩形(&R)', icon: 'rect', tip: '矩形', enabled: canEdit, run: () => L.drawtool && L.drawtool.start({ geom: 'rect' }) });
  C('drawOval', { label: '椭圆(&O)', icon: 'oval', tip: '椭圆', enabled: canEdit, run: () => L.drawtool && L.drawtool.start({ geom: 'ellipse' }) });
  C('fillColor', { label: '填充颜色', icon: 'fillColor', tip: '填充颜色', enabled: () => canEdit() && objSel(), run: () => L.drawing.modifySelected('Fill Color', (it) => { it.fill = { t: 'solid', c: A.lastFill || '#CCFFCC', a: 1 }; }) });
  C('lineColor', { label: '线条颜色', icon: 'lineColor', tip: '线条颜色', enabled: () => canEdit() && objSel(), run: () => L.drawing.modifySelected('线条颜色', (it) => { it.line = Object.assign({ w: 0.75, dash: 'solid' }, it.line && it.line.t !== 'none' ? it.line : {}, { c: A.lastLine || '#000000' }); }) });
  const order = (label, fn) => ({ label, enabled: () => canEdit() && floatSel(), run: () => L.drawing.modifySelected(ui.plainAmp(label), fn) });
  C('bringFront', Object.assign(order('置于顶层(&n)', (it) => { it.float.z = A.maxZ() + 1024; it.float.behind = false; if (it.float.wrap === 'behind') it.float.wrap = 'none'; }), { icon: 'bringFront' }));
  C('sendBack', Object.assign(order('置于底层(&k)', (it) => { it.float.z = Math.max(0, A.minZ() - 1024); }), { icon: 'sendBack' }));
  C('bringForward', Object.assign(order('上移一层(&F)', (it) => { it.float.z = (it.float.z || 0) + 1024 * 1024; }), { icon: 'bringForward' }));
  C('sendBackward', Object.assign(order('下移一层(&w)', (it) => { it.float.z = Math.max(0, (it.float.z || 0) - 1024 * 1024); }), { icon: 'sendBackward' }));
  C('bringFrontText', order('浮于文字上方(&T)', (it) => { it.float.behind = false; if (it.float.wrap === 'behind') it.float.wrap = 'none'; }));
  C('sendBehindText', order('衬于文字下方(&e)', (it) => { it.float.behind = true; it.float.wrap = 'behind'; }));
  C('groupObj', { label: '组合(&G)', icon: 'group', enabled: () => false, run: () => {} });
  C('ungroupObj', { label: '取消组合(&U)', icon: 'ungroup', enabled: () => canEdit() && selectedObjIs((it) => it.t === 'group'), run: () => A.ungroup() });
  for (const [k, dir, lbl] of [['nudgeUp', 'ArrowUp', '向上(&U)'], ['nudgeDown', 'ArrowDown', '向下(&D)'], ['nudgeLeft', 'ArrowLeft', '左(&L)'], ['nudgeRight', 'ArrowRight', '右(&R)']]) C(k, { label: lbl, enabled: () => canEdit() && floatSel(), run: () => L.drawing.nudge(dir, 1) });
  C('rotateLeft', { label: '向左旋转 90°(&L)', icon: 'rotateL', enabled: () => canEdit() && objSel(), run: () => L.drawing.modifySelected(L.N('旋转'), (it) => { it.rot = (((it.rot || 0) - 90) % 360 + 360) % 360 || undefined; }) });
  C('rotateRight', { label: '向右旋转 90°(&R)', icon: 'rotateR', enabled: () => canEdit() && objSel(), run: () => L.drawing.modifySelected(L.N('旋转'), (it) => { it.rot = ((it.rot || 0) + 90) % 360 || undefined; }) });
  C('flipH', { label: '水平翻转(&H)', icon: 'flipH', enabled: () => canEdit() && selectedObjIs((it) => it.t === 'shape' || it.t === 'img'), run: () => L.drawing.modifySelected(L.N('翻转'), (it) => { it.flipH = !it.flipH || undefined; }) });
  C('flipV', { label: '垂直翻转(&V)', icon: 'flipV', enabled: () => canEdit() && selectedObjIs((it) => it.t === 'shape' || it.t === 'img'), run: () => L.drawing.modifySelected(L.N('翻转'), (it) => { it.flipV = !it.flipV || undefined; }) });
  C('setDefaults', { label: '设置自选图形默认值(&D)', enabled: () => selectedObjIs((it) => it.t === 'shape'), run: () => { const it = L.drawing.selectedItem(); A.shapeDefaults = { fill: L.clone(it.fill), line: L.clone(it.line), shadow: it.shadow ? L.clone(it.shadow) : undefined }; A.status('新的自选图形将使用这些设置。'); } });
  C('wordartEdit', { label: '编辑文字(&x)...', enabled: () => canEdit() && selectedObjIs((it) => !!it.wordart), run: () => L.dlg.wordart(true) });
  C('wordartGallery', { label: 'LettersArt 库', icon: 'waGallery', enabled: () => canEdit() && selectedObjIs((it) => !!it.wordart), run: () => L.dlg.wordart('gallery') });

  /* ================= Outlining ================= */
  const ol = (fn, ...a) => () => L.outline && L.outline[fn](...a);
  C('promoteH1', { label: '提升为标题 1', icon: 'promoteH1', enabled: canEdit, run: ol('setLevel', 1) });
  C('promote', { label: '升级', icon: 'promote', key: 'Alt+Shift+Left', enabled: canEdit, run: ol('shift', -1) });
  C('demote', { label: '降级', icon: 'demote', key: 'Alt+Shift+Right', enabled: canEdit, run: ol('shift', 1) });
  C('demoteBody', { label: '降为正文', icon: 'demoteBody', enabled: canEdit, run: ol('setLevel', 0) });
  C('moveUp', { label: '上移', icon: 'moveUp', key: 'Alt+Shift+Up', enabled: canEdit, run: ol('move', -1) });
  C('moveDown', { label: '下移', icon: 'moveDown', key: 'Alt+Shift+Down', enabled: canEdit, run: ol('move', 1) });
  C('expand', { label: '展开', icon: 'expand', enabled: () => A.view === 'outline', run: ol('expand', true) });
  C('collapse', { label: '折叠', icon: 'collapse', enabled: () => A.view === 'outline', run: ol('expand', false) });
  C('olFirstLine', { label: '仅显示首行', icon: 'firstLine', enabled: () => A.view === 'outline', checked: () => !!(L.outline && L.outline.firstLine), run: ol('toggle', 'firstLine') });
  C('olFormatting', { label: '显示文字格式', icon: 'showFormatting', enabled: () => A.view === 'outline', checked: () => !(L.outline && L.outline.plain), run: ol('toggle', 'plain') });
  C('updateTOC', { label: '更新目录', icon: 'updateTOC', tip: '更新目录', enabled: canEdit, run: () => L.fields && L.fields.updateAllOfType('TOC') });
  C('gotoTOC', { label: '转到目录', icon: 'gotoTOC', tip: '转到目录', run: () => L.fields && L.fields.gotoFirst('TOC') });

  /* ================= Header and Footer ================= */
  const hfOnly = () => !!LY.hfEdit;
  C('hfPageNum', { label: '插入页码', icon: 'pageNum', enabled: hfOnly, run: () => insertFieldNow('PAGE', '插入页码') });
  C('hfNumPages', { label: '插入总页数', icon: 'numPages', enabled: hfOnly, run: () => insertFieldNow('NUMPAGES', '插入总页数') });
  C('hfFormatPageNum', { label: '设置页码格式', icon: 'formatPageNum', enabled: () => !!D.doc, run: () => L.dlg.pageNumbers('format') });
  C('hfDate', { label: '插入日期', icon: 'date', enabled: hfOnly, run: () => insertFieldNow('DATE \\@ "M/d/yyyy"', '插入日期') });
  C('hfTime', { label: '插入时间', icon: 'time', enabled: hfOnly, run: () => insertFieldNow('TIME \\@ "h:mm:ss am/pm"', '插入时间') });
  C('hfShowText', { label: '显示/隐藏文档文字', icon: 'showDocText', enabled: hfOnly, checked: () => LY.root && LY.root.classList.contains('hidebody'), run: () => LY.root.classList.toggle('hidebody') });
  C('hfLinkPrev', { label: '链接到前一节', icon: 'linkPrev', enabled: () => hfOnly() && A.hfSectionIndex() > 0, checked: () => hfOnly() && A.hfLinked(), run: () => A.hfToggleLink() });
  C('hfSwitch', { label: '在页眉和页脚间切换', icon: 'hfSwitch', enabled: hfOnly, run: () => A.hfSwitch() });
  C('hfPrev', { label: '显示上一项', icon: 'showPrev', enabled: hfOnly, run: () => A.hfStep(-1) });
  C('hfNext', { label: '显示下一项', icon: 'showNext', enabled: hfOnly, run: () => A.hfStep(1) });
  C('hfClose', { label: '关闭页眉和页脚', enabled: hfOnly, run: () => A.closeHeaderFooter() });

  /* ================= Mail Merge toolbar ================= */
  const mm = (fn, ...a) => () => L.mailmerge && L.mailmerge[fn](...a);
  const hasData = () => !!(L.mailmerge && L.mailmerge.data);
  C('mmMainDoc', { label: '主文档设置', icon: 'mmMain', run: pane('mailmerge') });
  C('mmOpenSource', { label: '打开数据源', icon: 'open', run: mm('openSource') });
  C('mmRecipients', { label: '邮件合并收件人', icon: 'mmRecipients', enabled: hasData, run: mm('recipientsDialog') });
  C('mmAddress', { label: '插入地址块', icon: 'mmAddress', enabled: () => canEdit() && hasData(), run: mm('insertBlock', 'address') });
  C('mmGreeting', { label: '插入问候语', icon: 'mmGreeting', enabled: () => canEdit() && hasData(), run: mm('insertBlock', 'greeting') });
  C('mmInsertField', { label: '插入合并域', icon: 'mmField', enabled: () => canEdit() && hasData(), run: mm('fieldDialog') });
  C('mmViewData', { label: '查看合并数据', icon: 'mmView', enabled: hasData, checked: () => !!(L.mailmerge && L.mailmerge.preview), run: mm('togglePreview') });
  C('mmFirst', { label: '第一条记录', icon: 'first', enabled: hasData, run: mm('go', 0) });
  C('mmPrev', { label: '上一条记录', icon: 'showPrev', enabled: hasData, run: mm('step', -1) });
  C('mmNext', { label: '下一条记录', icon: 'showNext', enabled: hasData, run: mm('step', 1) });
  C('mmLast', { label: '最后一条记录', icon: 'last', enabled: hasData, run: mm('go', -1) });
  C('mmMergeNew', { label: '合并到新文档', icon: 'mmMergeNew', enabled: hasData, run: mm('mergeToNew') });
  C('mmMergePrint', { label: '合并到打印机', icon: 'mmMergePrint', enabled: hasData, run: mm('mergeToPDF') });

  /* ================= navigation ================= */
  C('browsePrev', { label: '上一个', icon: 'browseUp', tip: '上一页', run: () => A.browse(-1) });
  C('browseNext', { label: '下一步', icon: 'browseDown', tip: '下一页', run: () => A.browse(1) });
  C('extendSel', { label: '扩展选区', key: 'F8', run: () => A.extendStep() });
  C('goBack', { label: '返回', key: 'Shift+F5', run: () => A.goBack() });

  /* ================= helpers used by commands ================= */
  A.maxZ = () => { let m = 251659264; D.walk(doc().main, (b) => { if (b.t === 'p') for (const it of b.runs) if (it.float && it.float.z > m) m = it.float.z; }); return m; };
  A.minZ = () => { let m = 251659264; D.walk(doc().main, (b) => { if (b.t === 'p') for (const it of b.runs) if (it.float && it.float.z < m) m = it.float.z; }); return m; };
  A.toggleFullScreen = function () {
    const app = L.$('#app');
    const on = !app.classList.contains('fullscreen');
    app.classList.toggle('fullscreen', on);
    let btn = L.$('.closefs');
    if (on && !btn) { btn = L.h('button', { class: 'btn closefs', type: 'button', text: '关闭全屏显示', onclick: () => A.toggleFullScreen() }); document.body.appendChild(btn); }
    if (!on && btn) btn.remove();
    LY.fitSizer();
    if (L.rulers) L.rulers.draw();
  };
  A.cycleCase = function () {
    if (!E.sel) return;
    let [a, b] = E.range();
    if (D.eqPos(a, b)) { const w = O.wordAt(a); if (!w) return; [a, b] = w; }
    const t = O.textRange(a, b, '\n');
    const mode = t === t.toUpperCase() && t !== t.toLowerCase() ? 'lower' : t === t.toLowerCase() ? 'title' : 'upper';
    E.edit('更改大小写', () => { O.changeCase(a, b, mode); return { a, f: b }; });
  };
  A.removeHyperlink = function () {
    if (!E.sel) return;
    let [a, b] = E.range();
    if (D.eqPos(a, b)) {
      /* the whole link around the caret */
      const p = a.p;
      let pos = 0, s = -1, e = -1;
      for (const it of p.runs) { const n = D.ilen(it); if (it.rPr && it.rPr.link && pos <= a.o && a.o <= pos + n) { if (s < 0) s = pos; e = pos + n; } else if (s >= 0 && pos > a.o) break; pos += n; }
      if (s < 0) return;
      a = D.pos(p, s); b = D.pos(p, e);
    }
    E.edit('删除超链接', () => { O.formatRange(a, b, (r) => { delete r.link; if (r.style === 'Hyperlink' || r.style === 'FollowedHyperlink') delete r.style; return r; }); return E.sel; });
  };
  A.formatObject = function (editContent) {
    const it = L.drawing && L.drawing.selectedItem();
    if (!it) return;
    if (editContent && it.t === 'chart' && L.chart && L.chart.edit) { A.editChart(); return; }
    if (editContent && it.wordart) { L.dlg.wordart(true); return; }
    if (L.dlg && L.dlg.formatObject) L.dlg.formatObject();
  };
  A.insertChart = function () {
    const c = L.chart.sample ? L.chart.sample('bar') : null;
    if (!c) return;
    const it = { t: 'chart', w: 360, h: 216, chart: c };
    E.edit('Insert Chart', () => O.insertItem(E.deleteSelection(), it));
    setTimeout(() => { const os = A.selectItemAt(E.sel.f.p, E.sel.f.o - 1); if (os) A.editChart(); }, 30);
  };
  A.selectItemAt = function (p, o) {
    const f = LY.fragFor(p, o);
    const el = f && f.querySelector(`[data-o="${o}"]`);
    if (el) { E.selectInlineObject(D.pos(p, o), el); return E.objSel; }
    return null;
  };
  A.editChart = function () {
    const os = E.objSel;
    const it = L.drawing.selectedItem();
    if (!os || !it || it.t !== 'chart') return;
    const shape = { chart: L.clone(it.chart), w: it.w, h: it.h };
    L.chart.edit(shape, (res) => { if (!res) return; L.drawing.modifySelected(L.N('图表'), (x) => { x.chart = res.chart || shape.chart; }); });
  };
  A.ungroup = function () {
    const os = E.objSel;
    const it = L.drawing.selectedItem();
    if (!os || !it || it.t !== 'group') return;
    E.edit(L.N('取消组合'), () => {
      D.touch(os.p);
      const i = os.p.runs.indexOf(it);
      const baseH = it.float ? it.float.posH : null, baseV = it.float ? it.float.posV : null;
      const kids = (it.kids || []).map((k) => {
        const c = L.clone(k);
        c.float = L.clone(it.float || { posH: { rel: 'column', off: 0 }, posV: { rel: 'paragraph', off: 0 }, wrap: 'none', dist: { t: 0, b: 0, l: 9, r: 9 }, z: 251659264 });
        c.float.posH = { rel: baseH ? baseH.rel : 'column', off: L.round(((baseH && baseH.off) || 0) + (k.x || 0), 2) };
        c.float.posV = { rel: baseV ? baseV.rel : 'paragraph', off: L.round(((baseV && baseV.off) || 0) + (k.y || 0), 2) };
        delete c.x; delete c.y;
        return c;
      });
      os.p.runs.splice(i, 1, ...kids);
      return D.pos(os.p, os.o);
    });
    E.clearObjSel();
  };
  A.insertFileDialog = async function () {
    const files = await L.pickFiles('.docx,.docm,.dotx,.txt,.htm,.html');
    const f = files[0];
    if (!f) return;
    try {
      let blocks;
      if (/\.txt$/i.test(f.name)) blocks = A.docFromText(await L.readAsText(f)).main.blocks;
      else if (/\.html?$/i.test(f.name)) blocks = await L.htmlio.htmlToBlocks(await L.readAsText(f));
      else {
        const res = await L.docx.read(await L.readAsArrayBuffer(f));
        const src = res.doc;
        A.mergeStylesFrom(src);
        blocks = src.main.blocks;
        /* carry notes and comments along */
        const d = doc();
        E.edit('插入文件', () => {
          D.touchKey(d, 'fn'); D.touchKey(d, 'en'); D.touchKey(d, 'comments'); D.touchKey(d, 'numbering');
          for (const k in src.fn) if (!d.fn[k]) d.fn[k] = src.fn[k];
          for (const k in src.en) if (!d.en[k]) d.en[k] = src.en[k];
          for (const k in src.comments) if (!d.comments[k]) d.comments[k] = src.comments[k];
          A.mergeNumbering(src, blocks);
          return E.sel;
        });
      }
      E.pasteBlocks(blocks);
    } catch (e) { ui.msg(L.t('无法插入文件。{0}', e.message || ''), { icon: 'error' }); }
  };
  /** copy styles missing from the current document */
  A.mergeStylesFrom = function (src, overwrite) {
    const d = doc();
    E.edit('复制样式', () => {
      D.touchKey(d, 'styles');
      for (const k in src.styles) if (overwrite || !d.styles[k]) d.styles[k] = L.clone(src.styles[k]);
      D.stylesChanged();
      return E.sel;
    });
  };
  /** remap list ids of blocks coming from another document */
  A.mergeNumbering = function (src, blocks) {
    const d = doc();
    const numbering = L.opc.duplicate(src.numbering);
    const map = {};
    const absMap = {};
    for (const nid in numbering.nums) {
      const n = numbering.nums[nid];
      if (absMap[n.abs] == null) { const abs = numbering.abs[n.abs]; let aid = 0; while (d.numbering.abs[aid] != null) aid++; abs.id = aid; d.numbering.abs[aid] = abs; absMap[n.abs] = aid; }
      let k = 1; while (d.numbering.nums[k] != null) k++;
      d.numbering.nums[k] = { ...n, abs: absMap[n.abs] };
      map[nid] = String(k);
    }
    const fix = (bl) => D.walk({ blocks: bl }, (b) => { if (b.t === 'p' && b.pPr.num && map[b.pPr.num.id]) b.pPr.num = Object.assign({}, b.pPr.num, { id: map[b.pPr.num.id] }); });
    fix(blocks);
  };

  /* ---------- split window ---------- */
  A.toggleSplit = function () {
    A.splitOn = !A.splitOn;
    let pane2 = L.$('#split2');
    if (A.splitOn) {
      if (!pane2) {
        pane2 = L.h('div', { id: 'split2', class: 'split2', 'aria-label': '第二窗格' });
        const bar = L.h('div', { class: 'splitter h', id: 'split2-bar', role: 'separator' });
        L.$('#center').insertBefore(bar, L.$('#viewbar'));
        L.$('#center').insertBefore(pane2, bar);
        bar.addEventListener('pointerdown', (e) => {
          if (e.button !== 0) return;
          e.preventDefault();
          const h0 = pane2.offsetHeight, style0 = pane2.style.height;
          L.ui.drag(e, {
            move: (ev, d) => { pane2.style.height = L.clamp(h0 - d.dy, 60, window.innerHeight * 0.7) + 'px'; },
            cancel: () => { pane2.style.height = style0; },
          });
        });
        pane2.addEventListener('click', (e) => {
          const pg = e.target.closest('.pg');
          if (!pg) return;
          const i = +pg.dataset.pg;
          const real = LY.pages[i];
          if (real) real.el.scrollIntoView({ block: 'start' });
        });
      }
      L.$('#split2-bar').hidden = false; pane2.hidden = false;
      A.refreshSplit();
    } else if (pane2) { pane2.hidden = true; L.$('#split2-bar').hidden = true; }
  };
  A.refreshSplit = L.debounce(function () {
    const pane2 = L.$('#split2');
    if (!pane2 || pane2.hidden) return;
    const top = pane2.scrollTop;
    pane2.textContent = '';
    const inner = LY.root.cloneNode(true);
    inner.removeAttribute('id');
    inner.contentEditable = 'false';
    inner.style.position = 'relative'; inner.style.left = '0'; inner.style.top = '8px';
    inner.style.transform = `scale(${LY.zoom})`;
    const sizer = L.h('div', { style: `height:${LY.root.offsetHeight * LY.zoom + 20}px;position:relative;display:flex;justify-content:center` }, inner);
    inner.style.transformOrigin = '50% 0';
    pane2.appendChild(sizer);
    pane2.scrollTop = top;
  }, 300);
  L.bus.on('layout-done', () => { if (A.splitOn) A.refreshSplit(); });

  /* ---------- extend selection (F8) ---------- */
  A.extendStep = function () {
    if (!E.sel) return;
    A.extLevel = E.extend ? (A.extLevel || 0) + 1 : 0;
    E.extend = true;
    if (!A.extLevel) A.extAnchor = E.sel.a;
    else if (A.extLevel === 1) { const w = O.wordExtent(E.sel.f); E.setSel(w[0], w[1]); }
    else if (A.extLevel === 2) { const s = O.sentenceExtent(E.sel.f); E.setSel(s[0], s[1]); }
    else if (A.extLevel === 3) E.selectPara(E.sel.a.p);
    else E.selectAll();
    A.updateStatus();
  };
  L.bus.on('sel', () => {
    if (!E.extend || !A.extAnchor || A.extBusy) return;
    if (A.extLevel) return;
    if (E.sel && !D.eqPos(E.sel.a, A.extAnchor) && D.storyOf(doc(), E.sel.f.p) === D.storyOf(doc(), A.extAnchor.p)) { A.extBusy = true; E.setSel(A.extAnchor, E.sel.f, { noScroll: true }); A.extBusy = false; }
  });
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && E.extend) { E.extend = false; A.extAnchor = null; A.updateStatus(); } }, true);

  /* ---------- go back (Shift+F5): last three edit locations ---------- */
  A.editTrail = [];
  L.bus.on('doc-changed', (info) => {
    if (!E.sel || (info && info.history)) return;
    const p = E.sel.f;
    const last = A.editTrail[0];
    if (last && last.p === p.p && Math.abs(last.o - p.o) < 40) { A.editTrail[0] = { p: p.p, o: p.o }; return; }
    A.editTrail.unshift({ p: p.p, o: p.o });
    A.editTrail.length = Math.min(A.editTrail.length, 4);
  });
  A.goBack = function () {
    if (A.editTrail.length < 2) { if (A.editTrail[0]) A.gotoPos(D.pos(A.editTrail[0].p, A.editTrail[0].o)); return; }
    const t = A.editTrail.shift();
    A.editTrail.push(t);
    const target = A.editTrail[0];
    if (D.info(doc(), target.p)) A.gotoPos(D.pos(target.p, Math.min(target.o, D.plen(target.p))));
  };

  /* ---------- header/footer navigation ---------- */
  A.hfSectionIndex = () => { if (!LY.hfEdit) return 0; const pg = LY.pages[LY.hfEdit.page]; return pg ? pg.si : 0; };
  A.hfLinked = function () {
    const pg = LY.pages[LY.hfEdit.page];
    if (!pg) return false;
    const secs = D.sections(doc());
    const own = secs[pg.si].sect.refs && secs[pg.si].sect.refs[LY.hfEdit.kind] && secs[pg.si].sect.refs[LY.hfEdit.kind][pg.type];
    return !own;
  };
  A.hfToggleLink = function () {
    const d = doc();
    const pg = LY.pages[LY.hfEdit.page];
    const secs = D.sections(d);
    const kind = LY.hfEdit.kind;
    const s = secs[pg.si];
    const holder = s.endPara || d;
    E.edit('链接到前一节', () => {
      D.touchKey(holder, 'sect');
      const sect = holder === d ? d.sect : holder.sect;
      sect.refs = sect.refs || { hdr: {}, ftr: {} };
      sect.refs[kind] = Object.assign({}, sect.refs[kind] || {});
      if (sect.refs[kind][pg.type]) delete sect.refs[kind][pg.type];
      else {
        /* unlink: copy the inherited story into a new one for this section */
        const inh = LY.hfFor(secs, pg.si, kind, pg.type);
        const id = (kind === 'hdr' ? 'h' : 'f') + D.nid();
        D.touchKey(d, 'hf');
        d.hf[id] = { kind, id, blocks: inh ? D.cloneBlocks(inh.blocks) : [D.para([], { style: kind === 'hdr' ? 'Header' : 'Footer' })] };
        sect.refs[kind][pg.type] = id;
      }
      return E.sel;
    });
    A.editHeaderFooter(kind, pg.idx);
  };
  A.hfSwitch = () => { const pg = LY.hfEdit.page; A.editHeaderFooter(LY.hfEdit.kind === 'hdr' ? 'ftr' : 'hdr', pg); const p = LY.pages[pg]; if (p) p.el.scrollIntoView({ block: LY.hfEdit.kind === 'hdr' ? 'start' : 'end' }); };
  A.hfStep = function (dir) {
    /* next page whose header/footer story differs */
    const cur = LY.pages[LY.hfEdit.page];
    const kind = LY.hfEdit.kind;
    const story = (pg) => (kind === 'hdr' ? pg.hdrStory : pg.ftrStory);
    for (let i = LY.hfEdit.page + dir; i >= 0 && i < LY.pages.length; i += dir) {
      const pg = LY.pages[i];
      if (pg.si !== cur.si || pg.type !== cur.type || story(pg) !== story(cur)) { A.editHeaderFooter(kind, i); pg.el.scrollIntoView({ block: 'start' }); return; }
    }
  };
})();
