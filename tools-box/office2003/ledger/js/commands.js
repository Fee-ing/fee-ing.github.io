/* Ledger — command definitions (menus, toolbars, shortcuts all run these). */
(function (root) {
  'use strict';
  const L = root.L;
  const M = L.model, LY = L.layout, O = L.ops, F = L.formula, NF = L.numfmt, C = L.calc;
  const ui = L.ui;
  const cmd = ui.cmd;
  const G = () => L.grid, A = () => L.app, D = () => L.dlg, E = () => L.editor, CL = () => L.clip;
  const MAXR = M.MAXR, MAXC = M.MAXC;
  const sh = () => G().sheet();
  const wb = () => G().wb;
  const sel = () => G().sel();
  const rgs = () => G().ranges();
  const act = () => { const s = sel(); return LY.styleOf(sh(), s.r, s.c, sh().get(s.r, s.c)); };
  const notEditing = () => !E().active;
  const guard = (key) => G().guardProtect(key, !key);
  const run = (fn) => () => { if (E().active && !E().commit()) return; try { fn(); } catch (e) { A().error(e); } G().paint(); };
  /** apply a style change to the selection */
  const fmt = (delta, label) => run(() => { if (guard('formatCells')) return; O.format(sh(), rgs(), delta, label); });
  const hasMulti = () => rgs().length > 1;
  const wholeCols = () => rgs().every((x) => G().isWholeCols(x));
  const wholeRows = () => rgs().every((x) => G().isWholeRows(x));

  /* ------------------------------------------------------------ File */
  /* the toolbar button makes a blank workbook; the menu and Ctrl+N show Blank and the templates */
  cmd('newBook', { label: '新建(&N)...', menuLabel: '新建(&N)...', tip: '新建', icon: 'new', key: 'Ctrl+N', run: (arg, ev) => { if ((ev && ev.currentTarget && ev.currentTarget.closest && ev.currentTarget.closest('.toolbar')) || !window.VO) A().newWorkbook(); else VO.newDialog('ledger', (id) => (id ? L.panes.newFromTemplate(id) : A().newWorkbook())); } });
  cmd('newBlank', { label: '新建空白工作簿', icon: 'new', run: () => A().newWorkbook() });
  cmd('open', { label: '打开(&O)...', icon: 'open', key: 'Ctrl+O', run: () => A().openDialog() });
  cmd('close', { label: '关闭(&C)', run: () => A().closeWorkbook() });
  cmd('save', { label: '保存(&S)', icon: 'save', key: 'Ctrl+S', run: run(() => A().save()) });
  cmd('saveAs', { label: '另存为(&A)...', key: 'F12', run: run(() => D().saveAs()) });
  cmd('saveWeb', { label: '另存为网页(&g)...', run: run(() => A().saveAs('html')) });
  cmd('exportPDF', { label: '导出为 PDF(&P)...', icon: 'print', run: run(() => L.print.exportPDF()) });
  cmd('exportCSV', { label: '另存为 CSV(&C)', run: run(() => A().saveAs('csv')) });
  cmd('pageSetup', { label: '页面设置(&u)...', icon: 'pageSetup', run: run(() => D().pageSetup()) });
  cmd('setPrintArea', { label: '设置打印区域(&S)', run: run(() => A().setPrintArea(true)) });
  cmd('clearPrintArea', { label: '清除打印区域(&C)', run: run(() => A().setPrintArea(false)), enabled: () => !!A().printAreaOf(sh()) });
  cmd('printPreview', { label: '打印预览(&v)', icon: 'preview', run: run(() => L.print.preview()) });
  cmd('print', { label: '打印(&P)...', icon: 'print', key: 'Ctrl+P', run: run(() => D().print()) });
  cmd('printQuick', { label: '打印', tip: '打印', icon: 'print', run: run(() => L.print.exportPDF({ quick: true })) });
  cmd('properties', { label: '属性(&i)', icon: 'properties', run: run(() => D().properties()) });
  /* close every workbook (asking to save each), then return to the Start Center */
  cmd('exitApp', { label: '退出(&x)', run: async () => { for (;;) { const last = A().books.length <= 1; if (!(await A().closeWorkbook())) return; if (last) break; } if (window.VO) VO.exit(); } });

  /* ------------------------------------------------------------ Edit */
  /* the step's name (an English label, L.N at its O.tx) shown translated */
  cmd('undo', { label: '撤消(&U)', icon: 'undo', key: 'Ctrl+Z', get menuLabel() { return wb() && wb().undo.canUndo() ? L.t('撤消 {0}(&U)', L.t(wb().undo.undoLabel)) : "无法撤消(&U)"; }, enabled: () => wb() && wb().undo.canUndo(), run: () => A().undo() });
  cmd('redo', { label: '重做(&R)', icon: 'redo', key: 'Ctrl+Y', get menuLabel() { return wb() && wb().undo.canRedo() ? L.t('重做 {0}(&R)', L.t(wb().undo.redoLabel)) : A().lastCommand ? L.t('重复 {0}(&R)', ui.stripAmp(A().lastCommand.label)) : "无法重复(&R)"; }, enabled: () => wb() && (wb().undo.canRedo() || !!A().lastCommand), run: () => A().redo() });
  cmd('cut', { label: '剪切(&t)', icon: 'cut', key: 'Ctrl+X', run: run(() => CL().copy(null, true)) });
  cmd('copy', { label: '复制(&C)', icon: 'copy', key: 'Ctrl+C', run: run(() => CL().copy(null, false)) });
  cmd('paste', { label: '粘贴(&P)', icon: 'paste', key: 'Ctrl+V', run: run(() => CL().pasteCommand()) });
  cmd('officeClipboard', { label: 'Office 剪贴板(&b)...', icon: 'clipboard', run: () => L.panes.task.show('clipboard') });
  cmd('pasteSpecial', { label: '选择性粘贴(&S)...', enabled: () => CL().canPaste(), run: run(() => D().pasteSpecial()) });
  cmd('pasteValues', { label: '粘贴值(&V)', enabled: () => CL().canPaste(), run: run(() => CL().pasteInternal({ what: 'values' })) });
  cmd('pasteHyperlink', { label: '粘贴为超链接(&H)', enabled: () => CL().canPaste(), run: run(() => A().pasteAsHyperlink()) });
  cmd('fillDown', { label: '向下(&D)', key: 'Ctrl+D', run: run(() => A().fillDir('down')) });
  cmd('fillRight', { label: '右(&R)', key: 'Ctrl+R', run: run(() => A().fillDir('right')) });
  cmd('fillUp', { label: '向上(&U)', run: run(() => A().fillDir('up')) });
  cmd('fillLeft', { label: '左(&L)', run: run(() => A().fillDir('left')) });
  cmd('fillSeries', { label: '序列(&S)...', run: run(() => D().series()) });
  cmd('fillJustify', { label: '两端对齐(&J)', run: run(() => A().justifyFill()) });
  cmd('fillAcross', { label: '至同组工作表(&A)...', enabled: () => (A().selectedSheets() || []).length > 1, run: run(() => D().fillAcross()) });
  cmd('clearAll', { label: '全部(&A)', run: run(() => { if (guard()) return; O.clear(sh(), rgs(), 'all'); }) });
  cmd('clearFormats', { label: '格式(&F)', run: run(() => { if (guard('formatCells')) return; O.clear(sh(), rgs(), 'formats'); }) });
  cmd('clearContents', { label: '内容(&C)', key: 'Del', run: run(() => { if (guard()) return; const d = G().selectedObject(); if (d) { A().deleteObject(d); return; } O.clear(sh(), rgs(), 'contents'); }) });
  cmd('clearComments', { label: '批注(&m)', run: run(() => O.clear(sh(), rgs(), 'comments')) });
  cmd('clearHyperlinks', { label: '超链接(&H)', run: run(() => O.clear(sh(), rgs(), 'hyperlinks')) });
  cmd('deleteCells', { label: '删除(&D)...', key: 'Ctrl+-', run: run(() => A().deleteSelection()) });
  cmd('deleteSheet', { label: '删除工作表(&l)', icon: 'deleteSheet', run: run(() => A().deleteSheets()) });
  cmd('moveCopySheet', { label: '移动或复制工作表(&M)...', run: run(() => D().moveCopySheet()) });
  cmd('find', { label: '查找(&F)...', icon: 'find', key: 'Ctrl+F', run: run(() => D().find('find')) });
  cmd('replace', { label: '替换(&e)...', icon: 'replace', key: 'Ctrl+H', run: run(() => D().find('replace')) });
  cmd('goTo', { label: '定位(&G)...', key: 'Ctrl+G', run: run(() => D().goTo()) });
  cmd('goToSpecial', { label: '定位条件(&S)...', run: run(() => D().goToSpecial()) });
  cmd('links', { label: '链接(&k)...', enabled: () => !!(wb() && wb().externalLinks && wb().externalLinks.length), run: run(() => D().links()) });
  cmd('selectAll', { label: '全选(&A)', key: 'Ctrl+A', run: run(() => A().selectAllSmart()) });

  /* ------------------------------------------------------------ View */
  cmd('viewNormal', { label: '普通(&N)', icon: 'normalGrid', radio: true, checked: () => !sh().view.pageBreakPreview, run: run(() => A().setPageBreakPreview(false)) });
  cmd('viewPageBreak', { label: '分页预览(&P)', icon: 'pageBreakView', radio: true, checked: () => !!sh().view.pageBreakPreview, run: run(() => A().setPageBreakPreview(true)) });
  cmd('taskPane', { label: '任务窗格(&k)', key: 'Ctrl+F1', checked: () => !L.$('#taskpane').hidden, run: () => A().toggleTaskPane() });
  cmd('formulaBar', { label: '编辑栏(&F)', checked: () => A().opts.formulaBar !== false, run: () => { A().opts.formulaBar = A().opts.formulaBar === false; A().applyOpts(); } });
  cmd('statusBar', { label: '状态栏(&S)', checked: () => A().opts.statusBar !== false, run: () => { A().opts.statusBar = A().opts.statusBar === false; A().applyOpts(); } });
  cmd('headerFooter', { label: '页眉和页脚(&H)...', run: run(() => D().pageSetup(2)) });
  cmd('showComments', { label: '批注(&C)', icon: 'comment', checked: () => !!wb()._showComments, run: run(() => { wb()._showComments = !wb()._showComments; A().toggleToolbar('reviewing', wb()._showComments ? true : undefined); }) });
  cmd('fullScreen', { label: '全屏显示(&u)', icon: 'fullScreen', checked: () => L.$('#app').classList.contains('fullscreen'), run: () => A().toggleFullScreen() });
  cmd('zoomDlg', { label: '缩放(&Z)...', icon: 'zoom', run: run(() => D().zoom()) });
  cmd('zoomIn', { label: '放大', run: () => G().setZoom((sh().view.zoom || 100) + 10) });
  cmd('zoomOut', { label: '缩小', run: () => G().setZoom((sh().view.zoom || 100) - 10) });
  cmd('showFormulas', { label: '公式审核模式(&M)', icon: 'showFormulas', key: 'Ctrl+`', checked: () => !!sh().view.formulas, run: () => { sh().view.formulas = !sh().view.formulas; G().paint(); ui.refresh(); } });
  cmd('gridlinesView', { label: '网格线(&G)', checked: () => sh().view.grid !== false, run: () => { sh().view.grid = sh().view.grid === false; G().paint(); } });
  cmd('headingsView', { label: '行列标题(&H)', checked: () => sh().view.headings !== false, run: () => { sh().view.headings = sh().view.headings === false; G().paint(); } });
  for (const [id, label] of [['standard', '常用(&S)'], ['formatting', '格式设置(&F)'], ['drawing', '绘图(&D)'], ['chart', '图表(&C)'], ['auditing', '公式审核(&A)'], ['reviewing', '审阅(&R)'], ['picture', '图片(&P)'], ['borders', '边框(&B)'], ['pivot', '数据透视表(&b)'], ['forms', '窗体(&m)'], ['list', '列表(&L)']]) {
    cmd('tb_' + id, { label, checked: () => A().toolbarVisible(id), run: () => A().toggleToolbar(id) });
  }
  cmd('customize', { label: '自定义(&C)...', run: run(() => D().customize()) });

  /* ------------------------------------------------------------ Insert */
  cmd('insertCells', { label: '单元格(&l)...', icon: 'insertCells', key: 'Ctrl+Shift++', run: run(() => A().insertSelection()) });
  cmd('insertRows', { label: '行(&R)', icon: 'insertRows', run: run(() => A().insertLines('r')) });
  cmd('insertCols', { label: '列(&C)', icon: 'insertCols', run: run(() => A().insertLines('c')) });
  cmd('insertSheet', { label: '工作表(&W)', icon: 'insertSheet', key: 'Shift+F11', run: run(() => A().insertSheet()) });
  cmd('insertChart', { label: '图表(&h)...', icon: 'chartWizard', key: 'F11', run: run(() => L.chartWizard.open()) });
  cmd('chartWizard', { label: '图表向导', tip: '图表向导', icon: 'chartWizard', run: run(() => L.chartWizard.open()) });
  cmd('insertSymbol', { label: '符号(&S)...', icon: 'symbol', run: run(() => D().symbol()) });
  cmd('insertPageBreak', { get label() { return A().pageBreakAtSel() ? '删除分页符(&B)' : '分页符(&B)'; }, icon: 'pageBreak', run: run(() => A().togglePageBreak()) });
  cmd('resetPageBreaks', { label: '重置所有分页符', run: run(() => A().resetPageBreaks()) });
  cmd('insertFunction', { label: '函数(&F)...', icon: 'fx', key: 'Shift+F3', run: () => D().insertFunction() });
  cmd('functionArgs', { label: '函数参数', run: () => D().functionArgs() });
  cmd('nameDefine', { label: '定义(&D)...', key: 'Ctrl+F3', run: run(() => D().defineName()) });
  cmd('pasteName', { label: '粘贴(&P)...', key: 'F3', run: () => D().pasteName() });
  cmd('nameCreate', { label: '创建(&C)...', key: 'Ctrl+Shift+F3', run: run(() => D().createNames()) });
  cmd('nameApply', { label: '应用(&A)...', run: run(() => D().applyNames()) });
  cmd('insertComment', { get label() { const s = sel(); return sh().comments.get(M.key(s.r, s.c)) ? '编辑批注(&m)' : '批注(&m)'; }, icon: 'comment', key: 'Shift+F2', run: run(() => A().editComment()) });
  cmd('deleteComment', { label: '删除批注(&m)', enabled: () => { const s = sel(); return !!sh().comments.get(M.key(s.r, s.c)); }, run: run(() => { const s = sel(); O.tx(wb(), '删除批注', () => O.setComment(sh(), s.r, s.c, null)); }) });
  cmd('showHideComment', { label: '显示/隐藏批注(&m)', enabled: () => { const s = sel(); return !!sh().comments.get(M.key(s.r, s.c)); }, run: run(() => { const s = sel(); const cm = sh().comments.get(M.key(s.r, s.c)); O.tx(wb(), '显示批注', () => O.setComment(sh(), s.r, s.c, Object.assign({}, cm, { visible: !cm.visible }))); }) });
  cmd('insertPicture', { label: '来自文件(&F)...', icon: 'picture', run: run(() => A().insertPicture()) });
  cmd('insertClipArt', { label: '剪贴画(&C)...', icon: 'clipart', run: run(() => A().insertPicture()) });
  cmd('insertShape', { label: '自选图形(&A)', icon: 'autoshapes', run: run(() => L.drawing.startTool('rect')) });
  cmd('insertTextBox', { label: '文本框(&x)', icon: 'textbox', run: run(() => L.drawing.startTool('textbox')) });
  cmd('hyperlink', { label: '超链接(&i)...', icon: 'hyperlink', key: 'Ctrl+K', run: run(() => D().hyperlink()) });
  cmd('removeHyperlink', { label: '删除超链接(&R)', enabled: () => { const s = sel(); return sh().links.some((l) => M.rangeContains(l.ref, s.r, s.c)); }, run: run(() => O.clear(sh(), rgs(), 'hyperlinks')) });

  /* ------------------------------------------------------------ Format */
  cmd('formatCells', { label: '单元格(&e)...', icon: 'formatCells', key: 'Ctrl+1', run: run(() => { const d = G().selectedObject(); if (d) { L.drawing.format(d); return; } if (guard('formatCells')) return; D().formatCells(); }) });
  cmd('rowHeight', { label: '高度(&e)...', icon: 'rowHeight', run: run(() => { if (guard('formatRows')) return; D().rowHeight(); }) });
  cmd('rowAutofit', { label: '自动调整(&A)', run: run(() => { if (guard('formatRows')) return; O.autofitRows(sh(), A().selRows()); }) });
  cmd('rowHide', { label: '隐藏(&H)', key: 'Ctrl+9', run: run(() => { if (guard('formatRows')) return; O.hideLines(sh(), 'r', A().selRows(), true); }) });
  cmd('rowUnhide', { label: '取消隐藏(&U)', key: 'Ctrl+Shift+(', run: run(() => { if (guard('formatRows')) return; O.hideLines(sh(), 'r', A().selRows(true), false); }) });
  cmd('colWidth', { label: '宽度(&i)...', icon: 'colWidth', run: run(() => { if (guard('formatColumns')) return; D().colWidth(); }) });
  cmd('colAutofit', { label: '根据所选内容自动调整(&A)', run: run(() => { if (guard('formatColumns')) return; const s = rgs()[0]; O.autofitCols(sh(), A().selCols(), G().isWholeCols(s) ? null : { r1: s.r1, r2: s.r2 }); }) });
  cmd('colHide', { label: '隐藏(&H)', key: 'Ctrl+0', run: run(() => { if (guard('formatColumns')) return; O.hideLines(sh(), 'c', A().selCols(), true); }) });
  cmd('colUnhide', { label: '取消隐藏(&U)', key: 'Ctrl+Shift+)', run: run(() => { if (guard('formatColumns')) return; O.hideLines(sh(), 'c', A().selCols(true), false); }) });
  cmd('colStandard', { label: '标准列宽(&S)...', run: run(() => D().standardWidth()) });
  cmd('sheetRename', { label: '重命名(&R)', run: run(() => A().renameSheetInline()) });
  cmd('sheetHide', { label: '隐藏(&H)', run: run(() => A().hideSheets()) });
  cmd('sheetUnhide', { label: '取消隐藏(&U)...', enabled: () => wb().sheets.some((s) => s.state === 'hidden'), run: run(() => D().unhideSheet()) });
  cmd('sheetBackground', { get label() { return sh().background ? '删除背景(&k)' : '背景(&B)...'; }, run: run(() => A().sheetBackground()) });
  cmd('tabColor', { label: '工作表标签颜色(&T)...', icon: 'tabColor', run: run(() => D().tabColor()) });
  cmd('autoFormat', { label: '自动套用格式(&A)...', icon: 'tableAutoFormat', run: run(() => { if (guard('formatCells')) return; D().autoFormat(); }) });
  cmd('conditionalFormat', { label: '条件格式(&d)...', run: run(() => { if (guard('formatCells')) return; D().conditionalFormat(); }) });
  cmd('styleDlg', { label: '样式(&S)...', run: run(() => { if (guard('formatCells')) return; D().style(); }) });
  cmd('formatObject', { label: '选定对象(&O)...', enabled: () => !!G().selectedObject(), run: run(() => L.drawing.format(G().selectedObject())) });

  /* formatting toolbar */
  cmd('bold', { label: '加粗(&B)', icon: 'bold', key: 'Ctrl+B', checked: () => !!(act().font && act().font.b), run: run(() => { if (guard('formatCells')) return; O.toggleFont(sh(), rgs(), 'b', act()); }) });
  cmd('italic', { label: '倾斜(&I)', icon: 'italic', key: 'Ctrl+I', checked: () => !!(act().font && act().font.i), run: run(() => { if (guard('formatCells')) return; O.toggleFont(sh(), rgs(), 'i', act()); }) });
  cmd('underline', { label: '下划线(&U)', icon: 'underline', key: 'Ctrl+U', checked: () => !!(act().font && act().font.u), run: run(() => { if (guard('formatCells')) return; O.toggleFont(sh(), rgs(), 'u', act()); }) });
  cmd('strike', { label: '删除线(&k)', icon: 'strike', key: 'Ctrl+5', checked: () => !!(act().font && act().font.strike), run: run(() => { if (guard('formatCells')) return; O.toggleFont(sh(), rgs(), 'strike', act()); }) });
  const halign = (h) => run(() => { if (guard('formatCells')) return; const cur = act().align && act().align.h; O.format(sh(), rgs(), { align: { h: cur === h ? undefined : h } }, L.N('对齐')); });
  cmd('alignLeft', { label: '左对齐(&L)', icon: 'alignLeft', checked: () => (act().align || {}).h === 'left', run: halign('left') });
  cmd('alignCenter', { label: '居中(&C)', icon: 'alignCenter', checked: () => (act().align || {}).h === 'center', run: halign('center') });
  cmd('alignRight', { label: '右对齐(&R)', icon: 'alignRight', checked: () => (act().align || {}).h === 'right', run: halign('right') });
  cmd('mergeCenter', { label: '合并及居中(&M)', tip: '合并及居中', icon: 'mergeCenter', checked: () => { const s = sel(); return !!LY.mergeAt(sh(), s.r, s.c); }, run: run(() => {
    if (guard('formatCells')) return;
    const s = sel();
    if (LY.mergeAt(sh(), s.r, s.c) && rgs().every((r) => sh().merges.some((m) => m.r1 === r.r1 && m.c1 === r.c1 && m.r2 === r.r2 && m.c2 === r.c2))) { O.unmerge(sh(), rgs()); return; }
    A().confirmMerge(() => O.merge(sh(), rgs(), { center: true }));
  }) });
  cmd('currencyStyle', { label: '货币样式(&C)', icon: 'currency', run: fmt({ nf: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)' }, '货币样式') });
  cmd('percentStyle', { label: '百分比样式(&P)', icon: 'percent', key: 'Ctrl+Shift+%', run: fmt({ nf: '0%' }, '百分比样式') });
  cmd('commaStyle', { label: '千位分隔样式(&m)', icon: 'comma', run: fmt({ nf: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)' }, '千位分隔样式') });
  cmd('incDecimal', { label: '增加小数位数(&I)', icon: 'incDecimal', run: run(() => { if (guard('formatCells')) return; const s = sel(); const v = sh().val(s.r, s.c); const code = NF.withDecimals(act().nf || 'General', 1, v); O.format(sh(), rgs(), { nf: code }, '增加小数位数'); }) });
  cmd('decDecimal', { label: '减少小数位数(&D)', icon: 'decDecimal', run: run(() => { if (guard('formatCells')) return; const s = sel(); const v = sh().val(s.r, s.c); const code = NF.withDecimals(act().nf || 'General', -1, v); O.format(sh(), rgs(), { nf: code }, '减少小数位数'); }) });
  cmd('incIndent', { label: '增加缩进量(&n)', icon: 'incIndent', run: run(() => { if (guard('formatCells')) return; O.format(sh(), rgs(), (st) => ({ align: { indent: Math.min(15, ((st.align && st.align.indent) || 0) + 1), h: st.align && (st.align.h === 'right' || st.align.h === 'distributed') ? st.align.h : 'left' } }), '增加缩进量'); }) });
  cmd('decIndent', { label: '减少缩进量(&d)', icon: 'decIndent', run: run(() => { if (guard('formatCells')) return; O.format(sh(), rgs(), (st) => { const n = Math.max(0, ((st.align && st.align.indent) || 0) - 1); return { align: { indent: n || undefined } }; }, '减少缩进量'); }) });
  cmd('wrapText', { label: '自动换行(&W)', icon: 'wrapText', checked: () => !!(act().align && act().align.wrap), run: run(() => { if (guard('formatCells')) return; const on = !(act().align && act().align.wrap); O.format(sh(), rgs(), { align: { wrap: on || undefined } }, '自动换行'); O.tx(wb(), L.N('自动调整'), () => { for (const r of A().selRows()) O.autoRow(sh(), r); }); }) });
  cmd('fillColorApply', { label: '填充颜色', tip: '填充颜色', icon: 'fillColor', swatch: () => A().lastFill || '#FFFF00', run: run(() => A().applyFill(A().lastFill || '#FFFF00')) });
  cmd('fontColorApply', { label: '字体颜色', tip: '字体颜色', icon: 'fontColor', swatch: () => A().lastFontColor || '#FF0000', run: run(() => A().applyFontColor(A().lastFontColor || '#FF0000')) });
  cmd('bordersApply', { label: '边框', tip: '边框', icon: 'bordersBottom', run: run(() => A().applyBorderPreset(A().lastBorder || 'bottom')) });
  cmd('formatPainter', { label: '格式刷', tip: '格式刷', icon: 'painter', checked: () => A().painter.on(), run: (arg, ev) => A().painter.click(ev) });
  /* number format shortcuts */
  cmd('fmtGeneral', { label: '常规', key: 'Ctrl+Shift+~', run: fmt({ nf: 'General' }, L.N('格式')) });
  cmd('fmtCurrency', { label: '货币', key: 'Ctrl+Shift+$', run: fmt({ nf: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)' }, L.N('格式')) });
  cmd('fmtExp', { label: '科学记数', key: 'Ctrl+Shift+^', run: fmt({ nf: '0.00E+00' }, L.N('格式')) });
  cmd('fmtDate', { label: '日期', key: 'Ctrl+Shift+#', run: fmt({ nf: 'd-mmm-yy' }, L.N('格式')) });
  cmd('fmtTime', { label: '时间', key: 'Ctrl+Shift+@', run: fmt({ nf: 'h:mm AM/PM' }, L.N('格式')) });
  cmd('fmtNumber', { label: '数字', key: 'Ctrl+Shift+!', run: fmt({ nf: '#,##0.00' }, L.N('格式')) });
  cmd('borderOutline', { label: '外边框', key: 'Ctrl+Shift+&', run: run(() => A().applyBorderPreset('outside')) });
  cmd('borderNone', { label: '清除边框', key: 'Ctrl+Shift+_', run: run(() => A().applyBorderPreset('none')) });

  /* ------------------------------------------------------------ Tools */
  cmd('errorChecking', { label: '错误检查(&E)...', run: run(() => D().errorChecking()) });
  cmd('protectSheet', { get label() { return sh().protection ? '撤消工作表保护(&p)...' : '保护工作表(&P)...'; }, icon: 'protect', run: run(() => D().protectSheet()) });
  cmd('protectWorkbook', { get label() { return wb().protection ? '撤消工作簿保护(&W)...' : '保护工作簿(&W)...'; }, run: run(() => D().protectWorkbook()) });
  cmd('allowEditRanges', { label: '允许用户编辑区域(&n)...', run: run(() => ui.msg('Ledger 的工作表保护遵循各单元格的“锁定”设置（“格式 ▸ 单元格 ▸ 保护”）。请先解除允许编辑的单元格的锁定，再保护工作表。', { icon: 'info' })) });
  cmd('goalSeek', { label: '单变量求解(&G)...', icon: 'goalSeek', run: run(() => D().goalSeek()) });
  cmd('scenarios', { label: '方案(&e)...', run: run(() => D().scenarios()) });
  cmd('tracePrecedents', { label: '追踪引用单元格(&P)', icon: 'tracePrec', run: run(() => A().trace('prec')) });
  cmd('traceDependents', { label: '追踪从属单元格(&D)', icon: 'traceDep', run: run(() => A().trace('dep')) });
  cmd('traceError', { label: '追踪错误(&E)', icon: 'traceError', run: run(() => A().trace('err')) });
  cmd('removeArrows', { label: '移去所有箭头(&A)', icon: 'removeArrows', run: () => A().trace(null) });
  cmd('circleInvalid', { label: '圈释无效数据(&C)', icon: 'circleInvalid', run: run(() => A().circleInvalid(true)) });
  cmd('clearCircles', { label: '清除无效数据标识圈(&l)', icon: 'clearCircles', run: () => A().circleInvalid(false) });
  cmd('evaluateFormula', { label: '公式求值(&F)', icon: 'evaluate', run: run(() => D().evaluateFormula()) });
  cmd('watchWindow', { label: '显示监视窗口(&W)', icon: 'watch', run: run(() => D().watchWindow()) });
  cmd('addWatch', { label: '添加监视(&W)', icon: 'watch', run: run(() => D().watchWindow(true)) });
  cmd('calcNow', { label: '立即计算(&N)', icon: 'calcNow', key: 'F9', run: run(() => A().calculateNow()) });
  cmd('calcSheet', { label: '计算工作表', key: 'Shift+F9', run: run(() => A().calculateNow()) });
  cmd('autoCorrectDlg', { label: '自动更正选项(&A)...', run: run(() => D().autoCorrect()) });
  cmd('options', { label: '选项(&O)...', run: run(() => D().options()) });
  cmd('macros', { label: '宏(&M)...', key: 'Alt+F8', run: () => ui.msg(wb().macroEnabled ? '此工作簿包含宏。保存为 .xlsm 时，Ledger 会保留宏，但无法运行它们。' : 'Ledger 不运行宏。包含宏的工作簿 (.xlsm) 可打开和保存，其中的宏会原样保留。', { icon: 'info' }) });
  cmd('addIns', { label: '加载宏(&I)...', run: () => ui.msg('已内置分析工具库函数（工程、财务、日期）。', { icon: 'info' }) });

  /* ------------------------------------------------------------ Data */
  cmd('sortDlg', { label: '排序(&S)...', icon: 'sortDlg', run: run(() => { if (guard('sort')) return; D().sort(); }) });
  cmd('sortAsc', { label: '升序排序', tip: '升序排序', icon: 'sortAsc', run: run(() => { if (guard('sort')) return; A().quickSort(false); }) });
  cmd('sortDesc', { label: '降序排序', tip: '降序排序', icon: 'sortDesc', run: run(() => { if (guard('sort')) return; A().quickSort(true); }) });
  cmd('autoFilter', { label: '自动筛选(&F)', icon: 'autoFilter', checked: () => !!sh().autoFilter || !!A().tableAt(), run: run(() => { if (guard('autoFilter')) return; L.filter.toggle(); }) });
  cmd('showAll', { label: '全部显示(&S)', enabled: () => L.filter && L.filter.active(), run: run(() => L.filter.showAll()) });
  cmd('advancedFilter', { label: '高级筛选(&A)...', run: run(() => D().advancedFilter()) });
  cmd('dataForm', { label: '窗体(&o)...', icon: 'dataForm', run: run(() => D().dataForm()) });
  cmd('subtotals', { label: '分类汇总(&b)...', icon: 'subtotal', run: run(() => { if (guard()) return; D().subtotals(); }) });
  cmd('validation', { label: '数据验证(&l)...', icon: 'validation', run: run(() => { if (guard()) return; D().validation(); }) });
  cmd('dataTable', { label: '模拟运算表(&T)...', run: run(() => D().dataTable()) });
  cmd('textToColumns', { label: '分列(&e)...', icon: 'textToCols', run: run(() => { if (guard()) return; D().textToColumns(); }) });
  cmd('consolidate', { label: '合并计算(&n)...', run: run(() => D().consolidate()) });
  cmd('groupRows', { label: '组合(&G)...', icon: 'group', key: 'Alt+Shift+ArrowRight', run: run(() => A().group(1)) });
  cmd('ungroupRows', { label: '取消组合(&U)...', icon: 'ungroup', key: 'Alt+Shift+ArrowLeft', run: run(() => A().group(-1)) });
  cmd('hideDetail', { label: '隐藏明细数据(&H)', run: run(() => A().outlineDetail(false)) });
  cmd('showDetail', { label: '显示明细数据(&S)', run: run(() => A().outlineDetail(true)) });
  cmd('autoOutline', { label: '自动建立分级显示(&A)', run: run(() => A().autoOutline()) });
  cmd('clearOutline', { label: '清除分级显示(&C)', run: run(() => A().clearOutline()) });
  cmd('outlineSettings', { label: '设置(&e)...', run: run(() => D().outlineSettings()) });
  cmd('createList', { label: '创建列表(&C)...', key: 'Ctrl+L', run: run(() => D().createList()) });
  cmd('totalRow', { label: '汇总行(&T)', checked: () => { const t = A().tableAt(); return !!(t && t.totals); }, enabled: () => !!A().tableAt(), run: run(() => A().toggleTotalRow()) });
  cmd('convertToRange', { label: '转换为区域(&v)', enabled: () => !!A().tableAt(), run: run(() => A().convertTableToRange()) });
  cmd('importText', { label: '导入文本文件(&T)...', icon: 'open', run: async () => { const files = await L.pickFiles('.csv,.txt,.tsv,.prn,.tab'); if (files.length) A().openFile(files[0], null, { wizard: true }).catch((e) => A().error(e)); } });
  cmd('pivotTable', { label: '数据透视表和数据透视图(&P)...', icon: 'pivot', run: run(() => D().pivotTable()) });
  cmd('refreshData', { label: '刷新数据(&R)', run: run(() => A().calculateNow()) });

  /* ------------------------------------------------------------ Window */
  cmd('freezePanes', { get label() { return sh().view.freeze ? '取消冻结窗格(&f)' : '冻结窗格(&F)'; }, icon: 'freeze', run: run(() => A().toggleFreeze()) });
  cmd('splitWindow', { get label() { return sh().view.freeze ? '取消拆分(&S)' : '拆分(&S)'; }, icon: 'split', run: run(() => A().toggleFreeze(true)) });
  cmd('newWindow', { label: '新建窗口(&N)', run: run(() => ui.msg('Ledger 为每个工作簿显示一个窗口。使用“窗口”菜单可在打开的工作簿之间切换。', { icon: 'info' })) });
  cmd('arrangeWindows', { label: '排列(&A)...', run: run(() => ui.msg('已打开的工作簿列在“窗口”菜单底部。', { icon: 'info' })) });

  /* ------------------------------------------------------------ Help */
  cmd('help', { label: 'Ledger 帮助(&H)', icon: 'help', key: 'F1', run: () => L.panes.task.show('help') });
  cmd('keyboardHelp', { label: '键盘快捷键(&K)', run: () => L.panes.task.show('keys') });
  cmd('functionsHelp', { label: '函数参考(&F)', run: () => D().insertFunction() });
  cmd('about', { label: '关于 Ledger(&A)', run: () => D().about() });

  /* ------------------------------------------------------------ sheets & misc */
  cmd('nextSheet', { label: '下一张工作表', key: 'Ctrl+PageDown', run: () => A().nextSheet(1) });
  cmd('prevSheet', { label: '上一张工作表', key: 'Ctrl+PageUp', run: () => A().nextSheet(-1) });
  cmd('autoSum', { label: '自动求和', tip: '自动求和', icon: 'autoSum', key: 'Alt+=', run: run(() => { if (guard()) return; A().autoSum('SUM'); }) });
  cmd('insertDate', { label: '插入日期', key: 'Ctrl+;', run: run(() => { if (guard()) return; const s = sel(); O.tx(wb(), L.N('键入'), () => O.enter(sh(), s.r, s.c, NF.text('m/d/yyyy', Math.floor(NF.jsDateToSerial(new Date(), wb().date1904))))); }) });
  cmd('insertTime', { label: '插入时间', key: 'Ctrl+Shift+:', run: run(() => { if (guard()) return; const s = sel(); O.tx(wb(), L.N('键入'), () => O.enter(sh(), s.r, s.c, NF.text('h:mm AM/PM', NF.jsDateToSerial(new Date(), wb().date1904) % 1))); }) });
  cmd('copyAbove', { label: '复制上方单元格的公式', key: "Ctrl+'", run: run(() => A().copyFromAbove(false)) });
  cmd('copyValueAbove', { label: '复制上方单元格的值', key: 'Ctrl+Shift+"', run: run(() => A().copyFromAbove(true)) });
  cmd('selectRegion', { label: '选择当前区域', key: 'Ctrl+Shift+*', run: () => { const s = sel(); G().selectRange(O.currentRegion(sh(), s.r, s.c), { r: s.r, c: s.c }); } });
  cmd('selectArray', { label: '选择当前数组', key: 'Ctrl+/', run: () => { const s = sel(); const a = O.arrayAt(sh(), s.r, s.c); if (a) G().selectRange(a, { r: s.r, c: s.c }); } });
  cmd('pickList', { label: '从下拉列表中选择(&L)...', run: () => A().pickList() });
  cmd('dropDownList', { label: '从下拉列表中选择(&L)...', key: 'Alt+ArrowDown', run: () => A().pickList() });
  cmd('repeatLast', { label: '重复', key: 'F4', run: () => A().redo() });
  cmd('contextKey', { label: '快捷菜单', key: 'Shift+F10', run: () => A().contextMenuAtSel() });
  cmd('newSheetChart', { label: '新工作表上的图表', run: run(() => L.chartWizard.open({ sheet: true })) });
  void hasMulti; void wholeCols; void wholeRows; void notEditing; void F; void C; void MAXR; void MAXC;
})(typeof window !== 'undefined' ? window : globalThis);
