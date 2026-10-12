/* Lectern — command definitions (menus, toolbars, shortcuts). */
(function () {
  'use strict';
  const L = window.L;
  const ui = L.ui;
  const C = ui.cmd;
  const E = () => L.ed;
  const A = () => L.app;
  const T = L.txt;
  const editing = () => L.te.active();
  const hasSel = () => E().selected().length > 0;
  const canRotate = () => hasSel() && E().selected().every(L.model.canRotate);
  const hasSlide = () => !!E().slide();
  const normal = () => A().view === 'normal' || A().view === 'master';
  const textCtx = () => editing() || L.fmt.textShapes().length > 0;
  const fmtState = () => (textCtx() ? L.fmt.state() : null);
  const one = (type) => () => { const s = E().selected(); return s.length > 0 && s.every((x) => (typeof type === 'function' ? type(x) : x.type === type)); };
  const isPic = one('image'), isTable = () => { const s = E().selected(); return s.length === 1 && s[0].type === 'table'; }, isWA = one('wordart');
  const slideCtx = () => A().focusArea === 'slides' || A().view === 'sorter';

  /* ---------- File ---------- */
  /* the toolbar button makes a blank presentation; the menu and Ctrl+N show Blank and the templates */
  C('new', { label: '新建(&N)...', menuLabel: '新建(&N)...', icon: 'new', key: 'Ctrl+N', tip: 'New', run: (arg, ev) => { if ((ev && ev.currentTarget && ev.currentTarget.closest && ev.currentTarget.closest('.toolbar')) || !window.VO) A().newPresentation(); else VO.newDialog('lectern', (id) => (id ? A().newFromTemplate(id) : A().newPresentation())); } });
  C('open', { label: '打开(&O)...', icon: 'open', key: 'Ctrl+O', tip: 'Open', run: () => A().openDialog() });
  C('close', { label: '关闭(&C)', run: () => A().closePresentation() });
  C('save', { label: '保存(&S)', icon: 'save', key: 'Ctrl+S', tip: 'Save', run: () => A().save() });
  C('saveAs', { label: '另存为(&A)...', run: () => A().saveAs() });
  C('saveWeb', { label: '另存为网页(&G)...', icon: 'webPage', run: () => A().exportAs('html') });
  C('pageSetup', { label: '页面设置(&U)...', run: () => L.dlg.pageSetup() });
  C('printPreview', { label: '打印预览(&V)', icon: 'preview', tip: 'Print Preview', run: () => A().setView('preview') });
  C('print', { label: '打印(&P)...', icon: 'print', key: 'Ctrl+P', tip: 'Print', run: () => L.dlg.print() });
  C('properties', { label: '属性(&I)', icon: 'properties', run: () => L.dlg.properties() });
  C('exitApp', { label: '退出(&X)', icon: 'exit', run: () => A().exit() });

  /* ---------- Edit ---------- */
  C('undo', { label: '撤销(&U)', icon: 'undo', key: 'Ctrl+Z', get menuLabel() { return L.hist.undo.length ? '撤销 ' + L.hist.undo[L.hist.undo.length - 1].label : "无法撤销(&C)"; }, enabled: () => L.hist.undo.length > 0, run: () => A().undo() });
  C('redo', { label: '恢复(&R)', icon: 'redo', key: 'Ctrl+Y', get menuLabel() { return L.hist.redo.length ? '恢复 ' + L.hist.redo[L.hist.redo.length - 1].label : "无法恢复(&C)"; }, enabled: () => L.hist.redo.length > 0, run: () => A().redo() });
  C('cut', { label: '剪切(&T)', icon: 'cut', key: 'Ctrl+X', enabled: () => editing() || hasSel() || slideCtx(), run: () => A().cmdClipboard('cut') });
  C('copy', { label: '复制(&C)', icon: 'copy', key: 'Ctrl+C', enabled: () => editing() || hasSel() || slideCtx(), run: () => A().cmdClipboard('copy') });
  C('paste', { label: '粘贴(&P)', icon: 'paste', key: 'Ctrl+V', enabled: () => !!A().clip, run: () => A().paste() });
  C('pasteSpecial', { label: '选择性粘贴(&S)...', enabled: () => !!A().clip, run: () => A().pasteSpecial() });
  C('officeClipboard', { label: 'Office 剪贴板(&B)...', icon: 'clipboard', run: () => L.panes.task.show('clipboard') });
  C('clear', { label: '清除(&A)', key: 'Del', enabled: () => hasSel() || slideCtx(), run: () => (slideCtx() ? A().deleteSlides() : A().deleteSel()) });
  C('selectAll', { label: '全选(&L)', icon: 'selectAll', key: 'Ctrl+A', run: () => A().selectAll() });
  C('duplicate', { label: '创建副本(&L)', icon: 'duplicate', key: 'Ctrl+D', enabled: () => hasSel() || slideCtx(), run: () => (slideCtx() || !hasSel() ? A().duplicateSlides() : A().duplicateSel()) });
  C('deleteSlide', { label: '删除幻灯片(&D)', icon: 'delete', enabled: () => L.pres.slides.length > 0 && normal() || A().view === 'sorter', run: () => A().deleteSlides() });
  C('find', { label: '查找(&F)...', icon: 'find', key: 'Ctrl+F', run: () => L.dlg.find(false) });
  C('replace', { label: '替换(&E)...', icon: 'replace', key: 'Ctrl+H', run: () => L.dlg.find(true) });

  /* ---------- View ---------- */
  C('viewNormal', { label: '普通(&N)', icon: 'normalView', radio: true, checked: () => A().view === 'normal', tip: 'Normal View', run: () => A().setView('normal') });
  C('viewSorter', { label: '幻灯片浏览(&D)', icon: 'sorterView', radio: true, checked: () => A().view === 'sorter', tip: 'Slide Sorter View', run: () => A().setView('sorter') });
  C('viewNotes', { label: '备注页(&P)', icon: 'notesView', radio: true, checked: () => A().view === 'notes', run: () => A().setView('notes') });
  C('viewMaster', { label: '幻灯片母版(&S)', icon: 'master', checked: () => A().view === 'master', run: () => A().setView('master') });
  C('closeMaster', { label: '关闭母版视图(&C)', tbLabel: '&Close Master View', run: () => A().setView('normal') });
  C('insertTitleMaster', { label: '插入新标题母版(&T)', icon: 'newSlide', enabled: () => A().view === 'master' && !E().design().titleDeco, run: () => A().insertTitleMaster() });
  C('deleteTitleMaster', { label: '删除标题母版(&D)', icon: 'delete', enabled: () => A().view === 'master' && E().masterKind === 'title', run: () => A().deleteTitleMaster() });
  C('viewColor', { label: '彩色(&C)', radio: true, checked: () => !A().opts.gray && !A().opts.bw, run: () => A().setColorMode('color') });
  C('viewGray', { label: '灰度(&G)', icon: 'grayscale', radio: true, checked: () => !!A().opts.gray, tip: 'Color/Grayscale', run: () => A().setColorMode(A().opts.gray ? 'color' : 'gray') });
  C('viewBW', { label: '纯黑白(&P)', radio: true, checked: () => !!A().opts.bw, run: () => A().setColorMode('bw') });
  C('taskPane', { label: '任务窗格(&K)', icon: 'taskpane', key: 'Ctrl+F1', checked: () => A().taskOpen(), run: () => A().toggleTask() });
  C('slidesPane', { label: '幻灯片和大纲窗格', checked: () => A().leftOpen(), run: () => A().toggleLeft() });
  for (const [id, label] of [['standard', '常用(&S)'], ['formatting', '格式(&F)'], ['drawing', '绘图(&D)'], ['outlining', '大纲(&O)'], ['picture', '图片(&P)'], ['tables', '表格和边框(&T)'], ['wordart', '艺术字(&L)']]) {
    C('tb_' + id, { label, checked: () => A().tbShown(id), run: () => A().toggleToolbar(id) });
  }
  C('ruler', { label: '标尺(&R)', icon: 'ruler', checked: () => !!A().opts.ruler, run: () => { A().opts.ruler = !A().opts.ruler; A().saveOpts(); E().layout(); } });
  C('showGrid', { label: '网格和参考线(&G)...', icon: 'grid', tip: 'Show/Hide Grid', checked: () => E().grid.show, run: () => { E().grid.show = !E().grid.show; E().layout(); A().saveOpts(); } });
  C('gridGuides', { label: '网格和参考线(&I)...', icon: 'grid', run: () => L.dlg.grid() });
  C('headerFooter', { label: '页眉和页脚(&H)...', run: () => L.dlg.headerFooter() });
  C('zoomDlg', { label: '显示比例(&Z)...', icon: 'zoom', run: () => L.dlg.zoom() });
  C('zoomFit', { label: '适应(&F)', run: () => E().setZoom('fit') });

  /* ---------- Insert ---------- */
  C('newSlide', { label: '新幻灯片...', menuLabel: '新幻灯片(&N)...', tbLabel: '新幻灯片', icon: 'newSlide', key: 'Ctrl+M', tip: '新幻灯片', enabled: () => A().view !== 'master', run: () => A().newSlideCmd() });
  C('duplicateSlide', { label: '复制幻灯片(&D)', icon: 'duplicate', enabled: () => L.pres.slides.length > 0 && A().view !== 'master', run: () => A().duplicateSlides() });
  C('insertSlideNumber', { label: '幻灯片编号(&U)', icon: 'slideNumber', run: () => A().insertField('slidenum') });
  C('insertDateTime', { label: '日期和时间(&T)...', icon: 'date', run: () => L.dlg.dateTime() });
  C('insertSymbol', { label: '符号(&S)...', icon: 'symbol', run: () => { A().ensureTextTarget(); L.dlg.symbol(); } });
  C('slidesFromFiles', { label: '幻灯片（从文件）(&F)...', icon: 'insertSlides', run: () => L.dlg.slideFinder() });
  C('slidesFromOutline', { label: '幻灯片（从大纲）(&L)...', icon: 'outlineTxt', run: () => A().slidesFromOutline() });
  C('insertClipArt', { label: '剪贴画(&C)...', icon: 'clipart', tip: 'Insert Clip Art', run: () => L.panes.task.show('clipart') });
  C('insertPicture', { label: '来自文件(&F)...', icon: 'picture', tip: 'Insert Picture', run: (arg) => A().insertPictureDialog(arg && arg.ph) });
  C('photoAlbum', { label: '新建相册(&N)...', icon: 'photoAlbum', run: () => L.dlg.photoAlbum() });
  C('insertWordArt', { label: '艺术字(&L)...', icon: 'wordart', tip: 'Insert LettersArt', keys: 'WordArt', run: () => A().insertWordArtDialog() });
  C('insertDiagram', { label: '图示(&G)...', icon: 'diagram', tip: 'Insert IntelliArt Graphic', keys: 'IntelliArt SmartArt organization chart org', run: (arg) => L.diagram.gallery((k) => L.diagram.insert(k, arg && arg.ph)) });
  C('textBox', { label: '文本框(&X)', icon: 'textbox', tip: 'Text Box', checked: () => !!(E().tool && E().tool.kind === 'text' && !E().tool.vert), run: () => E().setTool(E().tool && E().tool.kind === 'text' ? null : { kind: 'text' }) });
  C('vTextBox', { label: '垂直文本框(&V)', icon: 'vtextbox', checked: () => !!(E().tool && E().tool.kind === 'text' && E().tool.vert), run: () => E().setTool({ kind: 'text', vert: true }) });
  C('insertChart', { label: '图表(&H)...', icon: 'chart', tip: 'Insert Chart', run: (arg) => A().insertChart(arg && arg.ph) });
  C('insertTable', { label: '表格(&B)...', icon: 'table', tip: 'Insert Table', run: (arg) => L.dlg.insertTable((r, c) => A().insertTable(r, c, arg && arg.ph)) });
  C('insertMedia', { label: '影片和声音(&V)...', icon: 'movie', run: () => ui.msg('Movies and sounds are not supported in this edition. Pictures, charts, tables and diagrams are available from the Insert menu.') });
  C('hyperlink', { label: '超链接(&I)...', icon: 'hyperlink', key: 'Ctrl+K', tip: 'Insert Hyperlink', enabled: () => editing() || hasSel(), run: () => L.dlg.hyperlink() });

  /* ---------- Format ---------- */
  C('fontDlg', { label: '字体(&F)...', icon: 'font', key: 'Ctrl+T', enabled: textCtx, run: () => L.dlg.font() });
  C('bulletsDlg', { label: '项目符号和编号(&B)...', enabled: textCtx, run: () => L.dlg.bullets() });
  C('lineSpacing', { label: '行距(&S)...', icon: 'lineSpacing', enabled: textCtx, run: () => L.dlg.lineSpacing() });
  C('changeCase', { label: '更改大小写(&H)...', icon: 'changeCase', key: 'Shift+F3', enabled: textCtx, run: () => L.dlg.changeCase() });
  C('replaceFonts', { label: '替换字体(&P)...', run: () => L.dlg.replaceFonts() });
  C('slideDesign', { label: '幻灯片设计(&E)...', tbLabel: 'Design', icon: 'design', tip: 'Slide Design', run: () => L.panes.task.show('design') });
  C('slideLayout', { label: '幻灯片版式(&L)...', icon: 'layout', run: () => L.panes.task.show('layout') });
  C('background', { label: '背景(&K)...', icon: 'background', enabled: hasSlide, run: () => L.dlg.background() });
  C('formatObject', { get label() { const s = E().selected()[0]; return s ? (s.type === 'image' ? '图片(&I)...' : s.type === 'wordart' ? '艺术字(&A)...' : s.ph ? '占位符(&C)...' : s.type === 'text' ? '文本框(&O)...' : s.type === 'table' ? '表格(&B)...' : '自选图形(&P)...') : '自选图形(&P)...'; }, icon: 'formatPicture', enabled: hasSel, run: (arg) => L.dlg.formatShape(arg) });
  C('alignLeft', { label: '左对齐(&L)', icon: 'alignLeft', key: 'Ctrl+L', radio: true, enabled: textCtx, checked: () => { const s = fmtState(); return !!s && s.algn === 'l'; }, run: () => A().setAlign('l') });
  C('alignCenter', { label: '居中(&C)', icon: 'alignCenter', key: 'Ctrl+E', radio: true, enabled: textCtx, checked: () => { const s = fmtState(); return !!s && s.algn === 'ctr'; }, run: () => A().setAlign('ctr') });
  C('alignRight', { label: '右对齐(&R)', icon: 'alignRight', key: 'Ctrl+R', radio: true, enabled: textCtx, checked: () => { const s = fmtState(); return !!s && s.algn === 'r'; }, run: () => A().setAlign('r') });
  C('justify', { label: '两端对齐(&J)', icon: 'justify', key: 'Ctrl+J', radio: true, enabled: textCtx, checked: () => { const s = fmtState(); return !!s && s.algn === 'just'; }, run: () => A().setAlign('just') });
  C('bold', { label: 'Bold', icon: 'bold', key: 'Ctrl+B', enabled: textCtx, checked: () => { const s = fmtState(); return !!s && s.b; }, run: () => A().toggleRun('b') });
  C('italic', { label: 'Italic', icon: 'italic', key: 'Ctrl+I', enabled: textCtx, checked: () => { const s = fmtState(); return !!s && s.i; }, run: () => A().toggleRun('i') });
  C('underline', { label: 'Underline', icon: 'underline', key: 'Ctrl+U', enabled: textCtx, checked: () => { const s = fmtState(); return !!s && s.u; }, run: () => A().toggleRun('u') });
  C('shadowText', { label: 'Shadow', icon: 'shadowText', enabled: textCtx, checked: () => { const s = fmtState(); return !!s && s.shd; }, run: () => A().toggleRun('shd') });
  C('superscript', { label: '上标', key: 'Ctrl+Shift+=', enabled: textCtx, run: () => A().toggleBaseline(30000) });
  C('subscript', { label: '下标', key: 'Ctrl+=', enabled: textCtx, run: () => A().toggleBaseline(-25000) });
  C('resetChar', { label: '重设字符格式', key: 'Ctrl+Space', enabled: textCtx, run: () => L.fmt.run((r) => ({ t: r.t, fld: r.fld, link: r.link }), 'Reset Formatting') });
  C('growFont', { label: '增大字号', icon: 'growFont', key: 'Ctrl+Shift+>', enabled: textCtx, run: () => A().stepFont(1) });
  C('shrinkFont', { label: '减小字号', icon: 'shrinkFont', key: 'Ctrl+Shift+<', enabled: textCtx, run: () => A().stepFont(-1) });
  C('numbering', { label: '编号', icon: 'numbering', enabled: textCtx, checked: () => { const s = fmtState(); return !!s && s.bu === 'num'; }, run: () => A().toggleBullets('num') });
  C('bullets', { label: '项目符号', icon: 'bullets', enabled: textCtx, checked: () => { const s = fmtState(); return !!s && s.bu === 'char'; }, run: () => A().toggleBullets('char') });
  C('promote', { label: '减少缩进', icon: 'decIndent', key: 'Alt+Shift+Left', tip: 'Decrease Indent (Promote)', enabled: textCtx, run: () => A().indent(-1) });
  C('demote', { label: '增加缩进', icon: 'incIndent', key: 'Alt+Shift+Right', tip: 'Increase Indent (Demote)', enabled: textCtx, run: () => A().indent(1) });
  C('moveParaUp', { label: '上移', icon: 'moveUp', key: 'Alt+Shift+Up', enabled: textCtx, run: () => A().moveParas(-1) });
  C('moveParaDown', { label: '下移', icon: 'moveDown', key: 'Alt+Shift+Down', enabled: textCtx, run: () => A().moveParas(1) });
  C('fontColor', { label: 'Font Color', icon: 'fontColor', tip: 'Font Color', enabled: textCtx, swatch: () => A().lastFontColorHex(), run: () => A().applyFontColor(A().lastFontColor) });
  C('fillColor', { label: 'Fill Color', icon: 'fillColor', tip: 'Fill Color', enabled: hasSel, swatch: () => A().lastFillHex(), run: () => A().applyFill(A().lastFill) });
  C('lineColor', { label: 'Line Color', icon: 'lineColor', tip: 'Line Color', enabled: hasSel, swatch: () => A().lastLineHex(), run: () => A().applyLineColor(A().lastLineColor) });
  C('painter', { label: '格式刷', icon: 'painter', tip: '格式刷', enabled: () => hasSel() || A().painter.on(), checked: () => A().painter.on(), run: (arg, ev) => A().painter.click(ev) });
  C('showFormatting', { label: '显示格式', icon: 'showFormatting', checked: () => !A().opts.outlinePlain, run: () => { A().opts.outlinePlain = !A().opts.outlinePlain; L.panes.left.render(); } });
  C('expandAll', { label: '展开全部', icon: 'expandAll', run: () => L.panes.left.setTab('outline') });
  C('summarySlide', { label: '摘要幻灯片', icon: 'summary', tip: 'Summary Slide', run: () => A().summarySlide() });

  /* ---------- Tools ---------- */
  C('spelling', { label: '拼写(&S)...', icon: 'spell', key: 'F7', tip: 'Spelling', run: () => A().spelling() });
  C('autocorrectDlg', { label: '自动更正选项(&A)...', run: () => L.dlg.autocorrect() });
  C('optionsDlg', { label: '选项(&O)...', icon: 'options', run: () => L.dlg.options() });

  /* ---------- Slide Show ---------- */
  C('showFromStart', { label: '观看放映(&V)', icon: 'slideshow', key: 'F5', tip: 'Slide Show (from start)', enabled: () => L.pres.slides.length > 0, run: () => L.show.start({ from: 0 }) });
  C('showFromCurrent', { label: '从当前幻灯片开始放映', icon: 'slideshow', key: 'Shift+F5', tip: 'Slide Show from current slide (Shift+F5)', enabled: () => L.pres.slides.length > 0, run: () => L.show.start({ from: A().view === 'master' ? 0 : E().idx }) });
  C('setupShow', { label: '设置放映方式(&S)...', icon: 'setupShow', run: () => L.dlg.setupShow() });
  C('rehearse', { label: '排练计时(&R)', icon: 'rehearse', tip: 'Rehearse Timings', enabled: () => L.pres.slides.length > 0, run: () => L.show.start({ from: 0, rehearse: true }) });
  C('actionSettings', { label: '动作设置(&I)...', icon: 'actionSettings', enabled: hasSel, run: () => L.dlg.actionSettings() });
  C('animSchemes', { label: '动画方案(&N)...', run: () => L.panes.task.show('animSchemes') });
  C('customAnim', { label: '自定义动画(&M)...', icon: 'animation', run: () => L.panes.task.show('customAnim') });
  C('transitionPane', { label: '幻灯片切换(&T)...', icon: 'transition', tbLabel: 'Transition', tip: 'Slide Transition', run: () => L.panes.task.show('transition') });
  C('hideSlide', { label: '隐藏幻灯片(&H)', icon: 'hideSlide', tip: 'Hide Slide', checked: () => { const s = E().slide(); return !!s && !!s.hidden; }, enabled: () => L.pres.slides.length > 0 && A().view !== 'master', run: () => A().toggleHidden() });
  C('speakerNotes', { label: 'Speaker Notes', icon: 'notesView', run: () => A().speakerNotesDlg() });

  /* ---------- Help ---------- */
  C('help', { label: 'Lectern 帮助(&H)', icon: 'help', key: 'F1', tip: 'Lectern Help', run: () => L.panes.task.show('help') });
  C('about', { label: '关于 Lectern(&A)', icon: 'about', run: () => L.dlg.about() });
  C('gettingStarted', { label: '入门(&G)', run: () => L.panes.task.show('getting-started') });

  /* ---------- Drawing ---------- */
  C('select', { label: '选择对象', icon: 'select', checked: () => !E().tool, run: () => { E().cancelFreeform(false); E().setTool(null); } });
  C('drawLine', { label: 'Line', icon: 'line', checked: () => !!(E().tool && E().tool.kind === 'line' && !E().tool.arrows && E().tool.geom === 'line'), run: (a, ev) => E().setTool({ kind: 'line', geom: 'line', sticky: ev && ev.detail === 2 }) });
  C('drawArrow', { label: '箭头', icon: 'arrowTool', checked: () => !!(E().tool && E().tool.kind === 'line' && E().tool.arrows === 'end'), run: (a, ev) => E().setTool({ kind: 'line', geom: 'straightConnector1', arrows: 'end', sticky: ev && ev.detail === 2 }) });
  C('drawDblArrow', { label: '双向箭头', icon: 'doubleArrow', run: () => E().setTool({ kind: 'line', geom: 'straightConnector1', arrows: 'both' }) });
  C('drawRect', { label: '矩形', icon: 'rect', checked: () => !!(E().tool && E().tool.geom === 'rect' && E().tool.kind === 'shape'), run: (a, ev) => E().setTool({ kind: 'shape', geom: 'rect', sticky: ev && ev.detail === 2 }) });
  C('drawOval', { label: '椭圆', icon: 'oval', checked: () => !!(E().tool && E().tool.geom === 'ellipse'), run: (a, ev) => E().setTool({ kind: 'shape', geom: 'ellipse', sticky: ev && ev.detail === 2 }) });
  C('drawCurve', { label: '曲线', icon: 'curve', run: () => E().setTool({ kind: 'curve' }) });
  C('drawFreeform', { label: 'Freeform', icon: 'freeform', run: () => E().setTool({ kind: 'freeform' }) });
  C('drawScribble', { label: '自由曲线', icon: 'scribble', run: () => E().setTool({ kind: 'scribble' }) });
  C('drawElbow', { label: '肘形连接符', icon: 'elbow', run: () => E().setTool({ kind: 'line', geom: 'bentConnector3' }) });
  C('drawCurvedConn', { label: '曲线连接符', icon: 'curve', run: () => E().setTool({ kind: 'line', geom: 'curvedConnector3' }) });
  C('group', { label: '组合(&G)', icon: 'group', key: 'Ctrl+G', enabled: () => E().selected().length > 1, run: () => A().group() });
  C('ungroup', { label: '取消组合(&U)', icon: 'ungroup', key: 'Ctrl+Shift+G', enabled: () => E().selected().some((s) => s.type === 'group'), run: () => A().ungroup() });
  C('regroup', { label: '重新组合(&E)', enabled: () => !!A().lastUngroup, run: () => A().regroup() });
  C('bringFront', { label: '置于顶层(&N)', icon: 'bringFront', enabled: hasSel, run: () => A().order('front') });
  C('sendBack', { label: '置于底层(&K)', icon: 'sendBack', enabled: hasSel, run: () => A().order('back') });
  C('bringForward', { label: '上移一层(&F)', icon: 'bringForward', enabled: hasSel, run: () => A().order('forward') });
  C('sendBackward', { label: '下移一层(&B)', icon: 'sendBackward', enabled: hasSel, run: () => A().order('backward') });
  for (const [k, l, ic] of [['L', '左对齐(&L)', 'alignL'], ['C', '居中对齐(&C)', 'alignC'], ['R', '右对齐(&R)', 'alignR'], ['T', '顶端对齐(&T)', 'alignT'], ['M', '中部对齐(&M)', 'alignM'], ['B', '底端对齐(&B)', 'alignB']]) C('align' + k, { label: l, icon: ic, enabled: hasSel, run: () => A().align(k) });
  C('distH', { label: '水平分布(&H)', icon: 'distH', enabled: () => E().selected().length > 1 || A().opts.alignToSlide, run: () => A().distribute('h') });
  C('distV', { label: '垂直分布(&V)', icon: 'distV', enabled: () => E().selected().length > 1 || A().opts.alignToSlide, run: () => A().distribute('v') });
  C('alignToSlide', { label: '相对于幻灯片(&S)', checked: () => !!A().opts.alignToSlide, run: () => { A().opts.alignToSlide = !A().opts.alignToSlide; } });
  C('freeRotate', { label: '自由旋转(&T)', icon: 'freeRotate', enabled: canRotate, run: () => ui.toast('Drag the green rotation handle to rotate. Hold Shift for 15° steps.') });
  C('rotateLeft', { label: '向左旋转 90°(&L)', icon: 'rotateL', tip: 'Rotate Left 90°', enabled: canRotate, run: () => A().rotate(-90) });
  C('rotateRight', { label: '向右旋转 90°(&R)', icon: 'rotateR', enabled: canRotate, run: () => A().rotate(90) });
  C('flipH', { label: '水平翻转(&H)', icon: 'flipH', enabled: canRotate, run: () => A().flip('h') });
  C('flipV', { label: '垂直翻转(&V)', icon: 'flipV', enabled: canRotate, run: () => A().flip('v') });
  C('nudgeUp', { label: '向上填充(&U)', enabled: hasSel, run: () => A().nudge(0, -1) });
  C('nudgeDown', { label: '向下填充(&D)', enabled: hasSel, run: () => A().nudge(0, 1) });
  C('nudgeLeft', { label: '靠左(&L)', enabled: hasSel, run: () => A().nudge(-1, 0) });
  C('nudgeRight', { label: '靠右(&R)', enabled: hasSel, run: () => A().nudge(1, 0) });
  C('setDefaults', { label: '设置自选图形默认值(&D)', enabled: () => { const s = E().primary(); return !!s && (s.type === 'shape' || s.type === 'line'); }, run: () => A().setDefaults() });
  C('editPoints', { label: '编辑顶点(&N)', enabled: () => false, run: () => {} });

  /* ---------- Picture ---------- */
  C('picMoreContrast', { label: '提高对比度', icon: 'moreContrast', enabled: isPic, run: () => A().picAdjust('contrast', 0.15) });
  C('picLessContrast', { label: '降低对比度', icon: 'lessContrast', enabled: isPic, run: () => A().picAdjust('contrast', -0.15) });
  C('picMoreBright', { label: '提高亮度', icon: 'moreBright', enabled: isPic, run: () => A().picAdjust('bright', 0.1) });
  C('picLessBright', { label: '降低亮度', icon: 'lessBright', enabled: isPic, run: () => A().picAdjust('bright', -0.1) });
  C('picCrop', { label: 'Crop', icon: 'crop', enabled: isPic, checked: () => !!E().cropMode, run: () => { E().cropMode = !E().cropMode; E().drawOverlay(); } });
  C('picReset', { label: '重设图片', icon: 'resetPicture', enabled: isPic, run: () => A().picReset() });
  C('picTransparent', { label: '设置透明色', icon: 'transparentColor', enabled: isPic, run: () => A().picTransparentTool() });

  /* ---------- Table ---------- */
  C('rowAbove', { label: '在上方插入行(&A)', icon: 'rowAbove', enabled: isTable, run: () => A().tableCmd('rowAbove') });
  C('rowBelow', { label: '在下方插入行(&B)', icon: 'rowBelow', enabled: isTable, run: () => A().tableCmd('rowBelow') });
  C('colLeft', { label: '在左侧插入列(&L)', icon: 'colLeft', enabled: isTable, run: () => A().tableCmd('colLeft') });
  C('colRight', { label: '在右侧插入列(&R)', icon: 'colRight', enabled: isTable, run: () => A().tableCmd('colRight') });
  C('deleteRows', { label: '删除行(&R)', icon: 'deleteRow', enabled: isTable, run: () => A().tableCmd('deleteRows') });
  C('deleteCols', { label: '删除列(&C)', icon: 'deleteCol', enabled: isTable, run: () => A().tableCmd('deleteCols') });
  C('mergeCells', { label: '合并单元格(&M)', icon: 'mergeCells', enabled: () => isTable() && !!E().cellSel, run: () => A().tableCmd('merge') });
  C('splitCell', { label: '拆分单元格(&P)', icon: 'splitCells', enabled: isTable, run: () => A().tableCmd('split') });
  C('cellTop', { label: '顶端对齐', icon: 'alignTop', enabled: isTable, run: () => A().tableCmd('anchor', 't') });
  C('cellMiddle', { label: '垂直居中', icon: 'alignMiddle', enabled: isTable, run: () => A().tableCmd('anchor', 'ctr') });
  C('cellBottom', { label: '底端对齐', icon: 'alignBottom', enabled: isTable, run: () => A().tableCmd('anchor', 'b') });
  C('distRows', { label: '平均分布各行', icon: 'distRows', enabled: isTable, run: () => A().tableCmd('distRows') });
  C('distCols', { label: '平均分布各列', icon: 'distCols', enabled: isTable, run: () => A().tableCmd('distCols') });
  C('selectTable', { label: '选择表格(&T)', enabled: isTable, run: () => A().tableCmd('selectAll') });
  C('tableBordersFill', { label: '边框和填充(&F)...', enabled: isTable, run: () => A().tableBordersDlg() });

  /* ---------- WordArt ---------- */
  C('wordartEdit', { label: '编辑文本(&X)...', icon: 'editText', tbLabel: 'Edit Text...', enabled: () => E().selected().length === 1 && isWA(), run: () => A().editWordArt() });
  C('wordartGallery', { label: 'LettersArt Gallery', icon: 'waGallery', enabled: isWA, run: () => A().wordartRestyle() });
  C('waSameHeight', { label: '艺术字字母等高', icon: 'sameHeight', enabled: isWA, checked: () => { const s = E().primary(); return !!(s && s.wa && s.wa.cap); }, run: () => A().waToggle('cap') });
  C('waVertical', { label: '艺术字竖排文字', icon: 'vertText', enabled: isWA, checked: () => { const s = E().primary(); return !!(s && s.wa && s.wa.vert); }, run: () => A().waToggle('vert') });
})();
