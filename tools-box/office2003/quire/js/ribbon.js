/* Quire — the ribbon of the Paper 2016 look (common/ribbon.js): Word's tabs and groups over Quire's commands.
 * Read when the ribbon is first shown, so it may use what app.js defines (A.parts, A.menus). */
(function () {
  'use strict';
  const L = window.L, ui = L.ui;
  const A = () => L.app;
  const open = (items) => (r) => ui.openMenu(items, r);
  const ctx = (k) => () => A().inContext(k);

  L.ribbonSpec = () => ({
    icons: { autocorrectDlg: 'options', clearFormats: 'eraser', goTo: 'forward', insertBreak: 'pageBreak', bookmark: 'bookmark', crossReference: 'crossRef', pageNumbers: 'pageNum', insertField: 'field',
      insertObject: 'object', insertFile: 'insertFile', themeDlg: 'themes', autoFormat: 'tableAutoFormat', printedWatermark: 'watermark', hyphenation: 'hyphenation', tabsDlg: 'tabStops',
      indexTables: 'toc', insertFootnoteNow: 'footnote', insertEndnoteNow: 'endnote', footnotes: 'footnote', insertCaption: 'caption', letterWizard: 'templates', thesaurus: 'thesaurus',
      setLanguage: 'language', deleteAllComments: 'delete', showComments: 'comment', compareDocs: 'compare', newWindow: 'newWindow', arrangeAll: 'arrangeAll', splitWindow: 'split',
      macros: 'macros', headingRows: 'headingRows', splitTable: 'splitTable', tableToText: 'tableToText', formulaDlg: 'fx', wordartEdit: 'editText', hfClose: 'close' },
    qat: ['save', { split: 'undo', menu: (r) => A().historyMenu(r, 'undo') }, { split: 'redo', menu: (r) => A().historyMenu(r, 'redo') }],
    file: () => ['new', 'open', 'close', 'templates', 'fileSearch', '-', 'save', 'saveAs', 'saveWeb', 'webPreview', '-', 'properties', 'permission', 'links', '-', 'pageSetup', 'printPreview', 'print', '-', 'exportPDF', 'sendMail', '-', 'optionsDlg', 'help', 'about', '-', 'exitApp'],
    tabs: [
      { id: 'home', label: '开始(&H)', groups: [
        { label: '剪贴板', launcher: 'officeClipboard', items: [{ split: 'paste', size: 'large', menu: open(['paste', 'pasteSpecial', 'pasteHyperlink']) }, 'cut', 'copy', 'painter'] },
        { label: '字体', icon: 'font', launcher: 'fontDlg', items: [
          { row: [{ make: () => A().parts.font({ width: 128, id: 'rb-font' }) }, { make: () => A().parts.size({ width: 42, id: 'rb-size' }) }, 'growFont', 'shrinkFont', 'changeCase', 'clearFormats'] },
          { row: ['bold', 'italic', 'underline', 'strike', 'subscript', 'superscript', '|', { split: 'highlightBtn', menu: (r, b) => A().menus.highlight(r, b) }, { split: 'fontColorBtn', menu: (r, b) => A().menus.fontColor(r, b) }] }] },
        { label: '段落', icon: 'paragraph', launcher: 'paragraphDlg', items: [
          { row: [{ split: 'bullets', menu: open(['bulletsDlg']) }, 'numbering', 'decIndent', 'incIndent', '|', 'sortDlg', 'showMarks'] },
          { row: ['alignLeft', 'alignCenter', 'alignRight', 'justify', '|', { drop: '行距', icon: 'lineSpacing', menu: (r) => A().menus.lineSpacing(r) }, { split: 'borderBtn', menu: (r, b) => A().menus.borders(r, b) }] }] },
        { label: '样式', launcher: 'stylesPane', items: [{ make: () => A().parts.style({ width: 150, id: 'rb-style' }) }, 'stylesPane', 'revealFormatting'] },
        { label: '编辑', items: ['find', 'replace', { drop: '选择', icon: 'selectAll', menu: open(['selectAll', 'goTo']) }] },
      ] },
      { id: 'insert', label: '插入(&I)', groups: [
        { label: '页', items: [{ cmd: 'insertBreak', size: 'large', label: '分隔符' }] },
        { label: '表格', items: [{ drop: '表格', icon: 'table', size: 'large', menu: (r) => A().menus.tableGrid(r, ['-', 'insertTable', 'drawTable', 'textToTable']) }] },
        { label: '插图', items: [{ cmd: 'insertPicture', size: 'large', label: '图片' }, 'insertClipArt', { drop: '形状', icon: 'autoshapes', menu: (r) => A().openAutoShapes(r) }, 'insertDiagram', 'insertChart'] },
        { label: '链接', items: ['hyperlink', 'bookmark', 'crossReference'] },
        { label: '批注', items: [{ cmd: 'insertComment', size: 'large', label: '批注' }] },
        { label: '页眉和页脚', items: ['headerFooter', 'pageNumbers'] },
        { label: '文本', items: [{ cmd: 'textBox', size: 'large', label: '文本框' }, 'insertWordArt', 'dropCap', 'dateTime', 'insertField', 'insertObject', 'insertFile'] },
        { label: '符号', items: [{ cmd: 'insertSymbol', size: 'large', label: '符号' }] },
      ] },
      { id: 'design', label: '设计(&D)', groups: [
        { label: '文档格式', items: [{ cmd: 'themeDlg', size: 'large', label: '主题' }, 'autoFormat'] },
        { label: '页面背景', items: [{ cmd: 'printedWatermark', size: 'large', label: '水印' }, { cmd: 'bordersDlg', size: 'large', label: '页面边框' }] },
      ] },
      { id: 'layout', label: '布局', groups: [
        { label: '页面设置', launcher: 'pageSetup', items: [{ cmd: 'pageSetup', size: 'large', label: '页边距' }, { drop: '分栏', icon: 'columns', size: 'large', menu: (r) => A().menus.columns(r, ['-', 'columnsDlg']) }, 'insertBreak', 'hyphenation', 'textDirection'] },
        { label: '段落', icon: 'paragraph', launcher: 'paragraphDlg', items: ['incIndent', 'decIndent', 'tabsDlg'] },
        { label: '排列', items: ['bringFront', 'sendBack', { drop: '自动换行', icon: 'textWrap', menu: (r) => A().wrapMenu(r) }, 'groupObj', 'ungroupObj', 'rotateRight'] },
      ] },
      { id: 'references', label: '引用(&r)', groups: [
        { label: '目录', items: [{ cmd: 'indexTables', size: 'large', label: '目录' }, 'updateTOC'] },
        { label: '脚注', launcher: 'insertFootnote', items: [{ cmd: 'insertFootnoteNow', size: 'large', label: '插入脚注' }, 'insertEndnoteNow', 'footnotes'] },
        { label: '题注', items: [{ cmd: 'insertCaption', size: 'large', label: '插入题注' }, 'crossReference'] },
      ] },
      { id: 'mailings', label: '邮件(&g)', groups: [
        { label: '创建', items: [{ cmd: 'envelopes', size: 'large', label: '信封' }, 'letterWizard'] },
        { label: '开始邮件合并', items: [{ cmd: 'mailMerge', size: 'large', label: '开始邮件合并' }, 'mmOpenSource', 'mmRecipients'] },
        { label: '编写和插入域', items: ['mmAddress', 'mmGreeting', 'mmInsertField'] },
        { label: '预览结果', items: [{ cmd: 'mmViewData', size: 'large', label: '预览结果' }, { row: ['mmFirst', 'mmPrev', 'mmNext', 'mmLast'] }] },
        { label: '完成', items: [{ cmd: 'mmMergeNew', size: 'large', label: '完成并合并' }, 'mmMergePrint'] },
      ] },
      { id: 'review', label: '审阅(&v)', groups: [
        { label: '校对', items: ['wordCount', 'autocorrectDlg', 'autoSummarize'] },
        { label: '洞察', items: [{ cmd: 'research', size: 'large', label: '信息检索' }] },
        { label: '语言', items: [{ cmd: 'setLanguage', size: 'large', label: '语言' }] },
        { label: '批注', items: [{ cmd: 'insertComment', size: 'large', label: '新建批注' }, 'deleteAllComments', 'showComments'] },
        { label: '修订跟踪', items: [{ cmd: 'trackChanges', size: 'large', label: '修订' },
          { drop: '显示标记的最终状态', icon: 'reveal', menu: open(['markupFinalMarkup', 'markupFinal', 'markupOriginalMarkup', 'markupOriginal']), tip: '显示以供审阅' },
          { drop: '显示标记', icon: 'comment', menu: open(['markup', '-', 'showComments', 'showInsDel', 'showFormatting', '-', 'reviewersMenu', '-', 'balloonsToggle']) }, 'reviewingPane'] },
        { label: '更改', items: [{ split: 'acceptChange', size: 'large', label: '接受', menu: open(['acceptChange', 'acceptShown', 'acceptAll']) }, { split: 'rejectChange', label: '拒绝', menu: open(['rejectChange', 'rejectShown', 'rejectAll', '-', 'deleteAllComments']) }, 'prevChange', 'nextChange'] },
        { label: '比较', items: [{ cmd: 'compareDocs', size: 'large', label: '比较' }] },
        { label: '保护', items: [{ cmd: 'protectDoc', size: 'large', label: '保护文档' }] },
      ] },
      { id: 'view', label: '视图(&w)', groups: [
        { label: '视图', items: [{ cmd: 'viewReading', size: 'large', label: '阅读模式' }, { cmd: 'viewPrint', size: 'large', label: '页面视图' }, { cmd: 'viewWeb', size: 'large', label: 'Web 版式' }, 'viewOutline', 'viewNormal'] },
        { label: '显示', items: ['ruler', 'docMap', 'thumbnails', 'taskPane'] },
        { label: '缩放', items: [{ cmd: 'zoomDlg', size: 'large', label: '缩放' }, { make: () => A().parts.zoom({ width: 90, id: 'rb-zoom' }) }, 'fullScreen'] },
        { label: '窗口', items: [{ cmd: 'newWindow', size: 'large', label: '新建窗口' }, 'arrangeAll', 'splitWindow'] },
        { label: '宏', items: [{ cmd: 'macros', size: 'large', label: '宏' }] },
      ] },
    ],
    contextual: [
      { id: 'table', set: '表格工具', when: ctx('table'), tabs: [
        { id: 'tableDesign', label: '设计', groups: [
          { label: '表格样式', items: [{ cmd: 'tableAutoFormat', size: 'large', label: '表格样式' }, 'headingRows'] },
          { label: '边框', launcher: 'bordersDlg', items: [{ drop: '边框样式', icon: 'dashStyle', menu: (r) => A().borderStyleMenu(r) }, { drop: '笔粗细', icon: 'lineStyle', menu: (r) => A().borderWeightMenu(r) }, { split: 'borderBtn', size: 'large', label: '边框', menu: (r, b) => A().menus.borders(r, b) }] },
        ] },
        { id: 'tableLayout', label: '布局', groups: [
          { label: '表格', items: [{ drop: '选择', icon: 'selectTable', menu: open(['selectCell', 'selectCol', 'selectRow', 'selectTable']) }, 'gridlines', { cmd: 'tableProps', size: 'large', label: '属性' }] },
          { label: '绘制', items: [{ cmd: 'drawTable', size: 'large', label: '绘制表格' }, { cmd: 'eraser', size: 'large', label: '橡皮擦' }] },
          { label: '行和列', launcher: 'insertCells', items: [{ drop: '删除', icon: 'deleteTable', size: 'large', menu: open(['deleteCells', 'deleteCols', 'deleteRows', 'deleteTable']) }, 'rowAbove', 'rowBelow', 'colLeft', 'colRight'] },
          { label: '合并', items: ['mergeCells', 'splitCells', 'splitTable'] },
          { label: '单元格大小', items: [{ drop: '自动调整', icon: 'autofit', menu: open(['autofitContents', 'autofitWindow', 'fixedWidth']) }, 'distRows', 'distCols'] },
          { label: '对齐方式', items: ['cellTop', 'cellMiddle', 'cellBottom', 'textDirection'] },
          { label: '数据', items: [{ cmd: 'sortDlg', size: 'large', label: '排序' }, 'headingRows', 'tableToText', 'formulaDlg'] },
        ] },
      ] },
      { id: 'picture', set: '图片工具', when: ctx('picture'), tabs: [
        { id: 'pictureFormat', label: '格式', groups: [
          { label: '调整', items: [{ row: ['picMoreBright', 'picLessBright', 'picMoreContrast', 'picLessContrast'] }, { drop: '颜色', icon: 'colorPic', menu: open([{ label: '自动(&A)', run: () => A().picMode('') }, { label: '灰度(&G)', run: () => A().picMode('gray') }, { label: '黑白(&B)', run: () => A().picMode('bw') }, { label: '冲蚀(&W)', run: () => A().picMode('washout') }]) }, 'picReset'] },
          { label: '排列', items: [{ drop: '自动换行', icon: 'textWrap', size: 'large', menu: (r) => A().wrapMenu(r) }, 'bringFront', 'sendBack', 'picRotate'] },
          { label: '大小', launcher: 'formatObject', items: [{ cmd: 'picCrop', size: 'large', label: '裁剪' }] },
        ] },
      ] },
      { id: 'drawing', set: '绘图工具', when: ctx('shape'), tabs: [
        { id: 'drawingFormat', label: '格式', groups: [
          { label: '插入形状', items: [{ drop: '形状', icon: 'autoshapes', size: 'large', menu: (r) => A().openAutoShapes(r) }, 'textBox'] },
          { label: '形状样式', launcher: 'formatObject', items: [{ split: 'fillColor', menu: (r, b) => A().objColorMenu(b, 'fill') }, { split: 'lineColor', menu: (r, b) => A().objColorMenu(b, 'line') }, { drop: '阴影', icon: 'shadowStyle', menu: (r) => A().objShadowMenu(r) }] },
          { label: 'LettersArt 样式', items: [{ cmd: 'wordartGallery', size: 'large', label: 'LettersArt 样式' }, 'wordartEdit'] },
          { label: '排列', items: [{ drop: '自动换行', icon: 'textWrap', menu: (r) => A().wrapMenu(r) }, 'bringFront', 'sendBack', 'groupObj', 'ungroupObj', 'rotateRight'] },
        ] },
      ] },
      { id: 'headerFooter', set: '页眉和页脚工具', open: true, when: ctx('headerFooter'), tabs: [
        { id: 'hfDesign', label: '设计', groups: [
          { label: '页眉和页脚', items: [{ cmd: 'hfPageNum', size: 'large', label: '页码' }, 'hfNumPages', 'hfFormatPageNum'] },
          { label: '插入', items: [{ cmd: 'hfDate', size: 'large', label: '日期' }, 'hfTime'] },
          { label: '导航', items: [{ cmd: 'hfSwitch', size: 'large', label: '转到页脚' }, 'hfPrev', 'hfNext', 'hfLinkPrev'] },
          { label: '选项', items: ['hfShowText', 'pageSetup'] },
          { label: '关闭', items: [{ cmd: 'hfClose', size: 'large', label: '关闭页眉和页脚' }] },
        ] },
      ] },
      { id: 'outlining', set: L.N('大纲'), open: true, when: ctx('outlining'), tabs: [
        { id: 'outlining', label: '大纲', groups: [
          { label: '大纲工具', items: [{ row: ['promoteH1', 'promote', 'demote', 'demoteBody'] }, { row: ['moveUp', 'moveDown', 'expand', 'collapse'] }, 'olFirstLine', 'olFormatting'] },
          { label: '目录', items: ['updateTOC', 'gotoTOC'] },
          { label: '关闭', items: [{ cmd: 'viewPrint', size: 'large', label: '关闭大纲视图' }] },
        ] },
      ] },
    ],
  });
})();
