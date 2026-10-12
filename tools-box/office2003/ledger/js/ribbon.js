/* Ledger — the ribbon of the Paper 2016 look (common/ribbon.js): Excel's tabs and groups over Ledger's commands.
 * Read when the ribbon is first shown, so it may use what app.js defines (A.parts, A.menus). */
(function () {
  'use strict';
  const L = window.L, ui = L.ui;
  const A = () => L.app;
  const open = (items) => (r) => ui.openMenu(items, r);
  const ctx = (k) => () => A().inContext(k);

  L.ribbonSpec = () => ({
    icons: { autoCorrectDlg: 'options', taskPane: 'taskpane', pasteValues: 'paste', fillSeries: 'fillColor', clearAll: 'eraser', goTo: 'forward', insertSheet: 'insertSheet', moveCopySheet: 'insertSheet', nameDefine: 'nameBox', pasteName: 'nameBox',
      nameCreate: 'nameBox', nameApply: 'nameBox', functionsHelp: 'help', calcSheet: 'calcNow', errorChecking: 'traceError', showAll: 'filter', advancedFilter: 'filterOn', consolidate: 'subtotal',
      groupRows: 'group', ungroupRows: 'ungroup', showDetail: 'expand', hideDetail: 'collapse', autoOutline: 'group', clearOutline: 'ungroup', importText: 'textToCols', refreshData: 'redo',
      protectWorkbook: 'protect', allowEditRanges: 'protect', deleteComment: 'delete', showHideComment: 'comment', showComments: 'comment', viewNormal: 'normalGrid', newWindow: 'newWindow',
      arrangeWindows: 'arrangeAll', macros: 'macros', formulaBar: 'fx', statusBar: 'layout', gridlinesView: 'grid', headingsView: 'headingRows', setPrintArea: 'pageSetup', insertPageBreak: 'pageBreak',
      resetPageBreaks: 'pageBreak', sheetBackground: 'background', createList: 'table', totalRow: 'autoSum', convertToRange: 'tableToText', newSheetChart: 'chart', exportCSV: 'excel', conditionalFormat: 'condFormat', styleDlg: 'stylesPane', headerFooter: 'headerFooter', formatObject: 'formatPicture' },
    qat: ['save', 'undo', 'redo'],
    file: () => ['newBook', 'open', 'close', '-', 'save', 'saveAs', 'saveWeb', '-', 'properties', 'links', '-', 'pageSetup', 'printPreview', 'print', '-', 'exportPDF', 'exportCSV', '-', 'options', 'addIns', 'help', 'about', '-', 'exitApp'],
    tabs: [
      { id: 'home', label: '开始(&H)', groups: [
        { label: '剪贴板', launcher: 'officeClipboard', items: [{ split: 'paste', size: 'large', menu: open(['paste', 'pasteValues', 'pasteSpecial', 'pasteHyperlink']) }, 'cut', 'copy', 'formatPainter'] },
        { label: '字体', icon: 'font', launcher: 'formatCells', items: [
          { row: [{ make: () => A().parts.font({ width: 132, id: 'rb-font' }) }, { make: () => A().parts.size({ width: 44, id: 'rb-size' }) }] },
          { row: ['bold', '倾斜', 'underline', 'strike', '|', { split: 'bordersApply', menu: (r, b) => A().menus.borders(r, b) }, { split: 'fillColorApply', menu: (r, b) => A().menus.fill(r, b) }, { split: 'fontColorApply', menu: (r, b) => A().menus.fontColor(r, b) }] }] },
        { label: '对齐', launcher: 'formatCells', items: [
          { row: ['alignLeft', 'alignCenter', 'alignRight', '|', 'decIndent', 'incIndent'] },
          { row: [{ cmd: 'wrapText', size: 'medium', label: '文字环绕' }, { cmd: 'mergeCenter', size: 'medium', label: '合并及居中' }] }] },
        { label: '数字', launcher: 'formatCells', items: [
          { row: [{ drop: 'Number Format', icon: 'formatCells', size: 'medium', menu: open(['fmtGeneral', 'fmtNumber', 'fmtCurrency', 'fmtDate', 'fmtTime', 'fmtExp', '-', 'formatCells']) }] },
          { row: ['currencyStyle', 'percentStyle', 'commaStyle', '|', 'incDecimal', 'decDecimal'] }] },
        { label: '样式', items: [{ cmd: 'conditionalFormat', size: 'large', label: 'Conditional Formatting' }, { cmd: 'autoFormat', size: 'large', label: '套用表格格式' }, { cmd: 'styleDlg', size: 'large', label: '单元格样式' }] },
        { label: '单元格', items: [
          { drop: '插入', icon: 'insertCells', size: 'large', menu: open(['insertCells', 'insertRows', 'insertCols', 'insertSheet']) },
          { drop: '删除', icon: 'deleteSheet', size: 'large', menu: open(['deleteCells', 'deleteSheet']) },
          { drop: '格式', icon: 'formatCells', size: 'large', menu: open([{ label: '行', sub: ['rowHeight', 'rowAutofit', 'rowHide', 'rowUnhide'] }, { label: 'Column', sub: ['colWidth', 'colAutofit', 'colHide', 'colUnhide', 'colStandard'] }, { label: '工作表', sub: ['sheetRename', 'sheetHide', 'sheetUnhide', 'moveCopySheet', 'tabColor'] }, '-', 'protectSheet', 'formatCells']) }] },
        { label: '编辑', items: [
          { split: 'autoSum', label: 'AutoSum', menu: (r) => A().menus.autoSum(r) },
          { drop: 'Fill', icon: 'fillColor', menu: open(['fillDown', 'fillRight', 'fillUp', 'fillLeft', 'fillAcross', 'fillSeries', 'fillJustify']) },
          { drop: 'Clear', icon: 'eraser', menu: open(['clearAll', 'clearFormats', 'clearContents', 'clearComments', 'clearHyperlinks']) },
          { drop: 'Sort & Filter', icon: 'sortAsc', size: 'large', menu: open(['sortAsc', 'sortDesc', 'sortDlg', '-', 'autoFilter', 'showAll', 'advancedFilter']) },
          { drop: 'Find & Select', icon: 'find', size: 'large', menu: open(['find', 'replace', 'goTo', 'goToSpecial']) }] },
      ] },
      { id: 'insert', label: '插入(&I)', groups: [
        { label: '表格', items: [{ cmd: 'pivotTable', size: 'large', label: 'PivotTable' }, { cmd: 'createList', size: 'large', label: '表格' }] },
        { label: '插图', items: [{ cmd: 'insertPicture', size: 'large', label: '图片' }, 'insertClipArt', { cmd: 'insertShape', label: '形状', icon: 'autoshapes' }] },
        { label: '图表', items: [{ cmd: 'chartWizard', size: 'large', label: 'Chart' }, 'insertChart', 'newSheetChart'] },
        { label: '链接', items: [{ cmd: 'hyperlink', size: 'large', label: 'Hyperlink' }] },
        { label: '文字', items: [{ cmd: 'insertTextBox', size: 'large', label: '文本框' }, 'headerFooter'] },
        { label: '符号', items: [{ cmd: 'insertSymbol', size: 'large', label: '符号' }] },
      ] },
      { id: 'pageLayout', label: '页面布局', groups: [
        { label: '页面设置', launcher: 'pageSetup', items: [{ cmd: 'pageSetup', size: 'large', label: '页边距' }, { drop: 'Print Area', icon: 'pageSetup', size: 'large', menu: open(['setPrintArea', 'clearPrintArea']) },
          { drop: 'Breaks', icon: 'pageBreak', size: 'large', menu: open(['insertPageBreak', 'resetPageBreaks']) }, { cmd: 'sheetBackground', size: 'large', label: 'Background' }] },
        { label: '工作表选项', items: [{ cmd: 'gridlinesView', label: '网格线' }, { cmd: 'headingsView', label: 'Headings' }] },
        { label: '排列', items: [{ drop: 'Bring Forward / Send Backward', icon: 'bringFront', size: 'large', label: '排列', menu: (r) => A().menus.draw(r) }] },
      ] },
      { id: 'formulas', label: '公式', groups: [
        { label: '函数库', items: [{ cmd: 'insertFunction', size: 'large', label: 'Insert Function' }, { split: 'autoSum', size: 'large', label: 'AutoSum', menu: (r) => A().menus.autoSum(r) }, 'functionsHelp'] },
        { label: '定义的名称', items: [{ cmd: 'nameDefine', size: 'large', label: 'Define Name' }, 'pasteName', 'nameCreate', 'nameApply'] },
        { label: '公式审核', items: ['tracePrecedents', 'traceDependents', 'removeArrows', 'showFormulas', 'traceError', 'evaluateFormula', { cmd: 'watchWindow', size: 'large', label: 'Watch Window' }] },
        { label: '重新计算', items: [{ cmd: 'calcNow', size: 'large', label: '立即计算' }, 'calcSheet'] },
      ] },
      { id: 'data', label: '&Data', groups: [
        { label: '获取外部数据', items: [{ cmd: 'importText', size: 'large', label: '来自文本' }, { cmd: 'refreshData', size: 'large', label: '全部刷新' }] },
        { label: '排序和筛选', items: [{ row: ['sortAsc', 'sortDesc'] }, { cmd: 'sortDlg', size: 'large', label: '排序' }, { cmd: 'autoFilter', size: 'large', label: '筛选' }, 'showAll', 'advancedFilter'] },
        { label: '数据工具', items: [{ cmd: 'textToColumns', size: 'large', label: '分列' }, { split: 'validation', menu: open(['validation', 'circleInvalid', 'clearCircles']) }, 'consolidate', 'dataForm'] },
        { label: '预测', items: [{ drop: 'What-If Analysis', icon: 'goalSeek', size: 'large', menu: open(['scenarios', 'goalSeek', 'dataTable']) }] },
        { label: '大纲', launcher: 'outlineSettings', items: [{ cmd: 'groupRows', size: 'large', label: 'Group' }, { split: 'ungroupRows', size: 'large', label: '取消组合', menu: open(['ungroupRows', 'clearOutline']) }, { cmd: 'subtotals', size: 'large', label: 'Subtotal' }, 'showDetail', 'hideDetail', 'autoOutline'] },
      ] },
      { id: 'review', label: '审阅(&V)', groups: [
        { label: '校对', items: ['errorChecking', 'autoCorrectDlg'] },
        { label: '批注', items: [{ cmd: 'insertComment', size: 'large', label: '新建批注' }, 'deleteComment', 'showHideComment', 'showComments'] },
        { label: '修订', items: [{ cmd: 'protectSheet', size: 'large', label: 'Protect Sheet' }, { cmd: 'protectWorkbook', size: 'large', label: 'Protect Workbook' }, 'allowEditRanges'] },
      ] },
      { id: 'view', label: '视图(&W)', groups: [
        { label: '工作簿视图', items: [{ cmd: 'viewNormal', size: 'large', label: 'Normal' }, { cmd: 'viewPageBreak', size: 'large', label: '分页预览' }, 'fullScreen'] },
        { label: '显示', items: ['formulaBar', { cmd: 'gridlinesView', label: '网格线' }, { cmd: 'headingsView', label: 'Headings' }, 'statusBar', 'taskPane'] },
        { label: '显示比例', items: [{ cmd: 'zoomDlg', size: 'large', label: '显示比例' }, { make: () => A().parts.zoom({ width: 70, id: 'rb-zoom' }) }] },
        { label: '窗口', items: [{ cmd: 'newWindow', size: 'large', label: '新建窗口' }, 'arrangeWindows', { cmd: 'freezePanes', size: 'large', label: '冻结窗格' }, 'splitWindow'] },
        { label: '宏', items: [{ cmd: 'macros', size: 'large', label: '宏' }] },
      ] },
    ],
    contextual: [
      { id: 'chart', set: 'Chart Tools', when: ctx('chart'), tabs: [
        { id: 'chartDesign', label: '设计', groups: [
          { label: 'Type', items: [{ drop: 'Change Chart Type', icon: 'chart', size: 'large', menu: (r) => L.chartWizard && L.chartWizard.typeMenu(r) }] },
          { label: '数据', items: [{ cmd: 'chartWizard', size: 'large', label: '图表向导' }] },
        ] },
        { id: 'chartFormat', label: '格式', groups: [{ label: '当前选定区域', items: [{ cmd: 'formatObject', size: 'large', label: '设置所选内容格式' }] }, { label: '排列', items: [{ drop: '排列', icon: 'bringFront', size: 'large', menu: (r) => A().menus.draw(r) }] }] },
      ] },
      { id: 'picture', set: '图片工具', when: ctx('picture'), tabs: [
        { id: 'pictureFormat', label: '格式', groups: [{ label: 'Picture', items: [{ cmd: 'formatObject', size: 'large', label: 'Format Picture' }] }, { label: '排列', items: [{ drop: '排列', icon: 'bringFront', size: 'large', menu: (r) => A().menus.draw(r) }] }] },
      ] },
      { id: 'drawing', set: '绘图工具', when: ctx('shape'), tabs: [
        { id: 'drawingFormat', label: '格式', groups: [
          { label: '插入形状', items: [{ cmd: 'insertShape', size: 'large', label: '形状', icon: 'autoshapes' }, 'insertTextBox'] },
          { label: '形状样式', launcher: 'formatObject', items: [{ cmd: 'fillColorApply', size: 'large', label: '形状填充' }] },
          { label: '排列', items: [{ drop: '排列', icon: 'bringFront', size: 'large', menu: (r) => A().menus.draw(r) }] },
        ] },
      ] },
      { id: 'list', set: '表格工具', when: ctx('list'), tabs: [
        { id: 'listDesign', label: '设计', groups: [
          { label: 'Tools', items: [{ cmd: 'convertToRange', size: 'large', label: '转换为区域' }, 'sortDlg', 'autoFilter'] },
          { label: '表格样式选项', items: [{ cmd: 'totalRow', label: '汇总行' }, 'insertRows', 'insertCols'] },
        ] },
      ] },
    ],
  });
})();
