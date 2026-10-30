/* Lectern — the ribbon of the Paper 2016 look (common/ribbon.js): PowerPoint's tabs and groups over Lectern's
 * commands. Read when the ribbon is first shown, so it may use what app.js defines (A.parts, A.menus). */
(function () {
  'use strict';
  const L = window.L, ui = L.ui;
  const A = () => L.app;
  const open = (items) => (r) => ui.openMenu(items, r);
  const ctx = (k) => () => A().inContext(k);

  L.ribbonSpec = () => ({
    icons: { clear: 'delete', duplicate: 'duplicate', deleteSlide: 'delete', vTextBox: 'vtextbox', slidesFromFiles: 'insertSlides', slidesFromOutline: 'outlineTxt', replaceFonts: 'font',
      viewNotes: 'notesView', viewColor: 'grayscale', viewBW: 'grayscale', zoomFit: 'zoom', closeMaster: 'close', insertTitleMaster: 'master', deleteTitleMaster: 'delete', speakerNotes: 'notesView',
      selectTable: 'selectTable', tableBordersFill: 'tablesBorders', regroup: 'group', alignToSlide: 'alignC', editPoints: 'freeform', summarySlide: 'summary', slidesPane: 'thumbnails',
      gridGuides: 'grid', autocorrectDlg: 'options', duplicateSlide: 'duplicate', subscript: 'subscript', superscript: 'superscript', headerFooter: 'headerFooter',
      pageSetup: 'pageSetup', animSchemes: 'animation' },
    qat: ['save', 'undo', 'redo', 'showFromStart'],
    file: () => ['new', 'open', 'close', '-', 'save', 'saveAs', 'saveWeb', '-', 'properties', ...(L.file.native() ? ['versions'] : []), '-', 'pageSetup', 'printPreview', 'print', '-', 'optionsDlg', 'help', 'about', '-', 'exitApp'],
    tabs: [
      { id: 'home', label: '开始(&H)', groups: [
        { label: '剪贴板', launcher: 'officeClipboard', items: [{ split: 'paste', size: 'large', menu: open(['paste', 'pasteSpecial']) }, 'cut', { split: 'copy', menu: open(['copy', 'duplicate']) }, 'painter'] },
        { label: '幻灯片', items: [{ split: 'newSlide', size: 'large', label: '新建幻灯片', menu: open(['newSlide', 'duplicateSlide', '-', 'slidesFromFiles', 'slidesFromOutline']) }, { cmd: 'slideLayout', label: '布局' }, 'duplicateSlide', 'deleteSlide'] },
        { label: '字体', icon: 'font', launcher: 'fontDlg', items: [
          { row: [{ make: () => A().parts.font({ width: 130, id: 'rb-font' }) }, { make: () => A().parts.size({ width: 44, id: 'rb-size' }) }, 'growFont', 'shrinkFont', 'changeCase'] },
          { row: ['bold', 'italic', 'underline', 'shadowText', 'subscript', 'superscript', '|', { split: 'fontColor', menu: (r, b) => A().menus.color('font')(r, b) }] }] },
        { label: '段落', icon: 'paragraph', launcher: 'bulletsDlg', items: [
          { row: ['bullets', 'numbering', '|', 'promote', 'demote', '|', 'lineSpacing'] },
          { row: ['alignLeft', 'alignCenter', 'alignRight', 'justify'] }] },
        { label: '绘图', launcher: 'formatObject', items: [
          { drop: '形状', icon: 'autoshapes', size: 'large', menu: (r) => A().menus.autoShapes(r) },
          { drop: '排列', icon: 'bringFront', size: 'large', menu: (r) => A().menus.draw(r) },
          { split: 'fillColor', label: '形状填充', menu: (r, b) => A().menus.color('fill')(r, b) },
          { split: 'lineColor', label: '形状轮廓', menu: (r, b) => A().menus.color('line')(r, b) },
          { drop: '形状效果', icon: 'shadowStyle', menu: (r) => A().menus.shadow(r) }] },
        { label: '编辑', items: ['find', 'replace', 'selectAll'] },
      ] },
      { id: 'insert', label: '插入(&I)', groups: [
        { label: '幻灯片', items: [{ split: 'newSlide', size: 'large', label: '新建幻灯片', menu: open(['newSlide', 'duplicateSlide', '-', 'slidesFromFiles', 'slidesFromOutline']) }] },
        { label: '表格', items: [{ cmd: 'insertTable', size: 'large', label: '表格' }] },
        { label: '图像', items: [{ cmd: 'insertPicture', size: 'large', label: '图片' }, 'insertClipArt', 'photoAlbum'] },
        { label: '插图', items: [{ drop: '形状', icon: 'autoshapes', size: 'large', menu: (r) => A().menus.autoShapes(r) }, { cmd: 'insertDiagram', size: 'large', label: 'IntelliArt' }, { cmd: 'insertChart', size: 'large', label: '图表' }] },
        { label: '链接', items: [{ cmd: 'hyperlink', size: 'large', label: '超链接' }, { cmd: 'actionSettings', size: 'large', label: '动作' }] },
        { label: '文本', items: [{ cmd: 'textBox', size: 'large', label: '文本框' }, { cmd: 'headerFooter', size: 'large', label: '页眉和页脚' }, { cmd: 'insertWordArt', size: 'large', label: 'LettersArt' }, 'insertDateTime', 'insertSlideNumber', 'vTextBox'] },
        { label: '符号', items: [{ cmd: 'insertSymbol', size: 'large', label: '符号' }] },
        { label: '媒体', items: [{ cmd: 'insertMedia', size: 'large', label: '媒体' }] },
      ] },
      { id: 'design', label: '设计(&g)', groups: [
        { label: '主题', items: [{ cmd: 'slideDesign', size: 'large', label: '主题' }, 'replaceFonts'] },
        { label: '自定义', items: [{ cmd: 'pageSetup', size: 'large', label: '幻灯片大小' }, { cmd: 'background', size: 'large', label: '设置背景格式' }] },
      ] },
      { id: 'transitions', label: '切换', groups: [
        { label: '切换到此幻灯片', items: [{ cmd: 'transitionPane', size: 'large', label: '切换' }] },
        { label: '计时', items: [{ cmd: 'rehearse', size: 'large', label: '排练计时' }] },
      ] },
      { id: 'animations', label: '动画', groups: [
        { label: '动画', items: [{ cmd: 'animSchemes', size: 'large', label: '动画方案' }] },
        { label: '高级动画', items: [{ cmd: 'customAnim', size: 'large', label: '动画窗格' }] },
      ] },
      { id: 'slideShow', label: '幻灯片放映(&w)', groups: [
        { label: '开始幻灯片放映', items: [{ cmd: 'showFromStart', size: 'large', label: '从头开始' }, { cmd: 'showFromCurrent', size: 'large', label: '从当前幻灯片开始' }] },
        { label: '设置', items: [{ cmd: 'setupShow', size: 'large', label: '设置幻灯片放映' }, 'hideSlide', 'rehearse'] },
      ] },
      { id: 'review', label: '审阅(&v)', groups: [
        { label: '校对', items: [{ cmd: 'spelling', size: 'large', label: '拼写' }, 'autocorrectDlg'] },
      ] },
      { id: 'view', label: '视图', groups: [
        { label: '演示文稿视图', items: [{ cmd: 'viewNormal', size: 'large', label: '普通' }, { cmd: 'viewSorter', size: 'large', label: '幻灯片浏览' }, { cmd: 'viewNotes', size: 'large', label: '备注页' }] },
        { label: '母版视图', items: [{ cmd: 'viewMaster', size: 'large', label: '幻灯片母版' }] },
        { label: '显示', items: ['ruler', 'showGrid', 'gridGuides', { cmd: 'speakerNotes', label: '备注' }, 'slidesPane', 'taskPane'] },
        { label: '缩放', items: [{ cmd: 'zoomDlg', size: 'large', label: '缩放' }, { cmd: 'zoomFit', size: 'large', label: '适应窗口' }, { make: () => A().parts.zoom({ width: 64, id: 'rb-zoom' }) }] },
        { label: '颜色/灰度', items: ['viewColor', 'viewGray', 'viewBW'] },
      ] },
    ],
    contextual: [
      { id: 'master', set: '幻灯片母版', open: true, when: ctx('master'), tabs: [
        { id: 'masterTab', label: '幻灯片母版', groups: [
          { label: '编辑母版', items: [{ cmd: 'insertTitleMaster', size: 'large', label: '插入标题母版' }, 'deleteTitleMaster'] },
          { label: '关闭', items: [{ cmd: 'closeMaster', size: 'large', label: '关闭母版视图' }] },
        ] },
      ] },
      { id: 'smartart', set: 'IntelliArt 工具', when: ctx('diagram'), tabs: [
        { id: 'smartartDesign', label: '设计', groups: [
          { label: '创建图形', items: [{ split: 'diagramAddShape', size: 'large', label: '添加形状', menu: open(['diagramAddAfter', 'diagramAddBefore', 'diagramAddAbove', 'diagramAddBelow']) },
            { row: ['diagramPromote', 'diagramDemote'] }, { row: ['diagramMoveUp', 'diagramMoveDown'] }, { cmd: 'diagramTextPane', size: 'large' }] },
          { label: '布局', items: [{ drop: '更改布局', icon: 'layout', size: 'large', menu: (r) => L.diagram.layoutMenu(r) }] },
          { label: 'IntelliArt 样式', items: [{ drop: '更改颜色', icon: 'colorPic', size: 'large', menu: (r) => L.diagram.colorsMenu(r) }, { drop: '样式', icon: 'themes', size: 'large', menu: (r) => L.diagram.styleMenu(r) }] },
          { label: '重置', items: [{ cmd: 'diagramReset', size: 'large' }, 'diagramToShapes'] },
        ] },
        { id: 'smartartFormat', label: '格式', groups: [
          { label: '形状样式', launcher: 'formatObject', items: [{ split: 'fillColor', label: '形状填充', menu: (r, b) => A().menus.color('fill')(r, b) }, { split: 'lineColor', label: '形状轮廓', menu: (r, b) => A().menus.color('line')(r, b) }, { drop: '阴影', icon: 'shadowStyle', menu: (r) => A().menus.shadow(r) }] },
          { label: '排列', items: ['bringFront', 'sendBack', { drop: '对齐', icon: 'alignC', menu: open(['distH', 'distV', 'alignToSlide']) }] },
        ] },
      ] },
      { id: 'drawing', set: '绘图工具', when: ctx('shape'), tabs: [
        { id: 'drawingFormat', label: '格式', groups: [
          { label: '插入形状', items: [{ drop: '形状', icon: 'autoshapes', size: 'large', menu: (r) => A().menus.autoShapes(r) }, 'textBox', 'editPoints'] },
          { label: '形状样式', launcher: 'formatObject', items: [{ split: 'fillColor', label: '形状填充', menu: (r, b) => A().menus.color('fill')(r, b) }, { split: 'lineColor', label: '形状轮廓', menu: (r, b) => A().menus.color('line')(r, b) },
            { drop: '线型', icon: 'lineStyle', menu: (r) => A().menus.lineStyle(r) }, { drop: '虚线线型', icon: 'dashStyle', menu: (r) => A().menus.dash(r) }, { drop: '箭头样式', icon: 'arrowStyle', menu: (r) => A().menus.arrow(r) }, { drop: '阴影样式', icon: 'shadowStyle', menu: (r) => A().menus.shadow(r) }] },
          { label: '排列', items: ['bringFront', 'sendBack', 'group', 'ungroup', { drop: '对齐', icon: 'alignC', menu: open(['distH', 'distV', 'alignToSlide']) }, 'rotateRight', 'flipH', 'flipV'] },
        ] },
      ] },
      { id: 'picture', set: '图片工具', when: ctx('picture'), tabs: [
        { id: 'pictureFormat', label: '格式', groups: [
          { label: '调整', items: [{ row: ['picMoreBright', 'picLessBright', 'picMoreContrast', 'picLessContrast'] }, { drop: '颜色', icon: 'colorPic', menu: (r) => A().menus.picColor(r) }, 'picTransparent', 'picReset'] },
          { label: '图片样式', launcher: 'formatObject', items: [{ drop: '图片边框', icon: 'lineStyle', menu: (r) => A().menus.lineStyle(r) }] },
          { label: '排列', items: ['bringFront', 'sendBack', 'rotateLeft'] },
          { label: '大小', items: [{ cmd: 'picCrop', size: 'large', label: '裁剪' }] },
        ] },
      ] },
      { id: 'table', set: '表格工具', when: ctx('table'), tabs: [
        { id: 'tableDesign', label: '设计', groups: [
          { label: '表格样式', launcher: 'tableBordersFill', items: [{ drop: '底纹', icon: 'fillColor', size: 'large', menu: (r, b) => A().menus.tableFill(r, b) }, { drop: '边框', icon: 'bordersAll', size: 'large', menu: (r) => A().menus.tableBorders(r) }] },
        ] },
        { id: 'tableLayout', label: '布局', groups: [
          { label: '表格', items: [{ cmd: 'selectTable', size: 'large', label: '选择' }] },
          { label: '行和列', items: [{ drop: '删除', icon: 'deleteTable', size: 'large', menu: open(['deleteRows', 'deleteCols']) }, 'rowAbove', 'rowBelow', 'colLeft', 'colRight'] },
          { label: '合并', items: ['mergeCells', 'splitCell'] },
          { label: '单元格大小', items: ['distRows', 'distCols'] },
          { label: '对齐方式', items: ['cellTop', 'cellMiddle', 'cellBottom'] },
        ] },
      ] },
      { id: 'wordart', set: 'LettersArt 工具', when: ctx('wordart'), tabs: [
        { id: 'wordartFormat', label: '格式', groups: [
          { label: '文本', items: [{ cmd: 'wordartEdit', size: 'large', label: '编辑文字' }, 'wordartGallery'] },
          { label: 'LettersArt 样式', launcher: 'formatObject', items: [{ drop: 'LettersArt 形状', icon: 'waShape', size: 'large', menu: (r) => A().menus.waShape(r) }, 'waSameHeight', 'waVertical'] },
        ] },
      ] },
    ],
  });
})();
