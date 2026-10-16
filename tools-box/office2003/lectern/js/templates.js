/* Lectern — content templates: ready-made decks for everyday talks, each on one of the design
 * templates (model.js). Used by the Start Center (lectern/?template=<id>); the text on every slide
 * says what to put there.
 */
(function () {
  'use strict';
  const L = window.L, M = L.model, T = L.txt;
  const LIST = [
    { id: 'status', name: '项目状态报告', design: 'blueprint' },
    { id: 'meeting', name: '团队会议', design: 'meadow' },
    { id: 'lesson', name: '课程', design: 'parchment' },
    { id: 'event', name: '活动公告', design: 'night' },
  ];
  const TL = (L.templates = {});
  TL.list = () => LIST;

  /** a table in place of the slide's table placeholder; rows[0] is the header */
  function table(pres, s, tp, rows, widths) {
    const t = M.newTable(pres, s, rows.length, rows[0].length, tp.x, tp.y, tp.w);
    t.ph = L.clone(tp.ph);
    const rh = Math.min(44, tp.h / rows.length);
    rows.forEach((row, ri) => row.forEach((v, ci) => {
      const cell = t.tbl.rows[ri].cells[ci];
      cell.tx.ps = [T.para(v, { algn: 'l' }, ri === 0 ? { sz: 18, b: true, color: 'lt1' } : { sz: 16 })];
      cell.fill = ri === 0 ? { t: 'solid', c: 'accent1', a: 1 } : { t: 'solid', c: ri % 2 ? '#F2F5FA' : '#FFFFFF', a: 1 };
      for (const k of ['l', 'r', 't', 'b']) cell.bd[k] = { c: '#B4C3DC', w: 1, dash: 'solid' };
      cell.tx.anchor = 'ctr';
    }));
    t.tbl.rows.forEach((r) => { r.h = rh; });
    t.h = rh * rows.length;
    t.tbl.cols = widths.map((f) => tp.w * f);
    s.shapes.splice(s.shapes.indexOf(tp), 1, t);
    return t;
  }

  /** a SmartArt diagram in place of the slide's body placeholder (js/diagram.js) */
  const diagram = (pres, s, layout, lines, opts) => L.diagram.make(pres, s, s.shapes.find((x) => x.ph && x.ph.type === 'body'), layout, lines, opts);

  const B = {
    status(add, pres) {
      add('title', '项目名称', ['状态报告 · 报告周期', '汇报人：您的姓名']);
      add('text', '摘要', ['总体状态：正常 / 有风险 / 已偏离', '一句话说明自上次报告以来的进展', '一句话说明需要关注的问题']);
      diagram(pres, add('text', '项目阶段'), 'chevron', ['调研', '设计', '构建', '上线']);
      const s = add('table', '里程碑');
      table(pres, s, s.shapes.find((x) => x.ph && x.ph.type === 'tbl'), [['里程碑', '负责人', '截止日期', '状态'], ['需求确认完毕', '姓名', '日期', '已完成'], ['设计完成', '姓名', '日期', '正常'], ['首个版本发布', '姓名', '日期', '有风险'], ['正式上线', '姓名', '日期', '未开始']], [0.42, 0.2, 0.16, 0.22]);
      add('twoText', '已完成与下一步', ['本期已完成', '\t完成了什么', '\t交付了什么'], ['下期计划', '\t接下来做什么', '\t由谁负责']);
      add('text', '风险与问题', ['风险：可能出什么问题', '\t影响以及我们的应对措施', '问题：已经出了什么问题', '\t它阻塞了什么，由谁解决']);
      add('text', '需要决策的事项', ['需要与会者给出什么', '截止时间，以及延误的后果']);
    },
    meeting(add, pres) {
      add('title', '团队会议', ['日期 · 时间 · 会议室或视频链接']);
      diagram(pres, add('text', '议程'), 'vProcess', ['欢迎与近况交流', '上周进展', '讨论议题', '行动事项与后续步骤']);
      add('twoText', '进展', ['进展顺利', '\t哪些做得好', '\t要感谢谁'], ['遇到的阻碍', '\t卡在哪里', '\t需要什么帮助']);
      add('text', '讨论', ['今天要决定的问题', '\t方案 A', '\t方案 B']);
      add('text', '行动事项', ['行动 — 负责人 — 截止日期', '行动 — 负责人 — 截止日期', '行动 — 负责人 — 截止日期']);
      add('title', '谢谢', ['下次会议：日期与时间']);
    },
    lesson(add, pres) {
      add('title', '课程标题', ['课程名称 · 教师姓名']);
      add('text', '今天的学习目标', ['了解……', '练习如何……', '能够解释……']);
      diagram(pres, add('text', '核心概念'), 'radial', ['概念', '\t示例', '\t为什么重要', '\t反例', '\t应用场景']);
      add('twoText', '比较', ['第一项', '\t特点', '\t特点'], ['第二项', '\t特点', '\t特点']);
      add('text', '轮到你了', ['给全班的问题或练习', '有多少时间，两人一组还是独立完成']);
      add('text', '小结与作业', ['要记住的三件事', '作业：做什么，何时交']);
    },
    event(add, pres) {
      add('title', '活动名称', ['星期、日期 · 时间', '地点']);
      add('text', '活动看点', ['活动的亮点', '谁来讲演或表演', '餐饮与活动']);
      diagram(pres, add('text', '活动流程'), 'timeline', ['时间 — 开场', '时间 — 主体环节', '时间 — 休息', '时间 — 结束']);
      add('text', '到达方式', ['地址', '公共交通与停车', '无障碍设施']);
      add('title', '期待您的到来', ['请于某日期前回复 · 联系方式']);
    },
  };

  /** a new presentation from the template; size: { w, h } (default M.DEFAULT_SIZE, 16:9) */
  TL.build = function (id, size) {
    const t = LIST.find((x) => x.id === id);
    if (!t) return null;
    const pres = M.newPresentation(Object.assign({ design: t.design, empty: true }, size || {}));
    const did = Object.keys(pres.designs)[0];
    /** add(layout, title, body lines[, second column lines]): '\t' indents a line one level */
    const add = (layout, title, ...cols) => {
      const s = M.newSlide(pres, layout, did);
      pres.slides.push(s);
      const ph = (types) => s.shapes.filter((x) => x.ph && types.includes(x.ph.type));
      const tt = ph(['ctrTitle', 'title'])[0];
      if (tt) T.setPlain(tt.tx, title);
      if (layout === 'title') { const sub = ph(['subTitle'])[0]; if (sub && cols[0]) sub.tx.ps = cols[0].map((ln) => T.para(ln)); }
      else ph(['body']).forEach((b, i) => { if (cols[i]) b.tx.ps = T.fromLines(cols[i]); });
      s.trans = { type: 'fade', spd: 'med', click: true, after: null };
      return s;
    };
    B[id](add, pres);
    pres.props.title = '';
    pres.hf = Object.assign(pres.hf, { num: true, notOnTitle: true });
    return pres;
  };
})();
