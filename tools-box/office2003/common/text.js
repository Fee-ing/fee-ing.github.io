/* Office 2003 — the text the UI shows and the {0}, {1}… placeholders in it.
 *
 * Embedded copy (tools-box): the suite is Chinese only, so upstream's i18n runtime is gone from
 * this copy (common/lang.js, common/i18n.js and the dictionaries under common/i18n/). Every
 * string the UI shows is written in Chinese where it is used, and nothing is looked up at run
 * time.
 *
 * VO.t / L.t fill the placeholders from the arguments that follow; a string with no argument
 * comes back as it is. VO.N / L.N mark UI text in the tables that mix it with data (they return
 * the string unchanged, as upstream's N does when there is no dictionary). VO.plain is upstream's
 * "the text as it is shown": the Chinese is already plain, so it returns the string.
 */
(function (root) {
  'use strict';
  const VO = (root.VO = root.VO || {});
  const t = (s, ...args) => {
    if (typeof s !== 'string' || !args.length) return s;
    return s.replace(/\{(\d+)\}/g, (m, i) => (args[i] === undefined || args[i] === null ? m : args[i]));
  };
  VO.t = t;
  VO.N = (s) => s;
  VO.plain = (s) => String(s == null ? '' : s);
  if (root.L) { root.L.t = t; root.L.N = VO.N; }
})(typeof window !== 'undefined' ? window : globalThis);
