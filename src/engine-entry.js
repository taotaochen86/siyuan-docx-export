const siyuan = require("./siyuan");
const builder = require("./builder");

window.__docxExport = {
  version: "0.3.0",
  fetchTree: siyuan.fetchTree,
  api: siyuan.api,
  buildDocument: builder.buildDocument,
  collectZoteroKeys: siyuan.collectZoteroKeys,
  fetchZoteroRefs: siyuan.fetchZoteroRefs,
  collectImageRefs: siyuan.collectImageRefs,
  prefetchImages: siyuan.prefetchImages,
};
console.log("[docx-export] engine ready v0.3");
