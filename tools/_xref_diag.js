const siyuan = require("../src/siyuan");
const M = require("../src/markdown");
const D = require("../src/docx");
const ROOT = process.argv[2];
(async () => {
  const raw = await siyuan.fetchTree(ROOT, true);
  const anchors = M.collectAnchors(raw.docs);
  const refId = (raw.docs.find((d) => d.title.trim().toLowerCase() === "reference") || {}).id || "";
  const caps = new Set();
  for (const d of raw.docs) {
    for (const b of M.parseDoc(d, d.markdown, anchors, refId)) {
      if (b.type === "caption") caps.add(b.kind + "|" + b.num);
      if (b.type === "math" && b.num) caps.add("式|" + b.num);
    }
  }
  console.log("=== 题注/公式编号（kind|num）===");
  console.log([...caps].join("  "));
  console.log("\n=== 引用 label -> 是否命中 ===");
  const re = /\[([^\]]*)\]\(#([^)]*)\)/g;
  const miss = [];
  for (const d of raw.docs) {
    let m; re.lastIndex = 0;
    while ((m = re.exec(d.markdown)) !== null) {
      const label = m[1].trim();
      const cm = /^(表|图|式)\s*([0-9]+(?:-[0-9]+)?)\s*$/.exec(label);
      if (!cm) { miss.push(`[非编号] ${label} @${d.title}`); continue; }
      const hit = caps.has(cm[1] + "|" + cm[2]);
      if (!hit) miss.push(`[未命中] ${label} @${d.title}`);
    }
  }
  console.log(miss.length ? miss.join("\n") : "全部命中");
})();
