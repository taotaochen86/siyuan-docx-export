const siyuan = require("../src/siyuan");
const M = require("../src/markdown");
const ROOT = process.argv[2];
(async () => {
  const raw = await siyuan.fetchTree(ROOT, true);
  const anchors = M.collectAnchors(raw.docs);
  const refId = "";
  console.log("=== parseDoc 识别出的 caption ===");
  for (const d of raw.docs) {
    const bs = M.parseDoc(d, d.markdown, anchors, refId);
    const caps = bs.filter((b) => b.type === "caption");
    if (caps.length) console.log(`[${d.title}]`, caps.map((c) => `${c.kind}|${c.num}|${c.title}`).join("  "));
  }
  console.log("\n=== markdown 中以 表/图/式 开头的行（前 40）===");
  let n = 0;
  for (const d of raw.docs) {
    for (const ln of d.markdown.split("\n")) {
      const t = ln.trim();
      if (!/^(表|图|式)\s*[0-9]/.test(t)) continue;
      if (++n > 40) break;
      console.log(`[${d.title}] ${JSON.stringify(t.slice(0, 80))}`);
    }
  }
})();
