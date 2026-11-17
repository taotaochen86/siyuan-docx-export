const siyuan = require("../src/siyuan");
const ROOT = process.argv[2];
(async () => {
  const raw = await siyuan.fetchTree(ROOT, true);
  for (const d of raw.docs) {
    const lines = d.markdown.split("\n");
    let inCode = false;
    lines.forEach((ln, idx) => {
      if (/^\s*```/.test(ln)) { inCode = !inCode; return; }
      if (inCode) return;
      if (/\$[^$]*\$/.test(ln)) {
        console.log(`[${d.title}] L${idx + 1}: ${ln.trim()}`);
      }
    });
  }
})();
