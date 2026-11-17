const siyuan = require("../src/siyuan");
const ROOT = process.argv[2];
(async () => {
  const raw = await siyuan.fetchTree(ROOT, true);
  for (const d of raw.docs) {
    if (!d.title) continue;
    const heads = d.markdown.split("\n").filter((l) => /^#{1,2}\s/.test(l));
    console.log(`[${d.title}]`);
    heads.slice(0, 6).forEach((h) => console.log("   " + h.trim()));
  }
})();
