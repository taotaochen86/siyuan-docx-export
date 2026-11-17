const siyuan = require("../src/siyuan");
const ROOT = process.argv[2] || "20261107193423-jjppxdp";
(async () => {
  const raw = await siyuan.fetchTree(ROOT, true);
  for (const d of raw.docs) {
    const marks = {
      dollar: (d.markdown.match(/\$[^$]+\$/g) || []),
      paren: (d.markdown.match(/\\\([^)]*\\\)/g) || []),
      ddollar: (d.markdown.match(/\$\$[^$]+\$\$/g) || []),
    };
    const all = [...marks.dollar, ...marks.paren, ...marks.ddollar];
    if (all.length) {
      console.log("=== " + d.title + " (" + d.id + ") ===");
      all.slice(0, 12).forEach((m) => console.log("  " + m));
    }
  }
})();
