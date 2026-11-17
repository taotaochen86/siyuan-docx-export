const siyuan = require("../src/siyuan");
const ROOT = process.argv[2];
(async () => {
  const raw = await siyuan.fetchTree(ROOT, true);
  let total = 0;
  for (const d of raw.docs) {
    for (const ln of d.markdown.split("\n")) {
      if (!/\]\(#/.test(ln)) continue;
      total += 1;
      if (total <= 25) console.log(`[${d.title}] ${ln.trim().slice(0, 200)}`);
    }
  }
  console.log("含 #链接 的行数:", total);
  // 分类统计链接形态
  const re = /\[([^\]]*)\]\(#([^)]*)\)/g;
  const forms = new Map();
  for (const d of raw.docs) {
    let m;
    re.lastIndex = 0;
    while ((m = re.exec(d.markdown)) !== null) {
      const label = m[1];
      const key = /^(表|图|式)\s*[0-9]+(?:-[0-9]+)?$/.test(label.trim()) ? "纯编号label" : "其他label:" + JSON.stringify(label.slice(0, 20));
      forms.set(key, (forms.get(key) || 0) + 1);
    }
  }
  console.log("\n=== label 形态 ===");
  [...forms.entries()].sort((a, b) => b[1] - a[1]).forEach(([k, v]) => console.log(v, k));
})();
