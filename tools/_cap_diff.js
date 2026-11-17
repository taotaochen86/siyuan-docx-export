// 逐条对比：源 markdown 题注原文 vs 产物题注渲染文本
// 用法: node tools/_cap_diff.js <rootId> <docx>
const fs = require("fs");
const F = require("fflate");
const siyuan = require("../src/siyuan");

const ROOT = process.argv[2];
const DOCX = process.argv[3];

const dec = (s) => s.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
const textOf = (xml) => dec(
  xml
    .replace(/<w:instrText[\s\S]*?<\/w:instrText>/g, "")   // 域指令不是可见文本
    .replace(/<[^>]+>/g, "")
);

(async () => {
  const raw = await siyuan.fetchTree(ROOT, true);

  // 源：每条题注原文（行首 表/图/式 + 编号），记录所属文档
  const src = [];
  for (const d of raw.docs) {
    for (const ln of d.markdown.split("\n")) {
      const t = ln.trim().replace(/\*\*/g, "").replace(/\\\./g, ".");
      const m = /^(表|图|式)\s*([0-9]+(?:\s*[-–—]\s*[0-9]+)?)\s*(.*)$/.exec(t);
      if (m) src.push({ doc: d.title, kind: m[1], num: m[2].replace(/\s+/g, ""), title: (m[3] || "").trim() });
    }
  }

  // 产物：含 SEQ 域的段落，按其纯文本取
  const entries = F.unzipSync(new Uint8Array(fs.readFileSync(DOCX)));
  const doc = new TextDecoder().decode(entries["word/document.xml"]);
  const paras = doc.match(/<w:p\b[\s\S]*?<\/w:p>/g) || [];
  const out = [];
  for (const p of paras) {
    if (!/SEQ (表|图|式) /.test(p)) continue;
    out.push(textOf(p).replace(/\s+/g, " ").trim());
  }

  console.log(`源题注（行首匹配）: ${src.length} 条 | 产物题注段: ${out.length} 段\n`);
  console.log("=== 产物题注（按出现顺序）===");
  out.forEach((t, i) => console.log(`${String(i + 1).padStart(2)}. ${t}`));

  console.log("\n=== 源题注（按文档顺序）===");
  src.forEach((s, i) => console.log(`${String(i + 1).padStart(2)}. [${s.doc}] ${s.kind} ${s.num} ${s.title}`.slice(0, 90)));

  // 关键比对：源里带编号的题注编号集合 vs 产物里的编号集合
  const srcNums = new Set(src.filter((s) => /-/.test(s.num)).map((s) => s.kind + s.num));
  console.log("\n源里「章-序」编号:", [...srcNums].join("  ") || "(无)");

  const parsedNums = [];
  for (const t of out) {
    const m = /^(表|图|式)\s*([0-9]+(?:-[0-9]+)?)/.exec(t);
    if (m) parsedNums.push(m[1] + m[2]);
  }
  console.log("产物里题注编号:", parsedNums.join("  "));
  const mismatch = parsedNums.filter((k) => !srcNums.has(k));
  console.log(mismatch.length ? `\n产物编号不在源清单: ${mismatch.join("  ")}` : "\n产物编号全部能在源清单里找到");
})();
