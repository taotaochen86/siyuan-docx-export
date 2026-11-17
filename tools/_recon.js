// 内容对账：思源 Markdown 统计 vs docx 实际内容
const fs = require("fs");
const { unzipSync } = require("fflate");

const stat = JSON.parse(fs.readFileSync(process.argv[2], "utf8"));
const u = unzipSync(new Uint8Array(fs.readFileSync(process.argv[3])));
const d = new TextDecoder().decode(u["word/document.xml"]);
const text = [...d.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1]).join("");

// 松比对：剥掉自动编号前缀 + 空白 + 零宽 + 标点差异
const norm = (s) =>
  s.replace(/[\s\u200b-\u200d\ufeff]/g, "")
    .replace(/^第[一二三四五六七八九十百零\d]+[章节篇]/, "")
    .replace(/^[一二三四五六七八九十]+[、.．]/, "")
    .replace(/^[（(][一二三四五六七八九十\d]+[）)]/, "")
    .replace(/^\d+(\.\d+)*[.．、]?\s*/, "")
    .replace(/[：:，,。.、（）()【】\[\]""''"‘’…—\-]/g, "");

const all = norm(text);
let missing = [];
let checked = 0;
for (const doc of stat.docs) {
  for (const h of doc.heads) {
    if (!h) continue;
    checked += 1;
    if (!all.includes(norm(h))) missing.push(doc.title + " :: " + h);
  }
}

// docx 侧结构计数
const cnt = {
  Heading: (d.match(/w:val="Heading\d"/g) || []).length,
  HeadingDoc: (d.match(/w:val="HeadingDoc"/g) || []).length,
  Tables: (d.match(/<w:tbl>/g) || []).length,
  TableRows: (d.match(/<w:tr>/g) || []).length,
  Drawings: (d.match(/<w:drawing>/g) || []).length,
  PageBreaks: (d.match(/w:type="page"/g) || []).length,
  MathPara: (d.match(/<m:oMathPara>/g) || []).length,
  MathInline: (d.match(/<m:oMath>/g) || []).length,
  Fields: (d.match(/w:fldCharType="begin"/g) || []).length,
};

console.log("=== 标题文本覆盖率 ===");
console.log(`思源 Markdown 标题 ${checked} 条，docx 中命中 ${checked - missing.length} 条`);
if (missing.length) {
  console.log("未命中：");
  missing.forEach((m) => console.log("  - " + m));
} else {
  console.log("全部命中，无标题丢失");
}

console.log("\n=== 结构数量对账 ===");
const rows = [
  ["文档篇数", stat.totalDocs, cnt.PageBreaks + 1],
  ["标题（含篇标题）", stat.sum.heads, cnt.Heading + cnt.HeadingDoc],
  ["图片引用", stat.sum.imgs, cnt.Drawings],
  ["表格行", stat.sum.tblRows, cnt.TableRows],
  ["公式块", stat.sum.mathBlocks, cnt.MathPara],
  ["内联公式", stat.sum.inlineMath, cnt.MathInline],
  ["正文/其他行", stat.sum.bodyLines, cnt.TableRows + cnt.Drawings],
];
for (const [k, a, b] of rows) console.log(`  ${k}: 思源 ${a} / docx ${b}`);
console.log("\ndocx 明细: " + JSON.stringify(cnt));
