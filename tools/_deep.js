// 深检：公式残留 / 表格边框与守恒 / 空单元格 / 分页
const fs = require("fs");
const { unzipSync } = require("fflate");

const u = unzipSync(new Uint8Array(fs.readFileSync(process.argv[2])));
const d = new TextDecoder().decode(u["word/document.xml"]);
const text = [...d.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1]).join("");

console.log("=== 公式残留检查 ===");
const dollar = (text.match(/\$/g) || []).length;
const backslash = (text.match(/\\[a-zA-Z]+/g) || []).slice(0, 12);
console.log(`可见文本中 "$" 个数: ${dollar}${dollar ? "  ← 存在 LaTeX 原文残留" : "（干净）"}`);
if (backslash.length) console.log("疑似 LaTeX 命令残留: " + backslash.join(" "));
const sub = (d.match(/<w:vertAlign w:val="subscript"\/>/g) || []).length;
const sup = (d.match(/<w:vertAlign w:val="superscript"\/>/g) || []).length;
console.log(`上下标 run: 下标 ${sub} / 上标 ${sup}`);
// 残留样本
const bad = [];
const runs = [...d.matchAll(/<w:t[^>]*>([^<]*\$[^<]*)<\/w:t>/g)];
runs.slice(0, 8).forEach((m) => bad.push(m[1].slice(0, 60)));
if (bad.length) { console.log("含 $ 的文本片段（前 8）:"); bad.forEach((b) => console.log("   " + b)); }
// 真正的行内公式残留：成对 $ 且内部带 LaTeX 特征（_ ^ \）
const latexLike = [...text.matchAll(/\$([^$]{1,80})\$/g)].filter((m) => /[\\_^]/.test(m[1]));
console.log(`\n疑似未渲染行内公式（$..$ 且含 _ ^ \\）: ${latexLike.length} 处`);
latexLike.slice(0, 10).forEach((m) => console.log("   $" + m[1].slice(0, 50) + "$"));

console.log("\n=== 表格检查 ===");
const tbls = d.split("<w:tbl>").slice(1);
console.log(`表格数 ${tbls.length}`);
let noBorder = 0, nonThreeLine = 0, emptyCell = 0, zeroRow = 0;
const dims = [];
tbls.forEach((t, i) => {
  const rows = (t.match(/<w:tr>/g) || []).length;
  const firstRow = t.match(/<w:tr>[\s\S]*?<\/w:tr>/);
  const cols = firstRow ? (firstRow[0].match(/<w:tc>/g) || []).length : 0;
  dims.push(`${rows}x${cols}`);
  if (rows === 0) zeroRow += 1;
  const b = t.match(/<w:tblBorders>([\s\S]*?)<\/w:tblBorders>/);
  if (!b) { noBorder += 1; return; }
  const top = /<w:top w:val="single" w:sz="12"/.test(b[1]);
  const bottom = /<w:bottom w:val="single" w:sz="12"/.test(b[1]);
  const noV = /<w:insideV w:val="nil"/.test(b[1]) && /<w:left w:val="nil"/.test(b[1]) && /<w:right w:val="nil"/.test(b[1]);
  if (!(top && bottom && noV)) nonThreeLine += 1;
  emptyCell += (t.match(/<w:tc><w:tcPr[^>]*>[\s\S]*?<\/w:tc>/g) || []).filter((c) => !/<w:t[^>]*>[^<]+<\/w:t>/.test(c)).length;
});
console.log(`无边框定义 ${noBorder} / 非三线 ${nonThreeLine} / 零行表 ${zeroRow} / 空单元格 ${emptyCell}`);
console.log("每表 行x列: " + dims.join(" "));

console.log("\n=== 其他 ===");
console.log(`分页符 ${(d.match(/w:type="page"/g) || []).length}`);
console.log(`段落 ${(d.match(/<w:p[ >]/g) || []).length}`);
console.log(`引用缺失占位 ${(text.match(/\[图片缺失: /g) || []).length}`);
console.log(`脚注定义行残留 ${(text.match(/\[\^\d+\]:/g) || []).length}`);
console.log(`制表符 run ${(d.match(/<w:tab\/>/g) || []).length}`);
