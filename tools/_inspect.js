const fs = require("fs");
const path = require("path");
const { zipSync, unzipSync } = require("fflate");

const p = process.argv[2];
const u = unzipSync(new Uint8Array(fs.readFileSync(p)));
const doc = new TextDecoder().decode(u["word/document.xml"]);
const sty = new TextDecoder().decode(u["word/styles.xml"]);
const num = new TextDecoder().decode(u["word/numbering.xml"]);

console.log("=== styles.xml: Heading1 定义 ===");
const h1 = /<w:style[^>]*w:styleId="Heading1"[^>]*>[\s\S]*?<\/w:style>/.exec(sty);
console.log(h1 ? h1[0].replace(/\s+/g, " ").slice(0, 500) : "未找到");

console.log("\n=== numbering.xml 前 400 字 ===");
console.log(num.replace(/\s+/g, " ").slice(0, 400));

console.log("\n=== 标题段落（前 14 个，看有无重号）===");
const hs = [...doc.matchAll(/<w:p>(?:(?!<\/w:p>)[\s\S])*?<w:pStyle w:val="(Heading\d|HeadingDoc)"\/>[\s\S]*?<\/w:p>/g)];
let n = 0;
for (const m of hs) {
  const style = m[1];
  const text = (m[0].match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || [])
    .map((t) => t.replace(/<[^>]+>/g, ""))
    .join("");
  console.log(`  [${style}] ${text}`);
  if (++n >= 14) break;
}
console.log(`  标题段落总数: ${hs.length}`);

const vert = (doc.match(/<w:vertAlign w:val="(subscript|superscript)"/g) || []).length;
console.log(`\n上下标 run: ${vert}`);
const numPr = (doc.match(/<w:numPr>/g) || []).length;
console.log(`段落内 numPr: ${numPr}（应为 0，编号由样式继承）`);
const tabs = (doc.match(/<w:tab w:val="center"/g) || []).length;
console.log(`公式段 tab stops: ${tabs}`);
const zotItem = (doc.match(/ADDIN ZOTERO_ITEM/g) || []).length;
console.log(`ZOTERO_ITEM 域: ${zotItem}`);
const zotBibl = (doc.match(/ADDIN ZOTERO_BIBL/g) || []).length;
console.log(`ZOTERO_BIBL 域: ${zotBibl}`);

const seqShi = (doc.match(/SEQ 式/g) || []).length;
const seqBiao = (doc.match(/SEQ 表/g) || []).length;
const seqTu = (doc.match(/SEQ 图/g) || []).length;
const refFld = (doc.match(/\bREF bkmk_/g) || []).length;
const bms = (doc.match(/<w:bookmarkStart /g) || []).length;
const pgbk = (doc.match(/w:type="page"/g) || []).length;
const tbls = (doc.match(/<w:tbl>/g) || []).length;
const nilV = (doc.match(/w:insideV w:val="nil"/g) || []).length;
const fldB = (doc.match(/w:fldCharType="begin"/g) || []).length;
const fldE = (doc.match(/w:fldCharType="end"/g) || []).length;
console.log(`SEQ 式/表/图: ${seqShi} / ${seqBiao} / ${seqTu}`);
console.log(`REF 交叉引用域: ${refFld}  书签: ${bms}`);
console.log(`分页符: ${pgbk}  表格: ${tbls}  三线表(insideV nil): ${nilV}`);
console.log(`域配对 begin/end: ${fldB} / ${fldE}`);

// ---- 图片嵌入 ----
const media = Object.keys(u).filter((k) => k.startsWith("word/media/"));
const ct = new TextDecoder().decode(u["[Content_Types].xml"]);
const rels = new TextDecoder().decode(u["word/_rels/document.xml.rels"]);
const drawings = (doc.match(/<w:drawing>/g) || []).length;
const embeds = (doc.match(/r:embed="rIdImg\d+"/g) || []).length;
const imgRelCount = (rels.match(/relationships\/image/g) || []).length;
const ctImg = (ct.match(/ContentType="image\/[a-z]+"/g) || []);
console.log(`\n媒体部件: ${media.length} ${media.length ? "[" + media.map((m) => m.replace("word/media/", "") + "(" + u[m].length + "B)").join(", ") + "]" : ""}`);
console.log(`drawing 节点: ${drawings}  r:embed 引用: ${embeds}`);
console.log(`image 关系: ${imgRelCount}  Content_Types 图片扩展名: ${ctImg.length ? ctImg.join(", ") : "无"}`);
const miss = (doc.match(/\[图片缺失: [^\]]+\]/g) || []).length;
const ph = (doc.match(/\[图片: [^\]]+\]/g) || []).length;
console.log(`图片缺失占位: ${miss}  旧版图片占位: ${ph}`);
