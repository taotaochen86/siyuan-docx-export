// 校验下载到 Downloads 的 docx：结构对账
const fs = require("fs");
const path = require("path");
const fflate = require("fflate");

const dl = path.join(process.env.USERPROFILE, "Downloads");
const arg = process.argv[2];
const cands0 = arg
  ? [{ f: path.basename(arg), p: arg, m: fs.statSync(arg).mtimeMs }]
  : fs.readdirSync(dl)
  .filter((f) => f.toLowerCase().endsWith(".docx"))
  .map((f) => ({ f, p: path.join(dl, f), m: fs.statSync(path.join(dl, f)).mtimeMs }))
  .sort((a, b) => b.m - a.m);
const cands = cands0;
if (!cands.length) { console.log("Downloads 无 docx"); process.exit(0); }
const top = cands[0];
console.log("最新文件:", top.f, new Date(top.m).toLocaleString(), fs.statSync(top.p).size, "bytes");
if (cands[1]) console.log("上一个:", cands[1].f, new Date(cands[1].m).toLocaleString());

let files;
try { files = fflate.unzipSync(fs.readFileSync(top.p)); }
catch (e) { console.log("解压失败:", e.message); process.exit(1); }
console.log("部件:", Object.keys(files).join(", "));

const doc = Buffer.from(files["word/document.xml"]).toString("utf-8");
const c = (re) => (doc.match(re) || []).length;
console.log("Heading:", c(/w:val="Heading[1-4]"/g));
console.log("表格:", c(/<w:tbl>/g), " 行:", c(/<w:tr>/g));
console.log("图片(内嵌):", c(/<w:drawing>/g), " 图片占位:", c(/\[图片/g));
console.log("公式OMML:", c(/<m:oMath/g), " 公式降级$:", c(/\$\$/g));
console.log("SEQ题注:", c(/SEQ (表|图)/g));
console.log("REF交叉引用:", c(/REF bkmk/g));
console.log("ZOTERO_ITEM:", c(/ZOTERO_ITEM/g));
console.log("ZOTERO_BIBL:", c(/ZOTERO_BIBL/g));
console.log("bookmarkStart:", c(/bookmarkStart/g));
console.log("分页符:", c(/w:type="page"/g));
console.log("三线表top:", c(/w:top w:val="single"/g), " insideV nil:", c(/insideV w:val="nil"/g));
console.log("域配对 begin/end:", c(/fldCharType="begin"/g), "/", c(/fldCharType="end"/g));

// 标签配对
let bad = 0;
for (const [name, data] of Object.entries(files)) {
  if (!name.endsWith(".xml") && !name.endsWith(".rels")) continue;
  const t = Buffer.from(data).toString("utf-8");
  const stack = []; const re = /<\/?([a-zA-Z0-9:]+)([^>]*?)(\/?)>/g; let m;
  while ((m = re.exec(t)) !== null) {
    if (m[0][1] === "/") { if (stack.pop() !== m[1]) { console.log("标签不配对", m[1], name); bad++; break; } }
    else if (!m[3]) stack.push(m[1]);
  }
  if (stack.length) { console.log("未闭合", stack.slice(-2), name); bad++; }
}
console.log(bad ? "XML 有问题:" + bad : "XML 良构 OK");
