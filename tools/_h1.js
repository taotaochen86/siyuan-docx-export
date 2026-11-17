const fs = require("fs");
const { unzipSync } = require("fflate");
const u = unzipSync(new Uint8Array(fs.readFileSync(process.argv[2])));
const d = new TextDecoder().decode(u["word/document.xml"]);
// 用正则找所有段落起点（覆盖 <w:p> / <w:p xxx> / <w:p/>）
const starts = [...d.matchAll(/<w:p(?:\s[^>]*)?(?:\/?)>/g)].map((m) => m.index);
const paras = starts.map((s, i) => d.slice(s, i + 1 < starts.length ? starts[i + 1] : s + 4000));
let shown = 0;
paras.forEach((p, i) => {
  const m = p.match(/<w:pStyle w:val="(Heading\d|HeadingDoc)"/);
  if (!m) return;
  shown += 1;
  if (shown > 12) return;
  const ts = [...p.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((x) => x[1]).join("");
  console.log(`${i + 1} ${m[1]} [${ts.slice(0, 34)}]`);
});
console.log("段落总数(TAG计数)=" + starts.length + "  Heading 段落=" + shown);
const cnt = (d.match(/w:val="Heading1"/g) || []).length;
console.log("Heading1 出现次数=" + cnt);
