const fs = require("fs");
const F = require("fflate");
const entries = F.unzipSync(new Uint8Array(fs.readFileSync(process.argv[2])));
const xml = new TextDecoder().decode(entries["word/document.xml"]);
const paras = xml.match(/<w:p\b[\s\S]*?<\/w:p>/g) || [];
const CN = { 零: 0, 一: 1, 二: 2, 三: 3, 四: 4, 五: 5, 六: 6, 七: 7, 八: 8, 九: 9 };
function cnNum(s) {
  if (/^[0-9]+$/.test(s)) return parseInt(s, 10);
  if (s.length === 1 && CN[s] != null) return CN[s];
  const m = /^([一二三四五六七八九])?十([一二三四五六七八九])?$/.exec(s);
  if (m) return (m[1] ? CN[m[1]] : 1) * 10 + (m[2] ? CN[m[2]] : 0);
  return null;
}
let n = 0;
for (const p of paras) {
  if (!/w:val="Heading1"/.test(p)) continue;
  n++;
  const t = p.replace(/<w:instrText[\s\S]*?<\/w:instrText>/g, "").replace(/<[^>]+>/g, "").trim();
  const m = /第\s*([0-9]+|[一二三四五六七八九十]+)\s*章/.exec(t);
  const r = /SEQ 章 \\\* ARABIC \\r ([0-9]+)/.exec(p.replace(/<[^>]+>/g, " "));
  console.log(`${n}. [${t}] 解析章号=${m ? cnNum(m[1]) : "null"} 插入\\r=${r ? r[1] : "(无)"}`);
}
console.log("Heading1 总数:", n);
