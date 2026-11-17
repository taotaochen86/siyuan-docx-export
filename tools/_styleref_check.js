const fflate = require("fflate"), fs = require("fs");
const u = fflate.unzipSync(fs.readFileSync(process.argv[2]));
const x = Buffer.from(u["word/document.xml"]).toString();
const styles = Buffer.from(u["word/styles.xml"]).toString();
console.log("=== styles.xml 标题样式名 ===");
[...styles.matchAll(/<w:style [^>]*w:styleId="(Heading\d+|HeadingDoc)"[^>]*>[\s\S]*?<w:name w:val="([^"]*)"/g)]
  .forEach((m) => console.log(m[1], "=> name=", m[2]));
console.log("\n=== 含 STYLEREF 的段落 ===");
for (const p of x.split("</w:p>")) {
  if (!/STYLEREF/.test(p)) continue;
  const instr = [...p.matchAll(/<w:instrText[^>]*>([^<]*)<\/w:instrText>/g)].map((z) => z[1].trim());
  const txt = [...p.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((z) => z[1]).join("");
  console.log("指令:", JSON.stringify(instr), "| 文本:", txt);
}
