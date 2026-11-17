const fflate = require("fflate"), fs = require("fs");
const x = Buffer.from(fflate.unzipSync(fs.readFileSync(process.argv[2]))["word/document.xml"]).toString();
for (const p of x.split("</w:p>")) {
  if (!/eqArr/.test(p)) continue;
  const m = p.match(/<m:oMath[\s\S]*?<\/m:oMath>/)[0];
  console.log("begChr={ :", /begChr m:val="\{"/.test(m));
  console.log("endChr=. :", /endChr m:val="\."/.test(m));
  console.log("eqArr 行 m:e 数:", (m.match(/<m:e>/g) || []).length);
  console.log("分数:", /<m:f>/.test(m), " 竖线组:", /begChr m:val="\|"/.test(m));
  console.log("纯文本:", m.replace(/<[^>]+>/g, ""));
}
