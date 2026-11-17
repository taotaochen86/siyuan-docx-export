const fs = require("fs");
const F = require("fflate");
const entries = F.unzipSync(new Uint8Array(fs.readFileSync(process.argv[2])));
const xml = new TextDecoder().decode(entries["word/document.xml"]);
const paras = xml.match(/<w:p\b[\s\S]*?<\/w:p>/g) || [];
let n = 0;
for (const p of paras) {
  if (!/w:val="Heading1"/.test(p)) continue;
  n++;
  const t = p.replace(/<w:instrText[\s\S]*?<\/w:instrText>/g, "").replace(/<[^>]+>/g, "");
  if (n <= 4) {
    console.log(`H1#${n} text=[${t}] vanish=${(p.match(/<w:vanish/g) || []).length}`);
    console.log("  raw:", p.replace(/\s+/g, " ").slice(0, 500));
  }
}
console.log("Heading1 总数:", n);
