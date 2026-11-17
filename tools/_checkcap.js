// 检查题注段的书签范围：bookmarkEnd 必须早于题名文字（否则 F9 更新 REF 会塞进整条题注）
const fs = require("fs");
const { unzipSync } = require("fflate");

const doc = new TextDecoder().decode(unzipSync(new Uint8Array(fs.readFileSync(process.argv[2])))["word/document.xml"]);
const pRe = /<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g;
const tokRe = /<w:bookmarkStart[^>]*w:name="([^"]*)"[^>]*\/>|<w:bookmarkEnd[^>]*\/>|<w:instrText[^>]*>([^<]*)<\/w:instrText>|<w:t[^>]*>([^<]*)<\/w:t>/g;

let n = 0, bad = 0, skipped = 0;
for (const m of doc.matchAll(pRe)) {
  const p = m[0];
  if (!/SEQ\s+(表|图|式)/.test(p)) continue;
  n += 1;
  const ev = [];
  for (const t of p.matchAll(tokRe)) {
    if (t[1] !== undefined) ev.push({ k: "BM_START", v: t[1] });
    else if (t[2] !== undefined) ev.push({ k: "INSTR", v: t[2].trim() });
    else if (t[3] !== undefined) ev.push({ k: "TEXT", v: t[3] });
    else ev.push({ k: "BM_END" });
  }
  let s = -1, e = -1;
  ev.forEach((x, i) => {
    if (x.k === "BM_START" && s < 0) s = i;
    if (x.k === "BM_END" && s >= 0 && e < 0) e = i;
  });
  const label = ev.filter((x) => x.k !== "BM_START").map((x) => (x.k === "TEXT" ? `"${x.v}"` : x.k === "INSTR" ? `[${x.v}]` : "END")).join(" → ");
  if (s < 0) { skipped += 1; console.log(`SKIP 无书签  ${label}`); continue; }
  const inside = ev.slice(s, e + 1).filter((x) => x.k === "TEXT").map((x) => x.v).join("").trim();
  const outside = ev.slice(e + 1).filter((x) => x.k === "TEXT").map((x) => x.v).join("").trim();
  // 正确：书签内只有「表 4-1」这类编号；题名必须在书签外
  const ok = outside.length > 0;
  if (!ok) bad += 1;
  console.log(`${ok ? "OK  " : "BAD "} 书签内="${inside}"  书签外="${outside}"`);
}
console.log(`\n题注段: ${n}  有书签: ${n - skipped}  范围错误: ${bad}`);

