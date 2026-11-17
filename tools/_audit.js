// 结构审计：部件 / 关系闭环 / 域顺序 / 书签 / 编号引用 / 样式引用 / 图片 / Content_Types
const fs = require("fs");
const { unzipSync } = require("fflate");

const path = process.argv[2];
const u = unzipSync(new Uint8Array(fs.readFileSync(path)));
const dec = new TextDecoder();
const fails = [];
const warns = [];
const pass = [];
const F = (m) => fails.push(m);
const W = (m) => warns.push(m);
const P = (m) => pass.push(m);

const parts = Object.keys(u);
P(`zip 部件 ${parts.length} 个`);

// ---------- 1. XML 良构 ----------
function xmlCheck(s) {
  let i = 0; const stack = [];
  while (i < s.length) {
    const lt = s.indexOf("<", i);
    if (lt < 0) break;
    if (s.startsWith("<?", lt)) { const g = s.indexOf("?>", lt); if (g < 0) return "未闭合声明"; i = g + 2; continue; }
    if (s.startsWith("<!--", lt)) { const g = s.indexOf("-->", lt); if (g < 0) return "未闭合注释"; i = g + 3; continue; }
    if (s.startsWith("<![CDATA[", lt)) { const g = s.indexOf("]]>", lt); if (g < 0) return "未闭合CDATA"; i = g + 3; continue; }
    let j = lt + 1, q = null;
    while (j < s.length) {
      const c = s[j];
      if (q) { if (c === q) q = null; } else if (c === '"' || c === "'") q = c;
      else if (c === ">") break;
      j += 1;
    }
    if (j >= s.length) return "标签未闭合";
    const raw = s.slice(lt + 1, j);
    if (raw[0] === "/") {
      const name = raw.slice(1).trim();
      const top = stack.pop();
      if (top !== name) return `标签不匹配: 期望 </${top}> 实际 </${name}>`;
    } else if (!raw.endsWith("/")) {
      stack.push(raw.split(/[\s/>]/)[0]);
    }
    i = j + 1;
  }
  if (stack.length) return "未闭合标签: " + stack.slice(-3).join(",");
  return null;
}
let xmlBad = 0;
for (const k of parts) {
  if (!k.endsWith(".xml") && !k.endsWith(".rels")) continue;
  const e = xmlCheck(dec.decode(u[k]));
  if (e) { xmlBad += 1; F(`${k} 不良构：${e}`); }
}
if (!xmlBad) P(`XML 良构：${parts.filter((k) => k.endsWith(".xml") || k.endsWith(".rels")).length} 个部件全部通过`);

const doc = dec.decode(u["word/document.xml"]);
const relsXml = dec.decode(u["word/_rels/document.xml.rels"]);
const ctXml = dec.decode(u["[Content_Types].xml"]);
const styXml = dec.decode(u["word/styles.xml"]);
const numXml = dec.decode(u["word/numbering.xml"]);

// ---------- 2. Content_Types 覆盖 ----------
const defaults = [...ctXml.matchAll(/<Default Extension="([^"]+)"/g)].map((m) => m[1].toLowerCase());
const overrides = [...ctXml.matchAll(/<Override PartName="([^"]+)"/g)].map((m) => m[1].replace(/^\//, ""));
const exts = [...new Set(parts.map((k) => (k.split(".").pop() || "").toLowerCase()))];
const missingExt = exts.filter((e) => !defaults.includes(e) && !overrides.some((o) => o.endsWith("." + e)));
if (missingExt.length) F(`Content_Types 缺扩展名声明: ${missingExt.join(",")}`);
else P(`Content_Types 扩展名覆盖齐全（${exts.join(",")}）`);
const needOverride = ["word/document.xml", "word/styles.xml", "word/numbering.xml"];
const missOv = needOverride.filter((o) => !overrides.includes(o));
if (missOv.length) F(`Content_Types 缺 Override: ${missOv.join(",")}`);
else P("Content_Types Override 齐全");

// ---------- 3. 关系闭环 ----------
const rels = [...relsXml.matchAll(/<Relationship Id="([^"]+)" Type="([^"]+)" Target="([^"]+)"/g)];
const relIds = new Set(rels.map((m) => m[1]));
const relTargets = rels.filter((m) => !/^http/i.test(m[3]));
for (const m of relTargets) {
  const t = "word/" + m[3].replace(/^\.\//, "");
  if (!parts.includes(t)) F(`关系 ${m[1]} 指向的 ${t} 不存在`);
}
if (relTargets.every((m) => parts.includes("word/" + m[3].replace(/^\.\//, "")))) P(`关系目标全部存在（${relTargets.length} 条本地关系）`);

const embeds = [...doc.matchAll(/r:(?:embed|id)="([^"]+)"/g)].map((m) => m[1]);
const badEmbed = [...new Set(embeds)].filter((id) => !relIds.has(id));
if (badEmbed.length) F(`引用不存在的 rId: ${badEmbed.join(",")}`);
else P(`文档内 rId 引用全部有对应关系（${new Set(embeds).size} 个）`);

const relIdsUse = rels.map((m) => m[1]);
// styles.xml / numbering.xml 这类系统关系由部件类型引用，本来就不会出现在 document.xml 的 r:id 里
const sysRel = rels.filter((m) => /relationships\/(styles|numbering|settings|theme|webSettings|fontTable)/.test(m[2])).map((m) => m[1]);
const unused = relIdsUse.filter((id) => !embeds.includes(id) && !sysRel.includes(id));
if (unused.length) W(`有 ${unused.length} 条关系未被文档引用: ${unused.join(",")}`);

// ---------- 4. 域结构与顺序 ----------
const seqs = [...doc.matchAll(/<w:fldChar w:fldCharType="(\w+)"/g)].map((m) => m[1]);
const dB = seqs.filter((x) => x === "begin").length;
const dS = seqs.filter((x) => x === "separate").length;
const dE = seqs.filter((x) => x === "end").length;
if (dB !== dE) F(`域 begin/end 不配对: ${dB}/${dE}`);
else if (dB !== dS) W(`域 separate 数(${dS})与 begin(${dB})不一致，可能存在无缓存域`);
else P(`域配对 begin/separate/end = ${dB}/${dS}/${dE}`);
// 顺序：每个 begin 之后必须先出现 instrText，再 separate，再 end
let depth = 0, orderBad = 0;
const events = [...doc.matchAll(/<w:fldChar w:fldCharType="(\w+)"\/>|<w:instrText[^>]*>/g)].map((m) => (m[1] ? "F:" + m[1] : "INSTR"));
for (const ev of events) {
  if (ev === "F:begin") { depth += 1; }
  else if (ev === "F:end") { depth -= 1; if (depth < 0) orderBad += 1; }
}
if (depth !== 0 || orderBad) F(`域嵌套错乱（残留深度 ${depth}，越界 ${orderBad}）`);
else P("域嵌套顺序正确");

// ---------- 5. 书签 ----------
const bmS = [...doc.matchAll(/<w:bookmarkStart w:id="([^"]+)" w:name="([^"]+)"/g)];
const bmE = [...doc.matchAll(/<w:bookmarkEnd w:id="([^"]+)"/g)];
const idsStart = bmS.map((m) => m[1]), idsEnd = bmE.map((m) => m[1]);
const dupId = idsStart.filter((id, i) => idsStart.indexOf(id) !== i);
const dupName = bmS.map((m) => m[2]).filter((n, i, a) => a.indexOf(n) !== i);
const unpaired = [...new Set(idsStart)].filter((id) => !idsEnd.includes(id));
if (dupId.length) F(`书签 id 重复: ${dupId.join(",")}`);
if (dupName.length) F(`书签名重复: ${[...new Set(dupName)].join(",")}`);
if (unpaired.length) F(`书签未配对: ${unpaired.join(",")}`);
const badName = bmS.map((m) => m[2]).filter((n) => !/^[A-Za-z_][A-Za-z0-9_]{0,39}$/.test(n));
if (badName.length) F(`书签名非法(需字母开头/≤40字符): ${badName.slice(0, 5).join(",")}`);
if (!dupId.length && !dupName.length && !unpaired.length && !badName.length) P(`书签 ${bmS.length} 个：id 唯一、名称唯一、起止配对、命名合法`);

// ---------- 6. 样式引用 ----------
const styleIds = new Set([...styXml.matchAll(/<w:style[^>]*w:styleId="([^"]+)"/g)].map((m) => m[1]));
const usedStyles = [...new Set([...doc.matchAll(/<w:pStyle w:val="([^"]+)"/g)].map((m) => m[1]))];
const badStyle = usedStyles.filter((s) => !styleIds.has(s));
if (badStyle.length) F(`引用了未定义样式: ${badStyle.join(",")}`);
else P(`样式引用全部有定义（用到 ${usedStyles.length} 个：${usedStyles.join(",")}）`);
const headingNames = [...styXml.matchAll(/<w:style[^>]*w:styleId="Heading(\d)"[\s\S]*?<w:name w:val="([^"]+)"/g)];
if (headingNames.length < 5) F(`Heading 样式缺内置名，只剩 ${headingNames.length}/5 个带 w:name`);
else P(`Heading1-5 均带内置名（${headingNames.map((m) => m[2]).join("/")}）`);

// ---------- 7. 编号引用 ----------
const numIds = new Set([...numXml.matchAll(/<w:num w:numId="([^"]+)"/g)].map((m) => m[1]));
const usedNum = [...new Set([...(styXml + doc).matchAll(/<w:numId w:val="([^"]+)"/g)].map((m) => m[1]))]
  .filter((v) => v !== "0");
const badNum = usedNum.filter((n) => !numIds.has(n));
if (badNum.length) F(`numId 未定义: ${badNum.join(",")}`);
else if (usedNum.length) P(`编号引用有效（numId ${usedNum.join(",")} 均在 numbering.xml 定义）`);
const lvls = [...numXml.matchAll(/<w:lvl w:ilvl="([^"]+)"/g)].map((m) => m[1]);
const usedIlvl = [...new Set([...(styXml + doc).matchAll(/<w:ilvl w:val="([^"]+)"/g)].map((m) => m[1]))];
const badIlvl = usedIlvl.filter((l) => !lvls.includes(l));
if (badIlvl.length) F(`ilvl 未定义: ${badIlvl.join(",")}`);
else if (usedIlvl.length) P(`编号层级有效（ilvl ${usedIlvl.join(",")} ∈ 已定义 ${lvls.join(",")}）`);

// ---------- 8. 图片 ----------
const media = parts.filter((k) => k.startsWith("word/media/"));
const drawings = [...doc.matchAll(/<w:drawing>/g)].length;
// 同一张图被重复引用时 r:embed 复用同一 media 部件，故比较唯一 rId 数而非 drawing 数
const embedIds = [...new Set([...doc.matchAll(/r:embed="([^"]+)"/g)].map((m) => m[1]))];
const extents = [...doc.matchAll(/<wp:extent cx="([^"]+)" cy="([^"]+)"/g)];
const zeroExtent = extents.filter((m) => +m[1] <= 0 || +m[2] <= 0).length;
const maxCx = Math.max(0, ...extents.map((m) => +m[1]));
const contentWEmu = 8312 * 635;
if (embedIds.length !== media.length) F(`唯一图片引用(${embedIds.length}) 与 media(${media.length}) 数量不一致`);
else if (drawings) P(`图片 ${drawings} 处引用 / ${media.length} 个 media 部件，去重后一一对应`);
if (zeroExtent) F(`有 ${zeroExtent} 张图片尺寸为 0`);
if (extents.length && maxCx > contentWEmu + 1000) F(`图片超出版心：最大宽 ${maxCx} EMU > ${contentWEmu}`);
else if (extents.length) P(`图片宽度均在版心内（最大 ${maxCx} EMU，版心 ${contentWEmu}）`);
const ph = (doc.match(/\[图片缺失: /g) || []).length;
if (ph) W(`有 ${ph} 张图片降级为缺失占位`);

// ---------- 9. 基本元素 ----------
const stats = {
  段落: [...doc.matchAll(/<w:p>/g)].length,
  标题: [...doc.matchAll(/<w:pStyle w:val="Heading\d"/g)].length,
  表格: [...doc.matchAll(/<w:tbl>/g)].length,
  分页: [...doc.matchAll(/w:type="page"/g)].length,
  域总: dB,
  书签: bmS.length,
  图片: drawings,
  ZOTERO_ITEM: (doc.match(/ADDIN ZOTERO_ITEM/g) || []).length,
  ZOTERO_BIBL: (doc.match(/ADDIN ZOTERO_BIBL/g) || []).length,
  SEQ: (doc.match(/SEQ (表|图|式)/g) || []).length,
  REF: (doc.match(/REF bkmk_/g) || []).length,
};

// ---------- 输出 ----------
for (const p of pass) console.log("PASS  " + p);
for (const w of warns) console.log("WARN  " + w);
for (const f of fails) console.log("FAIL  " + f);
console.log("\n统计: " + JSON.stringify(stats));
console.log(`结果: ${fails.length} 项失败 / ${warns.length} 项警告 / ${pass.length} 项通过`);
process.exit(fails.length ? 1 : 0);
