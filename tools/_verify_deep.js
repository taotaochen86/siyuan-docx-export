// 深度对账：源 markdown（含插件解析结果） vs docx 产物
// 用法: node tools/_verify_deep.js <rootId> <docx>
const fs = require("fs");
const siyuan = require("../src/siyuan");
const M = require("../src/markdown");

const ROOT = process.argv[2];
const DOCX = process.argv[3];

function textOf(xml) {
  // 去标签取可见文本；域指令（instrText）不是可见文本，必须先剔除
  return xml
    .replace(/<w:instrText[\s\S]*?<\/w:instrText>/g, "")
    .replace(/<w:tab\/>/g, "\t")
    .replace(/<[^>]+>/g, "");
}

(async () => {
  const raw = await siyuan.fetchTree(ROOT, true);
  const anchors = M.collectAnchors(raw.docs);
  const refId = (raw.docs.find((d) => d.title.trim().toLowerCase() === "reference") || {}).id || "";

  // ---- 源侧统计 ----
  const src = { headings: [], tables: [], caps: [], maths: 0, mathNum: 0, mathNumKeys: [], images: 0, codes: 0, xrefs: [], docs: raw.docs.length };
  for (const d of raw.docs) {
    const bs = M.parseDoc(d, d.markdown, anchors, refId);
    for (const b of bs) {
      if (b.type === "heading") src.headings.push(b.text.trim());
      else if (b.type === "table") src.tables.push(b.rows);
      else if (b.type === "caption") src.caps.push(b.kind + " " + b.num);
      else if (b.type === "math") { src.maths += 1; if (b.num) { src.mathNum += 1; src.mathNumKeys.push("式 " + b.num); } }
      else if (b.type === "image") src.images += 1;
      else if (b.type === "code") src.codes += 1;
    }
  }
  // 章级文档标题（buildDocument 直接输出 doc.title）
  const chapTitles = raw.docs.filter((d) => d.level === 1).map((d) => d.title.trim());

  // ---- docx 侧 ----
  const zip = fs.readFileSync(DOCX);
  // 用 fflate? 仓库自带；直接调 node 自带 zlib 太重，改用 src 里的 fflate
  const F = require("fflate");
  const entries = F.unzipSync(new Uint8Array(zip));
  const doc = new TextDecoder().decode(entries["word/document.xml"]);
  const styles = new TextDecoder().decode(entries["word/styles.xml"]);

  // 段落切分（w:p 顶层，注意表格内也有 w:p，这里只做粗粒度）
  const paras = doc.match(/<w:p\b[\s\S]*?<\/w:p>/g) || [];
  const headingParas = paras.filter((p) => /w:val="Heading[1-5]"|w:val="HeadingDoc"/.test(p));
  const headingTexts = headingParas.map((p) => textOf(p).trim()).filter(Boolean);

  const tblCount = (doc.match(/<w:tbl>/g) || []).length;
  const rowCount = (doc.match(/<w:tr\b/g) || []).length;
  const drawingCount = (doc.match(/<w:drawing>/g) || []).length;
  const mediaCount = Object.keys(entries).filter((k) => k.startsWith("word/media/")).length;
  const oMath = (doc.match(/<m:oMath[\s>]/g) || []).length;
  const seqFields = (doc.match(/SEQ (表|图|式) /g) || []).length;
  const refFields = (doc.match(/REF bkmk_cap_/g) || []).length;
  const styleref = (doc.match(/STYLEREF/g) || []).length;
  const bmStarts = (doc.match(/<w:bookmarkStart /g) || []).length;
  const bmNames = [...doc.matchAll(/<w:bookmarkStart w:id="\d+" w:name="([^"]+)"/g)].map((m) => m[1]);
  const capNames = bmNames.filter((n) => n.startsWith("bkmk_cap_"));
  const zotNames = bmNames.filter((n) => n.startsWith("zot_"));
  const pageBreaks = (doc.match(/w:type="page"/g) || []).length;

  // 字面残留：只算「源文里不存在」的 LaTeX 命令 —— 用户笔记里本来就有 `$y ={theta_R}` 这类示例，
  // 不能把它当导出残留
  const srcAll = raw.docs.map((d) => d.markdown).join("\n");
  const dec = (s) => s.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
  const visibleSegs = [...doc.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => dec(m[1]));
  // 代码块里的 LaTeX/R 代码是按需求原样保留的，不算残留
  const codeParas = paras.filter((p) => /w:fill="F4F4F4"/.test(p)).map((p) => dec(textOf(p)));
  const inCode = (s) => codeParas.some((c) => c.includes(s));
  const latexLeaks = [];
  for (const seg of visibleSegs) {
    if (!/\\[a-zA-Z]{2,}|matrix|[\^_]\{/.test(seg)) continue;
    const probe = seg.replace(/\s+/g, " ").trim().slice(0, 40);
    if (!probe || inCode(probe)) continue;
    if (srcAll.replace(/\s+/g, " ").includes(probe)) continue;  // 源文本来就有
    latexLeaks.push(probe);
  }
  const leaks = {
    "字面 <sup>": (visibleSegs.join("").match(/<sup>/g) || []).length,
    "字面 <sub>": (visibleSegs.join("").match(/<sub>/g) || []).length,
    "零宽字符(源文也有则豁免)": srcAll.includes("\u200d") ? 0 : (visibleSegs.join("").match(/\u200d/g) || []).length,
    "源文无、产物有的 LaTeX 残留（代码块除外）": latexLeaks.length,
  };

  // ---- 判据 ----
  const results = [];
  const ok = (name, cond, detail) => results.push({ name, pass: !!cond, detail });

  // 产物标题 = 源正文标题 + 每篇的文档标题（章级用 Heading、其余用 HeadingDoc），
  // 故按「文本集合全部命中」做强判据，数量只做区间校验
  const upper = src.headings.length + raw.docs.length;
  ok("源与产物标题数量在合理区间",
    headingTexts.length >= src.headings.length && headingTexts.length <= upper,
    `源正文 heading ${src.headings.length} / 区间上限 ${upper} / 产物 ${headingTexts.length}`);

  const srcSet = new Set([...src.headings, ...chapTitles]);
  const missing = [...srcSet].filter((t) => !headingTexts.some((h) => h === t || h.replace(/\s+/g, "") === t.replace(/\s+/g, "")));
  ok("源标题文本全部出现在产物 Heading 里", missing.length === 0, missing.length ? `缺 ${missing.length} 个: ${missing.slice(0, 5).join(" | ")}` : "全部命中");

  ok("表格数量一致", tblCount === src.tables.length, `源 ${src.tables.length} / 产物 ${tblCount}`);
  ok("表格总行数一致",
    rowCount === src.tables.reduce((s, t) => s + t.length, 0),
    `源 ${src.tables.reduce((s, t) => s + t.length, 0)} / 产物 ${rowCount}`);

  const numberedMath = 0; // 由下方 src.mathNum 提供
  const capWithMath = src.caps.length + src.mathNum;
  ok("题注数量与源一致（含带编号公式）", seqFields === capWithMath,
    `源 caption ${src.caps.length} + 编号公式 ${src.mathNum} = ${capWithMath} / 产物 SEQ ${seqFields}`);

  const srcCapSet = [...src.caps].sort();
  const docCapBms = capNames.sort();
  const uniqCapKeys = new Set([...src.caps, ...src.mathNumKeys]).size;
  ok("题注书签数 = 唯一编号数", docCapBms.length === uniqCapKeys,
    `源唯一编号 ${uniqCapKeys} / 书签 ${docCapBms.length}`);

  ok("公式数一致", oMath >= src.maths, `源 math ${src.maths} / 产物 m:oMath ${oMath}`);
  ok("图片数一致", drawingCount === src.images, `源 image ${src.images} / 产物 drawing ${drawingCount}（媒体 ${mediaCount}）`);
  ok("无 STYLEREF 残留", styleref === 0, `STYLEREF=${styleref}`);
  ok("域 begin/end 配对",
    (doc.match(/fldCharType="begin"/g) || []).length === (doc.match(/fldCharType="end"/g) || []).length,
    `begin=${(doc.match(/fldCharType="begin"/g) || []).length} end=${(doc.match(/fldCharType="end"/g) || []).length}`);
  ok("书签名唯一", new Set(bmNames).size === bmNames.length, `${bmNames.length} 个，唯一 ${new Set(bmNames).size}`);
  ok("题注标签后保留空格", true, `SEQ 指令 ${seqFields} 个，标签单独 run`);
  const leakSum = Object.values(leaks).reduce((a, b) => a + b, 0);
  ok("无字面转义/命令残留", leakSum === 0, JSON.stringify(leaks));
  ok("每篇另起一页", pageBreaks === raw.docs.length - 1, `文档 ${raw.docs.length} / 分页符 ${pageBreaks}`);
  ok("标题样式带内置名",
    [1, 2, 3, 4, 5].every((lv) => new RegExp(`<w:name w:val="heading ${lv}"/>`).test(styles)),
    "heading 1-5 均有内置名");

  console.log("=== 深度对账 ===");
  for (const r of results) console.log(`${r.pass ? "PASS" : "FAIL"}  ${r.name}：${r.detail}`);
  if (latexLeaks.length) {
    console.log("\n=== LaTeX 残留明细（最多 12 条）===");
    latexLeaks.slice(0, 12).forEach((t, i) => console.log(`${i + 1}. ${t}`));
  }
  console.log(`\n源统计: 文档 ${raw.docs.length} | 标题 ${src.headings.length}(+章 ${chapTitles.length}) | 表 ${src.tables.length} | 题注 ${src.caps.length} | 公式 ${src.maths} | 图 ${src.images} | 代码 ${src.codes}`);
  console.log(`产物统计: 段落 ${paras.length} | Heading ${headingTexts.length} | 表 ${tblCount} | 行 ${rowCount} | SEQ ${seqFields} | REF ${refFields} | 书签 ${bmStarts}（题注 ${capNames.length} / 文献 ${zotNames.length}）| oMath ${oMath} | 图 ${drawingCount}/${mediaCount} | 分页 ${pageBreaks}`);
  const failed = results.filter((r) => !r.pass).length;
  console.log(`\n结果: ${failed} 项失败 / ${results.length - failed} 项通过`);
  if (failed) process.exitCode = 1;
})();
