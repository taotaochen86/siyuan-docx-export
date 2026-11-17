// 把解析后的文档渲染为 OOXML 并打包成 .docx
// 对接三个插件：tree-doc-numbering（标题自动编号）、siyuan-xref（图表/公式题注与交叉引用）、
// siyuan-plugin-zotero-cite（Zotero 文内引用域 + 书目域）
const D = require("./docx");
const M = require("./markdown");

const HEADING_SIZES = { 1: 16, 2: 14, 3: 13, 4: 12, 5: 12 };
const CONTENT_W = 8312; // A4 21cm - 左右各 3.17cm，单位 twips

function para(opts, inner) {
  return `<w:p>${D.pPr(opts)}${inner || ""}</w:p>`;
}

function pageBreakPara() {
  return `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;
}

// ---------------- 标题 ----------------
// 从「第N章」解析真实章号（阿拉伯或中文数字）；解析不到返回 null
const CN_DIGIT = { 零: 0, 一: 1, 二: 2, 三: 3, 四: 4, 五: 5, 六: 6, 七: 7, 八: 8, 九: 9 };
function cnNum(s) {
  if (/^[0-9]+$/.test(s)) return parseInt(s, 10);
  if (s.length === 1 && CN_DIGIT[s] != null) return CN_DIGIT[s];
  const m = /^([一二三四五六七八九])?十([一二三四五六七八九])?$/.exec(s);
  if (m) return (m[1] ? CN_DIGIT[m[1]] : 1) * 10 + (m[2] ? CN_DIGIT[m[2]] : 0);
  return null;
}
function parseChapterNo(text) {
  const m = /第\s*([0-9]+|[一二三四五六七八九十]+)\s*章/.exec(text);
  return m ? cnNum(m[1]) : null;
}

// level 1 -> Heading1（标题编号是思源静态文本，只套样式）
// 章标题段末追加一个隐藏的 `SEQ 章 \r N` 计数器：题注用 `SEQ 章 \c` 取章号，
// 这样「表 4-1」是整段域、F9 后章号仍是 4（每章显式重置，不会因为缺章被重排）。
function renderHeading(text, level, strip) {
  const lv = Math.min(Math.max(level, 1), 5);
  const size = HEADING_SIZES[lv] || 12;
  let clean = M.stripHtml(text).replace(/[\u200b-\u200d\ufeff]/g, "").replace(/\*\*/g, "").trim();
  if (strip) clean = M.stripHeadingNum(clean);
  if (!clean) return "";
  // 标题编号保留思源原文（静态文本），只套 Heading 样式，不绑 Word 列表。
  let inner = D.runXml(clean, { cn: D.CN_HEI, en: D.CN_HEI, size, bold: true, color: "000000" });
  if (lv === 1) {
    const ch = parseChapterNo(clean);
    // 缓存文本留空：更新前后都不产生可见字符，标题里不会多出数字
    inner += D.fieldRuns(ch != null ? `SEQ 章 \\* ARABIC \\r ${ch}` : "SEQ 章 \\* ARABIC", "", { size, hidden: true });
  }
  return para(
    {
      style: "Heading" + lv,
      jc: "left",
      spacing: { before: lv <= 2 ? 200 : 120, after: 120, line: 240, lineRule: "auto" },
    },
    inner
  );
}

// 不参与自动编号的大标题（节级文档、Reference 等）
function renderDocHeading(text, strip) {
  let clean = M.stripHtml(text).replace(/[\u200b-\u200d\ufeff]/g, "").replace(/\*\*/g, "").trim();
  if (strip) clean = M.stripHeadingNum(clean);
  if (!clean) return "";
  return para(
    { style: "HeadingDoc", jc: "left", spacing: { before: 240, after: 160, line: 240, lineRule: "auto" } },
    D.runXml(clean, { cn: D.CN_HEI, en: D.CN_HEI, size: 16, bold: true, color: "000000" })
  );
}

// ---------------- 正文 ----------------
function renderBody(text, refsByKey, listItem, ctx) {
  if (M.stripHtml(text).replace(/\u200d/g, "").trim() === "") return "";
  const opts = { spacing: { line: 360, lineRule: "auto", after: 80 } };
  if (listItem) {
    opts.ind = { left: 420, hanging: 210 };
  } else {
    opts.ind = { firstLineChars: 200, firstLine: 480 };
  }
  let inner = listItem ? D.runXml("• ", { size: 12 }) : "";
  inner += D.buildInline(text, refsByKey, 12, false, false, refsByKey, ctx ? ctx.imgHook : null);
  return para(opts, inner);
}

function renderCode(code) {
  const lines = code.split("\n");
  let inner = "";
  lines.forEach((ln, i) => {
    inner += D.runXml(ln === "" ? " " : ln, { size: 9, mono: true });
    if (i < lines.length - 1) inner += "<w:r><w:br/></w:r>";
  });
  return para(
    { spacing: { line: 276, lineRule: "auto", before: 80, after: 80 }, ind: { left: 170 }, shd: "F4F4F4" },
    inner
  );
}

// 公式片段：优先 Word 原生 OMML；解析为空回退轻量文本 run
function mathFragment(code) {
  return D.mathOmml(code) || D.mathXml(code, { size: 12 });
}

// 独立公式：有编号走「居中 + 右对齐编号(SEQ 式 域 + 书签)」，无编号居中
function renderMath(code, num, anchor, bmId, ctx) {
  if (!num) {
    return para(
      { jc: "center", spacing: { before: 80, after: 80 } },
      mathFragment(code)
    );
  }
  const ppr = `<w:pPr><w:tabs><w:tab w:val="center" w:pos="${Math.round(CONTENT_W / 2)}"/>` +
    `<w:tab w:val="right" w:pos="${CONTENT_W}"/></w:tabs>` +
    `<w:spacing w:before="80" w:after="80"/></w:pPr>`;
  const key = D.capKey("式", num);
  let name = null;
  if (!ctx || !ctx.bmNames || !ctx.bmNames.has(key)) {
    if (ctx && ctx.bmNames) ctx.bmNames.add(key);
    name = D.capBmName("式", num);
  }
  let inner = `<w:r><w:tab/></w:r>` + mathFragment(code);
  inner += `<w:r><w:tab/></w:r>`;
  if (name) inner += D.bookmarkStart(bmId, name);
  const { chap, seq } = splitChapSeq(num);
  inner += D.runXml("(", { size: 12 });
  if (chap) {
    // 同题注：章号取章标题里隐藏的 `SEQ 章 \r N`，整段编号都是域
    inner += D.fieldRuns(`SEQ 章 \\c`, chap, { size: 12 });
    inner += D.runXml("-", { size: 12 });
  }
  inner += D.fieldRuns(`SEQ \u5f0f \\* ARABIC \\s 1`, seq, { size: 12 });
  inner += D.runXml(")", { size: 12 });
  if (name) inner += D.bookmarkEnd(bmId);
  return `<w:p>${ppr}${inner}</w:p>`;
}

const EMU_PX = 9525;                              // 96dpi：1px = 9525 EMU
const MAX_W_EMU = Math.round(CONTENT_W * 635 * 0.98); // 版心宽（向下留 2% 余量）
const MIME_BY_EXT = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", webp: "image/webp" };

function baseName(url) {
  return String(url || "").split("?")[0].split("#")[0].split("/").pop();
}

// 图片：有二进制走真嵌入（drawing + rels + media），取不到降级占位文本
function imageXml(url, alt, ctx, size) {
  const a = ctx && ctx.assets && ctx.assets[url];
  if (!a) {
    return D.runXml("[图片缺失: " + baseName(url) + "]", { size: size || 10.5, italic: true });
  }
  let idx = ctx.media.findIndex((m) => m.key === url);
  if (idx < 0) {
    const name = "image" + (ctx.media.length + 1) + "." + a.ext;
    ctx.media.push({ key: url, asset: a, name });
    idx = ctx.media.length - 1;
  }
  const m = ctx.media[idx];
  const rid = "rIdImg" + (idx + 1);
  let w, h;
  if (a.w && a.h) {
    w = a.w * EMU_PX;
    h = a.h * EMU_PX;
    if (w > MAX_W_EMU) { h = Math.round(h * MAX_W_EMU / w); w = MAX_W_EMU; }
  } else {
    w = Math.round(MAX_W_EMU * 0.6);
    h = Math.round(w * 0.75);
  }
  return D.drawingXml(rid, Math.round(w), Math.round(h), idx + 1, m.name);
}

function renderImage(rel, ctx) {
  return para(
    { jc: "center", spacing: { before: 80, after: 40 } },
    imageXml(rel, "", ctx, 12)
  );
}

// 拆「4-1」-> 章 4、序 1；无连字符时章号为空、整段为序号
function splitChapSeq(num) {
  const i = String(num).indexOf("-");
  return i > 0 ? { chap: num.slice(0, i), seq: num.slice(i + 1) } : { chap: "", seq: num };
}

function renderCaption(kind, num, title, anchor, bmId, ctx) {
  let inner = "";
  let start = "", end = "";
  // 书签按「标签+编号」命名并全局去重：同一编号在多篇文档里只建一次，
  // 交叉引用按同一编号命中，跨文档也不会撞名、不会「未定义书签」。
  const key = D.capKey(kind, num);
  if (!ctx || !ctx.bmNames || !ctx.bmNames.has(key)) {
    if (ctx && ctx.bmNames) ctx.bmNames.add(key);
    start = D.bookmarkStart(bmId, D.capBmName(kind, num));
    end = D.bookmarkEnd(bmId);
  }
  // 书签只覆盖「表/图/式 + 编号」，题名放到书签外。
  // 否则 F9 更新 REF 域时会把书签内全文（含题名）整条塞进正文。
  const { chap, seq } = splitChapSeq(num);
  inner += start;
  inner += D.runXml(kind + " ", { size: 12, bold: true });
  if (chap) {
    // 章号是域：取本章标题里隐藏的 `SEQ 章 \r N`（`\c` = 重复最近序号）。
    // 「表 4-1」整段都是域结果，F9 后章号仍精确等于思源原章号。
    inner += D.fieldRuns(`SEQ 章 \\c`, chap, { size: 12, bold: true });
    inner += D.runXml("-", { size: 12, bold: true });
  }
  inner += D.fieldRuns(`SEQ ${kind} \\* ARABIC \\s 1`, seq, { size: 12, bold: true });
  inner += end;
  if (title) inner += D.runXml(" " + title, { size: 12, bold: true });
  return para(
    { jc: "center", spacing: { before: 80, after: 80 } },
    inner
  );
}

function threeLineBorders() {
  const b = (tag, val, sz) =>
    `<w:${tag} w:val="${val}"${val !== "nil" ? ` w:sz="${sz}" w:space="0" w:color="000000"` : ""}/>`;
  return "<w:tblBorders>" +
    b("top", "single", 12) + b("bottom", "single", 12) +
    b("left", "nil") + b("right", "nil") +
    b("insideH", "nil") + b("insideV", "nil") + "</w:tblBorders>";
}

function renderTable(rows, refsByKey, ctx) {
  // 列数按真实内容取，不能设下限 3：思源里 2 列的表照 2 列出，否则会多出一条空列
  const ncol = Math.max(1, rows.reduce((m, r) => Math.max(m, r.length), 1));
  let grid = "";
  for (let c = 0; c < ncol; c++) grid += `<w:gridCol w:w="${Math.floor(9000 / ncol)}"/>`;

  let trs = "";
  rows.forEach((row, ri) => {
    let tcs = "";
    for (let ci = 0; ci < ncol; ci++) {
      const cell = ci < row.length ? row[ci] : "";
      const cellBorder = ri === 0
        ? `<w:tcBorders><w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/></w:tcBorders>` : "";
      const cellInner = para(
        { spacing: { line: 276, lineRule: "auto" } },
        D.buildInline(cell, refsByKey, 10.5, ri === 0, false, refsByKey, ctx ? ctx.imgHook : null)
      );
      tcs += `<w:tc><w:tcPr><w:tcW w:w="${Math.floor(9000 / ncol)}" w:type="dxa"/>${cellBorder}</w:tcPr>${cellInner}</w:tc>`;
    }
    trs += `<w:tr>${tcs}</w:tr>`;
  });

  const tbl = "<w:tbl>" +
    `<w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:jc w:val="center"/>${threeLineBorders()}</w:tblPr>` +
    `<w:tblGrid>${grid}</w:tblGrid>${trs}</w:tbl>`;
  return tbl + para({ spacing: { after: 40 } }, "");
}

function renderReference(entries, refsByKey, bmIdStart) {
  let bibItems = [];
  entries.forEach((e, i) => {
    if (refsByKey && refsByKey[e.key] && refsByKey[e.key].itemData) {
      bibItems.push({ id: i + 1, uris: [`http://zotero.org/users/local/1/items/${e.key}`], itemData: refsByKey[e.key].itemData });
    }
  });
  const biblInstr = "ADDIN ZOTERO_BIBL " + JSON.stringify({ citationItems: bibItems, properties: {} });

  let body = "";
  entries.forEach((e, i) => {
    const bm = (refsByKey && refsByKey[e.key]) ? D.bookmarkStart(bmIdStart + i, "zot_" + e.key) + D.bookmarkEnd(bmIdStart + i) : "";
    const inner = D.runXml(e.order + ". ", { size: 12 }) +
      D.buildInline(e.text, refsByKey, 12, false, false, refsByKey);
    body += para(
      { spacing: { line: 360, lineRule: "auto", after: 80 }, ind: { left: 420, hanging: 420 } },
      bm + inner
    );
  });

  // 跨段域：begin/instr/separate 插首段最前，end 末段最后
  const begin = `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve"> ${D.esc(biblInstr)} </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>`;
  const end = `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
  body = body.replace("<w:p>", "<w:p>" + begin, 1);
  body = body.replace(/<\/w:p>(?![\s\S]*<w:p>)/, end + "</w:p>");
  return body;
}

function renderBlocks(blocks, refsByKey, bmId, ctx) {
  let out = "";
  let bm = (ctx && ctx.bm != null) ? ctx.bm : bmId;
  let n = -1;
  while (++n < blocks.length) {
    const b = blocks[n];
    switch (b.type) {
      case "heading": {
        out += renderHeading(b.text, b.level, ctx && ctx.stripHeadingNum);
        if (ctx && b.level === 1) ctx.seenH1 = true;
        break;
      }
      case "body": out += renderBody(b.text, refsByKey, b.listItem, ctx); break;
      case "code": out += renderCode(b.code); break;
      case "math": out += renderMath(b.code, b.num, b.anchor, bm, ctx); bm += 1; break;
      case "image": out += renderImage(b.rel, ctx); break;
      case "table": out += renderTable(b.rows, refsByKey, ctx); break;
      case "caption": out += renderCaption(b.kind, b.num, b.title, b.anchor, bm, ctx); bm += 1; break;
      case "reference": out += renderReference(b.entries, refsByKey, bm); bm += b.entries.length; break;
      default: break;
    }
  }
  // 书签 id 必须全篇唯一：计数器挂在 ctx 上跨文档累加，不能每篇从 1000 重来
  if (ctx) ctx.bm = bm;
  return out;
}

// ---------------- 样式 / 编号 ----------------
function stylesXml() {
  // 必须带 <w:name w:val="heading N"/>（Word 内置样式名，小写），
  // 否则 Word/WPS 只当它是匿名自定义样式（显示为 _Style N），导航窗格与大纲视图不认它是标题
  const head = (lv, sz, ilvl) => `
  <w:style w:type="paragraph" w:styleId="Heading${lv}">
    <w:name w:val="heading ${lv}"/>
    <w:basedOn w:val="Normal"/><w:next w:val="Normal"/>
    <w:uiPriority w:val="9"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:outlineLvl w:val="${lv - 1}"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="${D.CN_HEI}" w:hAnsi="${D.CN_HEI}" w:eastAsia="${D.CN_HEI}"/><w:b/><w:color w:val="000000"/><w:sz w:val="${sz * 2}"/></w:rPr>
  </w:style>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr>
    <w:rFonts w:ascii="${D.EN_FONT}" w:hAnsi="${D.EN_FONT}" w:eastAsia="${D.CN_SONG}"/>
    <w:sz w:val="24"/>
  </w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:rPr><w:rFonts w:ascii="${D.EN_FONT}" w:hAnsi="${D.EN_FONT}" w:eastAsia="${D.CN_SONG}"/><w:sz w:val="24"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="HeadingDoc">
    <w:name w:val="文档标题"/>
    <w:basedOn w:val="Normal"/><w:next w:val="Normal"/>
    <w:uiPriority w:val="9"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="${D.CN_HEI}" w:hAnsi="${D.CN_HEI}" w:eastAsia="${D.CN_HEI}"/><w:b/><w:color w:val="000000"/><w:sz w:val="32"/></w:rPr>
  </w:style>
  ${head(1, 16)}${head(2, 14)}${head(3, 13)}${head(4, 12)}${head(5, 12)}
</w:styles>`;
}

// 标题编号保留思源写死的静态文本，不再使用 Word 多级列表；numbering.xml 仅留空壳以兼容打包链路
function numberingXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`;
}

function buildDocument(raw, opts, refsByKey, assets) {
  const docs = raw.docs;
  const ctx = { assets: assets || {}, media: [], bm: 1000, seenH1: false, stripHeadingNum: !!opts.stripHeadingNum };
  ctx.imgHook = (alt, url) => imageXml(url, alt, ctx, 10.5);
  const refId = opts.referenceDocId ||
    (docs.find((d) => d.title.trim().toLowerCase() === "reference") || {}).id || "";
  const anchors = M.collectAnchors(docs);

  // 预扫描：先把所有题注/公式编号收齐，交叉引用才能判断「目标题注是否存在」，
  // 不存在的引用降级为普通文本，避免 F9 后出现「错误！未定义书签。」
  ctx.capKeys = new Set();
  ctx.bmNames = new Set();
  const blocksByDoc = docs.map((d) => M.parseDoc(d, d.markdown, anchors, refId));
  for (const bs of blocksByDoc) {
    for (const b of bs) {
      if (b.type === "caption") ctx.capKeys.add(D.capKey(b.kind, b.num));
      else if (b.type === "math" && b.num) ctx.capKeys.add(D.capKey("式", b.num));
    }
  }
  const refsInline = Object.assign({}, refsByKey || {});
  refsInline.__caps = ctx.capKeys;

  let body = "";

  // 封面
  if (opts.coverTitle) {
    body += para({ jc: "center", spacing: { before: 2268, after: 0 } },
      D.runXml(opts.coverTitle, { cn: D.CN_HEI, en: D.CN_HEI, size: 26, bold: true }));
    for (let k = 0; k < 6; k++) body += para({}, "");
    if (opts.coverAuthor) body += para({ jc: "center" }, D.runXml("作　　者：" + opts.coverAuthor, { size: 14 }));
    if (opts.coverDate) body += para({ jc: "center" }, D.runXml("发布日期：" + opts.coverDate, { size: 14 }));
    body += pageBreakPara();
  }

  const renderDoc = (doc, blocks) => renderBlocks(blocks, refsInline, ctx.bm, ctx);

  body += renderDoc(docs[0], blocksByDoc[0]);
  for (let i = 1; i < docs.length; i++) {
    const doc = docs[i];
    body += pageBreakPara();
    if (doc.level === 1 && doc.id !== refId) {
      // 章级文档：标题原样保留「第N章 题名」（与思源一致的静态编号），只套 Heading1
      body += renderHeading(doc.title, 1, ctx.stripHeadingNum);
      ctx.seenH1 = true;
    } else {
      body += renderDocHeading(doc.title, ctx.stripHeadingNum);
    }
    body += renderDoc(doc, blocksByDoc[i]);
  }

  const sect = `<w:sectPr>
    <w:pgSz w:w="11906" w:h="16838"/>
    <w:pgMar w:top="1440" w:right="1797" w:bottom="1440" w:left="1797" w:gutter="0"/>
  </w:sectPr>`;

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <w:body>${body}${sect}</w:body>
</w:document>`;

  const exts = Array.from(new Set(ctx.media.map((m) => m.asset.ext)));
  const extDefaults = exts.map((e) => `<Default Extension="${e}" ContentType="${MIME_BY_EXT[e] || "image/png"}"/>`).join("\n  ");
  const imgRels = ctx.media.map((m, i) =>
    `<Relationship Id="rIdImg${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${m.name}"/>`
  ).join("\n  ");

  const files = {
    "[Content_Types].xml": enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  ${extDefaults}
</Types>`),
    "_rels/.rels": enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`),
    "word/document.xml": enc(documentXml),
    "word/styles.xml": enc(stylesXml()),
    "word/numbering.xml": enc(numberingXml()),
    "word/_rels/document.xml.rels": enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  ${imgRels}
</Relationships>`),
  };

  // 图片部件：word/media/imageN.ext
  ctx.media.forEach((m) => { files["word/media/" + m.name] = m.asset.bytes; });

  return D.fflate.zipSync(files, { deflate: D.fflate.deflateSync });
}

function enc(s) {
  return new TextEncoder().encode(s);
}

module.exports = { buildDocument };
