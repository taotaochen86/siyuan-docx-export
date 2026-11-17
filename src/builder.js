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
// level 1 -> Heading1（带「第N章」自动编号 + 书签级 outlineLvl 0）
// level 2..5 -> Heading2..5（一、/（一）/1./(1）自动编号）
// noNum: 该标题在章节结构里是"悬空"的（出现在全书第一个一级标题之前，如书名页后的「序言」），
//        用 numId=0 去掉编号。否则 Word 会把它隐式计入上级编号，导致其后所有章号 +1。
function renderHeading(text, level, noNum) {
  const lv = Math.min(Math.max(level, 1), 5);
  const size = HEADING_SIZES[lv] || 12;
  const clean = M.stripHtml(text).replace(/[\u200b-\u200d\ufeff]/g, "").replace(/\*\*/g, "").trim();
  if (!clean) return "";
  const xml = para(
    {
      style: "Heading" + lv,
      jc: "left",
      spacing: { before: lv <= 2 ? 200 : 120, after: 120, line: 240, lineRule: "auto" },
      ind: { left: 0, firstLine: 0 },
    },
    D.runXml(clean, { cn: D.CN_HEI, en: D.CN_HEI, size, bold: true, color: "000000" })
  );
  if (!noNum) return xml;
  // numId=0 是 OOXML 标准的"移除编号"写法；numPr 必须紧跟 pStyle 之后，保持 pPr 子元素顺序合法
  return xml.replace(/(<w:pStyle w:val="Heading\d"\/>)/,
    '$1<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>');
}

// 不参与自动编号的大标题（节级文档、Reference 等）
function renderDocHeading(text) {
  const clean = M.stripHtml(text).replace(/[\u200b-\u200d\ufeff]/g, "").replace(/\*\*/g, "").trim();
  if (!clean) return "";
  return para(
    { style: "HeadingDoc", jc: "left", spacing: { before: 240, after: 160, line: 240, lineRule: "auto" } },
    D.runXml(clean, { cn: D.CN_HEI, en: D.CN_HEI, size: 16, bold: true, color: "000000" })
  );
}

// 「第3章 AAAA」-> 「AAAA」：章号交给 Word 多级列表生成，避免重号
function stripChapterNum(title) {
  const out = String(title || "").replace(/^\s*第\s*[0-9一二三四五六七八九十百零两]+\s*[章]\s*/, "").trim();
  return out || String(title || "").trim();
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

// 独立公式：有编号走「居中 + 右对齐编号(SEQ 式 域 + 书签)」，无编号居中
function renderMath(code, num, anchor, bmId) {
  if (!num) {
    return para(
      { jc: "center", spacing: { before: 80, after: 80 } },
      D.mathXml(code, { size: 12 })
    );
  }
  const ppr = `<w:pPr><w:tabs><w:tab w:val="center" w:pos="${Math.round(CONTENT_W / 2)}"/>` +
    `<w:tab w:val="right" w:pos="${CONTENT_W}"/></w:tabs>` +
    `<w:spacing w:before="80" w:after="80"/></w:pPr>`;
  const name = anchor ? D.bkmkName(anchor) : null;
  let inner = `<w:r><w:tab/></w:r>` + D.mathXml(code, { size: 12 });
  inner += `<w:r><w:tab/></w:r>`;
  if (name) inner += D.bookmarkStart(bmId, name);
  const { chap, seq } = splitChapSeq(num);
  inner += D.runXml("(", { size: 12 });
  if (chap) {
    inner += D.fieldRuns('STYLEREF "Heading 1" \\s', chap, { size: 12 });
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

function renderCaption(kind, num, title, anchor, bmId) {
  let inner = "";
  let start = "", end = "";
  if (anchor) {
    const name = D.bkmkName(anchor);
    start = D.bookmarkStart(bmId, name);
    end = D.bookmarkEnd(bmId);
  }
  // 书签只覆盖「表/图/式 + 编号」，题名放到书签外。
  // 否则 F9 更新 REF 域时会把书签内全文（含题名）整条塞进正文。
  const { chap, seq } = splitChapSeq(num);
  inner += start;
  inner += D.runXml(kind + " ", { size: 12, bold: true });
  if (chap) {
    // 章号：STYLEREF 取最近 Heading 1 编号；序号：SEQ 按级别 1 每章重置
    inner += D.fieldRuns('STYLEREF "Heading 1" \\s', chap, { size: 12, bold: true });
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
        const orphan = ctx && !ctx.seenH1 && b.level > 1;
        out += renderHeading(b.text, b.level, orphan);
        if (ctx && b.level === 1) ctx.seenH1 = true;
        break;
      }
      case "body": out += renderBody(b.text, refsByKey, b.listItem, ctx); break;
      case "code": out += renderCode(b.code); break;
      case "math": out += renderMath(b.code, b.num, b.anchor, bm); bm += 1; break;
      case "image": out += renderImage(b.rel, ctx); break;
      case "table": out += renderTable(b.rows, refsByKey, ctx); break;
      case "caption": out += renderCaption(b.kind, b.num, b.title, b.anchor, bm); bm += 1; break;
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
    <w:pPr><w:keepNext/><w:outlineLvl w:val="${lv - 1}"/>
      <w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="1"/></w:numPr>
    </w:pPr>
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
  ${head(1, 16, 0)}${head(2, 14, 1)}${head(3, 13, 2)}${head(4, 12, 3)}${head(5, 12, 4)}
</w:styles>`;
}

// 多级列表：L1 第N章 / L2 一、/ L3 （一）/ L4 1. / L5 (1）
function numberingXml() {
  const lvl = (ilvl, fmt, text, sz) => `
    <w:lvl w:ilvl="${ilvl}">
      <w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:suff w:val="space"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="${D.CN_HEI}" w:hAnsi="${D.CN_HEI}" w:eastAsia="${D.CN_HEI}"/><w:b/><w:sz w:val="${sz * 2}"/></w:rPr>
    </w:lvl>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="multilevel"/>
    ${lvl(0, "decimal", "第%1章", 16)}
    ${lvl(1, "chineseCounting", "%2、", 14)}
    ${lvl(2, "chineseCounting", "（%3）", 13)}
    ${lvl(3, "decimal", "%4.", 12)}
    ${lvl(4, "decimal", "(%5）", 12)}
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>`;
}

function buildDocument(raw, opts, refsByKey, assets) {
  const docs = raw.docs;
  const ctx = { assets: assets || {}, media: [], bm: 1000, seenH1: false };
  ctx.imgHook = (alt, url) => imageXml(url, alt, ctx, 10.5);
  const refId = opts.referenceDocId ||
    (docs.find((d) => d.title.trim().toLowerCase() === "reference") || {}).id || "";
  const anchors = M.collectAnchors(docs);

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

  const renderDoc = (doc) => renderBlocks(M.parseDoc(doc, doc.markdown, anchors, refId), refsByKey, ctx.bm, ctx);

  body += renderDoc(docs[0]);
  for (let i = 1; i < docs.length; i++) {
    const doc = docs[i];
    body += pageBreakPara();
    if (doc.level === 1 && doc.id !== refId) {
      // 章级文档：Heading1 自动编号「第N章」，原文「第3章」前缀剥掉
      body += renderHeading(stripChapterNum(doc.title), 1);
      ctx.seenH1 = true;
    } else {
      body += renderDocHeading(doc.title);
    }
    body += renderDoc(doc);
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
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
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
