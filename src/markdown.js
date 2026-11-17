// Markdown 切块：规则平移自 siyuan-to-word/build_docx.py，并对接
// tree-doc-numbering（标题正文编号）、siyuan-xref（图表/公式题注与交叉引用）、
// siyuan-plugin-zotero-cite（zotero:// 引用与书目）三个插件的真实输出形态。

function stripHtml(s) {
  s = s.replace(/<span[^>]*>/g, "").replace(/<\/span>/g, "");
  return s.replace(/<[^>]+>/g, "");
}

function stripFrontmatter(md) {
  if (md.startsWith("---")) {
    const end = md.indexOf("\n---", 3);
    if (end !== -1) md = md.slice(end + 4);
  }
  return md.replace(/^\n+/, "");
}

function normTitle(s) {
  return stripHtml(s).replace(/[\\\s​]/g, "");
}

const LABELS = "表|图|式";

// 交叉引用锚点：[表 4-1](#blockId) / [图 2](#id) / [式 6-1](#id)
function collectAnchors(docs) {
  const m = {};
  const re = new RegExp("\\[((?:" + LABELS + ")\\s*[0-9]+(?:-[0-9]+)?)\\]\\(#([^)]+)\\)", "g");
  for (const doc of docs) {
    let mm;
    re.lastIndex = 0;
    while ((mm = re.exec(doc.markdown)) !== null) {
      m[mm[1].replace(/\s+/g, "")] = mm[2];
    }
  }
  return m;
}

// 题注：表 4-1 标题 / 图 4-1 标题 / 式(6-1) 说明
const CAPTION_RE = /^\s*(表|图|式)\s*[0-9]+(?:-[0-9]+)?\s+\S/;
function isCaption(line) { return CAPTION_RE.test(line); }

const REF_RE = /^\s*(\d+)\\?\.\s+(.*?)\s*\[↗\]\(zotero:\/\/[^)]*\/items\/([A-Z0-9]+)[^)]*\)\s*$/;

function parseReference(md) {
  const entries = [];
  for (const line of md.split("\n")) {
    const mm = REF_RE.exec(line);
    if (mm) entries.push({ order: mm[1], text: mm[2].trim(), key: mm[3] });
  }
  return entries;
}

function splitTableRow(line) {
  line = line.trim();
  if (line.startsWith("|")) line = line.slice(1);
  if (line.endsWith("|")) line = line.slice(0, -1);
  return line.split("|").map((c) => c.trim());
}

function isSepRow(cells) {
  const nonempty = cells.filter((c) => c !== "");
  if (!nonempty.length) return false;
  return nonempty.every((c) => /^:?-{2,}:?$/.test(c));
}

// 标题层级映射：文档标题 Heading1，正文 #..#### 落 Heading2..5
function mapLevel(docLevel, hashN) {
  return Math.min(hashN + 1, 5);
}

// 剥掉 tree-doc-numbering 写死在标题文本里的静态编号前缀（可选开关）。
// 覆盖：第N章/节/篇、（一）/(1）括号序号、一、中文顿号、1. / 3.1.4 阿拉伯点号。
// 只剥开头第一段编号，题名原样保留；题注编号（表/图/式）不走这里。
function stripHeadingNum(s) {
  if (!s) return s;
  let t = String(s).trim();
  const patterns = [
    /^第\s*[0-9一二三四五六七八九十百千]+\s*[章节篇]\s*/,
    /^[（(]\s*[0-9一二三四五六七八九十百千]+\s*[)）]\s*/,
    /^[一二三四五六七八九十百千]+\s*、\s*/,
    /^[0-9]+(?:[.、-][0-9]+)+[.、]?\s*/,  // 多级 3.1.4 / 3.1.4.40.1
    /^[0-9]+[.、]\s*/,                    // 单级 1. / 1、
  ];
  for (const re of patterns) {
    const m = re.exec(t);
    if (m) { t = t.slice(m[0].length); break; }
  }
  return t.trim();
}

// ---------- 公式 ----------
function splitMath(text) {
  let num = "";
  const m = /\\tag\{([^}]+)\}/.exec(text);
  if (m) {
    num = m[1].trim();
    text = text.replace(/\\tag\{[^}]*\}/g, "");
  }
  return { code: text.trim(), num };
}

function parseDoc(meta, md, anchors, refId) {
  md = stripFrontmatter(md);
  const blocks = [];

  if (meta.id === refId) {
    const entries = parseReference(md);
    if (entries.length) blocks.push({ type: "reference", entries });
    return blocks;
  }

  const lines = md.split("\n");
  let i = 0, n = lines.length, firstSkip = false;

  while (i < n) {
    const line = lines[i];
    const fm = /^```(\w*)\s*$/.exec(line.trim());
    if (fm) {
      const block = [];
      i += 1;
      while (i < n && !lines[i].trim().startsWith("```")) { block.push(lines[i]); i += 1; }
      i += 1;
      blocks.push({ type: "code", code: block.join("\n") });
      continue;
    }

    if (line.trim().startsWith("$$")) {
      const head = line.trim();
      let body;
      i += 1;
      // 单行完整公式 $$...$$：必须在这里就闭合，否则会把后面所有段落一路吃进公式里
      if (head.length > 4 && head.endsWith("$$")) {
        body = head.slice(2, -2);
      } else {
        const block = [head.replace(/^\$+/, "")];
        while (i < n && !lines[i].trim().endsWith("$$")) { block.push(lines[i]); i += 1; }
        if (i < n) block.push(lines[i].trim().replace(/\$+$/, ""));
        i += 1;
        body = block.join("\n");
      }
      const mm = splitMath(body.trim());
      blocks.push({ type: "math", code: mm.code, num: mm.num, anchor: mm.num ? (anchors["式" + mm.num] || null) : null });
      continue;
    }

    if (line.trim().startsWith("|")) {
      const trows = [];
      while (i < n && lines[i].trim().startsWith("|")) { trows.push(lines[i]); i += 1; }
      const rows = trows.map(splitTableRow).filter((r) => !isSepRow(r));
      blocks.push({ type: "table", rows });
      continue;
    }

    const im = /^!\[[^\]]*\]\(([^)]+)\)/.exec(line.trim());
    if (im) { blocks.push({ type: "image", rel: im[1] }); i += 1; continue; }

    const hm = /^(#{1,6})\s+(.*)$/.exec(line);
    if (hm) {
      const hashN = hm[1].length;
      if (!firstSkip) {
        firstSkip = true;
        if (meta.level === 0 || normTitle(hm[2]) === normTitle(meta.title)) { i += 1; continue; }
      }
      const level = mapLevel(meta.level, hashN);
      // 编号（第N章 / 一、 / （一） / 1. / (1））是 tree-doc-numbering 写死在标题文本里的，
      // 严格对齐思源：原样保留为普通标题文本，不再剥号交给 Word 列表生成。
      const text = stripHtml(hm[2]).replace(/[\u200b-\u200d\ufeff]/g, "").replace(/\*\*/g, "").trim();
      blocks.push({ type: "heading", level, text });
      i += 1; continue;
    }

    const lm = /^[-*]\s+(.*)$/.exec(line);
    if (lm) { blocks.push({ type: "body", text: lm[1], listItem: true }); i += 1; continue; }

    let body = line;
    if (body.startsWith("> ")) body = body.slice(2);
    if (body.trim() === "---" || stripHtml(body).replace(/\u200d/g, "").trim() === "") { i += 1; continue; }

    if (isCaption(body)) {
      const cm = /^\s*(表|图|式)\s*([0-9]+(?:-[0-9]+)?)\s+(.*)$/.exec(body);
      const kind = cm[1], num = cm[2], title = stripHtml(cm[3]).replace(/\*\*/g, "").trim();
      blocks.push({ type: "caption", kind, num, title, anchor: anchors[kind + num] || null });
      i += 1; continue;
    }

    blocks.push({ type: "body", text: body, listItem: false });
    i += 1;
  }

  return blocks;
}

module.exports = {
  stripHtml, collectAnchors, parseDoc, parseReference, mapLevel,
  splitMath, isCaption, stripHeadingNum,
};
