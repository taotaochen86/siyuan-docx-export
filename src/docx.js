// OOXML(.docx) 生成：字符串拼 XML + fflate 打包。规则平移自 build_docx.py。
const fflate = require("fflate");

const CN_SONG = "宋体", CN_HEI = "黑体", EN_FONT = "Times New Roman", MONO = "Consolas";
const M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";

const CM = 567;            // twips per cm
const EMU_CM = 360000;

function esc(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// 排版间距：删除汉字与英文/数字之间的半角/全角空格；
// 数字↔字母（单位，如 10 cm、m 3）、英文内部、汉字之间的空格保留；
// 题注标签「表/图/式」与编号之间的空格保留（表 4-1）。
function cnSpacing(s) {
  return String(s == null ? "" : s)
    .replace(/([\u4e00-\u9fa5])[ 　]+([A-Za-z0-9])/g, (m, cn, an) =>
      /[表图式]/.test(cn) ? m : cn + an)
    .replace(/([A-Za-z0-9])[ 　]+([\u4e00-\u9fa5])/g, "$1$2");
}

// ---------------- run / rPr ----------------
function rPr(opts) {
  const o = opts || {};
  const cn = o.mono ? MONO : (o.cn || CN_SONG);
  const en = o.mono ? MONO : (o.en || EN_FONT);
  let x = "<w:rPr>";
  x += `<w:rFonts w:ascii="${en}" w:hAnsi="${en}" w:eastAsia="${cn}" w:cs="${en}"/>`;
  if (o.bold) x += "<w:b/>";
  if (o.italic) x += "<w:i/>";
  if (o.vert) x += `<w:vertAlign w:val="${o.vert}"/>`;
  if (o.color) x += `<w:color w:val="${o.color}"/>`;
  const sz = o.size || 12;
  x += `<w:sz w:val="${sz * 2}"/><w:szCs w:val="${sz * 2}"/>`;
  x += "</w:rPr>";
  return x;
}

function runXml(text, opts) {
  const t = String(text == null ? "" : text);
  const shown = (opts && opts.mono) ? t : cnSpacing(t);
  return `<w:r>${rPr(opts)}<w:t xml:space="preserve">${esc(shown)}</w:t></w:r>`;
}

// ---------------- paragraph pPr ----------------
function pPr(opts) {
  const o = opts || {};
  let x = "<w:pPr>";
  if (o.style) x += `<w:pStyle w:val="${o.style}"/>`;
  if (o.jc) x += `<w:jc w:val="${o.jc}"/>`;
  if (o.spacing) {
    const sp = o.spacing;
    x += `<w:spacing ${sp.line ? `w:line="${sp.line}" w:lineRule="${sp.lineRule || "auto"}"` : ""}` +
      `${sp.before != null ? ` w:before="${sp.before}"` : ""}${sp.after != null ? ` w:after="${sp.after}"` : ""}/>`;
  }
  if (o.ind) {
    const ind = o.ind;
    x += "<w:ind" +
      `${ind.firstLineChars != null ? ` w:firstLineChars="${ind.firstLineChars}"` : ""}` +
      `${ind.firstLine != null ? ` w:firstLine="${ind.firstLine}"` : ""}` +
      `${ind.left != null ? ` w:left="${ind.left}"` : ""}` +
      `${ind.hanging != null ? ` w:hanging="${ind.hanging}"` : ""}/>`;
  }
  if (o.shd) x += `<w:shd w:val="clear" w:color="auto" w:fill="${o.shd}"/>`;
  if (o.pageBreakBefore) x += "<w:pageBreakBefore/>";
  x += "</w:pPr>";
  return x;
}

// ---------------- 域 / 书签 ----------------
function fieldRuns(instr, cached, opts) {
  const o = opts || {};
  let x = "";
  x += `<w:r><w:fldChar w:fldCharType="begin"/></w:r>`;
  x += `<w:r><w:instrText xml:space="preserve"> ${esc(instr)} </w:instrText></w:r>`;
  x += `<w:r><w:fldChar w:fldCharType="separate"/></w:r>`;
  x += `<w:r>${rPr({ size: o.size || 12, bold: o.bold, italic: o.italic, cn: o.cn }) }<w:t xml:space="preserve">${esc(cnSpacing(cached || ""))}</w:t></w:r>`;
  x += `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
  return x;
}

function bookmarkStart(id, name) { return `<w:bookmarkStart w:id="${id}" w:name="${esc(name)}"/>`; }
function bookmarkEnd(id) { return `<w:bookmarkEnd w:id="${id}"/>`; }

// ---------------- 行内解析 ----------------
function splitMarks(text, regex) { return text.split(regex); }

function inlineXml(pieces) { return pieces.join(""); }

// 解析一段行内文本，产出 run/域 XML 片段
// imgHook(alt, url) 可选：命中 ![alt](url) 时回调，返回图片 XML（嵌入用）
function buildInline(text, refs, size, baseBold, mono, refsByKey, imgHook) {
  let clean = mono ? text : text.replace(/​/g, "").trim();
  if (!clean && !mono) return "";
  if (!imgHook) return linkPass(clean, refs, size, baseBold, mono, refsByKey);

  const out = [];
  const imgRe = /!\[([^\]]*)\]\(([^)\s]+)[^)]*\)/g;
  let last = 0, m;
  while ((m = imgRe.exec(clean)) !== null) {
    if (m.index > last) out.push(linkPass(clean.slice(last, m.index), refs, size, baseBold, mono, refsByKey));
    out.push(imgHook(m[1], m[2]));
    last = m.index + m[0].length;
  }
  if (last < clean.length) out.push(linkPass(clean.slice(last), refs, size, baseBold, mono, refsByKey));
  return out.join("");
}

function linkPass(clean, refs, size, baseBold, mono, refsByKey) {
  if (!clean) return "";
  let out = [];
  // 链接切分
  const linkRe = /\[([^\]]*)\]\(([^)]+)\)/g;
  let last = 0, m;
  const segs = [];
  while ((m = linkRe.exec(clean)) !== null) {
    segs.push({ kind: "text", value: clean.slice(last, m.index) });
    segs.push({ kind: "link", label: m[1], url: m[2] });
    last = m.index + m[0].length;
  }
  segs.push({ kind: "text", value: clean.slice(last) });

  for (const seg of segs) {
    if (seg.kind === "link") {
      if (seg.url.startsWith("zotero://")) {
        out.push(buildZoteroCite(seg.label, seg.url, refsByKey, size));
      } else if (seg.url.startsWith("#")) {
        out.push(fieldRuns(`REF ${bkmkName(seg.url.slice(1))} \\h`, seg.label, { size }));
      } else if (seg.url.startsWith("http")) {
        out.push(runXml(seg.label, { size, color: "0563C1" }));
      } else {
        out.push(marksXml(seg.label, refs, size, baseBold, mono, refsByKey));
      }
    } else {
      out.push(marksXml(seg.value, refs, size, baseBold, mono, refsByKey));
    }
  }
  return out.join("");
}

// 图片：inline drawing（cx/cy 单位 EMU）
function drawingXml(rid, cx, cy, id, name) {
  return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:docPr id="${id}" name="Picture ${id}"/>` +
    `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="${esc(name)}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
    `</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
}

function marksXml(text, refs, size, baseBold, mono, refsByKey) {
  if (!text) return "";
  const out = [];
  // 先切 HTML 上下标：<sup>…</sup> / <sub>…</sub>（思源正文常用 HTML 标签写单位上标）
  const vertRe = /<(sup|sub)(?:\s[^>]*)?>([\s\S]*?)<\/\1>/g;
  let last = 0, m;
  while ((m = vertRe.exec(text)) !== null) {
    if (m.index > last) out.push(boldMarks(text.slice(last, m.index), refs, size, baseBold, mono, refsByKey));
    const vert = m[1] === "sup" ? "superscript" : "subscript";
    out.push(vertXml(m[2], vert, size, baseBold, mono));
    last = m.index + m[0].length;
  }
  if (last < text.length) out.push(boldMarks(text.slice(last), refs, size, baseBold, mono, refsByKey));
  return out.join("");
}

// 上下标 run：内容仍走加粗/斜体/代码/公式链，vertAlign 注入到每个 run 的 rPr
function vertXml(text, vert, size, bold, mono) {
  const inner = codeItalic(text, size, bold, mono);
  return inner.replace(/(<w:rFonts[^\/]*\/>)/g, `$1<w:vertAlign w:val="${vert}"/>`);
}

function boldMarks(text, refs, size, baseBold, mono, refsByKey) {
  if (!text) return "";
  const out = [];
  const re = /\*\*([^*]+)\*\*/g;
  let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) out.push(codeItalic(text.slice(last, m.index), size, baseBold, mono, refsByKey));
    out.push(codeItalic(m[1], size, true, mono, refsByKey));
    last = m.index + m[0].length;
  }
  if (last < text.length) out.push(codeItalic(text.slice(last), size, baseBold, mono, refsByKey));
  return out.join("");
}

function italicXml(text, size, bold, mono) {
  if (!text) return "";
  const out = [];
  const re = /\*([^*]+)\*/g;
  let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) out.push(mathAndPlain(text.slice(last, m.index), size, bold, mono));
    out.push(runXml(m[1], { size, bold, italic: true, mono }));
    last = m.index + m[0].length;
  }
  if (last < text.length) out.push(mathAndPlain(text.slice(last), size, bold, mono));
  return out.join("");
}

function codeItalic(text, size, bold, mono, refsByKey) {
  if (!text) return "";
  const out = [];
  const re = /`([^`]+)`/g;
  let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) out.push(italicXml(text.slice(last, m.index), size, bold, mono));
    out.push(runXml(m[1], { size, bold, mono: true }));
    last = m.index + m[0].length;
  }
  if (last < text.length) out.push(italicXml(text.slice(last), size, bold, mono));
  return out.join("");
}

function mathAndPlain(seg, size, bold, mono) {
  const parts = seg.split(/(\$[^$]+\$)/g);
  let x = "";
  for (const f of parts) {
    if (f.startsWith("$") && f.endsWith("$") && f.length > 2) {
      // 轻量公式渲染：上下标走 Word 上下标 run，希腊字母/常用符号转字符
      // （完整 OMML 见 README「已知边界」；此处至少保证 A_1 不再显示成 "$A_1$" 原文）
      x += mathXml(f.slice(1, -1), { size, bold, mono });
    } else if (f) {
      x += runXml(f, { size, bold, mono });
    }
  }
  return x;
}

// ---------------- 轻量 LaTeX -> runs ----------------
const GREEK = {
  alpha: "α", beta: "β", gamma: "γ", delta: "δ", epsilon: "ε", varepsilon: "ε",
  zeta: "ζ", eta: "η", theta: "θ", iota: "ι", kappa: "κ", lambda: "λ", mu: "μ",
  nu: "ν", xi: "ξ", pi: "π", rho: "ρ", sigma: "σ", tau: "τ", upsilon: "υ",
  phi: "φ", varphi: "φ", chi: "χ", psi: "ψ", omega: "ω",
  Gamma: "Γ", Delta: "Δ", Theta: "Θ", Lambda: "Λ", Xi: "Ξ", Pi: "Π",
  Sigma: "Σ", Phi: "Φ", Psi: "Ψ", Omega: "Ω",
};
const SYMBOLS = {
  cdot: "·", times: "×", pm: "±", mp: "∓", le: "≤", leq: "≤", ge: "≥", geq: "≥",
  neq: "≠", approx: "≈", infty: "∞", sum: "∑", int: "∫", partial: "∂",
  rightarrow: "→", to: "→", leftarrow: "←", Rightarrow: "⇒", prime: "′",
  quad: " ", qquad: " ", ",": " ", ";": " ", ":": " ", "!": "",
};

function readGroup(s, i) {
  if (s[i] === "{") {
    let depth = 1, j = i + 1, buf = "";
    while (j < s.length && depth > 0) {
      if (s[j] === "{") depth += 1;
      else if (s[j] === "}") depth -= 1;
      if (depth > 0) buf += s[j];
      j += 1;
    }
    return { value: buf, next: j };
  }
  if (s[i] === "\\") {
    let j = i + 1;
    while (j < s.length && /[A-Za-z]/.test(s[j])) j += 1;
    return { value: s.slice(i, j), next: j };
  }
  return { value: s[i] || "", next: i + 1 };
}

function latexFragments(src) {
  const frags = [];
  const s = String(src == null ? "" : src).replace(/\\left|\\right/g, "").replace(/\\[,;:!]/g, " ");
  let i = 0;
  while (i < s.length) {
    const ch = s[i];
    if (ch === "\\") {
      let k = i + 1;
      while (k < s.length && /[A-Za-z]/.test(s[k])) k += 1;
      const name = s.slice(i + 1, k);
      if (name === "frac") {
        const a = readGroup(s, k), b = readGroup(s, a.next);
        frags.push({ text: "(" }); frags.push(...latexFragments(a.value));
        frags.push({ text: ")/(" }); frags.push(...latexFragments(b.value));
        frags.push({ text: ")" });
        i = b.next; continue;
      }
      if (name === "sqrt") {
        const a = readGroup(s, k);
        frags.push({ text: "√(" }); frags.push(...latexFragments(a.value)); frags.push({ text: ")" });
        i = a.next; continue;
      }
      if (name === "text" || name === "mathrm" || name === "mathbf" || name === "operatorname" ||
          name === "overline" || name === "hat" || name === "bar") {
        const a = readGroup(s, k);
        frags.push(...latexFragments(a.value));
        i = a.next; continue;
      }
      if (GREEK[name]) { frags.push({ text: GREEK[name] }); i = k; continue; }
      if (SYMBOLS[name] !== undefined) { frags.push({ text: SYMBOLS[name] }); i = k; continue; }
      // 未知命令：丢掉命令名，保留后续内容
      i = k; continue;
    }
    if (ch === "_" || ch === "^") {
      const g = readGroup(s, i + 1);
      frags.push({ text: latexFragments(g.value).map((f) => f.text).join(""), vert: ch === "_" ? "subscript" : "superscript" });
      i = g.next; continue;
    }
    if (ch === "{" || ch === "}") { i += 1; continue; }
    frags.push({ text: ch });
    i += 1;
  }
  return frags;
}

function mathXml(latex, opts) {
  const o = opts || {};
  return latexFragments(latex)
    .filter((f) => f.text)
    .map((f) => runXml(f.text, { size: o.size || 12, bold: o.bold, mono: o.mono, vert: f.vert }))
    .join("");
}


function bkmkName(id) { return "bkmk_" + id.replace(/[^A-Za-z0-9]/g, ""); }

// ---------------- Zotero 引用 ----------------
function parseZoteroKeys(url) {
  const zm = /zkeys=([A-Z0-9,]+)/.exec(url);
  if (zm) return zm[1].split(",");
  const km = /\/items\/([A-Z0-9]+)/.exec(url);
  return km ? [km[1]] : [];
}

function buildZoteroCite(label, url, refsByKey, size) {
  const keys = parseZoteroKeys(url).filter((k) => refsByKey && refsByKey[k] && refsByKey[k].itemData);
  if (!keys.length) return runXml(label, { size });
  const items = keys.map((k, i) => ({
    id: i + 1,
    uris: [`http://zotero.org/users/local/1/items/${k}`],
    itemData: refsByKey[k].itemData,
  }));
  const citation = {
    citationID: keys[0],
    properties: { formattedCitation: label, plainCitation: label, noteIndex: 0 },
    citationItems: items,
    schema: "https://github.com/citation-style-language/schema/raw/master/csl-citation.json",
  };
  const instr = "ADDIN ZOTERO_ITEM CSL_CITATION " + JSON.stringify(citation);
  return fieldRuns(instr, label, { size });
}

module.exports = {
  esc, rPr, runXml, pPr, fieldRuns, bookmarkStart, bookmarkEnd,
  buildInline, bkmkName, mathXml, latexFragments, drawingXml,
  CN_SONG, CN_HEI, EN_FONT, MONO, CM, EMU_CM, M_NS, fflate,
};
