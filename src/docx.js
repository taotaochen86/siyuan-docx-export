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

const KIND_CODE = { "表": "tbl", "图": "fig", "式": "eq" };
// 题注/公式编号 -> 稳定书签键。思源编号是静态文本，按「标签+编号」建书签才能让
// 交叉引用（REF）跨文档命中同一条题注；用块 ID 时同一编号在多篇文档里会撞名。
function capKey(kind, num) {
  return (KIND_CODE[kind] || "x") + "_" + String(num == null ? "" : num).replace(/-/g, "_");
}
function capBmName(kind, num) { return "bkmk_cap_" + capKey(kind, num); }
const CAP_LABEL_RE = /^(表|图|式)\s*([0-9]+(?:-[0-9]+)?)\s*$/;

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
        // 交叉引用：按「表/图/式 + 编号」命中题注书签；没有对应题注就保留普通文本，
        // 不生成 REF 域，避免 F9 刷新后显示「错误！未定义书签。」
        const cm = CAP_LABEL_RE.exec(seg.label.trim());
        const caps = refs && refs.__caps;
        if (cm && caps && caps.has(capKey(cm[1], cm[2]))) {
          out.push(fieldRuns(`REF ${capBmName(cm[1], cm[2])} \\h`, seg.label, { size }));
        } else {
          out.push(marksXml(seg.label, refs, size, baseBold, mono, refsByKey));
        }
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
      // 优先转 Word 原生 OMML 公式对象；解析为空时回退轻量文本 run
      const omml = mathOmml(f.slice(1, -1));
      x += omml || mathXml(f.slice(1, -1), { size, bold, mono });
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

// ---------------- LaTeX -> OMML（Word 原生公式对象） ----------------
// 子集覆盖本稿公式：变量/数字/希腊字母/运算符、上下标与双标、分数、根号、
// 帽/上划线、正体组、括号/竖线定界组、matrix/cases 多行阵列、\left/\right 定界符。
// 未识别命令跳过命令名保留后续；解析为空时由调用方回退普通文本 run。
function ommlRun(text, upright) {
  const rpr = upright ? "<m:rPr><m:sty m:val=\"normal\"/></m:rPr>" : "";
  return `<m:r>${rpr}<m:t xml:space="preserve">${esc(text)}</m:t></m:r>`;
}

// 配对括号/竖线内部（按深度配对）
function readDelim(s, i, open, close) {
  let depth = 1, j = i + 1, buf = "";
  while (j < s.length && depth > 0) {
    if (s[j] === close) {
      depth -= 1;
      if (depth === 0) { j += 1; break; }
    } else if (s[j] === open) {
      depth += 1;
    }
    if (depth > 0) buf += s[j];
    j += 1;
  }
  return { value: buf, next: j };
}

// 读取 \left/\right 后的定界字符（\{ 取 {，. 表示空端）
function readDelimChar(s, i) {
  if (s[i] === "\\") return { chr: s[i + 1] || ".", next: i + 2 };
  return { chr: s[i] || ".", next: i + 1 };
}

function delimNode(beg, end, inner) {
  const pr = `<m:dPr><m:begChr m:val="${esc(beg)}"/><m:endChr m:val="${esc(end)}"/></m:dPr>`;
  return `<m:d>${pr}<m:e>${inner}</m:e></m:d>`;
}

function ommlAtom(s, i, up) {
  const ch = s[i];
  // 括号 / 竖线定界组整体作为一个基元（便于后续上下标绑定到整组）
  if (ch === "(" || ch === "|") {
    const close = ch === "(" ? ")" : "|";
    const g = readDelim(s, i, ch, close);
    return { node: delimNode(ch, close, ommlSeq(g.value, up)), next: g.next };
  }
  if (ch === "{") {
    const g = readGroup(s, i);
    return { node: ommlSeq(g.value, up), next: g.next };
  }
  if (ch === "\\") {
    let k = i + 1;
    while (k < s.length && /[A-Za-z]/.test(s[k])) k += 1;
    const name = s.slice(i + 1, k);
    if (name === "frac") {
      const a = readGroup(s, k), b = readGroup(s, a.next);
      return {
        node: `<m:f><m:fPr><m:type m:val="bar"/></m:fPr><m:num>${ommlSeq(a.value, up)}</m:num><m:den>${ommlSeq(b.value, up)}</m:den></m:f>`,
        next: b.next,
      };
    }
    if (name === "sqrt") {
      const a = readGroup(s, k);
      return {
        node: `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${ommlSeq(a.value, up)}</m:e></m:rad>`,
        next: a.next,
      };
    }
    if (name === "hat" || name === "bar" || name === "overline") {
      const a = readGroup(s, k);
      const chr = name === "hat" ? "^" : name === "bar" ? "¯" : "‾";
      return {
        node: `<m:acc><m:accPr><m:chr m:val="${chr}"/></m:accPr><m:e>${ommlSeq(a.value, up)}</m:e></m:acc>`,
        next: a.next,
      };
    }
    if (name === "text" || name === "mathrm" || name === "mathbf" || name === "operatorname") {
      const a = readGroup(s, k);
      return { node: ommlSeq(a.value, true), next: a.next };
    }
    // 多行阵列：\begin{matrix}/\begin{cases} ... \\ 行 ... \end{...}
    if (name === "begin") {
      const envG = readGroup(s, k);
      const env = envG.value;
      const endTag = `\\end{${env}}`;
      const ei = s.indexOf(endTag, envG.next);
      const inner = ei < 0 ? s.slice(envG.next) : s.slice(envG.next, ei);
      const rows = inner.split(/\\\\/);
      const es = rows.map((r) => `<m:e>${ommlSeq(r, up)}</m:e>`).join("");
      const next = ei < 0 ? s.length : ei + endTag.length;
      return { node: `<m:eqArr>${es}</m:eqArr>`, next };
    }
    if (name === "end") {
      const g = readGroup(s, k);
      return { node: "", next: g.next };
    }
    // \left { ... \right . 定界符包裹整段
    if (name === "left") {
      const bc = readDelimChar(s, k);
      const ri = s.indexOf("\\right", bc.next);
      let ec = { chr: ".", next: bc.next };
      let innerEnd = ri;
      if (ri >= 0) {
        ec = readDelimChar(s, ri + 6);
        innerEnd = ri;
      }
      const inner = ommlSeq(s.slice(bc.next, innerEnd), up);
      return { node: delimNode(bc.chr, ec.chr, inner), next: ri >= 0 ? ec.next : innerEnd };
    }
    if (name === "right") {
      const ec = readDelimChar(s, k);
      return { node: "", next: ec.next };
    }
    if (GREEK[name]) return { node: ommlRun(GREEK[name], up), next: k };
    if (SYMBOLS[name] !== undefined) return { node: ommlRun(SYMBOLS[name], up), next: k };
    // 转义花括号/括号字面（\{、\}）；未知命令名丢弃
    if (!name) {
      const c = s[i + 1];
      if (c === "{" || c === "}" || c === "|" || c === "(" || c === ")") {
        return { node: ommlRun(c, up), next: i + 2 };
      }
      return { node: "", next: i + 1 };
    }
    return { node: "", next: k };
  }
  // 连续字母/数字合并为一个基元 run（NH、SSE 等整体作为上下标的 e）
  let j = i;
  while (j < s.length && /[A-Za-z0-9]/.test(s[j])) j += 1;
  if (j > i) return { node: ommlRun(s.slice(i, j), up), next: j };
  return { node: ommlRun(ch, up), next: i + 1 };
}

function ommlSupSub(eNode, s, i, up) {
  let sub = null, sup = null;
  while (s[i] === "_" || s[i] === "^") {
    const isSub = s[i] === "_";
    i += 1;
    while (s[i] === " " || s[i] === "~") i += 1;
    const g = readGroup(s, i);
    const content = ommlSeq(g.value, up);
    i = g.next;
    if (isSub) sub = content; else sup = content;
  }
  if (sub && sup) {
    return { node: `<m:sSubSup><m:e>${eNode}</m:e><m:sub>${sub}</m:sub><m:sup>${sup}</m:sup></m:sSubSup>`, next: i };
  }
  if (sub) {
    return { node: `<m:sSub><m:sSubPr><m:ctrlPr/></m:sSubPr><m:e>${eNode}</m:e><m:sub>${sub}</m:sub></m:sSub>`, next: i };
  }
  if (sup) {
    return { node: `<m:sSup><m:sSupPr><m:ctrlPr/></m:sSupPr><m:e>${eNode}</m:e><m:sup>${sup}</m:sup></m:sSup>`, next: i };
  }
  return { node: eNode, next: i };
}

function ommlSeq(src, up) {
  const s = String(src == null ? "" : src).replace(/\\[,;:!]/g, " ");
  let out = "", i = 0;
  while (i < s.length) {
    const ch = s[i];
    if (ch === "{" || ch === "}") { i += 1; continue; }
    const atom = ommlAtom(s, i, up);
    if (!atom.node) { i = atom.next; continue; }
    // 基元后先跳过空白再判定上下标：命令与 _/^ 间的空格不打断绑定
    let j = atom.next;
    while (s[j] === " " || s[j] === "~") j += 1;
    if (s[j] === "_" || s[j] === "^") {
      const ss = ommlSupSub(atom.node, s, j, up);
      out += ss.node;
      i = ss.next;
      continue;
    }
    out += atom.node;
    if (j > atom.next) out += ommlRun(s.slice(atom.next, j), up);
    i = atom.next;
  }
  return out;
}

// 行内公式 -> <m:oMath>；块级由 builder 再包 <m:oMathPara>。
// 解析为空返回空串，调用方回退普通文本。
function mathOmml(latex) {
  const body = ommlSeq(latex);
  if (!body) return "";
  return `<m:oMath xmlns:m="${M_NS}">${body}</m:oMath>`;
}

// 旧轻量文本降级（OMML 失败时回退用）
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
  buildInline, bkmkName, capKey, capBmName, mathXml, mathOmml, latexFragments, drawingXml,
  CN_SONG, CN_HEI, EN_FONT, MONO, CM, EMU_CM, M_NS, fflate,
};
