// 思源内核 API 封装（插件渲染进程同源，免 Token）
const TIMEOUT_MS = 30000;
const BASE = "http://127.0.0.1:6806";
function resolveUrl(path) {
  if (/^https?:/.test(path)) return path;
  if (typeof location !== "undefined" && location.origin) return path;
  return BASE + path;
}

async function api(path, payload) {
  const ctl = (typeof AbortController === "function") ? new AbortController() : null;
  const timer = ctl ? setTimeout(() => { try { ctl.abort(); } catch (e) { /* ignore */ } }, TIMEOUT_MS) : null;
  let res;
  try {
    res = await fetch(resolveUrl(path), {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload || {}),
      signal: ctl ? ctl.signal : undefined,
    });
  } finally {
    if (timer) clearTimeout(timer);
  }
  const j = await res.json();
  if (j.code !== 0) throw new Error(path + "：" + (j.msg || ("code " + j.code)));
  return j.data;
}

function collectImages(md) {
  const out = [];
  const re = /!\[[^\]]*\]\(([^)]+)\)/g;
  let m;
  while ((m = re.exec(md)) !== null) out.push(m[1]);
  return out;
}

// 递归拉取根文档及其子文档
async function fetchTree(rootId, includeSub) {
  const info = await api("/api/block/getBlockInfo", { id: rootId });
  const box = info && info.box ? info.box : "";

  const docs = [{ id: rootId, title: "", path: "/" + rootId + ".sy", level: 0, markdown: "", images: [] }];

  if (includeSub) {
    const walk = async (docPath, level) => {
      const data = await api("/api/filetree/listDocsByPath", { notebook: box, path: docPath });
      const files = (data && data.files) || [];
      files.sort((a, b) => (a.sort || 0) - (b.sort || 0));
      for (const f of files) {
        docs.push({ id: f.id, title: f.name || "", path: f.path, level, markdown: "", images: [] });
        if ((f.subFileCount || 0) > 0) await walk(f.path, level + 1);
      }
    };
    await walk("/" + rootId + ".sy", 1);
  }

  for (const d of docs) {
    const md = (await api("/api/export/exportMdContent", { id: d.id })) || "";
    d.markdown = typeof md === "string" ? md : (md && md.content) || "";
    if (!d.title) {
      const tm = /^title:\s*(.+)$/m.exec(d.markdown);
      d.title = tm ? tm[1].trim() : d.id;
    }
    d.images = collectImages(d.markdown);
  }

  return { notebook: box, root: rootId, docs };
}

// 图片 -> dataURL（浏览器读不了本地文件，走内核 getFile）
async function fetchImageDataURL(rel) {
  const base = rel.split("/").pop();
  const candidates = [rel, "assets/" + base, base];
  for (const p of candidates) {
    try {
      const r = await fetch(resolveUrl("/api/file/getFile?path=" + encodeURIComponent(p)));
      if (!r.ok) continue;
      const blob = await r.blob();
      if (!blob || blob.size === 0) continue;
      return await new Promise((resolve) => {
        const fr = new FileReader();
        fr.onload = () => resolve(fr.result);
        fr.readAsDataURL(blob);
      });
    } catch (e) { /* 试下一个候选 */ }
  }
  return null;
}

// ---------------- 图片二进制预取（嵌入 docx 用） ----------------
// 渲染进程与内核同源，/assets/<文件名> 可直接 GET；失败再回退 /api/file/getFile。
async function loadImageBytes(rel) {
  const clean = String(rel).split("?")[0].split("#")[0];
  const base = clean.split("/").pop();
  const cands = [];
  if (/^https?:/i.test(clean)) {
    cands.push(clean);
  } else {
    cands.push("/assets/" + encodeURIComponent(base));
    cands.push("/assets/" + base);
    cands.push("/api/file/getFile?path=" + encodeURIComponent("assets/" + base));
  }
  for (const u of cands) {
    try {
      const r = await fetch(u);
      if (!r.ok) continue;
      const buf = new Uint8Array(await r.arrayBuffer());
      if (buf.length < 24) continue;
      return buf;
    } catch (e) { /* 下一个 */ }
  }
  return null;
}

// 只读文件头判断格式与像素尺寸
function sniffImage(b) {
  if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) {
    const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
    return { ext: "png", mime: "image/png", w: dv.getUint32(16), h: dv.getUint32(20) };
  }
  if (b[0] === 0xff && b[1] === 0xd8) {
    let i = 2;
    while (i + 8 < b.length) {
      if (b[i] !== 0xff) { i += 1; continue; }
      const mk = b[i + 1];
      if (mk === 0xd8 || mk === 0xd9) { i += 2; continue; }
      const len = (b[i + 2] << 8) | b[i + 3];
      if (mk >= 0xc0 && mk <= 0xcf && mk !== 0xc4 && mk !== 0xc8 && mk !== 0xcc) {
        return { ext: "jpg", mime: "image/jpeg", w: (b[i + 7] << 8) | b[i + 8], h: (b[i + 5] << 8) | b[i + 6] };
      }
      i += 2 + (len || 2);
    }
    return { ext: "jpg", mime: "image/jpeg", w: 0, h: 0 };
  }
  if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) {
    return { ext: "gif", mime: "image/gif", w: b[6] | (b[7] << 8), h: b[8] | (b[9] << 8) };
  }
  if (b[0] === 0x52 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x46 &&
      b[8] === 0x57 && b[9] === 0x45 && b[10] === 0x42 && b[11] === 0x50) {
    return { ext: "webp", mime: "image/webp", w: 0, h: 0 };
  }
  return null;
}

// 收集文档里所有图片引用（![alt](url)）
function collectImageRefs(docs) {
  const re = /!\[[^\]]*\]\(([^)\s]+)/g;
  const out = [];
  for (const d of docs) {
    let m;
    re.lastIndex = 0;
    while ((m = re.exec(d.markdown || "")) !== null) out.push(m[1]);
  }
  return Array.from(new Set(out));
}

// 预取图片：返回 { [url]: { bytes, ext, mime, w, h } }
async function prefetchImages(docs) {
  const refs = collectImageRefs(docs);
  const out = {};
  for (const rel of refs) {
    if (/^data:/i.test(rel)) continue;
    const bytes = await loadImageBytes(rel);
    if (!bytes) continue;
    const info = sniffImage(bytes);
    if (!info) continue;
    out[rel] = { bytes, ext: info.ext, mime: info.mime, w: info.w, h: info.h };
  }
  return out;
}

// ---------------- Zotero（本机 23119 本地 API） ----------------
const ZOT_BASE = "http://127.0.0.1:23119";
const ZOT_SKIP = { attachment: 1, note: 1, annotation: 1 };

const CSL_TYPE = {
  journalArticle: "article-journal", magazineArticle: "article-magazine",
  newspaperArticle: "article-newspaper", book: "book", bookSection: "chapter",
  conferencePaper: "paper-conference", thesis: "thesis", dissertation: "thesis",
  report: "report", webpage: "webpage", preprint: "article", manuscript: "manuscript",
  encyclopediaArticle: "entry-encyclopedia", patent: "patent", standard: "standard",
};

// Zotero API item.data -> CSL itemData（Word 域里 citationItems[].itemData 用这套字段）
function toCSL(d) {
  const creators = (d.creators || [])
    .filter((c) => c.creatorType === "author" || c.creatorType === "editor")
    .map((c) => {
      const o = { family: c.lastName || "", given: c.firstName || "" };
      if (c.name) { o.family = c.name; delete o.given; }
      if (!o.family && !o.given) return null;
      return o;
    })
    .filter(Boolean);
  const ym = /(\d{4})/.exec(d.date || "");
  const item = {
    type: CSL_TYPE[d.itemType] || "article",
    title: d.title || "",
    author: creators,
    "container-title": d.publicationTitle || d.bookTitle || d.proceedingsTitle || "",
    publisher: d.publisher || "",
    "publisher-place": d.place || "",
    volume: d.volume || "",
    issue: d.issue || "",
    page: d.pages || "",
    DOI: d.DOI || "",
    URL: d.url || "",
    ISBN: d.ISBN || "",
    language: d.language || "",
  };
  if (ym) item.issued = { "date-parts": [[parseInt(ym[1], 10)]] };
  const out = {};
  for (const k of Object.keys(item)) {
    const v = item[k];
    if (v === "" || v == null) continue;
    if (Array.isArray(v) && !v.length) continue;
    out[k] = v;
  }
  return out;
}

// 按 itemKey 批量取条目 -> { KEY: {itemKey, itemData} }
async function fetchZoteroRefs(keys) {
  const uniq = Array.from(new Set((keys || []).filter(Boolean)));
  const out = {};
  if (!uniq.length) return out;
  for (let i = 0; i < uniq.length; i += 50) {
    const batch = uniq.slice(i, i + 50);
    try {
      const res = await fetch(ZOT_BASE + "/api/users/0/items?itemKey=" +
        encodeURIComponent(batch.join(",")) + "&format=json&limit=100");
      if (!res.ok) continue;
      const arr = await res.json();
      if (!Array.isArray(arr)) continue;
      for (const it of arr) {
        const d = (it && it.data) || {};
        if (!d.key || ZOT_SKIP[d.itemType] || !d.title) continue;
        if (batch.indexOf(d.key) < 0) continue;
        out[d.key] = { itemKey: d.key, itemData: toCSL(d) };
      }
    } catch (e) { /* Zotero 未启动或未授权：整批跳过，走降级 */ }
  }
  // 批量偶有漏项（出版变体等），逐个兜底
  for (const k of uniq) {
    if (out[k]) continue;
    try {
      const res = await fetch(ZOT_BASE + "/api/users/0/items/" + encodeURIComponent(k) + "?format=json");
      if (!res.ok) continue;
      const it = await res.json();
      const d = (it && it.data) || {};
      if (!d.key || ZOT_SKIP[d.itemType] || !d.title) continue;
      out[d.key] = { itemKey: d.key, itemData: toCSL(d) };
    } catch (e) { /* 单条取不到也降级 */ }
  }
  return out;
}

// 从 Markdown 里收集所有 zotero:// 的 key（含 zkeys= 多文献）
function collectZoteroKeys(docs) {
  const out = [];
  const re = /zotero:\/\/([^)\s]+)/g;
  for (const doc of docs || []) {
    let m;
    re.lastIndex = 0;
    while ((m = re.exec(doc.markdown || "")) !== null) {
      const url = m[1];
      const zm = /zkeys=([A-Z0-9,]+)/.exec(url);
      if (zm) { out.push.apply(out, zm[1].split(",")); continue; }
      const km = /\/items\/([A-Z0-9]+)/.exec(url);
      if (km) out.push(km[1]);
    }
  }
  return Array.from(new Set(out.filter(Boolean)));
}

module.exports = {
  api, fetchTree, fetchImageDataURL, collectImages,
  fetchZoteroRefs, collectZoteroKeys, toCSL,
  prefetchImages, collectImageRefs,
};
