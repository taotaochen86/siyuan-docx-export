// 本地冒烟：思源 API 拉真实文档 -> 插件 buildDocument -> 解压 docx 核对关键点
const fs = require("fs");
const path = require("path");
const B = require("../src/builder");

const TOKEN = process.env.SIYUAN_TOKEN || "";
const BASE = "http://127.0.0.1:6806";
const H = { "Content-Type": "application/json" };
if (TOKEN) H.Authorization = "Token " + TOKEN;
const ROOT = process.argv[2] || "20261107193423-jjppxdp";

async function api(p, body) {
  const r = await fetch(BASE + p, { method: "POST", headers: H, body: JSON.stringify(body || {}) });
  const j = await r.json();
  if (j.code !== 0) throw new Error(p + ": " + j.msg);
  return j.data;
}

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
    if (!d.title) { const tm = /^title:\s*(.+)$/m.exec(d.markdown); d.title = tm ? tm[1].trim() : d.id; }
    d.images = [...d.markdown.matchAll(/!\[[^\]]*\]\(([^)]+)\)/g)].map(m => m[1]);
  }
  return { notebook: box, root: rootId, docs };
}

function sniff(b) {
  if (b[0] === 0x89 && b[1] === 0x50) return { ext: "png", mime: "image/png", w: 0, h: 0 };
  if (b[0] === 0xff && b[1] === 0xd8) return { ext: "jpg", mime: "image/jpeg", w: 0, h: 0 };
  if (b[0] === 0x47 && b[1] === 0x49) return { ext: "gif", mime: "image/gif", w: 0, h: 0 };
  if (b[0] === 0x52 && b[1] === 0x49) return { ext: "webp", mime: "image/webp", w: 0, h: 0 };
  return null;
}

async function prefetch(docs) {
  const refs = [...new Set(docs.flatMap(d => d.images).filter(Boolean))];
  const out = {};
  for (const rel of refs) {
    const clean = rel.split("?")[0].split("#")[0];
    const base = clean.split("/").pop();
    for (const u of ["/assets/" + encodeURIComponent(base), "/api/file/getFile?path=" + encodeURIComponent("assets/" + base)]) {
      try {
        const r = await fetch(BASE + u, { headers: H });
        if (!r.ok) continue;
        const buf = new Uint8Array(await r.arrayBuffer());
        if (buf.length < 24) continue;
        const info = sniff(buf);
        if (info) { out[rel] = { bytes: buf, ext: info.ext, mime: info.mime, w: info.w, h: info.h }; }
        break;
      } catch (e) {}
    }
  }
  return out;
}

async function main() {
  const raw = await fetchTree(ROOT, true);
  console.log("文档数: " + raw.docs.length);
  const assets = await prefetch(raw.docs);
  console.log("图片命中: " + Object.keys(assets).length + "/" + raw.docs.flatMap(d => d.images).length);

  let refsByKey = {};
  try {
    const keys = [];
    for (const d of raw.docs) {
      for (const m of d.markdown.matchAll(/zotero:\/\/([^)\s]+)/g)) {
        const zm = /zkeys=([A-Z0-9,]+)/.exec(m[1]);
        if (zm) { keys.push(...zm[1].split(",")); continue; }
        const km = /\/items\/([A-Z0-9]+)/.exec(m[1]);
        if (km) keys.push(km[1]);
      }
    }
    const uniq = [...new Set(keys.filter(Boolean))];
    for (const k of uniq) {
      try { const r = await fetch("http://127.0.0.1:23119/api/users/0/items/" + k + "?format=json"); if (!r.ok) continue; const it = await r.json(); const d = it && it.data; if (d && d.key && d.title) refsByKey[k] = { itemKey: k, itemData: { title: d.title, type: "article" } }; } catch (e) {}
    }
    console.log("Zotero 命中: " + Object.keys(refsByKey).length + "/" + uniq.length);
  } catch (e) { console.log("Zotero 跳过: " + e.message); }

  const bytes = B.buildDocument(raw, {}, refsByKey, assets);
  const outPath = path.join(__dirname, "..", "_smoke_" + ROOT + ".docx");
  fs.writeFileSync(outPath, Buffer.from(bytes));
  console.log("已写: " + outPath + " (" + bytes.length + " B)");

  // 解压核对
  const { execSync } = require("child_process");
  const AdmZip = null;
  const zlib = require("zlib");
  // 用 fflate 读包
  const fflate = require("fflate");
  const un = fflate.unzipSync(bytes);
  const docXml = un["word/document.xml"] ? Buffer.from(un["word/document.xml"]).toString("utf8") : "";
  const numXml = un["word/numbering.xml"] ? Buffer.from(un["word/numbering.xml"]).toString("utf8") : "";

  const cnt = (re, s) => (s.match(re) || []).length;
  console.log("\n=== 核对 ===");
  console.log("字面 <sup>: " + cnt(/<sup>/gi, docXml) + "  <sub>: " + cnt(/<sub>/gi, docXml));
  console.log("vertAlign 上标: " + cnt(/superscript/g, docXml) + " 下标: " + cnt(/subscript/g, docXml));
  console.log("numbering suff=space: " + cnt(/w:suff w:val="space"/g, numXml));
  console.log("SEQ 域: " + cnt(/SEQ /g, docXml) + "  REF 域: " + cnt(/REF /g, docXml) + "  STYLEREF: " + cnt(/STYLEREF/g, docXml));
  console.log("SEQ 带 \\s 1 重置: " + cnt(/SEQ [^\s]+ \\?\*? ARABIC \\s 1/g, docXml));
  console.log("ZOTERO_ITEM: " + cnt(/ZOTERO_ITEM/g, docXml) + "  ZOTERO_BIBL: " + cnt(/ZOTERO_BIBL/g, docXml));
  console.log("fldChar begin: " + cnt(/fldCharType="begin"/g, docXml) + " end: " + cnt(/fldCharType="end"/g, docXml));
  console.log("表格: " + cnt(/<w:tbl>/g, docXml) + "  嵌入图片: " + cnt(/<a:blip/g, docXml));

  // 残留：w:t 文本里 汉字 空格 英文/数字
  const isCaptionLabel = ch => "表图式".includes(ch);
  const bad = [];
  for (const rm of docXml.matchAll(/<w:r\b[^>]*>([\s\S]*?)<\/w:r>/g)) {
    const body = rm[1];
    if (/Consolas/.test(body)) continue; // 代码块 mono 原样保留
    const tm = body.match(/<w:t[^>]*>([^<]*)<\/w:t>/);
    if (!tm) continue;
    const t = tm[1];
    let hit = false;
    for (const mm of t.matchAll(/[\u4e00-\u9fa5]\s+[A-Za-z0-9]/g)) if (!isCaptionLabel(mm[0][0])) { hit = true; break; }
    if (!hit) for (const mm of t.matchAll(/[A-Za-z0-9]\s+[\u4e00-\u9fa5]/g)) { hit = true; break; }
    if (hit) bad.push(t);
  }
  console.log("汉字-英数残留空格文本段: " + bad.length);
  bad.slice(0, 8).forEach(t => console.log("  残留: [" + t + "]"));
}

main().catch(e => { console.error("FAIL", e); process.exit(1); });
