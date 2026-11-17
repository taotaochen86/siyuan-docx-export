const { Plugin, Dialog, showMessage } = require("siyuan");
const siyuan = require("./siyuan");
const builder = require("./builder");

const ICON_SVG = `<svg viewBox="0 0 16 16" width="16" height="16"><path fill="currentColor" d="M9.5 1h3.2L15 3.3V15H1V3.3L3.3 1h6.2zm.5 1H3.8L2 3.8V14h12V3.8L12.2 2H10zM4 6h8v1.2H4V6zm0 3h8v1.2H4V9zm0 3h5v1.2H4V12z"/></svg>`;

module.exports = class DocxExport extends Plugin {
  async onload() {
    this.curProtyle = null;
    this.eventBus.on("loaded-protyle-static", (e) => this.track(e));
    this.eventBus.on("switch-protyle", (e) => this.track(e));

    this.addTopBar({
      icon: "iconExport",
      title: "导出为规范 Word",
      position: "right",
      callback: () => this.openDialog(),
    });

    this.addCommand({
      langKey: "exportDocxDialog",
      langText: "导出规范 Word（选范围）",
      hotkey: "",
      callback: () => this.openDialog(),
    });

    this._docTreeMenuHandler = (e) => this.onDocTreeMenu(e);
    this.eventBus.on("open-menu-doctree", this._docTreeMenuHandler);
  }

  onunload() {
    if (this._docTreeMenuHandler) {
      this.eventBus.off("open-menu-doctree", this._docTreeMenuHandler);
      this._docTreeMenuHandler = null;
    }
  }

  track(e) {
    if (e && e.detail && e.detail.protyle) this.curProtyle = e.detail.protyle;
  }

  currentDocId() {
    const p = this.curProtyle;
    if (p && p.block && p.block.rootID) return p.block.rootID;
    if (p && typeof p.getInstance === "function") {
      const i = p.getInstance();
      if (i && i.block && i.block.rootID) return i.block.rootID;
    }
    const t = document.querySelector(".protyle-title[data-node-id]");
    if (t) return t.getAttribute("data-node-id");
    return null;
  }

  onDocTreeMenu(e) {
    const d = (e && e.detail) || {};
    if (!d.menu || !d.elements || !d.elements.length) return;
    const el = d.elements[0];
    const node = (el && el.closest && el.closest("li[data-type]")) || el;
    if (!node || !node.getAttribute) return;
    if (String(node.getAttribute("data-type") || "").toLowerCase() !== "navigation-file") return;
    const docId = String(node.getAttribute("data-node-id") || "");
    if (!/^\d{14}-[a-z0-9]{7}$/i.test(docId)) return;
    d.menu.addSeparator();
    d.menu.addItem({
      icon: "iconExport",
      label: "导出规范 Word",
      click: () => this.openDialog(docId),
    });
  }

  openDialog(docId) {
    docId = docId || this.currentDocId();
    if (!docId) { showMessage("请先在思源里打开一篇文档", 5000); return; }

    const dlg = new Dialog({
      title: "导出规范 Word",
      width: "560px",
      content: `<div class="de-dlg" style="padding:16px;">
        <div style="margin-bottom:12px;">
          <label style="display:block;margin-bottom:6px;">导出范围</label>
          <label style="margin-right:16px;"><input type="radio" name="de-scope" value="tree" checked> 当前文档及全部子文档（整册）</label>
          <label><input type="radio" name="de-scope" value="current"> 仅当前文档</label>
        </div>
        <div style="margin-bottom:10px;">
          <label><input type="checkbox" class="de-strip-num"> 导出前剔除标题编号（剥掉 第N章 / 一、/（一）/ 1. 等前缀，标题只留题名；表/图/式题注编号不受影响）</label>
        </div>
        <div style="margin-bottom:10px;"><label>封面书名（留空不生成封面）</label><input class="b3-text-field de-cover-title" style="width:100%;" placeholder="留空则不生成封面"></div>
        <div style="display:flex;gap:12px;margin-bottom:10px;">
          <div style="flex:1;"><label>作者</label><input class="b3-text-field de-author" style="width:100%;" placeholder="陈涛涛"></div>
          <div style="flex:1;"><label>日期</label><input class="b3-text-field de-date" style="width:100%;" placeholder="2026 年 11 月 17 日"></div>
        </div>
        <div style="color:#888;font-size:12px;margin-bottom:14px;">标题套 Heading 1–5 样式、编号保留思源原文（第N章 / 一、/ （一）/ 1. / (1））；表、图、式题注走 SEQ 域，交叉引用走 REF 域；图片真嵌入并按版心自适应；Zotero 引用与书目走 Zotero 域（需 Zotero 运行，否则降级原文）。</div>
        <div style="text-align:right;">
          <button class="b3-button b3-button--outline de-cancel">取消</button>
          <button class="b3-button de-export">导出 .docx</button>
        </div>
      </div>`,
      destroyCallback: () => {},
    });

    const el = dlg.element;
    el.querySelector(".de-cancel").addEventListener("click", () => dlg.destroy());
    el.querySelector(".de-export").addEventListener("click", () => {
      const scope = el.querySelector('input[name="de-scope"]:checked').value;
      const coverTitle = el.querySelector(".de-cover-title").value.trim();
      const coverAuthor = el.querySelector(".de-author").value.trim();
      const coverDate = el.querySelector(".de-date").value.trim();
      const stripHeadingNum = el.querySelector(".de-strip-num").checked;
      this.exportDoc(docId, scope === "tree", { coverTitle, coverAuthor, coverDate, stripHeadingNum }, dlg).catch((e) => {
        showMessage("导出失败：" + ((e && e.message) || e), 8000);
      });
    });
  }

  async exportDoc(docId, includeSub, cover, dlg) {
    showMessage("正在拉取思源内容…", 2000);
    const raw = await siyuan.fetchTree(docId, includeSub);

    showMessage("正在读取 Zotero 文献数据…", 2000);
    const keys = siyuan.collectZoteroKeys(raw.docs);
    let refsByKey = {};
    try {
      refsByKey = await siyuan.fetchZoteroRefs(keys);
    } catch (e) { refsByKey = {}; }
    const hit = Object.keys(refsByKey).length;

    showMessage("正在读取图片…", 2000);
    const imgTotal = siyuan.collectImageRefs(raw.docs).length;
    let assets = {};
    try { assets = await siyuan.prefetchImages(raw.docs); } catch (e) { assets = {}; }
    const imgHit = Object.keys(assets).length;

    showMessage("正在生成 Word…", 2000);
    const uint8 = builder.buildDocument(raw, cover, refsByKey, assets);

    const blob = new Blob([uint8], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const base = cover.coverTitle || (raw.docs[0] && raw.docs[0].title) || "导出文档";
    a.href = url;
    a.download = base + ".docx";
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);

    if (dlg) dlg.destroy();
    const zot = keys.length ? ("，Zotero " + hit + "/" + keys.length) : "";
    const img = imgTotal ? ("，图片 " + imgHit + "/" + imgTotal) : "";
    showMessage("已导出：" + base + ".docx（" + (zot + img).replace(/^，/, "") + "）", 4000);
  }
};
