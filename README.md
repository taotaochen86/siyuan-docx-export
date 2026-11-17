# 导出规范 Word · siyuan-docx-export

把思源笔记的**文档树**导出成一份可直接交稿的 `.docx`：标题带 Word 自动编号、题注是 `SEQ` 域、交叉引用是 `REF` 域、Zotero 引用与书目是 Word 域、图片真嵌入、表格三线、空表照留、每篇文档另起一页。

**纯前端实现**：所有转换在思源的渲染进程里完成，不依赖本地 Python、不需要额外服务。

---

## 特性

| 能力 | 说明 |
|---|---|
| 标题 | 套 Word 内置 `Heading 1–5`（导航窗格可见），带五级自动编号：`第1章` / `一、` / `（一）` / `1.` / `(1）` |
| 正文 | 宋体 + Times New Roman，小四 12pt，1.5 倍行距，首行缩进 2 字符 |
| 表格 | 三线表（顶/底 sz12、表头 sz6），无竖线无内部横线；**空表与空单元格完整保留**，列数按原文不补齐 |
| 图片 | 从内核取二进制写入 `word/media/`，真嵌入；超版心等比缩放；取不到才降级为占位文字 |
| 题注 | `表 N` / `图 N` / `式 N` 段落 → `SEQ 表\|图\|式 \* ARABIC` 域 + 书签，题名在书签外 |
| 交叉引用 | `[表 4-1](#块ID)` → `REF bkmk_xxx \h` 域；F9 后只刷新编号，不会把题名一起带进正文 |
| Zotero 引用 | `zotero://…/items/KEY` → `ADDIN ZOTERO_ITEM CSL_CITATION` 域（含完整 CSL itemData） |
| Zotero 书目 | 文末条目 → 跨段 `ADDIN ZOTERO_BIBL` 域 + `zot_<KEY>` 书签 |
| 分页 | 每篇思源文档另起一页；可选封面（书名 / 作者 / 日期） |
| 公式 | 上下标、希腊字母、`\frac`、`\sqrt` 轻量渲染为 Word run（**不生成 OMML**，见边界） |

## 安装

**方式一（推荐）**：思源 → 设置 → 集市 → 搜索「导出规范 Word」。

**方式二（手动）**：下载 Release 里的 `siyuan-docx-export.zip`，解压到
`<思源工作空间>/data/plugins/siyuan-docx-export/`，重启思源。

## 用法

1. **顶栏按钮**：右上角一排找 `iconExport` 图标的「导出规范 Word」。
2. **命令面板**：`Shift+Alt+P` → 搜「导出规范 Word」。
3. 面板里选范围（整册含子文档 / 仅当前文档），填封面信息，点导出。文件落到浏览器下载目录。

> 顶栏按钮不显示时先试命令面板。插件早期版本曾把整段 SVG 字符串传给 `addTopBar` 的 `icon`，思源按图标名拼 `<use>` 导致按钮空白；现已改用内置图标名 `iconExport`。

## Word 域说明

| 对象 | 指令 | 在 Word 里 |
|---|---|---|
| 表题注 | `SEQ 表 \* ARABIC` | F9 后按出现顺序重排为 表 1、表 2… |
| 图题注 | `SEQ 图 \* ARABIC` | 同上 |
| 公式编号 | `SEQ 式 \* ARABIC` | 右对齐 `(1)` 形式 |
| 交叉引用 | `REF bkmk_xxx \h` | 超链接，只显示编号 |
| Zotero 文内 | `ADDIN ZOTERO_ITEM CSL_CITATION {…}` | 装了 Zotero Word 插件可回写联动，否则显示缓存文本 |
| Zotero 书目 | `ADDIN ZOTERO_BIBL {…}` | 跨段域，begin 在首条、end 在末条 |

## 与三个思源插件的对接

| 插件 | 思源里的形态 | 落到 Word 里 |
|---|---|---|
| `tree-doc-numbering` | 标题文本写死编号（`一、` / `（一）` / `1.` / `3.1.4.40.1`） | 剥掉原编号，改由 `numbering.xml` 五级列表生成 |
| `siyuan-xref` | 题注独立段落；引用 `[表 4-1](#块ID)` | 题注 → `SEQ` 域 + 书签；引用 → `REF` 域 |
| `siyuan-plugin-zotero-cite` | 文内 `[作者, 年](zotero://…)`；书目 `1\. 题录 [↗](zotero://…?bib=1)` | 文内 → `ZOTERO_ITEM` 域；书目 → `ZOTERO_BIBL` 域 |

Zotero 元数据经本机 `127.0.0.1:23119` 实时取：批量 `itemKey=` 请求，漏项单条兜底；Zotero 未运行则整批降级保留原文，不阻断出稿。

## 已知边界

- **公式不生成 OMML**。只做轻量 run 渲染，矩阵、多行对齐、大型运算符会退化成线性文本。纯浏览器链路里没有可信的 LaTeX→OMML 转换器，这是有意的取舍。
- **图片仅支持 PNG / JPEG / GIF / WebP**，BMP / TIFF 降级为占位。
- **题注编号由 Word 重排**。思源的「表 4-1」F9 后会变成「表 3」：正文引用同步变，前后一致，但数字与原文不同。要保留「章-序」需改成 `STYLEREF 1 \s` + `SEQ 表 \* ARABIC \s 1`。
- **脚注定义行** `[^1]: …` 目前作为可见文字保留，尚未转成 Word 真脚注。
- **题注书签只覆盖「表/图/式 + 编号」**，题名必须在书签外。把整段包进书签的话，F9 更新 REF 域时 Word 会把题名整条复制进正文。
- **标题样式必须带 `<w:name w:val="heading N"/>`**。只写 `w:styleId` 不写 `w:name`，Word 会当成匿名样式（显示 `_Style 2`），表现为「章节在导航窗格全部消失」。
- 网页端不加载本插件（`frontends` 只声明 `desktop`、`desktop-window`）。

## 开发

```bash
npm install
npm run build      # 打包并写入思源插件目录
```

单独打引擎（改 `src/` 后按需重新打包，思源按时间戳按需拉取，无需重启）：

```bash
./node_modules/.bin/esbuild src/engine-entry.js --bundle --format=iife \
  --platform=browser --target=es2019 \
  --outfile="<思源工作空间>/data/plugins/siyuan-docx-export/engine.js"
```

`require("siyuan")` 为外部依赖（`--external:siyuan`），`fflate` 已内联。

### 目录

```
plugin.json     思源插件清单
index.js        插件主入口（构建产物）
engine.js       导出引擎（独立 IIFE，挂 window.__docxExport）
i18n/           中英文语言包
src/            源码：index / siyuan / markdown / docx / builder
tools/          结构审计与对账脚本
docs/           全面诊断报告
```

### 自检工具

```bash
node tools/_audit.js "任意.docx"     # 15 项结构审计，有 FAIL 退出码 1
node tools/_deep.js  "任意.docx"     # 公式残留、表格边框、空单元格
node tools/_checkcap.js "任意.docx"  # 题注书签范围是否正确
```

审计覆盖：XML 良构、Content_Types、关系闭环、rId 引用、域配对、书签 id/名称唯一与配对、`pStyle` 引用、Heading 内置名、`numId`/`ilvl` 引用、图片 rId 与 media 对应、图片尺寸不超版心。

## License

MIT © 2026 ttChen
