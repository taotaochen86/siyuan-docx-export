# Export Word · siyuan-docx-export

Export a SiYuan document tree to a submission-ready `.docx` — entirely in the browser, no Python, no local service.

## What it does

- **Headings**: native Word `Heading 1–5` styles (visible in the navigation pane); the source numbering (`第1章` / `一、` / `（一）` / `1.` / `(1）`) is kept as static text, not bound to a Word list, so the number sits right next to the title.
- **Body**: SimSun + Times New Roman, 12 pt, 1.5 line spacing, 2-character first-line indent.
- **Tables**: three-line style, borders only on top/bottom and under the header row. Empty tables and empty cells are preserved; column counts are never padded.
- **Images**: fetched from the kernel and written into `word/media/`, truly embedded, scaled to the text column when oversized.
- **Captions**: `表 N` / `图 N` / `式 N` paragraphs become `SEQ` fields with bookmarks; the caption title stays outside the bookmark.
- **Cross-references**: `[表 4-1](#blockid)` becomes a `REF … \h` field. Pressing F9 refreshes the number only — it does not paste the caption text into the body.
- **Zotero**: in-text citations become `ADDIN ZOTERO_ITEM CSL_CITATION` fields; the bibliography becomes a cross-paragraph `ADDIN ZOTERO_BIBL` field.
- One page break per source document; optional cover page.

Works with documents produced by `tree-doc-numbering`, `siyuan-xref` and `siyuan-plugin-zotero-cite`.

## Install

From the SiYuan Marketplace (search "Export Word"), or drop `siyuan-docx-export.zip` from Releases into `<workspace>/data/plugins/siyuan-docx-export/` and restart SiYuan.

## Usage

Click the `iconExport` button in the top bar, or press `Shift+Alt+P` and search "导出规范 Word". Choose the scope (whole tree or current document), fill in the cover fields, export.

## Known limits

- **OMML subset.** Inline `$...$` and block `$$...$$` become native Word OMML `<m:oMath>` objects via a hand-written pure-JS LaTeX→OMML converter (variables/digits/Greek/operators, single or double sub/superscript, fractions, roots, accents, upright `\mathrm`/`\text`, parenthesis/bar delimiter groups, `\left\{…\right.` brace, `matrix/cases` two-row `eqArr`, grouped-bracket sub/superscript). Large integral/sum limits, multi-column true matrices and multi-row alignment are out of subset; unsupported commands drop the command name, and empty parses fall back to linear runs. Real WPS/Word editability needs on-machine review (this host has no Office; only structural validation was run).
- Images support PNG / JPEG / GIF / WebP only.
- Caption numbers are re-sequenced by Word on F9: "表 4-1" becomes "表 3", consistently with the cross-references but different from the source numbering.
- Footnote definition lines (`[^1]: …`) are kept as visible text for now.
- Desktop frontends only; the plugin is not loaded in the web client.

## Build

```bash
npm install
npm run build
```

## License

MIT © 2026 ttChen
