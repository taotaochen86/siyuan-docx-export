/**
 * 构建脚本：打包插件入口与导出引擎，并同步到思源插件目录。
 *
 *   npm run build              # 打包 + 自动同步（默认探测 %USERPROFILE%\SiYuan）
 *   SIYUAN_HOME=D:\SiYuan npm run build   # 指定思源工作空间
 *
 * 产物：
 *   index.js   插件入口（CJS，siyuan 为外部依赖）
 *   engine.js  导出引擎（独立 IIFE，挂 window.__docxExport，供 jsrunner 等外部调用）
 */
import esbuild from "esbuild";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.dirname(fileURLToPath(import.meta.url));
const PLUGIN_NAME = "siyuan-docx-export";

const common = {
  bundle: true,
  platform: "browser",
  target: "es2019",
  logLevel: "info",
};

// 插件入口：require("siyuan") 由宿主提供，不打进产物
await esbuild.build({
  ...common,
  entryPoints: [path.join(root, "src", "index.js")],
  format: "cjs",
  external: ["siyuan"],
  minify: true,
  outfile: path.join(root, "index.js"),
});

// 导出引擎：独立 IIFE，可被 jsrunner 直接 eval
await esbuild.build({
  ...common,
  entryPoints: [path.join(root, "src", "engine-entry.js")],
  format: "iife",
  outfile: path.join(root, "engine.js"),
});

// 同步到思源插件目录（找不到就跳过，不影响构建）
const workspace = process.env.SIYUAN_HOME || path.join(os.homedir(), "SiYuan");
const pluginDir = path.join(workspace, "data", "plugins", PLUGIN_NAME);

if (fs.existsSync(path.join(workspace, "data", "plugins"))) {
  fs.mkdirSync(pluginDir, { recursive: true });
  for (const f of ["index.js", "engine.js", "plugin.json", "icon.png", "README.md"]) {
    fs.copyFileSync(path.join(root, f), path.join(pluginDir, f));
  }
  for (const f of ["zh_CN.json", "en_US.json"]) {
    fs.copyFileSync(path.join(root, "i18n", f), path.join(pluginDir, f));
  }
  console.log("synced ->", pluginDir);
} else {
  console.log("SiYuan workspace not found, skipped sync:", workspace);
}
