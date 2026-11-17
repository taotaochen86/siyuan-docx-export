const D = require("../src/docx.js");
const fs = require("fs");
const path = require("path");

function check(name, cond, extra) {
  if (!cond) { console.error("FAIL " + name + " " + (extra || "")); process.exit(1); }
  console.log("PASS " + name);
}

function visible(x) {
  return [...x.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(m => m[1]).join("");
}

// 1. 汉字↔英文/数字空格删除；数字↔字母（单位）空格保留
const s = "浓度 10 cm 以上 3 个 pacman 包";
const v = visible(D.buildInline(s, {}, 12, false, false, {}, null));
check("汉字-数字空格删", v.includes("浓度10"), v);
check("数字-单位字母空格保留", v.includes("10 cm"), v);
check("字母-汉字空格删", v.includes("cm以上"), v);
check("数字-汉字空格删", v.includes("以上3") && v.includes("3个"), v);
check("英文-汉字两侧删", v.includes("个pacman") && v.includes("pacman包"), v);

// 2. 单位与数字保留：kg、m3、ha-1 形态
const v2 = visible(D.buildInline("10 kg m<sup>3</sup> ha<sup>-1</sup>", {}, 12, false, false, {}, null));
check("kg 前空格保留", v2.includes("10 kg"), v2);

// 3. 代码块 mono 不动间距
const v3 = visible(D.buildInline("结 3 个 `x  y`", {}, 12, false, true, {}, null));
check("mono 内空格保留", v3.includes("x  y"), v3);

// 3b. 题注标签「表/图/式」与编号空格保留
const v4 = visible(D.buildInline("见表 4-1 和图 2 所示", {}, 12, false, false, {}, null));
check("题注标签空格保留", v4.includes("表 4-1") && v4.includes("图 2"), v4);

// 4. 标题编号为静态文本：不绑 Word 列表（无 numPr/numId/abstractNum/lvl），编号原文保留
const builderSrc = fs.readFileSync(path.join(__dirname, "../src/builder.js"), "utf8");
const mdSrc = fs.readFileSync(path.join(__dirname, "../src/markdown.js"), "utf8");
check("标题路径无 numPr/numId 注入", !/<w:numPr/.test(builderSrc), "numPr 仍存在");
check("numbering.xml 为空壳（无 abstractNum/lvl）", !/<w:abstractNum|<w:lvl /.test(builderSrc), "仍含列表定义");
check("解析层标题分支默认不剥编号", !/const text = stripHeadingNum/.test(mdSrc), "parseDoc 仍默认剥号");

// 5. 可选剥号函数：覆盖各静态编号形态，无编号题名原样保留
const M = require("../src/markdown.js");
const stripCases = [
  ["第1章 绪论", "绪论"],
  ["第3章 材料与方法", "材料与方法"],
  ["一、文档的标题变", "文档的标题变"],
  ["（一）研究目的与意义", "研究目的与意义"],
  ["1. 研究意义", "研究意义"],
  ["(1）撒旦法撒旦法", "撒旦法撒旦法"],
  ["3.1.4.40.1 多级编号标题", "多级编号标题"],
  ["2026 年工作总结", "2026 年工作总结"],
  ["绪论（无编号）", "绪论（无编号）"],
];
stripCases.forEach(([input, expect]) => {
  check("剥号：" + input, M.stripHeadingNum(input) === expect, "=> " + M.stripHeadingNum(input));
});

console.log("\n可见文本: " + v);
console.log("可见文本2: " + v2);
console.log("\nALL PASS");
