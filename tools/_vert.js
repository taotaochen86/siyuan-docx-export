const D = require("../src/docx.js");

function check(name, cond, extra) {
  if (!cond) { console.error("FAIL " + name + " " + (extra || "")); process.exit(1); }
  console.log("PASS " + name);
}

// 截图原句形态：上标单位 + 一个下标
const s = "3.1 × 10<sup>4</sup> m<sup>3</sup> ha<sup>-1</sup> 浓度 g cm<sub>2</sub>";
const out = D.buildInline(s, {}, 12, false, false, {}, null);

check("无字面 sup 标签", !/<\/?sup>/i.test(out));
check("无字面 sub 标签", !/<\/?sub>/i.test(out));
check("上标 run 数=3", (out.match(/vertAlign w:val="superscript"/g) || []).length === 3,
  `got ${(out.match(/superscript/g) || []).length}`);
check("下标 run 数=1", (out.match(/vertAlign w:val="subscript"/g) || []).length === 1,
  `got ${(out.match(/subscript/g) || []).length}`);

const visible = [...out.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(m => m[1]).join("");
check("上下标文字保留", ["4", "3", "-1", "2"].every(x => visible.includes(x)), visible);

// 加粗与上标共存
const o2 = D.buildInline("**强调** ha<sup>-1</sup>", {}, 12, false, false, {}, null);
check("加粗仍在", /<w:b\/>/.test(o2));
check("共存上标在", /superscript/.test(o2));

// 斜体包上标
const o3 = D.buildInline("m<sup>*3*</sup>", {}, 12, false, false, {}, null);
check("上标内斜体在", /superscript/.test(o3) && /<w:i\/>/.test(o3));

console.log("\n可见文本: " + visible);
console.log("\nALL PASS");
