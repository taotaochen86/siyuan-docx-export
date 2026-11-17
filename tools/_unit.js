const D = require("../src/docx.js");
const s = "行内 $A_1$ 与 $\\frac{a}{b}$ 结束。";
console.log("输入: " + s);
const out = D.buildInline(s, {}, 12, false, false, {}, null);
console.log("输出XML: " + out);
console.log("\n可见文本: " + [...out.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(m=>m[1]).join("|"));
console.log("\nmathXml 直测: " + D.mathXml("A_1", {size:12}));
