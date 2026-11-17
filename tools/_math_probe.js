const D = require("../src/docx.js");
const cases = ["$A_1B_1$", "$R^2$", "$r$", "$\\hat{y}$", "$\\hat{\\beta}_0$", "$t~ha^{-1}$", "$\\alpha$"];
const vis = (x) => [...x.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1]).join("");
for (const c of cases) {
  const xml = D.buildInline(c, {}, 12, false, false, {});
  console.log("IN : " + c + "\nOUT: " + vis(xml) + "  vert=" + (xml.match(/vertAlign/g) || []).length);
}
// 代码块走 mono，$ 应原样
const code = "dt$GrainYield <- runif(18)";
const codeXml = D.buildInline(code, {}, 12, false, true, {});
console.log("CODE OUT: " + vis(codeXml) + "  mono=" + /Consolas/.test(codeXml));
