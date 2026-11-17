const D = require("../src/docx.js");
const latex = `\\theta ( h ) = \\left\\{ \\begin{matrix} \\theta _ { r } + \\frac { ( \\theta _ { s } - \\theta _ { r } ) } { ( 1 + | \\alpha h | ^ { n } ) ^ { m } },h < 0 \\\\ \\theta _ { s }, h \\geq 0 \\end{matrix} \\right.`;
const omml = D.mathOmml(latex);
console.log(omml);
// 粗良构：每个开标签有对应闭标签
const opens = omml.match(/<m:[a-zA-Z]+[ >]/g) || [];
const closes = omml.match(/<\/m:[a-zA-Z]+>/g) || [];
const tag = (t) => t.replace(/[ <\/>]/g, "");
const stack = [];
let ok = true;
for (const t of omml.match(/<\/?m:[a-zA-Z]+[^>]*>/g) || []) {
  if (t.endsWith("/>")) continue;
  if (t.startsWith("</")) { if (stack.pop() !== tag(t)) { ok = false; break; } }
  else stack.push(tag(t));
}
console.log("\\n开标签:", opens.length, " 闭标签:", closes.length, " 栈余:", stack.length, " 良构:", ok && stack.length === 0);
