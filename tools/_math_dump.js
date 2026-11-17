const fflate = require("fflate"), fs = require("fs");
const b = fs.readFileSync(process.argv[2]);
const u = fflate.unzipSync(b);
const x = Buffer.from(u["word/document.xml"]).toString("utf8");
const cnt = (re) => (x.match(re) || []).length;
console.log("m:oMath:", cnt(/<m:oMath>/g) + cnt(/<m:oMath /g));
console.log("m:sSub:", cnt(/<m:sSub>/g), " m:sSup:", cnt(/<m:sSup>/g), " m:sSubSup:", cnt(/<m:sSubSup>/g));
console.log("m:f 分数:", cnt(/<m:f>/g), " m:rad:", cnt(/<m:rad>/g), " m:acc:", cnt(/<m:acc>/g));
console.log("w:vertAlign 上标:", cnt(/vertAlign w:val="superscript"/g), " 下标:", cnt(/vertAlign w:val="subscript"/g));
// 普通 w:t 文本里是否残留希腊/命令（非公式 run）
const wt = [...x.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1]);
const stray = wt.filter((t) => /[βδεσμλσθφπΣΔ]|\\(beta|epsilon|sigma|frac|hat|mathrm|cdot)/.test(t));
console.log("普通文本 run 中希腊/命令残留:", stray.length);
stray.slice(0, 15).forEach((t) => console.log("  [" + t.slice(0, 60) + "]"));
// 良构：oMath begin/end 配对
console.log("oMath begin/end 配对:", cnt(/<m:oMath[ >]/g), "/", cnt(/<\/m:oMath>/g));
