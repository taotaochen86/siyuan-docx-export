const fs=require("fs");const{unzipSync}=require("fflate");
const u=unzipSync(new Uint8Array(fs.readFileSync(process.argv[2])));
const text=[...new TextDecoder().decode(u["word/document.xml"]).matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(m=>m[1]).join("");
let i=text.indexOf("$");
let n=0;
while(i>=0){ n++; console.log(n+" | ["+text.slice(Math.max(0,i-30),i+30).replace(/\s+/g," ")+"]"); i=text.indexOf("$", i+1); }
console.log("合计 "+n);
console.log("\n全部可见文本:\n"+text.slice(0,900));
