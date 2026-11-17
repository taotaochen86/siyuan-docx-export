const fs=require("fs");const{unzipSync}=require("fflate");
const u=unzipSync(new Uint8Array(fs.readFileSync(process.argv[2])));
const text=[...new TextDecoder().decode(u["word/document.xml"]).matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(m=>m[1]).join("");
let i=text.indexOf("\\frac");
while(i>=0){ console.log("…"+text.slice(Math.max(0,i-40),i+60).replace(/\s+/g," ")+"…"); i=text.indexOf("\\frac", i+1); }
