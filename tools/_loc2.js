const fs=require("fs");const{unzipSync}=require("fflate");
const u=unzipSync(new Uint8Array(fs.readFileSync(process.argv[2])));
const d=new TextDecoder().decode(u["word/document.xml"]);
const text=[...d.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(m=>m[1]).join("");
const re=/\$([^$]{1,80})\$/g;let m,n=0;
while((m=re.exec(text))){
  if(!/[\\_^]/.test(m[1])) continue;
  n++;
  const ctx=text.slice(Math.max(0,m.index-28), m.index+52).replace(/\s+/g," ");
  console.log(n+" | …"+ctx+"…");
}
console.log("合计 "+n);
