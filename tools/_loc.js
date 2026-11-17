const fs=require("fs");const{unzipSync}=require("fflate");
const u=unzipSync(new Uint8Array(fs.readFileSync("C:/Users/Taotao Chen/DoubaoWork/chats/2026-11-17/new-chat-7/docx-export-plugin/book1.docx")));
const d=new TextDecoder().decode(u["word/document.xml"]);
const tblRanges=[...d.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)].map(m=>[m.index,m.index+m[0].length]);
const starts=[...d.matchAll(/<w:p(?:\s[^>]*)?>/g)].map(m=>m.index);
let shown=0;
starts.forEach((s,i)=>{
  const p=d.slice(s, i+1<starts.length?starts[i+1]:s+3000);
  if(!/\$[^$]{1,80}\$/.test(p)) return;
  if(!/[\\_^]/.test(p)) return;
  const inTbl=tblRanges.some(([a,b])=>s>=a&&s<=b);
  const st=(p.match(/<w:pStyle w:val="([^"]+)"/)||[])[1]||"正文";
  const t=[...p.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(x=>x[1]).join("");
  if(!/\$[^$]{1,80}\$/m.test(t)) return;
  shown++;
  if(shown<=12) console.log((inTbl?"[表格内]":"[正文]")+" "+st+" | "+t.slice(0,72));
});
console.log("命中段落="+shown);
