const fs=require("fs");const{unzipSync}=require("fflate");
const md=JSON.parse(fs.readFileSync("_tbl.json","utf8"));
const u=unzipSync(new Uint8Array(fs.readFileSync(process.argv[2])));
const d=new TextDecoder().decode(u["word/document.xml"]);
const tbls=d.split("<w:tbl>").slice(1);
const dx=tbls.map(t=>{
  const rows=(t.match(/<w:tr>/g)||[]).length;
  const fr=t.match(/<w:tr>[\s\S]*?<\/w:tr>/);
  const cols=fr?(fr[0].match(/<w:tc>/g)||[]).length:0;
  return {rows,cols};
});
console.log("思源块 "+md.count+" / docx 表 "+dx.length);
let bad=0;
md.blocks.forEach((b,i)=>{
  const a=dx[i]; if(!a){console.log("缺表 #"+(i+1));bad++;return;}
  const dr=a.rows-b.rows, dc=a.cols-b.cols;
  if(dr!==0||Math.abs(dc)>1){ bad++; console.log(`#${i+1} ${b.doc} 思源 ${b.rows}x${b.cols} → docx ${a.rows}x${a.cols}`); }
});
console.log("行列不一致的表: "+bad+" / "+md.count);
const sumMd=md.blocks.reduce((s,b)=>s+b.rows,0), sumDx=dx.reduce((s,b)=>s+b.rows,0);
console.log("总行数 思源 "+sumMd+" / docx "+sumDx);
console.log("首 8 表对照: "+md.blocks.slice(0,8).map((b,i)=>`${b.rows}x${b.cols}/${dx[i].rows}x${dx[i].cols}`).join("  "));
