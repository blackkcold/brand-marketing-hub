#!/usr/bin/env node
"use strict";
/**
 * v4.1 deterministic PPTX fallback.
 * Never-silent-loss is enforced by block chunking + continuation slides.
 *
 * Usage:
 * node runtime/pptxgenjs/render.js deck_spec.json output.pptx [assets.json] [template-manifest.json]
 */
const fs=require("fs");
const path=require("path");
const pptxgen=require("@lofcz/pptxgenjs");

const [,,deckPath,outPath,assetsPath,templatePath]=process.argv;
if(!deckPath||!outPath){
  console.error("usage: node render.js deck_spec.json output.pptx [assets.json] [template-manifest.json]");
  process.exit(2);
}
const load=p=>JSON.parse(fs.readFileSync(p,"utf8"));
const spec=load(deckPath);
if(spec.version!=="4.1") throw new Error("deck_spec version must be 4.1");
const assetManifest=assetsPath&&fs.existsSync(assetsPath)?load(assetsPath):{assets:[]};
const template=templatePath&&fs.existsSync(templatePath)?load(templatePath):{
  allowed_fonts:["微软雅黑"],allowed_palettes:{house:["1E46E6","06175E","D1EBFE","FFFFFF","111111"]}
};
const assets=new Map((assetManifest.assets||[]).map(a=>[a.asset_id,a]));
const assetBase=assetsPath?path.dirname(path.resolve(assetsPath)):process.cwd();

function resolveAssetPath(asset){
  if(!asset||!asset.local_path)return null;
  return path.isAbsolute(asset.local_path)?asset.local_path:path.resolve(assetBase,asset.local_path);
}
function fitAssetRect(asset,x,y,w,h){
  const iw=Number(asset&&asset.width), ih=Number(asset&&asset.height);
  if(!(iw>0&&ih>0))return {x,y,w,h};
  const scale=Math.min(w/iw,h/ih);
  const rw=iw*scale, rh=ih*scale;
  return {x:x+(w-rw)/2,y:y+(h-rh)/2,w:rw,h:rh};
}

const pptx=new pptxgen();
pptx.layout="LAYOUT_WIDE";
pptx.author="brand-marketing-hub v4.1";
pptx.subject=spec.deck.purpose||"";
pptx.title=spec.deck.title||"";
pptx.company="vivo";
pptx.lang="zh-CN";
pptx.theme={
  headFontFace:(template.allowed_fonts||[])[0]||"微软雅黑",
  bodyFontFace:(template.allowed_fonts||[])[0]||"微软雅黑",
  lang:"zh-CN"
};

const font=(template.allowed_fonts||[])[0]||"微软雅黑";
const house=(template.allowed_palettes&&template.allowed_palettes.house)||[];
const BLUE=(house[0]||"1E46E6").replace("#","");
const NAVY=(house.find(x=>String(x).toUpperCase().replace("#","")==="06175E")||"06175E").replace("#","");
const INK="111111",MUTED="565656",WHITE="FFFFFF",LIGHT="EAF2FF";
const CONTENT_TOP=1.28,CONTENT_BOTTOM=6.72;

function safeText(v){return v===null||v===undefined?"":String(v);}
function addChrome(slide,slideId){
  slide.addText(slideId,{x:12.05,y:7.02,w:.65,h:.18,fontFace:font,fontSize:8,color:MUTED,align:"right",margin:0});
}
function addTitle(slide,takeaway){
  slide.addText(takeaway,{x:.62,y:.42,w:11.8,h:.62,fontFace:font,fontSize:22,bold:true,color:NAVY,margin:0,fit:"shrink"});
}
function splitText(text,maxChars=520){
  const value=safeText(text);
  if(value.length<=maxChars)return [value];
  const chunks=[];
  let rest=value;
  while(rest.length>maxChars){
    let cut=-1;
    for(const mark of ["\n","。","；",";","，",","," "]){
      const idx=rest.lastIndexOf(mark,maxChars);
      if(idx>Math.floor(maxChars*.55)){cut=idx+mark.length;break;}
    }
    if(cut<1)cut=maxChars;
    chunks.push(rest.slice(0,cut));
    rest=rest.slice(cut);
  }
  if(rest)chunks.push(rest);
  return chunks;
}
function chunkBlock(b){
  if(!b||!b.type)throw new Error("deck block missing type");
  if(b.type==="body_text"){
    const parts=splitText(b.text,520);
    return parts.map((text,i)=>({...b,block_id:`${b.block_id||"BLK"}-p${i+1}`,text}));
  }
  if(b.type==="bullets"&&(b.items||[]).length>8){
    const out=[],items=b.items||[];
    for(let i=0;i<items.length;i+=8)out.push({...b,block_id:`${b.block_id||"BLK"}-p${out.length+1}`,items:items.slice(i,i+8)});
    return out;
  }
  if(b.type==="table"&&(b.rows||[]).length>9){
    const out=[],rows=b.rows||[];
    for(let i=0;i<rows.length;i+=9)out.push({...b,block_id:`${b.block_id||"BLK"}-p${out.length+1}`,rows:rows.slice(i,i+9)});
    return out;
  }
  if((b.type==="comparison"||b.type==="timeline")&&(b.items||[]).length>4){
    const out=[],items=b.items||[];
    for(let i=0;i<items.length;i+=4)out.push({...b,block_id:`${b.block_id||"BLK"}-p${out.length+1}`,items:items.slice(i,i+4)});
    return out;
  }
  return [b];
}
function estimateHeight(b){
  switch(b.type){
    case "headline":return .70;
    case "body_text":return 1.00;
    case "bullets":return Math.min(3.6,.44*Math.max(1,(b.items||[]).length))+.12;
    case "stat":return 1.22;
    case "callout":return .90;
    case "table":return Math.min(3.7,.32*((b.rows||[]).length+(Array.isArray(b.headers)?1:0))+.55);
    case "image":return 3.35;
    case "chart":return 3.55;
    case "comparison":
    case "timeline":return 1.75;
    case "source_footer":return .42;
    default:throw new Error(`unsupported deck block type: ${b.type}`);
  }
}
function addBlock(slide,b,cursor){
  const t=b.type;
  if(t==="headline"){
    slide.addText(safeText(b.text),{x:.7,y:cursor.y,w:11.9,h:.55,fontFace:font,fontSize:18,bold:true,color:INK,margin:0,fit:"shrink"});
    cursor.y+=.70;
  }else if(t==="body_text"){
    slide.addText(safeText(b.text),{x:.7,y:cursor.y,w:11.9,h:.85,fontFace:font,fontSize:12.5,color:INK,margin:.02,fit:"shrink",valign:"top"});
    cursor.y+=1.00;
  }else if(t==="bullets"){
    const runs=(b.items||[]).map(x=>({text:safeText(x),options:{bullet:{indent:14},hanging:3,breakLine:true}}));
    slide.addText(runs,{x:.75,y:cursor.y,w:11.7,h:Math.min(3.4,.42*Math.max(1,runs.length)),fontFace:font,fontSize:13,color:INK,margin:.02,fit:"shrink",paraSpaceAfterPt:6});
    cursor.y+=estimateHeight(b);
  }else if(t==="stat"){
    slide.addShape(pptx.ShapeType.roundRect,{x:.72,y:cursor.y,w:3.4,h:1.05,fill:{color:LIGHT},line:{color:LIGHT}});
    slide.addText(safeText(b.value)+(b.unit||""),{x:.95,y:cursor.y+.12,w:2.95,h:.38,fontFace:font,fontSize:23,bold:true,color:BLUE,margin:0,fit:"shrink"});
    slide.addText(safeText(b.label),{x:.95,y:cursor.y+.6,w:2.95,h:.24,fontFace:font,fontSize:10.5,color:MUTED,margin:0,fit:"shrink"});
    cursor.y+=1.22;
  }else if(t==="callout"){
    slide.addShape(pptx.ShapeType.roundRect,{x:.7,y:cursor.y,w:11.9,h:.72,fill:{color:LIGHT},line:{color:BLUE,width:1}});
    slide.addText(safeText(b.text),{x:.95,y:cursor.y+.16,w:11.35,h:.32,fontFace:font,fontSize:13,bold:true,color:NAVY,margin:0,fit:"shrink"});
    cursor.y+=.90;
  }else if(t==="table"){
    const rows=[];
    if(Array.isArray(b.headers))rows.push(b.headers.map(safeText));
    for(const row of (b.rows||[]))rows.push(row.map(safeText));
    if(rows.length){
      slide.addTable(rows,{x:.7,y:cursor.y,w:11.9,h:Math.min(3.5,.32*rows.length+.35),fontFace:font,fontSize:9.5,color:INK,
        border:{type:"solid",color:"D6E0F2",pt:.5},fill:WHITE,margin:.04,autoFit:false});
      cursor.y+=estimateHeight(b);
    }
  }else if(t==="image"){
    const a=assets.get(b.asset_id);
    const assetPath=resolveAssetPath(a);
    if(a&&assetPath&&fs.existsSync(assetPath)){
      const box=fitAssetRect(a,.75,cursor.y,5.7,3.2);
      slide.addImage({path:assetPath,...box});
    }else{
      slide.addShape(pptx.ShapeType.rect,{x:.75,y:cursor.y,w:5.7,h:2.2,fill:{color:"F2F7FF"},line:{color:BLUE}});
      slide.addText("MISSING VERIFIED IMAGE",{x:1.0,y:cursor.y+.9,w:5.2,h:.3,fontFace:font,fontSize:11,bold:true,color:BLUE,align:"center",margin:0});
    }
    cursor.y+=3.35;
  }else if(t==="chart"){
    const d=b.data||{},labels=d.labels||[];
    const series=(d.series||[]).map(s=>({name:s.name||"",labels,values:s.values||[]}));
    const typeMap={bar:pptx.ChartType.bar,column:pptx.ChartType.bar,line:pptx.ChartType.line,pie:pptx.ChartType.pie,
      doughnut:pptx.ChartType.doughnut,area:pptx.ChartType.area,scatter:pptx.ChartType.scatter};
    if(!series.length)throw new Error(`chart block ${b.block_id||""} has no series`);
    slide.addChart(typeMap[b.chart_type]||pptx.ChartType.bar,series,{x:.75,y:cursor.y,w:11.6,h:3.4,showLegend:true,showTitle:false,
      showValue:false,catAxisLabelFontFace:font,valAxisLabelFontFace:font});
    cursor.y+=3.55;
  }else if(t==="comparison"||t==="timeline"){
    const items=b.items||[];
    const cols=Math.max(items.length,1),gap=.16,w=(11.9-gap*(cols-1))/cols;
    items.forEach((it,i)=>{
      slide.addShape(pptx.ShapeType.roundRect,{x:.7+i*(w+gap),y:cursor.y,w,h:1.55,fill:{color:"F7F9FC"},line:{color:"D6E0F2"}});
      slide.addText(safeText(it.title||it.label||it.time||""),{x:.88+i*(w+gap),y:cursor.y+.15,w:w-.36,h:.3,fontFace:font,fontSize:11,bold:true,color:NAVY,margin:0,fit:"shrink"});
      slide.addText(safeText(it.text||it.description||it.value||""),{x:.88+i*(w+gap),y:cursor.y+.55,w:w-.36,h:.72,fontFace:font,fontSize:9.5,color:INK,margin:0,fit:"shrink"});
    });
    cursor.y+=1.75;
  }else if(t==="source_footer"){
    const sources=(b.source_ids||[]).map(safeText).join(" · ");
    slide.addText(sources,{x:.75,y:cursor.y,w:11.4,h:.22,fontFace:font,fontSize:8.5,color:MUTED,margin:0,fit:"shrink"});
    cursor.y+=.42;
  }else{
    throw new Error(`unsupported deck block type: ${t}`);
  }
}
function makeCover(){
  const slide=pptx.addSlide();
  slide.background={color:BLUE};
  slide.addText(spec.deck.title,{x:.9,y:2.0,w:9.4,h:1.2,fontFace:font,fontSize:34,bold:true,color:WHITE,margin:0,fit:"shrink"});
  slide.addText(spec.deck.purpose||"",{x:.92,y:3.55,w:8.8,h:.6,fontFace:font,fontSize:16,color:"D1EBFE",margin:0,fit:"shrink"});
  const logoPath=path.resolve(__dirname,"../../assets/vivo-deck/vivo_wordmark_white.png");
  if(fs.existsSync(logoPath))slide.addImage({path:logoPath,x:11.45,y:.42,w:1.15,h:.36});
  else slide.addText("vivo",{x:11.45,y:.42,w:1.1,h:.4,fontFace:"Arial",fontSize:20,bold:true,color:WHITE,margin:0,align:"right"});
}
function newContentSlide(s,part){
  const slide=pptx.addSlide();
  slide.background={color:WHITE};
  addTitle(slide,(s.takeaway||s.intent)+(part>1?`（续 ${part}）`:""));
  return {slide,cursor:{y:CONTENT_TOP},part};
}

makeCover();
for(const s of spec.slides){
  const blocks=(s.content||[]).flatMap(chunkBlock);
  let state=newContentSlide(s,1);
  for(const b of blocks){
    const h=estimateHeight(b);
    if(h>CONTENT_BOTTOM-CONTENT_TOP+.01){
      throw new Error(`block ${b.block_id||""} cannot fit a single slide even after chunking`);
    }
    if(state.cursor.y+h>CONTENT_BOTTOM&&state.cursor.y>CONTENT_TOP+.01){
      addChrome(state.slide,`${s.slide_id}.${state.part}`);
      state=newContentSlide(s,state.part+1);
    }
    addBlock(state.slide,b,state.cursor);
  }
  addChrome(state.slide,state.part===1?s.slide_id:`${s.slide_id}.${state.part}`);
}

(async()=>{
  await pptx.writeFile({fileName:outPath});
  console.log(`PASS render: ${outPath}`);
})().catch(err=>{console.error(err);process.exit(1);});
