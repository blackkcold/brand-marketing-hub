#!/usr/bin/env node
"use strict";
/**
 * v4.1 deterministic PPTX fallback.
 * Usage:
 * node runtime/pptxgenjs/render.js deck_spec.json output.pptx [assets.json] [template-manifest.json]
 *
 * This renderer is intentionally conservative. Host-native presentation runtimes
 * remain preferred because they can preserve real reference/master decks.
 */
const fs=require("fs");
const path=require("path");
const pptxgen=require("pptxgenjs");

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
const INK="111111", MUTED="565656", WHITE="FFFFFF", LIGHT="EAF2FF";
const SW=13.333, SH=7.5;

function addChrome(slide,slideId){
  slide.addText(slideId,{x:12.15,y:7.03,w:.55,h:.18,fontFace:font,fontSize:8,color:MUTED,align:"right",margin:0});
}
function addTitle(slide,takeaway){
  slide.addText(takeaway,{x:.62,y:.42,w:11.8,h:.62,fontFace:font,fontSize:22,bold:true,color:NAVY,margin:0,breakLine:false,fit:"shrink"});
}
function safeText(v){ return v===null||v===undefined?"":String(v); }
function addBlock(slide,b,cursor){
  const t=b.type;
  if(t==="headline"){
    slide.addText(safeText(b.text),{x:.7,y:cursor.y,w:11.9,h:.55,fontFace:font,fontSize:18,bold:true,color:INK,margin:0,fit:"shrink"});
    cursor.y+=.7;
  }else if(t==="body_text"){
    slide.addText(safeText(b.text),{x:.7,y:cursor.y,w:11.9,h:.85,fontFace:font,fontSize:12.5,color:INK,margin:.02,fit:"shrink",valign:"top"});
    cursor.y+=1.0;
  }else if(t==="bullets"){
    const runs=(b.items||[]).map(x=>({text:safeText(x),options:{bullet:{indent:14},hanging:3,breakLine:true}}));
    slide.addText(runs,{x:.75,y:cursor.y,w:11.7,h:Math.min(2.8,.42*Math.max(1,runs.length)),fontFace:font,fontSize:13,color:INK,margin:.02,breakLine:false,fit:"shrink",paraSpaceAfterPt:6});
    cursor.y+=Math.min(3.0,.44*Math.max(1,runs.length))+.12;
  }else if(t==="stat"){
    slide.addShape(pptx.ShapeType.roundRect,{x:.72,y:cursor.y,w:3.4,h:1.05,rectRadius:.06,fill:{color:LIGHT},line:{color:LIGHT}});
    slide.addText(safeText(b.value)+(b.unit||""),{x:.95,y:cursor.y+.12,w:2.95,h:.38,fontFace:font,fontSize:23,bold:true,color:BLUE,margin:0,fit:"shrink"});
    slide.addText(safeText(b.label),{x:.95,y:cursor.y+.6,w:2.95,h:.24,fontFace:font,fontSize:10.5,color:MUTED,margin:0,fit:"shrink"});
    cursor.y+=1.22;
  }else if(t==="callout"){
    slide.addShape(pptx.ShapeType.roundRect,{x:.7,y:cursor.y,w:11.9,h:.72,fill:{color:LIGHT},line:{color:BLUE,width:1}});
    slide.addText(safeText(b.text),{x:.95,y:cursor.y+.16,w:11.35,h:.32,fontFace:font,fontSize:13,bold:true,color:NAVY,margin:0,fit:"shrink"});
    cursor.y+=.9;
  }else if(t==="table"){
    const rows=[];
    if(Array.isArray(b.headers)) rows.push(b.headers.map(safeText));
    for(const r of (b.rows||[])) rows.push(r.map(safeText));
    if(rows.length){
      slide.addTable(rows,{x:.7,y:cursor.y,w:11.9,h:Math.min(3.5,.32*rows.length+.35),fontFace:font,fontSize:9.5,color:INK,
        border:{type:"solid",color:"D6E0F2",pt:.5},fill:WHITE,margin:.04,
        bold:false,autoFit:false,breakLine:false});
      cursor.y+=Math.min(3.7,.32*rows.length+.55);
    }
  }else if(t==="image"){
    const a=assets.get(b.asset_id);
    if(a&&a.local_path&&fs.existsSync(a.local_path)){
      slide.addImage({path:a.local_path,x:.75,y:cursor.y,w:5.7,h:3.2,sizing:"contain"});
      cursor.y+=3.35;
    }else{
      slide.addShape(pptx.ShapeType.rect,{x:.75,y:cursor.y,w:5.7,h:2.2,fill:{color:"F2F7FF"},line:{color:BLUE}});
      slide.addText("MISSING VERIFIED IMAGE",{x:1.0,y:cursor.y+.9,w:5.2,h:.3,fontFace:font,fontSize:11,bold:true,color:BLUE,align:"center",margin:0});
      cursor.y+=2.35;
    }
  }else if(t==="chart"){
    const d=b.data||{}; const labels=d.labels||[];
    const series=(d.series||[]).map(s=>({name:s.name||"",labels,values:s.values||[]}));
    const typeMap={bar:pptx.ChartType.bar,column:pptx.ChartType.bar,line:pptx.ChartType.line,pie:pptx.ChartType.pie,doughnut:pptx.ChartType.doughnut,area:pptx.ChartType.area,scatter:pptx.ChartType.scatter};
    if(series.length){
      slide.addChart(typeMap[b.chart_type]||pptx.ChartType.bar,series,{x:.75,y:cursor.y,w:11.6,h:3.4,showLegend:true,showTitle:false,showValue:false,catAxisLabelFontFace:font,valAxisLabelFontFace:font});
      cursor.y+=3.55;
    }
  }else if(t==="comparison"||t==="timeline"){
    const items=b.items||[];
    const cols=Math.min(Math.max(items.length,1),4); const gap=.16; const w=(11.9-gap*(cols-1))/cols;
    items.slice(0,cols).forEach((it,i)=>{
      slide.addShape(pptx.ShapeType.roundRect,{x:.7+i*(w+gap),y:cursor.y,w,h:1.55,fill:{color:"F7F9FC"},line:{color:"D6E0F2"}});
      slide.addText(safeText(it.title||it.label||it.time||""),{x:.88+i*(w+gap),y:cursor.y+.15,w:w-.36,h:.3,fontFace:font,fontSize:11,bold:true,color:NAVY,margin:0,fit:"shrink"});
      slide.addText(safeText(it.text||it.description||it.value||""),{x:.88+i*(w+gap),y:cursor.y+.55,w:w-.36,h:.72,fontFace:font,fontSize:9.5,color:INK,margin:0,fit:"shrink"});
    });
    cursor.y+=1.75;
  }
}
function makeCover(){
  const slide=pptx.addSlide();
  slide.background={color:BLUE};
  slide.addText(spec.deck.title,{x:.9,y:2.0,w:9.4,h:1.2,fontFace:font,fontSize:34,bold:true,color:WHITE,margin:0,fit:"shrink"});
  slide.addText(spec.deck.purpose||"",{x:.92,y:3.55,w:8.8,h:.6,fontFace:font,fontSize:16,color:"D1EBFE",margin:0,fit:"shrink"});
  slide.addText("vivo",{x:11.45,y:.42,w:1.1,h:.4,fontFace:"Arial",fontSize:20,bold:true,color:WHITE,margin:0,align:"right"});
}
makeCover();
for(const s of spec.slides){
  const slide=pptx.addSlide();
  slide.background={color:WHITE};
  addTitle(slide,s.takeaway||s.intent);
  const cursor={y:1.28};
  for(const b of (s.content||[])){
    if(cursor.y>6.65) break;
    addBlock(slide,b,cursor);
  }
  addChrome(slide,s.slide_id);
}
pptx.writeFile({fileName:outPath});
