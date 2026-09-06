import { parseCrop, isSlideHidden, extractTextRuns, findOffCanvasShapes, getNotesText, parseCoreMetadata, scoreSeverity } from './ghost-core.mjs';

const td=new TextDecoder();
const txt=v=>v?td.decode(v):'';
function norm(path){
  const out=[]; for(const p of path.split('/')){ if(!p||p==='.') continue; if(p==='..') out.pop(); else out.push(p); } return out.join('/');
}
function resolve(base,target){ const dir=base.split('/').slice(0,-1).join('/'); return norm(dir+'/'+target); }
function parseRels(xml=''){
  const out={};
  for(const m of xml.matchAll(/<Relationship\b([^>]*)\/?\s*>/g)){
    const attrs={}; for(const a of m[1].matchAll(/\b(Id|Type|Target)="([^"]*)"/g)) attrs[a[1]]=a[2];
    if(attrs.Id) out[attrs.Id]=attrs;
  }
  return out;
}
function slideNum(path){return Number(path.match(/slide(\d+)\.xml$/)?.[1]||0)}
function slideSize(xml){const m=xml.match(/<p:sldSz\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"/);return m?{cx:Number(m[1]),cy:Number(m[2])}:{cx:9144000,cy:5143500}}
function pictureBlocks(xml=''){return [...xml.matchAll(/<p:pic\b[\s\S]*?<\/p:pic>/g)].map(m=>m[0])}
function picName(block){return block.match(/<p:cNvPr\b[^>]*name="([^"]*)"/)?.[1]||'Picture'}
function picRid(block){return block.match(/<a:blip\b[^>]*r:embed="([^"]+)"/)?.[1]||''}

export function scanPackage(files){
  const pres=txt(files['ppt/presentation.xml']); const size=slideSize(pres);
  const slidePaths=Object.keys(files).filter(p=>/^ppt\/slides\/slide\d+\.xml$/.test(p)).sort((a,b)=>slideNum(a)-slideNum(b));
  const findings=[],slides=[],media=[]; const seenMedia=new Set();
  for(const path of slidePaths){
    const num=slideNum(path), xml=txt(files[path]);
    const relPath=`ppt/slides/_rels/slide${num}.xml.rels`; const rels=parseRels(txt(files[relPath]));
    const hidden=isSlideHidden(xml); const slide={number:num,hidden,text:extractTextRuns(xml),pictures:[],notes:[],offCanvas:[]};
    if(hidden){const f={type:'hidden-slide',slide:num,title:`Hidden slide ${num}`,text:slide.text.join(' · ')}; f.severity=scoreSeverity(f); findings.push(f)}
    for(const block of pictureBlocks(xml)){
      const crop=parseCrop(block), rid=picRid(block), rel=rels[rid];
      if(!crop||!rel) continue;
      const mediaPath=resolve(path,rel.Target); const bytes=files[mediaPath];
      const pic={name:picName(block),crop,rid,mediaPath}; slide.pictures.push(pic);
      if(bytes && !seenMedia.has(mediaPath)){media.push({path:mediaPath,bytes});seenMedia.add(mediaPath)}
      if(crop.hiddenPercent>0){const f={type:'crop',slide:num,title:`Recoverable crop on slide ${num}`,name:pic.name,mediaPath,hiddenPercent:crop.hiddenPercent,crop};f.severity=scoreSeverity(f);findings.push(f)}
    }
    const noteRel=Object.values(rels).find(r=>/\/notesSlide$/i.test(r.Type||''));
    if(noteRel){const np=resolve(path,noteRel.Target), notes=getNotesText(txt(files[np])); slide.notes=notes; if(notes.length){const f={type:'notes',slide:num,title:`Speaker notes on slide ${num}`,text:notes.join('\n')};f.severity=scoreSeverity(f);findings.push(f)}}
    const off=findOffCanvasShapes(xml,size.cx,size.cy); slide.offCanvas=off;
    for(const item of off){const f={type:'off-canvas',slide:num,title:`Off-slide object on slide ${num}`,name:item.name,text:item.text,geometry:item};f.severity=scoreSeverity(f);findings.push(f)}
    slides.push(slide);
  }
  const metadata=parseCoreMetadata(txt(files['docProps/core.xml']));
  if(metadata.creator||metadata.lastModifiedBy){const f={type:'metadata',title:'Document identity metadata',text:[metadata.creator&&`Author: ${metadata.creator}`,metadata.lastModifiedBy&&`Last editor: ${metadata.lastModifiedBy}`].filter(Boolean).join(' · ')};f.severity=scoreSeverity(f);findings.push(f)}
  const counts={critical:0,warning:0,info:0}; for(const f of findings) counts[f.severity]++;
  return {slides,findings,media,metadata,size,counts,totalFindings:findings.length};
}
