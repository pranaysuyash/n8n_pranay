function decodeXml(s=''){
  return s.replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&amp;/g,'&');
}

export function parseCrop(xml=''){
  const m=xml.match(/<a:srcRect\b([^>]*)\/?\s*>/i);
  if(!m) return null;
  const attrs={};
  for(const [,k,v] of m[1].matchAll(/\b(l|t|r|b)="(-?\d+)"/g)) attrs[k]=Number(v)/1000;
  const left=attrs.l||0, top=attrs.t||0, right=attrs.r||0, bottom=attrs.b||0;
  const visibleWidth=Math.max(0,Math.min(100,100-left-right));
  const visibleHeight=Math.max(0,Math.min(100,100-top-bottom));
  const visiblePercent=Number((visibleWidth*visibleHeight/100).toFixed(4));
  const hiddenPercent=Number((100-visiblePercent).toFixed(4));
  return {left,top,right,bottom,hiddenPercent,visiblePercent};
}

export function isSlideHidden(xml=''){
  const m=xml.match(/<p:sld\b([^>]*)>/i);
  if(!m) return false;
  return /\bshow="(?:0|false)"/i.test(m[1]);
}

export function extractTextRuns(xml=''){
  return [...xml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map(m=>decodeXml(m[1])).filter(Boolean);
}

function attr(tag,name){const m=tag.match(new RegExp(`\\b${name}="(-?\\d+)"`));return m?Number(m[1]):0}
export function findOffCanvasShapes(xml='', slideCx=9144000, slideCy=5143500){
  const out=[];
  for(const m of xml.matchAll(/<p:sp\b[\s\S]*?<\/p:sp>/g)){
    const block=m[0];
    const off=block.match(/<a:off\b[^>]*>/), ext=block.match(/<a:ext\b[^>]*>/);
    if(!off||!ext) continue;
    const x=attr(off[0],'x'), y=attr(off[0],'y'), cx=attr(ext[0],'cx'), cy=attr(ext[0],'cy');
    const outside=(x+cx<=0)||(y+cy<=0)||(x>=slideCx)||(y>=slideCy);
    if(!outside) continue;
    const name=decodeXml(block.match(/<p:cNvPr\b[^>]*\bname="([^"]*)"/)?.[1]||'Off-canvas shape');
    out.push({name,text:extractTextRuns(block).join(' '),x,y,cx,cy});
  }
  return out;
}

export function getNotesText(xml=''){
  const out=[];
  for(const m of xml.matchAll(/<p:sp\b[\s\S]*?<\/p:sp>/g)){
    const block=m[0];
    const ph=block.match(/<p:ph\b[^>]*\btype="([^"]+)"/)?.[1]||'';
    if(['sldNum','hdr','ftr','dt'].includes(ph)) continue;
    const text=extractTextRuns(block).join(' ').trim();
    if(text) out.push(text);
  }
  return out;
}

function firstTag(xml,name){return decodeXml(xml.match(new RegExp(`<${name}>([\\s\\S]*?)<\\/${name}>`))?.[1]||'')}
export function parseCoreMetadata(xml=''){
  return {creator:firstTag(xml,'dc:creator'),lastModifiedBy:firstTag(xml,'cp:lastModifiedBy'),title:firstTag(xml,'dc:title')};
}

export function scoreSeverity(f){
  if(f.type==='crop') return (f.hiddenPercent||0)>=50?'critical':'warning';
  if(['hidden-slide','notes','off-canvas'].includes(f.type)) return 'warning';
  return 'info';
}
