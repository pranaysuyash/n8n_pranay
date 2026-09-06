import { scanPackage } from './pptx-scan.mjs';

const $=id=>document.getElementById(id);
const state={report:null,fileName:'',selected:0,urls:new Map()};
const stageMessages=[
  ['Opening the package','PowerPoint files are ZIP packages of XML, media, relationships, and metadata. GhostLayer is reading them locally.'],
  ['Comparing what you see with what is embedded','A crop changes the visible rectangle. It does not necessarily replace the original image bytes.'],
  ['Looking behind presentation mode','Hidden slides and speaker notes remain separate parts inside the package.'],
  ['Expanding the slide canvas','Objects can sit completely outside the visible slide and still remain in the file.'],
  ['Reading the file’s identity trail','Core properties can preserve author and last-editor information after the deck is shared.']
];

function mimeFor(path){const e=path.split('.').pop().toLowerCase();return e==='png'?'image/png':e==='gif'?'image/gif':e==='svg'?'image/svg+xml':e==='webp'?'image/webp':'image/jpeg'}
function mediaUrl(path){if(state.urls.has(path))return state.urls.get(path);const m=state.report.media.find(x=>x.path===path);if(!m)return'';const u=URL.createObjectURL(new Blob([m.bytes],{type:mimeFor(path)}));state.urls.set(path,u);return u}
function clearUrls(){for(const u of state.urls.values())URL.revokeObjectURL(u);state.urls.clear()}
function toast(s){$('toast').textContent=s;$('toast').classList.add('show');setTimeout(()=>$('toast').classList.remove('show'),1800)}
function esc(s=''){return String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function labelFor(t){return {'crop':'CROP X-RAY','hidden-slide':'HIDDEN SLIDE','notes':'PRESENTER NOTES','off-canvas':'OFF-CANVAS OBJECT','metadata':'IDENTITY METADATA'}[t]||t.toUpperCase()}
function blurb(f){switch(f.type){case'crop':return `${f.hiddenPercent.toFixed(1)}% of the embedded image sits outside the visible crop.`;case'hidden-slide':return 'This slide is excluded from slideshow playback, but its content remains in the PPTX.';case'notes':return 'Presenter-only notes remain attached to the slide inside the file package.';case'off-canvas':return 'This object is positioned beyond the visible slide boundary.';case'metadata':return 'The file preserves document identity properties that are not shown in slideshow mode.';default:return''}}

async function scanBuffer(buffer,name){
  if(!window.fflate){toast('ZIP engine did not load. Check your connection and retry.');return}
  clearUrls();
  state.fileName=name;
  $('hero').hidden=true;$('stage').hidden=false;$('report').hidden=true;$('scanView').hidden=false;
  $('fileLabel').textContent=name.toUpperCase(); $('reportTitle').textContent='Recoverable state';
  document.querySelectorAll('.scan-steps li').forEach(x=>x.className=''); $('scanProgress').style.width='0%';
  let files;
  try{files=window.fflate.unzipSync(new Uint8Array(buffer))}catch(e){toast('That file is not a readable PPTX package.');reset();return}
  for(let i=0;i<stageMessages.length;i++){
    $('scanHeadline').textContent=stageMessages[i][0];$('scanDetail').textContent=stageMessages[i][1];$('scanProgress').style.width=`${(i+1)*20}%`;
    document.querySelectorAll('.scan-steps li').forEach((el,j)=>{el.className=j<i?'done':j===i?'active':''});
    await new Promise(r=>setTimeout(r,i===0?420:330));
  }
  try{state.report=scanPackage(files)}catch(e){console.error(e);toast('The PPTX opened, but its structure was not recognized.');reset();return}
  renderReport();
}

function renderReport(){
  const r=state.report;$('scanView').hidden=true;$('report').hidden=false;
  $('totalFindings').textContent=r.totalFindings;$('criticalCount').textContent=r.counts.critical;$('warningCount').textContent=r.counts.warning;$('infoCount').textContent=r.counts.info;$('slideCount').textContent=`${r.slides.length} slides`;
  $('summarySentence').textContent=r.totalFindings?`${r.totalFindings} recoverable layer${r.totalFindings===1?'':'s'} remain inside ${state.fileName}.`:'No supported hidden-state patterns were found in this deck.';
  const list=$('findingList');list.innerHTML='';
  if(!r.findings.length){list.innerHTML='<div class="finding-card"><h4>No findings</h4><p>Try the planted specimen to see the experiment.</p></div>';$('findingBody').innerHTML='<p class="body-copy">No supported recoverable state was detected. GhostLayer currently checks cropped pictures, hidden slides, presenter notes, off-canvas shapes, and core document metadata.</p>';return}
  r.findings.forEach((f,i)=>{const el=document.createElement('div');el.className='finding-card'+(i===0?' active':'');el.innerHTML=`<div class="finding-row"><div><h4>${esc(f.title)}</h4><p>${esc(blurb(f))}</p></div><span class="mini-risk ${f.severity}">${f.severity}</span></div>`;el.onclick=()=>selectFinding(i);list.appendChild(el)});
  selectFinding(0);
  window.scrollTo({top:$('stage').offsetTop-12,behavior:'smooth'});
}

function selectFinding(i){
  state.selected=i;document.querySelectorAll('.finding-card').forEach((x,j)=>x.classList.toggle('active',j===i));const f=state.report.findings[i];
  $('findingType').textContent=labelFor(f.type);$('findingTitle').textContent=f.title;$('riskBadge').textContent=f.severity.toUpperCase();$('riskBadge').dataset.risk=f.severity;
  $('findingBody').innerHTML=bodyFor(f);
  if(f.type==='crop')requestAnimationFrame(()=>setupCrop(f));
}

function bodyFor(f){
  if(f.type==='crop'){
    return `<p class="body-copy">The slide references the full embedded image and uses an OOXML <code>a:srcRect</code> to display only part of it. A recipient with the PPTX can recover the original media.</p>
    <div class="crop-compare"><div class="compare-card"><div class="compare-label"><span>WHAT THE SLIDE SHOWS</span><span>VISIBLE CROP</span></div><div class="media-stage"><div class="crop-viewport" id="cropViewport"><img id="cropImage" alt="Visible cropped portion"></div></div></div><div class="compare-card"><div class="compare-label"><span>WHAT THE FILE CONTAINS</span><span>FULL EMBEDDED IMAGE</span></div><div class="media-stage"><div class="full-image" id="fullImage"><img id="fullImageEl" alt="Full recoverable embedded image"><div class="visible-box" id="visibleBox"></div></div></div></div>
    <div class="metric-strip"><div><span>Hidden by crop</span><b>${f.hiddenPercent.toFixed(1)}%</b></div><div><span>Embedded asset</span><b>${esc(f.mediaPath.split('/').pop())}</b></div><div><span>Slide</span><b>${f.slide}</b></div></div>`;
  }
  if(f.type==='hidden-slide')return `<p class="body-copy">PowerPoint’s slide XML can carry <code>show="0"</code>. Presentation mode skips the slide, but the part remains in the package with its text and objects intact.</p><div class="text-reveal"><blockquote>${esc(f.text||'Hidden slide contains non-text objects.')}</blockquote><small>Recovered from ppt/slides/slide${f.slide}.xml</small></div>`;
  if(f.type==='notes')return `<p class="body-copy">Speaker notes are stored as dedicated NotesSlide parts. They are invisible to an audience watching the slideshow but travel with the file.</p><div class="text-reveal"><blockquote>${esc(f.text)}</blockquote><small>Presenter notes · slide ${f.slide}</small></div>`;
  if(f.type==='off-canvas')return `<p class="body-copy">The object’s coordinates place its entire bounding box beyond the visible slide dimensions. It is off-stage, not deleted.</p><div class="canvas-demo"><div class="canvas-world"><div class="slide-outline"></div><div class="off-object"><small>OFF-CANVAS OBJECT · ${esc(f.name)}</small>${esc(f.text||'Object contains no text')}</div></div></div><div class="metric-strip"><div><span>Object X</span><b>${Math.round(f.geometry.x/914400)} in</b></div><div><span>Slide width</span><b>${Math.round(state.report.size.cx/914400)} in</b></div><div><span>Slide</span><b>${f.slide}</b></div></div>`;
  if(f.type==='metadata'){const m=state.report.metadata;return `<p class="body-copy">Core document properties can preserve authorship and editing identity even when none of it appears on a slide.</p><div class="metadata-table"><div><span>Document title</span><b>${esc(m.title||'—')}</b></div><div><span>Author</span><b>${esc(m.creator||'—')}</b></div><div><span>Last modified by</span><b>${esc(m.lastModifiedBy||'—')}</b></div></div>`}
  return `<p class="body-copy">${esc(blurb(f))}</p>`;
}

function setupCrop(f){
  const url=mediaUrl(f.mediaPath);const full=$('fullImageEl'),crop=$('cropImage');if(!url||!full||!crop)return;full.src=url;crop.src=url;
  const c=f.crop;full.onload=()=>{
    const stage=full.parentElement, sw=stage.clientWidth, sh=stage.clientHeight, iw=full.naturalWidth, ih=full.naturalHeight, scale=Math.min(sw*.95/iw,sh*.95/ih), rw=iw*scale,rh=ih*scale,left=(sw-rw)/2,top=(sh-rh)/2;
    const vb=$('visibleBox');vb.style.left=`${left+rw*c.left/100}px`;vb.style.top=`${top+rh*c.top/100}px`;vb.style.width=`${rw*(1-(c.left+c.right)/100)}px`;vb.style.height=`${rh*(1-(c.top+c.bottom)/100)}px`;
  };
  crop.onload=()=>{
    const vp=$('cropViewport'), iw=crop.naturalWidth, ih=crop.naturalHeight;const visibleW=iw*(1-(c.left+c.right)/100),visibleH=ih*(1-(c.top+c.bottom)/100);const scale=Math.max(vp.clientWidth/visibleW,vp.clientHeight/visibleH);crop.style.width=`${iw*scale}px`;crop.style.height=`${ih*scale}px`;crop.style.left=`${-iw*scale*c.left/100 + (vp.clientWidth-visibleW*scale)/2}px`;crop.style.top=`${-ih*scale*c.top/100 + (vp.clientHeight-visibleH*scale)/2}px`;
  };
}

function reset(){clearUrls();state.report=null;state.fileName='';$('stage').hidden=true;$('hero').hidden=false;window.scrollTo({top:0,behavior:'smooth'})}
function makeSpecimenFiles(){
  const E=s=>new TextEncoder().encode(s);
  const image=`<svg xmlns="http://www.w3.org/2000/svg" width="1600" height="900" viewBox="0 0 1600 900"><rect width="1600" height="900" fill="#102430"/><rect x="0" width="620" height="900" fill="#731f2d"/><text x="70" y="120" font-family="Arial" font-size="54" font-weight="700" fill="#ffd9dc">CONFIDENTIAL</text><text x="70" y="205" font-family="Arial" font-size="48" font-weight="700" fill="white">Project MICA</text><text x="70" y="290" font-family="Arial" font-size="27" fill="#ffd9dc">Acquisition target</text><text x="70" y="335" font-family="Arial" font-size="37" font-weight="700" fill="white">Northstar Labs</text><text x="70" y="430" font-family="Arial" font-size="27" fill="#ffd9dc">Offer ceiling</text><text x="70" y="485" font-family="Arial" font-size="50" font-weight="700" fill="white">$18.5M</text><text x="70" y="615" font-family="Arial" font-size="24" fill="#ffd9dc">Do not circulate outside board.</text><text x="720" y="115" font-family="Arial" font-size="55" font-weight="700" fill="white">Q4 Growth Outlook</text><text x="720" y="170" font-family="Arial" font-size="25" fill="#b9d3da">Board preview · September 2026</text><text x="720" y="260" font-family="Arial" font-size="27" fill="#b9d3da">Revenue trajectory</text><g fill="#9ee3c2"><rect x="760" y="570" width="95" height="200" rx="18"/><rect x="940" y="500" width="95" height="270" rx="18"/><rect x="1120" y="435" width="95" height="335" rx="18"/><rect x="1300" y="350" width="95" height="420" rx="18"/></g></svg>`;
  return {
    '[Content_Types].xml':E(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="svg" ContentType="image/svg+xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`),
    'ppt/presentation.xml':E(`<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="12192000" cy="6858000"/></p:presentation>`),
    'ppt/slides/slide1.xml':E(`<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:pic><p:nvPicPr><p:cNvPr id="4" name="Board screenshot"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdImg"/><a:srcRect l="38750"/></p:blipFill></p:pic><p:sp><p:nvSpPr><p:cNvPr id="7" name="Fallback target"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="13500000" y="1800000"/><a:ext cx="2600000" cy="800000"/></a:xfrm></p:spPr><p:txBody><a:p><a:r><a:t>Client fallback: if Project MICA stalls, approach Helios immediately.</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`),
    'ppt/slides/_rels/slide1.xml.rels':E(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.svg"/><Relationship Id="rIdNotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/></Relationships>`),
    'ppt/notesSlides/notesSlide1.xml':E(`<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>If asked about the margin dip, do not mention Project MICA until legal clears the announcement.</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>`),
    'ppt/slides/slide2.xml':E(`<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" show="0"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Project MICA · acquisition scenario</a:t></a:r></a:p><a:p><a:r><a:t>Target: Northstar Labs · Offer ceiling: $18.5M · Board vote: 4–1 preliminary · confidential / not for distribution</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`),
    'ppt/slides/_rels/slide2.xml.rels':E(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`),
    'ppt/slides/slide3.xml':E(`<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Operating priorities</a:t></a:r></a:p><a:p><a:r><a:t>Improve conversion · Reduce infrastructure cost · Expand partner channel</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`),
    'ppt/slides/_rels/slide3.xml.rels':E(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`),
    'ppt/media/image1.svg':E(image),
    'docProps/core.xml':E(`<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Board Meeting — Final External</dc:title><dc:creator>Board Strategy Team</dc:creator><cp:lastModifiedBy>M. Chen / CorpDev</cp:lastModifiedBy></cp:coreProperties>`)
  }
}
function specimenZip(){return window.fflate.zipSync(makeSpecimenFiles(),{level:6})}
async function loadSpecimen(){try{const z=specimenZip();await scanBuffer(z.buffer.slice(z.byteOffset,z.byteOffset+z.byteLength),'board-meeting-final.pptx')}catch(e){console.error(e);toast('Could not load the planted specimen.')}}

$('specimenBtn').onclick=loadSpecimen;$('scanAnother').onclick=reset;$('fileInput').onchange=async e=>{const f=e.target.files?.[0];if(!f)return;if(!f.name.toLowerCase().endsWith('.pptx'))return toast('Choose a .pptx file.');await scanBuffer(await f.arrayBuffer(),f.name)};$('downloadSpecimen').onclick=()=>{const a=document.createElement('a');const u=URL.createObjectURL(new Blob([specimenZip()],{type:'application/vnd.openxmlformats-officedocument.presentationml.presentation'}));a.href=u;a.download='ghostlayer-specimen.pptx';a.click();setTimeout(()=>URL.revokeObjectURL(u),500)};
