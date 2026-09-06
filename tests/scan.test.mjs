import test from 'node:test';
import assert from 'node:assert/strict';
import { scanPackage } from '../pptx-scan.mjs';

const enc=s=>new TextEncoder().encode(s);

test('scanPackage links cropped image, notes, hidden slide, off-canvas shape and metadata',()=>{
  const files={
    'ppt/presentation.xml':enc(`<p:presentation><p:sldSz cx="9144000" cy="5143500"/></p:presentation>`),
    'ppt/slides/slide1.xml':enc(`<p:sld><p:cSld><p:spTree><p:pic><p:nvPicPr><p:cNvPr id="4" name="Board screenshot"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdImg"/><a:srcRect l="38750"/></p:blipFill></p:pic><p:sp><p:nvSpPr><p:cNvPr id="7" name="Offstage"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="10000000" y="100000"/><a:ext cx="800000" cy="400000"/></a:xfrm></p:spPr><p:txBody><a:p><a:r><a:t>Fallback target Helios</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`),
    'ppt/slides/_rels/slide1.xml.rels':enc(`<Relationships><Relationship Id="rIdImg" Type="x/image" Target="../media/image1.png"/><Relationship Id="rIdNotes" Type="x/notesSlide" Target="../notesSlides/notesSlide1.xml"/></Relationships>`),
    'ppt/notesSlides/notesSlide1.xml':enc(`<p:notes><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>Keep MICA confidential.</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>`),
    'ppt/slides/slide2.xml':enc(`<p:sld show="0"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Project MICA</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`),
    'ppt/slides/_rels/slide2.xml.rels':enc(`<Relationships/>`),
    'ppt/media/image1.png':new Uint8Array([137,80,78,71]),
    'docProps/core.xml':enc(`<cp:coreProperties><dc:title>External Board Deck</dc:title><dc:creator>Board Team</dc:creator><cp:lastModifiedBy>CorpDev</cp:lastModifiedBy></cp:coreProperties>`),
  };
  const result=scanPackage(files);
  assert.equal(result.slides.length,2);
  assert.equal(result.findings.filter(x=>x.type==='crop').length,1);
  assert.equal(result.findings.find(x=>x.type==='crop').hiddenPercent,38.75);
  assert.equal(result.findings.filter(x=>x.type==='notes').length,1);
  assert.equal(result.findings.filter(x=>x.type==='hidden-slide').length,1);
  assert.equal(result.findings.filter(x=>x.type==='off-canvas').length,1);
  assert.equal(result.metadata.creator,'Board Team');
  assert.equal(result.media[0].path,'ppt/media/image1.png');
});
