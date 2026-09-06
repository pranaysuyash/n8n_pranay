import test from 'node:test';
import assert from 'node:assert/strict';
import {
  parseCrop,
  isSlideHidden,
  extractTextRuns,
  findOffCanvasShapes,
  getNotesText,
  parseCoreMetadata,
  scoreSeverity,
} from '../ghost-core.mjs';

test('parseCrop converts OOXML srcRect percentages and reports hidden fraction', () => {
  const xml = `<p:pic><p:blipFill><a:srcRect l="35000" t="5000" r="10000" b="0"/></p:blipFill></p:pic>`;
  const crop = parseCrop(xml);
  assert.deepEqual(crop, { left: 35, top: 5, right: 10, bottom: 0, hiddenPercent: 47.75, visiblePercent: 52.25 });
});

test('isSlideHidden detects show="0" and defaults visible when omitted', () => {
  assert.equal(isSlideHidden(`<p:sld xmlns:p="x" show="0"><p:cSld/></p:sld>`), true);
  assert.equal(isSlideHidden(`<p:sld xmlns:p="x"><p:cSld/></p:sld>`), false);
});

test('extractTextRuns returns decoded text in document order', () => {
  const xml = `<a:p><a:r><a:t>Secret &amp; internal</a:t></a:r></a:p><a:p><a:r><a:t>Q4</a:t></a:r></a:p>`;
  assert.deepEqual(extractTextRuns(xml), ['Secret & internal', 'Q4']);
});

test('findOffCanvasShapes flags shapes wholly outside slide bounds', () => {
  const slide = `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Hidden aside"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="10000000" y="100000"/><a:ext cx="1000000" cy="500000"/></a:xfrm></p:spPr><p:txBody><a:p><a:r><a:t>Do not show client</a:t></a:r></a:p></p:txBody></p:sp>`;
  const items = findOffCanvasShapes(slide, 9144000, 5143500);
  assert.equal(items.length, 1);
  assert.equal(items[0].name, 'Hidden aside');
  assert.match(items[0].text, /Do not show client/);
});

test('getNotesText ignores placeholder labels and keeps presenter text', () => {
  const notes = `<p:notes><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>If asked about margin, skip slide 9.</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:nvPr><p:ph type="sldNum"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>1</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>`;
  assert.deepEqual(getNotesText(notes), ['If asked about margin, skip slide 9.']);
});

test('parseCoreMetadata exposes author and last modifier', () => {
  const core = `<cp:coreProperties xmlns:cp="x" xmlns:dc="y"><dc:creator>Pranay Demo</dc:creator><cp:lastModifiedBy>Board Ops</cp:lastModifiedBy></cp:coreProperties>`;
  assert.deepEqual(parseCoreMetadata(core), { creator: 'Pranay Demo', lastModifiedBy: 'Board Ops', title: '' });
});

test('scoreSeverity maps crop and hidden slide findings to useful levels', () => {
  assert.equal(scoreSeverity({ type: 'crop', hiddenPercent: 65 }), 'critical');
  assert.equal(scoreSeverity({ type: 'crop', hiddenPercent: 12 }), 'warning');
  assert.equal(scoreSeverity({ type: 'hidden-slide' }), 'warning');
  assert.equal(scoreSeverity({ type: 'metadata' }), 'info');
});
