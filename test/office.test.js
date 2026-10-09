'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { zipSync } = require('../vendor/node_modules/fflate');
const { analyzeOffice } = require('../src/office');
const Module = require('node:module');
const originalLoad = Module._load;
Module._load = function(request, parent, isMain) {
  if (request === 'vscode') return { EventEmitter: class { constructor() { this.event = () => ({ dispose() {} }); } fire() {} dispose() {} } };
  return originalLoad.call(this, request, parent, isMain);
};
const { JinSheetDocument, viewModel } = require('../src/extension');
Module._load = originalLoad;

function packageBytes(files) {
  return zipSync(Object.fromEntries(Object.entries(files).map(([name, text]) => [name, new TextEncoder().encode(text)])), { level: 6 });
}
function docx(body = '<w:p><w:r><w:t>Hello</w:t></w:r></w:p>') {
  return packageBytes({
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
    'word/document.xml': `<w:document xmlns:w="urn:word"><w:body>${body}<w:sectPr/></w:body></w:document>`
  });
}
function pptx() {
  return packageBytes({
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
    'ppt/presentation.xml': '<p:presentation xmlns:p="urn:p" xmlns:r="urn:r"><p:sldIdLst><p:sldId r:id="second"/><p:sldId r:id="first"/></p:sldIdLst></p:presentation>',
    'ppt/_rels/presentation.xml.rels': '<Relationships><Relationship Id="first" Type="urn/slide" Target="slides/slide1.xml"/><Relationship Id="second" Type="urn/slide" Target="slides/slide2.xml"/></Relationships>',
    'ppt/slides/slide1.xml': '<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><a:p><a:r><a:t>First</a:t></a:r></a:p></p:sld>',
    'ppt/slides/slide2.xml': '<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><a:p><a:r><a:t>Second</a:t></a:r></a:p><a:p><a:r><a:t>Line two</a:t></a:r></a:p></p:sld>'
  });
}
function uri(extension) { return { fsPath: `/tmp/preview${extension}`, scheme: 'file' }; }

test('DOCX와 DOCM에서 문단 텍스트를 추출하고 편집을 비활성화한다', () => {
  const bytes = docx('<w:p><w:r><w:t>Hello</w:t><w:tab/><w:t>Word</w:t></w:r></w:p><w:p><w:r><w:t>Next</w:t></w:r></w:p>');
  for (const extension of ['.docx', '.docm']) {
    const state = analyzeOffice(bytes, extension);
    assert.deepEqual(state.sections[0].paragraphs, ['Hello\tWord', 'Next']);
    const document = new JinSheetDocument(uri(extension), 'office', state, bytes, false, 'Office 문서는 텍스트 미리보기 전용입니다.');
    assert.equal(document.writable, false);
    assert.equal(viewModel(document).kind, 'office');
    assert.throws(() => document.applyCellEdit(0, 0, 'changed'), /전용/);
    document.dispose();
  }
});

test('PPTX와 PPTM은 프레젠테이션 순서로 슬라이드 텍스트를 반환한다', () => {
  const bytes = pptx();
  for (const extension of ['.pptx', '.pptm']) {
    const state = analyzeOffice(bytes, extension);
    assert.deepEqual(state.sections.map(section => section.paragraphs), [['Second', 'Line two'], ['First']]);
  }
});

test('Office ZIP 안의 DTD와 패키지 밖 관계 경로를 거부한다', () => {
  assert.throws(() => analyzeOffice(docx('<!DOCTYPE x [<!ENTITY e SYSTEM "file:///etc/passwd">]><w:p><w:r><w:t>&e;</w:t></w:r></w:p>'), '.docx'), /DTD|엔터티|XML/);
  const hostile = packageBytes({
    '[Content_Types].xml': '<Types/>',
    'ppt/presentation.xml': '<p:presentation xmlns:p="urn:p" xmlns:r="urn:r"><p:sldIdLst><p:sldId r:id="bad"/></p:sldIdLst></p:presentation>',
    'ppt/_rels/presentation.xml.rels': '<Relationships><Relationship Id="bad" Type="urn/slide" Target="../../outside.xml"/></Relationships>'
  });
  assert.throws(() => analyzeOffice(hostile, '.pptx'), /패키지 외부 경로/);
});

test('구형 바이너리 .doc/.ppt를 OOXML 파일로 오인하지 않는다', () => {
  assert.throws(() => analyzeOffice(Buffer.from('legacy'), '.doc'), /구형/);
  assert.throws(() => analyzeOffice(Buffer.from('legacy'), '.ppt'), /구형/);
});
