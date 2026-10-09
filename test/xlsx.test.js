'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { zipSync } = require('../vendor/node_modules/fflate');
const { analyze, inspectZip, patchCell } = require('../src/xlsx');

function fixture({ formula = false, extra = undefined, workbookXml = undefined } = {}) {
  const files = {
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': workbookXml || `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Hidden" sheetId="2" state="hidden" r:id="rId2"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`,
    'xl/worksheets/sheet1.xml': `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>before</t></is></c><c r="B1"><v>42</v></c></row>${formula ? '<row r="2"><c r="A2"><f>1+1</f><v>2</v></c></row>' : ''}</sheetData></worksheet>`,
    'xl/worksheets/sheet2.xml': `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>9</v></c></row></sheetData></worksheet>`
  };
  if (extra) files[extra.name] = extra.content;
  return zipSync(Object.fromEntries(Object.entries(files).map(([name, content]) => [name, new TextEncoder().encode(content)])), { level: 6 });
}

test('표준 단순 XLSX에서 여러 시트·값을 읽고 단일 셀만 수정한다', () => {
  const original = fixture();
  const model = analyze(original);
  assert.equal(model.writable, true);
  assert.equal(model.sheets.length, 2);
  assert.equal(model.sheets[0].cells.get('A1').value, 'before');
  const result = patchCell(model, 0, 'A1', 'after < &');
  assert.equal(result.model.sheets[0].cells.get('A1').value, 'after < &');
  for (const name of Object.keys(model.files)) {
    if (name !== 'xl/worksheets/sheet1.xml') assert.deepEqual(Buffer.from(result.model.files[name]), Buffer.from(model.files[name]));
  }
  assert.equal(result.model.sheets[0].cells.get('B1').value, '42');
});

test('수식이 있는 통합문서는 수식과 캐시값을 읽되 쓰기를 거부한다', () => {
  const model = analyze(fixture({ formula: true }));
  assert.equal(model.writable, false);
  assert.equal(model.sheets[0].cells.get('A2').formula, '1+1');
  assert.equal(model.sheets[0].cells.get('A2').cachedValue, '2');
  assert.throws(() => patchCell(model, 0, 'A1', 'blocked'), /수식이 포함/);
});

test('알 수 없는 패키지 구성요소는 표시를 허용하되 저장을 막는다', () => {
  const model = analyze(fixture({ extra: { name: 'xl/customXml/item1.xml', content: '<custom/>' } }));
  assert.equal(model.sheets[0].cells.get('A1').value, 'before');
  assert.equal(model.writable, false);
  assert.match(model.readOnlyReason, /지원하지 않는 OOXML/);
});

test('외부 엔터티와 DTD가 있는 XML을 거부한다', () => {
  const bytes = fixture();
  assert.throws(() => analyze(bytes.subarray(0, 10)), /ZIP/);
  const hostile = fixture({ workbookXml: '<!DOCTYPE x [<!ENTITY a SYSTEM "file:///etc/passwd">]><workbook>&a;</workbook>' });
  assert.throws(() => analyze(hostile), /XML .* 거부/);
});

test('중복·상위 경로 ZIP 항목, 크기 변조와 CRC 오류를 거부한다', () => {
  assert.throws(() => analyze(zipSync({ '../escape.xml': new Uint8Array([1]) })), /경로/);
  const sizeTampered = Buffer.from(zipSync({ 'part.xml': new TextEncoder().encode('<x/>') }));
  const directory = sizeTampered.indexOf(Buffer.from([0x50, 0x4b, 0x01, 0x02]));
  sizeTampered.writeUInt32LE(64 * 1024 * 1024 + 1, directory + 24);
  assert.throws(() => inspectZip(sizeTampered), /크기 제한/);
  const crcTampered = Buffer.from(zipSync({ 'part.xml': new TextEncoder().encode('<x/>') }));
  const crcDirectory = crcTampered.indexOf(Buffer.from([0x50, 0x4b, 0x01, 0x02]));
  crcTampered.writeUInt32LE((crcTampered.readUInt32LE(crcDirectory + 16) ^ 0xffffffff) >>> 0, crcDirectory + 16);
  assert.throws(() => inspectZip(crcTampered), /CRC/);
});
