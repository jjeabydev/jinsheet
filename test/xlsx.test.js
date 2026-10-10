'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { unzipSync, zipSync } = require('../vendor/node_modules/fflate');
const { analyze, formatCellValue, inspectZip, patchCell, patchCells, patchColumnWidths, transformStructure } = require('../src/xlsx');

function fixture({ formula = false, extra = undefined, workbookXml = undefined, formatted = false, columnDefinitions = '', sheetViews = '', cellExtension = '' } = {}) {
  const files = {
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': workbookXml || `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Hidden" sheetId="2" state="hidden" r:id="rId2"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>${formatted ? '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' : ''}</Relationships>`,
    'xl/worksheets/sheet1.xml': `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="${formatted ? 'A1:D1' : 'A1:C1'}"/>${sheetViews}${columnDefinitions}<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>before</t></is>${cellExtension}</c>${formatted ? '<c r="B1" s="1"><v>45292</v></c><c r="C1" s="2"><v>0.125</v></c><c r="D1" s="3"><v>45292</v></c>' : '<c r="B1"><v>42</v></c><c r="C1" t="b"><v>1</v></c>'}</row>${formula ? '<row r="2"><c r="A2"><f>1+1</f><v>2</v></c></row>' : ''}</sheetData></worksheet>`,
    'xl/worksheets/sheet2.xml': `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>9</v></c></row></sheetData></worksheet>`
  };
  if (formatted) {
    files['xl/styles.xml'] = '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts><fonts count="2"><font/><font><b/><i/><color rgb="FF123456"/><sz val="14"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFEEAA"/></patternFill></fill></fills><borders count="2"><border/><border><left style="thin"><color rgb="FFABCDEF"/></left></border></borders><cellXfs count="4"><xf numFmtId="0"/><xf numFmtId="14"/><xf numFmtId="10"/><xf numFmtId="164" fontId="1" fillId="2" borderId="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf></cellXfs></styleSheet>';
    files['[Content_Types].xml'] = files['[Content_Types].xml'].replace('</Types>', '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
  }
  if (extra) files[extra.name] = extra.content;
  return zipSync(Object.fromEntries(Object.entries(files).map(([name, content]) => [name, new TextEncoder().encode(content)])), { level: 6 });
}

test('XLSX freeze panes are displayed and preserved when saving an edited cell', () => {
  const paneXml = '<sheetViews><sheetView workbookViewId="0"><pane xSplit="2" ySplit="3" topLeftCell="C4" activePane="bottomRight" state="frozen"/><selection pane="bottomRight" activeCell="C4" sqref="C4"/></sheetView></sheetViews>';
  const model = analyze(fixture({ sheetViews: paneXml }));
  assert.equal(model.writable, true, model.readOnlyReason);
  assert.deepEqual(model.sheets[0].freezePane, { rows: 3, columns: 2 });
  const saved = patchCell(model, 0, 'A1', 'changed');
  const reopened = analyze(saved.bytes);
  assert.equal(reopened.writable, true, reopened.readOnlyReason);
  assert.deepEqual(reopened.sheets[0].freezePane, { rows: 3, columns: 2 });
  assert.ok(reopened.sheets[0].xml.includes(paneXml));
  assert.deepEqual(Buffer.from(reopened.files['xl/worksheets/sheet2.xml']), Buffer.from(model.files['xl/worksheets/sheet2.xml']));
});

test('XLSX cell extensions and rich inline strings are read-only before an edit can discard them', () => {
  const extensionModel = analyze(fixture({ cellExtension: '<extLst><ext uri="urn:test"/></extLst>' }));
  assert.equal(extensionModel.writable, false);
  assert.match(extensionModel.readOnlyReason, /지원하지 않는 셀 구성요소/);
  assert.throws(() => patchCell(extensionModel, 0, 'A1', 'changed'), /지원하지 않는 셀 구성요소/);

  const files = unzipSync(fixture());
  const worksheet = new TextDecoder().decode(files['xl/worksheets/sheet1.xml']);
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(worksheet.replace('<is><t>before</t></is>', '<is><r><rPr><b/></rPr><t>before</t></r></is>'));
  const richTextModel = analyze(zipSync(files));
  assert.equal(richTextModel.writable, false);
  assert.match(richTextModel.readOnlyReason, /서식 또는 확장이 포함된 인라인 문자열/);
});

test('XLSX split panes, malformed freeze panes and oversized frozen regions are read-only', () => {
  const cases = [
    ['<sheetViews><sheetView><pane xSplit="1" ySplit="0" topLeftCell="B1" state="split"/></sheetView></sheetViews>', /분할 창은 고정 창/],
    ['<sheetViews><sheetView><pane xSplit="1.5" ySplit="0" topLeftCell="B1" state="frozen"/></sheetView></sheetViews>', /소수 또는 잘못된 고정 창 범위/],
    ['<sheetViews><sheetView><pane xSplit="0" ySplit="51" topLeftCell="A52" state="frozen"/></sheetView></sheetViews>', /화면 렌더링 한도/],
    ['<sheetViews><sheetView><pane xSplit="1" ySplit="0" topLeftCell="?1" state="frozen"/></sheetView></sheetViews>', /시작 셀이 시트 범위 또는 고정 영역/],
    ['<sheetViews><sheetView><pane xSplit="1" ySplit="0" topLeftCell="A1" state="frozen"/></sheetView></sheetViews>', /시작 셀이 시트 범위 또는 고정 영역/],
    ['<sheetViews><sheetView><pane xSplit="1" ySplit="0" topLeftCell="XFE1" state="frozen"/></sheetView></sheetViews>', /시작 셀이 시트 범위 또는 고정 영역/] ,
    ['<sheetViews><sheetView><pane xSplit="0" ySplit="1" topLeftCell="A1048577" state="frozen"/></sheetView></sheetViews>', /시작 셀이 시트 범위 또는 고정 영역/],
    ['<sheetViews><sheetView><pane xSplit="1" ySplit="0" topLeftCell="B1" state="frozen"/><pane xSplit="2" ySplit="0" topLeftCell="C1" state="frozen"/></sheetView></sheetViews>', /여러 분할 창/]
  ];
  for (const [sheetViews, reason] of cases) {
    const model = analyze(fixture({ sheetViews }));
    assert.equal(model.writable, false);
    assert.match(model.readOnlyReason, reason);
  }
});

test('XLSX column width spans are parsed as bounded zero-based metadata', () => {
  const model = analyze(fixture({ columnDefinitions: '<cols><col min="1" max="2" width="24" customWidth="1"/><col min="4" max="4" width="9.5"/></cols>' }));
  assert.deepEqual(model.sheets[0].columnWidths, { 0: 24, 1: 24, 3: 9.5 });
  assert.equal(model.writable, true);
  const saved = patchCell(model, 0, 'A1', 'after');
  const reopened = analyze(saved.bytes);
  assert.deepEqual(reopened.sheets[0].columnWidths, { 0: 24, 1: 24, 3: 9.5 });
  assert.match(reopened.sheets[0].xml, /<cols><col min="1" max="2" width="24"/);
  assert.throws(() => analyze(fixture({ columnDefinitions: '<cols><col min="0" max="1" width="12"/></cols>' })), /열 너비 범위/);
  assert.throws(() => analyze(fixture({ columnDefinitions: '<cols><col min="1" max="1" width="256"/></cols>' })), /열 너비 값/);
  assert.throws(() => analyze(fixture({ columnDefinitions: '<cols><col min="1" max="1" width=""/></cols>' })), /열 너비 값/);
});

test('XLSX 열 너비 저장은 범위 정의를 나누고 숨김·스타일 속성과 다른 패키지 항목을 보존한다', () => {
  const original = fixture({ columnDefinitions: '<cols><col min="1" max="3" width="24" customWidth="1" hidden="1" style="0"/><col min="5" max="5" width="9.5"/></cols>' });
  const model = analyze(original);
  const result = patchColumnWidths(model, [{ sheetIndex: 0, column: 1, width: 12.25 }, { sheetIndex: 0, column: 3, width: 18 }]);
  const reopened = analyze(result.bytes);
  assert.deepEqual(reopened.sheets[0].columnWidths, { 0: 24, 1: 12.25, 2: 24, 3: 18, 4: 9.5 });
  assert.match(reopened.sheets[0].xml, /<col min="1" max="1" width="24" customWidth="1" hidden="1" style="0"\/>/);
  assert.match(reopened.sheets[0].xml, /<col min="2" max="2" width="12.25" customWidth="1" hidden="1" style="0"\/>/);
  assert.match(reopened.sheets[0].xml, /<col min="3" max="3" width="24" customWidth="1" hidden="1" style="0"\/>/);
  assert.match(reopened.sheets[0].xml, /<col min="4" max="4" width="18" customWidth="1"\/>/);
  assert.deepEqual(Buffer.from(reopened.files['xl/worksheets/sheet2.xml']), Buffer.from(model.files['xl/worksheets/sheet2.xml']));
  assert.deepEqual(Buffer.from(reopened.files['xl/styles.xml'] || []), Buffer.from(model.files['xl/styles.xml'] || []));
});

test('XLSX 열 너비 저장은 없는 정의를 만들고 범위·형식 오류에서 저장을 거부한다', () => {
  const model = analyze(fixture());
  const result = patchColumnWidths(model, [{ sheetIndex: 0, column: 2, width: 20 }]);
  const reopened = analyze(result.bytes);
  assert.deepEqual(reopened.sheets[0].columnWidths, { 2: 20 });
  assert.match(reopened.sheets[0].xml, /<cols><col min="3" max="3" width="20" customWidth="1"\/><\/cols><sheetData>/);
  assert.throws(() => patchColumnWidths(model, [{ sheetIndex: 0, column: 16384, width: 12 }]), /요청이 유효하지 않습니다/);
  assert.throws(() => patchColumnWidths(analyze(fixture({ formula: true })), [{ sheetIndex: 0, column: 0, width: 12 }]), /수식이 포함/);
});

test('XLSX standard date/percent and common custom formats affect display only', () => {
  const model = analyze(fixture({ formatted: true }));
  const date = model.sheets[0].cells.get('B1'), percent = model.sheets[0].cells.get('C1');
  assert.equal(model.writable, true);
  assert.equal(date.value, '45292');
  assert.equal(date.numFmtId, 14);
  assert.equal(date.styleProperties, null);
  assert.equal(formatCellValue(date, 'en'), '1/1/24');
  assert.equal(formatCellValue(date, 'ko'), '2024-01-01');
  assert.equal(formatCellValue({ value: '60', kind: 'number', numFmtId: 14 }, 'en'), '2/29/00');
  assert.equal(percent.value, '0.125');
  assert.equal(formatCellValue(percent, 'en'), '12.50%');
  assert.equal(formatCellValue({ value: '1234.5', kind: 'number', numFmtId: 164, formatCode: '#,##0.00' }, 'en'), '1,234.50');
  assert.equal(formatCellValue({ value: '1234.5', kind: 'number', numFmtId: 5 }, 'en'), '$1,235');
  assert.equal(formatCellValue({ value: '1234.5', kind: 'number', numFmtId: 6 }, 'en'), '$1,235');
  assert.equal(formatCellValue({ value: '1234.5', kind: 'number', numFmtId: 7 }, 'en'), '$1,234.50');
  assert.equal(formatCellValue({ value: '1234.5', kind: 'number', numFmtId: 8 }, 'en'), '$1,234.50');
  assert.equal(formatCellValue({ value: '-1234.5', kind: 'number', numFmtId: 39 }, 'en'), '(1,234.50)');
  assert.equal(formatCellValue({ value: '1234.5', kind: 'number', numFmtId: 166, formatCode: '0 \"days\"' }, 'en'), '1234.5');
  const customDate = model.sheets[0].cells.get('D1');
  assert.equal(customDate.numFmtId, 164);
  assert.equal(customDate.formatCode, 'yyyy-mm-dd');
  assert.deepEqual(customDate.styleProperties, { bold: true, italic: true, color: '#123456', fontSize: 14, fill: '#FFEEAA', horizontal: 'center', vertical: 'center', borders: { left: { color: '#ABCDEF', line: 'thin' } } });
  assert.equal(formatCellValue(customDate, 'en'), '2024-01-01');
  const editedModel = patchCell(model, 0, 'C1', '0.25').model;
  const edited = editedModel.sheets[0].cells.get('C1');
  assert.equal(edited.value, '0.25');
  assert.equal(formatCellValue(edited, 'en'), '25.00%');
  assert.deepEqual(Buffer.from(editedModel.files['xl/styles.xml']), Buffer.from(model.files['xl/styles.xml']));
});

test('100,000 numeric-cell workbook survives distant multi-cell edit, save and reopen', () => {
  const files = unzipSync(fixture());
  const rows = [];
  for (let row = 1; row <= 10000; row++) {
    let cells = '';
    for (let column = 1; column <= 10; column++) {
      const address = String.fromCharCode(64 + column) + row;
      cells += `<c r="${address}"><v>${row * column}</v></c>`;
    }
    rows.push(`<row r="${row}">${cells}</row>`);
  }
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:J10000"/><sheetData>${rows.join('')}</sheetData></worksheet>`);
  files['xl/worksheets/sheet2.xml'] = new TextEncoder().encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>');
  const model = analyze(zipSync(files, { level: 0 }));
  assert.equal(model.sheets[0].cells.size, 100000);
  assert.equal(model.writable, true);
  const untouchedEntries = Object.fromEntries(Object.entries(model.files).filter(([name]) => name !== model.sheets[0].path));
  const output = patchCells(model, [
    { sheetIndex: 0, address: 'A1', value: 'first-edited' },
    { sheetIndex: 0, address: 'J10000', value: 'last-edited' }
  ]).bytes;
  const reopened = analyze(output);
  assert.equal(reopened.writable, true, reopened.readOnlyReason);
  assert.equal(reopened.sheets[0].cells.size, 100000);
  assert.equal(reopened.sheets[0].cells.get('A1').value, 'first-edited');
  assert.equal(reopened.sheets[0].cells.get('J10000').value, 'last-edited');
  assert.equal(reopened.sheets[0].cells.get('E5678').value, String(5678 * 5));
  for (const [name, bytes] of Object.entries(untouchedEntries)) assert.deepEqual(Buffer.from(reopened.files[name]), Buffer.from(bytes), `untouched package entry changed: ${name}`);
});

test("빈 셀에 새 값을 입력하면 셀·행 순서와 사용 범위가 갱신되고 값을 비울 수 있다", () => {
  const model = analyze(fixture());
  const addedCell = patchCell(model, 0, "D1", "12").model;
  assert.equal(addedCell.sheets[0].cells.get("D1").kind, "number");
  assert.equal(addedCell.sheets[0].cells.get("D1").value, "12");
  assert.match(addedCell.sheets[0].xml, /<dimension ref="A1:D1"\/>/);
  const addedRow = patchCell(model, 0, "B2", "new row").model;
  assert.equal(addedRow.sheets[0].cells.get("B2").value, "new row");
  assert.match(addedRow.sheets[0].xml, /<dimension ref="A1:C2"\/>/);
  const multiple = patchCells(model, [
    { sheetIndex: 0, address: "E1", value: "last" },
    { sheetIndex: 0, address: "D1", value: "12" },
    { sheetIndex: 0, address: "B2", value: "second row" },
    { sheetIndex: 0, address: "A2", value: "first cell" }
  ]).model;
  assert.deepEqual([...multiple.sheets[0].cells.values()].filter(cell => cell.row === 0).sort((a, b) => a.column - b.column).map(cell => cell.address), ["A1", "B1", "C1", "D1", "E1"]);
  assert.equal(multiple.sheets[0].cells.get("A2").value, "first cell");
  assert.equal(multiple.sheets[0].cells.get("B2").value, "second row");
  const cleared = patchCell(model, 0, "A1", "").model;
  assert.equal(cleared.sheets[0].cells.get("A1").value, "");
});

test("단순 XLSX에서 행·열 삽입·삭제가 셀 주소와 사용 범위를 함께 이동한다", () => {
  const original = fixture(), model = analyze(original);
  const insertedRow = transformStructure(model, { sheetIndex: 0, axis: "row", action: "insert", index: 0 });
  assert.equal(insertedRow.model.sheets[0].cells.get("A2").value, "before");
  assert.equal(insertedRow.model.sheets[0].cells.get("B2").value, "42");
  assert.match(insertedRow.model.sheets[0].xml, /<dimension ref="A1:C2"\/>/);
  const removedRow = transformStructure(insertedRow.model, { sheetIndex: 0, axis: "row", action: "delete", index: 0 });
  assert.equal(removedRow.model.sheets[0].cells.get("A1").value, "before");
  assert.equal(removedRow.model.sheets[0].cells.has("A2"), false);
  const insertedColumn = transformStructure(model, { sheetIndex: 0, axis: "column", action: "insert", index: 1 });
  assert.equal(insertedColumn.model.sheets[0].cells.get("A1").value, "before");
  assert.equal(insertedColumn.model.sheets[0].cells.get("C1").value, "42");
  assert.equal(insertedColumn.model.sheets[0].cells.get("D1").value, "TRUE");
  assert.match(insertedColumn.model.sheets[0].xml, /<dimension ref="A1:D1"\/>/);
  const removedColumn = transformStructure(insertedColumn.model, { sheetIndex: 0, axis: "column", action: "delete", index: 1 });
  assert.equal(removedColumn.model.sheets[0].cells.get("B1").value, "42");
  assert.equal(removedColumn.model.sheets[0].cells.get("C1").value, "TRUE");
  for (const name of Object.keys(model.files)) {
    if (name !== model.sheets[0].path) assert.deepEqual(Buffer.from(insertedRow.model.files[name]), Buffer.from(model.files[name]));
  }
  assert.throws(() => transformStructure(analyze(fixture({ formula: true })), { sheetIndex: 0, axis: "row", action: "insert", index: 0 }), /수식이 포함/);
});

test("셀 값 편집은 숫자와 불리언 데이터 형식을 보존한다", () => {
  const model = analyze(fixture());
  const numeric = patchCell(model, 0, "B1", "58").model.sheets[0].cells.get("B1");
  assert.equal(numeric.kind, "number");
  assert.equal(numeric.value, "58");
  const boolean = patchCell(model, 0, "C1", "FALSE").model.sheets[0].cells.get("C1");
  assert.equal(boolean.kind, "boolean");
  assert.equal(boolean.value, "FALSE");
  const text = patchCell(model, 0, "B1", "문자").model.sheets[0].cells.get("B1");
  assert.equal(text.kind, "string");
  assert.equal(text.value, "문자");
});

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


test("실제 JinSheet XLSX 샘플은 편집·행 삽입 후 재파싱되고 비수정 패키지 항목을 보존한다", () => {
  const fs = require("node:fs");
  const path = require("node:path");
  const originalBytes = fs.readFileSync(path.join(__dirname, "fixtures", "JinSheet-Sample.xlsx"));
  const original = analyze(originalBytes);
  assert.equal(original.writable, true, original.readOnlyReason);
  assert.equal(original.sheets[0].cells.get("B2").value, "12");

  const edited = patchCell(original, 0, "B2", "13").model;
  const changed = transformStructure(edited, { sheetIndex: 0, axis: "row", action: "insert", index: 0 });
  const reopened = analyze(changed.bytes);
  assert.equal(reopened.writable, true, reopened.readOnlyReason);
  assert.equal(reopened.sheets[0].cells.get("A2").value, "상품");
  assert.equal(reopened.sheets[0].cells.get("B3").value, "13");
  assert.equal(reopened.sheets[0].cells.get("C3").value, "1500");

  for (const [name, bytes] of Object.entries(original.files)) {
    if (name === original.sheets[0].path) continue;
    assert.deepEqual(Buffer.from(reopened.files[name]), Buffer.from(bytes), `untouched package entry changed: ${name}`);
  }
});


test('실제 서식 샘플의 셀 색상·글꼴·정렬·테두리를 파싱하고 안전 편집을 허용한다', () => {
  const fs = require('node:fs');
  const path = require('node:path');
  const bytes = fs.readFileSync(path.join(__dirname, 'fixtures', 'JinSheet-Styled-Sample.xlsx'));
  const model = analyze(bytes);
  assert.equal(model.writable, true, model.readOnlyReason);
  assert.equal(model.sheets[0].cells.get('A1').styleProperties.fill, '#1F4E78');
  assert.equal(model.sheets[0].cells.get('A1').styleProperties.color, '#FFFFFF');
  assert.equal(model.sheets[0].cells.get('A1').styleProperties.bold, true);
  assert.equal(model.sheets[0].cells.get('A1').styleProperties.horizontal, 'center');
  assert.equal(model.sheets[0].cells.get('A1').styleProperties.borders.left.color, '#B4C7E7');
  assert.equal(model.sheets[0].cells.get('C2').styleProperties.horizontal, 'right');
  assert.equal(formatCellValue(model.sheets[0].cells.get('C2'), 'en'), '1,500');
});
