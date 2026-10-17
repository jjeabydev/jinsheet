'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { applyEdits, decode, detectDelimiter, encode, extendCells, load, parse, transformStructure } = require('../src/csv');

test('인용 필드 안의 줄바꿈과 escape quote를 파싱한다', () => {
  const text = 'a,b\r\n"line 1\nline 2","say ""hi"""\r\n';
  const rows = parse(text, ',');
  assert.equal(rows.length, 2);
  assert.deepEqual(rows[1].map(cell => cell.value), ['line 1\nline 2', 'say "hi"']);
});

test('수정하지 않은 바이트 구간, BOM, 혼합 개행을 유지한다', () => {
  const original = Buffer.from([0xef, 0xbb, 0xbf, ...Buffer.from('a,b\r\nkeep,"x,y"\n')]);
  const doc = load(original, '.csv');
  const changed = encode(applyEdits(doc.text, doc.rows, [{ row: 1, column: 1, value: 'new,value' }], doc.delimiter), doc.encoding, doc.bom);
  const result = changed.subarray(doc.bom.length).toString('utf8');
  assert.equal(result, 'a,b\r\nkeep,"new,value"\n');
});

test('UTF-16LE BOM을 보존하고 왕복한다', () => {
  const bytes = Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from('a\tb', 'utf16le')]);
  const doc = load(bytes, '.tsv');
  assert.equal(doc.encoding, 'utf-16le');
  assert.deepEqual(encode(doc.text, doc.encoding, doc.bom), bytes);
});

test('모호한 구분자와 인용 오류를 거부한다', () => {
  assert.throws(() => detectDelimiter('one\ntwo'), /구분자를/);
  assert.throws(() => parse('a,"unfinished', ','), /닫히지 않은/);
  assert.throws(() => parse('a,b"c', ','), /따옴표가 잘못된/);
  assert.throws(() => parse('"a"tail,b', ','), /인용 필드 뒤/);
});

test('필드 source offset으로 수정 영역만 교체한다', () => {
  const text = 'left," keep ",right\n';
  const rows = parse(text, ',');
  assert.equal(applyEdits(text, rows, [{ row: 0, column: 1, value: 'new' }], ','), 'left,"new",right\n');
  assert.throws(() => applyEdits(text, rows, [{ row: 0, column: 1, value: 'first' }, { row: 0, column: 1, value: 'second' }], ','), /겹치는 셀 수정 요청/);
});

test('CSV 대량 셀 수정과 열 구조 편집은 모든 패치를 보존한다', () => {
  const rowCount = 4000;
  const text = Array.from({ length: rowCount }, (_, row) => `row${row},value${row},extra${row}`).join('\n');
  const rows = parse(text, ',');
  const edits = rows.map((_, row) => ({ row, column: 1, value: `updated${row}` }));
  const changed = parse(applyEdits(text, rows, edits, ','), ',');
  assert.equal(changed.length, rowCount);
  assert.deepEqual(changed[0].map(cell => cell.value), ['row0', 'updated0', 'extra0']);
  assert.deepEqual(changed.at(-1).map(cell => cell.value), [`row${rowCount - 1}`, `updated${rowCount - 1}`, `extra${rowCount - 1}`]);

  const state = load(Buffer.from(text), '.csv');
  const columnDeleted = transformStructure(state, '.csv', { axis: 'column', action: 'delete', index: 1 });
  assert.equal(columnDeleted.rows.length, rowCount);
  assert.deepEqual(columnDeleted.rows[0].map(cell => cell.value), ['row0', 'extra0']);
  assert.deepEqual(columnDeleted.rows.at(-1).map(cell => cell.value), [`row${rowCount - 1}`, `extra${rowCount - 1}`]);
});

test('CSV는 100,000셀 한도에서 마지막 셀과 필드 오프셋을 보존한다', () => {
  const rowCount = 10000, columnCount = 10;
  const text = Array.from({ length: rowCount }, (_, row) => Array.from({ length: columnCount }, (_, column) => `${row}:${column}`).join(',')).join('\n');
  const rows = parse(text, ',');
  assert.equal(rows.length, rowCount);
  assert.equal(rows.reduce((count, row) => count + row.length, 0), rowCount * columnCount);
  assert.equal(rows[0][0].value, '0:0');
  assert.equal(rows.at(-1).at(-1).value, '9999:9');
  assert.equal(text.slice(rows.at(-1).at(-1).start, rows.at(-1).at(-1).end), '9999:9');
});

test('CSV/TSV 행·열 구조 변경은 원래 필드와 혼합 개행을 보존한다', () => {
  const original = Buffer.from([0xef, 0xbb, 0xbf, ...Buffer.from('name,value\r\n"keep,as is",1\nbeta,2')]);
  const state = load(original, '.csv');
  const rowInserted = transformStructure(state, '.csv', { axis: 'row', action: 'insert', index: 1 });
  assert.equal(rowInserted.text, 'name,value\r\n,\n"keep,as is",1\nbeta,2');
  const rowDeleted = transformStructure(state, '.csv', { axis: 'row', action: 'delete', index: 1 });
  assert.equal(rowDeleted.text, 'name,value\r\nbeta,2');
  const columnInserted = transformStructure(state, '.csv', { axis: 'column', action: 'insert', index: 1 });
  assert.equal(columnInserted.text, 'name,,value\r\n"keep,as is",,1\nbeta,,2');
  const columnDeleted = transformStructure(state, '.csv', { axis: 'column', action: 'delete', index: 1 });
  assert.equal(columnDeleted.text, 'name\r\n"keep,as is"\nbeta');
  assert.deepEqual(encode(rowDeleted.text, rowDeleted.encoding, rowDeleted.bom).subarray(0, 3), Buffer.from([0xef, 0xbb, 0xbf]));
});

test('빈 TSV에서 첫 행 또는 열을 삽입하면 빈 셀 하나가 생성된다', () => {
  const state = load(Buffer.alloc(0), '.tsv');
  for (const axis of ['row', 'column']) {
    const inserted = transformStructure(state, '.tsv', { axis, action: 'insert', index: 0 });
    assert.equal(inserted.text, '""');
    assert.deepEqual(inserted.rows.map(row => row.map(cell => cell.value)), [['']]);
  }
});

test('CSV 빈 영역의 멀리 떨어진 셀을 한 번에 확장하고 기존 개행 형식을 보존한다', () => {
  const state = load(Buffer.from('name,value\r\nalpha,1\r\n'), '.csv');
  const expanded = extendCells(state, '.csv', [{ row: 1, column: 4 }, { row: 3, column: 2 }]);
  assert.deepEqual(expanded.rows[1].map(cell => cell.value), ['alpha', '1', '', '', '']);
  assert.deepEqual(expanded.rows[3].map(cell => cell.value), ['', '', '']);
  assert.equal(expanded.text, 'name,value\r\nalpha,1,,,\r\n,\r\n,,');
});
