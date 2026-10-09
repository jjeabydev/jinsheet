'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { applyEdits, decode, detectDelimiter, encode, load, parse } = require('../src/csv');

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
});
