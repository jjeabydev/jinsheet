'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { createColumnOffsets, columnAtPixel, visibleColumnRange, excelColumnWidthPixels, estimateBestFitWidth } = require('../src/grid');

test('열 너비 재배치에 따라 prefix offset과 픽셀 좌표 조회가 일치한다', () => {
  const offsets = createColumnOffsets(5, new Map([[1, 240], [3, 72]]));
  assert.deepEqual([...offsets], [0, 144, 384, 528, 600, 744]);
  assert.equal(columnAtPixel(offsets, 0), 0);
  assert.equal(columnAtPixel(offsets, 143.9), 0);
  assert.equal(columnAtPixel(offsets, 144), 1);
  assert.equal(columnAtPixel(offsets, 383.9), 1);
  assert.equal(columnAtPixel(offsets, 384), 2);
  assert.equal(columnAtPixel(offsets, 9999), 4);
});

test('수평 스크롤과 열 너비에 맞춰 가상화 범위를 계산한다', () => {
  const offsets = createColumnOffsets(100, new Map([[0, 240], [1, 72], [2, 320]]));
  assert.deepEqual(visibleColumnRange(offsets, 0, 500), { columnStart: 0, columnCount: 8 });
  assert.deepEqual(visibleColumnRange(offsets, 700, 500), { columnStart: 0, columnCount: 12 });
  assert.deepEqual(visibleColumnRange(createColumnOffsets(3), 0, 500), { columnStart: 0, columnCount: 3 });
  assert.deepEqual(visibleColumnRange(createColumnOffsets(100), 5000, 400), { columnStart: 30, columnCount: 13 });
});

test('잘못된 열 개수와 너비는 기본값으로 안전하게 처리한다', () => {
  assert.deepEqual([...createColumnOffsets(0)], [0]);
  assert.deepEqual([...createColumnOffsets(-2)], [0]);
  const offsets = createColumnOffsets(3, new Map([[0, 12], [1, 900], [2, Number.NaN]]));
  assert.deepEqual([...offsets], [0, 48, 688, 832]);
});

test('Excel character widths convert to bounded pixels for the grid', () => {
  assert.equal(excelColumnWidthPixels(8.43), 64);
  assert.equal(excelColumnWidthPixels(24), 173);
  assert.equal(excelColumnWidthPixels(0), 48);
  assert.equal(excelColumnWidthPixels(256), 144);
});

test('열 자동 맞춤은 라틴 문자·한글·탭 너비와 설정 한계를 반영한다', () => {
  assert.equal(estimateBestFitWidth([]), 48);
  assert.equal(estimateBestFitWidth(['ABCD']), 48);
  assert.equal(estimateBestFitWidth(['안녕하세요']), 90);
  assert.equal(estimateBestFitWidth(['a\tb']), 62);
  assert.equal(estimateBestFitWidth(['x'.repeat(2000)]), 640);
});
