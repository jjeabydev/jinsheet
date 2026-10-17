'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { createColumnOffsets, columnAtPixel, visibleColumnRange, createRowView, createRowMetrics, clampRowHeightPixels, excelColumnWidthPixels, estimateBestFitWidth } = require('../src/grid');

test('희소 행 높이는 백만 행과 숨긴 행의 표시 좌표·역방향 픽셀 조회에 반영된다', () => {
  assert.equal(clampRowHeightPixels(0), 1); assert.equal(clampRowHeightPixels(1000), 546);
  assert.ok(Math.abs(clampRowHeightPixels(26.4) - 26.4) < 1e-10); assert.equal(clampRowHeightPixels(NaN), 26);
  assert.ok(Math.abs(clampRowHeightPixels((45.33 + 1) * 4 / 3) * .75 - 46.33) < 1e-10);
  const rows = createRowView(1048576, undefined, [1, 3]);
  const metrics = createRowMetrics(rows, { 0: 52, 1: 200, 2: 13, 999999: 80 });
  assert.equal(metrics.top(0), 0); assert.equal(metrics.height(0), 52);
  assert.equal(metrics.top(1), 52); assert.equal(metrics.top(2), 65);
  assert.equal(metrics.indexAtPixel(51.99), 0); assert.equal(metrics.indexAtPixel(52), 1);
  assert.equal(metrics.indexAtPixel(65), 2);
  const distant = rows.indexOf(999999), pixel = metrics.top(distant);
  assert.equal(pixel, distant * 26 + 13); assert.equal(metrics.height(distant), 80);
  assert.equal(metrics.indexAtPixel(pixel + 79.9), distant);
  assert.equal(metrics.totalHeight, rows.displayCount * 26 + 67);
  assert.equal(createRowMetrics(createRowView(3), {}, 20).top(3), 60);
});

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

test('숨긴 열은 공간을 차지하지 않고 픽셀 조회에서 건너뛴다', () => {
  const offsets = createColumnOffsets(5, new Map(), [0, 2, 3]);
  assert.deepEqual([...offsets], [0, 0, 144, 144, 144, 288]);
  assert.equal(columnAtPixel(offsets, 0), 1);
  assert.equal(columnAtPixel(offsets, 143.9), 1);
  assert.equal(columnAtPixel(offsets, 144), 4);
  assert.deepEqual(visibleColumnRange(offsets, 0, 100), { columnStart: 0, columnCount: 5 });
  assert.equal(columnAtPixel(createColumnOffsets(2, new Map(), [0, 1]), 0), -1);
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

test('OOXML 열 너비는 추가 패딩이나 표시 제한 없이 픽셀로 변환한다', () => {
  assert.equal(excelColumnWidthPixels(9.140625), 64);
  assert.equal(excelColumnWidthPixels(8.43), 59);
  assert.equal(excelColumnWidthPixels(24), 168);
  assert.equal(excelColumnWidthPixels(0), 0);
  assert.equal(excelColumnWidthPixels(.140625), 1);
  assert.equal(excelColumnWidthPixels(255), 1785);
  assert.equal(excelColumnWidthPixels(undefined, 64), 64);
  assert.equal(excelColumnWidthPixels(256), 144);
  assert.deepEqual([...createColumnOffsets(4, new Map([[1, 1], [2, 1785]]), [], 64, true)], [0, 64, 65, 1850, 1914]);
  for (let pixel = 1; pixel <= 1785; pixel++) assert.equal(excelColumnWidthPixels(Math.round(pixel / 7 * 100) / 100), pixel);
});

test('열 자동 맞춤은 라틴 문자·한글·탭 너비와 설정 한계를 반영한다', () => {
  assert.equal(estimateBestFitWidth([]), 48);
  assert.equal(estimateBestFitWidth(['ABCD']), 48);
  assert.equal(estimateBestFitWidth(['안녕하세요']), 90);
  assert.equal(estimateBestFitWidth(['a\tb']), 62);
  assert.equal(estimateBestFitWidth(['x'.repeat(2000)]), 640);
});
test('표시 행 매핑은 실제 행 주소와 압축된 화면 순번을 왕복하고 잘못된 행을 거부한다', () => {
  const identity = createRowView(5);
  assert.equal(identity.displayCount, 5);
  assert.equal(identity.rowAt(3), 3);
  assert.equal(identity.indexOf(3), 3);
  const filtered = createRowView(8, [0, 2, 5, 7]);
  assert.equal(filtered.displayCount, 4);
  assert.equal(filtered.rowAt(1), 2);
  assert.equal(filtered.indexOf(5), 2);
  assert.equal(filtered.indexOf(3), -1);
  assert.equal(filtered.rowAt(4), -1);
  assert.throws(() => createRowView(4, [0, 2, 2]), /중복 없이 오름차순/);
  assert.throws(() => createRowView(4, [0, 4]), /중복 없이 오름차순/);
});
test('희소 숨김 행 매핑은 백만 행에서도 주소와 화면 순번을 압축 없이 계산한다', () => {
  const filtered = createRowView(1048576, undefined, [2, 100, 1048575]);
  assert.equal(filtered.displayCount, 1048573);
  assert.equal(filtered.rowAt(2), 3);
  assert.equal(filtered.rowAt(98), 99);
  assert.equal(filtered.indexOf(100), -1);
  assert.equal(filtered.indexOf(101), 99);
  assert.equal(filtered.rowAt(filtered.displayCount - 1), 1048574);
  assert.throws(() => createRowView(4, undefined, [1, 1]), /숨김 행 인덱스/);
});
