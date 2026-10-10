'use strict';

const DEFAULT_COLUMN_WIDTH = 144;
const MIN_COLUMN_WIDTH = 48;
const MAX_COLUMN_WIDTH = 640;
const ROW_HEADER_WIDTH = 52;
const COLUMN_CHUNK = 5;
const MAX_RANGE_COLUMNS = 30;

function clampColumnWidth(width) {
  const value = Number(width);
  return Number.isFinite(value) ? Math.max(MIN_COLUMN_WIDTH, Math.min(MAX_COLUMN_WIDTH, value)) : DEFAULT_COLUMN_WIDTH;
}

function excelColumnWidthPixels(width) {
  const value = Number(width);
  return Number.isFinite(value) && value >= 0 && value <= 255 ? clampColumnWidth(Math.round(value * 7 + 5)) : DEFAULT_COLUMN_WIDTH;
}

function createColumnOffsets(count, widths = new Map()) {
  const columnCount = Number.isInteger(count) && count > 0 ? count : 0;
  const offsets = new Float64Array(columnCount + 1);
  for (let column = 0; column < columnCount; column++) offsets[column + 1] = offsets[column] + clampColumnWidth(widths.get(column));
  return offsets;
}

function columnAtPixel(offsets, pixel) {
  const count = offsets.length - 1;
  if (count < 1) return -1;
  const position = Number.isFinite(pixel) ? Math.max(0, Math.min(pixel, offsets[count] - Number.EPSILON * Math.max(1, offsets[count]))) : 0;
  let low = 0, high = count - 1;
  while (low < high) {
    const middle = Math.floor((low + high) / 2);
    if (offsets[middle + 1] <= position) low = middle + 1;
    else high = middle;
  }
  return low;
}

function visibleColumnRange(offsets, scrollLeft, viewportWidth, options = {}) {
  const count = offsets.length - 1;
  if (count < 1) return { columnStart: 0, columnCount: 0 };
  const headerWidth = Number.isFinite(options.rowHeaderWidth) ? options.rowHeaderWidth : ROW_HEADER_WIDTH;
  const chunkSize = Number.isInteger(options.chunkSize) && options.chunkSize > 0 ? options.chunkSize : COLUMN_CHUNK;
  const overscan = Number.isInteger(options.overscan) && options.overscan >= 0 ? options.overscan : COLUMN_CHUNK;
  const maxColumns = Number.isInteger(options.maxColumns) && options.maxColumns > 0 ? options.maxColumns : MAX_RANGE_COLUMNS;
  const scroll = Number.isFinite(scrollLeft) ? Math.max(0, scrollLeft) : 0;
  const width = Number.isFinite(viewportWidth) ? Math.max(1, viewportWidth) : 1;
  const visibleStart = columnAtPixel(offsets, Math.max(0, scroll - headerWidth));
  const visibleEnd = columnAtPixel(offsets, Math.max(0, scroll - headerWidth) + width) + 1;
  const columnStart = Math.floor(visibleStart / chunkSize) * chunkSize;
  const columnCount = Math.min(maxColumns, Math.max(1, visibleEnd - columnStart + overscan), count - columnStart);
  return { columnStart, columnCount };
}


function estimateBestFitWidth(values, options = {}) {
  const minimum = Number.isFinite(options.minimum) ? options.minimum : MIN_COLUMN_WIDTH;
  const maximum = Number.isFinite(options.maximum) ? options.maximum : MAX_COLUMN_WIDTH;
  const padding = Number.isFinite(options.padding) ? Math.max(0, options.padding) : 20;
  const textWidth = value => {
    let width = 0;
    for (const character of String(value ?? "")) {
      if (character === "\t") width += 28;
      else if (/[\u1100-\u11ff\u2e80-\ua4cf\uac00-\ud7af\uf900-\ufaff\ufe10-\ufe6f\uff01-\uff60]/.test(character) || /[\u{1f300}-\u{1faff}]/u.test(character)) width += 14;
      else width += 7;
      if (width >= maximum) return maximum;
    }
    return width;
  };
  let contentWidth = 0;
  for (const value of values || []) contentWidth = Math.max(contentWidth, textWidth(value));
  return Math.max(minimum, Math.min(maximum, Math.ceil(contentWidth + padding)));
}

function browserGridSource() {
  return [
    `const DEFAULT_COLUMN_WIDTH=${DEFAULT_COLUMN_WIDTH};`, `const MIN_COLUMN_WIDTH=${MIN_COLUMN_WIDTH};`,
    `const MAX_COLUMN_WIDTH=${MAX_COLUMN_WIDTH};`, `const ROW_HEADER_WIDTH=${ROW_HEADER_WIDTH};`,
    `const COLUMN_CHUNK=${COLUMN_CHUNK};`, `const MAX_RANGE_COLUMNS=${MAX_RANGE_COLUMNS};`,
    clampColumnWidth.toString(), excelColumnWidthPixels.toString(), createColumnOffsets.toString(), columnAtPixel.toString(), visibleColumnRange.toString()
  ].join('\n');
}

module.exports = { DEFAULT_COLUMN_WIDTH, MIN_COLUMN_WIDTH, MAX_COLUMN_WIDTH, ROW_HEADER_WIDTH, createColumnOffsets, columnAtPixel, visibleColumnRange, clampColumnWidth, excelColumnWidthPixels, estimateBestFitWidth, browserGridSource };
