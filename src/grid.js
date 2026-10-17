'use strict';

const DEFAULT_COLUMN_WIDTH = 144;
const MIN_COLUMN_WIDTH = 48;
const MAX_COLUMN_WIDTH = 640;
const ROW_HEADER_WIDTH = 52;
const COLUMN_CHUNK = 5;
const MAX_RANGE_COLUMNS = 30;

function clampColumnWidth(width, xlsx = false) {
  const value = Number(width);
  return Number.isFinite(value) ? Math.max(xlsx ? 1 : MIN_COLUMN_WIDTH, Math.min(xlsx ? 1785 : MAX_COLUMN_WIDTH, value)) : DEFAULT_COLUMN_WIDTH;
}

function excelColumnWidthPixels(width, defaultWidth = DEFAULT_COLUMN_WIDTH, maximumDigitWidth = 7) {
  const value = Number(width);
  const digitWidth = Number.isFinite(maximumDigitWidth) && maximumDigitWidth >= 1 && maximumDigitWidth <= 600 ? maximumDigitWidth : 7;
  // OOXML stores the padded width in 1/256 character units, using the Normal
  // style font's maximum digit width in pixels at 96dpi.
  return width !== undefined && width !== null && Number.isFinite(value) && value >= 0 && value <= 255 ? Math.floor((256 * value + Math.floor(128 / digitWidth)) / 256 * digitWidth) : defaultWidth;
}

function createColumnOffsets(count, widths = new Map(), hiddenColumns = [], defaultWidth = DEFAULT_COLUMN_WIDTH, preserveWidths = false, scale = 1) {
  const columnCount = Number.isInteger(count) && count > 0 ? count : 0;
  const hidden = new Set(hiddenColumns);
  const offsets = new Float64Array(columnCount + 1);
  for (let column = 0; column < columnCount; column++) {
    const width = widths.get(column) ?? defaultWidth;
    const size = preserveWidths ? Number.isFinite(width) && width >= 0 && width <= 4096 ? width : defaultWidth : clampColumnWidth(width);
    offsets[column + 1] = offsets[column] + (hidden.has(column) ? 0 : size * (Number.isFinite(scale) && scale > 0 ? scale : 1));
  }
  return offsets;
}

function columnAtPixel(offsets, pixel) {
  const count = offsets.length - 1;
  if (count < 1) return -1;
  const position = Number.isFinite(pixel) ? Math.max(0, Math.min(pixel, offsets[count] - Number.EPSILON * Math.max(1, offsets[count]))) : 0;
  if (offsets[count] <= 0) return -1;
  let low = 0, high = count;
  while (low < high) {
    const middle = Math.floor((low + high) / 2);
    if (offsets[middle + 1] <= position) low = middle + 1;
    else high = middle;
  }
  return Math.min(count - 1, low);
}

function visibleColumnRange(offsets, scrollLeft, viewportWidth, options = {}) {
  const count = offsets.length - 1;
  if (count < 1) return { columnStart: 0, columnCount: 0 };
  if (offsets[count] <= 0) return { columnStart: 0, columnCount: Math.min(Number.isInteger(options.maxColumns) && options.maxColumns > 0 ? options.maxColumns : MAX_RANGE_COLUMNS, count) };
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

function createRowView(rowCount, visibleRows, hiddenRows) {
  const count = Number.isInteger(rowCount) && rowCount >= 0 ? rowCount : 0;
  if (visibleRows !== undefined && visibleRows !== null) {
    if (!Array.isArray(visibleRows) || visibleRows.some((row, index) => !Number.isInteger(row) || row < 0 || row >= count || (index > 0 && row <= visibleRows[index - 1]))) throw new RangeError('표시 행 인덱스는 중복 없이 오름차순이어야 합니다.');
    const indexes = new Map(visibleRows.map((row, index) => [row, index]));
    return { rowCount: count, displayCount: visibleRows.length, rowAt: index => Number.isInteger(index) && index >= 0 && index < visibleRows.length ? visibleRows[index] : -1, indexOf: row => indexes.get(row) ?? -1 };
  }
  if (hiddenRows !== undefined && hiddenRows !== null) {
    if (!Array.isArray(hiddenRows) || hiddenRows.some((row, index) => !Number.isInteger(row) || row < 0 || row >= count || (index > 0 && row <= hiddenRows[index - 1]))) throw new RangeError('숨김 행 인덱스는 중복 없이 오름차순이어야 합니다.');
    const upperBound = value => { let low = 0, high = hiddenRows.length; while (low < high) { const middle = (low + high) >> 1; if (hiddenRows[middle] <= value) low = middle + 1; else high = middle; } return low; };
    const lowerBound = value => { let low = 0, high = hiddenRows.length; while (low < high) { const middle = (low + high) >> 1; if (hiddenRows[middle] < value) low = middle + 1; else high = middle; } return low; };
    return {
      rowCount: count,
      displayCount: count - hiddenRows.length,
      rowAt: index => {
        if (!Number.isInteger(index) || index < 0 || index >= count - hiddenRows.length) return -1;
        let low = index, high = Math.min(count - 1, index + hiddenRows.length);
        while (low < high) { const middle = (low + high) >> 1; if (middle - upperBound(middle) < index) low = middle + 1; else high = middle; }
        return low;
      },
      indexOf: row => Number.isInteger(row) && row >= 0 && row < count && (hiddenRows[lowerBound(row)] !== row) ? row - lowerBound(row) : -1
    };
  }
  if (visibleRows === undefined || visibleRows === null) {
    return { rowCount: count, displayCount: count, rowAt: index => index >= 0 && index < count ? index : -1, indexOf: row => row >= 0 && row < count ? row : -1 };
  }
}


function clampRowHeightPixels(value) {
  return Number.isFinite(value) ? Math.min(546, Math.max(1, Math.round(value * .75 * 100) / 100 * 4 / 3)) : 26;
}

function createRowMetrics(rowView, heights = {}, defaultHeight = 26, scale = 1) {
  const normal = Number.isFinite(defaultHeight) && defaultHeight > 0 ? defaultHeight : 26;
  const factor = Number.isFinite(scale) && scale > 0 ? scale : 1;
  const entries = Object.entries(heights).map(([row, height]) => [rowView.indexOf(Number(row)), height])
    .filter(([index, height]) => index >= 0 && Number.isFinite(height) && height > 0 && height !== normal)
    .sort((left, right) => left[0] - right[0]);
  const byIndex = new Map(entries), indexes = entries.map(([index]) => index), sums = [0];
  for (const [, height] of entries) sums.push(sums.at(-1) + (height - normal) * factor);
  const top = index => {
    const target = Math.max(0, Math.min(rowView.displayCount, index));
    let low = 0, high = indexes.length;
    while (low < high) { const middle = (low + high) >> 1; if (indexes[middle] < target) low = middle + 1; else high = middle; }
    return target * normal * factor + sums[low];
  };
  const height = index => (byIndex.get(index) ?? normal) * factor;
  const indexAtPixel = pixel => {
    if (!rowView.displayCount) return -1;
    const position = Number.isFinite(pixel) ? Math.max(0, pixel) : 0;
    let low = 0, high = rowView.displayCount - 1;
    while (low < high) { const middle = Math.floor((low + high) / 2); if (top(middle + 1) <= position) low = middle + 1; else high = middle; }
    return low;
  };
  return { top, height, indexAtPixel, totalHeight: top(rowView.displayCount), uniformHeight: entries.length ? null : normal * factor };
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
    clampColumnWidth.toString(), excelColumnWidthPixels.toString(), createColumnOffsets.toString(), columnAtPixel.toString(), visibleColumnRange.toString(), createRowView.toString(), createRowMetrics.toString(), clampRowHeightPixels.toString()
  ].join('\n');
}

module.exports = { DEFAULT_COLUMN_WIDTH, MIN_COLUMN_WIDTH, MAX_COLUMN_WIDTH, ROW_HEADER_WIDTH, createColumnOffsets, columnAtPixel, visibleColumnRange, createRowView, createRowMetrics, clampRowHeightPixels, clampColumnWidth, excelColumnWidthPixels, estimateBestFitWidth, browserGridSource };
