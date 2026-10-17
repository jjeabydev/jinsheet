'use strict';

const { excelColumnWidthPixels } = require('./grid');

// Sparse prefix sums keep drawings near the end of a worksheet bounded in memory.
function printAxis(count, defaultSize, overrides, hidden) {
  const hiddenSet = new Set(hidden || []), indexes = [...new Set([...Object.keys(overrides || {}).map(Number), ...hiddenSet])].filter(index => Number.isInteger(index) && index >= 0 && index < count).sort((a, b) => a - b);
  const sums = [0];
  const size = index => hiddenSet.has(index) ? 0 : overrides[index] ?? defaultSize;
  for (const index of indexes) sums.push(sums.at(-1) + size(index) - defaultSize);
  const start = index => {
    let low = 0, high = indexes.length;
    while (low < high) { const middle = (low + high) >>> 1; if (indexes[middle] < index) low = middle + 1; else high = middle; }
    return index * defaultSize + sums[low];
  };
  const at = pixel => {
    let low = 0, high = count;
    while (low < high) { const middle = (low + high) >>> 1; if (start(middle + 1) <= pixel) low = middle + 1; else high = middle; }
    return low;
  };
  return { size, start, at, count };
}

function drawingAxes(model) {
  const columns = printAxis(16384, model.defaultColumnWidth ?? 144, Object.fromEntries(Object.entries(model.columnWidths || {}).map(([index, width]) => [index, excelColumnWidthPixels(width, model.defaultColumnWidth ?? 144, model.maximumDigitWidth ?? 7)])), model.hiddenColumns);
  const rows = printAxis(1048576, model.defaultRowHeight || 26, model.rowHeights || {}, model.hiddenRows);
  return { rows, columns };
}

function drawingPoint(value, axes) {
  return { x: axes.columns.start(value.column) + (axes.columns.size(value.column) ? value.x : 0), y: axes.rows.start(value.row) + (axes.rows.size(value.row) ? value.y : 0) };
}

function drawingMarker(point, axes) {
  if (![point.x, point.y].every(Number.isFinite) || point.x < 0 || point.y < 0 || point.x > axes.columns.start(axes.columns.count) || point.y > axes.rows.start(axes.rows.count)) throw new RangeError('고정 차트가 시트 범위를 벗어나 셀 크기를 변경할 수 없습니다.');
  const column = axes.columns.at(Math.min(point.x, axes.columns.start(axes.columns.count) - .00001)), row = axes.rows.at(Math.min(point.y, axes.rows.start(axes.rows.count) - .00001));
  if (column >= axes.columns.count || row >= axes.rows.count) throw new RangeError('고정 차트 위치를 표시할 셀이 없습니다.');
  return { column, row, x: Math.round((point.x - axes.columns.start(column)) * 9525) / 9525, y: Math.round((point.y - axes.rows.start(row)) * 9525) / 9525 };
}

function chartPrintGeometry(model, anchors, baseline = model) {
  const axes = drawingAxes(model), { rows, columns } = axes;
  const baseAxes = anchors.some(anchor => anchor?.editAs && anchor.editAs !== 'twoCell') ? drawingAxes(baseline) : axes;
  const boxes = anchors.map(anchor => {
    if (!anchor) throw new RangeError('차트 위치를 해석할 수 없습니다. 차트 포함을 해제하고 인쇄하세요.');
    if (!anchor.printable) return null;
    let from = anchor.type === 'absoluteAnchor' ? { x: anchor.x, y: anchor.y } : drawingPoint(anchor.from, axes);
    let to = anchor.type === 'twoCellAnchor' ? drawingPoint(anchor.to, axes) : { x: from.x + anchor.width, y: from.y + anchor.height };
    if (anchor.type === 'twoCellAnchor' && ['oneCell', 'absolute'].includes(anchor.editAs)) {
      const baseFrom = drawingPoint(anchor.from, baseAxes), baseTo = drawingPoint(anchor.to, baseAxes);
      if (anchor.editAs === 'absolute') from = baseFrom;
      to = { x: from.x + baseTo.x - baseFrom.x, y: from.y + baseTo.y - baseFrom.y };
    }
    if (![from.x, from.y, to.x, to.y].every(Number.isFinite) || from.x < 0 || from.y < 0 || to.x < from.x || to.y < from.y || to.x > columns.start(columns.count) || to.y > rows.start(rows.count)) throw new RangeError('차트 위치나 크기가 유효하지 않습니다. 차트 포함을 해제하고 인쇄하세요.');
    return to.x === from.x || to.y === from.y ? null : { x: from.x, y: from.y, width: to.x - from.x, height: to.y - from.y, bottom: rows.at(to.y - .00001), right: columns.at(to.x - .00001) };
  });
  return { boxes, rows, columns };
}

// Titles and body can have a gap in source coordinates. Clip each contiguous run
// separately, so a chart cannot bridge omitted cells or paint over headings.
function printAxisRuns(indexes, sizes, positions, heading) {
  const runs = []; let target = heading;
  for (const index of indexes) {
    const source = positions[index], length = sizes[index], previous = runs.at(-1);
    if (previous && Math.abs(previous.source + previous.length - source) < .00001) previous.length += length;
    else runs.push({ source, target, length });
    target += length;
  }
  return runs;
}

function chartPrintLayers(charts, xRuns, yRuns, budget) {
  let html = '';
  for (const [index, chart] of charts.entries()) for (const y of yRuns) for (const x of xRuns) {
    const left = Math.max(x.source, chart.x), top = Math.max(y.source, chart.y), right = Math.min(x.source + x.length, chart.x + chart.width), bottom = Math.min(y.source + y.length, chart.y + chart.height);
    if (right <= left || bottom <= top) continue;
    budget.pieces = (budget.pieces || 0) + 1; budget.chartCharacters = (budget.chartCharacters || 0) + chart.html.length;
    if (budget.pieces > 5000 || budget.chartCharacters > 4000000) throw new RangeError('차트 인쇄 출력이 너무 큽니다. 범위를 줄이거나 차트 포함을 해제하세요.');
    html += '<div class="print-chart-clip" data-print-chart="' + index + '" style="left:' + (x.target + left - x.source) + 'px;top:' + (y.target + top - y.source) + 'px;width:' + (right - left) + 'px;height:' + (bottom - top) + 'px"><div class="print-chart" style="left:' + (chart.x - left) + 'px;top:' + (chart.y - top) + 'px;width:' + chart.width + 'px;height:' + chart.height + 'px">' + chart.html + '</div></div>';
  }
  return html;
}

module.exports = { drawingAxes, drawingPoint, drawingMarker, chartPrintGeometry, printAxisRuns, chartPrintLayers };
