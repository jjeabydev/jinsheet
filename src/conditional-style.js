'use strict';
const {translateFormulaReferences}=require('./formula-references');

function iconThreshold(stop, statistics, range, getCellValue, evaluateFormula) {
  if (stop.type === 'min') return statistics.minimum;
  if (stop.type === 'max') return statistics.maximum;
  if (stop.type === 'num') return stop.value;
  if (stop.type === 'percent') return statistics.minimum + (statistics.maximum - statistics.minimum) * stop.value / 100;
  if (stop.type === 'percentile') {
    const position = (statistics.sorted.length - 1) * stop.value / 100, lower = Math.floor(position), fraction = position - lower;
    return statistics.sorted[lower] + (statistics.sorted[Math.min(lower + 1, statistics.sorted.length - 1)] - statistics.sorted[lower]) * fraction;
  }
  if (stop.type === 'formula' && evaluateFormula) {
    const result = evaluateFormula(stop.formula, { locale: getCellValue.locale, date1904: Boolean(getCellValue.date1904), now: getCellValue.now, getCell: (row, column, sheetName) => sheetName ? '#REF!' : getCellValue(row, column) });
    return typeof result === 'number' && Number.isFinite(result) ? result : null;
  }
  return null;
}

function iconForValue(rule, value, range, getCellValue, evaluateFormula, sheet) {
  const cached = getCellValue.rangeStatistics?.(range);
  const values = cached?.values ?? (getCellValue.rangeValues
    ? getCellValue.rangeValues(range)
    : [...(sheet.cells?.values?.() || [])].filter(item => item.row >= range.top && item.row <= range.bottom && item.column >= range.left && item.column <= range.right).map(item => getCellValue(item.row, item.column)).filter(candidate => typeof candidate === 'number' && Number.isFinite(candidate)));
  const sorted = cached?.sorted || [...values].sort((left, right) => left - right), statistics = cached || { values, sorted, minimum: sorted[0], maximum: sorted.at(-1) };
  if (!statistics.values.length || !statistics.sorted.length) return null;
  const thresholds = rule.iconSet.thresholds.map(stop => iconThreshold(stop, statistics, range, getCellValue, evaluateFormula));
  if (thresholds.some(threshold => threshold === null)) return null;
  let bucket = 0;
  for (let index = 1; index < thresholds.length; index++) {
    const threshold = thresholds[index], inclusive = rule.iconSet.thresholds[index].gte;
    if (inclusive ? value >= threshold : value > threshold) bucket = index;
  }
  const id = rule.iconSet.reverse ? bucket : thresholds.length - 1 - bucket;
  const customIcon = rule.iconSet.customIcons?.[id];
  return { set: customIcon?.iconSet || rule.iconSet.name, id: customIcon?.iconId ?? id, showValue: rule.iconSet.showValue };
}

function matchesTop10(rule, value, range, getCellValue, sheet) {
  const statistics = getCellValue.rangeStatistics?.(range);
  const values = statistics?.values ?? (getCellValue.rangeValues
    ? getCellValue.rangeValues(range)
    : [...(sheet.cells?.values?.() || [])].filter(item => item.row >= range.top && item.row <= range.bottom && item.column >= range.left && item.column <= range.right).map(item => getCellValue(item.row, item.column)).filter(candidate => typeof candidate === 'number' && Number.isFinite(candidate)));
  const sorted = statistics?.sorted ?? [...values].sort((left, right) => left - right);
  if (!sorted.length) return false;
  const { rank, percent, bottom } = rule.top10;
  const count = Math.min(sorted.length, percent ? Math.max(1, Math.ceil(sorted.length * rank / 100)) : rank);
  const threshold = bottom ? sorted[count - 1] : sorted[sorted.length - count];
  return bottom ? value <= threshold : value >= threshold;
}

function conditionalStyleForCell(sheet, row, column, cell, getCellValue = () => null, evaluateFormula) {
  if (!sheet?.conditionalFormats?.length) return null;
  const number = cell?.kind === 'blank' || !cell ? 0 : cell.kind === 'number' && Number.isFinite(Number(cell.value)) ? Number(cell.value) : null;
  const result = {};
  for (const rule of sheet.conditionalFormats) {
    const matchingRanges = rule.ranges.filter(range => row >= range.top && row <= range.bottom && column >= range.left && column <= range.right);
    if (!matchingRanges.length) continue;
    let matches = false;
    if (rule.type === 'iconSet') {
      if (number !== null && !Object.hasOwn(result, 'icon')) {
        const range = matchingRanges[0], icon = iconForValue(rule, number, range, getCellValue, evaluateFormula, sheet);
        if (icon) result.icon = icon;
      }
      if (rule.stopIfTrue && Object.hasOwn(result, 'icon')) break;
      continue;
    }
    if (rule.type === 'top10') {
      if (cell?.kind === 'number' && number !== null) matches = matchingRanges.some(range => matchesTop10(rule, number, range, getCellValue, sheet));
      if (!matches) continue;
      for (const [key, value] of Object.entries(rule.style)) {
        if (key === 'borders') {
          result.borders ||= {};
          for (const [side, border] of Object.entries(value)) if (!Object.hasOwn(result.borders, side)) result.borders[side] = border;
        } else if (!Object.hasOwn(result, key)) result[key] = value;
      }
      if (rule.stopIfTrue) break;
      continue;
    }
    if (rule.type === 'colorScale') {
      if (cell?.kind === 'number' && number !== null) {
        const range = matchingRanges[0], statistics = getCellValue.rangeStatistics?.(range), values = statistics?.values ?? (getCellValue.rangeValues
          ? getCellValue.rangeValues(range)
          : [...(sheet.cells?.values?.() || [])].filter(item => item.row >= range.top && item.row <= range.bottom && item.column >= range.left && item.column <= range.right && item.kind === 'number').map(item => getCellValue(item.row, item.column)).filter(value => typeof value === 'number' && Number.isFinite(value)));
        if (values.length) {
          let minimum = statistics?.minimum ?? Infinity, maximum = statistics?.maximum ?? -Infinity;
          if (!statistics) for (const value of values) { minimum = Math.min(minimum, value); maximum = Math.max(maximum, value); }
          const thresholds = rule.colorScale.stops.map(stop => {
            if (stop.type === 'min') return minimum;
            if (stop.type === 'max') return maximum;
            if (stop.type === 'num') return stop.value;
            if (stop.type === 'percent') return minimum + (maximum - minimum) * stop.value / 100;
            const sorted = statistics?.sorted || [...values].sort((left, right) => left - right), position = (sorted.length - 1) * stop.value / 100, lower = Math.floor(position), fraction = position - lower;
            return sorted[lower] + (sorted[Math.min(lower + 1, sorted.length - 1)] - sorted[lower]) * fraction;
          });
          const colors = rule.colorScale.colors.map(color => color.slice(1).match(/.{2}/g).map(channel => parseInt(channel, 16)));
          let stopIndex = thresholds.length - 2;
          for (let index = 0; index < thresholds.length - 1; index++) if (number <= thresholds[index + 1]) { stopIndex = index; break; }
          const low = thresholds[stopIndex], high = thresholds[stopIndex + 1], ratio = high === low ? 1 : Math.max(0, Math.min(1, (number - low) / (high - low)));
          const color = colors[stopIndex].map((channel, index) => Math.round(channel + (colors[stopIndex + 1][index] - channel) * ratio).toString(16).padStart(2, '0')).join('').toUpperCase();
          if (!Object.hasOwn(result, 'fill')) result.fill = '#' + color;
        }
      }
      continue;
    }
    if (rule.type === 'dataBar') {
      if (cell?.kind === 'number' && number !== null) {
        const range = matchingRanges[0], statistics = getCellValue.rangeStatistics?.(range), values = statistics?.values ?? (getCellValue.rangeValues
          ? getCellValue.rangeValues(range)
          : [...(sheet.cells?.values?.() || [])].filter(item => item.row >= range.top && item.row <= range.bottom && item.column >= range.left && item.column <= range.right).map(item => getCellValue(item.row, item.column)).filter(value => typeof value === 'number' && Number.isFinite(value)));
        if (values.length) {
          let minimum = statistics?.minimum ?? Infinity, maximum = statistics?.maximum ?? -Infinity;
          if (!statistics) for (const value of values) { minimum = Math.min(minimum, value); maximum = Math.max(maximum, value); }
          let width, start = 0, end, axis = null, negative = false;
          if (minimum < 0) {
            const mixedSigns = maximum > 0;
            axis = mixedSigns ? -minimum / (maximum - minimum) * 100 : 100;
            negative = number < 0;
            const sideMaximum = mixedSigns ? negative ? Math.abs(minimum) : maximum : Math.abs(minimum);
            const fraction = maximum === minimum ? (number === 0 ? 0 : 1) : sideMaximum === 0 ? 0 : Math.min(1, Math.abs(number) / sideMaximum);
            width = fraction === 0 ? 0 : rule.dataBar.minLength + fraction * (rule.dataBar.maxLength - rule.dataBar.minLength);
            if (negative) { end = axis; start = axis - width / 100 * axis; }
            else { start = axis; end = axis + width / 100 * (100 - axis); }
            width = Math.abs(end - start);
          } else {
            const ratio = maximum === minimum ? (maximum === 0 ? 0 : 1) : Math.max(0, Math.min(1, (number - minimum) / (maximum - minimum)));
            width = rule.dataBar.minLength + ratio * (rule.dataBar.maxLength - rule.dataBar.minLength);
            end = width;
          }
          if (!Object.hasOwn(result, 'dataBar')) result.dataBar = { color: rule.dataBar.color, width, ...(axis !== null ? { start, end, axis, negative } : {}), showValue: rule.dataBar.showValue, gradient: rule.dataBar.gradient };
        }
      }
      continue;
    } else if (rule.type === 'expression') {
      for (const range of matchingRanges) {
        const formula = translateFormulaReferences(rule.formula, row - range.top, column - range.left);
        const value = evaluateFormula(formula, { locale: getCellValue.locale, date1904: Boolean(cell?.date1904), now: getCellValue.now, getCell: (cellRow, cellColumn, sheetName) => sheetName ? '#REF!' : getCellValue(cellRow, cellColumn) });
        if (value === true || typeof value === 'number' && value !== 0) { matches = true; break; }
      }
    } else if (cell?.kind === 'number' && number !== null) {
      const [first, second] = rule.values;
      matches = rule.operator === 'equal' ? number === first
        : rule.operator === 'notEqual' ? number !== first
        : rule.operator === 'greaterThan' ? number > first
        : rule.operator === 'lessThan' ? number < first
        : rule.operator === 'greaterThanOrEqual' ? number >= first
        : rule.operator === 'lessThanOrEqual' ? number <= first
        : rule.operator === 'between' ? number >= Math.min(first, second) && number <= Math.max(first, second)
        : number < Math.min(first, second) || number > Math.max(first, second);
    }
    if (!matches) continue;
    for (const [key, value] of Object.entries(rule.style)) {
      if (key === 'borders') {
        result.borders ||= {};
        for (const [side, border] of Object.entries(value)) if (!Object.hasOwn(result.borders, side)) result.borders[side] = border;
      } else if (!Object.hasOwn(result, key)) result[key] = value;
    }
    if (rule.stopIfTrue) break;
  }
  return Object.keys(result).length ? result : null;

}
function createConditionalValueReader(document,calculator,sheetIndex){
  const conditionalRangeCache = new Map();
  const conditionalStatisticsCache = new Map();
  const conditionalCellValue = (row, column) => calculator.value(sheetIndex, row, column);
  conditionalCellValue.now=calculator.now;
  conditionalCellValue.locale=document.locale || 'en-US';
  conditionalCellValue.date1904=Boolean(document.state?.date1904);
  conditionalCellValue.rangeValues = bounds => {
    const key = [bounds.top, bounds.bottom, bounds.left, bounds.right].join(':');
    if (conditionalRangeCache.has(key)) return conditionalRangeCache.get(key);
    const sheet = document.state.sheets[sheetIndex], coordinates = new Map();
    for (const cell of sheet.cells.values()) if (cell.row >= bounds.top && cell.row <= bounds.bottom && cell.column >= bounds.left && cell.column <= bounds.right) coordinates.set(cell.row + ':' + cell.column, [cell.row, cell.column]);
    for (const draftKey of [...(document.drafts?.keys() || []), ...(document.formulaDrafts?.keys() || [])]) {
      const prefix = sheetIndex + ':', address = draftKey.startsWith(prefix) ? draftKey.slice(prefix.length) : '', match = /^([A-Z]+)([1-9]\d*)$/i.exec(address);
      if (!match) continue;
      let column = 0;
      for (const letter of match[1].toUpperCase()) column = column * 26 + letter.charCodeAt(0) - 64;
      const row = Number(match[2]) - 1; column--;
      if (row >= bounds.top && row <= bounds.bottom && column >= bounds.left && column <= bounds.right) coordinates.set(row + ':' + column, [row, column]);
    }
    const values = [];
    for (const [row, column] of coordinates.values()) {
      const value = conditionalCellValue(row, column);
      if (typeof value === 'number' && Number.isFinite(value)) values.push(value);
    }
    conditionalRangeCache.set(key, values);
    return values;
  };
  conditionalCellValue.rangeStatistics = bounds => {
    const key = [bounds.top, bounds.bottom, bounds.left, bounds.right].join(':');
    if (!conditionalStatisticsCache.has(key)) {
      const values = conditionalCellValue.rangeValues(bounds), sorted = [...values].sort((left, right) => left - right);
      conditionalStatisticsCache.set(key, { values, sorted, minimum: sorted[0], maximum: sorted.at(-1) });
    }
    return conditionalStatisticsCache.get(key);
  };
  return conditionalCellValue;
}
function createFilterStyleReader(document, calculator, evaluateFormula) {
  const readers = new Map(), cache = new Map();
  return (row, column, sheetIndex = 0) => {
    const key = sheetIndex + ':' + row + ':' + column;
    if (cache.has(key)) return cache.get(key);
    if (!readers.has(sheetIndex)) readers.set(sheetIndex, createConditionalValueReader(document, calculator, sheetIndex));
    const cell = calculator.cell(sheetIndex, row, column);
    const conditional = conditionalStyleForCell(document.state.sheets[sheetIndex], row, column, cell, readers.get(sheetIndex), evaluateFormula);
    const style = { ...(cell?.styleProperties || {}), ...(conditional || {}) };
    if (cache.size < 100000) cache.set(key, style);
    return style;
  };
}
module.exports = { conditionalStyleForCell, createConditionalValueReader, createFilterStyleReader };
