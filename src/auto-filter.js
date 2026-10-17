'use strict';

const OPERATORS = new Set(['equal', 'notEqual', 'lessThan', 'lessThanOrEqual', 'greaterThan', 'greaterThanOrEqual']);
const ICON_SET_LENGTHS = Object.freeze({
  '3Arrows': 3, '3ArrowsGray': 3, '3Flags': 3, '3TrafficLights1': 3, '3TrafficLights2': 3,
  '3Signs': 3, '3Symbols': 3, '3Symbols2': 3,
  '4Arrows': 4, '4ArrowsGray': 4, '4RedToBlack': 4, '4Rating': 4, '4TrafficLights': 4,
  '5Arrows': 5, '5ArrowsGray': 5, '5Rating': 5, '5Quarters': 5
});

function validateIconFilter(iconFilter) {
  if (!iconFilter || Object.keys(iconFilter).some(key => !['iconSet', 'iconId'].includes(key)) || !Object.hasOwn(ICON_SET_LENGTHS, iconFilter.iconSet) || !Number.isInteger(iconFilter.iconId) || iconFilter.iconId < 0 || iconFilter.iconId >= ICON_SET_LENGTHS[iconFilter.iconSet]) throw new RangeError('아이콘 필터의 집합이나 아이콘 번호가 유효하지 않습니다.');
  return iconFilter;
}

function validateCustomFilter(custom) {
  if (!custom || typeof custom.and !== 'boolean' || Object.keys(custom).some(key => !['and', 'conditions'].includes(key)) || !Array.isArray(custom.conditions) || custom.conditions.length < 1 || custom.conditions.length > 2) throw new RangeError('조건 필터는 한 개 또는 두 개의 조건과 AND/OR 조합을 입력하세요.');
  for (const condition of custom.conditions) if (!condition || Object.keys(condition).some(key => !['operator', 'value'].includes(key)) || !OPERATORS.has(condition.operator) || typeof condition.value !== 'string' || condition.value.length > 32767) throw new RangeError('조건 필터의 비교 연산자나 값이 유효하지 않습니다.');
  return custom;
}

function wildcardTokens(pattern) {
  const characters = Array.from(pattern.toLowerCase()), tokens = [];
  for (let index = 0; index < characters.length; index++) {
    const character = characters[index];
    if (character === '~' && index + 1 < characters.length && ['*', '?', '~'].includes(characters[index + 1])) tokens.push({ literal: characters[++index] });
    else if (character === '*') { if (tokens.at(-1)?.kind !== '*') tokens.push({ kind: '*' }); }
    else if (character === '?') tokens.push({ kind: '?' });
    else tokens.push({ literal: character });
  }
  return tokens;
}

// Literal searches use direct string operations. Complex patterns use a bitset
// automaton rather than backtracking over workbook-controlled text.
function wildcardMatcher(pattern) {
  const tokens = wildcardTokens(pattern), hasWildcards = tokens.some(token => token.kind);
  if (!hasWildcards) { const literal = tokens.map(token => token.literal).join(''); return value => String(value).toLowerCase() === literal; }
  if (tokens.length > 1024) throw new RangeError('와일드카드 필터는 1,024자 이하로 입력하세요.');
  const leading = tokens[0]?.kind === '*', trailing = tokens.at(-1)?.kind === '*', inner = tokens.slice(leading ? 1 : 0, trailing ? -1 : undefined);
  if (inner.every(token => token.literal !== undefined)) {
    const literal = inner.map(token => token.literal).join('');
    return value => { const text = String(value).toLowerCase(); return leading && trailing ? text.includes(literal) : leading ? text.endsWith(literal) : text.startsWith(literal); };
  }
  const points = Int32Array.from(tokens.map(token => token.kind === '*' ? -1 : token.kind === '?' ? -2 : token.literal.codePointAt(0)));
  const indexes = Int32Array.from([...points.keys()].filter(index => points[index] >= 0)).sort((left, right) => points[left] - points[right]);
  let stars = 0n, questions = 0n;
  for (let index = 0; index < points.length; index++) { if (points[index] === -1) stars |= 1n << BigInt(index); else if (points[index] === -2) questions |= 1n << BigInt(index); }
  const accept = 1n << BigInt(points.length), cache = new Map();
  const literalMask = point => {
    if (cache.has(point)) return cache.get(point);
    let low = 0, high = indexes.length, mask = 0n;
    while (low < high) { const middle = (low + high) >>> 1; if (points[indexes[middle]] < point) low = middle + 1; else high = middle; }
    while (low < indexes.length && points[indexes[low]] === point) mask |= 1n << BigInt(indexes[low++]);
    if (cache.size < 64) cache.set(point, mask);
    return mask;
  };
  return value => {
    let state = 1n | ((1n & stars) << 1n);
    for (const character of String(value).toLowerCase()) {
      const next = ((state & (questions | literalMask(character.codePointAt(0)))) << 1n) | (state & stars);
      state = next | ((next & stars) << 1n);
      if (!state) return false;
    }
    return Boolean(state & accept);
  };
}

function validateTopFilter(top10) {
  if (!top10 || Object.keys(top10).some(key => !['top', 'percent', 'value', 'filterValue'].includes(key)) || typeof top10.top !== 'boolean' || typeof top10.percent !== 'boolean' || !Number.isFinite(top10.value) || top10.value <= 0 || top10.value > (top10.percent ? 100 : 500) || !top10.percent && !Number.isInteger(top10.value) || top10.filterValue !== undefined && !Number.isFinite(top10.filterValue)) throw new RangeError('상위·하위 필터는 개수 1~500 또는 0 초과 100 이하의 백분율을 입력하세요.');
  return top10;
}

function numericFilterValue(cell) {
  const value = cell?.value;
  return value !== '' && value !== null && value !== undefined && (typeof value === 'number' || cell?.kind === 'number' && typeof value === 'string') && Number.isFinite(Number(value)) ? Number(value) : null;
}

function* filterColumnCells(filter, column, getCell) {
  for (let row = filter.top + 1; row <= filter.bottom; row++) yield getCell(row, column);
}

function topFilterThreshold(top10, population) {
  validateTopFilter(top10);
  const numbers = []; let count = 0;
  for (const cell of typeof population === 'function' ? population() : population) {
    if (++count > 100000) throw new RangeError('순위 필터는 100,000개 이하의 셀을 사용하세요.');
    const value = numericFilterValue(cell); if (value !== null) numbers.push(value);
  }
  if (!numbers.length) return null;
  numbers.sort((a, b) => a - b);
  const size = Math.min(numbers.length, top10.percent ? Math.max(1, Math.ceil(numbers.length * top10.value / 100)) : top10.value);
  return numbers[top10.top ? numbers.length - size : size - 1];
}

const DATE_GROUP_PARTS = ['year', 'month', 'day', 'hour', 'minute', 'second'];

function validateDateGroups(groups) {
  if (!Array.isArray(groups) || groups.length > 10000) throw new RangeError('날짜 그룹은 10,000개 이하로 입력하세요.');
  for (const group of groups) {
    const index = DATE_GROUP_PARTS.indexOf(group?.dateTimeGrouping);
    if (index < 0 || Object.keys(group).some(key => !['dateTimeGrouping', ...DATE_GROUP_PARTS.slice(0, index + 1)].includes(key)) || DATE_GROUP_PARTS.slice(0, index + 1).some(key => !Number.isInteger(group[key])) || group.year < 1 || group.year > 9999 || index >= 1 && (group.month < 1 || group.month > 12) || index >= 3 && (group.hour < 0 || group.hour > 23) || index >= 4 && (group.minute < 0 || group.minute > 59) || index >= 5 && (group.second < 0 || group.second > 59)) throw new RangeError('날짜 그룹의 연·월·일·시·분·초가 유효하지 않습니다.');
    if (index >= 2) {
      const date = new Date(0); date.setUTCFullYear(group.year, group.month, 0);
      const maximum = group.year === 1900 && group.month === 2 ? 29 : date.getUTCDate();
      if (group.day < 1 || group.day > maximum) throw new RangeError('날짜 그룹의 연·월·일·시·분·초가 유효하지 않습니다.');
    }
  }
  return groups;
}

function parseDateGroups(text) {
  if (typeof text !== 'string' || text.length > 32767) throw new RangeError('날짜 그룹 입력은 32,767자 이하로 입력하세요.');
  const groups = text.split(/\r?\n/).map(line => line.trim()).filter(Boolean).map(line => {
    const match = /^(\d{4})(?:-(\d{2})(?:-(\d{2})(?:[T ](\d{2})(?::(\d{2})(?::(\d{2}))?)?)?)?)?$/.exec(line);
    if (!match) throw new RangeError('날짜 그룹은 YYYY, YYYY-MM, YYYY-MM-DD 또는 날짜와 시각을 한 줄에 하나씩 입력하세요.');
    const fields = match.slice(1).filter(value => value !== undefined);
    return { dateTimeGrouping: DATE_GROUP_PARTS[fields.length - 1], ...Object.fromEntries(fields.map((value, index) => [DATE_GROUP_PARTS[index], Number(value)])) };
  });
  validateDateGroups(groups); return [...new Map(groups.map(group => [JSON.stringify(group), group])).values()];
}

function serialDateParts(serial, date1904 = false) {
  if (!Number.isFinite(serial) || serial < 0) return null;
  serial += date1904 ? 1462 : 0;
  let whole = Math.floor(serial), milliseconds = Math.round((serial - whole) * 86400000);
  whole += Math.floor(milliseconds / 86400000); milliseconds %= 86400000;
  if (whole > 2958465) return null;
  const leapDay = whole === 60, days = whole > 60 ? whole - 1 : whole;
  const date = new Date((leapDay ? Date.UTC(1900, 1, 28) : Date.UTC(1899, 11, 31) + days * 86400000) + milliseconds);
  return { year: leapDay ? 1900 : date.getUTCFullYear(), month: leapDay ? 2 : date.getUTCMonth() + 1, day: leapDay ? 29 : date.getUTCDate(), weekday: leapDay ? 3 : date.getUTCDay(), hour: date.getUTCHours(), minute: date.getUTCMinutes(), second: date.getUTCSeconds() };
}

const DYNAMIC_FILTER_TYPES = ['null', 'today', 'yesterday', 'tomorrow', 'thisWeek', 'lastWeek', 'nextWeek', 'thisMonth', 'lastMonth', 'nextMonth', 'thisQuarter', 'lastQuarter', 'nextQuarter', 'thisYear', 'lastYear', 'nextYear', 'yearToDate', ...Array.from({ length: 4 }, (_, index) => 'Q' + (index + 1)), ...Array.from({ length: 12 }, (_, index) => 'M' + (index + 1)), 'aboveAverage', 'belowAverage'];

function validateDynamicFilter(dynamic) {
  if (!dynamic || !DYNAMIC_FILTER_TYPES.includes(dynamic.type) || Object.keys(dynamic).some(key => !['type', 'value', 'maxValue'].includes(key)) || ['value', 'maxValue'].some(key => dynamic[key] !== undefined && !Number.isFinite(dynamic[key]))) throw new RangeError('동적 필터의 종류나 기준값이 유효하지 않습니다.');
  return dynamic;
}

function dynamicFilterUsesClock(dynamic) { return Boolean(dynamic && /^(?:today|yesterday|tomorrow|this|last|next|yearToDate)/.test(dynamic.type)); }

function validateColorFilter(colorFilter) {
  if (!colorFilter || typeof colorFilter.cellColor !== 'boolean' || colorFilter.color !== null && (typeof colorFilter.color !== 'string' || !/^#[0-9a-f]{6}$/i.test(colorFilter.color)) || Object.keys(colorFilter).some(key => !['cellColor', 'color', 'dxfId'].includes(key)) || colorFilter.dxfId !== undefined && (!Number.isInteger(colorFilter.dxfId) || colorFilter.dxfId < 0 || colorFilter.dxfId > 65535)) throw new RangeError('색상 필터의 대상이나 RGB 색이 유효하지 않습니다.');
  return colorFilter;
}

function dynamicFilterCriteria(dynamic, population = [], options = {}) {
  validateDynamicFilter(dynamic); const type = dynamic.type;
  if (type === 'null') return { kind: 'all' };
  if (/^M(?:[1-9]|1[0-2])$/.test(type)) return { kind: 'month', month: Number(type.slice(1)) };
  if (/^Q[1-4]$/.test(type)) return { kind: 'quarter', quarter: Number(type.slice(1)) };
  if (type === 'aboveAverage' || type === 'belowAverage') {
    const numbers = []; let count = 0, scale = 0;
    for (const cell of typeof population === 'function' ? population() : population) {
      if (++count > 100000) throw new RangeError('동적 필터는 100,000개 이하의 셀을 사용하세요.');
      const value = numericFilterValue(cell); if (value !== null) { numbers.push(value); scale = Math.max(scale, Math.abs(value)); }
    }
    const sum = numbers.reduce((total, value) => total + value, 0);
    const value = !numbers.length ? null : Number.isFinite(sum) ? sum / numbers.length : scale * (numbers.reduce((total, value) => total + value / scale, 0) / numbers.length);
    return { kind: 'average', value };
  }
  const now = options.now ?? new Date();
  if (!(now instanceof Date) || !Number.isFinite(now.getTime())) throw new RangeError('동적 날짜 필터의 기준 시각이 유효하지 않습니다.');
  const year = now.getFullYear(), month = now.getMonth(), day = now.getDate();
  const calendar = (y, m, d) => { const date = new Date(0); date.setUTCHours(0, 0, 0, 0); date.setUTCFullYear(y, m, d); return date; };
  const today = calendar(year, month, day), offset = type.startsWith('last') ? -1 : type.startsWith('next') ? 1 : 0;
  let start, end;
  if (['today', 'yesterday', 'tomorrow'].includes(type)) { const delta = type === 'yesterday' ? -1 : type === 'tomorrow' ? 1 : 0; start = calendar(year, month, day + delta); end = calendar(year, month, day + delta + 1); }
  else if (type.endsWith('Week')) { const delta = -today.getUTCDay() + offset * 7; start = calendar(year, month, day + delta); end = calendar(year, month, day + delta + 7); }
  else if (type.endsWith('Month')) { start = calendar(year, month + offset, 1); end = calendar(year, month + offset + 1, 1); }
  else if (type.endsWith('Quarter')) { const quarter = Math.floor(month / 3) * 3 + offset * 3; start = calendar(year, quarter, 1); end = calendar(year, quarter + 3, 1); }
  else if (type === 'yearToDate') { start = calendar(year, 0, 1); end = calendar(year, month, day + 1); }
  else { start = calendar(year + offset, 0, 1); end = calendar(year + offset + 1, 0, 1); }
  const serial = date => (date.getTime() - Date.UTC(1899, 11, 31)) / 86400000 + (date.getTime() >= Date.UTC(1900, 2, 1) ? 1 : 0) - (options.date1904 ? 1462 : 0);
  return { kind: 'range', value: serial(start), maxValue: serial(end) };
}

function compileAutoFilterColumn(column, population = [], options = {}) {
  if (column.iconFilter) {
    const filter = validateIconFilter(column.iconFilter);
    return cell => {
      const icon = options.getStyle ? options.getStyle(cell)?.icon : cell?.icon;
      return icon?.set === filter.iconSet && icon.id === filter.iconId;
    };
  }
  if (column.colorFilter) {
    const filter = validateColorFilter(column.colorFilter), expected = filter.color?.toUpperCase() ?? null;
    return cell => { const style = options.getStyle ? options.getStyle(cell) : cell?.styleProperties; return (style?.[filter.cellColor ? 'fill' : 'color']?.toUpperCase() ?? null) === expected; };
  }
  if (column.dynamic) {
    const criteria = dynamicFilterCriteria(column.dynamic, population, options);
    return cell => {
      if (criteria.kind === 'all') return true;
      const value = numericFilterValue(cell); if (value === null) return false;
      if (criteria.kind === 'average') return criteria.value !== null && (column.dynamic.type === 'aboveAverage' ? value > criteria.value : value < criteria.value);
      const date = serialDateParts(value, options.date1904); if (!date) return false;
      if (criteria.kind === 'month') return date.month === criteria.month;
      if (criteria.kind === 'quarter') return Math.ceil(date.month / 3) === criteria.quarter;
      return value >= criteria.value && value < criteria.maxValue;
    };
  }
  if (column.top10) {
    const threshold = topFilterThreshold(column.top10, population);
    return cell => { const value = numericFilterValue(cell); return threshold !== null && value !== null && (column.top10.top ? value >= threshold : value <= threshold); };
  }
  if (!column.custom) {
    const allowed = new Set(column.values);
    const groups = new Map();
    if (column.dateGroups) for (const group of validateDateGroups(column.dateGroups)) {
      const size = DATE_GROUP_PARTS.indexOf(group.dateTimeGrouping) + 1;
      if (!groups.has(size)) groups.set(size, new Set());
      groups.get(size).add(DATE_GROUP_PARTS.slice(0, size).map(key => group[key]).join('/'));
    }
    return cell => {
      const raw = cell?.value, value = typeof raw === 'boolean' ? raw ? 'TRUE' : 'FALSE' : raw ?? '';
      if (value === '') return column.blanks;
      if (allowed.has(String(value))) return true;
      if (!groups.size) return false;
      const numeric = numericFilterValue(cell), parts = numeric === null ? null : serialDateParts(numeric, options.date1904);
      return Boolean(parts && [...groups].some(([size, values]) => values.has(DATE_GROUP_PARTS.slice(0, size).map(key => parts[key]).join('/'))));
    };
  }
  validateCustomFilter(column.custom);
  const conjunction = column.custom.and;
  const predicates = column.custom.conditions.map(({ operator, value: operand }) => {
    const numeric = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(operand.trim()) && Number.isFinite(Number(operand));
    const equal = ['equal', 'notEqual'].includes(operator), match = equal ? wildcardMatcher(operand) : null;
    return cell => {
      const raw = cell?.value, value = typeof raw === 'boolean' ? raw ? 'TRUE' : 'FALSE' : raw ?? '', blank = value === '', numericText = options.coerceTextNumbers && typeof value === 'string' && value.trim() !== '' && Number.isFinite(Number(value.trim())), isNumber = cell?.kind === 'number' || typeof value === 'number' || numericText;
      let comparison;
      if (blank) comparison = operand === '' ? 0 : null;
      else if (numeric && isNumber && Number.isFinite(Number(value))) comparison = Math.sign(Number(value) - Number(operand));
      else if (equal) comparison = !isNumber && match(String(value)) ? 0 : null;
      else if (numeric || isNumber) comparison = null;
      else { const left = String(value).toLowerCase(), right = operand.toLowerCase(); comparison = left === right ? 0 : left < right ? -1 : 1; }
      if (operator === 'equal') return comparison === 0;
      if (operator === 'notEqual') return comparison !== 0;
      if (comparison === null) return false;
      return operator === 'lessThan' ? comparison < 0 : operator === 'lessThanOrEqual' ? comparison <= 0 : operator === 'greaterThan' ? comparison > 0 : comparison >= 0;
    };
  });
  return cell => conjunction ? predicates.every(predicate => predicate(cell)) : predicates.some(predicate => predicate(cell));
}

module.exports = { validateCustomFilter, validateTopFilter, topFilterThreshold, validateDateGroups, parseDateGroups, serialDateParts, DYNAMIC_FILTER_TYPES, validateDynamicFilter, dynamicFilterUsesClock, dynamicFilterCriteria, validateColorFilter, ICON_SET_LENGTHS, validateIconFilter, filterColumnCells, compileAutoFilterColumn };
