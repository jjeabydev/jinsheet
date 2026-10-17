'use strict';
const { filterColumnCells, compileAutoFilterColumn } = require('./auto-filter');
const { createFilterStyleReader } = require('./conditional-style');

const MAX_FORMULA_LENGTH = 100000;
const MAX_RANGE_CELLS = 10000;
const MAX_TOKENS = 20000;
const MAX_CALC_STEPS = 100000;
const MAX_DISTRIBUTION_SIZE = 100000;
const FORMULA_RANGE = Symbol('formulaRange');
const FORMULA_RANGE_METADATA = Symbol('formulaRangeMetadata');
const FORMULA_TEXT_LITERAL = Symbol('formulaTextLiteral');
const FORMULA_REFERENCE_VALUE = Symbol('formulaReferenceValue');
const ERROR_TYPE_CODES = new Map([['#NULL!', 1], ['#DIV/0!', 2], ['#VALUE!', 3], ['#REF!', 4], ['#NAME?', 5], ['#NUM!', 6], ['#N/A', 7], ['#N/A!', 7], ['#GETTING_DATA', 8], ['#SPILL!', 9], ['#CALC!', 14]]);

function columnNumber(letters) {
  let result = 0;
  for (const letter of letters.toUpperCase()) result = result * 26 + letter.charCodeAt(0) - 64;
  return result;
}

function tokenize(source) {
  if (typeof source !== 'string' || source.length > MAX_FORMULA_LENGTH) throw new Error('#VALUE!');
  const text = source.startsWith('=') ? source.slice(1) : source;
  const tokens = [];
  for (let index = 0; index < text.length;) {
    if (/\s/.test(text[index])) { index++; continue; }
    const rest = text.slice(index);
    const string = /^"(?:[^"]|"")*"/.exec(rest);
    if (string) { tokens.push({ type: 'value', value: string[0].slice(1, -1).replaceAll('""', '"') }); index += string[0].length; continue; }
    const error = /^#(?:REF!|DIV\/0!|VALUE!|NAME\?|NUM!|N\/A!?|NULL!|CALC!|CIRC!|SPILL!|GETTING_DATA)/i.exec(rest);
    if (error) { tokens.push({ type: 'error', value: error[0].toUpperCase() }); index += error[0].length; continue; }
    const tablePrefix = /^[\p{L}_][\p{L}\p{N}_.]*/u.exec(rest), tableName = tablePrefix && rest[tablePrefix[0].length] === '[' ? tablePrefix[0] : null;
    const structuredStart = tableName ? tableName.length : 0;
    if (rest[structuredStart] === '[' && (!tableName || !/^[A-Z]{1,3}[1-9]\d*$/i.test(tableName))) {
      let depth = 0, end = structuredStart;
      for (; end < rest.length; end++) {
        if (rest[end] === "'" && /['\[\]#@]/.test(rest[end + 1] || '')) { end++; continue; }
        if (rest[end] === '[') depth++;
        else if (rest[end] === ']' && --depth === 0) { end++; break; }
      }
      if (depth === 0) { tokens.push({ type: 'structuredReference', tableName, spec: rest.slice(structuredStart, end) }); index += end; continue; }
    }
    const quotedSheet = /^'((?:[^']|'')+)'!/.exec(rest);
    const plainSheet = quotedSheet ? null : /^([\p{L}_][\p{L}\p{N}_.]*)!/u.exec(rest);
    const sheetPrefix = quotedSheet || plainSheet;
    const sheetName = quotedSheet ? quotedSheet[1].replaceAll("''", "'") : plainSheet?.[1];
    const offset = sheetPrefix?.[0].length || 0;
    const referenceText = rest.slice(offset);
    const fullRow = /^\$?([1-9]\d*)/.exec(referenceText);
    if (fullRow && (/^\s*:/.test(referenceText.slice(fullRow[0].length)) || tokens.at(-1)?.type === ':')) {
      const row = Number(fullRow[1]);
      tokens.push(row <= 1048576 ? { type: 'reference', row: row - 1, column: 0, axis: 'row', sheetName } : { type: 'error', value: '#REF!' });
      index += offset + fullRow[0].length; continue;
    }
    const fullColumn = /^\$?([A-Za-z]{1,3})/.exec(referenceText);
    if (fullColumn && /^(?:\s*(?::|[,)]|$))/.test(referenceText.slice(fullColumn[0].length)) && (/^\s*:/.test(referenceText.slice(fullColumn[0].length)) || tokens.at(-1)?.type === ':')) {
      const column = columnNumber(fullColumn[1]);
      tokens.push(column <= 16384 ? { type: 'reference', row: 0, column: column - 1, axis: 'column', sheetName } : { type: 'error', value: '#REF!' });
      index += offset + fullColumn[0].length; continue;
    }
    const reference = /^\$?([A-Za-z]{1,3})\$?([1-9]\d*)/.exec(rest.slice(offset));
    const referenceEnd = offset + (reference?.[0].length || 0);
    if (reference && !/[A-Za-z0-9_.]/.test(rest[referenceEnd] || '') && rest[referenceEnd] !== '(') {
      const column = columnNumber(reference[1]), row = Number(reference[2]);
      if (column > 16384 || row > 1048576) tokens.push({ type: 'error', value: '#REF!' });
      else tokens.push({ type: 'reference', row: row - 1, column: column - 1, sheetName });
      index += referenceEnd; continue;
    }
    const number = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/.exec(rest);
    if (number) { const value = Number(number[0]); if (!Number.isFinite(value)) throw new Error('#NUM!'); tokens.push({ type: 'value', value }); index += number[0].length; continue; }
    const name = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(rest);
    if (name) { tokens.push({ type: 'name', value: name[0].toUpperCase() }); index += name[0].length; continue; }
    const operator = /^(?:<=|>=|<>|[+\-*/^&=<>:,()%#])/.exec(rest);
    if (!operator) throw new Error('#VALUE!');
    tokens.push({ type: operator[0], value: operator[0] }); index += operator[0].length;
    if (tokens.length > MAX_TOKENS) throw new Error('#VALUE!');
  }
  tokens.push({ type: 'eof' });
  return tokens;
}

function isError(value) { if (value?.[FORMULA_TEXT_LITERAL] || value?.[FORMULA_REFERENCE_VALUE]) value = value.value; return typeof value === 'string' && /^#(?:REF!|DIV\/0!|VALUE!|NAME\?|NUM!|N\/A|NULL!|CALC!|CIRC!|SPILL!|GETTING_DATA)/.test(value); }
function failValue() { return '#VALUE!'; }
function numeric(value) {
  if (value?.[FORMULA_TEXT_LITERAL] || value?.[FORMULA_REFERENCE_VALUE]) value = value.value;
  if (isError(value)) return value;
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return Number.isFinite(value) ? value : '#NUM!';
  if (typeof value === 'boolean') return value ? 1 : 0;
  if (typeof value === 'string' && value.trim() !== '' && Number.isFinite(Number(value.trim()))) return Number(value.trim());
  return '#VALUE!';
}
function isFormulaRange(value) { return Boolean(value?.[FORMULA_RANGE]); }
function flatten(values) { return values.flatMap(value => isFormulaRange(value) ? value.entries.map(entry => entry.value) : Array.isArray(value) ? flatten(value) : [value]); }
function matrixValue(value) {
  if (isFormulaRange(value)) {
    const { top, bottom, left, right } = value.bounds, height = bottom - top + 1, width = right - left + 1;
    if (height < 1 || width < 1 || height * width > MAX_RANGE_CELLS) return null;
    const result = Array.from({ length: height }, () => Array(width).fill(null));
    for (const entry of value.entries) result[entry.row - top][entry.column - left] = entry.value;
    return result;
  }
  if (!Array.isArray(value)) return null;
  const result = value.every(Array.isArray) ? value : [value];
  if (!result.length || !result[0]?.length || result.some(row => row.length !== result[0].length) || result.length * result[0].length > MAX_RANGE_CELLS) return null;
  return result;
}
function structuredReferenceParts(spec) {
  if (typeof spec !== 'string' || spec[0] !== '[' || spec.at(-1) !== ']') return null;
  let content = spec.slice(1, -1), thisRow = false;
  if (content.startsWith('@')) { thisRow = true; content = content.slice(1); }
  const decode = value => value.replace(/'(['\[\]#@])/g, '$1');
  const parts = []; let columnRange = false;
  if (content.trimStart().startsWith('[')) {
    let index = 0;
    while (index < content.length) {
      while (/\s/.test(content[index] || '') && index < content.length) index++;
      if (content[index++] !== '[') return null;
      const start = index; let raw = '', closed = false;
      for (; index < content.length; index++) {
        if (content[index] === "'" && /['\[\]#@]/.test(content[index + 1] || '')) { raw += content.slice(index, index + 2); index++; continue; }
        if (content[index] === ']') { index++; closed = true; break; }
        if (content[index] === '[') return null;
        raw += content[index];
      }
      if (!closed || index <= start) return null;
      parts.push({ value: decode(raw), selector: raw.startsWith('#'), row: raw.startsWith('@') });
      while (/\s/.test(content[index] || '') && index < content.length) index++;
      if (index === content.length) break;
      const separator = content[index++]; if (![',', ':'].includes(separator)) return null;
      if (separator === ':') { if (columnRange) return null; columnRange = true; }
      if (index === content.length) return null;
    }
  } else parts.push({ value: decode(content), selector: content.startsWith('#'), row: content.startsWith('@') });
  const selectors = parts.filter(part => part.selector), columns = parts.filter(part => !part.selector && !part.row);
  if (selectors.length > 1 || columnRange && columns.length !== 2 || !columnRange && columns.length > 1) return null;
  const selector = selectors[0]?.value.toLocaleLowerCase().replaceAll(' ', '') || null;
  if (selector && !['#data', '#headers', '#totals', '#all', '#thisrow'].includes(selector)) return null;
  if (selector === '#thisrow' || parts.some(part => part.row)) thisRow = true;
  const columnName = columns[0]?.value || (parts.find(part => part.row)?.value.slice(1)) || null;
  return { selector: selector === '#thisrow' ? '#data' : selector, columnName, columnEndName: columnRange ? columns[1].value : null, thisRow };
}
function weekendDays(value) {
  if (value === undefined) return new Set([5, 6]);
  if (typeof value === 'string' && /^[01]{7}$/.test(value)) {
    if (value === '1111111') return null;
    return new Set([...value].flatMap((flag, index) => flag === '1' ? [index] : []));
  }
  const code = numeric(value);
  if (typeof code !== 'number' || !Number.isInteger(code)) return null;
  if (code >= 1 && code <= 7) { const first = (code + 4) % 7; return new Set([first, (first + 1) % 7]); }
  if (code >= 11 && code <= 17) return new Set([(code + 2) % 7]);
  return null;
}
function isWorkingDay(serial, weekends) {
  const day = Math.floor(serial), sundayZero = ((day - (day > 60 ? 1 : 0)) % 7 + 7) % 7;
  return !weekends.has((sundayZero + 6) % 7);
}
function standardNormalCdf(value) {
  if (value === 0) return 0.5;
  const absolute = Math.abs(value) / Math.SQRT2, t = 1 / (1 + 0.5 * absolute);
  const coefficients = [0.17087277, -0.82215223, 1.48851587, -1.13520398, 0.27886807, -0.18628806, 0.09678418, 0.37409196, 1.00002368];
  const series = coefficients.reduce((result, coefficient) => result * t + coefficient, 0);
  const tail = t * Math.exp(-absolute * absolute - 1.26551223 + t * series);
  const error = 1 - tail;
  return value >= 0 ? (1 + error) / 2 : (1 - error) / 2;
}
function standardNormalInverse(probability) {
  if (!(probability > 0 && probability < 1)) return '#NUM!';
  const a = [-39.69683028665376, 220.9460984245205, -275.9285104469687, 138.357751867269, -30.66479806614716, 2.506628277459239];
  const b = [-54.47609879822406, 161.5858368580409, -155.6989798598866, 66.80131188771972, -13.28068155288572];
  const c = [-0.007784894002430293, -0.3223964580411365, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783];
  const d = [0.007784695709041462, 0.3224671290700398, 2.445134137142996, 3.754408661907416];
  const polynomial = (coefficients, value) => coefficients.reduce((result, coefficient) => result * value + coefficient, 0);
  let result;
  if (probability < 0.02425) {
    const q = Math.sqrt(-2 * Math.log(probability));
    result = polynomial(c, q) / (polynomial(d, q) * q + 1);
  } else if (probability > 1 - 0.02425) {
    const q = Math.sqrt(-2 * Math.log1p(-probability));
    result = -polynomial(c, q) / (polynomial(d, q) * q + 1);
  } else {
    const q = probability - 0.5, r = q * q;
    result = polynomial(a, r) * q / (polynomial(b, r) * r + 1);
  }
  const density = Math.exp(-result * result / 2) / Math.sqrt(2 * Math.PI);
  return result - (standardNormalCdf(result) - probability) / density;
}
function logFactorial(value) {
  let result = 0;
  for (let index = 2; index <= value; index++) result += Math.log(index);
  return result;
}
function compensatedSum(values) {
  let total = 0, compensation = 0;
  for (const value of values) {
    const corrected = value - compensation, next = total + corrected;
    compensation = (next - total) - corrected;
    total = next;
  }
  return total;
}
function binomialDistribution(trials, probability) {
  if (probability === 0) return Array.from({ length: trials + 1 }, (_, index) => index === 0 ? 1 : 0);
  if (probability === 1) return Array.from({ length: trials + 1 }, (_, index) => index === trials ? 1 : 0);
  const mode = Math.min(trials, Math.floor((trials + 1) * probability));
  const combinations = logFactorial(trials) - logFactorial(mode) - logFactorial(trials - mode);
  const probabilities = Array(trials + 1).fill(0);
  probabilities[mode] = Math.exp(combinations + mode * Math.log(probability) + (trials - mode) * Math.log1p(-probability));
  for (let successes = mode; successes > 0; successes--) probabilities[successes - 1] = probabilities[successes] * successes / (trials - successes + 1) * (1 - probability) / probability;
  for (let successes = mode; successes < trials; successes++) probabilities[successes + 1] = probabilities[successes] * (trials - successes) / (successes + 1) * probability / (1 - probability);
  const total = compensatedSum(probabilities);
  if (!(Number.isFinite(total) && total > 0)) return '#NUM!';
  for (let index = 0; index < probabilities.length; index++) probabilities[index] /= total;
  return probabilities;
}
function poissonDistribution(value, mean) {
  if (mean === 0) return Array.from({ length: value + 1 }, (_, index) => index === 0 ? 1 : 0);
  if (mean > MAX_DISTRIBUTION_SIZE || value > MAX_DISTRIBUTION_SIZE) return '#NUM!';
  const mode = Math.floor(mean), probabilities = Array(Math.max(value, mode) + 1).fill(0);
  let logModeProbability;
  if (mode === 0) logModeProbability = -mean;
  else if (mode < 32) logModeProbability = -mean + mode * Math.log(mean) - logFactorial(mode);
  else {
    const delta = mean - mode, inverse = 1 / mode, inverseSquared = inverse * inverse;
    const correction = inverse * (1 / 12 + inverseSquared * (-1 / 360 + inverseSquared * (1 / 1260 + inverseSquared * (-1 / 1680 + inverseSquared / 1188))));
    logModeProbability = mode * Math.log1p(delta / mode) - delta - 0.5 * Math.log(2 * Math.PI * mode) - correction;
  }
  const modeProbability = Math.exp(logModeProbability);
  probabilities[mode] = modeProbability;
  for (let count = mode; count > 0; count--) probabilities[count - 1] = probabilities[count] * count / mean;
  for (let count = mode; count < Math.max(value, mode); count++) probabilities[count + 1] = probabilities[count] * mean / (count + 1);
  return probabilities.slice(0, value + 1);
}
function matrix(value) { return Array.isArray(value) && value.every(Array.isArray) ? value : null; }
function comparable(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'number' || typeof value === 'boolean') return Number(value);
  if (typeof value === 'string' && value.trim() !== '' && Number.isFinite(Number(value.trim()))) return Number(value.trim());
  return typeof value === 'string' ? value.toLocaleLowerCase() : value;
}
function sortCompareValues(left, right, locale = 'en') {
  const rank = value => value === null || value === undefined || value === '' ? 5 : isError(value) ? 4 : typeof value === 'boolean' ? 3 : typeof value === 'number' ? 1 : 2;
  const leftRank = rank(left), rightRank = rank(right);
  if (leftRank !== rightRank) return leftRank - rightRank;
  if (leftRank === 1) return left - right;
  if (leftRank === 2) return String(left).localeCompare(String(right), locale, { sensitivity: 'base' });
  return String(left).localeCompare(String(right));
}
function interpolationFraction(lower, upper, value) {
  const span = upper - lower, offset = value - lower;
  if (Number.isFinite(span) && Number.isFinite(offset) && span !== 0) return offset / span;
  const scale = Math.max(Math.abs(lower), Math.abs(upper), Math.abs(value));
  if (scale === 0) return 0;
  return (value / scale - lower / scale) / (upper / scale - lower / scale);
}
function interpolateNumber(lower, upper, fraction) {
  if (fraction <= 0) return lower;
  if (fraction >= 1) return upper;
  return lower < 0 && upper > 0
    ? lower * (1 - fraction) + upper * fraction
    : lower + (upper - lower) * fraction;
}
function criterionMatch(value, criterion) {
  const text = typeof criterion === 'string' ? criterion : String(criterion ?? '');
  const match = /^(<=|>=|<>|=|<|>)(.*)$/.exec(text);
  const operator = match?.[1] || '=';
  const expected = comparable(match ? match[2] : criterion);
  const actual = comparable(value);
  if (isError(value)) return false;
  if ((operator === '=' || operator === '<>') && typeof expected === 'string' && /[*?]/.test(expected)) {
    if (typeof value !== 'string') return operator === '<>' && value !== null && value !== undefined;
    let pattern = '^';
    for (let index = 0; index < expected.length; index++) {
      const character = expected[index];
      if (character === '~' && /[*?~]/.test(expected[index + 1] || '')) pattern += expected[++index].replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      else if (character === '*') pattern += '.*';
      else if (character === '?') pattern += '.';
      else pattern += character.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }
    const equal = new RegExp(pattern + '$', 'i').test(String(value ?? ''));
    return operator === '<>' ? !equal : equal;
  }
  if (actual === null || actual === undefined) return operator === '<>' && expected !== '';
  if (operator === '=') return actual === expected;
  if (operator === '<>') return actual !== expected;
  if (typeof actual !== typeof expected) return false;
  if (operator === '<') return actual < expected;
  if (operator === '>') return actual > expected;
  if (operator === '<=') return actual <= expected;
  return actual >= expected;
}
function criteriaPairs(criteriaRange, sumRange) {
  const criteriaMetadata = criteriaRange?.[FORMULA_RANGE_METADATA];
  const sumMetadata = sumRange?.[FORMULA_RANGE_METADATA];
  if (criteriaMetadata && sumMetadata) {
    const criteriaBounds = criteriaMetadata.bounds, sumBounds = sumMetadata.bounds;
    const criteriaValues = new Map(criteriaMetadata.entries.map(entry => [`${entry.row}:${entry.column}`, entry.value]));
    const sumValues = new Map(sumMetadata.entries.map(entry => [`${entry.row - sumBounds.top}:${entry.column - sumBounds.left}`, entry.value]));
    const pairs = [];
    for (let row = criteriaBounds.top; row <= criteriaBounds.bottom; row++) {
      for (let column = criteriaBounds.left; column <= criteriaBounds.right; column++) {
        const criteria = criteriaValues.get(`${row}:${column}`) ?? null;
        const sumRow = row - criteriaBounds.top, sumColumn = column - criteriaBounds.left;
        pairs.push([criteria, sumValues.get(`${sumRow}:${sumColumn}`) ?? null]);
      }
    }
    return pairs;
  }
  const criteria = flatten([criteriaRange]), sums = flatten([sumRange]);
  if (criteria.length !== sums.length) return null;
  return criteria.map((value, index) => [value, sums[index]]);
}
function sparseRangeIndex(range) {
  return new Map(range.entries.map(entry => [(entry.row - range.bounds.top) + ':' + (entry.column - range.bounds.left), entry.value]));
}
function sameSparseShape(ranges) {
  return ranges.every(range => isFormulaRange(range) && range.bounds.bottom - range.bounds.top === ranges[0].bounds.bottom - ranges[0].bounds.top && range.bounds.right - range.bounds.left === ranges[0].bounds.right - ranges[0].bounds.left);
}
function sparsePositionUnion(ranges) {
  const positions = new Set();
  for (const range of ranges) for (const entry of range.entries) positions.add((entry.row - range.bounds.top) + ':' + (entry.column - range.bounds.left));
  return positions;
}
function lookupValue(value, approximate) {
  if (isError(value)) return value;
  return comparable(value);
}
function lookupIndex(values, target, approximate, lastEqual = false) {
  const sought = lookupValue(target, approximate);
  if (isError(sought)) return sought;
  if (!approximate) return values.findIndex(value => typeof sought === 'string' && /[*?~]/.test(sought)
    ? criterionMatch(value, target)
    : lookupValue(value, false) === sought);
  let found = -1;
  for (let index = 0; index < values.length; index++) {
    const candidate = lookupValue(values[index], true);
    if (isError(candidate)) continue;
    if (typeof candidate !== typeof sought) continue;
    if (candidate === sought) {
      if (!lastEqual) return index;
      found = index;
      continue;
    }
    if (candidate < sought) found = index;
    else break;
  }
  return found;
}
function aggregateNumbers(values) {
  const result = [];
  for (const value of flatten(values)) {
    if (isError(value)) return value;
    if (typeof value === 'number' && Number.isFinite(value)) result.push(value);
    else if (typeof value === 'boolean') result.push(value ? 1 : 0);
  }
  return result;
}
function aggregateDirectArguments(args) {
  const result = [];
  for (const argument of args) {
    if (argument?.[FORMULA_TEXT_LITERAL]) {
      const value = numeric(argument.value);
      if (typeof value !== 'number') return value;
      result.push(value);
      continue;
    }
    if (argument?.[FORMULA_REFERENCE_VALUE]) {
      const value = argument.value;
      if (isError(value)) return value;
      if (typeof value === 'number' && Number.isFinite(value)) result.push(value);
      continue;
    }
    if (!isFormulaRange(argument) && !Array.isArray(argument)) {
      if (isError(argument)) return argument;
      if (typeof argument === 'number' && Number.isFinite(argument)) result.push(argument);
      else if (typeof argument === 'boolean') result.push(argument ? 1 : 0);
      else if (typeof argument === 'string') {
        const value = numeric(argument);
        if (typeof value !== 'number') return value;
        result.push(value);
      }
      continue;
    }
    for (const value of flatten([argument])) {
      if (isError(value)) return value;
      if (typeof value === 'number' && Number.isFinite(value)) result.push(value);
    }
  }
  return result;
}
function countDirectArguments(args) {
  let count = 0;
  for (const argument of args) {
    if (argument?.[FORMULA_TEXT_LITERAL]) {
      if (isError(argument.value)) return argument.value;
      const value = numeric(argument.value);
      if (typeof value === 'number') count++;
      continue;
    }
    if (argument?.[FORMULA_REFERENCE_VALUE]) {
      if (typeof argument.value === 'number' && Number.isFinite(argument.value)) count++;
      continue;
    }
    if (!isFormulaRange(argument) && !Array.isArray(argument) && typeof argument === 'boolean') { count++; continue; }
    for (const value of flatten([argument])) if (typeof value === 'number' && Number.isFinite(value)) count++;
  }
  return count;
}
function varianceValue(numbers, sample, returnStandardDeviation = false) {
  const divisor = numbers.length - (sample ? 1 : 0);
  const directSum = numbers.reduce((sum, value) => sum + value, 0);
  if (Number.isFinite(directSum)) {
    const mean = directSum / numbers.length;
    const sumSquares = numbers.reduce((sum, value) => sum + (value - mean) ** 2, 0);
    const variance = sumSquares / divisor;
    if (Number.isFinite(variance)) return returnStandardDeviation ? Math.sqrt(variance) : variance;
  }
  const scale = Math.max(...numbers.map(Math.abs));
  if (scale === 0) return 0;
  const normalized = numbers.map(value => value / scale);
  const mean = normalized.reduce((sum, value) => sum + value, 0) / numbers.length;
  const normalizedVariance = normalized.reduce((sum, value) => sum + (value - mean) ** 2, 0) / divisor;
  if (normalizedVariance === 0) return 0;
  const standardDeviation = scale * Math.sqrt(normalizedVariance);
  if (!Number.isFinite(standardDeviation)) return '#NUM!';
  if (returnStandardDeviation) return standardDeviation;
  const variance = standardDeviation * standardDeviation;
  return Number.isFinite(variance) ? variance : '#NUM!';
}
function financialCashFlows(args) {
  const entries = [];
  let length = 0;
  for (const argument of args) {
    if (isFormulaRange(argument)) {
      const { top, left, right } = argument.bounds, width = right - left + 1;
      for (const entry of argument.entries) {
        const index = length + (entry.row - top) * width + (entry.column - left);
        if (isError(entry.value)) return entry.value;
        if (typeof entry.value === 'number' && Number.isFinite(entry.value) && entry.value !== 0) entries.push({ index, value: entry.value });
      }
      length += argument.cellCount;
      continue;
    }
    const values = flatten([argument]);
    for (let index = 0; index < values.length; index++) {
      const value = values[index];
      if (isError(value)) return value;
      if (typeof value === 'number' && Number.isFinite(value) && value !== 0) entries.push({ index: length + index, value });
    }
    length += values.length;
  }
  return { entries, length };
}
function financialDatePairs(valuesArgument, datesArgument) {
  let values, dates, length;
  if (isFormulaRange(valuesArgument) && isFormulaRange(datesArgument)) {
    if (!sameSparseShape([valuesArgument, datesArgument])) return '#NUM!';
    length = valuesArgument.cellCount;
    if (!Number.isSafeInteger(length) || length < 1 || length > MAX_DISTRIBUTION_SIZE) return '#NUM!';
    const valueMap = sparseRangeIndex(valuesArgument), dateMap = sparseRangeIndex(datesArgument);
    values = Array.from({ length }, (_, index) => valueMap.get(Math.floor(index / (valuesArgument.bounds.right - valuesArgument.bounds.left + 1)) + ':' + index % (valuesArgument.bounds.right - valuesArgument.bounds.left + 1)) ?? null);
    dates = Array.from({ length }, (_, index) => dateMap.get(Math.floor(index / (datesArgument.bounds.right - datesArgument.bounds.left + 1)) + ':' + index % (datesArgument.bounds.right - datesArgument.bounds.left + 1)) ?? null);
  } else {
    const valueMatrix = matrixValue(valuesArgument) ?? [[valuesArgument]], dateMatrix = matrixValue(datesArgument) ?? [[datesArgument]];
    if (valueMatrix.length !== dateMatrix.length || valueMatrix.some((row, index) => row.length !== dateMatrix[index].length)) return '#NUM!';
    values = valueMatrix.flat();
    dates = dateMatrix.flat();
    length = values.length;
    if (!length || length > MAX_DISTRIBUTION_SIZE) return '#NUM!';
  }
  const entries = [];
  for (let index = 0; index < length; index++) {
    if (isError(values[index]) || isError(dates[index])) return isError(values[index]) ? values[index] : dates[index];
    if (typeof dates[index] !== 'number' || !Number.isFinite(dates[index]) || dates[index] < 0 || dates[index] >= 2958466) return '#NUM!';
    if (typeof values[index] === 'number' && Number.isFinite(values[index]) && values[index] !== 0) entries.push({ index, value: values[index], date: dates[index] });
  }
  return { entries, length, startDate: dates[0] };
}
function pairedNumbers(left, right) {
  const pairs = [];
  if (isFormulaRange(left) || isFormulaRange(right)) {
    if (!isFormulaRange(left) || !isFormulaRange(right) || !sameSparseShape([left, right])) return '#N/A';
    const leftValues = sparseRangeIndex(left), rightValues = sparseRangeIndex(right);
    for (const position of new Set([...leftValues.keys(), ...rightValues.keys()])) {
      const x = leftValues.get(position), y = rightValues.get(position);
      if (isError(x)) return x;
      if (isError(y)) return y;
      if (typeof x === 'number' && Number.isFinite(x) && typeof y === 'number' && Number.isFinite(y)) pairs.push([x, y]);
    }
    return pairs;
  }
  const leftValues = flatten([left]), rightValues = flatten([right]);
  if (leftValues.length !== rightValues.length) return '#N/A';
  for (let index = 0; index < leftValues.length; index++) {
    const x = leftValues[index], y = rightValues[index];
    if (isError(x)) return x;
    if (isError(y)) return y;
    if (typeof x === 'number' && Number.isFinite(x) && typeof y === 'number' && Number.isFinite(y)) pairs.push([x, y]);
  }
  return pairs;
}

function logical(value) {
  if (isError(value)) return value;
  if (typeof value === 'boolean') return value;
  if (typeof value === 'number') return value !== 0;
  if (value === null || value === undefined || value === '') return false;
  if (typeof value === 'string' && /^(TRUE|FALSE)$/i.test(value.trim())) return value.trim().toUpperCase() === 'TRUE';
  return '#VALUE!';
}

function formatTextDateTime(value, format, locale = 'en-US', date1904 = false) {
  const text = String(format).trim(), timeSeparator = text.search(/\s+(?=h{1,2}:mm)/i), hasDateAndTime = timeSeparator >= 0;
  const dateFormat = hasDateAndTime ? text.slice(0, timeSeparator).trim() : text, timeFormat = hasDateAndTime ? text.slice(timeSeparator).trim() : text;
  const datePattern = /^(?:yyyy|yy)[-/.](?:m|mm)[-/.](?:d|dd)$/i.test(dateFormat)
    || /^(?:m|mm)[-/.](?:d|dd)[-/.](?:yyyy|yy)$/i.test(dateFormat)
    || /^(?:d|dd)[ -](?:mmm|mmmm)[ -](?:yyyy|yy)$/i.test(dateFormat);
  const timePattern = /^(?:h|hh):mm(?::ss)?(?:\s+AM\/PM)?$/i.test(hasDateAndTime ? timeFormat : text);
  if (hasDateAndTime && (!datePattern || !timePattern) || !hasDateAndTime && !datePattern && !timePattern) return null;
  const epochOffset = date1904 ? 1462 : 0;
  if (!Number.isFinite(value) || value < 0 || value > 2958465.999999 - epochOffset) return '#NUM!';
  const dateText = serial => {
    const date = dateParts(Math.floor(serial) + epochOffset);
    if (typeof date === 'string') return date;
    const monthName = (length) => new Intl.DateTimeFormat(locale, { month: length === 3 ? 'short' : 'long', timeZone: 'UTC' }).format(utcDate(date.year, date.month, 15));
    return dateFormat.replace(/yyyy|yy|mmmm|mmm|mm|m|dd|d/gi, token => {
      const lower = token.toLowerCase();
      if (lower === 'yyyy') return String(date.year).padStart(4, '0');
      if (lower === 'yy') return String(date.year % 100).padStart(2, '0');
      if (lower === 'mmmm' || lower === 'mmm') return monthName(lower.length);
      if (lower === 'mm') return String(date.month).padStart(2, '0');
      if (lower === 'm') return String(date.month);
      if (lower === 'dd') return String(date.day).padStart(2, '0');
      return String(date.day);
    });
  };
  const timeText = totalSeconds => {
    const secondsInDay = ((totalSeconds % 86400) + 86400) % 86400;
    let hour = Math.floor(secondsInDay / 3600);
    const minute = Math.floor(secondsInDay / 60) % 60, second = secondsInDay % 60;
    const meridiem = /AM\/PM/i.test(timeFormat);
    if (meridiem) hour = hour % 12 || 12;
    const hourToken = /hh/i.test(timeFormat) ? String(hour).padStart(2, '0') : String(hour);
    return hourToken + ':' + String(minute).padStart(2, '0') + (/:ss/i.test(timeFormat) ? ':' + String(second).padStart(2, '0') : '') + (meridiem ? (Math.floor(secondsInDay / 3600) < 12 ? ' AM' : ' PM') : '');
  };
  try {
    const totalSeconds = Math.round(value * 86400);
    if (!hasDateAndTime) return datePattern ? dateText(value) : timeText(totalSeconds);
    const dateResult = dateText(Math.floor(totalSeconds / 86400));
    return typeof dateResult === 'string' && dateResult.startsWith('#') ? dateResult : dateResult + ' ' + timeText(totalSeconds);
  } catch { return '#VALUE!'; }
}

const DATE_1904_SERIAL_RESULTS = new Set(['DATE', 'DATEVALUE', 'TODAY', 'NOW', 'EDATE', 'EOMONTH', 'WORKDAY', 'WORKDAY.INTL', 'COUPNCD', 'COUPPCD']);
const DATE_1904_INPUTS = new Map([
  ...['YEAR', 'MONTH', 'DAY', 'WEEKDAY', 'WEEKNUM', 'ISOWEEKNUM', 'EDATE', 'EOMONTH', 'HOUR', 'MINUTE', 'SECOND', 'WORKDAY', 'WORKDAY.INTL'].map(name => [name, [0]]),
  ...['DAYS', 'DATEDIF', 'DAYS360', 'YEARFRAC', 'NETWORKDAYS', 'NETWORKDAYS.INTL', 'ACCRINTM', 'DISC', 'INTRATE', 'RECEIVED', 'PRICEDISC', 'YIELDDISC', 'PRICE', 'YIELD', 'COUPDAYBS', 'COUPDAYS', 'COUPDAYSNC', 'COUPNCD', 'COUPNUM', 'COUPPCD'].map(name => [name, [0, 1]]),
  ['XNPV', []], ['XIRR', []]
]);

function call1904DateFunction(name, args, context) {
  const maximum = 2958465 - 1462;
  const shift = value => { const number = numeric(value); return typeof number !== 'number' ? number : number < 0 || number >= maximum + 1 ? '#NUM!' : number + 1462; };
  const mapDates = (value, coerce) => {
    if (isFormulaRange(value)) return { ...value, entries: value.entries.map(entry => ({ ...entry, value: mapDates(entry.value, coerce) })) };
    if (Array.isArray(value)) return value.map(item => mapDates(item, coerce));
    if (coerce) {
      const number = numeric(value);
      return typeof number === 'number' ? number + 1462 : value;
    }
    return typeof value === 'number' ? shift(value) : value;
  };
  const converted = [...args];
  for (const index of DATE_1904_INPUTS.get(name) || []) if (index < args.length) converted[index] = shift(args[index]);
  if (['WORKDAY', 'NETWORKDAYS', 'WORKDAY.INTL', 'NETWORKDAYS.INTL'].includes(name)) {
    const index = name.endsWith('.INTL') ? 3 : 2;
    if (index < args.length) converted[index] = mapDates(args[index], true);
  }
  if (name === 'XNPV' && args.length > 2) converted[2] = mapDates(args[2], false);
  if (name === 'XIRR' && args.length > 1) converted[1] = mapDates(args[1], false);
  // Nested bond/day-count calls now receive ordinary 1900 serials, so they must
  // not apply the epoch adjustment again. Other formula contexts remain intact.
  const next = { ...context, date1904: false, ...(Number.isFinite(context.todaySerial) ? { todaySerial: shift(context.todaySerial) } : {}) };
  const result = callFunction(name, converted, next);
  if (!DATE_1904_SERIAL_RESULTS.has(name) || typeof result !== 'number') return result;
  const serial = result - 1462;
  return Number.isFinite(serial) && serial >= 0 && serial < maximum + 1 ? serial : '#NUM!';
}

function callFunction(name, args, context = {}) {
  const sourceArgs = args;
  args = args.map(value => value?.[FORMULA_TEXT_LITERAL] || value?.[FORMULA_REFERENCE_VALUE] ? value.value : value);
  if (context.date1904 && (DATE_1904_INPUTS.has(name) || DATE_1904_SERIAL_RESULTS.has(name))) return call1904DateFunction(name, args, context);
  const sparseRangeSupported = ['SUM', 'PRODUCT', 'AVERAGE', 'AVERAGEA', 'MEDIAN', 'MIN', 'MINA', 'MAX', 'MAXA', 'LARGE', 'SMALL', 'MODE.SNGL', 'MODE.MULT', 'FREQUENCY', 'COUNT', 'COUNTA', 'COUNTBLANK', 'COUNTIF', 'SUMIF', 'AVERAGEIF', 'COUNTIFS', 'SUMIFS', 'AVERAGEIFS', 'MAXIFS', 'MINIFS', 'SUBTOTAL', 'ROWS', 'COLUMNS', 'INDEX', 'NETWORKDAYS', 'NETWORKDAYS.INTL', 'WORKDAY', 'WORKDAY.INTL', 'SUMSQ', 'DEVSQ', 'AVEDEV', 'GEOMEAN', 'HARMEAN', 'CORREL', 'PEARSON', 'COVARIANCE.P', 'COVARIANCE.S', 'SLOPE', 'INTERCEPT', 'RSQ', 'STEYX', 'FORECAST', 'FORECAST.LINEAR', 'GCD', 'LCM', 'MULTINOMIAL', 'SUMXMY2', 'SUMX2MY2', 'SUMX2PY2', 'TRANSPOSE', 'TOCOL', 'TOROW', 'WRAPROWS', 'WRAPCOLS', 'HSTACK', 'VSTACK', 'STDEV.S', 'STDEV.P', 'STDEVA', 'STDEVPA', 'VAR.S', 'VAR.P', 'VARA', 'VARPA', 'NPV', 'IRR', 'MIRR', 'XNPV', 'XIRR', 'FVSCHEDULE'];
  if (args.some(isFormulaRange) && !sparseRangeSupported.includes(name)) return '#VALUE!';
  const values = flatten(args);
  const textValue = value => value === null || value === undefined ? '' : value === true ? 'TRUE' : value === false ? 'FALSE' : String(value);
  if (['ISBLANK', 'ISNUMBER', 'ISTEXT', 'ISNONTEXT', 'ISLOGICAL'].includes(name)) {
    if (args.length !== 1) return '#VALUE!';
    const value = args[0];
    if (name === 'ISBLANK') return value === null || value === undefined;
    if (name === 'ISNUMBER') return typeof value === 'number' && Number.isFinite(value);
    if (name === 'ISTEXT') return typeof value === 'string';
    if (name === 'ISNONTEXT') return typeof value !== 'string';
    return typeof value === 'boolean';
  }
  if (name === 'ISEVEN' || name === 'ISODD') {
    if (args.length !== 1) return '#VALUE!';
    const value = numeric(args[0]);
    if (typeof value !== 'number') return value;
    const even = Math.abs(Math.trunc(value)) % 2 === 0;
    return name === 'ISEVEN' ? even : !even;
  }
  if (name === 'NA') return args.length ? '#VALUE!' : '#N/A';
  if (name === 'VALUE') {
    if (args.length !== 1) return '#VALUE!';
    if (isError(args[0])) return args[0];
    if (typeof args[0] === 'number') return Number.isFinite(args[0]) ? args[0] : '#VALUE!';
    const source = textValue(args[0]).trim();
    if (!source) return '#VALUE!';
    const match = /^([$€£¥₩])?([+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?(?:[eE][+-]?\d+)?)(%)?$/.exec(source);
    if (!match) return '#VALUE!';
    const value = Number(match[2].replaceAll(',', ''));
    return Number.isFinite(value) ? value / (match[3] ? 100 : 1) : '#VALUE!';
  }
  if (name === 'NUMBERVALUE') {
    if (args.length < 1 || args.length > 3) return '#VALUE!';
    if (isError(args[0])) return args[0];
    const decimalSeparator = args.length > 1 && args[1] !== null ? args[1] : '.';
    const groupSeparator = args.length > 2 && args[2] !== null ? args[2] : ',';
    if (typeof decimalSeparator !== 'string' || typeof groupSeparator !== 'string'
      || decimalSeparator.length > 1 || groupSeparator.length > 1
      || decimalSeparator === groupSeparator && decimalSeparator !== '') return '#VALUE!';
    let source = textValue(args[0]).replaceAll(/\s/gu, '');
    if (!source) return '#VALUE!';
    let percentCount = 0;
    while (source.endsWith('%')) { source = source.slice(0, -1); percentCount++; }
    if (source.includes('%')) return '#VALUE!';
    if (groupSeparator) source = source.replaceAll(groupSeparator, '');
    if (decimalSeparator && decimalSeparator !== '.') {
      if (source.split(decimalSeparator).length > 2) return '#VALUE!';
      source = source.replace(decimalSeparator, '.');
    } else if ((source.match(/\./g) || []).length > 1) return '#VALUE!';
    if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(source)) return '#VALUE!';
    const value = Number(source) / 100 ** percentCount;
    return Number.isFinite(value) ? value : '#VALUE!';
  }
  if (name === 'INDIRECT') {
    if (args.length < 1 || args.length > 2) return '#VALUE!';
    if (isError(args[0]) || isError(args[1])) return isError(args[0]) ? args[0] : args[1];
    const reference = textValue(args[0]).trim(), a1 = args.length < 2 || args[1] === null ? true : logical(args[1]);
    if (typeof a1 !== 'boolean') return '#VALUE!';
    if (!reference || reference.startsWith('[') || reference.length > 512) return '#REF!';
    const match = /^(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!(.+)$/.exec(reference);
    const sheetName = match ? (match[1] ? match[1].replaceAll("''", "'") : match[2]) : null;
    const address = match ? match[3] : reference;
    if (sheetName && context.getSheetIndex?.(sheetName) === -1) return '#REF!';
    let bounds;
    if (a1) {
      const range = /^(\$?[A-Z]{1,3}\$?[1-9]\d*)(?::(\$?[A-Z]{1,3}\$?[1-9]\d*))?$/i.exec(address);
      if (!range) return '#REF!';
      const parse = value => {
        const cell = /^\$?([A-Z]{1,3})\$?([1-9]\d*)$/i.exec(value);
        if (!cell) return null;
        const column = columnNumber(cell[1]) - 1, row = Number(cell[2]) - 1;
        return row < 1048576 && column < 16384 ? { row, column } : null;
      };
      const first = parse(range[1]), last = parse(range[2] || range[1]);
      if (!first || !last) return '#REF!';
      bounds = { top: Math.min(first.row, last.row), bottom: Math.max(first.row, last.row), left: Math.min(first.column, last.column), right: Math.max(first.column, last.column) };
    } else {
      const range = /^R(\[([+-]?\d+)\]|\d+)?C(\[([+-]?\d+)\]|\d+)?(?::R(\[([+-]?\d+)\]|\d+)?C(\[([+-]?\d+)\]|\d+)?)?$/i.exec(address);
      if (!range) return '#REF!';
      const coordinate = (token, offset, origin, limit) => {
        const value = token === undefined ? origin : token.startsWith('[') ? origin + Number(offset) : Number(token) - 1;
        return Number.isInteger(value) && value >= 0 && value < limit ? value : null;
      };
      const firstRow = coordinate(range[1], range[2], context.formulaRow || 0, 1048576), firstColumn = coordinate(range[3], range[4], context.formulaColumn || 0, 16384);
      const lastRow = coordinate(range[5], range[6], firstRow, 1048576), lastColumn = coordinate(range[7], range[8], firstColumn, 16384);
      if ([firstRow, firstColumn, lastRow, lastColumn].some(value => value === null)) return '#REF!';
      bounds = { top: Math.min(firstRow, lastRow), bottom: Math.max(firstRow, lastRow), left: Math.min(firstColumn, lastColumn), right: Math.max(firstColumn, lastColumn) };
    }
    const height = bounds.bottom - bounds.top + 1, width = bounds.right - bounds.left + 1;
    if (height * width > MAX_RANGE_CELLS) return '#REF!';
    const values = Array.from({ length: height }, (_, row) => Array.from({ length: width }, (_, column) => context.getCell?.(bounds.top + row, bounds.left + column, sheetName) ?? null));
    return height === 1 && width === 1 ? values[0][0] : values;
  }
  if (name === 'ROWS' || name === 'COLUMNS') {
    if (args.length !== 1) return '#VALUE!';
    const value = args[0];
    if (isError(value)) return value;
    if (isFormulaRange(value)) return name === 'ROWS' ? value.bounds.bottom - value.bounds.top + 1 : value.bounds.right - value.bounds.left + 1;
    if (!Array.isArray(value)) return 1;
    if (!value.length) return 0;
    return name === 'ROWS' ? value.length : (Array.isArray(value[0]) ? value[0].length : 1);
  }
  if (name === 'SEQUENCE') {
    if (args.length < 1 || args.length > 4 || args.some(isError)) return args.find(isError) || '#VALUE!';
    const dimension = (value, fallback) => value === undefined || value === null ? fallback : numeric(value);
    const rows = dimension(args[0], 1), columns = dimension(args[1], 1), start = dimension(args[2], 1), step = dimension(args[3], 1);
    if (![rows, columns].every(value => typeof value === 'number' && Number.isInteger(value) && value > 0) || typeof start !== 'number' || typeof step !== 'number') return '#VALUE!';
    if (!Number.isSafeInteger(rows * columns) || rows * columns > MAX_RANGE_CELLS) return '#NUM!';
    const result = [];
    for (let row = 0; row < rows; row++) {
      const line = [];
      for (let column = 0; column < columns; column++) {
        const value = start + (row * columns + column) * step;
        if (!Number.isFinite(value)) return '#NUM!';
        line.push(value);
      }
      result.push(line);
    }
    return result;
  }
  if (name === 'FILTER') {
    if (args.length < 2 || args.length > 3) return '#VALUE!';
    const source = matrixValue(args[0]), include = matrixValue(args[1]);
    if (!source || !include) return isError(args[0]) ? args[0] : isError(args[1]) ? args[1] : '#VALUE!';
    const height = source.length, width = source[0].length;
    const rowFilter = include.length === height && include[0].length === 1;
    const columnFilter = include.length === 1 && include[0].length === width;
    if (!rowFilter && !columnFilter) return '#VALUE!';
    const included = rowFilter ? include.map(row => logical(row[0])) : include[0].map(logical);
    const error = included.find(isError);
    if (error) return error;
    if (included.some(value => typeof value !== 'boolean')) return '#VALUE!';
    const selected = rowFilter
      ? source.filter((_, index) => included[index])
      : source.map(row => row.filter((_, index) => included[index]));
    if (selected.length && selected[0].length) return selected;
    return args.length > 2 ? args[2] : '#CALC!';
  }
  if (name === 'SORT') {
    if (args.length < 1 || args.length > 4) return '#VALUE!';
    if (args.some(isError)) return args.find(isError);
    const source = matrixValue(args[0]);
    if (!source) return '#VALUE!';
    const height = source.length, width = source[0].length;
    const sortIndex = args[1] === undefined || args[1] === null ? 1 : numeric(args[1]);
    const sortOrder = args[2] === undefined || args[2] === null ? 1 : numeric(args[2]);
    const byColumn = args[3] === undefined || args[3] === null ? false : logical(args[3]);
    const limit = byColumn === true ? height : width;
    if (!Number.isInteger(sortIndex) || sortIndex < 1 || sortIndex > limit || ![1, -1].includes(sortOrder) || typeof byColumn !== 'boolean') return '#VALUE!';
    const count = byColumn ? width : height;
    const indexes = Array.from({ length: count }, (_, index) => index);
    indexes.sort((left, right) => {
      const leftKey = byColumn ? source[sortIndex - 1][left] : source[left][sortIndex - 1];
      const rightKey = byColumn ? source[sortIndex - 1][right] : source[right][sortIndex - 1];
      return sortCompareValues(leftKey, rightKey, context.locale || 'en') * sortOrder || left - right;
    });
    return byColumn ? source.map(row => indexes.map(index => row[index])) : indexes.map(index => source[index]);
  }
  if (name === 'SORTBY') {
    if (args.length < 2 || args.length > 127) return '#VALUE!';
    if (args.some(isError)) return args.find(isError);
    const source = matrixValue(args[0]);
    if (!source) return '#VALUE!';
    const height = source.length, width = source[0].length, keys = [];
    for (let index = 1; index < args.length; index += 2) {
      const key = matrixValue(args[index]);
      if (!key) return '#VALUE!';
      const byRows = key.length === height && key[0].length === 1;
      const byColumns = key.length === 1 && key[0].length === width;
      if (!byRows && !byColumns || keys.length && keys[0].byRows !== byRows) return '#VALUE!';
      const orderValue = index + 1 < args.length && args[index + 1] !== null ? numeric(args[index + 1]) : 1;
      if (typeof orderValue !== 'number' || ![1, -1].includes(orderValue)) return '#VALUE!';
      keys.push({ values: byRows ? key.map(row => row[0]) : key[0], byRows, order: orderValue });
    }
    const byRows = keys[0].byRows, count = byRows ? height : width;
    const indexes = Array.from({ length: count }, (_, index) => index);
    indexes.sort((left, right) => {
      for (const key of keys) {
        const comparison = sortCompareValues(key.values[left], key.values[right], context.locale || 'en') * key.order;
        if (comparison) return comparison;
      }
      return left - right;
    });
    return byRows ? indexes.map(index => source[index]) : source.map(row => indexes.map(index => row[index]));
  }
  if (name === 'UNIQUE') {
    if (args.length < 1 || args.length > 3) return '#VALUE!';
    if (args.some(isError)) return args.find(isError);
    const source = matrixValue(args[0]);
    if (!source) return '#VALUE!';
    const byColumn = args[1] === undefined || args[1] === null ? false : logical(args[1]);
    const exactlyOnce = args[2] === undefined || args[2] === null ? false : logical(args[2]);
    if (typeof byColumn !== 'boolean' || typeof exactlyOnce !== 'boolean') return '#VALUE!';
    const vectors = byColumn
      ? Array.from({ length: source[0].length }, (_, column) => source.map(row => row[column]))
      : source;
    const keyFor = vector => JSON.stringify(vector.map(value => {
      if (value === null || value === undefined || value === '') return ['blank', ''];
      if (typeof value === 'string') return ['string', value.toLocaleLowerCase(context.locale || 'en')];
      return [typeof value, String(value)];
    }));
    const counts = new Map();
    for (const vector of vectors) {
      const key = keyFor(vector);
      counts.set(key, (counts.get(key) || 0) + 1);
    }
    const selected = [], seen = new Set();
    for (const vector of vectors) {
      const key = keyFor(vector);
      if (exactlyOnce ? counts.get(key) === 1 : !seen.has(key)) selected.push(vector);
      seen.add(key);
    }
    if (!selected.length) return '#CALC!';
    return byColumn
      ? source.map((_, row) => selected.map(vector => vector[row]))
      : selected;
  }
  if (name === 'TRANSPOSE') {
    if (args.length !== 1) return '#VALUE!';
    if (isError(args[0])) return args[0];
    const source = matrixValue(args[0]);
    if (!source || source.length * source[0].length > MAX_RANGE_CELLS) return source ? '#NUM!' : '#VALUE!';
    return source[0].map((_, column) => source.map(row => row[column]));
  }
  if (name === 'CHOOSECOLS' || name === 'CHOOSEROWS') {
    if (args.length < 2 || args.length > 254) return '#VALUE!';
    if (args.some(isError)) return args.find(isError);
    const source = matrixValue(args[0]);
    if (!source) return '#VALUE!';
    const size = name === 'CHOOSECOLS' ? source[0].length : source.length;
    const indexes = [];
    for (const argument of args.slice(1)) {
      const value = numeric(argument);
      if (typeof value !== 'number' || !Number.isInteger(value) || value === 0) return '#VALUE!';
      const index = value < 0 ? size + value : value - 1;
      if (index < 0 || index >= size) return '#VALUE!';
      indexes.push(index);
    }
    return name === 'CHOOSECOLS'
      ? source.map(row => indexes.map(index => row[index]))
      : indexes.map(index => source[index].slice());
  }
  if (name === 'TAKE' || name === 'DROP') {
    if (args.length < 2 || args.length > 3) return '#VALUE!';
    if (args.some(isError)) return args.find(isError);
    const source = matrixValue(args[0]);
    if (!source) return '#VALUE!';
    const dimension = value => {
      const result = numeric(value);
      return typeof result === 'number' ? Math.trunc(result) : null;
    };
    const rowCount = dimension(args[1]), columnCount = args.length > 2 ? dimension(args[2]) : null;
    if (rowCount === null || columnCount === null && args.length > 2) return '#VALUE!';
    if (name === 'TAKE' && (rowCount === 0 || columnCount === 0)) return '#VALUE!';
    const select = (size, count) => {
      if (count === null) return [0, size];
      if (name === 'TAKE') {
        const length = Math.abs(count);
        if (length > size) return null;
        return count > 0 ? [0, length] : [size - length, size];
      }
      if (Math.abs(count) >= size) return null;
      return count > 0 ? [count, size] : [0, size + count];
    };
    const rowBounds = select(source.length, rowCount), columnBounds = select(source[0].length, columnCount);
    if (!rowBounds || !columnBounds) return name === 'TAKE' ? '#VALUE!' : '#CALC!';
    const result = source.slice(rowBounds[0], rowBounds[1]).map(row => row.slice(columnBounds[0], columnBounds[1]));
    return result.length && result[0].length ? result : '#CALC!';
  }
  if (name === 'EXPAND') {
    if (args.length < 2 || args.length > 4) return '#VALUE!';
    if (isError(args[0]) || isError(args[1]) || args.length > 2 && isError(args[2])) return [args[0], args[1], args[2]].find(isError);
    const source = matrixValue(args[0]) || [[args[0]]];
    const rowsValue = numeric(args[1]);
    const columnsValue = args.length > 2 && args[2] !== null ? numeric(args[2]) : source[0].length;
    if (typeof rowsValue !== 'number' || typeof columnsValue !== 'number') return '#VALUE!';
    const rows = Math.trunc(rowsValue), columns = Math.trunc(columnsValue);
    if (rows < source.length || columns < source[0].length || rows < 1 || columns < 1) return '#VALUE!';
    if (!Number.isSafeInteger(rows * columns) || rows * columns > MAX_RANGE_CELLS) return '#NUM!';
    const pad = args.length > 3 ? args[3] : '#N/A';
    const result = Array.from({ length: rows }, (_, row) => Array.from({ length: columns }, (_, column) =>
      row < source.length && column < source[row].length ? source[row][column] : pad));
    return result;
  }
  if (name === 'TOCOL' || name === 'TOROW') {
    if (args.length < 1 || args.length > 3) return '#VALUE!';
    if (args.some(isError)) return args.find(isError);
    const source = matrixValue(args[0]);
    if (!source) return '#VALUE!';
    const ignoreValue = args.length > 1 && args[1] !== null ? numeric(args[1]) : 0;
    const scanByColumn = args.length > 2 && args[2] !== null ? logical(args[2]) : false;
    if (typeof ignoreValue !== 'number' || !Number.isInteger(ignoreValue) || ignoreValue < 0 || ignoreValue > 3 || typeof scanByColumn !== 'boolean') return '#VALUE!';
    const values = [];
    const append = value => {
      const blank = value === null || value === undefined || value === '';
      if ((ignoreValue & 1) && blank || (ignoreValue & 2) && isError(value)) return;
      values.push(value);
    };
    if (scanByColumn) for (let column = 0; column < source[0].length; column++) for (let row = 0; row < source.length; row++) append(source[row][column]);
    else for (const row of source) for (const value of row) append(value);
    if (!values.length) return '#CALC!';
    return name === 'TOROW' ? [values] : values.map(value => [value]);
  }
  if (name === 'WRAPROWS' || name === 'WRAPCOLS') {
    if (args.length < 2 || args.length > 3) return '#VALUE!';
    if (isError(args[0]) || isError(args[1])) return isError(args[0]) ? args[0] : args[1];
    const source = matrixValue(args[0]), countValue = numeric(args[1]);
    if (!source || typeof countValue !== 'number') return '#VALUE!';
    const wrapCount = Math.trunc(countValue);
    if (wrapCount < 1) return '#NUM!';
    const pad = args.length > 2 ? args[2] : '#N/A';
    const values = [];
    if (name === 'WRAPROWS') for (const row of source) values.push(...row);
    else for (let column = 0; column < source[0].length; column++) for (let row = 0; row < source.length; row++) values.push(source[row][column]);
    const outerSize = name === 'WRAPROWS' ? Math.ceil(values.length / wrapCount) : wrapCount;
    const innerSize = name === 'WRAPROWS' ? wrapCount : Math.ceil(values.length / wrapCount);
    if (outerSize * innerSize > MAX_RANGE_CELLS) return '#NUM!';
    const result = Array.from({ length: outerSize }, () => Array(innerSize).fill(pad));
    for (let index = 0; index < values.length; index++) {
      if (name === 'WRAPROWS') result[Math.floor(index / wrapCount)][index % wrapCount] = values[index];
      else result[index % wrapCount][Math.floor(index / wrapCount)] = values[index];
    }
    return result;
  }
  if (name === 'HSTACK' || name === 'VSTACK') {
    if (args.length < 1 || args.length > 254) return '#VALUE!';
    const arrays = [];
    for (const argument of args) {
      if (isError(argument)) return argument;
      arrays.push(matrixValue(argument) || [[argument]]);
    }
    const rows = name === 'HSTACK' ? Math.max(...arrays.map(array => array.length)) : arrays.reduce((sum, array) => sum + array.length, 0);
    const columns = name === 'HSTACK' ? arrays.reduce((sum, array) => sum + array[0].length, 0) : Math.max(...arrays.map(array => array[0].length));
    if (rows * columns > MAX_RANGE_CELLS) return '#NUM!';
    const result = Array.from({ length: rows }, () => Array(columns).fill('#N/A'));
    let rowOffset = 0, columnOffset = 0;
    for (const array of arrays) {
      for (let row = 0; row < array.length; row++) for (let column = 0; column < array[row].length; column++) {
        const targetRow = name === 'HSTACK' ? row : rowOffset + row;
        const targetColumn = name === 'HSTACK' ? columnOffset + column : column;
        result[targetRow][targetColumn] = array[row][column];
      }
      if (name === 'HSTACK') columnOffset += array[0].length;
      else rowOffset += array.length;
    }
    return result;
  }
  if (name === 'ADDRESS') {
    if (args.length < 2 || args.length > 5) return '#VALUE!';
    if (args.some(isError)) return args.find(isError);
    const row = numeric(args[0]), column = numeric(args[1]), absMode = args.length > 2 && args[2] !== null ? numeric(args[2]) : 1;
    const a1 = args.length > 3 && args[3] !== null ? logical(args[3]) : true;
    if (typeof row !== 'number' || !Number.isInteger(row) || row < 1 || row > 1048576 || typeof column !== 'number' || !Number.isInteger(column) || column < 1 || column > 16384 || typeof absMode !== 'number' || !Number.isInteger(absMode) || absMode < 1 || absMode > 4 || typeof a1 !== 'boolean') return '#VALUE!';
    let letters = '';
    for (let value = column; value > 0; value = Math.floor((value - 1) / 26)) letters = String.fromCharCode(65 + (value - 1) % 26) + letters;
    const absoluteRow = absMode === 1 || absMode === 2, absoluteColumn = absMode === 1 || absMode === 3;
    const reference = a1
      ? (absoluteColumn ? '$' : '') + letters + (absoluteRow ? '$' : '') + row
      : 'R' + (absoluteRow ? String(row) : '[' + row + ']') + 'C' + (absoluteColumn ? String(column) : '[' + column + ']');
    const sheetText = args.length > 4 && args[4] !== null ? String(args[4]) : '';
    const sheet = !sheetText ? '' : /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheetText) ? sheetText + '!' : "'" + sheetText.replaceAll("'", "''") + "'!";
    return sheet + reference;
  }
  if (name === 'TEXT') {
    if (args.length !== 2) return '#VALUE!';
    if (args.some(isError)) return args.find(isError);
    const format = textValue(args[1]).trim(), match = /^([$€£¥₩])?(#,##0|0)(?:\.([0#]{1,10}))?(%)?$/.exec(format);
    const value = numeric(args[0]);
    if (typeof value !== 'number') return value;
    if (!match) return formatTextDateTime(value, format, context.locale, context.date1904) ?? '#VALUE!';
    const percent = Boolean(match[4]), scaled = percent ? value * 100 : value;
    if (!Number.isFinite(scaled)) return '#VALUE!';
    const decimals = match[3] || '';
    const minimumFractionDigits = (decimals.match(/0/g) || []).length;
    const maximumFractionDigits = decimals.length;
    try {
      const number = new Intl.NumberFormat(context.locale || 'en-US', { useGrouping: match[2].startsWith('#,'), minimumFractionDigits, maximumFractionDigits }).format(scaled);
      return (match[1] || '') + number + (percent ? '%' : '');
    } catch { return '#VALUE!'; }
  }
  if (name === 'EXACT') {
    if (args.length !== 2) return '#VALUE!';
    if (args.some(isError)) return args.find(isError);
    return textValue(args[0]) === textValue(args[1]);
  }
  if (name === 'PROPER') {
    if (args.length !== 1) return '#VALUE!';
    if (isError(args[0])) return args[0];
    return textValue(args[0]).replace(/\p{L}+/gu, word => {
      const characters = Array.from(word);
      return characters[0].toLocaleUpperCase() + characters.slice(1).join('').toLocaleLowerCase();
    });
  }
  if (name === 'REPT') {
    if (args.length !== 2) return '#VALUE!';
    if (args.some(isError)) return args.find(isError);
    const count = numeric(args[1]);
    if (typeof count !== 'number' || count < 0) return '#VALUE!';
    const repetitions = Math.floor(count), source = textValue(args[0]);
    if (!source) return '';
    if (repetitions * source.length > 32767) return '#VALUE!';
    return source.repeat(repetitions);
  }
  if (name === 'CLEAN') {
    if (args.length !== 1) return '#VALUE!';
    if (isError(args[0])) return args[0];
    return textValue(args[0]).replace(/[\x00-\x1F]/g, '');
  }
  if (name === 'T') {
    if (args.length !== 1) return '#VALUE!';
    if (isError(args[0])) return args[0];
    return typeof args[0] === 'string' ? args[0] : '';
  }
  if (name === 'N') {
    if (args.length !== 1) return '#VALUE!';
    const value = args[0];
    if (isError(value)) return value;
    return typeof value === 'number' ? value : typeof value === 'boolean' ? Number(value) : 0;
  }
  if (name === 'UNICHAR') {
    if (args.length !== 1) return '#VALUE!';
    if (isError(args[0])) return args[0];
    const codePoint = numeric(args[0]);
    if (typeof codePoint !== 'number' || !Number.isInteger(codePoint) || codePoint < 1 || codePoint > 0x10ffff || codePoint >= 0xd800 && codePoint <= 0xdfff) return '#VALUE!';
    return String.fromCodePoint(codePoint);
  }
  if (name === 'UNICODE') {
    if (args.length !== 1) return '#VALUE!';
    if (isError(args[0])) return args[0];
    const characters = Array.from(textValue(args[0]));
    return characters.length ? characters[0].codePointAt(0) : '#VALUE!';
  }
  if (['ISERROR', 'ISERR', 'ISNA'].includes(name)) {
    if (args.length !== 1) return '#VALUE!';
    const value = args[0];
    if (name === 'ISNA') return value === '#N/A' || value === '#N/A!';
    if (name === 'ISERR') return isError(value) && value !== '#N/A' && value !== '#N/A!';
    return isError(value);
  }
  if (name === 'ERROR.TYPE') {
    if (args.length !== 1) return '#VALUE!';
    const error = args[0];
    const type = ERROR_TYPE_CODES.get(error);
    return type ?? '#N/A';
  }
  if (name === 'SIGN') {
    if (args.length !== 1) return '#VALUE!';
    if (isError(args[0])) return args[0];
    const value = numeric(args[0]);
    return typeof value === 'number' ? Math.sign(value) : '#VALUE!';
  }
  if (name === 'QUOTIENT') {
    if (args.length !== 2) return '#VALUE!';
    if (args.some(isError)) return args.find(isError);
    const numerator = numeric(args[0]), denominator = numeric(args[1]);
    if (typeof numerator !== 'number' || typeof denominator !== 'number') return '#VALUE!';
    return denominator === 0 ? '#DIV/0!' : Math.trunc(numerator / denominator);
  }
  if (name === 'EVEN' || name === 'ODD') {
    if (args.length !== 1) return '#VALUE!';
    if (isError(args[0])) return args[0];
    const value = numeric(args[0]);
    if (typeof value !== 'number') return '#VALUE!';
    if (name === 'EVEN' && value === 0) return 0;
    const magnitude = Math.ceil(Math.abs(value));
    const rounded = name === 'EVEN' ? magnitude + magnitude % 2 : magnitude % 2 ? magnitude : magnitude + 1;
    return Math.sign(value || 1) * rounded;
  }
  if (name === 'COUNTBLANK') {
    if (args.length !== 1) return '#VALUE!';
    if (isFormulaRange(args[0])) return args[0].cellCount - args[0].entries.filter(entry => entry.value !== null && entry.value !== undefined && entry.value !== '').length;
    return flatten([args[0]]).filter(value => value === null || value === undefined || value === '').length;
  }
  if (name === 'SUBTOTAL') {
    const error = args.find(isError);
    if (error) return error;
    if (args.length < 2 || isFormulaRange(args[0])) return '#VALUE!';
    const functionNumber = numeric(args[0]);
    if (typeof functionNumber !== 'number' || !Number.isInteger(functionNumber)) return '#VALUE!';
    const functions = { 1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT', 7: 'STDEV.S', 8: 'STDEV.P', 9: 'SUM', 10: 'VAR.S', 11: 'VAR.P' };
    const code = functionNumber > 100 ? functionNumber - 100 : functionNumber;
    const aggregate = functions[code];
    if (!aggregate || functionNumber > 111 || functionNumber < 1) return '#VALUE!';
    const included = [];
    for (const argument of args.slice(1)) {
      const range = isFormulaRange(argument) ? argument : argument?.[FORMULA_RANGE_METADATA];
      if (!range) { included.push(argument); continue; }
      for (const entry of range.entries) {
        if (context.isFilteredRow?.(range.sheetIndex, entry.row)) continue;
        if (functionNumber > 100 && context.isManuallyHiddenRow?.(range.sheetIndex, entry.row)) continue;
        if (entry.formula && /^\s*SUBTOTAL\s*\(/i.test(entry.formula)) continue;
        included.push(entry.value);
      }
    }
    return callFunction(aggregate, [included], context);
  }
  if (name === 'FIND' || name === 'SEARCH') {
    if (args.length < 2 || args.length > 3) return '#VALUE!';
    if (args.some(isError)) return args.find(isError);
    const needle = textValue(args[0]), haystack = textValue(args[1]), characters = Array.from(haystack);
    const start = args.length === 3 ? numeric(args[2]) : 1;
    if (typeof start !== 'number' || !Number.isInteger(start) || start < 1 || start > characters.length + 1) return '#VALUE!';
    const startOffset = characters.slice(0, start - 1).join('').length;
    if (name === 'FIND') {
      const index = haystack.indexOf(needle, startOffset);
      return index < 0 ? '#VALUE!' : Array.from(haystack.slice(0, index)).length + 1;
    }
    let pattern = '';
    for (let index = 0; index < needle.length; index++) {
      const character = needle[index];
      if (character === '~' && /[*?~]/.test(needle[index + 1] || '')) pattern += needle[++index].replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      else if (character === '*') pattern += '.*';
      else if (character === '?') pattern += '.';
      else pattern += character.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }
    const match = new RegExp(pattern, 'i').exec(haystack.slice(startOffset));
    return match ? start + Array.from(haystack.slice(startOffset, startOffset + match.index)).length : '#VALUE!';
  }
  if (name === 'TEXTSPLIT') {
    if (args.length < 2 || args.length > 6) return '#VALUE!';
    const inputError = flatten(args.slice(0, 5)).find(isError);
    if (inputError) return inputError;
    const source = textValue(args[0]);
    const delimiters = value => {
      if (value === null || value === undefined) return [];
      const array = matrixValue(value);
      return (array ? array.flat() : [value]).map(textValue);
    };
    const columnDelimiters = delimiters(args[1]), rowDelimiters = args.length > 2 ? delimiters(args[2]) : [];
    const ignoreEmpty = args.length > 3 && args[3] !== null ? logical(args[3]) : false;
    const matchMode = args.length > 4 && args[4] !== null ? numeric(args[4]) : 0;
    const pad = args.length > 5 && args[5] !== null ? args[5] : '#N/A';
    if (!columnDelimiters.length || columnDelimiters.length + rowDelimiters.length > 256 || columnDelimiters.some(delimiter => !delimiter) || rowDelimiters.some(delimiter => !delimiter) || typeof ignoreEmpty !== 'boolean' || typeof matchMode !== 'number' || ![0, 1].includes(matchMode)) return '#VALUE!';
    columnDelimiters.sort((left, right) => right.length - left.length);
    rowDelimiters.sort((left, right) => right.length - left.length);
    const split = (value, delimiterList) => {
      if (!delimiterList.length) return [value];
      const result = [], haystack = matchMode ? value.toLocaleLowerCase() : value;
      let start = 0;
      while (start <= value.length) {
        let index = -1, selected = null;
        for (const delimiter of delimiterList) {
          const needle = matchMode ? delimiter.toLocaleLowerCase() : delimiter;
          const candidate = haystack.indexOf(needle, start);
          if (candidate >= 0 && (index < 0 || candidate < index || candidate === index && delimiter.length > selected.length)) { index = candidate; selected = delimiter; }
        }
        if (index < 0) break;
        const part = value.slice(start, index);
        if (!ignoreEmpty || part !== '') result.push(part);
        start = index + selected.length;
      }
      const tail = value.slice(start);
      if (!ignoreEmpty || tail !== '') result.push(tail);
      return result;
    };
    if (source.length > 32767 || [...columnDelimiters, ...rowDelimiters].some(delimiter => delimiter.length > 32767)) return '#VALUE!';
    const rows = split(source, rowDelimiters).map(row => split(row, columnDelimiters));
    const width = Math.max(0, ...rows.map(row => row.length));
    if (!width) return '#CALC!';
    if (rows.length * width > MAX_RANGE_CELLS) return '#NUM!';
    return rows.map(row => Array.from({ length: width }, (_, index) => row[index] === undefined ? pad : row[index]));
  }
  if (name === 'TEXTBEFORE' || name === 'TEXTAFTER') {
    if (args.length < 2 || args.length > 6) return '#VALUE!';
    if (isError(args[0]) || isError(args[1])) return isError(args[0]) ? args[0] : args[1];
    for (const value of args.slice(2, 5)) if (isError(value)) return value;
    const text = Array.from(textValue(args[0])), delimiter = Array.from(textValue(args[1]));
    if (!delimiter.length) return '#VALUE!';
    const instance = args.length > 2 && args[2] !== null ? numeric(args[2]) : 1;
    const matchMode = args.length > 3 && args[3] !== null ? numeric(args[3]) : 0;
    const matchEnd = args.length > 4 && args[4] !== null ? numeric(args[4]) : 0;
    if (typeof instance !== 'number' || !Number.isInteger(instance) || instance === 0 || typeof matchMode !== 'number' || ![0, 1].includes(matchMode) || typeof matchEnd !== 'number' || ![0, 1].includes(matchEnd)) return '#VALUE!';
    const equal = (left, right) => matchMode ? left.toLocaleLowerCase() === right.toLocaleLowerCase() : left === right;
    const matches = [];
    for (let start = 0; start + delimiter.length <= text.length; start++) {
      let matchesDelimiter = true;
      for (let offset = 0; offset < delimiter.length; offset++) if (!equal(text[start + offset], delimiter[offset])) { matchesDelimiter = false; break; }
      if (matchesDelimiter) { matches.push(start); start += delimiter.length - 1; }
    }
    let selected = instance > 0 ? matches[instance - 1] : matches[matches.length + instance];
    if (selected === undefined && matchEnd) selected = text.length;
    if (selected === undefined) return args.length > 5 ? args[5] : '#N/A';
    const edge = selected + delimiter.length;
    const result = (name === 'TEXTBEFORE' ? text.slice(0, selected) : text.slice(edge)).join('');
    return result.length > 32767 ? '#VALUE!' : result;
  }
  if (name === 'SUBSTITUTE') {
    if (args.length < 3 || args.length > 4) return '#VALUE!';
    if (args.some(isError)) return args.find(isError);
    const source = textValue(args[0]), oldText = textValue(args[1]), newText = textValue(args[2]);
    if (!oldText) return source;
    if (args.length === 3) return source.split(oldText).join(newText);
    const instance = numeric(args[3]);
    if (typeof instance !== 'number' || !Number.isInteger(instance) || instance < 1) return '#VALUE!';
    let count = 0;
    return source.replaceAll(oldText, (match, offset) => ++count === instance ? newText : match);
  }
  if (name === 'REPLACE') {
    if (args.length !== 4) return '#VALUE!';
    if (args.some(isError)) return args.find(isError);
    const source = Array.from(textValue(args[0])), start = numeric(args[1]), count = numeric(args[2]);
    if (typeof start !== 'number' || typeof count !== 'number' || !Number.isInteger(start) || !Number.isInteger(count) || start < 1 || count < 0) return '#VALUE!';
    const beginning = source.slice(0, start - 1), ending = source.slice(start - 1 + count);
    return [...beginning, ...Array.from(textValue(args[3])), ...ending].join('');
  }
  if (name === 'SUMPRODUCT') {
    if (!args.length) return '#VALUE!';
    const arrays = args.map(value => Array.isArray(value) && value.every(Array.isArray) ? value : [[value]]);
    const rowCount = arrays[0].length, columnCount = arrays[0][0]?.length || 0;
    if (!rowCount || !columnCount || arrays.some(array => array.length !== rowCount || array.some(row => row.length !== columnCount))) return '#VALUE!';
    let total = 0;
    for (let row = 0; row < rowCount; row++) for (let column = 0; column < columnCount; column++) {
      let product = 1;
      for (const array of arrays) {
        const value = array[row][column];
        if (isError(value)) return value;
        const number = typeof value === 'number' ? value : typeof value === 'boolean' ? Number(value) : 0;
        product *= number;
      }
      total += product;
      if (!Number.isFinite(total)) return '#NUM!';
    }
    return total;
  }
  if (['LEFT', 'RIGHT', 'MID', 'LEN', 'LOWER', 'UPPER', 'TRIM'].includes(name)) {
    const expectedArgs = name === 'LEFT' || name === 'RIGHT' ? [1, 2] : name === 'MID' ? [3, 3] : [1, 1];
    if (args.length < expectedArgs[0] || args.length > expectedArgs[1]) return '#VALUE!';
    if (args.some(isError)) return args.find(isError);
    const characters = Array.from(textValue(args[0]));
    if (name === 'LEN') return characters.length;
    if (name === 'LOWER') return textValue(args[0]).toLowerCase();
    if (name === 'UPPER') return textValue(args[0]).toUpperCase();
    if (name === 'TRIM') return textValue(args[0]).replace(/^ +| +$/g, '').replace(/ {2,}/g, ' ');
    const rawStart = name === 'MID' ? numeric(args[1]) - 1 : 0;
    const rawCount = name === 'MID' ? numeric(args[2]) : (args.length === 1 ? 1 : numeric(args[1]));
    if (typeof rawStart !== 'number' || typeof rawCount !== 'number') return '#VALUE!';
    const start = Math.trunc(rawStart), count = Math.trunc(rawCount);
    if (count < 0 || (name === 'MID' && rawStart < 0)) return '#VALUE!';
    if (name === 'LEFT') return characters.slice(0, count).join('');
    if (name === 'RIGHT') return characters.slice(Math.max(0, characters.length - count)).join('');
    return characters.slice(start, start + count).join('');
  }
  if (name === 'CONCAT' || name === 'CONCATENATE') {
    if (!args.length || (name === 'CONCATENATE' && args.length > 255)) return '#VALUE!';
    if (values.some(isError)) return values.find(isError);
    return values.map(textValue).join('');
  }
  if (name === 'TEXTJOIN') {
    if (args.length < 3 || args.length > 254) return '#VALUE!';
    if (isError(args[0]) || isError(args[1])) return isError(args[0]) ? args[0] : args[1];
    const delimiter = textValue(args[0]), ignoreEmpty = logical(args[1]);
    if (isError(ignoreEmpty)) return ignoreEmpty;
    if (typeof ignoreEmpty !== 'boolean') return '#VALUE!';
    const parts = flatten(args.slice(2));
    if (parts.some(isError)) return parts.find(isError);
    const result = parts.filter(value => !ignoreEmpty || textValue(value) !== '').map(textValue).join(delimiter);
    return Array.from(result).length > 32767 ? '#VALUE!' : result;
  }
  if (name === 'COUNTIF' || name === 'SUMIF' || name === 'AVERAGEIF') {
    if ((name === 'COUNTIF' && args.length !== 2) || (name !== 'COUNTIF' && (args.length < 2 || args.length > 3))) return '#VALUE!';
    const criteriaRange = args[0], criteria = args[1], sumRange = args[2] ?? criteriaRange;
    if (isFormulaRange(criteriaRange) || isFormulaRange(sumRange)) {
      const criteriaBounds = isFormulaRange(criteriaRange) ? criteriaRange.bounds : criteriaRange?.[FORMULA_RANGE_METADATA]?.bounds;
      const sumMetadata = isFormulaRange(sumRange) ? sumRange : sumRange?.[FORMULA_RANGE_METADATA];
      const sumBounds = sumMetadata?.bounds;
      if (!criteriaBounds || !sumBounds) return '#VALUE!';
      const criteriaEntries = isFormulaRange(criteriaRange) ? criteriaRange.entries : criteriaRange[FORMULA_RANGE_METADATA].entries;
      const sumEntries = sumMetadata.entries;
      const criteriaValues = new Map(criteriaEntries.map(entry => [`${entry.row - criteriaBounds.top}:${entry.column - criteriaBounds.left}`, entry.value]));
      const sumValues = new Map(sumEntries.map(entry => [`${entry.row - sumBounds.top}:${entry.column - sumBounds.left}`, entry.value]));
      const criteriaHeight = criteriaBounds.bottom - criteriaBounds.top + 1;
      const criteriaWidth = criteriaBounds.right - criteriaBounds.left + 1;
      const sumHeight = sumBounds.bottom - sumBounds.top + 1;
      const sumWidth = sumBounds.right - sumBounds.left + 1;
      const positions = new Set(criteriaValues.keys());
      for (const entry of sumEntries) {
        const row = entry.row - sumBounds.top, column = entry.column - sumBounds.left;
        if (row < criteriaHeight && column < criteriaWidth) positions.add(`${row}:${column}`);
      }
      let count = 0, sum = 0;
      for (const position of positions) {
        const value = criteriaValues.get(position) ?? null;
        if (!criterionMatch(value, criteria)) continue;
        if (name === 'COUNTIF') { count++; continue; }
        const [row, column] = position.split(':').map(Number);
        const number = row < sumHeight && column < sumWidth ? sumValues.get(position) ?? null : null;
        if (isError(number)) return number;
        if (typeof number !== 'number' || !Number.isFinite(number)) continue;
        sum += number;
        if (!Number.isFinite(sum)) return '#NUM!';
        count++;
      }
      const criteriaCellCount = isFormulaRange(criteriaRange) ? criteriaRange.cellCount : criteriaHeight * criteriaWidth;
      if (name === 'COUNTIF') return count + (criterionMatch(null, criteria) ? criteriaCellCount - positions.size : 0);
      return name === 'SUMIF' ? sum : count ? sum / count : '#DIV/0!';
    }
    const pairs = criteriaPairs(criteriaRange, sumRange);
    if (!pairs) return '#VALUE!';
    let count = 0, sum = 0;
    for (const [value, number] of pairs) {
      if (!criterionMatch(value, criteria)) continue;
      if (name === 'COUNTIF') { count++; continue; }
      if (isError(number)) return number;
      if (typeof number !== 'number' || !Number.isFinite(number)) continue;
      sum += number;
      if (!Number.isFinite(sum)) return '#NUM!';
      count++;
    }
    if (name === 'COUNTIF') return count;
    return name === 'SUMIF' ? sum : count ? sum / count : '#DIV/0!';
  }
  if (name === 'COUNTIFS' || name === 'SUMIFS' || name === 'AVERAGEIFS' || name === 'MAXIFS' || name === 'MINIFS') {
    const hasSumRange = name !== 'COUNTIFS';
    const firstCriteria = hasSumRange ? 1 : 0;
    if (args.length < (hasSumRange ? 3 : 2) || (args.length - firstCriteria) % 2 !== 0) return '#VALUE!';
    const sparseSumRange = hasSumRange ? args[0] : null;
    const sparseConditions = [];
    for (let index = firstCriteria; index < args.length; index += 2) sparseConditions.push([args[index], args[index + 1]]);
    if (sparseConditions.some(([range]) => isFormulaRange(range)) || isFormulaRange(sparseSumRange)) {
      const sparseRanges = [...sparseConditions.map(([range]) => range), ...(hasSumRange ? [sparseSumRange] : [])];
      if (!sameSparseShape(sparseRanges)) return '#VALUE!';
      const conditionMaps = sparseConditions.map(([range, criterion]) => [sparseRangeIndex(range), criterion]);
      const sumValues = hasSumRange ? sparseRangeIndex(sparseSumRange) : null;
      const positions = sparsePositionUnion(sparseRanges);
      let count = 0, sum = 0, extreme = null;
      for (const position of positions) {
        if (!conditionMaps.every(([range, criterion]) => criterionMatch(range.has(position) ? range.get(position) : null, criterion))) continue;
        if (!hasSumRange) { count++; continue; }
        const value = sumValues.has(position) ? sumValues.get(position) : null;
        if (isError(value)) return value;
        if (name === 'MAXIFS' || name === 'MINIFS') {
          if (typeof value === 'number' && Number.isFinite(value)) extreme = extreme === null ? value : name === 'MAXIFS' ? Math.max(extreme, value) : Math.min(extreme, value);
          continue;
        }
        if (typeof value !== 'number' || !Number.isFinite(value)) continue;
        sum += value;
        if (!Number.isFinite(sum)) return '#NUM!';
        count++;
      }
      if (name === 'COUNTIFS') {
        if (conditionMaps.every(([, criterion]) => criterionMatch(null, criterion))) count += sparseRanges[0].cellCount - positions.size;
        return count;
      }
      if (name === 'MAXIFS' || name === 'MINIFS') return extreme ?? 0;
      return name === 'SUMIFS' ? sum : count ? sum / count : '#DIV/0!';
    }
    const sumRange = hasSumRange ? flatten([args[0]]) : null;
    const conditions = [];
    for (let index = firstCriteria; index < args.length; index += 2) conditions.push([flatten([args[index]]), args[index + 1]]);
    const length = conditions[0][0].length;
    if (conditions.some(([range]) => range.length !== length) || (sumRange && sumRange.length !== length)) return '#VALUE!';
    let count = 0, sum = 0, extreme = null;
    for (let offset = 0; offset < length; offset++) {
      if (!conditions.every(([range, criterion]) => criterionMatch(range[offset], criterion))) continue;
      if (!hasSumRange) { count++; continue; }
      const value = sumRange[offset];
      if (isError(value)) return value;
      if (name === 'MAXIFS' || name === 'MINIFS') {
        if (typeof value !== 'number' || !Number.isFinite(value)) continue;
        extreme = extreme === null ? value : name === 'MAXIFS' ? Math.max(extreme, value) : Math.min(extreme, value);
        continue;
      }
      if (typeof value !== 'number' || !Number.isFinite(value)) continue;
      sum += value;
      if (!Number.isFinite(sum)) return '#NUM!';
      count++;
    }
    if (name === 'COUNTIFS') return count;
    if (name === 'MAXIFS' || name === 'MINIFS') return extreme ?? 0;
    return name === 'SUMIFS' ? sum : count ? sum / count : '#DIV/0!';
  }
  if (name === 'XLOOKUP') {
    if (args.length < 3 || args.length > 6) return '#VALUE!';
    if (isError(args[1]) || isError(args[2])) return isError(args[1]) ? args[1] : args[2];
    const lookupMatrix = matrixValue(args[1]) ?? (!Array.isArray(args[1]) ? [[args[1]]] : null);
    const returnMatrix = matrixValue(args[2]) ?? (!Array.isArray(args[2]) ? [[args[2]]] : null);
    if (!lookupMatrix || !returnMatrix) return '#VALUE!';
    const lookupHorizontal = lookupMatrix.length === 1;
    if (!lookupHorizontal && lookupMatrix.some(row => row.length !== 1)) return '#VALUE!';
    const lookupValues = lookupHorizontal ? lookupMatrix[0] : lookupMatrix.map(row => row[0]);
    if (!lookupValues.length || (lookupHorizontal ? returnMatrix[0].length : returnMatrix.length) !== lookupValues.length) return '#VALUE!';
    const matchMode = args.length > 4 && args[4] !== null ? numeric(args[4]) : 0;
    const searchMode = args.length > 5 && args[5] !== null ? numeric(args[5]) : 1;
    if (typeof matchMode !== 'number' || ![-1, 0, 1, 2].includes(matchMode) || typeof searchMode !== 'number' || ![1, -1, 2, -2].includes(searchMode)) return '#VALUE!';
    if (isError(args[0])) return args[0];
    let found = -1;
    const matches = candidate => matchMode === 2
      ? criterionMatch(candidate, args[0])
      : lookupValue(candidate, false) === lookupValue(args[0], false);
    if ((searchMode === 2 || searchMode === -2) && matchMode !== 2) {
      const sought = comparable(args[0]), ascending = searchMode === 2;
      let low = 0, high = lookupValues.length - 1, candidateIndex = -1;
      while (low <= high) {
        const middle = (low + high) >> 1, candidate = lookupValues[middle];
        if (isError(candidate)) return candidate;
        if (matches(candidate)) { found = middle; break; }
        const current = comparable(candidate);
        if (typeof current !== typeof sought) return '#N/A';
        if (current < sought) {
          if (matchMode === -1) candidateIndex = middle;
          if (ascending) low = middle + 1;
          else high = middle - 1;
        } else {
          if (matchMode === 1) candidateIndex = middle;
          if (ascending) high = middle - 1;
          else low = middle + 1;
        }
      }
      if (found < 0) found = candidateIndex;
    } else {
      const order = Array.from({ length: lookupValues.length }, (_, index) => searchMode === -1 ? lookupValues.length - index - 1 : index);
      for (const index of order) if (matches(lookupValues[index])) { found = index; break; }
      if (found < 0 && matchMode !== 0 && matchMode !== 2) {
        const sought = comparable(args[0]);
        let best = null;
        for (const index of order) {
          const candidate = lookupValues[index];
          if (isError(candidate)) continue;
          const current = comparable(candidate);
          if (typeof current !== typeof sought) continue;
          const eligible = matchMode === -1 ? current < sought : current > sought;
          if (!eligible) continue;
          if (best === null || (matchMode === -1 ? current > best.value : current < best.value)) best = { value: current, index };
        }
        found = best?.index ?? -1;
      }
    }
    if (found < 0) return args.length > 3 && args[3] !== null ? args[3] : '#N/A';
    const normalizeBlank = value => value === null || value === undefined ? 0 : value;
    if (lookupHorizontal) {
      if (returnMatrix.length === 1) return normalizeBlank(returnMatrix[0][found]);
      return returnMatrix.map(row => [normalizeBlank(row[found])]);
    }
    const resultRow = returnMatrix[found].map(normalizeBlank);
    return resultRow.length === 1 ? resultRow[0] : [resultRow];
  }
  if (name === 'VLOOKUP' || name === 'HLOOKUP') {
    if (args.length < 3 || args.length > 4) return '#VALUE!';
    const table = matrix(args[1]), indexNumber = numeric(args[2]);
    if (!table || typeof indexNumber !== 'number' || !Number.isInteger(indexNumber) || indexNumber < 1) return '#VALUE!';
    const horizontal = name === 'HLOOKUP';
    const outerSize = horizontal ? table[0]?.length || 0 : table.length;
    const innerSize = horizontal ? table.length : Math.max(0, ...table.map(row => row.length));
    if (indexNumber > innerSize || !outerSize) return '#REF!';
    const approximate = args.length < 4 || (args[3] !== false && args[3] !== 0);
    const keys = Array.from({ length: outerSize }, (_, index) => horizontal ? table[0]?.[index] : table[index]?.[0]);
    const found = lookupIndex(keys, args[0], approximate, approximate);
    if (typeof found === 'string') return found;
    if (found < 0) return '#N/A';
    const result = horizontal ? table[indexNumber - 1]?.[found] : table[found]?.[indexNumber - 1];
    return result === null || result === undefined ? 0 : result;
  }
  if (name === 'LOOKUP') {
    if (args.length < 2 || args.length > 3) return '#VALUE!';
    if (isError(args[0]) || isError(args[1]) || isError(args[2])) return args.find(isError);
    const toVector = value => {
      const source = matrixValue(value) ?? (!Array.isArray(value) ? [[value]] : null);
      if (!source) return null;
      if (source.length === 1) return source[0];
      if (source.some(row => row.length !== 1)) return null;
      return source.map(row => row[0]);
    };
    const lookupVector = toVector(args[1]);
    if (!lookupVector?.length) return '#VALUE!';
    const resultVector = args.length === 3 ? toVector(args[2]) : lookupVector;
    if (!resultVector || resultVector.length !== lookupVector.length) return '#VALUE!';
    const found = lookupIndex(lookupVector, args[0], true, true);
    if (typeof found === 'string') return found;
    if (found < 0) return '#N/A';
    return resultVector[found] ?? 0;
  }
  if (name === 'MATCH' || name === 'XMATCH') {
    if (args.length < 2 || args.length > (name === 'XMATCH' ? 4 : 3)) return '#VALUE!';
    const range = flatten([args[1]]);
    if (name === 'XMATCH') {
      if (args.length > 4) return '#VALUE!';
      const mode = args.length < 3 ? 0 : numeric(args[2]), searchMode = args.length < 4 ? 1 : numeric(args[3]);
      if (typeof mode !== 'number' || ![-1, 0, 1, 2].includes(mode) || typeof searchMode !== 'number' || ![-2, -1, 1, 2].includes(searchMode)) return '#VALUE!';
      const sought = comparable(args[0]);
      const matches = value => mode === 2
        ? typeof args[0] === 'string' && criterionMatch(value, args[0])
        : comparable(value) === sought;
      if (searchMode === 1 || searchMode === -1) {
        const start = searchMode === 1 ? 0 : range.length - 1, end = searchMode === 1 ? range.length : -1, step = searchMode;
        let next = -1;
        for (let index = start; index !== end; index += step) {
          if (matches(range[index])) return index + 1;
          if (mode !== 0 && mode !== 2) {
            const candidate = comparable(range[index]);
            if (typeof candidate === typeof sought && (mode === -1 ? candidate < sought : candidate > sought)) {
              if (next < 0 || (mode === -1 ? comparable(range[next]) < candidate : comparable(range[next]) > candidate)) next = index;
            }
          }
        }
        return next < 0 ? '#N/A' : next + 1;
      }
      let low = 0, high = range.length - 1, candidateIndex = -1;
      while (low <= high) {
        const middle = (low + high) >> 1, candidate = comparable(range[middle]);
        if (matches(range[middle])) return middle + 1;
        if (typeof candidate !== typeof sought) return '#N/A';
        const ascending = searchMode === 2;
        if (candidate < sought) {
          if (mode === -1) candidateIndex = middle;
          if (ascending) low = middle + 1;
          else high = middle - 1;
        } else {
          if (mode === 1) candidateIndex = middle;
          if (ascending) high = middle - 1;
          else low = middle + 1;
        }
      }
      return candidateIndex < 0 ? '#N/A' : candidateIndex + 1;
    }
    const mode = args.length < 3 ? 1 : numeric(args[2]);
    if (typeof mode !== 'number' || ![-1, 0, 1].includes(mode)) return '#VALUE!';
    if (mode === 0) {
      const sought = lookupValue(args[0], false);
      const found = range.findIndex(value => typeof sought === 'string' && /[*?~]/.test(sought)
        ? criterionMatch(value, args[0])
        : lookupValue(value, false) === sought);
      return found < 0 ? '#N/A' : found + 1;
    }
    let found = -1;
    for (let index = 0; index < range.length; index++) {
      const candidate = comparable(range[index]), sought = comparable(args[0]);
      if (typeof candidate !== typeof sought) continue;
      if (candidate === sought) { found = index; continue; }
      if (mode === 1 && candidate < sought) found = index;
      if (mode === -1 && candidate > sought) found = index;
    }
    return found < 0 ? '#N/A' : found + 1;
  }
  if (name === 'INDEX') {
    if (args.length < 2 || args.length > 3) return '#VALUE!';
    const row = numeric(args[1]), column = args.length === 3 ? numeric(args[2]) : 1;
    if (typeof row !== 'number' || typeof column !== 'number' || !Number.isInteger(row) || !Number.isInteger(column) || row < 1 || column < 1) return '#VALUE!';
    if (isFormulaRange(args[0])) {
      const range = args[0], height = range.bounds.bottom - range.bounds.top + 1, width = range.bounds.right - range.bounds.left + 1;
      if (row > height || column > width) return '#REF!';
      const entry = range.entries.find(item => item.row === range.bounds.top + row - 1 && item.column === range.bounds.left + column - 1);
      return entry?.value ?? 0;
    }
    const range = matrix(args[0]);
    if (!range) return '#VALUE!';
    if (row > range.length || column > Math.max(0, ...range.map(line => line.length))) return '#REF!';
    return range[row - 1]?.[column - 1] ?? 0;
  }
  if (name === 'CHOOSE') {
    if (args.length < 2 || args.length > 255) return '#VALUE!';
    const index = numeric(args[0]);
    if (typeof index !== 'number') return index;
    const choice = Math.trunc(index);
    return choice < 1 || choice >= args.length ? '#VALUE!' : args[choice];
  }
  if (name === 'IFS') {
    if (!args.length || args.length % 2 || args.length > 254) return '#VALUE!';
    for (let index = 0; index < args.length; index += 2) {
      const condition = logical(args[index]);
      if (isError(condition)) return condition;
      if (condition) return args[index + 1];
    }
    return '#N/A';
  }
  if (name === 'SWITCH') {
    if (args.length < 3 || args.length > 255) return '#VALUE!';
    const expression = args[0];
    if (isError(expression)) return expression;
    const hasDefault = args.length % 2 === 0;
    const pairEnd = hasDefault ? args.length - 1 : args.length;
    for (let index = 1; index < pairEnd; index += 2) {
      const candidate = args[index];
      if (isError(candidate)) return candidate;
      const equal = typeof expression === 'string' && typeof candidate === 'string'
        ? expression.toLocaleLowerCase() === candidate.toLocaleLowerCase()
        : expression === candidate;
      if (equal) return args[index + 1];
    }
    return hasDefault ? args.at(-1) : '#N/A';
  }
  if (name === 'IF') {
    if (args.length < 2 || args.length > 3) return '#VALUE!';
    const condition = logical(args[0]);
    if (isError(condition)) return condition;
    return condition ? args[1] : (args.length === 3 ? args[2] : false);
  }
  if (name === 'IFERROR') {
    if (args.length !== 2) return '#VALUE!';
    return isError(args[0]) ? args[1] : args[0];
  }
  if (name === 'IFNA') {
    if (args.length !== 2) return '#VALUE!';
    return args[0] === '#N/A' || args[0] === '#N/A!' ? args[1] : args[0];
  }
  if (name === 'AND' || name === 'OR' || name === 'XOR') {
    if (!values.length) return '#VALUE!';
    let result = name === 'AND', trueCount = 0;
    for (const value of values) {
      const truth = logical(value);
      if (isError(truth)) return truth;
      if (name === 'AND') result = result && truth;
      else if (name === 'OR') result = result || truth;
      else if (truth) trueCount++;
    }
    return name === 'XOR' ? trueCount % 2 === 1 : result;
  }
  if (name === 'NOT') {
    if (args.length !== 1) return '#VALUE!';
    const truth = logical(args[0]);
    return isError(truth) ? truth : !truth;
  }
  if (name === 'RAND') {
    if (args.length) return '#VALUE!';
    const random = context.random?.() ?? Math.random();
    return Number.isFinite(random) && random >= 0 && random < 1 ? random : '#NUM!';
  }
  if (name === 'RANDBETWEEN') {
    if (args.length !== 2) return '#VALUE!';
    const bottom = numeric(args[0]), top = numeric(args[1]);
    if (typeof bottom !== 'number' || typeof top !== 'number') return typeof bottom !== 'number' ? bottom : top;
    const minimum = Math.trunc(bottom), maximum = Math.trunc(top);
    if (minimum > maximum) return '#NUM!';
    const random = context.random?.() ?? Math.random();
    if (!Number.isFinite(random) || random < 0 || random >= 1) return '#NUM!';
    return minimum + Math.floor(random * (maximum - minimum + 1));
  }
  if (name === 'RANDARRAY') {
    if (args.length > 5) return '#VALUE!';
    if (args.some(isError)) return args.find(isError);
    const dimension = (value, fallback) => value === undefined || value === null ? fallback : numeric(value);
    const rows = dimension(args[0], 1), columns = dimension(args[1], 1);
    const minimum = dimension(args[2], 0), maximum = dimension(args[3], 1);
    const whole = args[4] === undefined || args[4] === null ? false : args[4];
    if (![rows, columns].every(value => typeof value === 'number' && Number.isInteger(value) && value > 0)
      || typeof minimum !== 'number' || typeof maximum !== 'number' || typeof whole !== 'boolean') return '#VALUE!';
    if (!Number.isSafeInteger(rows * columns) || rows * columns > MAX_RANGE_CELLS) return '#NUM!';
    const bottom = Math.trunc(minimum), top = Math.trunc(maximum);
    if (whole && bottom > top || !whole && minimum > maximum) return '#NUM!';
    const result = [];
    for (let row = 0; row < rows; row++) {
      const line = [];
      for (let column = 0; column < columns; column++) {
        const random = context.random?.() ?? Math.random();
        if (!Number.isFinite(random) || random < 0 || random >= 1) return '#NUM!';
        const value = whole ? bottom + Math.floor(random * (top - bottom + 1)) : minimum + random * (maximum - minimum);
        if (!Number.isFinite(value)) return '#NUM!';
        line.push(value);
      }
      result.push(line);
    }
    return result;
  }
  if (name === 'NORM.S.DIST' || name === 'NORM.DIST') {
    const standard = name === 'NORM.S.DIST';
    if (args.length !== (standard ? 2 : 4)) return '#VALUE!';
    const value = numeric(args[0]), mean = standard ? 0 : numeric(args[1]), deviation = standard ? 1 : numeric(args[2]);
    const cumulative = standard ? args[1] : args[3];
    if (typeof value !== 'number' || typeof mean !== 'number' || typeof deviation !== 'number') return [value, mean, deviation].find(item => typeof item !== 'number');
    if (typeof cumulative !== 'boolean') return '#VALUE!';
    if (deviation <= 0) return '#NUM!';
    const difference = value - mean;
    const z = Number.isFinite(difference) ? difference / deviation : value / deviation - mean / deviation;
    if (Number.isNaN(z)) return '#NUM!';
    return cumulative ? standardNormalCdf(z) : Math.exp(-z * z / 2) / deviation / Math.sqrt(2 * Math.PI);
  }
  if (name === 'NORM.S.INV' || name === 'NORM.INV') {
    const standard = name === 'NORM.S.INV';
    if (args.length !== (standard ? 1 : 3)) return '#VALUE!';
    const probability = numeric(args[0]), mean = standard ? 0 : numeric(args[1]), deviation = standard ? 1 : numeric(args[2]);
    if (typeof probability !== 'number' || typeof mean !== 'number' || typeof deviation !== 'number') return [probability, mean, deviation].find(item => typeof item !== 'number');
    if (probability <= 0 || probability >= 1 || deviation <= 0) return '#NUM!';
    const value = standardNormalInverse(probability);
    if (typeof value !== 'number') return value;
    const result = mean + deviation * value;
    return Number.isFinite(result) ? result : '#NUM!';
  }
  if (name === 'BINOM.DIST') {
    if (args.length !== 4) return '#VALUE!';
    const successes = numeric(args[0]), trials = numeric(args[1]), probability = numeric(args[2]), cumulative = args[3];
    if ([successes, trials, probability].some(value => typeof value !== 'number')) return [successes, trials, probability].find(value => typeof value !== 'number');
    if (!Number.isInteger(successes) || !Number.isInteger(trials) || successes < 0 || trials < 0 || successes > trials || probability < 0 || probability > 1) return '#NUM!';
    if (typeof cumulative !== 'boolean') return '#VALUE!';
    if (trials > MAX_DISTRIBUTION_SIZE) return '#NUM!';
    const distribution = binomialDistribution(trials, probability);
    if (typeof distribution === 'string') return distribution;
    return cumulative ? compensatedSum(distribution.slice(0, successes + 1)) : distribution[successes];
  }
  if (name === 'POISSON.DIST') {
    if (args.length !== 3) return '#VALUE!';
    const value = numeric(args[0]), mean = numeric(args[1]), cumulative = args[2];
    if (typeof value !== 'number' || typeof mean !== 'number') return typeof value !== 'number' ? value : mean;
    if (!Number.isInteger(value) || value < 0 || mean < 0) return '#NUM!';
    if (typeof cumulative !== 'boolean') return '#VALUE!';
    const distribution = poissonDistribution(value, mean);
    if (typeof distribution === 'string') return distribution;
    return cumulative ? compensatedSum(distribution.slice(0, value + 1)) : distribution[value];
  }
  if (name === 'NPV') {
    if (args.length < 2) return '#VALUE!';
    const rate = numeric(args[0]);
    if (typeof rate !== 'number') return rate;
    if (rate <= -1) return '#NUM!';
    const cashFlows = financialCashFlows(args.slice(1));
    if (typeof cashFlows === 'string') return cashFlows;
    let result = 0;
    for (const { index, value } of cashFlows.entries) {
      result += value / (1 + rate) ** (index + 1);
      if (!Number.isFinite(result)) return '#NUM!';
    }
    return result;
  }
  if (name === 'RRI') {
    if (args.length !== 3) return '#VALUE!';
    const [periods, presentValue, futureValue] = args.map(numeric);
    for (const value of [periods, presentValue, futureValue]) if (typeof value !== 'number') return value;
    if (periods <= 0 || presentValue <= 0 || futureValue <= 0) return '#NUM!';
    const result = (futureValue / presentValue) ** (1 / periods) - 1;
    return Number.isFinite(result) ? result : '#NUM!';
  }
  if (name === 'PDURATION') {
    if (args.length !== 3) return '#VALUE!';
    const [rate, presentValue, futureValue] = args.map(numeric);
    for (const value of [rate, presentValue, futureValue]) if (typeof value !== 'number') return value;
    if (rate <= 0 || presentValue <= 0 || futureValue <= 0) return '#NUM!';
    const result = Math.log(futureValue / presentValue) / Math.log(1 + rate);
    return Number.isFinite(result) && result >= 0 ? result : '#NUM!';
  }
  if (name === 'ISPMT') {
    if (args.length !== 4) return '#VALUE!';
    const [rate, period, periods, presentValue] = args.map(numeric);
    for (const value of [rate, period, periods, presentValue]) if (typeof value !== 'number') return value;
    if (rate < 0 || periods <= 0 || !Number.isInteger(period) || period < 1 || period > periods) return '#NUM!';
    const result = presentValue * rate * (period / periods - 1);
    return Number.isFinite(result) ? result : '#NUM!';
  }
  if (name === 'DOLLARDE' || name === 'DOLLARFR') {
    if (args.length !== 2) return '#VALUE!';
    const value = numeric(args[0]), fractionValue = numeric(args[1]);
    if (typeof value !== 'number') return value;
    if (typeof fractionValue !== 'number') return fractionValue;
    const fraction = Math.trunc(fractionValue);
    if (fraction <= 0) return '#NUM!';
    const sign = Math.sign(value), absolute = Math.abs(value), whole = Math.floor(absolute);
    const fractional = absolute - whole;
    const result = name === 'DOLLARDE'
      ? whole + fractional * 100 / fraction
      : whole + fractional * fraction / 100;
    return Number.isFinite(result) ? sign * result : '#NUM!';
  }
  if (name === 'BASE') {
    if (args.length < 2 || args.length > 3) return '#VALUE!';
    const number = numeric(args[0]), radixValue = numeric(args[1]);
    const minLengthValue = args.length === 3 ? numeric(args[2]) : 0;
    for (const value of [number, radixValue, minLengthValue]) if (typeof value !== 'number') return value;
    if (number < 0 || number > Number.MAX_SAFE_INTEGER || radixValue < 2 || radixValue > 36 || minLengthValue < 0 || minLengthValue > 255) return '#NUM!';
    const integer = Math.trunc(number), radix = Math.trunc(radixValue), minLength = Math.trunc(minLengthValue);
    if (radix < 2 || radix > 36 || minLength < 0 || minLength > 255) return '#NUM!';
    return BigInt(integer).toString(radix).toUpperCase().padStart(minLength, '0');
  }
  if (name === 'DECIMAL') {
    if (args.length !== 2) return '#VALUE!';
    if (isError(args[0])) return args[0];
    const text = String(args[0]), radixValue = numeric(args[1]);
    if (typeof radixValue !== 'number') return radixValue;
    if (radixValue < 2 || radixValue > 36) return '#NUM!';
    const radix = Math.trunc(radixValue);
    if (!text.length || text.length > 255 || radix < 2 || radix > 36) return '#NUM!';
    const digits = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    let result = 0n;
    for (const character of text.toUpperCase()) {
      const digit = digits.indexOf(character);
      if (digit < 0 || digit >= radix) return '#NUM!';
      result = result * BigInt(radix) + BigInt(digit);
      if (result > BigInt(Number.MAX_SAFE_INTEGER)) return '#NUM!';
    }
    return Number(result);
  }
  if (name === 'ACCRINTM') {
    if (args.length < 4 || args.length > 5) return '#VALUE!';
    const issue = numeric(args[0]), settlement = numeric(args[1]);
    const rate = numeric(args[2]), par = numeric(args[3]);
    const basis = args.length === 5 ? numeric(args[4]) : 0;
    for (const value of [issue, settlement, rate, par, basis]) if (typeof value !== 'number') return value;
    if (issue < 0 || issue >= 2958466 || settlement < 0 || settlement >= 2958466 || issue >= settlement
      || rate < 0 || par <= 0 || !Number.isInteger(basis) || basis < 0 || basis > 4) return '#NUM!';
    const yearFraction = callFunction('YEARFRAC', [Math.floor(issue), Math.floor(settlement), basis]);
    if (typeof yearFraction !== 'number') return yearFraction;
    const result = par * rate * yearFraction;
    return Number.isFinite(result) ? result : '#NUM!';
  }
  if (['DISC', 'INTRATE', 'RECEIVED'].includes(name)) {
    if (args.length < 4 || args.length > 5) return '#VALUE!';
    const settlement = numeric(args[0]), maturity = numeric(args[1]);
    const first = numeric(args[2]), second = numeric(args[3]);
    const basis = args.length === 5 ? numeric(args[4]) : 0;
    for (const value of [settlement, maturity, first, second, basis]) if (typeof value !== 'number') return value;
    if (settlement < 0 || settlement >= 2958466 || maturity < 0 || maturity >= 2958466
      || maturity <= settlement || !Number.isInteger(basis) || basis < 0 || basis > 4) return '#NUM!';
    if (name === 'RECEIVED' ? first <= 0 || second < 0 : first <= 0 || second <= 0) return '#NUM!';
    const yearFraction = callFunction('YEARFRAC', [Math.floor(settlement), Math.floor(maturity), basis]);
    if (typeof yearFraction !== 'number' || yearFraction <= 0) return typeof yearFraction === 'number' ? '#NUM!' : yearFraction;
    let result;
    if (name === 'DISC') result = (second - first) / second / yearFraction;
    else if (name === 'INTRATE') result = (second - first) / first / yearFraction;
    else {
      const denominator = 1 - second * yearFraction;
      if (denominator <= 0) return '#NUM!';
      result = first / denominator;
    }
    return Number.isFinite(result) ? result : '#NUM!';
  }
  if (name === 'PRICEDISC' || name === 'YIELDDISC') {
    if (args.length < 4 || args.length > 5) return '#VALUE!';
    const settlement = numeric(args[0]), maturity = numeric(args[1]);
    const priceOrDiscount = numeric(args[2]), redemption = numeric(args[3]);
    const basis = args.length === 5 ? numeric(args[4]) : 0;
    for (const value of [settlement, maturity, priceOrDiscount, redemption, basis]) if (typeof value !== 'number') return value;
    if (settlement < 0 || settlement >= 2958466 || maturity < 0 || maturity >= 2958466 || maturity <= settlement
      || redemption <= 0 || !Number.isInteger(basis) || basis < 0 || basis > 4
      || name === 'PRICEDISC' && priceOrDiscount < 0 || name === 'YIELDDISC' && priceOrDiscount <= 0) return '#NUM!';
    const yearFraction = callFunction('YEARFRAC', [Math.floor(settlement), Math.floor(maturity), basis]);
    if (typeof yearFraction !== 'number' || yearFraction <= 0) return typeof yearFraction === 'number' ? '#NUM!' : yearFraction;
    const result = name === 'PRICEDISC'
      ? redemption * (1 - priceOrDiscount * yearFraction)
      : (redemption - priceOrDiscount) / priceOrDiscount / yearFraction;
    return Number.isFinite(result) && (name !== 'PRICEDISC' || result > 0) ? result : '#NUM!';
  }
  if (name === 'PRICE' || name === 'YIELD') {
    if (args.length < 6 || args.length > 7) return '#VALUE!';
    const settlement = numeric(args[0]), maturity = numeric(args[1]), rate = numeric(args[2]);
    const quote = numeric(args[3]), redemption = numeric(args[4]), frequency = numeric(args[5]);
    const basis = args.length === 7 ? numeric(args[6]) : 0;
    for (const value of [settlement, maturity, rate, quote, redemption, frequency, basis]) if (typeof value !== 'number') return value;
    if (settlement < 0 || settlement >= 2958466 || maturity < 0 || maturity >= 2958466 || settlement >= maturity
      || rate < 0 || redemption <= 0 || ![1, 2, 4].includes(frequency) || !Number.isInteger(basis) || basis < 0 || basis > 4
      || name === 'PRICE' && quote <= -frequency || name === 'YIELD' && quote <= 0) return '#NUM!';
    const priceAt = yieldRate => couponBondPrice(settlement, maturity, rate, yieldRate, redemption, frequency, basis);
    if (name === 'PRICE') return priceAt(quote);
    const lowLimit = -frequency, low = lowLimit + 1e-12;
    let lower = low, upper = 0.1, lowerPrice = priceAt(lower), upperPrice = priceAt(upper);
    if (typeof lowerPrice !== 'number' && lowerPrice !== '#NUM!') return lowerPrice;
    if (lowerPrice !== '#NUM!' && lowerPrice <= quote) return '#NUM!';
    for (let iteration = 0; upperPrice !== '#NUM!' && upperPrice > quote && iteration < 80; iteration++) {
      upper = upper * 2 + 0.1;
      upperPrice = priceAt(upper);
    }
    if (upperPrice !== '#NUM!' && upperPrice > quote) return '#NUM!';
    for (let iteration = 0; iteration < 200; iteration++) {
      const middle = (lower + upper) / 2, middlePrice = priceAt(middle);
      if (middlePrice === '#NUM!' || middlePrice < quote) upper = middle;
      else lower = middle;
      if (upper - lower <= 1e-12 * Math.max(1, Math.abs(middle))) return (lower + upper) / 2;
    }
    return (lower + upper) / 2;
  }
  if (name === 'IRR') {
    if (args.length < 1 || args.length > 2) return '#VALUE!';
    const cashFlows = financialCashFlows([args[0]]);
    if (typeof cashFlows === 'string') return cashFlows;
    if (cashFlows.length < 2 || !cashFlows.entries.some(entry => entry.value > 0) || !cashFlows.entries.some(entry => entry.value < 0)) return '#NUM!';
    const guess = args.length === 2 ? numeric(args[1]) : 0.1;
    if (typeof guess !== 'number') return guess;
    if (guess <= -1) return '#NUM!';
    const scale = cashFlows.entries.reduce((maximum, entry) => Math.max(maximum, Math.abs(entry.value)), 1), tolerance = scale * 1e-10;
    let rate = guess;
    for (let iteration = 0; iteration < 100; iteration++) {
      const base = 1 + rate;
      let balance = 0, derivative = 0;
      for (const { index: period, value } of cashFlows.entries) {
        const factor = base ** period;
        if (!Number.isFinite(factor) || factor === 0) return '#NUM!';
        balance += value / factor;
        if (period) derivative -= period * value / (factor * base);
      }
      if (!Number.isFinite(balance) || !Number.isFinite(derivative) || derivative === 0) return '#NUM!';
      if (Math.abs(balance) <= tolerance) return rate;
      const next = rate - balance / derivative;
      if (!Number.isFinite(next) || next <= -1) return '#NUM!';
      if (Math.abs(next - rate) <= 1e-12 * Math.max(1, Math.abs(rate))) return Math.abs(balance) <= tolerance ? next : '#NUM!';
      rate = next;
    }
    return '#NUM!';
  }
  if (name === 'MIRR') {
    if (args.length !== 3) return '#VALUE!';
    const cashFlows = financialCashFlows([args[0]]);
    if (typeof cashFlows === 'string') return cashFlows;
    const financeRate = numeric(args[1]), reinvestRate = numeric(args[2]);
    if (typeof financeRate !== 'number' || typeof reinvestRate !== 'number') return typeof financeRate !== 'number' ? financeRate : reinvestRate;
    if (financeRate <= -1 || reinvestRate <= -1) return '#NUM!';
    if (cashFlows.length < 2) return '#NUM!';
    let presentValue = 0, futureValue = 0, hasNegative = false, hasPositive = false;
    for (const { index, value } of cashFlows.entries) {
      if (value < 0) {
        hasNegative = true;
        presentValue += value / (1 + financeRate) ** index;
      } else {
        hasPositive = true;
        futureValue += value * (1 + reinvestRate) ** (cashFlows.length - 1 - index);
      }
      if (!Number.isFinite(presentValue) || !Number.isFinite(futureValue)) return '#NUM!';
    }
    if (!hasNegative || !hasPositive) return '#DIV/0!';
    const result = (futureValue / -presentValue) ** (1 / (cashFlows.length - 1)) - 1;
    return Number.isFinite(result) ? result : '#NUM!';
  }
  if (name === 'XNPV') {
    if (args.length !== 3) return '#VALUE!';
    const rate = numeric(args[0]);
    if (typeof rate !== 'number') return rate;
    if (rate <= -1) return '#NUM!';
    const pairs = financialDatePairs(args[1], args[2]);
    if (typeof pairs === 'string') return pairs;
    let result = 0;
    for (const item of pairs.entries) {
      result += item.value / (1 + rate) ** ((item.date - pairs.startDate) / 365);
      if (!Number.isFinite(result)) return '#NUM!';
    }
    return result;
  }
  if (name === 'XIRR') {
    if (args.length < 2 || args.length > 3) return '#VALUE!';
    const pairs = financialDatePairs(args[0], args[1]);
    if (typeof pairs === 'string') return pairs;
    if (pairs.length < 2 || !pairs.entries.some(item => item.value > 0) || !pairs.entries.some(item => item.value < 0)) return '#NUM!';
    const guess = args.length === 3 ? numeric(args[2]) : 0.1;
    if (typeof guess !== 'number') return guess;
    if (guess <= -1) return '#NUM!';
    const scale = pairs.entries.reduce((maximum, item) => Math.max(maximum, Math.abs(item.value)), 1), tolerance = scale * 1e-10;
    let rate = guess;
    for (let iteration = 0; iteration < 100; iteration++) {
      const base = 1 + rate;
      let balance = 0, derivative = 0;
      for (const item of pairs.entries) {
        const years = (item.date - pairs.startDate) / 365, factor = base ** years;
        if (!Number.isFinite(factor) || factor === 0) return '#NUM!';
        balance += item.value / factor;
        derivative -= years * item.value / (factor * base);
      }
      if (!Number.isFinite(balance) || !Number.isFinite(derivative) || derivative === 0) return '#NUM!';
      if (Math.abs(balance) <= tolerance) return rate;
      const next = rate - balance / derivative;
      if (!Number.isFinite(next) || next <= -1) return '#NUM!';
      if (Math.abs(next - rate) <= 1e-12 * Math.max(1, Math.abs(rate))) return Math.abs(balance) <= tolerance ? next : '#NUM!';
      rate = next;
    }
    return '#NUM!';
  }
  if (name === 'TODAY' || name === 'NOW') {
    if (args.length) return '#VALUE!';
    if (Number.isFinite(context.todaySerial)) return name === 'TODAY' ? Math.floor(context.todaySerial) : context.todaySerial;
    const now = context.now instanceof Date ? context.now : new Date();
    const date = utcDate(now.getFullYear(), now.getMonth() + 1, now.getDate());
    const serial = dateToSerial(date);
    if (name === 'TODAY') return serial;
    return serial + (now.getHours() * 3600 + now.getMinutes() * 60 + now.getSeconds()) / 86400;
  }
  if (name === 'DATE') {
    if (args.length !== 3) return '#VALUE!';
    const year = numeric(args[0]), month = numeric(args[1]), day = numeric(args[2]);
    if ([year, month, day].some(value => typeof value !== 'number')) return [year, month, day].find(value => typeof value === 'string');
    if (year < 0 || year > 9999) return '#NUM!';
    const normalizedYear = year < 1900 ? year + 1900 : year;
    if (Math.trunc(normalizedYear) === 1900 && Math.trunc(month) === 2 && Math.trunc(day) === 29) return 60;
    return dateToSerial(utcDate(Math.trunc(normalizedYear), Math.trunc(month), Math.trunc(day)));
  }
  if (name === 'DATEVALUE') {
    if (args.length !== 1) return '#VALUE!';
    if (isError(args[0])) return args[0];
    if (typeof args[0] !== 'string') return '#VALUE!';
    const date = parseDateValue(args[0]);
    if (!date) return '#VALUE!';
    return dateToSerial(utcDate(date.year, date.month, date.day));
  }
  if (['YEAR', 'MONTH', 'DAY'].includes(name)) {
    if (args.length !== 1) return '#VALUE!';
    const serial = numeric(args[0]);
    if (typeof serial !== 'number') return serial;
    const parts = dateParts(serial);
    if (typeof parts === 'string') return parts;
    return name === 'YEAR' ? parts.year : name === 'MONTH' ? parts.month : parts.day;
  }
  if (name === 'WEEKDAY') {
    if (args.length < 1 || args.length > 2) return '#VALUE!';
    const serial = numeric(args[0]);
    if (typeof serial !== 'number') return serial;
    if (serial < 0 || serial >= 2958466) return '#NUM!';
    const returnType = args.length === 1 ? 1 : numeric(args[1]);
    if (typeof returnType !== 'number') return returnType;
    if (![1, 2, 3, 11, 12, 13, 14, 15, 16, 17].includes(returnType)) return '#NUM!';
    const whole = Math.floor(serial), sundayZero = ((whole - (whole > 60 ? 1 : 0)) % 7 + 7) % 7;
    if (returnType === 1) return sundayZero + 1;
    if (returnType === 3) return (sundayZero + 6) % 7;
    const firstDay = returnType === 2 ? 1 : (returnType - 10) % 7;
    return (sundayZero - firstDay + 7) % 7 + 1;
  }
  if (name === 'WEEKNUM' || name === 'ISOWEEKNUM') {
    if (name === 'ISOWEEKNUM' ? args.length !== 1 : args.length < 1 || args.length > 2) return '#VALUE!';
    const serial = numeric(args[0]);
    if (typeof serial !== 'number') return serial;
    if (serial < 1 || serial >= 2958466) return '#NUM!';
    const whole = Math.floor(serial), returnType = name === 'ISOWEEKNUM' ? 21 : args.length === 1 ? 1 : numeric(args[1]);
    if (typeof returnType !== 'number') return returnType;
    if (returnType === 21) {
      const weekday = ((whole - (whole > 60 ? 1 : 0)) % 7 + 7) % 7;
      const mondayStart = whole - (weekday + 6) % 7;
      const thursday = dateParts(mondayStart + 3);
      if (typeof thursday === 'string') return thursday;
      const januaryFourth = dateToSerial(utcDate(thursday.year, 1, 4));
      const janWeekday = ((januaryFourth - (januaryFourth > 60 ? 1 : 0)) % 7 + 7) % 7;
      const firstMonday = januaryFourth - (janWeekday + 6) % 7;
      return Math.floor((mondayStart - firstMonday) / 7) + 1;
    }
    const firstDay = returnType === 2 ? 1 : [11, 12, 13, 14, 15, 16, 17].includes(returnType) ? (returnType - 10) % 7 : returnType === 1 ? 0 : null;
    if (firstDay === null) return '#NUM!';
    const weekdayIndex = value => ((value - (value > 60 ? 1 : 0)) % 7 + 7) % 7;
    const startOfWeek = value => value - (weekdayIndex(value) - firstDay + 7) % 7;
    const date = dateParts(whole);
    const januaryFirst = dateToSerial(utcDate(date.year, 1, 1));
    return Math.floor((startOfWeek(whole) - startOfWeek(januaryFirst)) / 7) + 1;
  }
  if (name === 'EDATE' || name === 'EOMONTH') {
    if (args.length !== 2) return '#VALUE!';
    const serial = numeric(args[0]), months = numeric(args[1]);
    if (typeof serial !== 'number') return serial;
    if (typeof months !== 'number') return months;
    if (serial < 0 || serial >= 2958466) return '#NUM!';
    const monthOffset = Math.trunc(months), source = dateParts(serial);
    if (typeof source === 'string') return source;
    if (serial === 60 && monthOffset === 0 && name === 'EDATE') return 60;
    const absoluteMonth = source.year * 12 + source.month - 1 + monthOffset;
    const year = Math.floor(absoluteMonth / 12), month = ((absoluteMonth % 12) + 12) % 12 + 1;
    if (year < 1900 || year > 9999) return '#NUM!';
    const lastDay = new Date(utcDate(year, month + 1, 0)).getUTCDate();
    if (name === 'EOMONTH') {
      if (year === 1900 && month === 2) return 60;
      return dateToSerial(utcDate(year, month, lastDay));
    }
    const day = Math.min(source.day, lastDay);
    return dateToSerial(utcDate(year, month, day));
  }
  if (name === 'DAYS') {
    if (args.length !== 2) return '#VALUE!';
    const end = numeric(args[0]), start = numeric(args[1]);
    if (typeof end !== 'number') return end;
    if (typeof start !== 'number') return start;
    if (start < 0 || start >= 2958466 || end < 0 || end >= 2958466) return '#NUM!';
    return Math.floor(end) - Math.floor(start);
  }
  if (name === 'DATEDIF') {
    if (args.length !== 3) return '#VALUE!';
    const startSerial = numeric(args[0]), endSerial = numeric(args[1]);
    if (typeof startSerial !== 'number') return startSerial;
    if (typeof endSerial !== 'number') return endSerial;
    if (isError(args[2])) return args[2];
    if (typeof args[2] !== 'string') return '#VALUE!';
    if ([startSerial, endSerial].some(value => value < 0 || value >= 2958466) || endSerial < startSerial) return '#NUM!';
    const unit = args[2].toUpperCase(), start = dateParts(Math.floor(startSerial)), end = dateParts(Math.floor(endSerial));
    if (typeof start === 'string') return start;
    if (typeof end === 'string') return end;
    const yearDifference = end.year - start.year;
    const monthDifference = yearDifference * 12 + end.month - start.month;
    const completeMonths = monthDifference - (end.day < start.day ? 1 : 0);
    const completeYears = yearDifference - (end.month < start.month || end.month === start.month && end.day < start.day ? 1 : 0);
    if (unit === 'Y') return completeYears;
    if (unit === 'M') return completeMonths;
    if (unit === 'D') return Math.floor(endSerial) - Math.floor(startSerial);
    if (unit === 'YM') return completeMonths - completeYears * 12;
    if (unit === 'MD') {
      const dayDifference = end.day - start.day;
      if (dayDifference >= 0) return dayDifference;
      const previousMonthLastDay = new Date(utcDate(end.year, end.month, 0)).getUTCDate();
      return dayDifference + previousMonthLastDay;
    }
    if (unit === 'YD') return Math.floor(endSerial) - dateToSerial(utcDate(start.year + completeYears, start.month, start.day));
    return '#NUM!';
  }
  if (name === 'NETWORKDAYS' || name === 'NETWORKDAYS.INTL') {
    const intl = name.endsWith('.INTL');
    if (args.length < 2 || args.length > (intl ? 4 : 3)) return '#VALUE!';
    const startValue = numeric(args[0]), endValue = numeric(args[1]);
    if (typeof startValue !== 'number') return startValue;
    if (typeof endValue !== 'number') return endValue;
    if ([startValue, endValue].some(value => value < 0 || value >= 2958466)) return '#NUM!';
    const holidaysIndex = intl ? 3 : 2, weekends = weekendDays(intl ? args[2] : undefined);
    if (!weekends) return '#NUM!';
    const first = Math.floor(Math.min(startValue, endValue)), last = Math.floor(Math.max(startValue, endValue));
    const totalDays = last - first + 1, fullWeeks = Math.floor(totalDays / 7);
    let count = fullWeeks * (7 - weekends.size);
    for (let offset = 0; offset < totalDays % 7; offset++) if (isWorkingDay(first + offset, weekends)) count++;
    if (args[holidaysIndex] !== undefined) {
      const holidays = flatten([args[holidaysIndex]]), unique = new Set();
      for (const holiday of holidays) {
        if (isError(holiday)) return holiday;
        const serial = numeric(holiday);
        if (typeof serial !== 'number') continue;
        const day = Math.floor(serial);
        if (day < first || day > last || unique.has(day)) continue;
        if (isWorkingDay(day, weekends)) { count--; unique.add(day); }
      }
    }
    return endValue < startValue ? -count : count;
  }
  if (name === 'WORKDAY' || name === 'WORKDAY.INTL') {
    const intl = name.endsWith('.INTL');
    if (args.length < 2 || args.length > (intl ? 4 : 3)) return '#VALUE!';
    const startValue = numeric(args[0]), daysValue = numeric(args[1]);
    if (typeof startValue !== 'number') return startValue;
    if (typeof daysValue !== 'number') return daysValue;
    if (startValue < 0 || startValue >= 2958466) return '#NUM!';
    const holidaysIndex = intl ? 3 : 2, weekends = weekendDays(intl ? args[2] : undefined);
    if (!weekends) return '#NUM!';
    const holidays = new Set();
    if (args[holidaysIndex] !== undefined) for (const holiday of flatten([args[holidaysIndex]])) {
      if (isError(holiday)) return holiday;
      const serial = numeric(holiday);
      if (typeof serial === 'number' && serial >= 0 && serial < 2958466) holidays.add(Math.floor(serial));
    }
    let day = Math.floor(startValue), remaining = Math.abs(Math.trunc(daysValue)), direction = Math.sign(daysValue);
    while (remaining > 0) {
      day += direction;
      if (day < 0 || day >= 2958466) return '#NUM!';
      if (isWorkingDay(day, weekends) && !holidays.has(day)) remaining--;
    }
    return day;
  }
  if (name === 'DAYS360') {
    if (args.length < 2 || args.length > 3) return '#VALUE!';
    const startSerial = numeric(args[0]), endSerial = numeric(args[1]);
    if (typeof startSerial !== 'number') return startSerial;
    if (typeof endSerial !== 'number') return endSerial;
    if ([startSerial, endSerial].some(value => value < 0 || value >= 2958466)) return '#NUM!';
    const method = args.length === 3 ? logical(args[2]) : false;
    if (typeof method !== 'boolean') return method;
    const sign = endSerial < startSerial ? -1 : 1;
    const start = dateParts(Math.floor(sign < 0 ? endSerial : startSerial));
    const end = dateParts(Math.floor(sign < 0 ? startSerial : endSerial));
    const lastDay = date => date.month === 2 && date.year === 1900 ? 29 : new Date(utcDate(date.year, date.month + 1, 0)).getUTCDate();
    const startLastFeb = start.month === 2 && start.day === lastDay(start);
    let startDay = start.day, endDay = end.day;
    if (method) {
      if (startDay === 31) startDay = 30;
      if (endDay === 31) endDay = 30;
    } else {
      if (startLastFeb) startDay = 30;
      if (end.month === 2 && end.day === lastDay(end) && startLastFeb) endDay = 30;
      if (endDay === 31 && startDay >= 30) endDay = 30;
      if (startDay === 31) startDay = 30;
    }
    return sign * (360 * (end.year - start.year) + 30 * (end.month - start.month) + endDay - startDay);
  }
  if (name === 'YEARFRAC') {
    if (args.length < 2 || args.length > 3) return '#VALUE!';
    const startSerial = numeric(args[0]), endSerial = numeric(args[1]);
    if (typeof startSerial !== 'number') return startSerial;
    if (typeof endSerial !== 'number') return endSerial;
    if ([startSerial, endSerial].some(value => value < 0 || value >= 2958466)) return '#NUM!';
    const basis = args.length === 3 ? numeric(args[2]) : 0;
    if (typeof basis !== 'number') return basis;
    if (!Number.isInteger(basis) || basis < 0 || basis > 4) return '#NUM!';
    const start = Math.min(startSerial, endSerial), end = Math.max(startSerial, endSerial), sign = endSerial < startSerial ? -1 : 1;
    if (basis === 0 || basis === 4) return sign * callFunction('DAYS360', [start, end, basis === 4]) / 360;
    if (basis === 2) return sign * (end - start) / 360;
    if (basis === 3) return sign * (end - start) / 365;
    let total = 0, cursor = start;
    while (cursor < end) {
      const date = dateParts(cursor);
      if (typeof date === 'string') return date;
      const yearEnd = date.year === 9999 ? 2958466 : dateToSerial(utcDate(date.year + 1, 1, 1));
      const segmentEnd = Math.min(end, yearEnd), yearLength = date.year === 1900 || date.year % 4 === 0 && (date.year % 100 !== 0 || date.year % 400 === 0) ? 366 : 365;
      total += (segmentEnd - cursor) / yearLength;
      cursor = segmentEnd;
    }
    return sign * total;
  }
  if (['COUPDAYBS', 'COUPDAYS', 'COUPDAYSNC', 'COUPNCD', 'COUPNUM', 'COUPPCD'].includes(name)) {
    if (args.length < 3 || args.length > 4) return '#VALUE!';
    const settlement = numeric(args[0]), maturity = numeric(args[1]), frequency = numeric(args[2]);
    const basis = args.length === 4 ? numeric(args[3]) : 0;
    for (const value of [settlement, maturity, frequency, basis]) if (typeof value !== 'number') return value;
    if (settlement < 0 || settlement >= 2958466 || maturity < 0 || maturity >= 2958466 || settlement >= maturity
      || ![1, 2, 4].includes(frequency) || !Number.isInteger(basis) || basis < 0 || basis > 4) return '#NUM!';
    const schedule = couponPeriodDates(Math.floor(settlement), Math.floor(maturity), frequency);
    if (typeof schedule === 'string') return schedule;
    if (name === 'COUPNCD') return schedule.next;
    if (name === 'COUPPCD') return schedule.previous;
    if (name === 'COUPNUM') return schedule.periodIndex + 1;
    const daysBetween = (start, end) => basis === 0 || basis === 4
      ? callFunction('DAYS360', [start, end, basis === 4])
      : end - start;
    if (name === 'COUPDAYS') return basis === 0 || basis === 4 ? 360 / frequency : schedule.next - schedule.previous;
    if (name === 'COUPDAYBS') return daysBetween(schedule.previous, Math.floor(settlement));
    return daysBetween(Math.floor(settlement), schedule.next);
  }
  if (name === 'TIME') {
    if (args.length !== 3) return '#VALUE!';
    const parts = args.map(numeric);
    if (parts.some(value => typeof value !== 'number')) return parts.find(value => typeof value === 'string');
    const seconds = Math.trunc(parts[0]) * 3600 + Math.trunc(parts[1]) * 60 + Math.trunc(parts[2]);
    return ((seconds % 86400) + 86400) % 86400 / 86400;
  }
  if (name === 'TIMEVALUE') {
    if (args.length !== 1) return '#VALUE!';
    if (isError(args[0])) return args[0];
    if (typeof args[0] !== 'string') return '#VALUE!';
    const seconds = parseTimeValue(args[0]);
    return seconds === null ? '#VALUE!' : seconds / 86400;
  }
  if (['HOUR', 'MINUTE', 'SECOND'].includes(name)) {
    if (args.length !== 1) return '#VALUE!';
    const serial = numeric(args[0]);
    if (typeof serial !== 'number') return serial;
    if (serial < 0 || serial >= 2958466) return '#NUM!';
    const seconds = Math.round((serial - Math.floor(serial)) * 86400) % 86400;
    return name === 'HOUR' ? Math.floor(seconds / 3600) : name === 'MINUTE' ? Math.floor(seconds / 60) % 60 : seconds % 60;
  }
  if (name === 'RATE') {
    if (args.length < 3 || args.length > 6) return '#VALUE!';
    const [periods, payment, presentValue, futureValue, paymentType, initialGuess] = [
      numeric(args[0]), numeric(args[1]), numeric(args[2]), numeric(args[3] ?? 0),
      args.length > 4 ? numeric(args[4]) : 0, args.length > 5 ? numeric(args[5]) : 0.1
    ];
    for (const value of [periods, payment, presentValue, futureValue, paymentType, initialGuess]) if (typeof value !== 'number') return value;
    if (periods <= 0 || paymentType !== 0 && paymentType !== 1 || initialGuess <= -1) return '#NUM!';
    if (payment === 0 && futureValue === 0 && presentValue !== 0) return '#NUM!';
    const tolerance = 1e-10 * Math.max(1, Math.abs(payment) * periods, Math.abs(presentValue), Math.abs(futureValue));
    let rate = initialGuess;
    for (let iteration = 0; iteration < 100; iteration++) {
      if (rate <= -1) return '#NUM!';
      let balance, derivative;
      if (Math.abs(rate) < 1e-8) {
        balance = presentValue + payment * periods + futureValue;
        derivative = presentValue * periods + payment * (periods * paymentType + periods * (periods - 1) / 2);
      } else {
        const factor = (1 + rate) ** periods;
        if (!Number.isFinite(factor)) return '#NUM!';
        const growth = factor - 1, periodGrowth = periods * (1 + rate) ** (periods - 1);
        const annuity = growth / rate * (1 + rate * paymentType);
        const annuityDerivative = ((periodGrowth * rate - growth) / (rate * rate)) * (1 + rate * paymentType) + growth / rate * paymentType;
        balance = presentValue * factor + payment * annuity + futureValue;
        derivative = presentValue * periodGrowth + payment * annuityDerivative;
      }
      if (!Number.isFinite(balance) || !Number.isFinite(derivative) || derivative === 0) return '#NUM!';
      if (Math.abs(balance) <= tolerance) return rate;
      const next = rate - balance / derivative;
      if (!Number.isFinite(next) || next <= -1) return '#NUM!';
      if (Math.abs(next - rate) <= 1e-10 * Math.max(1, Math.abs(rate))) return Math.abs(balance) <= tolerance ? next : '#NUM!';
      rate = next;
    }
    return '#NUM!';
  }
  if (name === 'EFFECT' || name === 'NOMINAL') {
    if (args.length !== 2) return '#VALUE!';
    const rate = numeric(args[0]), periods = numeric(args[1]);
    if (typeof rate !== 'number' || typeof periods !== 'number') return typeof rate !== 'number' ? rate : periods;
    if (rate < 0 || !Number.isInteger(periods) || periods < 1) return '#NUM!';
    const result = name === 'EFFECT' ? (1 + rate / periods) ** periods - 1 : periods * ((1 + rate) ** (1 / periods) - 1);
    return Number.isFinite(result) ? result : '#NUM!';
  }
  if (name === 'CUMIPMT' || name === 'CUMPRINC') {
    if (args.length !== 6) return '#VALUE!';
    const [rate, periods, presentValue, startPeriod, endPeriod, paymentType] = args.map(numeric);
    for (const value of [rate, periods, presentValue, startPeriod, endPeriod, paymentType]) if (typeof value !== 'number') return value;
    if (rate <= 0 || periods <= 0 || presentValue <= 0 || !Number.isInteger(periods) || !Number.isInteger(startPeriod) || !Number.isInteger(endPeriod)
      || startPeriod < 1 || endPeriod < startPeriod || endPeriod > periods || ![0, 1].includes(paymentType)) return '#NUM!';
    if (periods > MAX_DISTRIBUTION_SIZE) return '#NUM!';
    const payment = -callFunction('PMT', [rate, periods, presentValue, 0, paymentType]);
    if (!Number.isFinite(payment)) return '#NUM!';
    let balance = presentValue, total = 0;
    for (let period = 1; period <= endPeriod; period++) {
      let interest;
      if (paymentType === 1) {
        balance -= payment;
        interest = balance * rate;
        balance += interest;
      } else {
        interest = balance * rate;
        balance += interest - payment;
      }
      if (!Number.isFinite(balance) || !Number.isFinite(interest)) return '#NUM!';
      if (period >= startPeriod) total -= name === 'CUMIPMT' ? interest : paymentType === 1 ? payment : payment - interest;
    }
    return Number.isFinite(total) ? total : '#NUM!';
  }
  if (name === 'SLN') {
    if (args.length !== 3) return '#VALUE!';
    const [cost, salvage, life] = args.map(numeric);
    for (const value of [cost, salvage, life]) if (typeof value !== 'number') return value;
    if (cost < 0 || salvage < 0 || life <= 0) return '#NUM!';
    const result = (cost - salvage) / life;
    return Number.isFinite(result) ? result : '#NUM!';
  }
  if (name === 'SYD') {
    if (args.length !== 4) return '#VALUE!';
    const [cost, salvage, life, period] = args.map(numeric);
    for (const value of [cost, salvage, life, period]) if (typeof value !== 'number') return value;
    if (cost < 0 || salvage < 0 || life <= 0 || !Number.isInteger(period) || period < 1 || period > life) return '#NUM!';
    const result = (cost - salvage) * (life - period + 1) * 2 / (life * (life + 1));
    return Number.isFinite(result) ? result : '#NUM!';
  }
  if (name === 'DDB') {
    if (args.length < 4 || args.length > 5) return '#VALUE!';
    const [cost, salvage, life, period] = args.slice(0, 4).map(numeric), factor = args.length === 5 ? numeric(args[4]) : 2;
    for (const value of [cost, salvage, life, period, factor]) if (typeof value !== 'number') return value;
    const currentPeriod = Math.trunc(period);
    if (cost < 0 || salvage < 0 || life <= 0 || !Number.isInteger(life) || currentPeriod < 1 || currentPeriod > life || factor <= 0) return '#NUM!';
    if (life > MAX_DISTRIBUTION_SIZE) return '#NUM!';
    let bookValue = cost, depreciation = 0;
    for (let current = 1; current <= currentPeriod; current++) {
      depreciation = Math.min(Math.max(0, bookValue - salvage), bookValue * factor / life);
      bookValue -= depreciation;
      if (!Number.isFinite(bookValue) || !Number.isFinite(depreciation)) return '#NUM!';
    }
    return depreciation;
  }
  if (name === 'DB') {
    if (args.length < 4 || args.length > 5) return '#VALUE!';
    const [cost, salvage, life, period] = args.slice(0, 4).map(numeric), month = args.length === 5 && args[4] !== null ? numeric(args[4]) : 12;
    for (const value of [cost, salvage, life, period, month]) if (typeof value !== 'number') return value;
    const currentPeriod = Math.trunc(period), firstYearMonths = Math.trunc(month);
    if (cost <= 0 || salvage < 0 || salvage > cost || !Number.isInteger(life) || life <= 0 || !Number.isInteger(period)
      || firstYearMonths < 1 || firstYearMonths > 12 || currentPeriod < 1 || currentPeriod > life + (firstYearMonths < 12 ? 1 : 0)) return '#NUM!';
    if (life > MAX_DISTRIBUTION_SIZE) return '#NUM!';
    const rate = Math.round((1 - (salvage / cost) ** (1 / life)) * 1000) / 1000;
    let bookValue = cost, depreciation = cost * rate * firstYearMonths / 12;
    bookValue -= depreciation;
    if (currentPeriod === 1) return depreciation;
    for (let current = 2; current <= currentPeriod; current++) {
      const fraction = current === life + 1 ? (12 - firstYearMonths) / 12 : 1;
      depreciation = Math.min(Math.max(0, bookValue - salvage), bookValue * rate * fraction);
      bookValue -= depreciation;
      if (!Number.isFinite(bookValue) || !Number.isFinite(depreciation)) return '#NUM!';
    }
    return depreciation;
  }
  if (name === 'VDB') {
    if (args.length < 5 || args.length > 7) return '#VALUE!';
    const [cost, salvage, life, startPeriod, endPeriod] = args.slice(0, 5).map(numeric);
    const factor = args.length > 5 && args[5] !== null ? numeric(args[5]) : 2;
    const noSwitch = args.length > 6 ? logical(args[6]) : false;
    for (const value of [cost, salvage, life, startPeriod, endPeriod, factor]) if (typeof value !== 'number') return value;
    if (typeof noSwitch !== 'boolean') return noSwitch;
    if (cost < 0 || salvage < 0 || life <= 0 || startPeriod < 0 || endPeriod <= startPeriod || endPeriod > life || factor <= 0) return '#NUM!';
    if (life > MAX_DISTRIBUTION_SIZE) return '#NUM!';
    let bookValue = cost, total = 0;
    for (let period = 0; period < endPeriod; period++) {
      const declining = bookValue * factor / life;
      const straight = (bookValue - salvage) / (life - period);
      const depreciation = Math.min(Math.max(0, bookValue - salvage), noSwitch ? declining : Math.max(declining, straight));
      const elapsed = Math.max(0, Math.min(1, endPeriod - period));
      const overlap = Math.max(0, Math.min(endPeriod, period + 1) - Math.max(startPeriod, period));
      total += depreciation * overlap;
      bookValue -= depreciation * elapsed;
      if (!Number.isFinite(total) || !Number.isFinite(bookValue)) return '#NUM!';
    }
    return total;
  }
  if (name === 'FVSCHEDULE') {
    if (args.length !== 2) return '#VALUE!';
    const principal = numeric(args[0]);
    if (typeof principal !== 'number') return principal;
    const schedule = flatten([args[1]]);
    let factor = 1;
    for (const rate of schedule) {
      if (isError(rate)) return rate;
      if (typeof rate !== 'number' || !Number.isFinite(rate)) continue;
      if (rate <= -1) return '#NUM!';
      factor *= 1 + rate;
      if (!Number.isFinite(factor)) return '#NUM!';
    }
    const result = principal * factor;
    return Number.isFinite(result) ? result : '#NUM!';
  }
  if (['PMT', 'PV', 'FV', 'NPER'].includes(name)) {
    if (args.length < 3 || args.length > 5) return '#VALUE!';
    const parameters = [numeric(args[0]), numeric(args[1]), numeric(args[2]), numeric(args[3] ?? 0)];
    if (parameters.some(value => typeof value !== 'number')) return parameters.find(value => typeof value !== 'number');
    const [rate, second, third, optional] = parameters;
    const paymentType = args.length > 4 ? numeric(args[4]) : 0;
    if (typeof paymentType !== 'number') return paymentType;
    if (paymentType !== 0 && paymentType !== 1) return '#NUM!';
    if (name !== 'NPER' && second <= 0) return '#NUM!';
    if (rate <= -1) return '#NUM!';
    if (name === 'NPER') {
      const payment = second, presentValue = third, futureValue = optional;
      if (payment === 0) return presentValue + futureValue === 0 ? '#NUM!' : '#DIV/0!';
      if (rate === 0) {
        const result = -(presentValue + futureValue) / payment;
        return Number.isFinite(result) && result > 0 ? result : '#NUM!';
      }
      const paymentFactor = payment * (1 + rate * paymentType);
      const numerator = paymentFactor - futureValue * rate, denominator = presentValue * rate + paymentFactor;
      const base = numerator / denominator;
      if (!Number.isFinite(base) || base <= 0) return '#NUM!';
      const result = Math.log(base) / Math.log1p(rate);
      return Number.isFinite(result) && result > 0 ? result : '#NUM!';
    }
    const nper = second, factor = rate === 0 ? 1 : (1 + rate) ** nper;
    if (!Number.isFinite(factor) || factor <= 0) return '#NUM!';
    const annuityFactor = rate === 0 ? nper : ((factor - 1) / rate) * (1 + rate * paymentType);
    let result;
    if (name === 'PMT') result = rate === 0 ? -(third + optional) / nper : -(optional + third * factor) / annuityFactor;
    else if (name === 'PV') result = rate === 0 ? -(optional + third * nper) : -(optional + third * annuityFactor) / factor;
    else result = -(optional * factor + third * annuityFactor);
    return Number.isFinite(result) ? (result === 0 ? 0 : result) : '#NUM!';
  }
  if (values.some(isError)) return values.find(isError);
  if (['SUMXMY2', 'SUMX2MY2', 'SUMX2PY2'].includes(name)) {
    if (args.length !== 2) return '#VALUE!';
    const left = matrixValue(args[0]), right = matrixValue(args[1]);
    if (!left || !right) return '#VALUE!';
    if (left.length !== right.length || left[0].length !== right[0].length) return '#N/A';
    let total = 0;
    for (let row = 0; row < left.length; row++) for (let column = 0; column < left[row].length; column++) {
      const x = left[row][column], y = right[row][column];
      if (isError(x) || isError(y)) return isError(x) ? x : y;
      if (typeof x !== 'number' || !Number.isFinite(x) || typeof y !== 'number' || !Number.isFinite(y)) continue;
      const term = name === 'SUMXMY2' ? (x - y) ** 2 : name === 'SUMX2MY2' ? x ** 2 - y ** 2 : x ** 2 + y ** 2;
      total += term;
      if (!Number.isFinite(total)) return '#NUM!';
    }
    return total;
  }
  if (['SUMSQ', 'DEVSQ', 'AVEDEV', 'GEOMEAN', 'HARMEAN'].includes(name)) {
    const numbers = aggregateNumbers(args);
    if (typeof numbers === 'string') return numbers;
    if (name === 'SUMSQ') {
      let total = 0;
      for (const value of numbers) { total += value * value; if (!Number.isFinite(total)) return '#NUM!'; }
      return total;
    }
    if (!numbers.length) return '#NUM!';
    if (['GEOMEAN', 'HARMEAN'].includes(name) && numbers.some(value => value <= 0)) return '#NUM!';
    if (name === 'GEOMEAN') {
      const result = Math.exp(numbers.reduce((sum, value) => sum + Math.log(value), 0) / numbers.length);
      return Number.isFinite(result) ? result : '#NUM!';
    }
    if (name === 'HARMEAN') {
      const result = numbers.length / numbers.reduce((sum, value) => sum + 1 / value, 0);
      return Number.isFinite(result) ? result : '#NUM!';
    }
    const directSum = numbers.reduce((sum, value) => sum + value, 0);
    if (Number.isFinite(directSum)) {
      const directMean = directSum / numbers.length;
      const deviations = numbers.reduce((sum, value) => sum + Math.abs(value - directMean), 0);
      const result = name === 'AVEDEV' ? deviations / numbers.length : numbers.reduce((sum, value) => sum + (value - directMean) ** 2, 0);
      if (Number.isFinite(result)) return result;
    }
    const scale = Math.max(...numbers.map(Math.abs));
    if (scale === 0) return 0;
    const normalized = numbers.map(value => value / scale);
    const mean = normalized.reduce((sum, value) => sum + value, 0) / numbers.length;
    if (name === 'AVEDEV') {
      const normalizedDeviation = normalized.reduce((sum, value) => sum + Math.abs(value - mean), 0) / numbers.length;
      const result = normalizedDeviation * scale;
      return Number.isFinite(result) ? result : '#NUM!';
    }
    const normalizedSquares = normalized.reduce((sum, value) => sum + (value - mean) ** 2, 0);
    if (normalizedSquares === 0) return 0;
    const root = scale * Math.sqrt(normalizedSquares);
    const result = root * root;
    return Number.isFinite(result) ? result : '#NUM!';
  }
  if (name === 'COUNTA') return values.filter(value => value !== null && value !== undefined).length;
  if (name === 'COUNT') return countDirectArguments(sourceArgs);
  if (['AVERAGEA', 'MINA', 'MAXA'].includes(name)) {
    const items = flatten(args).filter(value => value !== null && value !== undefined).map(value => typeof value === 'number' ? value : typeof value === 'boolean' ? Number(value) : 0);
    if (name === 'AVERAGEA') {
      if (!items.length) return '#DIV/0!';
      let sum = 0;
      for (const value of items) { sum += value; if (!Number.isFinite(sum)) return '#NUM!'; }
      return sum / items.length;
    }
    if (!items.length) return 0;
    let result = items[0];
    for (let index = 1; index < items.length; index++) result = name === 'MINA' ? Math.min(result, items[index]) : Math.max(result, items[index]);
    return result;
  }
  if (['SUM', 'PRODUCT', 'AVERAGE', 'MEDIAN', 'MIN', 'MAX'].includes(name)) {
    if (name === 'PRODUCT' && !args.length) return '#VALUE!';
    const numbers = name === 'SUM' || name === 'AVERAGE' ? aggregateDirectArguments(sourceArgs) : aggregateNumbers(args);
    if (typeof numbers === 'string') return numbers;
    if (name === 'SUM') {
      let sum = 0;
      for (const value of numbers) { sum += value; if (!Number.isFinite(sum)) return '#NUM!'; }
      return sum;
    }
    if (name === 'PRODUCT') {
      let product = 1;
      for (const number of numbers) { product *= number; if (!Number.isFinite(product)) return '#NUM!'; }
      return product;
    }
    if (name === 'MEDIAN') {
      if (!numbers.length) return '#NUM!';
      numbers.sort((a, b) => a - b);
      const middle = Math.floor(numbers.length / 2);
      if (numbers.length % 2) return numbers[middle];
      const lower = numbers[middle - 1], upper = numbers[middle];
      return Math.sign(lower) === Math.sign(upper) ? lower + (upper - lower) / 2 : lower / 2 + upper / 2;
    }
    if (name === 'AVERAGE') {
      if (!numbers.length) return '#DIV/0!';
      let sum = 0;
      for (const value of numbers) { sum += value; if (!Number.isFinite(sum)) return '#NUM!'; }
      const average = sum / numbers.length;
      return Number.isFinite(average) ? average : '#NUM!';
    }
    if (!numbers.length) return 0;
    return name === 'MIN' ? Math.min(...numbers) : Math.max(...numbers);
  }
  if (name === 'LARGE' || name === 'SMALL') {
    if (args.length !== 2) return '#VALUE!';
    const numbers = aggregateNumbers([args[0]]);
    if (typeof numbers === 'string') return numbers;
    const rank = numeric(args[1]);
    if (typeof rank !== 'number') return rank;
    if (!Number.isInteger(rank) || rank < 1 || rank > numbers.length) return '#NUM!';
    numbers.sort((a, b) => name === 'LARGE' ? b - a : a - b);
    return numbers[rank - 1];
  }
  if (name === 'MODE.SNGL' || name === 'MODE.MULT') {
    if (args.length !== 1) return '#VALUE!';
    const numbers = flatten(args[0]).filter(value => typeof value === 'number' && Number.isFinite(value));
    if (!numbers.length) return '#N/A';
    const counts = new Map();
    for (const value of numbers) counts.set(value, (counts.get(value) || 0) + 1);
    let frequency = 0;
    for (const count of counts.values()) frequency = Math.max(frequency, count);
    if (frequency < 2) return '#N/A';
    if (name === 'MODE.SNGL') return numbers.find(value => counts.get(value) === frequency);
    return [...new Set(numbers)].filter(value => counts.get(value) === frequency).map(value => [value]);
  }
  if (name === 'FREQUENCY') {
    if (args.length !== 2) return '#VALUE!';
    const data = flatten(args[0]).filter(value => typeof value === 'number' && Number.isFinite(value));
    const bins = flatten(args[1]).filter(value => typeof value === 'number' && Number.isFinite(value));
    const counts = Array(bins.length + 1).fill(0);
    for (const value of data) {
      let bin = bins.findIndex(boundary => value <= boundary);
      if (bin < 0) bin = bins.length;
      counts[bin]++;
    }
    return counts.map(count => [count]);
  }
  if (name === 'RANK.EQ' || name === 'RANK.AVG') {
    if (args.length < 2 || args.length > 3) return '#VALUE!';
    const sought = numeric(args[0]);
    if (typeof sought !== 'number') return sought;
    const numbers = aggregateNumbers([args[1]]);
    if (typeof numbers === 'string') return numbers;
    if (args[2] !== undefined && args[2] !== null) {
      const order = numeric(args[2]);
      if (typeof order !== 'number') return order;
      numbers.sort((a, b) => order === 0 ? b - a : a - b);
    } else numbers.sort((a, b) => b - a);
    const firstRank = numbers.findIndex(value => value === sought);
    if (firstRank < 0) return '#N/A';
    if (name === 'RANK.EQ') return firstRank + 1;
    let ties = 0;
    for (const value of numbers) if (value === sought) ties++;
    return firstRank + (ties + 1) / 2;
  }
  if (name === 'PERCENTRANK.INC' || name === 'PERCENTRANK.EXC') {
    if (args.length < 2 || args.length > 3) return '#VALUE!';
    const sought = numeric(args[1]);
    if (typeof sought !== 'number') return sought;
    const numbers = aggregateNumbers([args[0]]);
    if (typeof numbers === 'string') return numbers;
    if (!numbers.length) return '#NUM!';
    let significance = 3;
    if (args[2] !== undefined && args[2] !== null) {
      significance = numeric(args[2]);
      if (typeof significance !== 'number') return significance;
      if (significance < 1 || significance > 15) return '#NUM!';
      significance = Math.trunc(significance);
    }
    numbers.sort((a, b) => a - b);
    if (sought < numbers[0] || sought > numbers[numbers.length - 1]) return '#N/A';
    const exact = numbers.indexOf(sought);
    let rank;
    if (exact >= 0) rank = exact + 1;
    else {
      const upper = numbers.findIndex(value => value > sought);
      const lower = upper - 1;
      rank = lower + 1 + interpolationFraction(numbers[lower], numbers[upper], sought);
    }
    const exclusive = name.endsWith('.EXC');
    const percentile = exclusive ? rank / (numbers.length + 1) : numbers.length === 1 ? 0 : (rank - 1) / (numbers.length - 1);
    const factor = 10 ** significance;
    return Math.round((percentile + Number.EPSILON) * factor) / factor;
  }
  if (['PERCENTILE', 'PERCENTILE.INC', 'PERCENTILE.EXC', 'QUARTILE', 'QUARTILE.INC', 'QUARTILE.EXC'].includes(name)) {
    if (args.length !== 2) return '#VALUE!';
    const numbers = aggregateNumbers([args[0]]);
    if (typeof numbers === 'string') return numbers;
    if (!numbers.length) return '#NUM!';
    let k = numeric(args[1]);
    if (typeof k !== 'number') return k;
    const exclusive = name.endsWith('.EXC');
    if (name.startsWith('QUARTILE')) {
      if (!Number.isInteger(k) || k < (exclusive ? 1 : 0) || k > (exclusive ? 3 : 4)) return '#NUM!';
      k /= 4;
    } else if (exclusive ? k <= 0 || k >= 1 : k < 0 || k > 1) return '#NUM!';
    numbers.sort((a, b) => a - b);
    const position = exclusive ? (numbers.length + 1) * k - 1 : (numbers.length - 1) * k;
    if (position < 0 || position > numbers.length - 1) return '#NUM!';
    const lower = Math.floor(position), fraction = position - lower;
    return interpolateNumber(numbers[lower], numbers[Math.min(lower + 1, numbers.length - 1)], fraction);
  }
  if (['STDEV.S', 'STDEV.P', 'VAR.S', 'VAR.P'].includes(name)) {
    const numbers = aggregateNumbers(args);
    if (typeof numbers === 'string') return numbers;
    const sample = name.endsWith('.S');
    if (numbers.length < (sample ? 2 : 1)) return '#DIV/0!';
    return varianceValue(numbers, sample, name.startsWith('STDEV'));
  }
  if (['STDEVA', 'STDEVPA', 'VARA', 'VARPA'].includes(name)) {
    const numbers = flatten(args).filter(value => value !== null && value !== undefined).map(value => typeof value === 'number' ? value : typeof value === 'boolean' ? Number(value) : 0);
    const sample = name === 'STDEVA' || name === 'VARA';
    if (numbers.length < (sample ? 2 : 1)) return '#DIV/0!';
    return varianceValue(numbers, sample, name.startsWith('STDEV'));
  }
  if (['CORREL', 'PEARSON', 'COVARIANCE.P', 'COVARIANCE.S', 'SLOPE', 'INTERCEPT', 'RSQ', 'STEYX', 'FORECAST', 'FORECAST.LINEAR'].includes(name)) {
    const forecast = name === 'FORECAST' || name === 'FORECAST.LINEAR';
    if (forecast ? args.length !== 3 : args.length !== 2) return '#VALUE!';
    const symmetric = ['CORREL', 'PEARSON', 'COVARIANCE.P', 'COVARIANCE.S'].includes(name);
    const pairs = pairedNumbers(forecast ? args[2] : symmetric ? args[0] : args[1], forecast ? args[1] : args[symmetric ? 1 : 0]);
    if (typeof pairs === 'string') return pairs;
    const minimum = name === 'COVARIANCE.P' ? 1 : name === 'STEYX' ? 3 : 2;
    if (pairs.length < minimum) return '#DIV/0!';
    const constantX = pairs.every(pair => pair[0] === pairs[0][0]);
    const constantY = pairs.every(pair => pair[1] === pairs[0][1]);
    if (['COVARIANCE.P', 'COVARIANCE.S'].includes(name) && (constantX || constantY)) return 0;
    if (constantX || (['CORREL', 'PEARSON', 'RSQ'].includes(name) && constantY)) return '#DIV/0!';
    if (constantY) {
      if (name === 'SLOPE' || name === 'STEYX') return 0;
      if (name === 'INTERCEPT' || forecast) return pairs[0][1];
    }
    const count = pairs.length;
    let workPairs = pairs, xScale = 1, yScale = 1, scaled = false;
    let meanX = pairs.reduce((sum, pair) => sum + pair[0], 0) / count;
    let meanY = pairs.reduce((sum, pair) => sum + pair[1], 0) / count;
    let sumXY = 0, sumXX = 0, sumYY = 0;
    if (Number.isFinite(meanX) && Number.isFinite(meanY)) {
      for (const [x, y] of pairs) { const dx = x - meanX, dy = y - meanY; sumXY += dx * dy; sumXX += dx * dx; sumYY += dy * dy; }
    }
    if (![meanX, meanY, sumXY, sumXX, sumYY].every(Number.isFinite)) {
      scaled = true;
      xScale = Math.max(...pairs.map(pair => Math.abs(pair[0]))) || 1;
      yScale = Math.max(...pairs.map(pair => Math.abs(pair[1]))) || 1;
      workPairs = pairs.map(([x, y]) => [x / xScale, y / yScale]);
      meanX = workPairs.reduce((sum, pair) => sum + pair[0], 0) / count;
      meanY = workPairs.reduce((sum, pair) => sum + pair[1], 0) / count;
      sumXY = 0; sumXX = 0; sumYY = 0;
      for (const [x, y] of workPairs) { const dx = x - meanX, dy = y - meanY; sumXY += dx * dy; sumXX += dx * dx; sumYY += dy * dy; }
      if (![meanX, meanY, sumXY, sumXX, sumYY].every(Number.isFinite)) return '#NUM!';
    }
    if (name === 'COVARIANCE.P' || name === 'COVARIANCE.S') {
      const covariance = sumXY / (name === 'COVARIANCE.P' ? count : count - 1);
      if (!scaled) return covariance;
      const smallerScale = Math.min(xScale, yScale), largerScale = Math.max(xScale, yScale);
      const result = covariance * smallerScale * largerScale;
      return Number.isFinite(result) ? result : '#NUM!';
    }
    if (sumXX === 0 || (name === 'CORREL' || name === 'PEARSON' || name === 'RSQ') && sumYY === 0) return '#DIV/0!';
    const normalizedSlope = sumXY / sumXX, scaleRatio = scaled ? yScale / xScale : 1;
    const slope = normalizedSlope * scaleRatio, correlation = sumXY / Math.sqrt(sumXX) / Math.sqrt(sumYY);
    const normalizedIntercept = meanY - normalizedSlope * meanX, intercept = scaled ? normalizedIntercept * yScale : normalizedIntercept;
    if (![slope, correlation, intercept].every(Number.isFinite)) return '#NUM!';
    if (name === 'CORREL' || name === 'PEARSON') return correlation;
    if (name === 'SLOPE') return slope;
    if (name === 'INTERCEPT') return intercept;
    if (name === 'RSQ') return correlation * correlation;
    if (name === 'STEYX') {
      const residuals = workPairs.reduce((sum, [x, y]) => sum + (y - (normalizedIntercept + normalizedSlope * x)) ** 2, 0);
      const error = Math.sqrt(residuals / (count - 2)) * (scaled ? yScale : 1);
      return Number.isFinite(error) ? error : '#NUM!';
    }
    const x = numeric(args[0]);
    if (typeof x !== 'number') return x;
    const forecastValue = scaled ? (normalizedIntercept + normalizedSlope * (x / xScale)) * yScale : intercept + slope * x;
    return Number.isFinite(forecastValue) ? forecastValue : '#NUM!';
  }
  const first = ['GCD', 'LCM', 'MULTINOMIAL'].includes(name) && (isFormulaRange(args[0]) || args[0]?.[FORMULA_RANGE_METADATA]) ? 0 : numeric(args[0]);
  if (typeof first === 'string') return first;
  if (name === 'FACT' || name === 'FACTDOUBLE') {
    if (args.length !== 1) return '#VALUE!';
    const number = Math.trunc(first);
    if (number < 0) return '#NUM!';
    let result = 1;
    for (let factor = name === 'FACT' ? 2 : number % 2 === 0 ? 2 : 1; factor <= number; factor += name === 'FACT' ? 1 : 2) {
      result *= factor;
      if (!Number.isFinite(result)) return '#NUM!';
    }
    return result;
  }
  if (name === 'COMBIN' || name === 'PERMUT') {
    if (args.length !== 2) return '#VALUE!';
    const second = numeric(args[1]);
    if (typeof second !== 'number') return second;
    const n = Math.trunc(first), k = Math.trunc(second);
    if (n < 0 || k < 0 || k > n || n > Number.MAX_SAFE_INTEGER) return '#NUM!';
    let result = 1;
    if (name === 'PERMUT') {
      for (let index = 0; index < k; index++) { result *= n - index; if (!Number.isFinite(result)) return '#NUM!'; }
      return result;
    }
    const choose = Math.min(k, n - k);
    for (let index = 1; index <= choose; index++) { result *= (n - choose + index) / index; if (!Number.isFinite(result)) return '#NUM!'; }
    return result;
  }
  if (name === 'GCD' || name === 'LCM' || name === 'MULTINOMIAL') {
    if (args.length < 1 || args.length > 255) return '#VALUE!';
    const numbers = flatten(args).map(numeric);
    if (numbers.some(value => typeof value !== 'number')) return numbers.find(value => typeof value !== 'number');
    const integers = numbers.map(value => Math.trunc(value));
    if (integers.some(value => Math.abs(value) > Number.MAX_SAFE_INTEGER || name === 'MULTINOMIAL' && value < 0)) return '#NUM!';
    if (name === 'GCD' || name === 'LCM') {
      let result = name === 'GCD' ? 0 : 1;
      for (const value of integers) {
        const number = Math.abs(value);
        if (name === 'GCD') {
          let left = result, right = number;
          while (right) [left, right] = [right, left % right];
          result = left;
        } else {
          if (!number || !result) { result = 0; continue; }
          let left = result, right = number;
          while (right) [left, right] = [right, left % right];
          result = result / left * number;
          if (!Number.isSafeInteger(result)) return '#NUM!';
        }
      }
      return result;
    }
    let total = 0, result = 1;
    for (const count of integers) {
      if (total + count > Number.MAX_SAFE_INTEGER) return '#NUM!';
      const choose = Math.min(total, count);
      let combinations = 1;
      for (let index = 1; index <= choose; index++) { combinations *= (total + count - choose + index) / index; if (!Number.isFinite(combinations)) return '#NUM!'; }
      result *= combinations;
      if (!Number.isFinite(result)) return '#NUM!';
      total += count;
    }
    return result;
  }
  if (name === 'ABS' && args.length === 1) return Math.abs(first);
  if (name === 'INT' && args.length === 1) return Math.floor(first);
  if (name === 'SQRT' && args.length === 1) return first < 0 ? '#NUM!' : Math.sqrt(first);
  if (name === 'PI') return args.length === 0 ? Math.PI : '#VALUE!';
  const unaryMath = {
    SIN: Math.sin, COS: Math.cos, TAN: Math.tan, ASIN: Math.asin, ACOS: Math.acos, ATAN: Math.atan,
    SINH: Math.sinh, COSH: Math.cosh, TANH: Math.tanh, ASINH: Math.asinh, ACOSH: Math.acosh, ATANH: Math.atanh,
    DEGREES: value => value * 180 / Math.PI, RADIANS: value => value * Math.PI / 180
  };
  if (Object.hasOwn(unaryMath, name)) {
    if (args.length !== 1) return '#VALUE!';
    if ((name === 'ASIN' || name === 'ACOS') && (first < -1 || first > 1)) return '#NUM!';
    if (name === 'ACOSH' && first < 1 || name === 'ATANH' && Math.abs(first) >= 1) return '#NUM!';
    const result = unaryMath[name](first);
    return Number.isFinite(result) ? (result === 0 ? 0 : result) : '#NUM!';
  }
  if (name === 'ATAN2') {
    if (args.length !== 2) return '#VALUE!';
    const y = numeric(args[1]);
    if (typeof y !== 'number') return y;
    if (first === 0 && y === 0) return '#DIV/0!';
    return Math.atan2(y, first);
  }
  if (name === 'EXP' && args.length === 1) {
    const result = Math.exp(first);
    return Number.isFinite(result) ? result : '#NUM!';
  }
  if (name === 'LN' || name === 'LOG10' || name === 'LOG') {
    if (name === 'LOG' ? args.length < 1 || args.length > 2 : args.length !== 1) return '#VALUE!';
    if (first <= 0) return '#NUM!';
    if (name === 'LN') return Math.log(first);
    if (name === 'LOG10') return Math.log10(first);
    const base = args.length === 1 ? 10 : numeric(args[1]);
    if (typeof base === 'string') return base;
    if (base <= 0) return '#NUM!';
    if (base === 1) return '#DIV/0!';
    return Math.log(first) / Math.log(base);
  }
  if (name === 'TRUNC') {
    if (args.length < 1 || args.length > 2) return '#VALUE!';
    const digits = args.length === 1 ? 0 : numeric(args[1]);
    if (typeof digits === 'string') return digits;
    if (Math.abs(digits) > 100) return '#NUM!';
    const factor = 10 ** digits;
    const result = Math.trunc(first * factor) / factor;
    return Number.isFinite(result) ? Number(result.toPrecision(15)) : '#NUM!';
  }
  if (name === 'CEILING.MATH' || name === 'FLOOR.MATH') {
    if (args.length < 1 || args.length > 3) return '#VALUE!';
    const significance = args[1] === null || args[1] === undefined ? 1 : numeric(args[1]);
    if (typeof significance === 'string') return significance;
    const mode = args.length < 3 || args[2] === null || args[2] === undefined ? 0 : numeric(args[2]);
    if (typeof mode === 'string') return mode;
    const magnitude = Math.abs(significance);
    if (magnitude === 0 || first === 0) return 0;
    const rawRatio = first / magnitude;
    const nearestInteger = Math.round(rawRatio);
    const ratio = Math.abs(rawRatio - nearestInteger) <= Number.EPSILON * Math.max(1, Math.abs(rawRatio)) * 4 ? nearestInteger : rawRatio;
    let rounded;
    if (name === 'CEILING.MATH') rounded = first < 0 && mode !== 0 ? Math.floor(ratio) : Math.ceil(ratio);
    else rounded = first < 0 && mode !== 0 ? Math.ceil(ratio) : Math.floor(ratio);
    const result = rounded * magnitude;
    return Number.isFinite(result) ? Number(result.toPrecision(15)) : '#NUM!';
  }
  if (name === 'CEILING' || name === 'FLOOR' || name === 'MROUND') {
    if (args.length !== 2) return '#VALUE!';
    const significance = numeric(args[1]);
    if (typeof significance === 'string') return significance;
    if (name === 'MROUND' && first !== 0 && significance !== 0 && Math.sign(first) !== Math.sign(significance)) return '#NUM!';
    if (first === 0 || significance === 0) return 0;
    if (name !== 'MROUND' && Math.sign(first) !== Math.sign(significance)) return '#NUM!';
    const magnitude = Math.abs(significance), rawRatio = Math.abs(first) / magnitude;
    const nearestInteger = Math.round(rawRatio);
    const ratio = Math.abs(rawRatio - nearestInteger) <= Number.EPSILON * Math.max(1, rawRatio) * 4 ? nearestInteger : rawRatio;
    const rounded = name === 'FLOOR' ? Math.floor(ratio) : name === 'CEILING' ? Math.ceil(ratio) : Math.floor(ratio + 0.5);
    const result = Math.sign(first) * rounded * magnitude;
    return Number.isFinite(result) ? Number(result.toPrecision(15)) : '#NUM!';
  }
  if (['ROUND', 'ROUNDUP', 'ROUNDDOWN'].includes(name)) {
    if (args.length !== 2) return '#VALUE!';
    const digits = numeric(args[1]);
    if (typeof digits === 'string') return digits;
    if (Math.abs(digits) > 100) return '#NUM!';
    const factor = 10 ** digits;
    const rawMagnitude = Math.abs(first) * factor;
    const nearestInteger = Math.round(rawMagnitude);
    const magnitude = Math.abs(rawMagnitude - nearestInteger) <= Number.EPSILON * Math.max(1, rawMagnitude) * 4
      ? nearestInteger
      : rawMagnitude;
    const rounded = name === 'ROUND'
      ? Math.round(magnitude)
      : name === 'ROUNDUP'
        ? Math.ceil(magnitude)
        : Math.floor(magnitude);
    const result = Math.sign(first) * rounded / factor;
    return Number.isFinite(result) ? result : '#NUM!';
  }
  if (name === 'MOD' && args.length === 2) {
    const divisor = numeric(args[1]);
    if (typeof divisor === 'string') return divisor;
    return divisor === 0 ? '#DIV/0!' : first - divisor * Math.floor(first / divisor);
  }
  if (name === 'POWER' && args.length === 2) {
    const exponent = numeric(args[1]);
    if (typeof exponent === 'string') return exponent;
    if (first === 0 && exponent < 0) return '#DIV/0!';
    const result = first ** exponent;
    return Number.isFinite(result) ? result : '#NUM!';
  }
  return '#NAME?';
}

function evaluateFormula(formula, context = {}) {
  try {
    const tokens = tokenize(formula);
    let position = 0;
    let captureFunctionArguments = 0;
    const peek = () => tokens[position];
    const take = type => { if (peek().type !== type) throw new Error('#VALUE!'); return tokens[position++]; };
    const precedence = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };
    function primary() {
      const token = peek();
      if (token.type === 'value') {
        position++;
        return captureFunctionArguments && typeof token.value === 'string' ? { [FORMULA_TEXT_LITERAL]: true, value: token.value } : token.value;
      }
      if (token.type === 'error') { position++; return token.value; }
      if (token.type === 'structuredReference') {
        position++;
        const parts = structuredReferenceParts(token.spec);
        if (!parts) return '#VALUE!';
        return context.getTableReference ? context.getTableReference(token.tableName, parts, context.formulaRow, context.formulaColumn) : '#NAME?';
      }
      if (token.type === 'reference') {
        position++;
        const start = token;
        let end = start;
        if (peek().type === ':') { position++; end = take('reference'); }
        if (peek().type === '#') {
          position++;
          if (start.axis || end !== start) return '#REF!';
          return context.getSpillRange?.(start.row, start.column, start.sheetName) ?? '#REF!';
        }
        if (start.axis || end.axis) {
          if (!start.axis || start.axis !== end.axis) return '#REF!';
          const sheetName = start.sheetName || end.sheetName;
          if (start.sheetName && end.sheetName && start.sheetName.toLocaleLowerCase() !== end.sheetName.toLocaleLowerCase()) return '#REF!';
          const bounds = start.axis === 'column'
            ? { top: 0, bottom: 1048575, left: Math.min(start.column, end.column), right: Math.max(start.column, end.column) }
            : { top: Math.min(start.row, end.row), bottom: Math.max(start.row, end.row), left: 0, right: 16383 };
          return context.getSparseRange?.(bounds, sheetName) ?? '#VALUE!';
        }
        const height = Math.abs(end.row - start.row) + 1, width = Math.abs(end.column - start.column) + 1;
        if (height * width > MAX_RANGE_CELLS) return '#VALUE!';
        const sheetName = start.sheetName || end.sheetName;
        if (start.sheetName && end.sheetName && start.sheetName.toLocaleLowerCase() !== end.sheetName.toLocaleLowerCase()) return '#REF!';
        if (height === 1 && width === 1 && end === start) {
          const value = context.getCell?.(start.row, start.column, sheetName) ?? null;
          return captureFunctionArguments ? { [FORMULA_REFERENCE_VALUE]: true, value } : value;
        }
        const top = Math.min(start.row, end.row), bottom = Math.max(start.row, end.row), left = Math.min(start.column, end.column), right = Math.max(start.column, end.column);
        const values = [], entries = [];
        for (let row = top; row <= bottom; row++) {
          const line = [];
          for (let column = left; column <= right; column++) {
            const value = context.getCell?.(row, column, sheetName) ?? null;
            line.push(value);
            entries.push({ row, column, value, formula: context.getCellFormula?.(row, column, sheetName) || null });
          }
          values.push(line);
        }
        Object.defineProperty(values, FORMULA_RANGE_METADATA, { value: { bounds: { top, bottom, left, right }, entries, sheetIndex: context.getSheetIndex?.(sheetName) ?? null } });
        return values;
      }
      if (token.type === '(') { position++; const value = expression(0); take(')'); return value; }
      if (token.type === 'name') {
        position++;
        if (token.value === 'TRUE' || token.value === 'FALSE') {
          const value = token.value === 'TRUE';
          if (peek().type === '(') { position++; take(')'); }
          return value;
        }
        if (peek().type !== '(') {
          const variable = context.variables?.get(token.value.toLocaleLowerCase());
          return variable === undefined ? context.getDefinedName?.(token.value) ?? '#NAME?' : variable;
        }
        position++;
        if (token.value === 'LET') {
          const previousVariables = context.variables;
          context.variables = new Map(previousVariables || []);
          try {
            const localNames = new Set();
            let pairs = 0;
            while (peek().type === 'name' && tokens[position + 1]?.type === ',') {
              const variable = take('name').value, normalized = variable.toLocaleLowerCase();
              if (/^[A-Z]{1,3}[1-9]\d*$/i.test(variable) || localNames.has(normalized)) return '#VALUE!';
              position++;
              const value = expression(0);
              if (isError(value)) return value;
              context.variables.set(normalized, value);
              localNames.add(normalized);
              if (++pairs > 126) return '#VALUE!';
              if (peek().type !== ',') return '#VALUE!';
              position++;
              if (peek().type === 'name' && tokens[position + 1]?.type === ',') continue;
              const result = expression(0);
              take(')');
              return result;
            }
            return '#VALUE!';
          } finally { context.variables = previousVariables; }
        }
        if (token.value === 'OFFSET') {
          if (peek().type !== 'reference') { expression(0); while (peek().type === ',') { position++; expression(0); } take(')'); return '#VALUE!'; }
          const start = take('reference');
          let end = start;
          if (peek().type === ':') { position++; end = take('reference'); }
          if (start.axis || end.axis || start.sheetName && end.sheetName && start.sheetName.toLocaleLowerCase() !== end.sheetName.toLocaleLowerCase()) { while (peek().type !== ')' && peek().type !== 'eof') position++; take(')'); return '#REF!'; }
          const sheetName = start.sheetName || end.sheetName;
          if (sheetName && context.getSheetIndex?.(sheetName) === -1) { while (peek().type !== ')' && peek().type !== 'eof') position++; take(')'); return '#REF!'; }
          if (peek().type !== ',') { take(')'); return '#VALUE!'; }
          position++;
          const rowOffset = expression(0);
          take(',');
          const columnOffset = expression(0);
          const originalTop = Math.min(start.row, end.row), originalLeft = Math.min(start.column, end.column);
          const originalHeight = Math.abs(end.row - start.row) + 1, originalWidth = Math.abs(end.column - start.column) + 1;
          let heightValue = originalHeight, widthValue = originalWidth;
          if (peek().type === ',') {
            position++;
            if (peek().type !== ',' && peek().type !== ')') heightValue = expression(0);
            if (peek().type === ',') { position++; if (peek().type !== ')') widthValue = expression(0); }
          }
          take(')');
          for (const value of [rowOffset, columnOffset, heightValue, widthValue]) if (isError(value)) return value;
          const rowDelta = numeric(rowOffset), columnDelta = numeric(columnOffset), heightNumber = numeric(heightValue), widthNumber = numeric(widthValue);
          if ([rowDelta, columnDelta, heightNumber, widthNumber].some(value => typeof value !== 'number')) return '#VALUE!';
          const top = originalTop + Math.trunc(rowDelta), left = originalLeft + Math.trunc(columnDelta), height = Math.trunc(heightNumber), width = Math.trunc(widthNumber);
          if (height < 1 || width < 1) return '#VALUE!';
          const bottom = top + height - 1, right = left + width - 1;
          if (top < 0 || left < 0 || bottom >= 1048576 || right >= 16384) return '#REF!';
          if (height * width > MAX_RANGE_CELLS) return '#NUM!';
          const result = Array.from({ length: height }, (_, row) => Array.from({ length: width }, (_, column) => context.getCell?.(top + row, left + column, sheetName) ?? null));
          return height === 1 && width === 1 ? result[0][0] : result;
        }
        if (token.value === 'ISFORMULA' || token.value === 'FORMULATEXT') {
          const reference = take('reference');
          take(')');
          if (reference.sheetName && context.getSheetIndex?.(reference.sheetName) === -1) return '#REF!';
          const formula = context.getCellFormula?.(reference.row, reference.column, reference.sheetName) || null;
          return token.value === 'ISFORMULA' ? Boolean(formula) : formula ? '=' + formula : '#N/A';
        }
        if (token.value === 'ISREF') {
          if (peek().type === 'reference') {
            const referenceStart = position;
            const start = take('reference');
            let end = start;
            if (peek().type === ':') { position++; end = take('reference'); }
            if (peek().type === ')') {
              position++;
              if (start.sheetName && end.sheetName && start.sheetName.toLocaleLowerCase() !== end.sheetName.toLocaleLowerCase()) return false;
              const sheetName = start.sheetName || end.sheetName;
              return !sheetName || context.getSheetIndex?.(sheetName) !== -1;
            }
            position = referenceStart;
          }
          expression(0);
          take(')');
          return false;
        }
        if (token.value === 'SHEET' || token.value === 'SHEETS') {
          if (peek().type === ')') {
            position++;
            if (token.value === 'SHEETS') return context.getSheetCount?.() ?? 1;
            const index = context.getSheetIndex?.();
            return Number.isInteger(index) && index >= 0 ? index + 1 : 1;
          }
          const start = take('reference');
          let end = start;
          if (peek().type === ':') { position++; end = take('reference'); }
          take(')');
          if (start.sheetName && end.sheetName && start.sheetName.toLocaleLowerCase() !== end.sheetName.toLocaleLowerCase()) return '#REF!';
          const sheetName = start.sheetName || end.sheetName;
          const index = sheetName ? context.getSheetIndex?.(sheetName) : context.getSheetIndex?.();
          if (sheetName && index === -1) return '#REF!';
          if (token.value === 'SHEETS') return 1;
          return Number.isInteger(index) && index >= 0 ? index + 1 : 1;
        }
        if (token.value === 'ROW' || token.value === 'COLUMN') {
          if (peek().type === ')') {
            position++;
            const origin = token.value === 'ROW' ? context.formulaRow : context.formulaColumn;
            return Number.isInteger(origin) && origin >= 0 ? origin + 1 : 1;
          }
          const start = take('reference');
          let end = start;
          if (peek().type === ':') { position++; end = take('reference'); }
          take(')');
          if (start.sheetName && end.sheetName && start.sheetName.toLocaleLowerCase() !== end.sheetName.toLocaleLowerCase()) return '#REF!';
          const sheetName = start.sheetName || end.sheetName;
          if (sheetName && context.getSheetIndex?.(sheetName) === -1) return '#REF!';
          return token.value === 'ROW' ? start.row + 1 : start.column + 1;
        }
        const args = [];
        if (peek().type !== ')') {
          do {
            if (peek().type === ',') args.push(null);
            else {
              captureFunctionArguments++;
              try { args.push(expression(0)); }
              finally { captureFunctionArguments--; }
            }
            if (peek().type !== ',') break;
            position++;
            if (peek().type === ')') { args.push(null); break; }
          } while (true);
        }
        take(')');
        return callFunction(token.value, args, context);
      }
      throw new Error('#VALUE!');
    }
    function unary() {
      if (peek().type === '+') { position++; return numeric(expression(5)); }
      if (peek().type === '-') { position++; const value = numeric(expression(5)); return typeof value === 'number' ? -value : value; }
      let value = primary();
      while (peek().type === '%') {
        position++;
        value = numeric(value);
        if (typeof value === 'string') return value;
        value /= 100;
      }
      return value;
    }
    function expression(minimum) {
      let left = unary();
      while (Object.hasOwn(precedence, peek().type) && precedence[peek().type] >= minimum) {
        const operator = tokens[position++].type, level = precedence[operator];
        let right = expression(level + (operator === '^' ? 0 : 1));
        if (isError(left)) return left;
        if (isError(right)) return right;
        if (operator === '&') {
          if (left?.[FORMULA_TEXT_LITERAL] || left?.[FORMULA_REFERENCE_VALUE]) left = left.value;
          if (right?.[FORMULA_TEXT_LITERAL] || right?.[FORMULA_REFERENCE_VALUE]) right = right.value;
          left = String(left ?? '') + String(right ?? '');
          continue;
        }
        if (left?.[FORMULA_TEXT_LITERAL] || left?.[FORMULA_REFERENCE_VALUE]) left = left.value;
        if (right?.[FORMULA_TEXT_LITERAL] || right?.[FORMULA_REFERENCE_VALUE]) right = right.value;
        if (['=', '<>', '<', '>', '<=', '>='].includes(operator)) {
          const a = typeof left === 'string' && typeof right === 'string' ? left.toLocaleLowerCase() : (typeof left === 'string' && typeof right === 'number' ? numeric(left) : left);
          const b = typeof left === 'string' && typeof right === 'string' ? right.toLocaleLowerCase() : (typeof right === 'string' && typeof left === 'number' ? numeric(right) : right);
          if (isError(a) || isError(b)) return isError(a) ? a : b;
          left = operator === '=' ? a === b : operator === '<>' ? a !== b : operator === '<' ? a < b : operator === '>' ? a > b : operator === '<=' ? a <= b : a >= b;
          continue;
        }
        const a = numeric(left), b = numeric(right);
        if (typeof a === 'string' || typeof b === 'string') { left = typeof a === 'string' ? a : b; continue; }
        if (operator === '+') left = a + b;
        else if (operator === '-') left = a - b;
        else if (operator === '*') left = a * b;
        else if (operator === '/') left = b === 0 ? '#DIV/0!' : a / b;
        else if (a === 0 && b < 0) left = '#DIV/0!';
        else { const value = a ** b; left = Number.isFinite(value) ? value : '#NUM!'; }
        if (typeof left === 'number' && !Number.isFinite(left)) left = '#NUM!';
      }
      return left;
    }
    const result = expression(0);
    return peek().type === 'eof' ? (isFormulaRange(result) || Array.isArray(result) && !context.allowArrayResult ? '#VALUE!' : result === null ? 0 : result) : '#VALUE!';
  } catch (error) {
    return /^#/.test(error?.message || '') ? error.message : failValue();
  }
}

function createFormulaCalculator(document) {
  const now = document.now instanceof Date ? new Date(document.now.getTime()) : new Date();
  let filterStyleReader;
  const filterStyle = (sheetIndex, row, column) => {
    filterStyleReader ||= createFilterStyleReader(document, calculator, evaluateFormula);
    return filterStyleReader(row, column, sheetIndex);
  };
  const filterPredicates = new WeakMap();
  const cache = new Map(), active = new Set(), spillValues = new Map(), spillRanges = new Map(), spillExtents = new Map();
  const spillScannedSheets = new Set();
  let steps = 0;
  function sparseRange(bounds, sheetName, defaultSheetIndex = 0) {
    const sheetIndex = sheetName
      ? document.state.sheets.findIndex(item => item.name.toLocaleLowerCase() === sheetName.toLocaleLowerCase())
      : defaultSheetIndex;
    const sheet = document.state.sheets[sheetIndex];
    if (!sheet) return '#REF!';
    const coordinates = new Map();
    const add = (row, column) => {
      if (row >= bounds.top && row <= bounds.bottom && column >= bounds.left && column <= bounds.right) coordinates.set(row + ':' + column, [row, column]);
    };
    for (const cell of sheet.cells.values()) add(cell.row, cell.column);
    const addAddress = (key, draftSheet) => {
      if (Number(draftSheet) !== sheetIndex) return;
      const address = key.slice(key.indexOf(':') + 1), match = /^([A-Z]+)([1-9]\d*)$/.exec(address);
      if (!match) return;
      add(Number(match[2]) - 1, columnNumber(match[1]) - 1);
    };
    for (const key of document.drafts?.keys() || []) addAddress(key, key.slice(0, key.indexOf(':')));
    for (const key of document.formulaDrafts?.keys() || []) addAddress(key, key.slice(0, key.indexOf(':')));
    for (const key of spillValues.keys()) if (key.startsWith(sheetIndex + ':')) addAddress(key, sheetIndex);
    const entries = [];
    for (const [row, column] of [...coordinates.values()].sort((left, right) => left[0] - right[0] || left[1] - right[1])) {
      if (++steps > MAX_CALC_STEPS) return '#NUM!';
      const cell = document.getCell(row, column, sheetIndex), value = cellValue(sheetIndex, row, column);
      if (value !== null && value !== '' || Boolean(cell?.formula || cell?.hasFormula)) entries.push({ row, column, value, formula: cell?.formula || null });
    }
    return { [FORMULA_RANGE]: true, bounds, entries, cellCount: (bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1), sheetIndex };
  }
  function isFilteredRow(sheetIndex, row) {
    const sheet = document.state.sheets[sheetIndex];
    const filters = [sheet?.autoFilter, ...(sheet?.tables || []).map(table => table.autoFilter)].filter(filter => filter?.columns.length);
    return filters.some(filter => row > filter.top && row <= filter.bottom && filter.columns.some(columnFilter => {
      const value = cellValue(sheetIndex, row, columnFilter.column);
      if (!filterPredicates.has(columnFilter)) {
        // A subtotal inside the ranked region can ask about visibility again.
        // Install a provisional predicate while the population is evaluated.
        filterPredicates.set(columnFilter, () => true);
        filterPredicates.set(columnFilter, compileAutoFilterColumn(columnFilter, () => filterColumnCells(filter, columnFilter.column, (r, c) => ({ value: cellValue(sheetIndex, r, c) })), { date1904: document.state.date1904, now, getStyle: cell => filterStyle(sheetIndex, cell.row, cell.column) }));
      }
      return !filterPredicates.get(columnFilter)({ value, row, column: columnFilter.column });
    }));
  }
  function isManuallyHiddenRow(sheetIndex, row) {
    return document.state.sheets[sheetIndex]?.hiddenRows?.includes(row) || false;
  }
  function scanSpills(sheetIndex) {
    if (spillScannedSheets.has(sheetIndex)) return;
    spillScannedSheets.add(sheetIndex);
    if (document.state?.formulas === false && !document.formulaDrafts?.size) return;
    const producesArray = formula => typeof formula === 'string' && (/\b(?:LET|SEQUENCE|FILTER|SORT|SORTBY|UNIQUE|TRANSPOSE|TEXTSPLIT|INDIRECT|OFFSET|CHOOSECOLS|CHOOSEROWS|TAKE|DROP|TOCOL|TOROW|WRAPROWS|WRAPCOLS|HSTACK|VSTACK|MODE\.MULT|FREQUENCY|RANDARRAY)\s*\(/i.test(formula) || document.state.definedNames?.some(item => new RegExp(`(?<![A-Z0-9_.])${item.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(?![A-Z0-9_.])`, 'i').test(formula)));
    const referencesSpill = formula => typeof formula === 'string' && /(?:'[^']+'|[A-Za-z_][\w. ]*)!\$?[A-Z]{1,3}\$?[1-9]\d*#|\$?[A-Z]{1,3}\$?[1-9]\d*#/i.test(formula);
    const formulas = [...(document.state?.sheets?.[sheetIndex]?.cells?.values?.() || [])]
      .filter(cell => cell.formula)
      .map(cell => ({ row: cell.row, column: cell.column, formula: cell.formula }));
    const prefix = sheetIndex + ':';
    for (const [key, formula] of document.formulaDrafts || []) {
      const address = key.startsWith(prefix) ? key.slice(prefix.length) : '';
      const match = /^([A-Z]+)([1-9]\d*)$/i.exec(address);
      if (!match) continue;
      let column = 0;
      for (const letter of match[1].toUpperCase()) column = column * 26 + letter.charCodeAt(0) - 64;
      formulas.push({ row: Number(match[2]) - 1, column: column - 1, formula });
    }
    // Calculate array-producing anchors first so A1# consumers can resolve their ranges.
    for (const item of formulas) if (producesArray(item.formula)) cellValue(sheetIndex, item.row, item.column);
    for (const item of formulas) if (referencesSpill(item.formula)) cellValue(sheetIndex, item.row, item.column);
  }
  function cellValue(sheetIndex, row, column) {
    const sheet = document.state.sheets[sheetIndex];
    if (!sheet || row < 0 || column < 0) return '#REF!';
    const address = addressFor(row + 1, column + 1), key = sheetIndex + ':' + address;
    if (spillValues.has(key)) return spillValues.get(key);
    const cell = document.getCell(row, column, sheetIndex);
    if (!cell) return null;
    if (!cell.formula) return scalarCellValue(cell);
    if (cache.has(key)) return cache.get(key);
    if (active.has(key) || active.size > 1000) return '#CIRC!';
    active.add(key);
    const evaluatedValue = evaluateFormula(cell.formula, { now, date1904: document.state.date1904, formulaRow: row, formulaColumn: column, allowArrayResult: true, getCell: (nextRow, nextColumn, sheetName) => {
      if (++steps > MAX_CALC_STEPS) return '#NUM!';
      if (!sheetName) return cellValue(sheetIndex, nextRow, nextColumn);
      const targetSheet = document.state.sheets.findIndex(item => item.name.toLocaleLowerCase() === sheetName.toLocaleLowerCase());
      return targetSheet < 0 ? '#REF!' : cellValue(targetSheet, nextRow, nextColumn);
    }, getCellFormula: (nextRow, nextColumn, sheetName) => {
      const targetSheet = sheetName ? document.state.sheets.findIndex(item => item.name.toLocaleLowerCase() === sheetName.toLocaleLowerCase()) : sheetIndex;
      return targetSheet < 0 ? null : document.getCell(nextRow, nextColumn, targetSheet)?.formula || null;
    }, getSheetIndex: sheetName => sheetName ? document.state.sheets.findIndex(item => item.name.toLocaleLowerCase() === sheetName.toLocaleLowerCase()) : sheetIndex,
    getDefinedName: name => {
      const normalized = name.toLocaleLowerCase();
      const definition = document.state.definedNames?.find(item => item.localSheetId === sheetIndex && item.name.toLocaleLowerCase() === normalized)
        || document.state.definedNames?.find(item => item.localSheetId === null && item.name.toLocaleLowerCase() === normalized);
      if (!definition?.source) return definition?.error || '#NAME?';
      const targetSheet = document.state.sheets.findIndex(item => item.name.toLocaleLowerCase() === definition.source.sheetName.toLocaleLowerCase());
      if (targetSheet < 0) return '#REF!';
      const { top, bottom, left, right } = definition.source, height = bottom - top + 1, width = right - left + 1;
      if (height < 1 || width < 1 || height * width > MAX_RANGE_CELLS) return '#REF!';
      const values = Array.from({ length: height }, (_, row) => Array.from({ length: width }, (_, column) => cellValue(targetSheet, top + row, left + column)));
      return height === 1 && width === 1 ? values[0][0] : values;
    },
    getSheetCount: () => document.state.sheets.length,
    getSparseRange: (bounds, sheetName) => sparseRange(bounds, sheetName, sheetIndex),
    getSpillRange: (spillRow, spillColumn, sheetName) => {
      const targetSheetIndex = sheetName ? document.state.sheets.findIndex(item => item.name.toLocaleLowerCase() === sheetName.toLocaleLowerCase()) : sheetIndex;
      if (targetSheetIndex < 0) return '#REF!';
      scanSpills(targetSheetIndex);
      const key = targetSheetIndex + ':' + addressFor(spillRow + 1, spillColumn + 1), bounds = spillRanges.get(key);
      return bounds ? matrixValue(sparseRange(bounds, document.state.sheets[targetSheetIndex].name, targetSheetIndex)) ?? '#REF!' : '#REF!';
    },
    getTableReference: (tableName, parts, formulaRow, formulaColumn) => {
      let targetSheetIndex = sheetIndex, table;
      if (tableName) {
        for (let candidate = 0; candidate < document.state.sheets.length && !table; candidate++) {
          const found = document.state.sheets[candidate].tables?.find(item => item.name.toLocaleLowerCase() === tableName.toLocaleLowerCase());
          if (found) { targetSheetIndex = candidate; table = found; }
        }
      } else table = document.state.sheets[targetSheetIndex]?.tables?.find(item => formulaRow >= item.range.top + (item.headerRowCount ?? 1) && formulaRow <= item.range.bottom && formulaColumn >= item.range.left && formulaColumn <= item.range.right);
      if (!table) return tableName ? '#NAME?' : '#VALUE!';
      const column = parts.columnName ? table.columns.findIndex(name => name.toLocaleLowerCase() === parts.columnName.toLocaleLowerCase()) : -1;
      const endColumn = parts.columnEndName ? table.columns.findIndex(name => name.toLocaleLowerCase() === parts.columnEndName.toLocaleLowerCase()) : column;
      if (parts.columnName && column < 0 || parts.columnEndName && endColumn < 0) return '#VALUE!';
      if (parts.columnEndName && endColumn < column) return '#VALUE!';
      const { top, bottom, left, right } = table.range, headerRows = table.headerRowCount ?? 1, totalsRows = table.totalsRowCount ?? 0;
      const dataTop = top + headerRows, dataBottom = bottom - totalsRows;
      if (parts.thisRow || parts.selector === '#thisrow') {
        if (targetSheetIndex !== sheetIndex || formulaRow < dataTop || formulaRow > dataBottom || column < 0) return '#VALUE!';
        if (endColumn === column) return cellValue(targetSheetIndex, formulaRow, left + column);
        return sparseRange({ top: formulaRow, bottom: formulaRow, left: left + column, right: left + endColumn }, document.state.sheets[targetSheetIndex].name, targetSheetIndex);
      }
      let rowTop = dataTop, rowBottom = dataBottom;
      if (parts.selector === '#headers') { if (!headerRows) return '#VALUE!'; rowTop = top; rowBottom = top + headerRows - 1; }
      else if (parts.selector === '#totals') { if (!totalsRows) return '#VALUE!'; rowTop = bottom - totalsRows + 1; rowBottom = bottom; }
      else if (parts.selector === '#all') { rowTop = top; rowBottom = bottom; }
      // An empty data body is a zero-cell range; headers and totals remain addressable.
      if (rowBottom < rowTop && rowBottom !== rowTop - 1) return '#VALUE!';
      return sparseRange({ top: rowTop, bottom: rowBottom, left: column < 0 ? left : left + column, right: column < 0 ? right : left + endColumn }, document.state.sheets[targetSheetIndex].name, targetSheetIndex);
    }, isFilteredRow, isManuallyHiddenRow });
    active.delete(key);
    let value = evaluatedValue;
    if (Array.isArray(evaluatedValue)) {
      const height = evaluatedValue.length, width = Array.isArray(evaluatedValue[0]) ? evaluatedValue[0].length : 0;
      let blocked = !height || !width || height * width > MAX_RANGE_CELLS || evaluatedValue.some(line => !Array.isArray(line) || line.length !== width) || row + height > 1048576 || column + width > 16384;
      if (!blocked) {
        const spillRange = { top: row, bottom: row + height - 1, left: column, right: column + width - 1 };
        blocked = (sheet.mergedCells || []).some(range => spillRange.top <= range.bottom && spillRange.bottom >= range.top && spillRange.left <= range.right && spillRange.right >= range.left)
          || (sheet.tables || []).some(table => spillRange.top <= table.range.bottom && spillRange.bottom >= table.range.top && spillRange.left <= table.range.right && spillRange.right >= table.range.left);
        for (let spillRow = row; !blocked && spillRow <= spillRange.bottom; spillRow++) for (let spillColumn = column; !blocked && spillColumn <= spillRange.right; spillColumn++) {
          if (spillRow === row && spillColumn === column) continue;
          if (spillValues.has(sheetIndex + ':' + addressFor(spillRow + 1, spillColumn + 1))) { blocked = true; break; }
          const occupied = document.getCell(spillRow, spillColumn, sheetIndex);
          if (occupied && (occupied.value !== '' && occupied.value !== null && occupied.value !== undefined || occupied.formula || occupied.hasFormula)) blocked = true;
        }
        if (!blocked) {
          spillRanges.set(key, spillRange);
          for (let spillRow = 0; spillRow < height; spillRow++) for (let spillColumn = 0; spillColumn < width; spillColumn++) {
            if (!spillRow && !spillColumn) continue;
            spillValues.set(sheetIndex + ':' + addressFor(row + spillRow + 1, column + spillColumn + 1), evaluatedValue[spillRow][spillColumn]);
          }
          const extent = spillExtents.get(sheetIndex) || { bottom: row, right: column };
          extent.bottom = Math.max(extent.bottom, row + height - 1); extent.right = Math.max(extent.right, column + width - 1);
          spillExtents.set(sheetIndex, extent);
        }
      }
      value = blocked ? '#SPILL!' : evaluatedValue[0][0];
    }
    cache.set(key, value);
    return value;
  }
  const calculator = { now, filterStyle, value: cellValue, spillRanges: sheetIndex => { scanSpills(sheetIndex); return [...spillRanges].filter(([key]) => key.startsWith(sheetIndex + ':')).map(([, range]) => ({ ...range })); }, spillExtent: sheetIndex => { scanSpills(sheetIndex); return spillExtents.get(sheetIndex) || null; }, cell: (sheetIndex, row, column) => {
    scanSpills(sheetIndex);
    const cell = document.getCell(row, column, sheetIndex);
    if (!cell?.formula) {
      const key = sheetIndex + ':' + addressFor(row + 1, column + 1);
      if (!spillValues.has(key)) return cell;
      const value = spillValues.get(key);
      return { ...cell, row, column, address: addressFor(row + 1, column + 1), value: typeof value === 'boolean' ? String(value).toUpperCase() : String(value ?? ''), cachedValue: value, kind: typeof value === 'number' ? 'number' : typeof value === 'boolean' ? 'boolean' : isError(value) ? 'error' : 'string', isSpill: true, ...(document.state.date1904 ? { date1904: true } : {}) };
    }
    const value = cellValue(sheetIndex, row, column);
    const kind = typeof value === 'number' ? 'number' : typeof value === 'boolean' ? 'boolean' : isError(value) ? 'error' : 'string';
    return { ...cell, value: kind === 'boolean' ? (value ? 'TRUE' : 'FALSE') : value === null ? '' : String(value), cachedValue: value, kind };
  } };
  return calculator;
}

function scalarCellValue(cell) {
  if (!cell || cell.kind === 'blank') return null;
  if (cell.kind === 'number') { const value = Number(cell.value); return Number.isFinite(value) ? value : '#VALUE!'; }
  if (cell.kind === 'boolean') return String(cell.value).toUpperCase() === 'TRUE';
  if (cell.kind === 'error') return cell.value;
  return cell.value;
}

function addressFor(row, column) {
  let letters = '';
  for (let value = column; value > 0; value = Math.floor((value - 1) / 26)) letters = String.fromCharCode(65 + (value - 1) % 26) + letters;
  return letters + row;
}

const DAY_MILLISECONDS = 86400000;
const EXCEL_EPOCH = Date.UTC(1899, 11, 31);
function utcDate(year, month, day) {
  const result = new Date(0);
  result.setUTCHours(0, 0, 0, 0);
  result.setUTCFullYear(year, month - 1, day);
  return result;
}
function dateToSerial(date) {
  const days = Math.floor((date.getTime() - EXCEL_EPOCH) / DAY_MILLISECONDS);
  const serial = days >= 60 ? days + 1 : days;
  return serial < 0 || serial >= 2958466 ? '#NUM!' : serial;
}
function dateParts(serial) {
  if (serial < 0 || serial >= 2958466) return '#NUM!';
  const whole = Math.floor(serial);
  if (whole === 60) return { year: 1900, month: 2, day: 29 };
  const adjusted = whole > 60 ? whole - 1 : whole;
  const date = new Date(EXCEL_EPOCH + adjusted * DAY_MILLISECONDS);
  if (!Number.isFinite(date.getTime())) return '#NUM!';
  return { year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: date.getUTCDate() };
}
function parseDateValue(input) {
  const text = input.trim().replace(/[T ](?=\d{1,2}:\d{2}(?::|$))[\s\S]*$/i, '');
  let match = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/.exec(text);
  if (!match) match = /^(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일?$/.exec(text);
  if (match) {
    const [, year, month, day] = match;
    return validDateValue(Number(year), Number(month), Number(day));
  }
  match = /^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/.exec(text);
  if (match) {
    const [, month, day, rawYear] = match;
    const year = rawYear.length === 2 ? Number(rawYear) + (Number(rawYear) < 30 ? 2000 : 1900) : Number(rawYear);
    return validDateValue(year, Number(month), Number(day));
  }
  const monthNames = new Map(['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'].map((name, index) => [name, index + 1]));
  match = /^(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+(\d{1,2}),?\s+(\d{2,4})$/i.exec(text);
  if (!match) match = /^(\d{1,2})\s+(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+(\d{2,4})$/i.exec(text);
  if (!match) return null;
  const monthName = (match[1].match(/[A-Za-z]+/)?.[0] || match[2]).slice(0, 3).toLowerCase();
  const day = Number(/^\d/.test(match[1]) ? match[1] : match[2]);
  const rawYear = /^\d/.test(match[1]) ? match[3] : match[3];
  const year = rawYear.length === 2 ? Number(rawYear) + (Number(rawYear) < 30 ? 2000 : 1900) : Number(rawYear);
  return validDateValue(year, monthNames.get(monthName), day);
}
function validDateValue(year, month, day) {
  if (![year, month, day].every(Number.isInteger) || year < 0 || year > 9999 || month < 1 || month > 12 || day < 1 || day > 31) return null;
  if (year === 1900 && month === 2 && day === 29) return { year, month, day };
  const date = utcDate(year, month, day);
  if (date.getUTCFullYear() !== year || date.getUTCMonth() + 1 !== month || date.getUTCDate() !== day) return null;
  return dateToSerial(date) >= 0 ? { year, month, day } : null;
}
function parseTimeValue(input) {
  let text = input.trim();
  const koreanPeriod = /^(오전|오후)\s+(.+)$/i.exec(text);
  if (koreanPeriod) text = `${koreanPeriod[2]} ${koreanPeriod[1]}`;
  const match = /(?:^|[T\s])(\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?\s*(AM|PM|오전|오후)?\s*Z?$/i.exec(text);
  if (!match) return null;
  let hour = Number(match[1]);
  const minute = Number(match[2]), second = Number(match[3] || 0);
  const fraction = match[4] ? Number(`0.${match[4]}`) : 0;
  const period = match[5]?.toUpperCase();
  if (minute > 59 || second > 59 || fraction >= 1) return null;
  if (period) {
    if (hour < 1 || hour > 12) return null;
    if (period === 'PM' || period === '오후') hour = hour % 12 + 12;
    else hour %= 12;
  } else if (hour > 23) return null;
  const total = hour * 3600 + minute * 60 + second + fraction;
  return total >= 0 && total < 86400 ? total : null;
}
function couponBondPrice(settlement, maturity, annualRate, yieldRate, redemption, frequency, basis) {
  const schedule = couponPeriodDates(Math.floor(settlement), Math.floor(maturity), frequency);
  if (typeof schedule === 'string') return schedule;
  const couponDays = callFunction('COUPDAYS', [settlement, maturity, frequency, basis]);
  const accruedDays = callFunction('COUPDAYBS', [settlement, maturity, frequency, basis]);
  const daysToNext = callFunction('COUPDAYSNC', [settlement, maturity, frequency, basis]);
  const couponCount = callFunction('COUPNUM', [settlement, maturity, frequency, basis]);
  for (const value of [couponDays, accruedDays, daysToNext, couponCount]) if (typeof value !== 'number') return value;
  const periodicCoupon = 100 * annualRate / frequency, periodicYield = yieldRate / frequency;
  const discountBase = 1 + periodicYield;
  if (discountBase <= 0 || couponDays <= 0) return '#NUM!';
  const firstPeriodFraction = daysToNext / couponDays;
  let price = -periodicCoupon * accruedDays / couponDays;
  for (let period = 0; period < couponCount; period++) {
    const discountedCoupon = periodicCoupon / discountBase ** (firstPeriodFraction + period);
    price += discountedCoupon;
    if (!Number.isFinite(price)) return '#NUM!';
  }
  const discountedRedemption = redemption / discountBase ** (firstPeriodFraction + couponCount - 1);
  price += discountedRedemption;
  return Number.isFinite(price) ? price : '#NUM!';
}
function couponPeriodDates(settlement, maturity, frequency) {
  const end = dateParts(maturity);
  if (typeof end === 'string') return end;
  const monthsPerPeriod = 12 / frequency;
  const endOfMonth = end.day === (end.year === 1900 && end.month === 2 ? 29 : new Date(utcDate(end.year, end.month + 1, 0)).getUTCDate());
  const couponAt = index => {
    const monthIndex = end.year * 12 + end.month - 1 - index * monthsPerPeriod;
    const year = Math.floor(monthIndex / 12), month = monthIndex % 12 + 1;
    const lastDay = year === 1900 && month === 2 ? 29 : new Date(utcDate(year, month + 1, 0)).getUTCDate();
    const day = endOfMonth ? lastDay : Math.min(end.day, lastDay);
    return dateToSerial(utcDate(year, month, day));
  };
  let next = null;
  for (let index = 0; index < 120000; index++) {
    const coupon = couponAt(index);
    if (typeof coupon !== 'number') return coupon;
    if (coupon > settlement) { next = coupon; continue; }
    if (coupon === settlement) return { previous: coupon, next, periodIndex: index - 1 };
    if (next === null) return '#NUM!';
    return { previous: coupon, next, periodIndex: index - 1 };
  }
  return '#NUM!';
}

module.exports = { evaluateFormula, createFormulaCalculator };
