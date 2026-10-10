'use strict';

const path = require('node:path').posix;
const { unzipSync, zipSync } = require('../vendor/node_modules/fflate');
const { SaxesParser } = require('../vendor/node_modules/saxes');

const LIMITS = Object.freeze({ bytes: 20 * 1024 * 1024, entries: 512, expanded: 64 * 1024 * 1024, entry: 16 * 1024 * 1024, ratio: 200, sheets: 100, rows: 1000000, columns: 16384, cells: 100000, xmlNodes: 500000 });
const CRC_TABLE = new Uint32Array(256);
for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; CRC_TABLE[n] = c >>> 0; }

function crc32(bytes) { let crc = 0xffffffff; for (const byte of bytes) crc = CRC_TABLE[(crc ^ byte) & 0xff] ^ (crc >>> 8); return (crc ^ 0xffffffff) >>> 0; }

function fail(message) { throw new Error(message); }

function inspectZip(input, limits = LIMITS) {
  const bytes = Buffer.from(input);
  if (bytes.length < 22 || bytes.length > limits.bytes) fail('XLSX ZIP 크기 제한을 벗어났습니다.');
  let eocd = -1;
  const lower = Math.max(0, bytes.length - 65557);
  for (let i = bytes.length - 22; i >= lower; i--) if (bytes.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  if (eocd < 0) fail('ZIP 종료 레코드를 찾을 수 없습니다.');
  const disk = bytes.readUInt16LE(eocd + 4), dirDisk = bytes.readUInt16LE(eocd + 6);
  const onDisk = bytes.readUInt16LE(eocd + 8), count = bytes.readUInt16LE(eocd + 10);
  const directorySize = bytes.readUInt32LE(eocd + 12), directoryOffset = bytes.readUInt32LE(eocd + 16);
  const commentLength = bytes.readUInt16LE(eocd + 20);
  if (disk || dirDisk || onDisk !== count || count === 0xffff || directoryOffset === 0xffffffff || directorySize === 0xffffffff) fail('다중 디스크 ZIP 또는 ZIP64는 지원하지 않습니다.');
  if (count < 1 || count > limits.entries || directoryOffset + directorySize !== eocd || eocd + 22 + commentLength !== bytes.length) fail('ZIP 중앙 디렉터리가 유효하지 않거나 엔트리 제한을 초과했습니다.');
  const records = new Map(); let cursor = directoryOffset, total = 0;
  for (let i = 0; i < count; i++) {
    if (cursor + 46 > eocd || bytes.readUInt32LE(cursor) !== 0x02014b50) fail('ZIP 중앙 디렉터리 항목이 손상되었습니다.');
    const madeBy = bytes.readUInt16LE(cursor + 4), flags = bytes.readUInt16LE(cursor + 8), method = bytes.readUInt16LE(cursor + 10);
    const expectedCrc = bytes.readUInt32LE(cursor + 16), compressed = bytes.readUInt32LE(cursor + 20), expanded = bytes.readUInt32LE(cursor + 24);
    const nameLength = bytes.readUInt16LE(cursor + 28), extraLength = bytes.readUInt16LE(cursor + 30), noteLength = bytes.readUInt16LE(cursor + 32);
    const diskStart = bytes.readUInt16LE(cursor + 34), externalAttrs = bytes.readUInt32LE(cursor + 38), localOffset = bytes.readUInt32LE(cursor + 42);
    const recordLength = 46 + nameLength + extraLength + noteLength;
    if (cursor + recordLength > eocd || diskStart !== 0 || flags & 1 || (flags & ~0x080e) || (method !== 0 && method !== 8)) fail('암호화·분할 또는 지원하지 않는 압축 항목을 거부했습니다.');
    if (compressed === 0xffffffff || expanded === 0xffffffff || localOffset === 0xffffffff) fail('ZIP64 엔트리는 지원하지 않습니다.');
    let name;
    try { name = flags & 0x800 ? new TextDecoder('utf-8', { fatal: true }).decode(bytes.subarray(cursor + 46, cursor + 46 + nameLength)) : bytes.subarray(cursor + 46, cursor + 46 + nameLength).toString('latin1'); }
    catch { fail('ZIP 항목 이름 인코딩이 유효하지 않습니다.'); }
    if (!name || name.includes('\\') || name.startsWith('/') || name.split('/').includes('..') || name.includes('\0')) fail('안전하지 않은 ZIP 항목 경로입니다.');
    if (records.has(name)) fail(`중복 ZIP 항목을 거부했습니다: ${name}`);
    let extraOffset = cursor + 46 + nameLength, extraEnd = extraOffset + extraLength;
    while (extraOffset < extraEnd) {
      if (extraOffset + 4 > extraEnd) fail(`ZIP 확장 필드가 손상되었습니다: ${name}`);
      const extraId = bytes.readUInt16LE(extraOffset), extraSize = bytes.readUInt16LE(extraOffset + 2);
      if (extraOffset + 4 + extraSize > extraEnd || [0x0001, 0x7075, 0x6375].includes(extraId)) fail(`지원하지 않거나 손상된 ZIP 확장 필드입니다: ${name}`);
      extraOffset += 4 + extraSize;
    }
    if (expanded > limits.entry || (compressed === 0 ? expanded > 0 : expanded / compressed > limits.ratio)) fail(`압축 해제 크기 제한을 초과했습니다: ${name}`);
    total += expanded;
    if (total > limits.expanded) fail('압축 해제된 전체 패키지 크기 제한을 초과했습니다.');
    if (localOffset + 30 > directoryOffset || bytes.readUInt32LE(localOffset) !== 0x04034b50) fail(`잘못된 로컬 ZIP 헤더입니다: ${name}`);
    const localFlags = bytes.readUInt16LE(localOffset + 6), localMethod = bytes.readUInt16LE(localOffset + 8);
    const localNameLength = bytes.readUInt16LE(localOffset + 26), localExtraLength = bytes.readUInt16LE(localOffset + 28);
    const localName = bytes.subarray(localOffset + 30, localOffset + 30 + localNameLength).toString(flags & 0x800 ? 'utf8' : 'latin1');
    if (localName !== name || localFlags !== flags || localMethod !== method) fail(`ZIP 로컬 항목과 중앙 디렉터리가 일치하지 않습니다: ${name}`);
    const dataOffset = localOffset + 30 + localNameLength + localExtraLength;
    const dataEnd = dataOffset + compressed;
    if (dataEnd > directoryOffset) fail(`ZIP 항목 데이터 범위가 유효하지 않습니다: ${name}`);
    if ((madeBy >>> 8) === 3 && ((externalAttrs >>> 16) & 0xf000) === 0xa000) fail('심볼릭 링크 ZIP 항목은 지원하지 않습니다.');
    if (method === 0 && compressed !== expanded) fail(`저장 항목 길이가 일치하지 않습니다: ${name}`);
    if (!(flags & 8) && (bytes.readUInt32LE(localOffset + 14) !== expectedCrc || bytes.readUInt32LE(localOffset + 18) !== compressed || bytes.readUInt32LE(localOffset + 22) !== expanded)) fail(`ZIP 로컬 크기·CRC가 중앙 디렉터리와 다릅니다: ${name}`);
    let entryEnd = dataEnd;
    if (flags & 8) {
      const signature = dataEnd + 4 <= directoryOffset && bytes.readUInt32LE(dataEnd) === 0x08074b50;
      const descriptor = dataEnd + (signature ? 4 : 0);
      if (descriptor + 12 > directoryOffset || bytes.readUInt32LE(descriptor) !== expectedCrc || bytes.readUInt32LE(descriptor + 4) !== compressed || bytes.readUInt32LE(descriptor + 8) !== expanded) fail(`ZIP 데이터 설명자가 유효하지 않습니다: ${name}`);
      entryEnd = descriptor + 12;
    }
    records.set(name, { compressed, expanded, expectedCrc, method, localOffset, dataOffset, entryEnd });
    cursor += recordLength;
  }
  if (cursor !== eocd) fail('ZIP 중앙 디렉터리 길이가 맞지 않습니다.');
  const ranges = [...records.entries()].map(([name, item]) => ({ name, start: item.localOffset, end: item.entryEnd })).sort((a, b) => a.start - b.start);
  for (let i = 1; i < ranges.length; i++) if (ranges[i].start < ranges[i - 1].end) fail(`ZIP 항목 데이터가 겹칩니다: ${ranges[i].name}`);
  return { bytes, records };
}

function parseXml(bytes, partName) {
  if (!bytes || bytes.length > LIMITS.entry) fail(`XML 항목이 없거나 너무 큽니다: ${partName}`);
  let xml;
  try { xml = new TextDecoder('utf-8', { fatal: true }).decode(bytes); }
  catch { fail(`UTF-8이 아닌 XML은 지원하지 않습니다: ${partName}`); }
  if (/<!ENTITY\b/i.test(xml)) fail(`XML 엔터티 선언을 거부했습니다: ${partName}`);
  const roots = [], stack = []; let nodeCount = 0, parseError;
  const parser = new SaxesParser({ xmlns: false, fragment: false });
  parser.on('doctype', () => { parseError = new Error(`DTD를 거부했습니다: ${partName}`); });
  parser.on('error', error => { parseError = error; });
  parser.on('opentag', tag => {
    if (stack.length >= 100 || ++nodeCount > LIMITS.xmlNodes) { parseError = new Error(`XML 구조 제한을 초과했습니다: ${partName}`); return; }
    const node = { name: tag.name, attributes: tag.attributes, children: [], text: '' };
    if (stack.length) stack[stack.length - 1].children.push(node); else roots.push(node);
    stack.push(node);
  });
  parser.on('closetag', () => { stack.pop(); });
  parser.on('text', text => { if (stack.length) stack[stack.length - 1].text += text; });
  parser.on('cdata', () => { parseError = new Error(`CDATA가 포함된 XML을 거부했습니다: ${partName}`); });
  try { parser.write(xml).close(); } catch (error) { parseError ||= error; }
  if (parseError) fail(`안전하지 않거나 손상된 XML (${partName}): ${parseError.message}`);
  if (roots.length !== 1) fail(`XML 루트가 하나가 아닙니다: ${partName}`);
  return { xml, root: roots[0] };
}

function children(node, name) { return node.children.filter(child => child.name === name); }
function child(node, name) { return node.children.find(item => item.name === name); }
function textContent(node) { return node.text + node.children.map(textContent).join(''); }


function readFreezePane(root) {
  const sheetViews = child(root, 'sheetViews');
  const empty = { freezePane: { rows: 0, columns: 0 }, reason: null };
  if (!sheetViews) return empty;
  const views = children(sheetViews, 'sheetView');
  if (views.length !== 1) return { freezePane: { rows: 0, columns: 0 }, reason: '여러 사용자 정의 시트 보기를 안전하게 해석할 수 없습니다.' };
  const panes = children(views[0], 'pane');
  if (!panes.length) return empty;
  if (panes.length !== 1) return { freezePane: { rows: 0, columns: 0 }, reason: '여러 분할 창을 안전하게 해석할 수 없습니다.' };
  const attributes = panes[0].attributes;
  const xText = attributes.xSplit ?? '0', yText = attributes.ySplit ?? '0';
  if (!/^(?:0|[1-9]\d*)$/.test(xText) || !/^(?:0|[1-9]\d*)$/.test(yText)) return { freezePane: { rows: 0, columns: 0 }, reason: '소수 또는 잘못된 고정 창 범위는 지원하지 않습니다.' };
  const columns = Number(xText), rows = Number(yText);
  if (!columns && !rows) return empty;
  if (attributes.state !== 'frozen') return { freezePane: { rows: 0, columns: 0 }, reason: '분할 창은 고정 창과 동작이 달라 저장 보호를 위해 읽기 전용으로 엽니다.' };
  if (columns > 10 || rows > 50) return { freezePane: { rows: 0, columns: 0 }, reason: '고정 창이 화면 렌더링 한도(행 50개·열 10개)를 초과했습니다.' };
  if (attributes.topLeftCell) {
    const topLeftColumn = columnNumber(attributes.topLeftCell), topLeftRow = rowNumber(attributes.topLeftCell);
    if (!topLeftColumn || !topLeftRow || topLeftColumn > LIMITS.columns || topLeftRow > LIMITS.rows || topLeftColumn <= columns || topLeftRow <= rows) return { freezePane: { rows: 0, columns: 0 }, reason: '고정 창 시작 셀이 시트 범위 또는 고정 영역과 일치하지 않습니다.' };
  }
  return { freezePane: { rows, columns }, reason: null };
}

function readColumnWidths(root) {
  const definitions = children(child(root, 'cols') || { children: [] }, 'col');
  if (definitions.length > LIMITS.columns) fail('워크시트 열 너비 정의 개수 제한을 초과했습니다.');
  const widths = {};
  let expandedColumns = 0;
  for (const definition of definitions) {
    const { min: minText, max: maxText, width: widthText } = definition.attributes;
    if (!/^[1-9]\d*$/.test(minText || '') || !/^[1-9]\d*$/.test(maxText || '')) fail('워크시트 열 너비 범위가 유효하지 않습니다.');
    const min = Number(minText), max = Number(maxText);
    if (min > max || max > LIMITS.columns) fail('워크시트 열 너비 범위가 제한을 벗어났습니다.');
    if (widthText === undefined) continue;
    if (!/^\+?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(widthText)) fail('워크시트 열 너비 값이 유효하지 않습니다.');
    const width = Number(widthText);
    if (!Number.isFinite(width) || width < 0 || width > 255) fail('워크시트 열 너비 값이 유효하지 않습니다.');
    expandedColumns += max - min + 1;
    if (expandedColumns > LIMITS.columns) fail('워크시트 열 너비 확장 제한을 초과했습니다.');
    for (let column = min; column <= max; column++) widths[column - 1] = width;
  }
  return widths;
}
function xmlEscape(value) { return String(value).replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&apos;'); }

function columnNumber(address) {
  const letters = /^([A-Z]+)([1-9]\d*)$/.exec(address)?.[1];
  if (!letters) return 0;
  let n = 0; for (const char of letters) n = n * 26 + char.charCodeAt(0) - 64;
  return n;
}
function rowNumber(address) { const match = /^[A-Z]+([1-9]\d*)$/.exec(address); return match ? Number(match[1]) : 0; }
function decodeCellValue(cell, sharedStrings) {
  const type = cell.attributes.t || 'n';
  if (type === 'inlineStr') return { value: textContent(child(cell, 'is') || { text: '', children: [] }), kind: 'string' };
  const raw = child(cell, 'v')?.text ?? '';
  if (type === 's') {
    const index = Number(raw);
    if (!Number.isInteger(index) || index < 0 || index >= sharedStrings.length) fail('공유 문자열 인덱스가 유효하지 않습니다.');
    return { value: sharedStrings[index], kind: 'string' };
  }
  if (type === 'b') return { value: raw === '1' ? 'TRUE' : 'FALSE', kind: 'boolean' };
  if (type === 'e') return { value: raw, kind: 'error' };
  if (type === 'str') return { value: raw, kind: 'string' };
  if (type === 'n') return { value: raw, kind: 'number' };
  fail(`지원하지 않는 셀 유형입니다: ${type}`);
}

const BUILTIN_NUMBER_FORMATS = Object.freeze({
  1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
  5: '$#,##0', 6: '$#,##0', 7: '$#,##0.00', 8: '$#,##0.00',
  9: '0%', 10: '0.00%', 14: 'm/d/yy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy',
  18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yy h:mm',
  37: '#,##0;(#,##0)', 38: '#,##0;(#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;(#,##0.00)'
});
const DATE_FORMAT_IDS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22]);

function styleColor(node) {
  const value = node?.attributes?.rgb;
  if (typeof value !== 'string' || !/^(?:FF)?[0-9A-F]{6}$/i.test(value)) return null;
  return `#${value.slice(-6)}`;
}

function readCellStyleProperties(xf, fonts, fills, borders) {
  const properties = {};
  const font = fonts[Number(xf.attributes.fontId) || 0];
  const fill = fills[Number(xf.attributes.fillId) || 0];
  const border = borders[Number(xf.attributes.borderId) || 0];
  const alignment = child(xf, 'alignment');
  if (font?.bold) properties.bold = true;
  if (font?.italic) properties.italic = true;
  if (font?.underline) properties.underline = true;
  if (font?.strike) properties.strike = true;
  if (font?.color) properties.color = font.color;
  if (font?.size) properties.fontSize = font.size;
  if (fill?.color) properties.fill = fill.color;
  const horizontal = alignment?.attributes?.horizontal;
  if (['left', 'center', 'right', 'justify'].includes(horizontal)) properties.horizontal = horizontal;
  const vertical = alignment?.attributes?.vertical;
  if (['top', 'center', 'bottom'].includes(vertical)) properties.vertical = vertical;
  if (border) {
    const sides = {};
    for (const side of ['left', 'right', 'top', 'bottom']) {
      const edge = child(border, side);
      const color = styleColor(child(edge || { children: [] }, 'color'));
      const line = edge?.attributes?.style;
      if (['hair', 'thin', 'medium', 'thick', 'dashed', 'dotted', 'double'].includes(line)) {
        sides[side] = { color: color || '#000000', line };
      }
    }
    if (Object.keys(sides).length) properties.borders = sides;
  }
  return Object.keys(properties).length ? properties : null;
}
function tokenizeDateFormat(format) {
  const tokens = [];
  for (let i = 0; i < format.length;) {
    if (format[i] === '"') {
      let end = i + 1, literal = '';
      while (end < format.length && format[end] !== '"') literal += format[end++];
      tokens.push({ literal }); i = Math.min(format.length, end + 1); continue;
    }
    if (format[i] === '[') { const end = format.indexOf(']', i + 1); i = end < 0 ? format.length : end + 1; continue; }
    if (format[i] === '\\' || format[i] === '_' || format[i] === '*') {
      if (i + 1 < format.length) tokens.push({ literal: format[i] === '_' ? ' ' : format[i + 1] });
      i += 2; continue;
    }
    const ampm = /^AM\/PM/i.exec(format.slice(i));
    if (ampm) { tokens.push({ token: 'ampm', code: ampm[0] }); i += ampm[0].length; continue; }
    if (/[ymdhs]/i.test(format[i])) {
      let end = i + 1; while (end < format.length && format[end].toLowerCase() === format[i].toLowerCase()) end++;
      tokens.push({ token: format[i].toLowerCase(), code: format.slice(i, end) }); i = end; continue;
    }
    tokens.push({ literal: format[i++] });
  }
  return tokens;
}
function formatExcelDate(serial, format, language) {
  if (!Number.isFinite(serial) || serial < 0 || serial > 2958465) return null;
  const whole = Math.floor(serial), dayMilliseconds = 86400000;
  const milliseconds = Math.round((serial - whole) * dayMilliseconds);
  const leapDay = whole === 60 && milliseconds < dayMilliseconds;
  const days = whole < 60 ? whole : whole - 1;
  const date = leapDay ? new Date(Date.UTC(1900, 1, 28)) : new Date(Date.UTC(1899, 11, 31) + days * dayMilliseconds + milliseconds);
  if (!Number.isFinite(date.getTime())) return null;
  const parts = {
    year: leapDay ? 1900 : date.getUTCFullYear(), month: leapDay ? 2 : date.getUTCMonth() + 1,
    day: leapDay ? 29 : date.getUTCDate(), weekday: leapDay ? 3 : date.getUTCDay(),
    hour: date.getUTCHours(), minute: date.getUTCMinutes(), second: date.getUTCSeconds()
  };
  const locale = language === 'ko' ? 'ko-KR' : 'en-US';
  const tokens = tokenizeDateFormat(format);
  const formatMonth = (width, month) => {
    if (width === 5) return new Intl.DateTimeFormat(locale, { month: 'narrow', timeZone: 'UTC' }).format(new Date(Date.UTC(2000, month - 1, 1)));
    if (width >= 3) return new Intl.DateTimeFormat(locale, { month: width === 3 ? 'short' : 'long', timeZone: 'UTC' }).format(new Date(Date.UTC(2000, month - 1, 1)));
    return width === 2 ? String(month).padStart(2, '0') : String(month);
  };
  const formatDayName = width => new Intl.DateTimeFormat(locale, { weekday: width === 3 ? 'short' : 'long', timeZone: 'UTC' }).format(new Date(Date.UTC(1900, 0, 7 + parts.weekday)));
  return tokens.map((item, index) => {
    if (item.literal !== undefined) return item.literal;
    const token = item.token, width = item.code.length;
    if (token === 'y') return width <= 2 ? String(parts.year % 100).padStart(2, '0') : String(parts.year).padStart(width, '0');
    if (token === 'd') return width >= 3 ? formatDayName(width) : width === 2 ? String(parts.day).padStart(2, '0') : String(parts.day);
    if (token === 'm') {
      const previous = tokens.slice(0, index).reverse().find(part => part.token);
      const next = tokens.slice(index + 1).find(part => part.token);
      if (previous?.token === 'h' || next?.token === 's') return width === 2 ? String(parts.minute).padStart(2, '0') : String(parts.minute);
      return formatMonth(width, parts.month);
    }
    if (token === 'h') {
      const twelveHour = tokens.some(part => part.token === 'ampm');
      const hour = twelveHour ? (parts.hour % 12 || 12) : parts.hour;
      return width === 2 ? String(hour).padStart(2, '0') : String(hour);
    }
    if (token === 's') return width === 2 ? String(parts.second).padStart(2, '0') : String(parts.second);
    if (token === 'ampm') return parts.hour < 12 ? (language === 'ko' ? '오전' : 'AM') : (language === 'ko' ? '오후' : 'PM');
    return item.code;
  }).join('');
}
function formatExcelNumber(value, numFmtId, formatCode, language) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return String(value);
  const format = formatCode || BUILTIN_NUMBER_FORMATS[numFmtId];
  if (!format || /^General$/i.test(format)) return String(value);
  const dateSyntax = format.replace(/"[^"]*"|\[[^\]]*\]|\\.|_.|\*./g, '');
  if (DATE_FORMAT_IDS.has(numFmtId) || /(?:y{2,4}|d{1,4}|AM\/PM|h{1,2}:[mhs]|m{1,2}:s{1,2})/i.test(dateSyntax)) {
    const dateCode = language === 'ko' && [14, 15, 16, 17, 22].includes(numFmtId)
      ? ({ 14: 'yyyy-mm-dd', 15: 'yyyy-mm-dd', 16: 'm-d', 17: 'yyyy-mm', 22: 'yyyy-mm-dd h:mm' }[numFmtId])
      : format;
    return formatExcelDate(numeric, dateCode, language) ?? String(value);
  }
  if (/[Ee][+-]?0/.test(format) || /[?]/.test(format)) return String(value);
  const sections = format.split(';');
  const selected = numeric < 0 && sections[1] ? sections[1] : sections[0];
  const hasCurrency = numFmtId >= 5 && numFmtId <= 8 || /\$|USD/.test(selected);
  const unquoted = selected.replace(/"[^"]*"|\\./g, '');
  const percent = unquoted.includes('%');
  const literals = [...selected.matchAll(/"([^"]*)"/g)].map(match => match[1]);
  if (literals.some(literal => /[^ $]/.test(literal))) return String(value);
  const cleaned = selected.replace(/"[^"]*"|\[[^\]]*\]|\\.|_.|\*./g, '').replace(/%|\$/g, '').replace(/USD/g, '');
  const numericPattern = cleaned.replace(/[()]/g, '');
  if (!/[0#]/.test(numericPattern) || !/^[#0,.+\-]*$/.test(numericPattern)) return String(value);
  const decimal = /\.([0#]+)/.exec(numericPattern)?.[1] || '';
  const minimumFractionDigits = (decimal.match(/0/g) || []).length;
  const maximumFractionDigits = Math.min(20, decimal.length);
  const options = { useGrouping: /#,##0/.test(numericPattern), minimumFractionDigits, maximumFractionDigits };
  if (hasCurrency) { options.style = 'currency'; options.currency = 'USD'; }
  else if (percent) options.style = 'percent';
  let displayed = new Intl.NumberFormat(language === 'ko' ? 'ko-KR' : 'en-US', options).format(numeric);
  if (numeric < 0 && sections[1] && /^\s*\([^)]*\)\s*$/.test(sections[1].replace(/\[[^\]]*\]/g, ''))) displayed = '(' + new Intl.NumberFormat(language === 'ko' ? 'ko-KR' : 'en-US', { ...options, style: hasCurrency ? 'currency' : 'decimal' }).format(Math.abs(numeric)) + ')';
  return displayed;
}
function formatCellValue(cell, language = 'en') {
  if (!cell || cell.kind !== 'number') return cell?.value ?? '';
  return formatExcelNumber(cell.value, Number(cell.numFmtId) || 0, cell.formatCode || null, language);
}

function parseRelationships(root, basePath, allow) {
  const result = new Map();
  for (const rel of children(root, 'Relationship')) {
    const { Id, Target, Type, TargetMode } = rel.attributes;
    if (!Id || !Target || !Type || result.has(Id)) fail('OOXML 관계 항목이 유효하지 않습니다.');
    if (TargetMode === 'External') { result.set(Id, { external: true, Type, Target }); continue; }
    let resolved = Target.startsWith('/') ? Target.slice(1) : path.join(basePath, Target);
    resolved = path.normalize(resolved);
    if (resolved.startsWith('../') || resolved === '..') fail('OOXML 관계가 패키지 외부 경로를 가리킵니다.');
    result.set(Id, { path: resolved, Type, Target });
    if (allow && !allow.some(suffix => Type.endsWith(suffix))) fail(`지원하지 않는 OOXML 관계입니다: ${Type}`);
  }
  return result;
}

function findCellRanges(xml) {
  const ranges = new Map(), stack = []; let i = 0;
  while (i < xml.length) {
    const lt = xml.indexOf('<', i); if (lt < 0) break;
    if (xml.startsWith('<!--', lt)) { const end = xml.indexOf('-->', lt + 4); if (end < 0) fail('닫히지 않은 XML 주석입니다.'); i = end + 3; continue; }
    if (xml.startsWith('<![CDATA[', lt)) { const end = xml.indexOf(']]>', lt + 9); if (end < 0) fail('닫히지 않은 CDATA입니다.'); i = end + 3; continue; }
    if (xml.startsWith('<?', lt) || xml.startsWith('<!', lt)) { const end = xml.indexOf('>', lt + 2); if (end < 0) fail('닫히지 않은 XML 선언입니다.'); i = end + 1; continue; }
    let end = lt + 1, quote = '';
    for (; end < xml.length; end++) { const ch = xml[end]; if (quote) { if (ch === quote) quote = ''; } else if (ch === '"' || ch === "'") quote = ch; else if (ch === '>') break; }
    if (end >= xml.length) fail('닫히지 않은 XML 태그입니다.');
    const raw = xml.slice(lt, end + 1), match = /^<\s*(\/?)\s*([A-Za-z_][\w.:-]*)/.exec(raw);
    if (!match) { i = end + 1; continue; }
    const closing = match[1] === '/', name = match[2], selfClosing = !closing && /\/\s*>$/.test(raw);
    if (closing) {
      const entry = stack.pop();
      if (!entry || entry.name !== name) fail('XML 요소 중첩이 유효하지 않습니다.');
      if (name === 'c') ranges.set(entry.address, { start: entry.start, startTagEnd: entry.startTagEnd, end: end + 1 });
    } else {
      const attrs = parseTagAttributes(raw);
      const address = name === 'c' ? attrs.r : null;
      if (name === 'c' && (!address || ranges.has(address))) fail('워크시트 셀 주소가 중복되거나 없습니다.');
      if (!selfClosing) stack.push({ name, start: lt, startTagEnd: end + 1, address });
      else if (name === 'c') ranges.set(address, { start: lt, startTagEnd: end + 1, end: end + 1 });
    }
    i = end + 1;
  }
  if (stack.length) fail('닫히지 않은 XML 요소가 있습니다.');
  return ranges;
}

function parseTagAttributes(tag) {
  const attrs = Object.create(null); let i = tag.indexOf(' ');
  if (i < 0) return attrs;
  while (i < tag.length) {
    while (/\s/.test(tag[i] || '')) i++;
    if (tag[i] === '>' || tag[i] === '/' || i >= tag.length) break;
    const nameStart = i; while (i < tag.length && /[\w:.-]/.test(tag[i])) i++;
    const name = tag.slice(nameStart, i); while (/\s/.test(tag[i] || '')) i++;
    if (!name || tag[i++] !== '=') fail('XML 속성 구문이 올바르지 않습니다.');
    while (/\s/.test(tag[i] || '')) i++;
    const quote = tag[i++]; if (quote !== '"' && quote !== "'") fail('XML 속성 인용부호가 올바르지 않습니다.');
    const end = tag.indexOf(quote, i); if (end < 0) fail('XML 속성 인용부호가 닫히지 않았습니다.');
    if (Object.hasOwn(attrs, name)) fail('XML 속성이 중복되었습니다.');
    attrs[name] = tag.slice(i, end).replaceAll('&quot;', '"').replaceAll('&apos;', "'").replaceAll('&lt;', '<').replaceAll('&gt;', '>').replaceAll('&amp;', '&');
    i = end + 1;
  }
  return attrs;
}

function safeUnzip(input) {
  const { bytes, records } = inspectZip(input);
  const files = unzipSync(bytes);
  for (const [name, record] of records) {
    const content = files[name];
    if (!content || content.length !== record.expanded) fail(`ZIP 항목 길이 검증에 실패했습니다: ${name}`);
    if (crc32(content) !== record.expectedCrc) fail(`ZIP CRC 검증에 실패했습니다: ${name}`);
  }
  return { bytes, records, files };
}

function analyze(input, extension = '.xlsx') {
  const { bytes, records, files } = safeUnzip(input);
  const unsupported = [];
  const markUnsupported = reason => { if (!unsupported.includes(reason)) unsupported.push(reason); };
  const known = new Set(['[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels', 'xl/styles.xml', 'xl/sharedStrings.xml', 'xl/theme/theme1.xml', 'docProps/core.xml', 'docProps/app.xml']);
  const rootRels = parseXml(files['_rels/.rels'], '_rels/.rels').root;
  const rootMap = parseRelationships(rootRels, '', ['officeDocument', 'core-properties', 'extended-properties']);
  const workbookPart = [...rootMap.values()].find(rel => rel.Type.endsWith('/officeDocument') && rel.path);
  if (!workbookPart || workbookPart.path !== 'xl/workbook.xml') fail('표준 위치의 통합문서 파트를 찾을 수 없습니다.');
  const workbook = parseXml(files[workbookPart.path], workbookPart.path).root;
  const workbookRelsPath = 'xl/_rels/workbook.xml.rels';
  const workbookRels = parseXml(files[workbookRelsPath], workbookRelsPath).root;
  const workbookMap = parseRelationships(workbookRels, 'xl', ['worksheet', 'styles', 'sharedStrings', 'theme']);
  if ([...rootMap.values(), ...workbookMap.values()].some(rel => rel.external)) markUnsupported('외부 링크 또는 외부 관계가 포함되어 있습니다.');
  const sheetsNode = child(workbook, 'sheets');
  if (!sheetsNode) fail('통합문서에 시트가 없습니다.');
  const sheetRefs = children(sheetsNode, 'sheet');
  if (!sheetRefs.length || sheetRefs.length > LIMITS.sheets) fail('워크시트 개수 제한을 벗어났습니다.');
  if (child(workbook, 'definedNames')) markUnsupported('이름 정의가 포함되어 있습니다.');
  if (child(workbook, 'externalReferences')) markUnsupported('외부 통합문서 참조가 포함되어 있습니다.');
  const sheetNames = new Set(), sheetEntries = [];
  let anyFormula = false, totalCells = 0;
  const sharedStrings = [];
  if (files['xl/sharedStrings.xml']) {
    known.add('xl/sharedStrings.xml');
    const shared = parseXml(files['xl/sharedStrings.xml'], 'xl/sharedStrings.xml').root;
    for (const si of children(shared, 'si')) sharedStrings.push(textContent(si));
  }
  const relationshipType = suffix => [...workbookMap].find(([, rel]) => rel.Type.endsWith(`/${suffix}`))?.[1].path;
  const stylesPath = relationshipType('styles');
  const styleFormats = [];
  if (stylesPath && files[stylesPath]) {
    known.add(stylesPath);
    const styleRoot = parseXml(files[stylesPath], stylesPath).root;
    const customFormats = new Map(children(child(styleRoot, 'numFmts') || { children: [] }, 'numFmt').map(item => [Number(item.attributes.numFmtId), item.attributes.formatCode || '']));
    const fonts = children(child(styleRoot, 'fonts') || { children: [] }, 'font').map(font => {
      const size = Number(child(font, 'sz')?.attributes?.val);
      return {
        bold: Boolean(child(font, 'b') && !['0', 'false'].includes(child(font, 'b').attributes?.val)),
        italic: Boolean(child(font, 'i') && !['0', 'false'].includes(child(font, 'i').attributes?.val)),
        underline: Boolean(child(font, 'u') && !['0', 'false'].includes(child(font, 'u').attributes?.val)),
        strike: Boolean(child(font, 'strike') && !['0', 'false'].includes(child(font, 'strike').attributes?.val)),
        color: styleColor(child(font, 'color')),
        size: Number.isFinite(size) && size >= 6 && size <= 48 ? size : null
      };
    });
    const fills = children(child(styleRoot, 'fills') || { children: [] }, 'fill').map(fill => {
      const pattern = child(fill, 'patternFill');
      return { color: pattern?.attributes?.patternType === 'solid' ? styleColor(child(pattern, 'fgColor')) : null };
    });
    const borders = children(child(styleRoot, 'borders') || { children: [] }, 'border');
    const cellXfs = child(styleRoot, 'cellXfs');
    for (const xf of children(cellXfs || { children: [] }, 'xf')) {
      const numFmtId = Number(xf.attributes.numFmtId) || 0;
      styleFormats.push({
        numFmtId,
        formatCode: customFormats.get(numFmtId) || BUILTIN_NUMBER_FORMATS[numFmtId] || null,
        properties: readCellStyleProperties(xf, fonts, fills, borders)
      });
    }
  }
  const themePath = relationshipType('theme'); if (themePath && files[themePath]) known.add(themePath);
  for (const sheetRef of sheetRefs) {
    const name = sheetRef.attributes.name, relId = sheetRef.attributes['r:id'];
    if (!name || sheetNames.has(name) || !relId) fail('워크시트 이름 또는 관계가 유효하지 않습니다.');
    sheetNames.add(name);
    const rel = workbookMap.get(relId);
    if (!rel?.path || !rel.Type.endsWith('/worksheet') || !files[rel.path]) fail(`시트 관계 대상이 유효하지 않습니다: ${name}`);
    known.add(rel.path);
    const parsed = parseXml(files[rel.path], rel.path); let root = parsed.root;
    const columnWidths = readColumnWidths(root);
    const paneResult = readFreezePane(root);
    if (paneResult.reason) markUnsupported(paneResult.reason);
    const unsupportedFeatures = ['mergeCells', 'autoFilter', 'conditionalFormatting', 'dataValidations', 'drawing', 'legacyDrawing', 'tableParts', 'pivotTableDefinition', 'extLst', 'hyperlinks', 'protectedRanges', 'sheetProtection', 'controls', 'oleObjects'];
    for (const feature of unsupportedFeatures) if (child(root, feature)) markUnsupported(`워크시트 기능 ${feature}은 현재 저장을 지원하지 않습니다.`);
    let sheetData = child(root, 'sheetData'); if (!sheetData) fail(`시트 데이터가 없습니다: ${name}`);
    const sheetDataStructureValid = sheetData.children.every(node => node.name === 'row');
    const rowStructureValid = sheetDataStructureValid && sheetData.children.every(row => /^[1-9]\d*$/.test(row.attributes.r || '') && row.children.every(node => node.name === 'c'));
    const hasColumnDefinitions = Boolean(child(root, 'cols'));
    const cells = new Map();
    for (const row of children(sheetData, 'row')) {
      for (const cell of children(row, 'c')) {
        if (++totalCells > LIMITS.cells) fail('초기 버전은 전체 100,000셀을 넘는 통합문서를 열지 않습니다.');
        const address = cell.attributes.r;
        const r = rowNumber(address), c = columnNumber(address);
        if (!r || r > LIMITS.rows || !c || c > LIMITS.columns || cells.has(address)) fail(`셀 주소가 유효하지 않습니다: ${address}`);
        if (Number(row.attributes.r) !== r) fail(`셀 주소와 행 인덱스가 일치하지 않습니다: ${address}`);
        const formula = child(cell, 'f');
        if (formula) { anyFormula = true; markUnsupported('수식이 포함되어 있어 저장할 수 없습니다. 수식과 캐시값은 보기 전용입니다.'); }
        const cellChildren = cell.children;
        if (cellChildren.some(node => !['f', 'v', 'is'].includes(node.name)) || ['f', 'v', 'is'].some(name => children(cell, name).length > 1)) markUnsupported('지원하지 않는 셀 구성요소가 포함되어 있습니다.');
        if (cellChildren.some(node => ['f', 'v'].includes(node.name) && node.children.length)) markUnsupported('복잡한 셀 값 구조는 현재 저장을 지원하지 않습니다.');
        const inlineString = child(cell, 'is');
        if (inlineString && (inlineString.children.length !== 1 || inlineString.children[0].name !== 't' || inlineString.children[0].children.length)) markUnsupported('서식 또는 확장이 포함된 인라인 문자열은 현재 저장을 지원하지 않습니다.');
        const decoded = decodeCellValue(cell, sharedStrings);
        if (decoded.kind === 'error') markUnsupported('오류 값이 포함되어 있습니다.');
        const styleIndex = Number(cell.attributes.s) || 0, cellFormat = styleFormats[styleIndex] || { numFmtId: 0, formatCode: BUILTIN_NUMBER_FORMATS[0] };
        const item = { address, row: r - 1, column: c - 1, value: decoded.value, kind: decoded.kind, formula: formula?.text || null, cachedValue: decoded.value, style: cell.attributes.s || null, styleProperties: cellFormat.properties || null, numFmtId: cellFormat.numFmtId, formatCode: cellFormat.formatCode };
        cells.set(address, item);

      }
    }
    const ranges = findCellRanges(parsed.xml);
    if (ranges.size !== cells.size) fail(`시트 XML 셀 인덱스 검증에 실패했습니다: ${name} (${[...ranges.keys()].join(',')}/${[...cells.keys()].join(',')})`);
    sheetEntries.push({ name, path: rel.path, cells, ranges, columnWidths, freezePane: paneResult.freezePane, hasColumnDefinitions, sheetDataStructureValid, rowStructureValid, xml: parsed.xml, visible: sheetRef.attributes.state !== 'hidden' && sheetRef.attributes.state !== 'veryHidden' });
    root = null; parsed.root = null; sheetData = null;
  }
  for (const name of records.keys()) {
    if (name.endsWith('/')) continue;
    if (!known.has(name) && name !== '[Content_Types].xml' && name !== '_rels/.rels' && name !== 'xl/workbook.xml' && name !== workbookRelsPath) markUnsupported(`지원하지 않는 OOXML 구성요소가 있습니다: ${name}`);
    if (/\.(xml|rels)$/i.test(name) && (known.has(name) || name === '[Content_Types].xml' || name === '_rels/.rels' || name === workbookRelsPath)) parseXml(files[name], name);
  }
  if (/\.xlsm$/i.test(extension)) markUnsupported('매크로 통합문서는 JinSheet에서 수정할 수 없습니다.');
  if (anyFormula) markUnsupported('수식이 포함되어 있어 저장할 수 없습니다. 수식과 캐시값은 보기 전용입니다.');
  const contentTypes = files['[Content_Types].xml'];
  if (!contentTypes) fail('OOXML 콘텐츠 형식 목록이 없습니다.');
  const types = parseXml(contentTypes, '[Content_Types].xml').root;
  for (const entry of children(types, 'Override')) {
    const type = entry.attributes.ContentType || '';
    if (/macroEnabled|vbaProject|vbaData|activeX|oleObject|pivot|chart|drawing|comments|connections|queryTable|slicer|signature|customXml|externalLink/i.test(type)) markUnsupported(`지원하지 않는 콘텐츠 형식이 있습니다: ${type}`);
  }
  return { bytes, files, records, sheets: sheetEntries, sharedStrings, writable: unsupported.length === 0, readOnlyReason: unsupported.join(' '), unsupported, formulas: anyFormula, extension };
}

function patchCell(model, sheetIndex, address, value) {
  return patchCells(model, [{ sheetIndex, address, value }]);
}

function findElementRanges(xml, targetName, keyAttribute) {
  const ranges = new Map(), stack = [];
  let i = 0;
  while (i < xml.length) {
    const lt = xml.indexOf('<', i);
    if (lt < 0) break;
    if (xml.startsWith('<!--', lt)) {
      const end = xml.indexOf('-->', lt + 4);
      if (end < 0) fail('닫히지 않은 XML 주석입니다.');
      i = end + 3; continue;
    }
    if (xml.startsWith('<![CDATA[', lt)) {
      const end = xml.indexOf(']]>', lt + 9);
      if (end < 0) fail('닫히지 않은 CDATA입니다.');
      i = end + 3; continue;
    }
    if (xml.startsWith('<?', lt) || xml.startsWith('<!', lt)) {
      const end = xml.indexOf('>', lt + 2);
      if (end < 0) fail('닫히지 않은 XML 선언입니다.');
      i = end + 1; continue;
    }
    let end = lt + 1, quote = '';
    for (; end < xml.length; end++) {
      const ch = xml[end];
      if (quote) { if (ch === quote) quote = ''; }
      else if (ch === '"' || ch === "'") quote = ch;
      else if (ch === '>') break;
    }
    if (end >= xml.length) fail('닫히지 않은 XML 태그입니다.');
    const raw = xml.slice(lt, end + 1), match = /^<\s*(\/?)\s*([A-Za-z_][\w.:-]*)/.exec(raw);
    if (!match) { i = end + 1; continue; }
    const closing = match[1] === '/', name = match[2], selfClosing = !closing && /\/\s*>$/.test(raw);
    if (closing) {
      const entry = stack.pop();
      if (!entry || entry.name !== name) fail('XML 요소 중첩이 유효하지 않습니다.');
      if (name === targetName) { if (ranges.has(entry.key)) fail('XML 요소 키가 중복되었습니다: ' + entry.key); ranges.set(entry.key, { start: entry.start, startTagEnd: entry.startTagEnd, closingStart: lt, end: end + 1, selfClosing: false }); }
    } else {
      const attrs = parseTagAttributes(raw);
      const key = name === targetName ? String(keyAttribute ? (attrs[keyAttribute] ?? '') : '') : null;
      if (name === targetName && selfClosing) { if (ranges.has(key)) fail('XML 요소 키가 중복되었습니다: ' + key); ranges.set(key, { start: lt, startTagEnd: end + 1, closingStart: end + 1, end: end + 1, selfClosing: true }); }
      else if (!selfClosing) stack.push({ name, start: lt, startTagEnd: end + 1, key });
    }
    i = end + 1;
  }
  if (stack.length) fail('닫히지 않은 XML 요소가 있습니다.');
  return ranges;
}

function serializeCellContent(value, kind, keep) {
  const text = String(value);
  const numeric = /^\s*[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?\s*$/.test(text) && Number.isFinite(Number(text.trim()));
  if ((kind === 'number' || kind === 'blank') && numeric) return '<c' + keep + '><v>' + text.trim() + '</v></c>';
  if ((kind === 'boolean' || kind === 'blank') && /^(?:TRUE|FALSE)$/i.test(text)) return '<c' + keep + ' t="b"><v>' + (text.toUpperCase() === 'TRUE' ? '1' : '0') + '</v></c>';
  const space = /^\s|\s$/.test(text) ? ' xml:space="preserve"' : '';
  return '<c' + keep + ' t="inlineStr"><is><t' + space + '>' + xmlEscape(text) + '</t></is></c>';
}

function columnLetters(number) {
  let result = '';
  for (let n = number; n > 0; n = Math.floor((n - 1) / 26)) result = String.fromCharCode(65 + (n - 1) % 26) + result;
  return result;
}

function expandWorksheetDimension(xml, address) {
  const dimension = findElementRanges(xml, 'dimension').get('');
  if (!dimension) return xml;
  const opening = xml.slice(dimension.start, dimension.startTagEnd);
  const ref = parseTagAttributes(opening).ref || '';
  const match = /^\$?([A-Z]+)\$?([1-9]\d*)(?::\$?([A-Z]+)\$?([1-9]\d*))?$/.exec(ref);
  if (!match) fail('워크시트 사용 범위를 안전하게 갱신할 수 없습니다.');
  const startColumn = columnNumber(match[1] + '1'), startRow = Number(match[2]);
  const endColumn = match[3] ? columnNumber(match[3] + '1') : startColumn, endRow = match[4] ? Number(match[4]) : startRow;
  const targetColumn = columnNumber(address), targetRow = rowNumber(address);
  const firstColumn = Math.min(startColumn, targetColumn), firstRow = Math.min(startRow, targetRow);
  const lastColumn = Math.max(endColumn, targetColumn), lastRow = Math.max(endRow, targetRow);
  const updatedRef = columnLetters(firstColumn) + firstRow + ':' + columnLetters(lastColumn) + lastRow;
  const nextOpening = opening.replace(/\bref\s*=\s*(["'])[^"']*\1/, function (_all, quote) { return 'ref=' + quote + updatedRef + quote; });
  if (nextOpening === opening && ref !== updatedRef) fail('워크시트 사용 범위를 갱신하지 못했습니다.');
  return xml.slice(0, dimension.start) + nextOpening + xml.slice(dimension.startTagEnd);
}

function addBlankCellOperations(sheet, additions, operations) {
  const rows = findElementRanges(sheet.xml, 'row', 'r');
  const sheetData = findElementRanges(sheet.xml, 'sheetData').get('');
  if (!sheetData || sheetData.selfClosing) fail('워크시트 데이터 영역을 안전하게 찾지 못했습니다.');
  const grouped = new Map();
  for (const cell of additions) {
    const row = rowNumber(cell.address);
    if (!grouped.has(row)) grouped.set(row, []);
    grouped.get(row).push(cell);
  }
  for (const [rowNumberValue, cells] of grouped) {
    cells.sort((a, b) => columnNumber(a.address) - columnNumber(b.address));
    const row = rows.get(String(rowNumberValue));
    if (row && row.selfClosing) {
      const opening = sheet.xml.slice(row.start, row.startTagEnd).replace(/\s+spans\s*=\s*(["'])[^"']*\1/, '').replace(/\/\s*>$/, '>');
      const content = cells.map(function (cell) { return serializeCellContent(cell.value, 'blank', ' r="' + cell.address + '"'); }).join('');
      operations.push({ start: row.start, end: row.end, replacement: opening + content + '</row>', column: 0 });
      continue;
    }
    if (row) {
      const existingColumns = [...sheet.ranges.entries()]
        .filter(function (entry) { return rowNumber(entry[0]) === rowNumberValue; })
        .sort(function (a, b) { return a[1].start - b[1].start; })
        .map(function (entry) { return columnNumber(entry[0]); });
      for (let i = 1; i < existingColumns.length; i++) if (existingColumns[i] <= existingColumns[i - 1]) fail('워크시트 셀 순서가 유효하지 않아 새 셀을 추가할 수 없습니다.');
      const opening = sheet.xml.slice(row.start, row.startTagEnd);
      const cleanOpening = opening.replace(/\s+spans\s*=\s*(["'])[^"']*\1/, '');
      if (cleanOpening !== opening) operations.push({ start: row.start, end: row.startTagEnd, replacement: cleanOpening });
      for (const cell of cells) {
        const column = columnNumber(cell.address);
        const next = [...sheet.ranges.entries()]
          .filter(function (entry) { return rowNumber(entry[0]) === rowNumberValue && columnNumber(entry[0]) > column; })
          .sort(function (a, b) { return columnNumber(a[0]) - columnNumber(b[0]); })[0];
        const position = next ? next[1].start : row.closingStart;
        operations.push({ start: position, end: position, replacement: serializeCellContent(cell.value, 'blank', ' r="' + cell.address + '"'), column });
      }
      continue;
    }
    const laterRow = [...rows.entries()].filter(function (entry) { return Number(entry[0]) > rowNumberValue; }).sort(function (a, b) { return Number(a[0]) - Number(b[0]); })[0];
    const position = laterRow ? laterRow[1].start : sheetData.closingStart;
    const cellXml = cells.map(function (cell) { return serializeCellContent(cell.value, 'blank', ' r="' + cell.address + '"'); }).join('');
    operations.push({ start: position, end: position, replacement: '<row r="' + rowNumberValue + '">' + cellXml + '</row>', column: rowNumberValue });
  }
}

function patchCells(model, edits) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  if (!Array.isArray(edits) || edits.length > 10000) fail('셀 수정 개수가 유효하지 않습니다.');
  const bySheet = new Map(), additions = new Map(), seen = new Set();
  for (const edit of edits) {
    const sheetIndex = edit.sheetIndex, address = edit.address, value = edit.value, sheet = model.sheets[sheetIndex];
    const row = rowNumber(address), column = columnNumber(address);
    if (!sheet || !row || row > LIMITS.rows || !column || column > LIMITS.columns || typeof value !== 'string' || value.length > 100000) fail('셀 수정 요청이 유효하지 않습니다.');
    const key = sheetIndex + ':' + address;
    if (seen.has(key)) fail('같은 셀을 중복 수정할 수 없습니다.');
    seen.add(key);
    const cell = sheet.cells.get(address), range = sheet.ranges.get(address);
    if (!cell && !range) {
      if (!additions.has(sheetIndex)) additions.set(sheetIndex, []);
      additions.get(sheetIndex).push({ address, value });
      if (!bySheet.has(sheetIndex)) bySheet.set(sheetIndex, []);
      continue;
    }
    if (!cell || !range || cell.formula || cell.kind === 'error') fail('기존의 단순 값 셀만 수정할 수 있습니다.');
    const originalOpen = sheet.xml.slice(range.start, range.startTagEnd);
    if (/\b[\w.-]+:/.test(originalOpen.replace(/^<[^\s>]+/, ''))) fail('접두사가 있는 셀 속성은 수정할 수 없습니다.');
    const attributes = parseTagAttributes(originalOpen);
    const keep = Object.entries(attributes).filter(function (entry) { return entry[0] !== 't'; }).map(function (entry) { return ' ' + entry[0] + '="' + xmlEscape(entry[1]) + '"'; }).join('');
    const replacement = serializeCellContent(value, cell.kind, keep);
    if (!bySheet.has(sheetIndex)) bySheet.set(sheetIndex, []);
    bySheet.get(sheetIndex).push(Object.assign({}, range, { replacement, column }));
  }
  for (const [sheetIndex, newCells] of additions) addBlankCellOperations(model.sheets[sheetIndex], newCells, bySheet.get(sheetIndex));
  const outputFiles = Object.assign({}, model.files), modifiedPaths = new Set();
  for (const [sheetIndex, patches] of bySheet) {
    const sheet = model.sheets[sheetIndex];
    patches.sort(function (a, b) { return b.start - a.start || (b.column || 0) - (a.column || 0); });
    let xml = sheet.xml, previousStart = xml.length + 1;
    for (const patch of patches) {
      if (patch.end > previousStart) fail('셀 수정 범위가 겹칩니다.');
      xml = xml.slice(0, patch.start) + patch.replacement + xml.slice(patch.end);
      previousStart = patch.start;
    }
    for (const cell of additions.get(sheetIndex) || []) xml = expandWorksheetDimension(xml, cell.address);
    outputFiles[sheet.path] = new TextEncoder().encode(xml);
    modifiedPaths.add(sheet.path);
  }
  const output = zipSync(outputFiles, { level: 0 });
  const checked = analyze(output, model.extension);
  if (!checked.writable) fail('저장 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  for (const [name, before] of Object.entries(model.files)) {
    if (modifiedPaths.has(name)) continue;
    if (!checked.files[name] || Buffer.compare(Buffer.from(before), Buffer.from(checked.files[name])) !== 0) fail('비수정 패키지 항목이 달라졌습니다: ' + name);
  }
  return { bytes: output, model: checked };
}

function setTagAttribute(tag, name, value) {
  const element = /^<([^\s/>]+)/.exec(tag);
  if (!element) fail('OOXML 시작 태그를 수정할 수 없습니다.');
  const attributes = parseTagAttributes(tag);
  attributes[name] = String(value);
  const serialized = Object.entries(attributes).map(function (entry) { return ' ' + entry[0] + '="' + xmlEscape(entry[1]) + '"'; }).join('');
  return '<' + element[1] + serialized + (/\/\s*>$/.test(tag) ? '/>' : '>');
}

function patchColumnWidths(model, edits) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  if (!Array.isArray(edits) || edits.length > 10000) fail('열 너비 수정 개수가 유효하지 않습니다.');
  if (!edits.length) return { bytes: model.bytes, model };
  const bySheet = new Map(), seen = new Set();
  for (const edit of edits) {
    const { sheetIndex, column, width } = edit || {};
    if (!Number.isInteger(sheetIndex) || !model.sheets[sheetIndex] || !Number.isInteger(column) || column < 0 || column >= LIMITS.columns || typeof width !== 'number' || !Number.isFinite(width) || width < 0 || width > 255) fail('열 너비 수정 요청이 유효하지 않습니다.');
    const key = sheetIndex + ':' + column;
    if (seen.has(key)) fail('같은 열 너비를 중복 수정할 수 없습니다.');
    seen.add(key);
    if (!bySheet.has(sheetIndex)) bySheet.set(sheetIndex, new Map());
    bySheet.get(sheetIndex).set(column, Math.round(width * 100) / 100);
  }
  const outputFiles = Object.assign({}, model.files), modifiedPaths = new Set();
  for (const [sheetIndex, requestedWidths] of bySheet) {
    const sheet = model.sheets[sheetIndex];
    const columnRanges = [...findElementRanges(sheet.xml, 'col', 'min').entries()]
      .map(function (entry) { const tag = sheet.xml.slice(entry[1].start, entry[1].startTagEnd); const attributes = parseTagAttributes(tag); return { range: entry[1], tag, attributes, min: Number(attributes.min), max: Number(attributes.max) }; })
      .sort(function (left, right) { return left.min - right.min; });
    for (let i = 0; i < columnRanges.length; i++) {
      const item = columnRanges[i];
      if (!item.range.selfClosing) fail('비표준 열 정의를 안전하게 수정할 수 없어 저장을 중단했습니다.');
      if (i && item.min <= columnRanges[i - 1].max) fail('겹치는 열 정의를 안전하게 수정할 수 없어 저장을 중단했습니다.');
    }
    const operations = [], inserted = [];
    for (const [column, width] of requestedWidths) {
      const coordinate = column + 1;
      const existing = columnRanges.find(item => coordinate >= item.min && coordinate <= item.max);
      const current = sheet.columnWidths[column];
      if (current !== undefined && Math.abs(current - width) < 0.005) continue;
      if (!existing) inserted.push({ column, width });
    }
    for (const item of columnRanges) {
      const targets = [...requestedWidths.entries()]
        .filter(([column, width]) => column + 1 >= item.min && column + 1 <= item.max && (item.attributes.width === undefined || Math.abs(Number(item.attributes.width) - width) >= 0.005))
        .sort((left, right) => left[0] - right[0]);
      if (!targets.length) continue;
      const parts = [];
      let cursor = item.min;
      const segment = (first, last, width, custom) => {
        let tag = setTagAttribute(item.tag, 'min', first);
        tag = setTagAttribute(tag, 'max', last);
        if (width !== undefined) tag = setTagAttribute(tag, 'width', width);
        if (custom) tag = setTagAttribute(tag, 'customWidth', '1');
        return tag;
      };
      for (const [column, width] of targets) {
        const coordinate = column + 1;
        if (cursor < coordinate) parts.push(segment(cursor, coordinate - 1, item.attributes.width, false));
        parts.push(segment(coordinate, coordinate, width, true));
        cursor = coordinate + 1;
      }
      if (cursor <= item.max) parts.push(segment(cursor, item.max, item.attributes.width, false));
      operations.push({ start: item.range.start, end: item.range.end, replacement: parts.join('') });
    }
    if (!inserted.length && !operations.length) continue;
    const cols = findElementRanges(sheet.xml, 'cols').get('');
    if (cols?.selfClosing) {
      const opening = sheet.xml.slice(cols.start, cols.startTagEnd).replace(/\/\s*>$/, '>');
      const tags = inserted.sort((a, b) => a.column - b.column).map(item => `<col min="${item.column + 1}" max="${item.column + 1}" width="${item.width}" customWidth="1"/>`).join('');
      operations.push({ start: cols.start, end: cols.end, replacement: opening + tags + '</cols>' });
    } else if (cols) {
      const insertions = new Map();
      for (const item of inserted) {
        const next = columnRanges.find(range => range.min > item.column + 1);
        const position = next ? next.range.start : cols.closingStart;
        if (!insertions.has(position)) insertions.set(position, []);
        insertions.get(position).push(item);
      }
      for (const [position, items] of insertions) {
        const tags = items.sort((a, b) => a.column - b.column).map(item => `<col min="${item.column + 1}" max="${item.column + 1}" width="${item.width}" customWidth="1"/>`).join('');
        operations.push({ start: position, end: position, replacement: tags });
      }
    } else if (inserted.length) {
      const sheetData = findElementRanges(sheet.xml, 'sheetData').get('');
      if (!sheetData || sheetData.selfClosing) fail('워크시트 데이터 앞에 열 너비를 안전하게 추가할 수 없습니다.');
      const tags = inserted.sort((a, b) => a.column - b.column).map(item => `<col min="${item.column + 1}" max="${item.column + 1}" width="${item.width}" customWidth="1"/>`).join('');
      operations.push({ start: sheetData.start, end: sheetData.start, replacement: '<cols>' + tags + '</cols>' });
    }
    operations.sort((left, right) => right.start - left.start || right.end - left.end);
    let xml = sheet.xml, previousStart = xml.length + 1;
    for (const operation of operations) {
      if (operation.end > previousStart) fail('열 너비 수정 범위가 겹칩니다.');
      xml = xml.slice(0, operation.start) + operation.replacement + xml.slice(operation.end);
      previousStart = operation.start;
    }
    outputFiles[sheet.path] = new TextEncoder().encode(xml);
    modifiedPaths.add(sheet.path);
  }
  if (!modifiedPaths.size) return { bytes: model.bytes, model };
  const output = zipSync(outputFiles, { level: 0 });
  const checked = analyze(output, model.extension);
  if (!checked.writable) fail('열 너비 저장 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  for (const [name, before] of Object.entries(model.files)) {
    if (modifiedPaths.has(name)) continue;
    if (!checked.files[name] || Buffer.compare(Buffer.from(before), Buffer.from(checked.files[name])) !== 0) fail('열 너비 저장 중 비수정 패키지 항목이 달라졌습니다: ' + name);
  }
  return { bytes: output, model: checked };
}

function replaceTagAttribute(tag, name, value) {
  const element = /^<([^\s/>]+)/.exec(tag);
  if (!element) fail('OOXML 시작 태그를 수정할 수 없습니다.');
  const attributes = parseTagAttributes(tag);
  if (!Object.hasOwn(attributes, name)) fail('OOXML 구조 변경 대상에 필수 속성이 없습니다: ' + name);
  attributes[name] = String(value);
  const serialized = Object.entries(attributes).map(function (entry) { return ' ' + entry[0] + '="' + xmlEscape(entry[1]) + '"'; }).join('');
  return '<' + element[1] + serialized + (/\/\s*>$/.test(tag) ? '/>' : '>');
}

function removeTagAttribute(tag, name) {
  const element = /^<([^\s/>]+)/.exec(tag);
  if (!element) fail('OOXML 시작 태그를 수정할 수 없습니다.');
  const attributes = parseTagAttributes(tag);
  if (!Object.hasOwn(attributes, name)) return tag;
  delete attributes[name];
  const serialized = Object.entries(attributes).map(function (entry) { return ' ' + entry[0] + '="' + xmlEscape(entry[1]) + '"'; }).join('');
  return '<' + element[1] + serialized + (/\/\s*>$/.test(tag) ? '/>' : '>');
}
function updateDimensionForStructure(xml, operation) {
  const dimension = findElementRanges(xml, 'dimension').get('');
  if (!dimension) return xml;
  const opening = xml.slice(dimension.start, dimension.startTagEnd);
  const attrs = parseTagAttributes(opening);
  const match = /^\$?([A-Z]+)\$?([1-9]\d*)(?::\$?([A-Z]+)\$?([1-9]\d*))?$/.exec(attrs.ref || '');
  if (!match) fail('워크시트 사용 범위를 안전하게 이동할 수 없습니다.');
  let firstColumn = columnNumber(match[1] + '1'), firstRow = Number(match[2]);
  let lastColumn = match[3] ? columnNumber(match[3] + '1') : firstColumn;
  let lastRow = match[4] ? Number(match[4]) : firstRow;
  const coordinate = operation.index + 1;
  const first = operation.axis === 'row' ? firstRow : firstColumn;
  const last = operation.axis === 'row' ? lastRow : lastColumn;
  if (operation.action === 'insert') {
    if (coordinate <= last) {
      if (coordinate < first) {
        if (operation.axis === 'row') firstRow = coordinate;
        else firstColumn = coordinate;
      }
      if (operation.axis === 'row') lastRow++;
      else lastColumn++;
    }
  } else if (coordinate < first) {
    if (operation.axis === 'row') { firstRow--; lastRow--; }
    else { firstColumn--; lastColumn--; }
  } else if (coordinate <= last) {
    if (operation.axis === 'row') lastRow--;
    else lastColumn--;
  }
  const nextRef = lastRow < firstRow || lastColumn < firstColumn
    ? 'A1'
    : columnLetters(firstColumn) + firstRow + ':' + columnLetters(lastColumn) + lastRow;
  const updated = replaceTagAttribute(opening, 'ref', nextRef);
  return xml.slice(0, dimension.start) + updated + xml.slice(dimension.startTagEnd);
}

function transformStructure(model, operation) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  const { sheetIndex, axis, action, index } = operation || {};
  const sheet = model.sheets[sheetIndex];
  if (!sheet || !['row', 'column'].includes(axis) || !['insert', 'delete'].includes(action) || !Number.isInteger(index) || index < 0) fail('XLSX 행·열 구조 변경 요청이 유효하지 않습니다.');
  const coordinate = index + 1;
  const maxCoordinate = axis === 'row' ? LIMITS.rows : LIMITS.columns;
  if (coordinate > maxCoordinate) fail('XLSX 행·열 제한을 벗어났습니다.');
  if (axis === 'column' && sheet.hasColumnDefinitions) fail('열 너비나 열 서식 정의가 있어 안전한 열 이동을 보장할 수 없습니다.');
  if (!sheet.sheetDataStructureValid) fail('워크시트 행 구조를 안전하게 이동할 수 없습니다.');
  if (!sheet.rowStructureValid) fail('행 번호 또는 행 내부 기능을 안전하게 이동할 수 없습니다.');
  const rowRanges = findElementRanges(sheet.xml, 'row', 'r');
  const rowsByPosition = [...rowRanges.entries()].sort((a, b) => a[1].start - b[1].start);
  for (let i = 1; i < rowsByPosition.length; i++) {
    if (Number(rowsByPosition[i][0]) <= Number(rowsByPosition[i - 1][0])) fail('워크시트 행 순서가 유효하지 않아 구조를 변경할 수 없습니다.');
  }
  for (const row of new Set([...sheet.cells.values()].map(cell => cell.row))) {
    const columns = [...sheet.cells.values()].filter(cell => cell.row === row).sort((a, b) => sheet.ranges.get(a.address).start - sheet.ranges.get(b.address).start).map(cell => cell.column);
    for (let i = 1; i < columns.length; i++) if (columns[i] <= columns[i - 1]) fail('워크시트 셀 순서가 유효하지 않아 구조를 변경할 수 없습니다.');
  }
  const patches = [];
  const targetRow = rowRanges.get(String(coordinate));
  if (axis === 'row' && action === 'delete' && targetRow) patches.push({ start: targetRow.start, end: targetRow.end, replacement: '' });
  const insertedRowAnchor = axis === 'row' && action === 'insert'
    ? rowsByPosition.find(entry => Number(entry[0]) >= coordinate)
    : null;
  const shiftedRowNumbers = new Set();
  for (const [number, range] of rowsByPosition) {
    const oldNumber = Number(number);
    if (axis === 'row' && action === 'delete' && oldNumber === coordinate) continue;
    let newNumber = oldNumber;
    if (axis === 'row') {
      if (action === 'insert' && oldNumber >= coordinate) newNumber++;
      if (action === 'delete' && oldNumber > coordinate) newNumber--;
      if (newNumber > LIMITS.rows) fail('행을 삽입하면 XLSX 행 제한을 초과합니다.');
      shiftedRowNumbers.add(oldNumber);
    }
    const original = sheet.xml.slice(range.start, range.startTagEnd);
    let updated = axis === 'row' && newNumber !== oldNumber ? replaceTagAttribute(original, 'r', String(newNumber)) : original;
    if (axis === 'column') updated = removeTagAttribute(updated, 'spans');
    if (insertedRowAnchor && insertedRowAnchor[0] === number) updated = '<row r="' + coordinate + '"/>' + updated;
    if (updated !== original) patches.push({ start: range.start, end: range.startTagEnd, replacement: updated });
  }
  if (axis === 'row' && action === 'insert' && !insertedRowAnchor) {
    const sheetData = findElementRanges(sheet.xml, 'sheetData').get('');
    if (!sheetData || sheetData.selfClosing) fail('워크시트 데이터 끝 위치를 찾을 수 없습니다.');
    patches.push({ start: sheetData.closingStart, end: sheetData.closingStart, replacement: '<row r="' + coordinate + '"/>' });
  }
  for (const cell of sheet.cells.values()) {
    const row = cell.row + 1, column = cell.column + 1;
    if (axis === 'row' && action === 'delete' && row === coordinate) continue;
    let nextRow = row, nextColumn = column;
    if (axis === 'row' && action === 'insert' && row >= coordinate) nextRow++;
    if (axis === 'row' && action === 'delete' && row > coordinate) nextRow--;
    if (axis === 'column' && action === 'insert' && column >= coordinate) nextColumn++;
    if (axis === 'column' && action === 'delete' && column === coordinate) {
      const range = sheet.ranges.get(cell.address);
      patches.push({ start: range.start, end: range.end, replacement: '' });
      continue;
    }
    if (axis === 'column' && action === 'delete' && column > coordinate) nextColumn--;
    if (nextRow > LIMITS.rows || nextColumn > LIMITS.columns) fail('구조 변경이 XLSX 행·열 제한을 초과합니다.');
    if (nextRow !== row || nextColumn !== column) {
      const range = sheet.ranges.get(cell.address);
      const opening = sheet.xml.slice(range.start, range.startTagEnd);
      const newAddress = columnLetters(nextColumn) + nextRow;
      patches.push({ start: range.start, end: range.startTagEnd, replacement: replaceTagAttribute(opening, 'r', newAddress) });
    }
  }
  patches.sort((a, b) => b.start - a.start);
  let xml = sheet.xml, previousStart = xml.length + 1;
  for (const patch of patches) {
    if (patch.end > previousStart) fail('XLSX 구조 변경 범위가 겹쳐 저장을 중단했습니다.');
    xml = xml.slice(0, patch.start) + patch.replacement + xml.slice(patch.end);
    previousStart = patch.start;
  }
  xml = updateDimensionForStructure(xml, { axis, action, index });
  const outputFiles = Object.assign({}, model.files);
  outputFiles[sheet.path] = new TextEncoder().encode(xml);
  const bytes = zipSync(outputFiles, { level: 0 });
  const checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('구조 변경 후 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  for (const [name, before] of Object.entries(model.files)) {
    if (name === sheet.path) continue;
    if (!checked.files[name] || Buffer.compare(Buffer.from(before), Buffer.from(checked.files[name])) !== 0) fail('비수정 패키지 항목이 달라졌습니다: ' + name);
  }
  return { bytes, model: checked };
}

module.exports = { LIMITS, analyze, children, formatCellValue, inspectZip, parseXml, patchCell, patchCells, patchColumnWidths, transformStructure, rowNumber, safeUnzip, textContent };
