'use strict';

function decode(bytes) {
  const data = Buffer.from(bytes);
  if (data.length >= 2 && data[0] === 0xff && data[1] === 0xfe) {
    return { text: new TextDecoder('utf-16le', { fatal: true }).decode(data.subarray(2)), encoding: 'utf-16le', bom: data.subarray(0, 2) };
  }
  if (data.length >= 2 && data[0] === 0xfe && data[1] === 0xff) {
    if ((data.length - 2) % 2) throw new Error('홀수 길이 UTF-16BE 파일은 편집할 수 없습니다.');
    const swapped = Buffer.from(data.subarray(2));
    for (let i = 0; i < swapped.length; i += 2) [swapped[i], swapped[i + 1]] = [swapped[i + 1], swapped[i]];
    return { text: new TextDecoder('utf-16le', { fatal: true }).decode(swapped), encoding: 'utf-16be', bom: data.subarray(0, 2) };
  }
  const bom = data.length >= 3 && data[0] === 0xef && data[1] === 0xbb && data[2] === 0xbf ? data.subarray(0, 3) : Buffer.alloc(0);
  return { text: new TextDecoder('utf-8', { fatal: true }).decode(data.subarray(bom.length)), encoding: 'utf-8', bom };
}

function detectDelimiter(text, preferred) {
  if (preferred === ',' || preferred === '\t' || preferred === ';') return preferred;
  const counts = new Map([[',', 0], ['\t', 0], [';', 0]]);
  let quoted = false;
  for (let i = 0; i < text.length && text[i] !== '\r' && text[i] !== '\n'; i++) {
    if (text[i] === '"') {
      if (quoted && text[i + 1] === '"') i++;
      else quoted = !quoted;
    } else if (!quoted && counts.has(text[i])) counts.set(text[i], counts.get(text[i]) + 1);
  }
  const ranked = [...counts].sort((a, b) => b[1] - a[1]);
  if (ranked[0][1] === 0 || (ranked[0][1] === ranked[1][1] && ranked[0][1] < 2)) {
    throw new Error('구분자를 확실히 판별할 수 없습니다. 쉼표 또는 탭 형식으로 저장된 파일인지 확인하세요.');
  }
  return ranked[0][0];
}

function parse(text, delimiter, maxCells = 100000) {
  const rows = [];
  let row = [], value = '', start = 0, i = 0, quoted = false, fieldQuoted = false, justClosed = false, cellCount = 0;
  const pushField = (end) => { if (++cellCount > maxCells) throw new Error(`초기 버전은 ${maxCells.toLocaleString()}셀을 넘는 CSV/TSV 파일을 열지 않습니다.`); row.push({ value, start, end, quoted: fieldQuoted }); value = ''; fieldQuoted = false; justClosed = false; start = end + 1; };
  const pushRow = (end) => { pushField(end); rows.push(row); row = []; };
  while (i < text.length) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"' && text[i + 1] === '"') { value += '"'; i += 2; continue; }
      if (ch === '"') { quoted = false; justClosed = true; i++; continue; }
      value += ch; i++; continue;
    }
    if (justClosed && ch !== delimiter && ch !== '\r' && ch !== '\n') throw new Error(`인용 필드 뒤의 문자가 유효하지 않습니다 (문자 위치 ${i}).`);
    if (ch === '"' && value.length === 0 && !fieldQuoted) { quoted = true; fieldQuoted = true; i++; continue; }
    if (ch === delimiter) { pushField(i); i++; start = i; continue; }
    if (ch === '\r' || ch === '\n') {
      const width = ch === '\r' && text[i + 1] === '\n' ? 2 : 1;
      pushRow(i); i += width; start = i; continue;
    }
    if (ch === '"') throw new Error(`따옴표가 잘못된 레코드입니다 (문자 위치 ${i}).`);
    value += ch; i++;
  }
  if (quoted) throw new Error('닫히지 않은 인용 필드가 있어 편집을 막았습니다.');
  if (start < text.length || row.length || value.length || (text.length > 0 && !/[\r\n]$/.test(text))) pushRow(text.length);
  return rows;
}

function encode(text, encoding, bom) {
  let body;
  if (encoding === 'utf-16le') body = Buffer.from(text, 'utf16le');
  else if (encoding === 'utf-16be') {
    body = Buffer.from(text, 'utf16le');
    for (let i = 0; i < body.length; i += 2) [body[i], body[i + 1]] = [body[i + 1], body[i]];
  } else body = Buffer.from(text, 'utf8');
  return Buffer.concat([bom, body]);
}

function quote(value, delimiter, forceQuote) {
  const text = String(value);
  return forceQuote || text.includes(delimiter) || /["\r\n]/.test(text) || /^\s|\s$/.test(text)
    ? `"${text.replaceAll('"', '""')}"` : text;
}

function applyEdits(text, rows, edits, delimiter) {
  const patches = edits.map(({ row, column, value }) => {
    const field = rows[row]?.[column];
    if (!field) throw new RangeError('수정할 셀 범위를 벗어났습니다.');
    return { start: field.start, end: field.end, replacement: quote(value, delimiter, field.quoted) };
  }).sort((a, b) => b.start - a.start);
  let result = text;
  let previousStart = text.length + 1;
  for (const patch of patches) {
    if (patch.end > previousStart) throw new Error('겹치는 셀 수정 요청을 거부했습니다.');
    result = result.slice(0, patch.start) + patch.replacement + result.slice(patch.end);
    previousStart = patch.start;
  }
  return result;
}

function load(bytes, extension, preferredDelimiter) {
  const decoded = decode(bytes);
  const delimiter = extension === '.tsv' ? '\t' : detectDelimiter(decoded.text, preferredDelimiter);
  const rows = parse(decoded.text, delimiter);
  return { ...decoded, delimiter, rows };
}

module.exports = { applyEdits, decode, detectDelimiter, encode, load, parse, quote };
