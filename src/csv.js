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
  let row = [], value = '', start = 0, rowStart = 0, i = 0, quoted = false, fieldQuoted = false, justClosed = false, quotedValueStart = 0, quotedParts = null, cellCount = 0;
  const pushField = (end) => {
    if (++cellCount > maxCells) throw new Error(`초기 버전은 ${maxCells.toLocaleString()}셀을 넘는 CSV/TSV 파일을 열지 않습니다.`);
    row.push({ value: fieldQuoted ? value : text.slice(start, end), start, end, quoted: fieldQuoted });
    value = ''; fieldQuoted = false; justClosed = false; quotedParts = null; start = end + 1;
  };
  const pushRow = (end, recordEnd = end) => {
    pushField(end);
    Object.defineProperties(row, {
      start: { value: rowStart },
      contentEnd: { value: end },
      end: { value: recordEnd },
      terminator: { value: text.slice(end, recordEnd) }
    });
    rows.push(row); row = []; rowStart = recordEnd;
  };
  while (i < text.length) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"' && text[i + 1] === '"') {
        if (!quotedParts) quotedParts = [];
        quotedParts.push(text.slice(quotedValueStart, i), '"');
        i += 2; quotedValueStart = i; continue;
      }
      if (ch === '"') {
        value = quotedParts ? quotedParts.concat(text.slice(quotedValueStart, i)).join('') : text.slice(quotedValueStart, i);
        quoted = false; justClosed = true; i++; continue;
      }
      i++; continue;
    }
    if (justClosed && ch !== delimiter && ch !== '\r' && ch !== '\n') throw new Error(`인용 필드 뒤의 문자가 유효하지 않습니다 (문자 위치 ${i}).`);
    if (ch === '"' && i === start && !fieldQuoted) { quoted = true; fieldQuoted = true; quotedValueStart = i + 1; i++; continue; }
    if (ch === delimiter) { pushField(i); i++; continue; }
    if (ch === '\r' || ch === '\n') {
      const width = ch === '\r' && text[i + 1] === '\n' ? 2 : 1;
      pushRow(i, i + width); i += width; start = i; continue;
    }
    if (ch === '"') throw new Error(`따옴표가 잘못된 레코드입니다 (문자 위치 ${i}).`);
    i++;
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

function applyTextPatches(text, patches, errorMessage, validPatch) {
  const ordered = patches.map((patch, index) => ({ ...patch, index }))
    .sort((left, right) => left.start - right.start || right.index - left.index);
  const chunks = [];
  let offset = 0;
  for (const patch of ordered) {
    if (!validPatch(patch) || patch.start < offset) throw new Error(errorMessage);
    chunks.push(text.slice(offset, patch.start), patch.replacement);
    offset = patch.end;
  }
  chunks.push(text.slice(offset));
  return chunks.join('');
}

function applyEdits(text, rows, edits, delimiter) {
  const patches = edits.map(({ row, column, value }) => {
    const field = rows[row]?.[column];
    if (!field) throw new RangeError('수정할 셀 범위를 벗어났습니다.');
    return { start: field.start, end: field.end, replacement: quote(value, delimiter, field.quoted) };
  });
  return applyTextPatches(text, patches, '겹치는 셀 수정 요청을 거부했습니다.', patch => patch.start >= 0 && patch.end >= patch.start && patch.end <= text.length);
}

function load(bytes, extension, preferredDelimiter) {
  const decoded = decode(bytes);
  const delimiter = extension === '.tsv' ? '\t' : detectDelimiter(decoded.text, preferredDelimiter);
  const rows = parse(decoded.text, delimiter);
  return { ...decoded, delimiter, rows };
}

function applyStructure(text, patches) {
  return applyTextPatches(text, patches, 'CSV 구조 변경 범위가 겹치거나 유효하지 않습니다.', patch => patch.start >= 0 && patch.end >= patch.start && patch.end <= text.length);
}

function recordEnding(state, index) {
  for (let i = index; i < state.rows.length; i++) if (state.rows[i].terminator) return state.rows[i].terminator;
  for (let i = Math.min(index - 1, state.rows.length - 1); i >= 0; i--) if (state.rows[i].terminator) return state.rows[i].terminator;
  if (state.text.includes('\r\n')) return '\r\n';
  if (state.text.includes('\r')) return '\r';
  return '\n';
}

function transformStructure(state, extension, operation) {
  const { axis, action, index } = operation || {};
  if (!['row', 'column'].includes(axis) || !['insert', 'delete'].includes(action) || !Number.isInteger(index) || index < 0) throw new RangeError('CSV 행·열 구조 변경 요청이 유효하지 않습니다.');
  const rows = state.rows;
  const columnCount = rows.reduce((count, row) => Math.max(count, row.length), 0);
  const patches = [];
  if (axis === 'row') {
    if (action === 'insert') {
      if (index > rows.length || rows.length >= 10000) throw new RangeError('CSV에서 행을 삽입할 위치 또는 행 수 제한이 유효하지 않습니다.');
      const width = Math.max(1, columnCount), record = Array(width).fill('').join(state.delimiter);
      if (!rows.length) patches.push({ start: 0, end: 0, replacement: quote('', state.delimiter, true) });
      else if (index < rows.length) patches.push({ start: rows[index].start, end: rows[index].start, replacement: record + recordEnding(state, index) });
      else {
        const last = rows[rows.length - 1];
        patches.push({ start: state.text.length, end: state.text.length, replacement: last.terminator ? record : recordEnding(state, index) + record });
      }
    } else {
      if (index >= rows.length) throw new RangeError('삭제할 CSV 행을 찾을 수 없습니다.');
      const row = rows[index];
      if (row.terminator) patches.push({ start: row.start, end: row.end, replacement: '' });
      else if (index > 0) patches.push({ start: rows[index - 1].contentEnd, end: row.end, replacement: '' });
      else patches.push({ start: row.start, end: row.end, replacement: '' });
    }
  } else {
    if (action === 'insert') {
      if (index > columnCount || columnCount >= 1000) throw new RangeError('CSV에서 열을 삽입할 위치 또는 열 수 제한이 유효하지 않습니다.');
      if (!rows.length) patches.push({ start: 0, end: 0, replacement: quote('', state.delimiter, true) });
      for (const row of rows) {
        if (index < row.length) patches.push({ start: row[index].start, end: row[index].start, replacement: state.delimiter });
        else if (index === columnCount && row.length === columnCount && row.length) patches.push({ start: row.contentEnd, end: row.contentEnd, replacement: state.delimiter });
      }
    } else {
      if (index >= columnCount) throw new RangeError('삭제할 CSV 열을 찾을 수 없습니다.');
      for (const row of rows) {
        if (index >= row.length) continue;
        if (row.length === 1) patches.push({ start: row[0].start, end: row[0].end, replacement: '' });
        else if (index === 0) patches.push({ start: row[0].start, end: row[1].start, replacement: '' });
        else patches.push({ start: row[index - 1].end, end: row[index].end, replacement: '' });
      }
    }
  }
  const text = applyStructure(state.text, patches);
  return load(encode(text, state.encoding, state.bom), extension, state.delimiter);
}

function extendCells(state, extension, coordinates) {
  if (!Array.isArray(coordinates) || coordinates.some(item => !Number.isInteger(item?.row) || item.row < 0 || !Number.isInteger(item?.column) || item.column < 0)) throw new RangeError('CSV 셀 확장 좌표가 유효하지 않습니다.');
  if (!coordinates.length) return state;
  const lastRow = Math.max(...coordinates.map(item => item.row));
  let expanded = state;
  if (lastRow >= 10000) throw new RangeError('CSV 행 제한을 벗어난 셀입니다.');
  if (expanded.rows.length <= lastRow) {
    const missingRows = lastRow + 1 - expanded.rows.length;
    const width = Math.max(1, expanded.rows.reduce((count, row) => Math.max(count, row.length), 0));
    const record = Array(width).fill('').join(expanded.delimiter);
    const ending = recordEnding(expanded, expanded.rows.length);
    let suffix;
    if (!expanded.rows.length) {
      // Preserve the established representation for the first empty CSV row.
      suffix = quote('', expanded.delimiter, true) + (missingRows > 1 ? ending + Array(missingRows - 1).fill(record).join(ending) : '');
    } else {
      const last = expanded.rows[expanded.rows.length - 1];
      suffix = (last.terminator ? '' : ending) + Array(missingRows).fill(record).join(ending);
    }
    const text = expanded.text + suffix;
    expanded = load(encode(text, expanded.encoding, expanded.bom), extension, expanded.delimiter);
  }
  const columnsByRow = new Map();
  for (const { row, column } of coordinates) columnsByRow.set(row, Math.max(columnsByRow.get(row) ?? -1, column));
  const patches = [];
  for (const [rowIndex, lastColumn] of columnsByRow) {
    const row = expanded.rows[rowIndex];
    const missingFields = lastColumn + 1 - row.length;
    if (missingFields > 0) patches.push({ start: row.contentEnd, end: row.contentEnd, replacement: expanded.delimiter.repeat(missingFields) });
  }
  if (!patches.length) return expanded;
  const text = applyStructure(expanded.text, patches);
  return load(encode(text, expanded.encoding, expanded.bom), extension, expanded.delimiter);
}

module.exports = { applyEdits, transformStructure, extendCells, decode, detectDelimiter, encode, load, parse, quote };
