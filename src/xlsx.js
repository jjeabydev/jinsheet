'use strict';

const path = require('node:path').posix;
const { unzipSync, zipSync } = require('../vendor/node_modules/fflate');
const { SaxesParser } = require('../vendor/node_modules/saxes');

const LIMITS = Object.freeze({ bytes: 20 * 1024 * 1024, entries: 512, expanded: 64 * 1024 * 1024, entry: 16 * 1024 * 1024, ratio: 200, sheets: 100, rows: 1000000, columns: 16384, cells: 100000, xmlNodes: 200000 });
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
  const stylesPath = relationshipType('styles'); if (stylesPath && files[stylesPath]) known.add(stylesPath);
  const themePath = relationshipType('theme'); if (themePath && files[themePath]) known.add(themePath);
  for (const sheetRef of sheetRefs) {
    const name = sheetRef.attributes.name, relId = sheetRef.attributes['r:id'];
    if (!name || sheetNames.has(name) || !relId) fail('워크시트 이름 또는 관계가 유효하지 않습니다.');
    sheetNames.add(name);
    const rel = workbookMap.get(relId);
    if (!rel?.path || !rel.Type.endsWith('/worksheet') || !files[rel.path]) fail(`시트 관계 대상이 유효하지 않습니다: ${name}`);
    known.add(rel.path);
    const parsed = parseXml(files[rel.path], rel.path), root = parsed.root;
    const unsupportedFeatures = ['mergeCells', 'autoFilter', 'conditionalFormatting', 'dataValidations', 'drawing', 'legacyDrawing', 'tableParts', 'pivotTableDefinition', 'extLst', 'hyperlinks', 'protectedRanges', 'sheetProtection', 'controls', 'oleObjects'];
    for (const feature of unsupportedFeatures) if (child(root, feature)) markUnsupported(`워크시트 기능 ${feature}은 현재 저장을 지원하지 않습니다.`);
    const sheetData = child(root, 'sheetData'); if (!sheetData) fail(`시트 데이터가 없습니다: ${name}`);
    const rows = new Map(), cells = new Map();
    for (const row of children(sheetData, 'row')) {
      for (const cell of children(row, 'c')) {
        if (++totalCells > LIMITS.cells) fail('초기 버전은 전체 100,000셀을 넘는 통합문서를 열지 않습니다.');
        const address = cell.attributes.r;
        const r = rowNumber(address), c = columnNumber(address);
        if (!r || r > LIMITS.rows || !c || c > LIMITS.columns || cells.has(address)) fail(`셀 주소가 유효하지 않습니다: ${address}`);
        if (Number(row.attributes.r) !== r) fail(`셀 주소와 행 인덱스가 일치하지 않습니다: ${address}`);
        const formula = child(cell, 'f');
        if (formula) { anyFormula = true; markUnsupported('수식이 포함되어 있어 저장할 수 없습니다. 수식과 캐시값은 보기 전용입니다.'); }
        const decoded = decodeCellValue(cell, sharedStrings);
        if (decoded.kind === 'error') markUnsupported('오류 값이 포함되어 있습니다.');
        const item = { address, row: r - 1, column: c - 1, value: decoded.value, kind: decoded.kind, formula: formula?.text || null, cachedValue: decoded.value, style: cell.attributes.s || null };
        cells.set(address, item);
        if (!rows.has(r - 1)) rows.set(r - 1, []);
        rows.get(r - 1)[c - 1] = item;
      }
    }
    const ranges = findCellRanges(parsed.xml);
    if (ranges.size !== cells.size) fail(`시트 XML 셀 인덱스 검증에 실패했습니다: ${name} (${[...ranges.keys()].join(',')}/${[...cells.keys()].join(',')})`);
    sheetEntries.push({ name, path: rel.path, rows, cells, ranges, xml: parsed.xml, root, visible: sheetRef.attributes.state !== 'hidden' && sheetRef.attributes.state !== 'veryHidden' });
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

function patchCells(model, edits) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  if (!Array.isArray(edits) || !edits.length || edits.length > 10000) fail('셀 수정 개수가 유효하지 않습니다.');
  const bySheet = new Map();
  for (const { sheetIndex, address, value } of edits) {
    const sheet = model.sheets[sheetIndex], cell = sheet?.cells.get(address), range = sheet?.ranges.get(address);
    if (!cell || !range || cell.formula || cell.kind === 'error') fail('기존의 단순 값 셀만 수정할 수 있습니다.');
    const originalOpen = sheet.xml.slice(range.start, range.startTagEnd);
    if (/\b[\w.-]+:/.test(originalOpen.replace(/^<[^\s>]+/, ''))) fail('접두사가 있는 셀 속성은 수정할 수 없습니다.');
    const attributes = parseTagAttributes(originalOpen);
    const keep = Object.entries(attributes).filter(([name]) => name !== 't').map(([name, v]) => ` ${name}="${xmlEscape(v)}"`).join('');
    const space = /^\s|\s$/.test(String(value)) ? ' xml:space="preserve"' : '';
    const replacement = `<c${keep} t="inlineStr"><is><t${space}>${xmlEscape(value)}</t></is></c>`;
    if (!bySheet.has(sheetIndex)) bySheet.set(sheetIndex, []);
    bySheet.get(sheetIndex).push({ ...range, replacement, sheet });
  }
  const outputFiles = { ...model.files }, modifiedPaths = new Set();
  for (const [sheetIndex, patches] of bySheet) {
    const sheet = model.sheets[sheetIndex];
    patches.sort((a, b) => b.start - a.start);
    let xml = sheet.xml, previousStart = xml.length + 1;
    for (const patch of patches) {
      if (patch.end > previousStart) fail('셀 수정 범위가 겹칩니다.');
      xml = xml.slice(0, patch.start) + patch.replacement + xml.slice(patch.end);
      previousStart = patch.start;
    }
    outputFiles[sheet.path] = new TextEncoder().encode(xml);
    modifiedPaths.add(sheet.path);
  }
  const output = zipSync(outputFiles, { level: 0 });
  const checked = analyze(output, model.extension);
  if (!checked.writable) fail(`저장 후 패키지 검사에서 거절되었습니다: ${checked.readOnlyReason}`);
  for (const [name, before] of Object.entries(model.files)) {
    if (modifiedPaths.has(name)) continue;
    if (!checked.files[name] || Buffer.compare(Buffer.from(before), Buffer.from(checked.files[name])) !== 0) fail(`비수정 패키지 항목이 달라졌습니다: ${name}`);
  }
  return { bytes: output, model: checked };
}

module.exports = { LIMITS, analyze, children, inspectZip, parseXml, patchCell, patchCells, rowNumber, safeUnzip, textContent };
