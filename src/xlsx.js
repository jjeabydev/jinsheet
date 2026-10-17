'use strict';

const path = require('node:path').posix;
const { unzipSync, zipSync } = require('../vendor/node_modules/fflate');
const { SaxesParser } = require('../vendor/node_modules/saxes');
const { createFormulaCalculator, evaluateFormula } = require('./formula');
const { translateFormulaReferences }=require('./formula-references');
const { conditionalStyleForCell: resolveConditionalStyle }=require('./conditional-style');
const { createColorContext, resolveStyleColor } = require('./style-colors');
const { defaultHeaderFooter, parseHeaderFooterText, encodeHeaderFooterText, validateHeaderFooter } = require('./header-footer');
const { drawingAxes, drawingPoint, drawingMarker } = require('./chart-print');
const { excelColumnWidthPixels } = require('./grid');
const { validateCustomFilter, validateTopFilter, topFilterThreshold, validateDateGroups, serialDateParts, validateDynamicFilter, dynamicFilterCriteria, validateColorFilter, ICON_SET_LENGTHS, validateIconFilter, filterColumnCells, compileAutoFilterColumn } = require('./auto-filter');

const LIMITS = Object.freeze({ bytes: 20 * 1024 * 1024, entries: 512, expanded: 64 * 1024 * 1024, entry: 16 * 1024 * 1024, ratio: 200, sheets: 100, rows: 1048576, columns: 16384, cells: 100000, xmlNodes: 500000 });
const ROW_VISIBILITY_NS = 'https://jinsheet.dev/ooxml/row-visibility/1';
const COLUMN_VISIBILITY_NS = 'https://jinsheet.dev/ooxml/column-visibility/1';
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
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
function localName(node) { return node?.name?.split(':').at(-1) || ''; }
function childrenByLocalName(node, name) { return (node?.children || []).filter(item => localName(item) === name); }
function firstDescendant(node, names) {
  if (!node) return null;
  if (names.includes(localName(node))) return node;
  for (const item of node.children) { const found = firstDescendant(item, names); if (found) return found; }
  return null;
}

function chartCacheValues(reference) {
  const cache = firstDescendant(reference, ['numCache', 'strCache', 'numLit', 'strLit']);
  if (!cache) return [];
  const points = childrenByLocalName(cache, 'pt');
  if (points.length > 10000) throw new Error('차트 캐시 점 개수 제한을 초과했습니다.');
  const countNode = childrenByLocalName(cache, 'ptCount')[0], countText = countNode?.attributes.val;
  if (countNode && (!/^(?:0|[1-9]\d*)$/.test(countText || '') || Number(countText) > 10000)) throw new Error('차트 캐시 점 개수가 유효하지 않습니다.');
  const maxIndex = points.reduce((max, point) => Math.max(max, Number(point.attributes.idx) + 1), 0);
  const count = countNode ? Number(countText) : maxIndex;
  if (maxIndex > count) throw new Error('차트 캐시 점 인덱스가 개수를 벗어납니다.');
  const values = Array(count).fill(null), seenIndexes = new Set();
  for (const point of points) {
    const indexText = point.attributes.idx, raw = childrenByLocalName(point, 'v')[0]?.text;
    if (!/^(?:0|[1-9]\d*)$/.test(indexText || '') || Number(indexText) >= count || raw === undefined || seenIndexes.has(Number(indexText))) throw new Error('차트 캐시 항목이 유효하지 않습니다.');
    const index = Number(indexText);
    seenIndexes.add(index);
    values[index] = localName(cache) === 'numCache' || localName(cache) === 'numLit' ? Number.isFinite(Number(raw)) ? Number(raw) : null : raw;
  }
  return values;
}

function chartReferenceValues(container, resolveReference) {
  const reference = firstDescendant(container, ['strRef', 'numRef', 'multiLvlStrRef', 'strLit', 'numLit']);
  return reference ? resolveReference?.(reference) ?? chartCacheValues(reference) : [];
}

function chartSeriesName(series, index, resolveReference) {
  const title = childrenByLocalName(series, 'tx')[0];
  if (title) {
    const values = chartReferenceValues(title, resolveReference);
    if (values.length && values[0] !== null) return String(values[0]);
    const literal = childrenByLocalName(title, 'v')[0]?.text;
    if (literal) return literal;
  }
  return `Series ${index + 1}`;
}

const DRAWING_CHART_ELEMENTS = new Set(['wsDr', 'twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor', 'pos', 'from', 'to', 'col', 'row', 'colOff', 'rowOff', 'ext', 'graphicFrame', 'nvGraphicFramePr', 'cNvPr', 'cNvGraphicFramePr', 'xfrm', 'off', 'graphic', 'graphicData', 'chart', 'clientData']);

// Keep drawing geometry separate from the live chart data/cache model.
function readChartAnchor(anchor, geometryOnly = false) {
  if (!anchor) return null;
  const client = childrenByLocalName(anchor, 'clientData')[0];
  const hidden = firstDescendant(anchor, ['cNvPr'])?.attributes.hidden;
  if (!geometryOnly && (['0', 'false'].includes(client?.attributes.fPrintsWithSheet) || ['1', 'true'].includes(hidden))) return { printable: false };
  if (!geometryOnly && client?.attributes.fPrintsWithSheet && !['1', 'true'].includes(client.attributes.fPrintsWithSheet)) return null;
  if (!geometryOnly && hidden && !['0', 'false'].includes(hidden)) return null;
  if (anchor.attributes.editAs && !['twoCell', 'oneCell', 'absolute'].includes(anchor.attributes.editAs)) return null;
  const transform = firstDescendant(anchor, ['xfrm']);
  if (!geometryOnly && transform && Object.entries(transform.attributes).some(([key, value]) => !['rot', 'flipH', 'flipV'].includes(key) || !['0', 'false'].includes(value))) return null;
  const number = (value, maximum) => /^\d+$/.test(value ?? '') && Number.isSafeInteger(Number(value)) && Number(value) <= maximum ? Number(value) : NaN;
  const marker = name => {
    const node = childrenByLocalName(anchor, name)[0];
    if (!node || ['col', 'row', 'colOff', 'rowOff'].some(key => childrenByLocalName(node, key).length !== 1)) return null;
    const read = (key, maximum) => number(childrenByLocalName(node, key)[0].text.trim(), maximum);
    const value = { column: read('col', 16383), row: read('row', 1048575), x: read('colOff', 2147483647) / 9525, y: read('rowOff', 2147483647) / 9525 };
    return Object.values(value).every(Number.isFinite) ? value : null;
  };
  const type = localName(anchor), from = marker('from');
  if (type === 'twoCellAnchor') {
    const to = marker('to');
    return from && to ? { printable: true, type, from, to, ...(['oneCell', 'absolute'].includes(anchor.attributes.editAs) ? { editAs: anchor.attributes.editAs } : {}) } : null;
  }
  const ext = childrenByLocalName(anchor, 'ext')[0], width = number(ext?.attributes.cx, 2147483647) / 9525, height = number(ext?.attributes.cy, 2147483647) / 9525;
  if (!(width > 0 && height > 0)) return null;
  if (type === 'oneCellAnchor') return from ? { printable: true, type, from, width, height } : null;
  const pos = childrenByLocalName(anchor, 'pos')[0], x = number(pos?.attributes.x, 2147483647) / 9525, y = number(pos?.attributes.y, 2147483647) / 9525;
  return type === 'absoluteAnchor' && Number.isFinite(x) && Number.isFinite(y) ? { printable: true, type, x, y, width, height } : null;
}

function hasUnsupportedDrawingObjects(root) {
  if (!DRAWING_CHART_ELEMENTS.has(localName(root))) return true;
  if (localName(root) === 'graphicFrame' && !firstDescendant(root, ['chart'])) return true;
  return root.children.some(hasUnsupportedDrawingObjects);
}

function readChartPreview(bytes, partName, defaultSheetName, resolveReference) {
  const { root } = parseXml(bytes, partName);
  const sourceRanges = [];
  const formulaNodes = [];
  (function collect(node) { if (localName(node) === 'f') formulaNodes.push(node); for (const item of node.children) collect(item); })(root);
  for (const formula of formulaNodes) {
    const source = parseRangeReference(formula.text.trim(), defaultSheetName);
    if (!source || source.top !== source.bottom && source.left !== source.right) return null;
    sourceRanges.push(source);
  }
  const plotArea = firstDescendant(root, ['plotArea']);
  if (!plotArea) return null;
  const chartNodes = plotArea.children.filter(item => ['barChart', 'lineChart', 'areaChart', 'scatterChart', 'pieChart', 'doughnutChart'].includes(localName(item)));
  if (chartNodes.length !== 1) return null;
  const [chartNode] = chartNodes;
  const chartKind = localName(chartNode);
  const type = chartKind === 'barChart' ? 'bar' : chartKind === 'lineChart' ? 'line' : chartKind === 'areaChart' ? 'area' : chartKind === 'scatterChart' ? 'scatter' : chartKind === 'doughnutChart' ? 'doughnut' : 'pie';
  const direction = type === 'bar' ? childrenByLocalName(chartNode, 'barDir')[0]?.attributes.val === 'bar' ? 'horizontal' : 'vertical' : 'vertical';
  const seriesNodes = childrenByLocalName(chartNode, 'ser');
  if (!seriesNodes.length || seriesNodes.length > 32) return null;
  const series = [];
  for (const [index, item] of seriesNodes.entries()) {
    const categories = chartReferenceValues(childrenByLocalName(item, type === 'scatter' ? 'xVal' : 'cat')[0], resolveReference);
    const values = chartReferenceValues(childrenByLocalName(item, type === 'scatter' ? 'yVal' : 'val')[0], resolveReference);
    if (!values.length || values.length > 10000 || values.some(value => value !== null && (typeof value !== 'number' || !Number.isFinite(value))) || type === 'scatter' && (!categories.length || categories.length !== values.length || categories.some(value => value !== null && (typeof value !== 'number' || !Number.isFinite(value))))) return null;
    series.push({ name: chartSeriesName(item, index, resolveReference).slice(0, 256), categories: categories.slice(0, 10000).map(value => value === null ? '' : String(value).slice(0, 256)), values, ...(type === 'scatter' ? { xValues: categories } : {}) });
  }
  const titleNode = firstDescendant(root, ['title']);
  const title = titleNode ? [...(function* walk(node) { if (localName(node) === 't') yield node.text; for (const item of node.children) yield* walk(item); })(titleNode)].join('').slice(0, 512) : '';
  return { type, direction, title, series, sourceRanges };
}

function readWorksheetCharts(root, worksheetPath, sheetName, files, known, markUnsupported) {
  const drawing = childrenByLocalName(root, 'drawing')[0];
  if (!drawing) return { charts: [], sourceRanges: [], chartPaths: [] };
  const relationId = drawing.attributes['r:id'];
  const relationshipsPath = path.join(path.dirname(worksheetPath), '_rels', path.basename(worksheetPath) + '.rels');
  if (!relationId || !files[relationshipsPath]) { markUnsupported('차트 그리기 관계를 찾을 수 없습니다.'); return { charts: [], sourceRanges: [], chartPaths: [] }; }
  let worksheetRelationships, drawingRelationship;
  try {
    worksheetRelationships = parseRelationships(parseXml(files[relationshipsPath], relationshipsPath).root, path.dirname(worksheetPath), ['table', 'hyperlink', 'drawing']);
    drawingRelationship = worksheetRelationships.get(relationId);
    if (!drawingRelationship?.path || !drawingRelationship.Type.endsWith('/drawing') || !files[drawingRelationship.path]) throw new Error('워크시트 그리기 대상이 유효하지 않습니다.');
    known.add(relationshipsPath); known.add(drawingRelationship.path);
    const drawingRoot = parseXml(files[drawingRelationship.path], drawingRelationship.path).root;
    const drawingRelsPath = path.join(path.dirname(drawingRelationship.path), '_rels', path.basename(drawingRelationship.path) + '.rels');
    if (!files[drawingRelsPath]) throw new Error('그리기 차트 관계 파일을 찾을 수 없습니다.');
    const drawingRelationships = parseRelationships(parseXml(files[drawingRelsPath], drawingRelsPath).root, path.dirname(drawingRelationship.path), ['chart', 'image']);
    known.add(drawingRelsPath);
    const charts = [], sourceRanges = [], chartPaths = [], chartAnchors = [], chartPlacements = [];
    const parents = new Map();
    (function collect(node, anchor = null) { if (['twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor'].includes(localName(node))) anchor = node; if (localName(node) === 'chart') parents.set(node, anchor); for (const item of node.children) collect(item, anchor); })(drawingRoot);
    const printUnsupported = hasUnsupportedDrawingObjects(drawingRoot);
    if (printUnsupported) markUnsupported('차트 외 그리기 개체가 있어 안전한 편집을 보장할 수 없습니다.');
    for (const reference of (function* walk(node) { if (localName(node) === 'chart' && node.attributes['r:id']) yield node; for (const item of node.children) yield* walk(item); })(drawingRoot)) {
      const relation = drawingRelationships.get(reference.attributes['r:id']);
      if (!relation?.path || !relation.Type.endsWith('/chart') || !files[relation.path]) throw new Error('차트 관계 대상이 유효하지 않습니다.');
      known.add(relation.path);
      chartPaths.push(relation.path);
      const chart = readChartPreview(files[relation.path], relation.path, sheetName);
      if (chart) {
        const anchor = parents.get(reference), kind = anchor && localName(anchor);
        charts.push({ type: chart.type, direction: chart.direction, title: chart.title, series: chart.series }); chartAnchors.push(readChartAnchor(anchor));
        chartPlacements.push(kind === 'twoCellAnchor' ? anchor.attributes.editAs || 'twoCell' : kind === 'oneCellAnchor' ? 'oneCell' : kind === 'absoluteAnchor' ? 'absolute' : null);
        sourceRanges.push(...chart.sourceRanges);
      }
      else markUnsupported('차트 형식 또는 원본 범위를 안전하게 해석할 수 없습니다.');
      if (charts.length > 32) throw new Error('워크시트 차트 개수 제한을 초과했습니다.');
    }
    for (const relation of drawingRelationships.values()) if (relation.external) markUnsupported('외부 그리기 관계가 있어 안전한 편집을 보장할 수 없습니다.');
    return { charts, sourceRanges, chartPaths, chartAnchors, chartPlacements, printUnsupported, drawingPath: drawingRelationship.path };
  } catch (error) {
    markUnsupported(`차트를 안전하게 미리볼 수 없습니다: ${error.message}`);
    return { charts: [], sourceRanges: [], chartPaths: [] };
  }
}

function readWorksheetHyperlinks(root, worksheetPath, files, known, markUnsupported) {
  const containers = childrenByLocalName(root, 'hyperlinks');
  if (!containers.length) return { relationshipIds: new Set(), references: [] };
  if (containers.length !== 1 || Object.keys(containers[0].attributes).length) { markUnsupported('워크시트 하이퍼링크 목록 구성이 유효하지 않습니다.'); return { relationshipIds: new Set(), references: [] }; }
  const container = containers[0], links = childrenByLocalName(container, 'hyperlink'), relationshipIds = new Set();
  const references = [];
  if (!links.length || links.length > 10000 || links.length !== container.children.length) { markUnsupported('워크시트 하이퍼링크 개수 또는 항목 구성이 유효하지 않습니다.'); return { relationshipIds, references }; }
  const relationshipsPath = path.join(path.dirname(worksheetPath), '_rels', path.basename(worksheetPath) + '.rels');
  let relationships = new Map();
  if (files[relationshipsPath]) {
    try {
      known.add(relationshipsPath);
      relationships = parseRelationships(parseXml(files[relationshipsPath], relationshipsPath).root, path.dirname(worksheetPath), ['table', 'hyperlink', 'drawing']);
    } catch (error) { markUnsupported(`워크시트 하이퍼링크 관계를 안전하게 보존할 수 없습니다: ${error.message}`); }
  }
  for (const link of links) {
    const { ref, 'r:id': id, location } = link.attributes;
    const range = ref && parseRangeReference(ref);
    if (!range || range.sheetName || Object.keys(link.attributes).some(attribute => !['ref', 'r:id', 'location', 'display', 'tooltip'].includes(attribute)) || link.children.length) {
      markUnsupported('워크시트 하이퍼링크 항목이 유효하지 않습니다.');
      continue;
    }
    let target = null;
    if (id) {
      const relation = relationships.get(id);
      if (!relation?.external || !relation.Type.endsWith('/hyperlink')) markUnsupported('워크시트 하이퍼링크 외부 관계가 유효하지 않습니다.');
      else { relationshipIds.add(id); target = relation.Target; }
    } else if (!location) markUnsupported('워크시트 하이퍼링크의 대상 위치가 없습니다.');
    references.push({ ...range, relationshipId: id || null, location: location || null, display: link.attributes.display || null, tooltip: link.attributes.tooltip || null, target });
  }
  for (const [id, relation] of relationships) if (relation.Type.endsWith('/hyperlink') && !relationshipIds.has(id)) markUnsupported('사용되지 않는 워크시트 하이퍼링크 관계가 있습니다.');
  return { relationshipIds, references };
}


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
  const format = child(root, 'sheetFormatPr'), defaultText = format?.attributes.defaultColWidth;
  const defaultCharacters = defaultText === undefined ? null : Number(defaultText);
  if (defaultCharacters !== null && (!/^\+?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(defaultText) || !Number.isFinite(defaultCharacters) || defaultCharacters < 0 || defaultCharacters > 255)) fail('기본 열 너비가 유효하지 않습니다.');
  const defaultWidthCharacters = defaultCharacters;
  const defaultWidth = defaultCharacters === null ? 64 : excelColumnWidthPixels(defaultCharacters);
  const definitions = children(child(root, 'cols') || { children: [] }, 'col');
  if (definitions.length > LIMITS.columns) fail('워크시트 열 너비 정의 개수 제한을 초과했습니다.');
  const widths = {}, hiddenColumns = new Set(), manualHiddenColumns = new Set(), manualVisibility = {};
  const visibilityPrefix = Object.keys(root.attributes).find(key => key.startsWith('xmlns:') && root.attributes[key] === COLUMN_VISIBILITY_NS)?.slice(6) || null;
  let expandedColumns = 0;
  for (const definition of definitions) {
    const { min: minText, max: maxText, width: widthText } = definition.attributes;
    if (!/^[1-9]\d*$/.test(minText || '') || !/^[1-9]\d*$/.test(maxText || '')) fail('워크시트 열 너비 범위가 유효하지 않습니다.');
    const min = Number(minText), max = Number(maxText);
    if (min > max || max > LIMITS.columns) fail('워크시트 열 너비 범위가 제한을 벗어났습니다.');
    const hiddenText = definition.attributes.hidden;
    if (hiddenText !== undefined && !['1', 'true', 'on', '0', 'false', 'off'].includes(hiddenText.toLocaleLowerCase())) fail('워크시트 숨김 열 속성이 유효하지 않습니다.');
    const manualHiddenText = visibilityPrefix ? definition.attributes[visibilityPrefix + ':manualHidden'] : undefined;
    if (manualHiddenText !== undefined && !['1', 'true', 'on', '0', 'false', 'off'].includes(String(manualHiddenText).toLocaleLowerCase())) fail('수동 숨김 열 메타데이터가 유효하지 않습니다.');
    expandedColumns += max - min + 1;
    if (expandedColumns > LIMITS.columns) fail('워크시트 열 너비 확장 제한을 초과했습니다.');
    const hidden = hiddenText !== undefined && ['1', 'true', 'on'].includes(hiddenText.toLocaleLowerCase());
    const manualHidden = visibilityPrefix && ['1', 'true', 'on'].includes(String(manualHiddenText || '').toLocaleLowerCase());
    let width;
    if (widthText !== undefined) {
      if (!/^\+?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(widthText)) fail('워크시트 열 너비 값이 유효하지 않습니다.');
      width = Number(widthText);
      if (!Number.isFinite(width) || width < 0 || width > 255) fail('워크시트 열 너비 값이 유효하지 않습니다.');
    }
    for (let column = min; column <= max; column++) {
      if (hidden || width !== undefined && excelColumnWidthPixels(width) === 0 || width === undefined && defaultWidth === 0) hiddenColumns.add(column - 1);
      if (manualHiddenText !== undefined) manualVisibility[column - 1] = Boolean(manualHidden);
      if (manualHidden) manualHiddenColumns.add(column - 1);
      if (width !== undefined) widths[column - 1] = width;
    }
  }
  if (defaultWidth === 0) for (let column = 0; column < LIMITS.columns; column++) if (widths[column] === undefined) hiddenColumns.add(column);
  return { widths, defaultWidth, defaultWidthCharacters, hiddenColumns: [...hiddenColumns].sort((left, right) => left - right), nativeHiddenColumns: [...hiddenColumns].sort((left, right) => left - right), manualHiddenColumns, manualVisibility };
}

function readOutlineMetadata(root, sheetData) {
  const readFlag = (value, name, fallback) => {
    if (value === undefined) return fallback;
    if (!['1', 'true', 'on', '0', 'false', 'off'].includes(String(value).toLocaleLowerCase())) fail('개요 설정이 유효하지 않습니다: ' + name);
    return ['1', 'true', 'on'].includes(String(value).toLocaleLowerCase());
  };
  const readLevel = (value, name) => {
    if (value === undefined) return 0;
    if (!/^(?:0|[1-7])$/.test(String(value))) fail('개요 수준이 0~7 범위를 벗어났습니다: ' + name);
    return Number(value);
  };
  const outlinePr = child(child(root, 'sheetPr') || { children: [] }, 'outlinePr');
  const summaryBelow = readFlag(outlinePr?.attributes.summaryBelow, 'summaryBelow', true);
  const summaryRight = readFlag(outlinePr?.attributes.summaryRight, 'summaryRight', true);
  const rowOutlineLevels = {}, columnOutlineLevels = {}, collapsedRows = new Set(), collapsedColumns = new Set();
  for (const row of children(sheetData, 'row')) {
    const rowNumber = Number(row.attributes.r);
    if (!Number.isInteger(rowNumber) || rowNumber < 1 || rowNumber > LIMITS.rows) continue;
    const index = rowNumber - 1, level = readLevel(row.attributes.outlineLevel, 'row ' + rowNumber);
    if (level) rowOutlineLevels[index] = level;
    if (readFlag(row.attributes.collapsed, 'row ' + rowNumber + ' collapsed', false)) collapsedRows.add(index);
  }
  for (const column of children(child(root, 'cols') || { children: [] }, 'col')) {
    const min = Number(column.attributes.min), max = Number(column.attributes.max);
    if (!Number.isInteger(min) || !Number.isInteger(max) || min < 1 || max < min || max > LIMITS.columns) continue;
    const level = readLevel(column.attributes.outlineLevel, 'column ' + min + ':' + max);
    if (level) for (let index = min - 1; index < max; index++) columnOutlineLevels[index] = level;
    if (readFlag(column.attributes.collapsed, 'column ' + min + ':' + max + ' collapsed', false)) {
      for (let index = min - 1; index < max; index++) collapsedColumns.add(index);
    }
  }
  const groupsFor = (levels, collapsed, summaryAfter) => {
    const groups = [];
    for (let level = 1; level <= 7; level++) {
      const coordinates = Object.keys(levels).map(Number).filter(index => levels[index] >= level).sort((a, b) => a - b);
      let start = null, end = null;
      const push = () => {
        if (start === null) return;
        const summary = summaryAfter ? end + 1 : start - 1;
        groups.push({ start, end, level, summary, collapsed: collapsed.has(summary) });
      };
      for (const coordinate of coordinates) {
        if (start === null) start = end = coordinate;
        else if (coordinate === end + 1) end = coordinate;
        else { push(); start = end = coordinate; }
      }
      push();
    }
    return groups.sort((a, b) => a.start - b.start || a.end - b.end || a.level - b.level);
  };
  const rowGroups = groupsFor(rowOutlineLevels, collapsedRows, summaryBelow);
  const columnGroups = groupsFor(columnOutlineLevels, collapsedColumns, summaryRight);
  const outlineHiddenRows = [...new Set(rowGroups.filter(group => group.collapsed).flatMap(group => Array.from({ length: group.end - group.start + 1 }, (_, offset) => group.start + offset)))].sort((a, b) => a - b);
  const outlineHiddenColumns = [...new Set(columnGroups.filter(group => group.collapsed).flatMap(group => Array.from({ length: group.end - group.start + 1 }, (_, offset) => group.start + offset)))].sort((a, b) => a - b);
  return { summaryBelow, summaryRight, rowOutlineLevels, columnOutlineLevels, rowGroups, columnGroups, outlineHiddenRows, outlineHiddenColumns };
}
function xmlEscape(value) { return String(value).replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&apos;'); }

function columnNumber(address) {
  const letters = /^([A-Z]+)([1-9]\d*)$/.exec(address)?.[1];
  if (!letters) return 0;
  let n = 0; for (const char of letters) n = n * 26 + char.charCodeAt(0) - 64;
  return n;
}
function rowNumber(address) { const match = /^[A-Z]+([1-9]\d*)$/.exec(address); return match ? Number(match[1]) : 0; }
function tableFilterRange(table) {
  if ((table.headerRowCount ?? 1) !== 1) return null;
  return { ...table.range, bottom: table.range.bottom - (table.totalsRowCount || 0) };
}

function tableForFilter(sheet, filter) {
  return filter && sheet.tables?.find(table => {
    const range = tableFilterRange(table);
    return range && ['top', 'bottom', 'left', 'right'].every(key => range[key] === filter[key]);
  });
}

function readFilterColor(dxf, cellColor, colorContext) {
  if (cellColor) {
    const pattern = child(child(dxf, 'fill') || { children: [] }, 'patternFill');
    if (pattern?.attributes.patternType === 'none' && pattern.children.every(item => ['fgColor', 'bgColor'].includes(item.name))) return null;
    const foreground = child(pattern || { children: [] }, 'fgColor');
    if (!pattern || pattern.attributes.patternType !== 'solid' || pattern.children.some(item => !['fgColor', 'bgColor'].includes(item.name)) || !foreground || !styleColor(foreground, colorContext)) fail('색상 필터의 단색 정의를 해석할 수 없습니다.');
    return styleColor(foreground, colorContext).toUpperCase();
  }
  const font = child(dxf, 'font'), color = child(font || { children: [] }, 'color');
  if (color && Object.keys(color.attributes).length === 1 && ['1', 'true'].includes(color.attributes.auto)) return null;
  if (!color || !styleColor(color, colorContext)) fail('글자색 필터의 색상 정의를 해석할 수 없습니다.');
  return styleColor(color, colorContext).toUpperCase();
}

function ensureFilterColorStyle(xml, colorFilter, colorContext) {
  validateColorFilter(colorFilter);
  const root = parseXml(new TextEncoder().encode(xml), 'filter-styles').root, nodes = children(child(root, 'dxfs') || { children: [] }, 'dxf');
  const color = colorFilter.color?.toUpperCase() ?? null;
  for (let index = 0; index < nodes.length; index++) {
    try { if (readFilterColor(nodes[index], colorFilter.cellColor, colorContext) === color) return { xml, dxfId: index }; } catch { /* Other differential styles remain untouched. */ }
  }
  if (nodes.length >= 65536) fail('색상 필터의 차등 서식 개수 제한을 초과했습니다.');
  const value = color === null ? '' : color.slice(1), tag = colorFilter.cellColor
    ? '<dxf><fill><patternFill patternType="' + (value ? 'solid' : 'none') + '">' + (value ? '<fgColor rgb="FF' + value + '"/>' : '') + '</patternFill></fill></dxf>'
    : '<dxf><font><color ' + (value ? 'rgb="FF' + value + '"' : 'auto="1"') + '/></font></dxf>';
  const range = findElementRanges(xml, 'dxfs').get('');
  if (range) {
    const opening = setTagAttribute(xml.slice(range.start, range.startTagEnd).replace(/\/\s*>$/, '>'), 'count', nodes.length + 1);
    xml = xml.slice(0, range.start) + opening + (range.selfClosing ? '' : xml.slice(range.startTagEnd, range.closingStart)) + tag + '</dxfs>' + xml.slice(range.end);
  } else {
    const following = ['tableStyles', 'colors', 'extLst'].flatMap(name => [...findElementRanges(xml, name).values()].map(range => range.start));
    const at = following.length ? Math.min(...following) : xml.lastIndexOf('</styleSheet>');
    if (at < 0) fail('색상 필터 서식을 추가할 위치가 없습니다.');
    xml = xml.slice(0, at) + '<dxfs count="1">' + tag + '</dxfs>' + xml.slice(at);
  }
  return { xml, dxfId: nodes.length };
}

function readAutoFilter(node, differentialNodes = [], colorContext, allowHeaderOnly = false) {
  if (!node) return null;
  const reference = String(node.attributes.ref || '').replaceAll('$', '').split(':');
  if (allowHeaderOnly && reference.length === 1) reference.push(reference[0]);
  if (reference.length !== 2) throw new Error('AutoFilter 범위가 지원되는 셀 참조가 아닙니다.');
  const first = reference[0], last = reference[1], top = rowNumber(first) - 1, bottom = rowNumber(last) - 1, left = columnNumber(first) - 1, right = columnNumber(last) - 1;
  const area = (bottom - top + 1) * (right - left + 1);
  if (top < 0 || bottom < top || bottom === top && !allowHeaderOnly || left < 0 || right < left || bottom >= LIMITS.rows || right >= LIMITS.columns || !Number.isSafeInteger(area) || area > LIMITS.cells) throw new Error('AutoFilter 범위가 유효하지 않거나 크기 제한을 넘었습니다.');
  if (Object.keys(node.attributes).some(name => !['ref', 'filterMode'].includes(name))) throw new Error('지원하지 않는 AutoFilter 속성이 있습니다.');
  const columns = [];
  for (const filterColumn of node.children) {
    if (filterColumn.name !== 'filterColumn' || Object.keys(filterColumn.attributes).some(name => name !== 'colId')) throw new Error('지원하지 않는 AutoFilter 구성이 있습니다.');
    const offset = Number(filterColumn.attributes.colId);
    if (!Number.isInteger(offset) || offset < 0 || left + offset > right || columns.some(item => item.column === left + offset) || filterColumn.children.length !== 1 || !['filters', 'customFilters', 'top10', 'dynamicFilter', 'colorFilter', 'iconFilter'].includes(filterColumn.children[0].name)) throw new Error('AutoFilter 열 필터 구성이 유효하지 않습니다.');
    const filters = filterColumn.children[0];
    if (filters.name === 'iconFilter') {
      const attrs = filters.attributes, iconId = Number(attrs.iconId), iconSet = attrs.iconSet || '3TrafficLights1';
      if (filters.children.length || filters.text.trim() || Object.keys(attrs).some(key => !['iconSet', 'iconId'].includes(key)) || !/^\d+$/.test(attrs.iconId || '') || !Number.isInteger(iconId)) throw new Error('아이콘 필터의 집합이나 아이콘 번호가 유효하지 않습니다.');
      columns.push({ column: left + offset, iconFilter: validateIconFilter({ iconSet, iconId }) }); continue;
    }
    if (filters.name === 'colorFilter') {
      const attrs=filters.attributes,index=Number(attrs.dxfId);
      if(filters.children.length||filters.text.trim()||Object.keys(attrs).some(key=>!['dxfId','cellColor'].includes(key))||!/^\d+$/.test(attrs.dxfId||'')||!Number.isInteger(index)||!differentialNodes[index]||attrs.cellColor!==undefined&&!['0','1','true','false'].includes(attrs.cellColor))throw new Error('색상 필터의 차등 서식 참조가 유효하지 않습니다.');
      const cellColor=!['0','false'].includes(attrs.cellColor),color=readFilterColor(differentialNodes[index],cellColor,colorContext);
      columns.push({column:left+offset,colorFilter:validateColorFilter({cellColor,color,dxfId:index})});continue;
    }
    if (filters.name === 'dynamicFilter') {
      const dynamic = { type: filters.attributes.type };
      if (filters.children.length || filters.text.trim() || Object.keys(filters.attributes).some(key => !['type', 'val', 'maxVal'].includes(key))) throw new Error('동적 필터의 종류나 기준값이 유효하지 않습니다.');
      for (const [attribute, key] of [['val', 'value'], ['maxVal', 'maxValue']]) if (filters.attributes[attribute] !== undefined) {
        if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(filters.attributes[attribute].trim())) throw new Error('동적 필터의 종류나 기준값이 유효하지 않습니다.');
        dynamic[key] = Number(filters.attributes[attribute]);
      }
      columns.push({ column: left + offset, dynamic: validateDynamicFilter(dynamic) }); continue;
    }
    if (filters.name === 'top10') {
      const attrs = filters.attributes;
      const numeric = value => typeof value === 'string' && /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(value.trim());
      if (filters.children.length || filters.text.trim() || Object.keys(attrs).some(key => !['top', 'percent', 'val', 'filterVal'].includes(key)) || ['top', 'percent'].some(key => attrs[key] !== undefined && !['0', '1', 'false', 'true'].includes(attrs[key])) || !numeric(attrs.val) || attrs.filterVal !== undefined && !numeric(attrs.filterVal)) throw new Error('AutoFilter 순위 필터 구성이 유효하지 않습니다.');
      const top10 = { top: !['0', 'false'].includes(attrs.top), percent: ['1', 'true'].includes(attrs.percent), value: Number(attrs.val), ...(attrs.filterVal !== undefined ? { filterValue: Number(attrs.filterVal) } : {}) };
      columns.push({ column: left + offset, top10: validateTopFilter(top10) }); continue;
    }
    if (filters.name === 'customFilters') {
      if (Object.keys(filters.attributes).some(name => name !== 'and') || filters.attributes.and !== undefined && !['0', '1', 'false', 'true'].includes(filters.attributes.and)) throw new Error('AutoFilter 조건 조합이 유효하지 않습니다.');
      const custom = { and: ['1', 'true'].includes(filters.attributes.and), conditions: filters.children.map(condition => {
        if (condition.name !== 'customFilter' || condition.children.length || Object.keys(condition.attributes).some(name => !['operator', 'val'].includes(name)) || typeof condition.attributes.val !== 'string') throw new Error('AutoFilter 비교 조건이 유효하지 않습니다.');
        return { operator: condition.attributes.operator ?? 'equal', value: condition.attributes.val };
      }) };
      const entry = { column: left + offset, custom: validateCustomFilter(custom) }; compileAutoFilterColumn(entry); columns.push(entry); continue;
    }
    if (Object.keys(filters.attributes).some(name => !['blank', 'calendarType'].includes(name)) || filters.attributes.calendarType !== undefined && !['none', 'gregorian'].includes(filters.attributes.calendarType) || (filters.attributes.blank !== undefined && !['0', '1', 'false', 'true'].includes(filters.attributes.blank))) throw new Error('지원하지 않는 AutoFilter 값 목록 속성이 있습니다.');
    const values = [], dateGroups = [];
    for (const filter of filters.children) {
      if (filter.name === 'dateGroupItem') {
        const group = {};
        for (const [key, value] of Object.entries(filter.attributes)) {
          if (!['dateTimeGrouping', 'year', 'month', 'day', 'hour', 'minute', 'second'].includes(key)) throw new Error('날짜 그룹의 연·월·일·시·분·초가 유효하지 않습니다.');
          if (key !== 'dateTimeGrouping' && !/^\d+$/.test(value)) throw new Error('날짜 그룹의 연·월·일·시·분·초가 유효하지 않습니다.');
          group[key] = key === 'dateTimeGrouping' ? value : Number(value);
        }
        if (filter.children.length || filter.text.trim()) throw new Error('날짜 그룹의 연·월·일·시·분·초가 유효하지 않습니다.');
        dateGroups.push(group); continue;
      }
      if (filter.name !== 'filter' || Object.keys(filter.attributes).length !== 1 || typeof filter.attributes.val !== 'string' || filter.children.length) throw new Error('단순 값 목록 이외의 AutoFilter 조건은 지원하지 않습니다.');
      values.push(filter.attributes.val);
    }
    validateDateGroups(dateGroups);
    columns.push({ column: left + offset, values: [...new Set(values)], blanks: ['1', 'true'].includes(filters.attributes.blank), ...(dateGroups.length ? { dateGroups } : {}) });
  }
  return { top, bottom, left, right, columns };
}

function readMergedCells(node) {
  if (!node) return [];
  if (Object.keys(node.attributes).some(name => name !== 'count')) throw new Error('병합 셀 목록 속성이 유효하지 않습니다.');
  const entries = children(node, 'mergeCell');
  if (node.children.some(item => item.name !== 'mergeCell') || entries.length > 1000 || node.attributes.count !== undefined && (!/^(?:0|[1-9]\d*)$/.test(node.attributes.count) || Number(node.attributes.count) !== entries.length)) throw new Error('병합 셀 목록 개수가 유효하지 않습니다.');
  const ranges = entries.map(entry => {
    if (Object.keys(entry.attributes).length !== 1 || typeof entry.attributes.ref !== 'string' || entry.children.length) throw new Error('병합 셀 참조가 유효하지 않습니다.');
    const reference = /^\$?([A-Z]{1,3})\$?([1-9]\d*):\$?([A-Z]{1,3})\$?([1-9]\d*)$/i.exec(entry.attributes.ref);
    if (!reference) throw new Error(`병합 셀 범위가 올바르지 않습니다: ${entry.attributes.ref}`);
    const left = columnNumber(reference[1] + '1') - 1, top = Number(reference[2]) - 1, right = columnNumber(reference[3] + '1') - 1, bottom = Number(reference[4]) - 1;
    if (left < 0 || top < 0 || right < left || bottom < top || right >= LIMITS.columns || bottom >= LIMITS.rows || right === left && bottom === top) throw new Error(`병합 셀 범위가 유효하지 않습니다: ${entry.attributes.ref}`);
    return { top, bottom, left, right, ref: entry.attributes.ref };
  });
  for (let index = 0; index < ranges.length; index++) for (let other = 0; other < index; other++) {
    const range = ranges[index], previous = ranges[other];
    if (range.top <= previous.bottom && range.bottom >= previous.top && range.left <= previous.right && range.right >= previous.left) throw new Error('병합 셀 범위가 서로 겹칩니다.');
  }
  return ranges;
}

function readHiddenRows(sheetData) {
  const hiddenRows = [];
  for (const row of children(sheetData, 'row')) {
    const raw = row.attributes.hidden;
    if (raw !== undefined && !['0', 'false', 'off', '1', 'true', 'on'].includes(raw.toLocaleLowerCase())) throw new Error('숨김 행 속성이 유효하지 않습니다.');
    const zeroHeight = row.attributes.ht !== undefined && Number(row.attributes.ht) === 0;
    if (raw === undefined && !zeroHeight) continue;
    const value = zeroHeight ? '1' : raw.toLocaleLowerCase();
    if (['0', 'false', 'off'].includes(value)) continue;
    if (!['1', 'true', 'on'].includes(value) || !/^[1-9]\d*$/.test(row.attributes.r || '')) throw new Error('숨김 행 속성이 유효하지 않습니다.');
    const index = Number(row.attributes.r) - 1;
    if (index >= LIMITS.rows) throw new Error('숨김 행이 워크시트 범위를 벗어났습니다.');
    hiddenRows.push(index);
  }
  return [...new Set(hiddenRows)].sort((left, right) => left - right);
}

function readRowDimensions(root, sheetData) {
  const toPixels = points => Math.round(points * 4 / 3 * 100) / 100;
  const rawDefault = child(root, 'sheetFormatPr')?.attributes.defaultRowHeight;
  const defaultPoints = rawDefault === undefined ? null : Number(rawDefault);
  if (defaultPoints !== null && (!Number.isFinite(defaultPoints) || defaultPoints <= 0 || defaultPoints > 409.5)) throw new Error('기본 행 높이가 유효하지 않습니다.');
  const heights = {}, customRows = [];
  for (const row of children(sheetData, 'row')) {
    if (row.attributes.ht === undefined) continue;
    const points = Number(row.attributes.ht), index = Number(row.attributes.r) - 1;
    if (!Number.isFinite(points) || points < 0 || points > 409.5 || !Number.isInteger(index) || index < 0 || index >= LIMITS.rows) throw new Error('행 높이 또는 행 번호가 유효하지 않습니다.');
    if (points > 0) heights[index] = toPixels(points);
    if (['1', 'true', 'on'].includes(row.attributes.customHeight)) customRows.push(index);
  }
  return { heights, customRows, defaultHeight: defaultPoints === null ? 20 : toPixels(defaultPoints) };
}

function parseRangeReference(formula, defaultSheetName = null, maximumCells = 1000) {
  const reference = /^=?((?:'((?:[^']|'')+)'|([\p{L}_][\p{L}\p{N}_.]*))!)?\$?([A-Z]{1,3})\$?([1-9]\d*)(?::\$?([A-Z]{1,3})\$?([1-9]\d*))?$/iu.exec(formula);
  if (!reference) return null;
  const sheetName = reference[2]?.replaceAll("''", "'") || reference[3] || defaultSheetName;
  const first = reference[4].toUpperCase() + reference[5], last = (reference[6] || reference[4]).toUpperCase() + (reference[7] || reference[5]);
  const top = rowNumber(first) - 1, bottom = rowNumber(last) - 1, left = columnNumber(first) - 1, right = columnNumber(last) - 1;
  const count = (bottom - top + 1) * (right - left + 1);
  if (top < 0 || left < 0 || bottom < top || right < left || bottom >= LIMITS.rows || right >= LIMITS.columns || !Number.isSafeInteger(count) || count > maximumCells) return null;
  return { ...(sheetName ? { sheetName } : {}), top, bottom, left, right };
}

function parseTableRange(reference) {
  if (typeof reference === 'string' && !reference.includes(':')) reference += ':' + reference;
  const match = /^\$?([A-Z]{1,3})\$?([1-9]\d*):\$?([A-Z]{1,3})\$?([1-9]\d*)$/i.exec(reference || '');
  if (!match) return null;
  const left = columnNumber(match[1] + '1') - 1, top = Number(match[2]) - 1, right = columnNumber(match[3] + '1') - 1, bottom = Number(match[4]) - 1;
  if (left < 0 || top < 0 || right < left || bottom < top || right >= LIMITS.columns || bottom >= LIMITS.rows) return null;
  return { top, bottom, left, right };
}

const TABLE_STYLE_NAMES = Object.freeze([
  ...Array.from({ length: 21 }, (_, index) => 'TableStyleLight' + (index + 1)),
  ...Array.from({ length: 28 }, (_, index) => 'TableStyleMedium' + (index + 1)),
  ...Array.from({ length: 11 }, (_, index) => 'TableStyleDark' + (index + 1))
]);

function readTableStyleInfo(node) {
  if (!node) return null;
  if (node.children.length || node.text.trim()) throw new Error('표 스타일 정보의 구성요소가 유효하지 않습니다.');
  const name = node.attributes.name ?? null;
  if (name !== null && (!name.length || name.length > 255 || /[\u0000-\u001f]/.test(name))) throw new Error('표 스타일 이름이 유효하지 않습니다.');
  const readFlag = key => {
    const value = node.attributes[key];
    if (value === undefined) return false;
    if (['1', 'true'].includes(value)) return true;
    if (['0', 'false'].includes(value)) return false;
    throw new Error('표 스타일 옵션이 유효하지 않습니다.');
  };
  return { name, showFirstColumn: readFlag('showFirstColumn'), showLastColumn: readFlag('showLastColumn'), showRowStripes: readFlag('showRowStripes'), showColumnStripes: readFlag('showColumnStripes') };
}

function parsePrintAreas(text, sheetName) {
  const parts = []; let start = 0, quoted = false;
  for (let index = 0; index < text.length; index++) {
    if (text[index] === "'") {
      if (quoted && text[index + 1] === "'") { index++; continue; }
      quoted = !quoted;
    } else if (text[index] === ',' && !quoted) { parts.push(text.slice(start, index).trim()); start = index + 1; }
  }
  parts.push(text.slice(start).trim());
  if (quoted || parts.length > 16) return null;
  const sources = parts.map(part => {
    const reference = /^=?((?:'(?:[^']|'')+'|[A-Z_][A-Z0-9_.]*)!)?(.*)$/i.exec(part);
    const columns = /^\$?([A-Z]{1,3}):\$?([A-Z]{1,3})$/i.exec(reference?.[2] || '');
    const rows = /^\$?([1-9]\d*):\$?([1-9]\d*)$/.exec(reference?.[2] || '');
    let expanded = part;
    if (columns) expanded = (reference[1] || '') + columns[1] + '1:' + columns[2] + LIMITS.rows;
    else if (rows) expanded = (reference[1] || '') + 'A' + rows[1] + ':XFD' + rows[2];
    return parseRangeReference(expanded, sheetName, LIMITS.rows * LIMITS.columns);
  });
  return sources.every(source => source && source.sheetName?.toLocaleLowerCase() === sheetName?.toLocaleLowerCase()) ? sources : null;
}

function readDefinedNames(node, sheetNames, markUnsupported) {
  if (!node) return [];
  if (Object.keys(node.attributes).length || node.children.some(item => item.name !== 'definedName') || node.children.length > 1024) { markUnsupported('정의된 이름 목록 구성이 지원되지 않습니다.'); return []; }
  const names = [];
  for (const entry of node.children) {
    const name = entry.attributes.name, localRaw = entry.attributes.localSheetId;
    if (!name || !/^[A-Z_\\][A-Z0-9_.]*$/i.test(name) || Object.keys(entry.attributes).some(attribute => !['name', 'localSheetId'].includes(attribute)) || entry.children.length) { markUnsupported('지원되지 않는 이름 정의 속성이 있습니다.'); continue; }
    const localSheetId = localRaw === undefined ? null : Number(localRaw);
    if (localSheetId !== null && (!Number.isInteger(localSheetId) || localSheetId < 0 || localSheetId >= sheetNames.length)) { markUnsupported('이름 정의의 시트 범위가 유효하지 않습니다.'); continue; }
    const duplicate = names.some(item => item.name.toLocaleLowerCase() === name.toLocaleLowerCase() && item.localSheetId === localSheetId);
    if (entry.text.trim() === '#REF!') {
      if (duplicate) markUnsupported(`이름 정의가 중복되었습니다: ${name}`);
      else names.push({ name, localSheetId, source: null, error: '#REF!' });
      continue;
    }
    const isPrintArea = name.toLocaleLowerCase() === '_xlnm.print_area';
    const isPrintTitles = name.toLocaleLowerCase() === '_xlnm.print_titles';
    const printRanges = (isPrintArea || isPrintTitles) && localSheetId !== null ? parsePrintAreas(entry.text.trim(), sheetNames[localSheetId]) : null;
    const printTitles = isPrintTitles && printRanges ? { rows: null, columns: null } : null;
    let validTitles = Boolean(printTitles);
    if (printTitles) for (const range of printRanges) {
      if (range.left === 0 && range.right === LIMITS.columns - 1 && !printTitles.rows) printTitles.rows = { start: range.top, end: range.bottom };
      else if (range.top === 0 && range.bottom === LIMITS.rows - 1 && !printTitles.columns) printTitles.columns = { start: range.left, end: range.right };
      else validTitles = false;
    }
    const source = isPrintArea || isPrintTitles ? printRanges?.[0] : parseRangeReference(entry.text.trim(), localSheetId === null ? null : sheetNames[localSheetId]);
    if (!source || !source.sheetName || isPrintTitles && !validTitles || duplicate) { markUnsupported(`이름 정의를 셀 범위로 해석할 수 없습니다: ${name}`); continue; }
    names.push({ name, localSheetId, source: isPrintTitles || isPrintArea && printRanges.length > 1 ? null : source, ...(isPrintArea ? { printAreas: printRanges } : {}), ...(isPrintTitles ? { printTitles } : {}) });
  }
  return names;
}

function readDataValidations(node, definedNames = [], sheetIndex = 0, date1904 = false) {
  if (!node) return [];
  if (Object.keys(node.attributes).some(name => name !== 'count')) throw new Error('데이터 유효성 목록 속성이 지원되지 않습니다.');
  const entries = children(node, 'dataValidation');
  if (node.children.some(item => item.name !== 'dataValidation') || !entries.length || entries.length > 256 || node.attributes.count !== undefined && Number(node.attributes.count) !== entries.length) throw new Error('데이터 유효성 규칙 개수가 유효하지 않습니다.');
  return entries.map(entry => {
    const allowed = ['type', 'errorStyle', 'imeMode', 'operator', 'allowBlank', 'showDropDown', 'showInputMessage', 'showErrorMessage', 'errorTitle', 'error', 'promptTitle', 'prompt', 'sqref'];
    const type = entry.attributes.type, list = type === 'list';
    if (Object.keys(entry.attributes).some(name => !allowed.includes(name)) || !['list', 'whole', 'decimal', 'date', 'time', 'textLength'].includes(type) || !entry.attributes.sqref || !['0', '1', undefined].includes(entry.attributes.allowBlank) || !['0', '1', undefined].includes(entry.attributes.showDropDown) || list && (entry.attributes.operator || entry.attributes.imeMode) || !list && (entry.attributes.showDropDown !== undefined || entry.attributes.imeMode)) throw new Error('지원하지 않는 데이터 유효성 종류 또는 속성입니다.');
    if (entry.attributes.errorStyle && !['stop', 'warning', 'information'].includes(entry.attributes.errorStyle) || ['showInputMessage', 'showErrorMessage'].some(name => entry.attributes[name] !== undefined && !['0', '1'].includes(entry.attributes[name])) || ['errorTitle', 'promptTitle'].some(name => (entry.attributes[name] || '').length > 32) || ['error', 'prompt'].some(name => (entry.attributes[name] || '').length > 255)) throw new Error('데이터 유효성 메시지 또는 알림 규칙 속성이 유효하지 않습니다.');
    const operator = list ? null : entry.attributes.operator || 'between';
    const operators = ['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'];
    if (!list && !operators.includes(operator)) throw new Error('지원하지 않는 데이터 유효성 비교 연산자입니다.');
    const firstFormula = child(entry, 'formula1')?.text.trim(), secondFormula = child(entry, 'formula2')?.text.trim();
    const needsSecond = !list && ['between', 'notBetween'].includes(operator);
    if (!firstFormula || firstFormula.length > 255 || (needsSecond ? !secondFormula || secondFormula.length > 255 : secondFormula !== undefined) || entry.children.some(item => !['formula1', 'formula2'].includes(item.name)) || children(entry, 'formula1').length !== 1 || children(entry, 'formula2').length !== (needsSecond ? 1 : 0)) throw new Error('데이터 유효성 비교값 구성이 유효하지 않습니다.');
    const formula = firstFormula;
    let values = null, source = null, sourceFormula = null;
    if (list && formula[0] === '"' && formula.at(-1) === '"') {
      values = [];
      const raw = formula.slice(1, -1);
      let value = '';
      for (let index = 0; index < raw.length; index++) {
        if (raw[index] === '"') {
          if (raw[index + 1] !== '"') throw new Error('데이터 유효성 목록 인용부호가 유효하지 않습니다.');
          value += '"'; index++;
        } else if (raw[index] === ',') { values.push(value); value = ''; }
        else value += raw[index];
      }
      values.push(value);
      if (!values.length || values.length > 50 || values.some(item => item.length > 255)) throw new Error('데이터 유효성 목록 값 개수가 제한을 벗어났습니다.');
    } else if (list && formula === '#REF!') {
      sourceFormula = formula;
    } else if (list) {
      source = parseRangeReference(formula);
      if (!source) {
        const name = /^=?([A-Z_\\][A-Z0-9_.]*)$/i.exec(formula)?.[1];
        const resolved = name && (definedNames.find(item => item.localSheetId === sheetIndex && item.name.toLocaleLowerCase() === name.toLocaleLowerCase()) || definedNames.find(item => item.localSheetId === null && item.name.toLocaleLowerCase() === name.toLocaleLowerCase()));
        if (resolved) source = resolved.source;
        else {
          const indirect = /^=?INDIRECT\((.+)\)$/i.exec(formula);
          if (!indirect) throw new Error('셀 범위 또는 지원되는 이름 정의 이외의 데이터 유효성 목록 원본은 지원되지 않습니다.');
          const argument = indirect[1].trim(), literal = /^"((?:[^"]|"")*)"$/.exec(argument);
          if (literal) {
            const reference = literal[1].replaceAll('""', '"');
            const directSource = parseRangeReference(reference);
            const directName = /^([A-Z_\\][A-Z0-9_.]*)$/i.exec(reference)?.[1];
            const directDefinition = directName && (definedNames.find(item => item.localSheetId === sheetIndex && item.name.toLocaleLowerCase() === directName.toLocaleLowerCase()) || definedNames.find(item => item.localSheetId === null && item.name.toLocaleLowerCase() === directName.toLocaleLowerCase()));
            if (!directSource && !directDefinition) throw new Error('INDIRECT 데이터 유효성 목록 원본이 셀 범위 또는 정의된 이름이 아닙니다.');
          } else {
            const sample = evaluateFormula(argument, { date1904, getCell: () => 'A1:A2', getSparseRange: () => [] });
            if (typeof sample !== 'string' || sample.startsWith('#')) throw new Error('지원하지 않는 INDIRECT 데이터 유효성 원본 수식입니다.');
          }
          sourceFormula = formula.startsWith('=') ? formula.slice(1) : formula;
        }
      }
    } else {
      const rawLimits = [firstFormula, ...(needsSecond ? [secondFormula] : [])];
      values = []; const formulas = [];
      for (const raw of rawLimits) {
        if (/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(raw) && Number.isFinite(Number(raw))) { values.push(Number(raw)); formulas.push(null); continue; }
        const expression = raw.startsWith('=') ? raw.slice(1) : raw;
        const sample = evaluateFormula(expression, { date1904, getCell: () => 1, getSparseRange: () => [] });
        if (typeof sample === 'string' && ['#NAME?', '#VALUE!', '#REF!'].includes(sample) && !(sample === '#REF!' && expression.includes('#REF!'))) throw new Error('지원하지 않는 데이터 유효성 기준 수식입니다.');
        values.push(null); formulas.push(expression);
      }
      var formulaLimits = formulas.some(Boolean) ? formulas : null;
    }
    const ranges = String(entry.attributes.sqref).trim().split(/\s+/).map(reference => {
      const parts = reference.replaceAll('$', '').split(':');
      if (parts.length > 2) throw new Error('데이터 유효성 적용 범위가 유효하지 않습니다.');
      const first = parts[0], last = parts[1] || parts[0];
      const top = rowNumber(first) - 1, bottom = rowNumber(last) - 1, left = columnNumber(first) - 1, right = columnNumber(last) - 1;
      if (top < 0 || left < 0 || bottom < top || right < left || bottom >= LIMITS.rows || right >= LIMITS.columns) throw new Error('데이터 유효성 적용 범위가 시트 제한을 벗어났습니다.');
      return { top, bottom, left, right };
    });
    if (!ranges.length || ranges.length > 64) throw new Error('데이터 유효성 적용 범위 개수가 제한을 벗어났습니다.');
    return { type, operator, ranges, values, formulas: formulaLimits || null, source, sourceFormula, errorStyle: entry.attributes.errorStyle || 'stop', allowBlank: entry.attributes.allowBlank === '1', showDropDown: entry.attributes.showDropDown === '1', showInputMessage: entry.attributes.showInputMessage !== '0', showErrorMessage: entry.attributes.showErrorMessage !== '0', errorTitle: entry.attributes.errorTitle || '', error: entry.attributes.error || '', promptTitle: entry.attributes.promptTitle || '', prompt: entry.attributes.prompt || '' };
  });
}
function isoDateSerial(raw, date1904) {
  if (raw.length > 128) fail('ISO 날짜 셀의 날짜·시각이 유효하지 않습니다.');
  const match = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}):(\d{2})(\.\d+)?)?(Z|[+-]\d{2}:\d{2})?$/.exec(raw);
  if (!match) fail('ISO 날짜 셀의 날짜·시각이 유효하지 않습니다.');
  const [, yearText, monthText, dayText, hourText, minuteText, secondText, fractionText, zone] = match;
  const year=Number(yearText), month=Number(monthText), day=Number(dayText);
  const hour=Number(hourText || 0), minute=Number(minuteText || 0), second=Number(secondText || 0), fraction=Number(fractionText || 0);
  const date=new Date(0); date.setUTCFullYear(year,month-1,day); date.setUTCHours(0,0,0,0);
  if (year < 1 || date.getUTCFullYear() !== year || date.getUTCMonth() !== month-1 || date.getUTCDate() !== day || hour > 24 || minute > 59 || second > 59 || hour === 24 && (minute || second || fraction)) fail('ISO 날짜 셀의 날짜·시각이 유효하지 않습니다.');
  let offset=0;
  if (zone && zone !== 'Z') {
    const hours=Number(zone.slice(1,3)), minutes=Number(zone.slice(4,6));
    if (hours > 14 || minutes > 59 || hours === 14 && minutes) fail('ISO 날짜 셀의 시간대가 유효하지 않습니다.');
    offset=(zone[0] === '-' ? -1 : 1)*(hours*60+minutes);
  }
  const instant=date.getTime()+(hour*3600+minute*60+second-offset*60)*1000;
  if (instant < -62135596800000 || instant >= 253402300800000) fail('ISO 날짜 셀이 지원하는 연도 범위를 벗어났습니다.');
  const epoch=date1904 ? Date.UTC(1904,0,1) : Date.UTC(1899,11,31);
  return (instant-epoch)/86400000+fraction/86400+(!date1904 && instant >= Date.UTC(1900,2,1) ? 1 : 0);
}

function decodeCellValue(cell, sharedStrings, date1904 = false) {
  const type = cell.attributes.t || 'n';
  if (type === 'inlineStr') return { value: textContent(child(cell, 'is') || { text: '', children: [] }), kind: 'string' };
  const raw = child(cell, 'v')?.text ?? '';
  if (type === 'd') return { value: String(isoDateSerial(raw.trim(), date1904)), kind: 'number', isoDate: raw };
  if (type === 's') {
    const index = Number(raw);
    if (!Number.isInteger(index) || index < 0 || index >= sharedStrings.length) fail('공유 문자열 인덱스가 유효하지 않습니다.');
    return { value: sharedStrings[index], kind: 'string' };
  }
  if (type === 'b') return { value: raw === '1' ? 'TRUE' : 'FALSE', kind: 'boolean' };
  if (type === 'e') return { value: raw, kind: 'error' };
  if (type === 'str') return { value: raw, kind: 'string' };
  if (type === 'n') return { value: raw, kind: raw === '' ? 'blank' : 'number' };
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

function styleColor(node, colorContext) { return resolveStyleColor(node, colorContext); }

function readCellStyleProperties(xf, fonts, fills, borders, colorContext) {
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
  if (font?.family) properties.fontFamily = font.family;
  if (fill?.color) properties.fill = fill.color;
  const horizontal = alignment?.attributes?.horizontal;
  if (['left', 'center', 'right', 'justify'].includes(horizontal)) properties.horizontal = horizontal;
  const vertical = alignment?.attributes?.vertical;
  if (['top', 'center', 'bottom'].includes(vertical)) properties.vertical = vertical;
  if (['1', 'true', 'on'].includes(alignment?.attributes?.wrapText)) properties.wrapText = true;
  if (border) {
    const sides = {};
    for (const side of ['left', 'right', 'top', 'bottom']) {
      const edge = child(border, side);
      const color = styleColor(child(edge || { children: [] }, 'color'), colorContext);
      const line = edge?.attributes?.style;
      if (['hair', 'thin', 'medium', 'thick', 'dashed', 'dotted', 'double'].includes(line)) {
        sides[side] = { color: color || '#000000', line };
      }
    }
    if (Object.keys(sides).length) properties.borders = sides;
  }
  return Object.keys(properties).length ? properties : null;
}

function readDifferentialStyle(dxf, colorContext) {
  if (children(dxf, 'numFmt').length || dxf.children.some(item => !['font', 'fill', 'border', 'alignment'].includes(item.name))) throw new Error('지원하지 않는 차등 서식 구성요소입니다.');
  const properties = {};
  const font = child(dxf, 'font');
  if (font) {
    if (font.children.some(item => !['b', 'i', 'u', 'strike', 'color'].includes(item.name))) throw new Error('지원하지 않는 조건부 글꼴 서식입니다.');
    for (const [tag, property] of [['b', 'bold'], ['i', 'italic'], ['u', 'underline'], ['strike', 'strike']]) {
      const element = child(font, tag);
      if (element && Object.keys(element.attributes).some(attribute => attribute !== 'val')) throw new Error('조건부 글꼴 속성이 지원되지 않습니다.');
      if (element) properties[property] = !['0', 'false'].includes(String(element.attributes.val ?? '').toLowerCase());
    }
    const colorNode = child(font, 'color');
    if (colorNode && (!styleColor(colorNode, colorContext))) throw new Error('조건부 글꼴 색상 정의를 해석할 수 없습니다.');
    const color = styleColor(colorNode, colorContext);
    if (color) properties.color = color;
  }
  const fill = child(dxf, 'fill'), pattern = fill && child(fill, 'patternFill');
  if (fill && (!pattern || pattern.attributes.patternType !== 'solid' || pattern.children.some(item => item.name !== 'fgColor'))) throw new Error('조건부 채우기 서식은 단색 전경색만 지원합니다.');
  if (pattern) {
    const colorNode = child(pattern, 'fgColor');
    if (!colorNode || !styleColor(colorNode, colorContext)) throw new Error('조건부 채우기 색상 정의를 해석할 수 없습니다.');
    const color = styleColor(colorNode, colorContext);
    if (color) properties.fill = color;
  }
  const alignment = child(dxf, 'alignment');
  if (alignment && Object.keys(alignment.attributes).some(attribute => !['horizontal', 'vertical'].includes(attribute))) throw new Error('조건부 맞춤은 가로·세로 정렬만 지원합니다.');
  if (['left', 'center', 'right', 'justify'].includes(alignment?.attributes.horizontal)) properties.horizontal = alignment.attributes.horizontal;
  if (['top', 'center', 'bottom'].includes(alignment?.attributes.vertical)) properties.vertical = alignment.attributes.vertical;
  const border = child(dxf, 'border');
  if (border) {
    if (border.children.some(item => !['left', 'right', 'top', 'bottom'].includes(item.name))) throw new Error('조건부 테두리는 셀 바깥 네 방향만 지원합니다.');
    const sides = {};
    for (const side of ['left', 'right', 'top', 'bottom']) {
      const edge = child(border, side), line = edge?.attributes.style;
      if (edge && Object.keys(edge.attributes).some(attribute => attribute !== 'style')) throw new Error('조건부 테두리 속성이 지원되지 않습니다.');
      if (edge && child(edge, 'color') && (!styleColor(child(edge, 'color'), colorContext))) throw new Error('조건부 테두리 색상 정의를 해석할 수 없습니다.');
      if (edge && line && !['hair', 'thin', 'medium', 'thick', 'dashed', 'dotted', 'double'].includes(line)) throw new Error('조건부 테두리 선 종류가 지원되지 않습니다.');
      if (!['hair', 'thin', 'medium', 'thick', 'dashed', 'dotted', 'double'].includes(line)) continue;
      sides[side] = { line, color: styleColor(child(edge, 'color'), colorContext) || '#000000' };
    }
    if (Object.keys(sides).length) properties.borders = sides;
  }
  return Object.keys(properties).length ? properties : null;
}

function readConditionalFormats(node, differentialStyles, colorContext) {
  if (!node) return [];
  if (Object.keys(node.attributes).some(attribute => attribute !== 'sqref')) throw new Error('조건부 서식 범위 속성이 지원되지 않습니다.');
  const references = String(node.attributes.sqref || '').trim().split(/\s+/).filter(Boolean);
  if (!references.length || references.length > 64) throw new Error('조건부 서식 범위가 비어 있거나 너무 많습니다.');
  const ranges = references.map(reference => {
    const parts = reference.replaceAll('$', '').split(':');
    if (parts.length > 2) throw new Error('조건부 서식 범위 참조가 유효하지 않습니다.');
    const first = parts[0], last = parts[1] || parts[0];
    const top = rowNumber(first) - 1, bottom = rowNumber(last) - 1, left = columnNumber(first) - 1, right = columnNumber(last) - 1;
    const area = (bottom - top + 1) * (right - left + 1);
    if (top < 0 || left < 0 || bottom < top || right < left || bottom >= LIMITS.rows || right >= LIMITS.columns || !Number.isSafeInteger(area)) throw new Error('조건부 서식 범위가 시트 제한을 벗어났습니다.');
    return { top, bottom, left, right };
  });
  const rules = children(node, 'cfRule');
  if (!rules.length || rules.length > 128) throw new Error('조건부 서식 규칙 개수 제한을 벗어났습니다.');
  return rules.map(rule => {
    const { type, operator, priority, dxfId, stopIfTrue } = rule.attributes;
    if (Object.keys(rule.attributes).some(attribute => !['type', 'operator', 'priority', 'dxfId', 'stopIfTrue', 'rank', 'percent', 'bottom'].includes(attribute)) || type !== 'top10' && ['rank', 'percent', 'bottom'].some(attribute => Object.hasOwn(rule.attributes, attribute)) || !['cellIs', 'expression', 'dataBar', 'colorScale', 'iconSet', 'top10'].includes(type) || type === 'cellIs' && !['equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual', 'between', 'notBetween'].includes(operator) || type !== 'cellIs' && operator !== undefined) throw new Error('지원하지 않는 조건부 서식 규칙 종류 또는 연산자입니다.');
    const rank = Number(priority), styleIndex = Number(dxfId);
    const visualRule = type === 'dataBar' || type === 'colorScale' || type === 'iconSet';
    if (!/^[1-9]\d*$/.test(priority || '') || !Number.isSafeInteger(rank) || rank > 100000 || (visualRule ? dxfId !== undefined : !/^(?:0|[1-9]\d*)$/.test(dxfId || '') || !Number.isInteger(styleIndex) || !differentialStyles[styleIndex])) throw new Error('조건부 서식 우선 순위 또는 차등 서식 참조가 유효하지 않습니다.');
    if (stopIfTrue !== undefined && !['0', '1'].includes(stopIfTrue)) throw new Error('조건부 서식 stopIfTrue 값이 유효하지 않습니다.');
    if (type === 'top10') {
      const { rank: rawRank = '10', percent: rawPercent = '0', bottom: rawBottom = '0' } = rule.attributes;
      const topRank = Number(rawRank), percent = rawPercent === '1', bottom = rawBottom === '1';
      if (rule.children.length || !/^[1-9]\d*$/.test(rawRank) || !Number.isInteger(topRank) || topRank > (percent ? 100 : 1000) || !['0', '1'].includes(rawPercent) || !['0', '1'].includes(rawBottom)) throw new Error('상위·하위 조건부 서식의 순위 설정이 유효하지 않습니다.');
      return { ranges, type, priority: rank, stopIfTrue: stopIfTrue === '1', top10: { rank: topRank, percent, bottom }, dxfId: styleIndex, style: differentialStyles[styleIndex] };
    }
    if (type === 'iconSet') {
      const iconSet = child(rule, 'iconSet'), attrs = iconSet?.attributes || {}, name = attrs.iconSet || '3TrafficLights1', iconCount = ICON_SET_LENGTHS[name];
      const custom = attrs.custom === '1', customIcons = children(iconSet || { children: [] }, 'cfIcon');
      if (!iconSet || rule.children.length !== 1 || Object.keys(attrs).some(attribute => !['iconSet', 'showValue', 'percent', 'reverse', 'custom'].includes(attribute)) || !iconCount || ['showValue', 'percent', 'reverse', 'custom'].some(attribute => attrs[attribute] !== undefined && !['0', '1'].includes(attrs[attribute])) || iconSet.children.some(item => !['cfvo', 'cfIcon'].includes(item.name)) || custom !== (customIcons.length > 0)) throw new Error('지원하지 않는 아이콘 집합 조건부 서식 구성입니다.');
      const customMappings = customIcons.map(configuration => {
        const { iconSet: mappedSet, iconId: mappedId } = configuration.attributes, id = Number(mappedId);
        if (configuration.children.length || configuration.text.trim() || Object.keys(configuration.attributes).some(attribute => !['iconSet', 'iconId'].includes(attribute)) || !Object.hasOwn(ICON_SET_LENGTHS, mappedSet) || !/^(?:0|[1-9]\d*)$/.test(mappedId || '') || !Number.isInteger(id) || id >= ICON_SET_LENGTHS[mappedSet]) throw new Error('사용자 지정 아이콘 번호 또는 집합이 유효하지 않습니다.');
        return { iconSet: mappedSet, iconId: id };
      });
      const thresholds = children(iconSet, 'cfvo').map(configuration => {
        const { type: stopType, val, gte } = configuration.attributes;
        if (configuration.children.length || Object.keys(configuration.attributes).some(attribute => !['type', 'val', 'gte'].includes(attribute)) || gte !== undefined && !['0', '1'].includes(gte)) throw new Error('아이콘 기준 구성이 유효하지 않습니다.');
        if (stopType === 'min' || stopType === 'max') {
          if (val !== undefined && val !== '0' || configuration.text.trim()) throw new Error('아이콘 최솟값·최댓값 기준이 유효하지 않습니다.');
          return { type: stopType, value: null, gte: gte !== '0' };
        }
        if (stopType === 'formula') {
          const formula = String(val ?? configuration.text).trim().replace(/^=/, '');
          if (!formula || formula.length > 256 || /[!\[\]]/.test(formula)) throw new Error('아이콘 기준 수식은 256자 이하의 같은 시트 수식이어야 합니다.');
          const sample = evaluateFormula(formula, { getCell: () => 1 });
          if (sample === '#NAME?' || sample === '#VALUE!') throw new Error('아이콘 기준 수식에 지원하지 않는 함수가 있습니다.');
          return { type: stopType, formula, value: null, gte: gte !== '0' };
        }
        if (!['num', 'percent', 'percentile'].includes(stopType) || typeof val !== 'string' || !/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(val) || !Number.isFinite(Number(val)) || stopType !== 'num' && (Number(val) < 0 || Number(val) > 100) || configuration.text.trim()) throw new Error('지원하지 않는 아이콘 기준입니다.');
        return { type: stopType, value: Number(val), gte: gte !== '0' };
      });
      if (thresholds.length !== iconCount || custom && customMappings.length !== iconCount) throw new Error('아이콘 집합과 기준 개수가 일치하지 않습니다.');
      return { ranges, type, priority: rank, stopIfTrue: stopIfTrue === '1', iconSet: { name, showValue: attrs.showValue !== '0', reverse: attrs.reverse === '1', thresholds, ...(custom ? { customIcons: customMappings } : {}) } };
    }
    if (type === 'dataBar') {
      const dataBar = child(rule, 'dataBar'), configurations = children(dataBar || { children: [] }, 'cfvo'), colors = children(dataBar || { children: [] }, 'color');
      if (!dataBar || rule.children.length !== 1 || Object.keys(dataBar.attributes).some(attribute => !['minLength', 'maxLength', 'showValue', 'gradient'].includes(attribute)) || dataBar.children.some(item => !['cfvo', 'color'].includes(item.name)) || configurations.length !== 2 || colors.length !== 1 || dataBar.children.filter(item => item.name === 'cfvo').length !== 2 || dataBar.children.filter(item => item.name === 'color').length !== 1) throw new Error('지원하지 않는 데이터 막대 구성입니다.');
      const [minimum, maximum] = configurations.map(item => item.attributes.type), color = styleColor(colors[0], colorContext);
      if (minimum !== 'min' || maximum !== 'max' || !color || configurations.some(item => Object.keys(item.attributes).some(attribute => attribute !== 'type' && attribute !== 'val') || item.children.length || item.attributes.val !== undefined && item.attributes.val !== '0') || colors[0].children.length) throw new Error('데이터 막대의 기준 또는 색상이 유효하지 않습니다.');
      const minLength = Number(dataBar.attributes.minLength ?? 0), maxLength = Number(dataBar.attributes.maxLength ?? 100);
      if (![minLength, maxLength].every(value => Number.isInteger(value) && value >= 0 && value <= 100) || minLength > maxLength || !['0', '1', undefined].includes(dataBar.attributes.showValue) || !['0', '1', undefined].includes(dataBar.attributes.gradient)) throw new Error('데이터 막대 길이 또는 표시 설정이 유효하지 않습니다.');
      return { ranges, type, priority: rank, stopIfTrue: stopIfTrue === '1', dataBar: { color, minLength, maxLength, showValue: dataBar.attributes.showValue !== '0', gradient: dataBar.attributes.gradient !== '0' } };
    }
    if (type === 'colorScale') {
      const colorScale = child(rule, 'colorScale'), configurations = children(colorScale || { children: [] }, 'cfvo'), colors = children(colorScale || { children: [] }, 'color');
      if (!colorScale || rule.children.length !== 1 || Object.keys(colorScale.attributes).length || colorScale.children.some(item => !['cfvo', 'color'].includes(item.name)) || (configurations.length !== 2 && configurations.length !== 3) || colors.length !== configurations.length || colorScale.children.filter(item => item.name === 'cfvo').length !== configurations.length || colorScale.children.filter(item => item.name === 'color').length !== colors.length) throw new Error('지원하지 않는 색조 조건부 서식 구성입니다.');
      const stops = configurations.map(configuration => {
        const { type: stopType, val } = configuration.attributes;
        if (configuration.children.length || Object.keys(configuration.attributes).some(attribute => attribute !== 'type' && attribute !== 'val')) throw new Error('색조 기준 구성이 유효하지 않습니다.');
        if (stopType === 'min' || stopType === 'max') {
          if (val !== undefined && val !== '0') throw new Error('색조 최소·최대 기준값이 유효하지 않습니다.');
          return { type: stopType, value: null };
        }
        if (!['num', 'percent', 'percentile'].includes(stopType) || typeof val !== 'string' || !/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(val) || !Number.isFinite(Number(val)) || (stopType !== 'num' && (Number(val) < 0 || Number(val) > 100))) throw new Error('지원하지 않는 색조 기준입니다.');
        return { type: stopType, value: Number(val) };
      });
      const parsedColors = colors.map(color => {
        const value = styleColor(color, colorContext);
        if (!value || color.children.length) throw new Error('색조 색상 값이 유효하지 않습니다.');
        return value;
      });
      if (stops[0].type !== 'min' || stops.at(-1).type !== 'max' || stops.length === 3 && !['num', 'percent', 'percentile'].includes(stops[1].type)) throw new Error('색조 기준 순서가 지원되지 않습니다.');
      return { ranges, type, priority: rank, stopIfTrue: stopIfTrue === '1', colorScale: { stops, colors: parsedColors } };
    }
    const formulas = children(rule, 'formula');
    const expected = type === 'expression' ? 1 : operator === 'between' || operator === 'notBetween' ? 2 : 1;
    if (formulas.length !== expected || rule.children.some(item => item.name !== 'formula')) throw new Error('조건부 서식 비교값 개수가 유효하지 않습니다.');
    if (type === 'expression') {
      const formula = formulas[0].text.trim().replace(/^=/, '');
      if (!formula || formula.length > 256 || /[!\[\]]/.test(formula.replaceAll('#REF!', ''))) throw new Error('조건부 서식 수식은 256자 이하의 같은 시트 셀·범위 참조만 지원합니다.');
      const sample = evaluateFormula(formula, { getCell: () => 1 });
      if (sample === '#NAME?' || sample === '#VALUE!') throw new Error('조건부 서식 수식에 지원하지 않는 함수 또는 참조가 있습니다.');
      return { ranges, type, formula, priority: rank, stopIfTrue: stopIfTrue === '1', dxfId: styleIndex, style: differentialStyles[styleIndex] };
    }
    const values = formulas.map(formula => {
      const text = formula.text.trim().replace(/^=/, '');
      if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(text) || !Number.isFinite(Number(text))) throw new Error('조건부 서식 비교값은 숫자 상수여야 합니다.');
      return Number(text);
    });
    return { ranges, type, operator, values, priority: rank, stopIfTrue: stopIfTrue === '1', dxfId: styleIndex, style: differentialStyles[styleIndex] };
  }).sort((a, b) => a.priority - b.priority);
}

function conditionalStyleForCell(sheet,row,column,cell,getCellValue){return resolveConditionalStyle(sheet,row,column,cell,getCellValue,evaluateFormula)}

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
function formatExcelDate(serial, format, language, date1904 = false) {
  const parts = serialDateParts(serial, date1904);
  if (!parts) return null;
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
function formatExcelNumber(value, numFmtId, formatCode, language, date1904 = false) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return String(value);
  const format = formatCode || BUILTIN_NUMBER_FORMATS[numFmtId];
  if (!format || /^General$/i.test(format)) return String(value);
  const dateSyntax = format.replace(/"[^"]*"|\[[^\]]*\]|\\.|_.|\*./g, '');
  if (DATE_FORMAT_IDS.has(numFmtId) || /(?:y{2,4}|d{1,4}|AM\/PM|h{1,2}:[mhs]|m{1,2}:s{1,2})/i.test(dateSyntax)) {
    const dateCode = language === 'ko' && [14, 15, 16, 17, 22].includes(numFmtId)
      ? ({ 14: 'yyyy-mm-dd', 15: 'yyyy-mm-dd', 16: 'm-d', 17: 'yyyy-mm', 22: 'yyyy-mm-dd h:mm' }[numFmtId])
      : format;
    return formatExcelDate(numeric, dateCode, language, date1904) ?? String(value);
  }
  if (/[Ee][+-]?0/.test(format) || /[?]/.test(format)) return String(value);
  const sections = format.split(';');
  const selected = numeric < 0 && sections[1] ? sections[1] : sections[0];
  const hasCurrency = numFmtId >= 5 && numFmtId <= 8 || /[$€£¥₩]|USD/.test(selected);
  const unquoted = selected.replace(/"[^"]*"|\\./g, '');
  const percent = unquoted.includes('%');
  const literals = [...selected.matchAll(/"([^"]*)"/g)].map(match => match[1]);
  if (literals.some(literal => /[^ $€£¥₩]/.test(literal))) return String(value);
  const cleaned = selected.replace(/"[^"]*"|\[[^\]]*\]|\\.|_.|\*./g, '').replace(/[%$€£¥₩]/g, '').replace(/USD/g, '');
  const numericPattern = cleaned.replace(/[()]/g, '');
  if (!/[0#]/.test(numericPattern) || !/^[#0,.+\-]*$/.test(numericPattern)) return String(value);
  const decimal = /\.([0#]+)/.exec(numericPattern)?.[1] || '';
  const minimumFractionDigits = (decimal.match(/0/g) || []).length;
  const maximumFractionDigits = Math.min(20, decimal.length);
  const options = { useGrouping: /#,##0/.test(numericPattern), minimumFractionDigits, maximumFractionDigits };
  const currencySymbol = literals.join('') || /[$€£¥₩]/.exec(selected)?.[0];
  const currency = ({ '$': 'USD', '€': 'EUR', '£': 'GBP', '¥': 'JPY', '₩': 'KRW' })[currencySymbol] || 'USD';
  if (hasCurrency) { options.style = 'currency'; options.currency = currency; }
  else if (percent) options.style = 'percent';
  let displayed = new Intl.NumberFormat(language === 'ko' ? 'ko-KR' : 'en-US', options).format(numeric);
  if (numeric < 0 && sections[1] && /^\s*\([^)]*\)\s*$/.test(sections[1].replace(/\[[^\]]*\]/g, ''))) displayed = '(' + new Intl.NumberFormat(language === 'ko' ? 'ko-KR' : 'en-US', { ...options, style: hasCurrency ? 'currency' : 'decimal' }).format(Math.abs(numeric)) + ')';
  return displayed;
}
function formatCellValue(cell, language = 'en') {
  if (!cell || cell.kind !== 'number') return cell?.value ?? '';
  return formatExcelNumber(cell.value, Number(cell.numFmtId) || 0, cell.formatCode || null, language, Boolean(cell.date1904));
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
  const date1904 = ['1', 'true', 'on'].includes(child(workbook, 'workbookPr')?.attributes.date1904);
  const workbookRelsPath = 'xl/_rels/workbook.xml.rels';
  const workbookRels = parseXml(files[workbookRelsPath], workbookRelsPath).root;
  const workbookMap = parseRelationships(workbookRels, 'xl', ['worksheet', 'styles', 'sharedStrings', 'theme']);
  if ([...rootMap.values(), ...workbookMap.values()].some(rel => rel.external)) markUnsupported('외부 링크 또는 외부 관계가 포함되어 있습니다.');
  const sheetsNode = child(workbook, 'sheets');
  if (!sheetsNode) fail('통합문서에 시트가 없습니다.');
  const sheetRefs = children(sheetsNode, 'sheet');
  if (!sheetRefs.length || sheetRefs.length > LIMITS.sheets) fail('워크시트 개수 제한을 벗어났습니다.');
  if (child(workbook, 'externalReferences')) markUnsupported('외부 통합문서 참조가 포함되어 있습니다.');
  const sheetNames = new Set(), sheetEntries = [];
  const definedNames = readDefinedNames(child(workbook, 'definedNames'), sheetRefs.map(item => item.attributes.name || ''), markUnsupported);
  const printAreaDefinitions = children(child(workbook, 'definedNames') || { children: [] }, 'definedName').filter(entry => entry.attributes.name?.toLocaleLowerCase() === '_xlnm.print_area');
  const printTitleDefinitions = children(child(workbook, 'definedNames') || { children: [] }, 'definedName').filter(entry => entry.attributes.name?.toLocaleLowerCase() === '_xlnm.print_titles');
  let anyFormula = false, totalCells = 0;
  const sharedStrings = [];
  if (files['xl/sharedStrings.xml']) {
    known.add('xl/sharedStrings.xml');
    const shared = parseXml(files['xl/sharedStrings.xml'], 'xl/sharedStrings.xml').root;
    for (const si of children(shared, 'si')) sharedStrings.push(textContent(si));
  }
  const relationshipType = suffix => [...workbookMap].find(([, rel]) => rel.Type.endsWith(`/${suffix}`))?.[1].path;
  const stylesPath = relationshipType('styles');
  const themePath = relationshipType('theme');
  const themeRoot = themePath && files[themePath] ? parseXml(files[themePath], themePath).root : null;
  const parsedStyles = stylesPath && files[stylesPath] ? parseXml(files[stylesPath], stylesPath).root : null;
  const colorContext = createColorContext(themeRoot, parsedStyles);
  const styleFormats = [], differentialStyles = [], differentialNodes = [];
  let normalFont = { family: 'Calibri', size: 11, bold: false, italic: false };
  if (stylesPath && files[stylesPath]) {
    known.add(stylesPath);
    const styleRoot = parsedStyles;
    const customFormats = new Map(children(child(styleRoot, 'numFmts') || { children: [] }, 'numFmt').map(item => [Number(item.attributes.numFmtId), item.attributes.formatCode || '']));
    const fonts = children(child(styleRoot, 'fonts') || { children: [] }, 'font').map(font => {
      const size = Number(child(font, 'sz')?.attributes?.val);
      return {
        bold: Boolean(child(font, 'b') && !['0', 'false'].includes(child(font, 'b').attributes?.val)),
        italic: Boolean(child(font, 'i') && !['0', 'false'].includes(child(font, 'i').attributes?.val)),
        underline: Boolean(child(font, 'u') && !['0', 'false'].includes(child(font, 'u').attributes?.val)),
        strike: Boolean(child(font, 'strike') && !['0', 'false'].includes(child(font, 'strike').attributes?.val)),
        color: styleColor(child(font, 'color'), colorContext),
        size: Number.isFinite(size) && size >= 1 && size <= 409 ? size : null,
        family: child(font, 'name')?.attributes?.val || null,
        scheme: child(font, 'scheme')?.attributes?.val || null
      };
    });
    const normalStyle = children(child(styleRoot, 'cellStyles') || { children: [] }, 'cellStyle').find(item => (item.attributes.name || '').toLocaleLowerCase() === 'normal');
    const baseXfs = children(child(styleRoot, 'cellStyleXfs') || { children: [] }, 'xf');
    const styleXfId = Number(normalStyle?.attributes.xfId);
    const styleXf = baseXfs[Number.isInteger(styleXfId) && styleXfId >= 0 && styleXfId < baseXfs.length ? styleXfId : 0] || baseXfs[0];
    const normalFontId = Number(styleXf?.attributes.fontId);
    const normalFontDefinition = fonts[Number.isInteger(normalFontId) && normalFontId >= 0 && normalFontId < fonts.length ? normalFontId : 0] || {};
    const fontScheme = child(child(themeRoot || { children: [] }, 'themeElements') || { children: [] }, 'fontScheme');
    const themeFont = normalFontDefinition.scheme === 'major' ? child(fontScheme || { children: [] }, 'majorFont') : normalFontDefinition.scheme === 'minor' ? child(fontScheme || { children: [] }, 'minorFont') : null;
    normalFont = {
      family: normalFontDefinition.family || child(themeFont || { children: [] }, 'latin')?.attributes.typeface || 'Calibri',
      size: normalFontDefinition.size ?? 11,
      bold: Boolean(normalFontDefinition.bold),
      italic: Boolean(normalFontDefinition.italic)
    };
    const fills = children(child(styleRoot, 'fills') || { children: [] }, 'fill').map(fill => {
      const pattern = child(fill, 'patternFill');
      return { color: pattern?.attributes?.patternType === 'solid' ? styleColor(child(pattern, 'fgColor'), colorContext) : null };
    });
    const borders = children(child(styleRoot, 'borders') || { children: [] }, 'border');
    const cellXfs = child(styleRoot, 'cellXfs');
    for (const xf of children(cellXfs || { children: [] }, 'xf')) {
      const numFmtId = Number(xf.attributes.numFmtId) || 0;
      styleFormats.push({
        numFmtId,
        formatCode: customFormats.get(numFmtId) || BUILTIN_NUMBER_FORMATS[numFmtId] || null,
        properties: readCellStyleProperties(xf, fonts, fills, borders, colorContext)
      });
    }
    const dxfs = child(styleRoot, 'dxfs'), dxfEntries = children(dxfs || { children: [] }, 'dxf');
    differentialNodes.push(...dxfEntries);
    if (dxfs && (!/^(?:0|[1-9]\d*)$/.test(dxfs.attributes.count || '') || Number(dxfs.attributes.count) !== dxfEntries.length)) {
      markUnsupported('차등 서식 목록의 개수가 일치하지 않습니다.');
      differentialStyles.push(...dxfEntries.map(() => null));
    } else for (const dxf of dxfEntries) {
      try { differentialStyles.push(readDifferentialStyle(dxf, colorContext)); }
      catch { differentialStyles.push(null); }
    }
  }
  if (themePath && files[themePath]) known.add(themePath);
  for (const [sheetIndex, sheetRef] of sheetRefs.entries()) {
    const name = sheetRef.attributes.name, relId = sheetRef.attributes['r:id'];
    const printAreas = definedNames.find(item => item.localSheetId === sheetIndex && item.printAreas)?.printAreas || [];
    const areaDefinitions = printAreaDefinitions.filter(entry => entry.attributes.localSheetId === undefined || !Number.isInteger(Number(entry.attributes.localSheetId)) || Number(entry.attributes.localSheetId) < 0 || Number(entry.attributes.localSheetId) >= sheetRefs.length || Number(entry.attributes.localSheetId) === sheetIndex);
    const printAreaUnsupported = areaDefinitions.length > 0 && (areaDefinitions.length !== 1 || areaDefinitions[0].attributes.localSheetId === undefined || !printAreas.length);
    const titleDefinition = definedNames.find(item => item.localSheetId === sheetIndex && item.printTitles);
    const printTitles = titleDefinition?.printTitles || { rows: null, columns: null };
    const titleDefinitions = printTitleDefinitions.filter(entry => entry.attributes.localSheetId === undefined || !Number.isInteger(Number(entry.attributes.localSheetId)) || Number(entry.attributes.localSheetId) < 0 || Number(entry.attributes.localSheetId) >= sheetRefs.length || Number(entry.attributes.localSheetId) === sheetIndex);
    const printTitlesUnsupported = titleDefinitions.length > 0 && (titleDefinitions.length !== 1 || titleDefinitions[0].attributes.localSheetId === undefined || !titleDefinition);
    if (!name || sheetNames.has(name) || !relId) fail('워크시트 이름 또는 관계가 유효하지 않습니다.');
    sheetNames.add(name);
    const rel = workbookMap.get(relId);
    if (!rel?.path || !rel.Type.endsWith('/worksheet') || !files[rel.path]) fail(`시트 관계 대상이 유효하지 않습니다: ${name}`);
    known.add(rel.path);
    const parsed = parseXml(files[rel.path], rel.path); let root = parsed.root;
    const sheetView = children(child(root, 'sheetViews') || { children: [] }, 'sheetView')[0];
    const rawZoomScale = sheetView?.attributes.zoomScale;
    let zoomScale = 100;
    if (rawZoomScale !== undefined) {
      const parsedZoomScale = Number(rawZoomScale);
      if (Number.isInteger(parsedZoomScale) && parsedZoomScale >= 10 && parsedZoomScale <= 400) zoomScale = parsedZoomScale;
      else markUnsupported('워크시트 확대 비율이 유효하지 않습니다.');
    }
    const columnProperties = readColumnWidths(root), columnWidths = columnProperties.widths;
    let hiddenColumns = columnProperties.hiddenColumns;
    const headerFooterState = readHeaderFooter(root), pageBreakState = readPageBreaks(root);
    const paneResult = readFreezePane(root);
    if (paneResult.reason) markUnsupported(paneResult.reason);
    const autoFilterNode = child(root, 'autoFilter');
    let autoFilter = null;
    if (autoFilterNode) try { autoFilter = readAutoFilter(autoFilterNode, differentialNodes, colorContext); }
    catch (error) { markUnsupported(`AutoFilter: ${error.message}`); }
    const conditionalFormats = [];
    for (const conditionalFormatting of children(root, 'conditionalFormatting')) try { conditionalFormats.push(...readConditionalFormats(conditionalFormatting, differentialStyles, colorContext)); }
    catch (error) { markUnsupported(`조건부 서식: ${error.message}`); }
    if (conditionalFormats.length > 256) markUnsupported('조건부 서식 규칙 개수 제한을 초과했습니다.');
    conditionalFormats.sort((a, b) => a.priority - b.priority);
    if (new Set(conditionalFormats.map(rule => rule.priority)).size !== conditionalFormats.length) markUnsupported('조건부 서식 우선 순위가 중복되었습니다.');
    let dataValidations = [];
    try { dataValidations = readDataValidations(child(root, 'dataValidations'), definedNames, sheetIndex, date1904); }
    catch (error) { markUnsupported(`데이터 유효성: ${error.message}`); }
    let mergedCells = [];
    try { mergedCells = readMergedCells(child(root, 'mergeCells')); }
    catch (error) { markUnsupported(`병합 셀: ${error.message}`); }
    const chartData = readWorksheetCharts(root, rel.path, name, files, known, markUnsupported);
    const charts = chartData.charts;
    const hyperlinkData = readWorksheetHyperlinks(root, rel.path, files, known, markUnsupported);
    const unsupportedFeatures = ['drawing', 'legacyDrawing', 'pivotTableDefinition', 'extLst', 'protectedRanges', 'sheetProtection', 'controls', 'oleObjects', 'scenarios', 'dataConsolidate', 'customSheetViews', 'phoneticPr', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'legacyDrawingHF', 'picture', 'webPublishItems'];
    for (const feature of unsupportedFeatures) if (child(root, feature) && !(feature === 'drawing' && charts.length > 0)) markUnsupported(`워크시트 기능 ${feature}은 현재 저장을 지원하지 않습니다.`);
    const tablePartsNode = child(root, 'tableParts'), tableParts = children(tablePartsNode || { children: [] }, 'tablePart');
    let tables = [];
    if (tablePartsNode) {
      const sheetRelsPath = path.join(path.dirname(rel.path), '_rels', path.basename(rel.path) + '.rels');
      if (!files[sheetRelsPath] || tablePartsNode.children.some(node => node.name !== 'tablePart') || !/^(?:0|[1-9]\d*)$/.test(tablePartsNode.attributes.count || '') || Number(tablePartsNode.attributes.count) !== tableParts.length || !tableParts.length) markUnsupported('워크시트 표 관계 목록 구성이 유효하지 않습니다.');
      else {
        known.add(sheetRelsPath);
        const sheetRelationships = parseRelationships(parseXml(files[sheetRelsPath], sheetRelsPath).root, path.dirname(rel.path), ['table', 'hyperlink', 'drawing']);
        const usedRelationships = new Set(hyperlinkData.relationshipIds);
        const drawingId = child(root, 'drawing')?.attributes['r:id'];
        if (drawingId) usedRelationships.add(drawingId);
        for (const part of tableParts) {
          const id = part.attributes['r:id'], relationship = sheetRelationships.get(id);
          if (!id || usedRelationships.has(id) || !relationship?.path || !relationship.Type.endsWith('/table') || !files[relationship.path] || tables.some(table => table.path === relationship.path)) {
            markUnsupported('워크시트 표 관계가 유효하지 않습니다.');
            continue;
          }
          usedRelationships.add(id); known.add(relationship.path);
          const tableRoot = parseXml(files[relationship.path], relationship.path).root;
          const range = parseTableRange(tableRoot.attributes.ref || ''), columnsNode = child(tableRoot, 'tableColumns'), tableColumns = children(columnsNode || { children: [] }, 'tableColumn');
          const tableStyleNodes = children(tableRoot, 'tableStyleInfo');
          const tableName = tableRoot.attributes.displayName || tableRoot.attributes.name;
          const headerRowCount = tableRoot.attributes.headerRowCount === undefined ? 1 : Number(tableRoot.attributes.headerRowCount), totalsRowCount = tableRoot.attributes.totalsRowCount === undefined ? 0 : Number(tableRoot.attributes.totalsRowCount);
          if (tableRoot.name !== 'table' || !tableName || !range || !columnsNode || columnsNode.children.some(node => node.name !== 'tableColumn') || !/^(?:0|[1-9]\d*)$/.test(columnsNode.attributes.count || '') || Number(columnsNode.attributes.count) !== tableColumns.length || range.right - range.left + 1 !== tableColumns.length || tableColumns.some(column => !column.attributes.name) || new Set(tableColumns.map(column => column.attributes.name.toLocaleLowerCase())).size !== tableColumns.length || ![0, 1].includes(headerRowCount) || ![0, 1].includes(totalsRowCount) || headerRowCount + totalsRowCount > range.bottom - range.top + 1) markUnsupported(`표 구성이 안전하게 보존 가능한 기본 형태가 아닙니다: ${tableName || relationship.path}`);
          else {
            let styleInfo = null;
            if (tableStyleNodes.length > 1) markUnsupported(`표 스타일 정보가 중복되었습니다: ${tableName}`);
            else try { styleInfo = readTableStyleInfo(tableStyleNodes[0]); }
            catch (error) { markUnsupported(`표 스타일: ${error.message}`); }
            let tableAutoFilter = null;
            try { tableAutoFilter = readAutoFilter(child(tableRoot, 'autoFilter'), differentialNodes, colorContext, true); }
            catch (error) { markUnsupported(`표 자동 필터: ${error.message}`); }
            const expected = tableFilterRange({ range, headerRowCount, totalsRowCount });
            const legacyTotalsFilter = Boolean(tableAutoFilter && totalsRowCount && tableAutoFilter.bottom === range.bottom);
            if (tableAutoFilter && (!expected || expected.bottom < expected.top || tableAutoFilter.top !== expected.top || tableAutoFilter.bottom !== expected.bottom && !legacyTotalsFilter || tableAutoFilter.left !== expected.left || tableAutoFilter.right !== expected.right)) markUnsupported(`표 자동 필터 범위가 머리글·데이터 범위와 일치하지 않습니다: ${tableName}`);
            if (tableAutoFilter && expected) tableAutoFilter = { ...tableAutoFilter, bottom: expected.bottom };
            tables.push({ name: tableName, ref: tableRoot.attributes.ref, range, columns: tableColumns.map(column => column.attributes.name), headerRowCount, totalsRowCount, styleInfo, calculatedColumns: tableColumns.map(column => child(column, 'calculatedColumnFormula')?.text || null), calculatedColumnUnsupported: tableColumns.map(column => { const formula = child(column, 'calculatedColumnFormula'); return Boolean(formula && (formula.children.length || ['1', 'true'].includes(formula.attributes.array))); }), totalColumns: tableColumns.map(column => ({ configured: Object.hasOwn(column.attributes, 'totalsRowFunction') || Object.hasOwn(column.attributes, 'totalsRowLabel') || Boolean(child(column, 'totalsRowFormula')), function: column.attributes.totalsRowFunction || (child(column, 'totalsRowFormula') ? 'custom' : 'none'), label: column.attributes.totalsRowLabel ?? null, formula: child(column, 'totalsRowFormula')?.text ?? null })), autoFilter: tableAutoFilter, filterRangeNeedsNormalization: legacyTotalsFilter, path: relationship.path });
          }
        }
        if (usedRelationships.size !== sheetRelationships.size) markUnsupported('표 외의 워크시트 관계가 있어 안전한 편집을 보장할 수 없습니다.');
      }
    }
    let sheetData = child(root, 'sheetData'); if (!sheetData) fail(`시트 데이터가 없습니다: ${name}`);
    let outlineMetadata = { summaryBelow: true, summaryRight: true, rowOutlineLevels: {}, columnOutlineLevels: {}, rowGroups: [], columnGroups: [], outlineHiddenRows: [], outlineHiddenColumns: [] };
    try { outlineMetadata = readOutlineMetadata(root, sheetData); }
    catch (error) { markUnsupported(`행·열 그룹: ${error.message}`); }
    const outlineHiddenColumnSet = new Set(outlineMetadata.outlineHiddenColumns);
    hiddenColumns = hiddenColumns.filter(column => !outlineHiddenColumnSet.has(column) || columnProperties.manualHiddenColumns.has(column));
    let hiddenRows = [];
    try { hiddenRows = readHiddenRows(sheetData); }
    catch (error) { markUnsupported(`숨김 행: ${error.message}`); }
    let rowDimensions = { heights: {}, defaultHeight: 26 };
    try { rowDimensions = readRowDimensions(root, sheetData); }
    catch (error) { markUnsupported(`행 높이: ${error.message}`); }
    const sheetDataStructureValid = sheetData.children.every(node => node.name === 'row');
    const rowStructureValid = sheetDataStructureValid && sheetData.children.every(row => /^[1-9]\d*$/.test(row.attributes.r || '') && row.children.every(node => node.name === 'c'));
    const hasColumnDefinitions = Boolean(child(root, 'cols'));
    const cells = new Map();
    const sharedFormulaGroups = new Map();
    const ranges = findCellRanges(parsed.xml);
    const formulaRanges = [...findElementRanges(parsed.xml, 'f', undefined, true).values()].sort((left, right) => left.start - right.start);
    let formulaRangeIndex = 0;
    for (const row of children(sheetData, 'row')) {
      for (const cell of children(row, 'c')) {
        if (++totalCells > LIMITS.cells) fail('초기 버전은 전체 100,000셀을 넘는 통합문서를 열지 않습니다.');
        const address = cell.attributes.r;
        const r = rowNumber(address), c = columnNumber(address);
        if (!r || r > LIMITS.rows || !c || c > LIMITS.columns || cells.has(address)) fail(`셀 주소가 유효하지 않습니다: ${address}`);
        if (Number(row.attributes.r) !== r) fail(`셀 주소와 행 인덱스가 일치하지 않습니다: ${address}`);
        const range = ranges.get(address);
        while (range && formulaRangeIndex < formulaRanges.length && formulaRanges[formulaRangeIndex].start < range.start) formulaRangeIndex++;
        const formulaRange = range && formulaRangeIndex < formulaRanges.length && formulaRanges[formulaRangeIndex].start < range.end ? formulaRanges[formulaRangeIndex++] : null;
        const formula = child(cell, 'f');
        if (formula) anyFormula = true;
        if (formula && !formulaRange) markUnsupported(`셀 수식 XML 범위를 찾을 수 없습니다: ${address}`);
        const cellChildren = cell.children;
        if (cellChildren.some(node => !['f', 'v', 'is'].includes(node.name)) || ['f', 'v', 'is'].some(name => children(cell, name).length > 1)) markUnsupported('지원하지 않는 셀 구성요소가 포함되어 있습니다.');
        if (cellChildren.some(node => ['f', 'v'].includes(node.name) && node.children.length)) markUnsupported('복잡한 셀 값 구조는 현재 저장을 지원하지 않습니다.');
        const inlineString = child(cell, 'is');
        if (inlineString && (inlineString.children.length !== 1 || inlineString.children[0].name !== 't' || inlineString.children[0].children.length)) markUnsupported('서식 또는 확장이 포함된 인라인 문자열은 현재 저장을 지원하지 않습니다.');
        const decoded = decodeCellValue(cell, sharedStrings, date1904);
        const styleIndex = Number(cell.attributes.s) || 0, cellFormat = styleFormats[styleIndex] || { numFmtId: 0, formatCode: BUILTIN_NUMBER_FORMATS[0] };
        const item = { address, row: r - 1, column: c - 1, value: decoded.value, kind: decoded.kind, formula: formula?.text || null, formulaType: formula?.attributes?.t || null, hasFormula: Boolean(formula), cachedValue: decoded.value, style: cell.attributes.s || null, styleProperties: cellFormat.properties || null, numFmtId: cellFormat.numFmtId, formatCode: cellFormat.formatCode, ...(date1904 ? { date1904: true } : {}) };
        if (decoded.isoDate !== undefined) item.isoDate = decoded.isoDate;
        cells.set(address, item);
        if (formula?.attributes?.t === 'shared') {
          const sharedIndex = formula.attributes.si;
          if (!/^\d+$/.test(sharedIndex || '')) { markUnsupported('공유 수식 그룹 번호가 유효하지 않습니다.'); continue; }
          if (!sharedFormulaGroups.has(sharedIndex)) sharedFormulaGroups.set(sharedIndex, []);
          sharedFormulaGroups.get(sharedIndex).push({ cell: item, formula: formula.text.trim(), attributes: formula.attributes, formulaRange });
        }
      }
    }
    const sharedFormulaGroupRecords = [];
    for (const [sharedIndex, members] of sharedFormulaGroups) {
      const masters = members.filter(member => member.formula);
      if (masters.length !== 1) { markUnsupported(`공유 수식 그룹 ${sharedIndex}에 기준 수식이 하나만 있어야 합니다.`); continue; }
      const master = masters[0], bounds = parseRangeReference(master.attributes.ref, null, LIMITS.cells);
      if (!bounds || bounds.sheetName || members.some(member => member.attributes.ref && member !== master)
        || members.some(({ cell }) => cell.row < bounds.top || cell.row > bounds.bottom || cell.column < bounds.left || cell.column > bounds.right)) {
        markUnsupported(`공유 수식 그룹 ${sharedIndex}의 기준 범위가 유효하지 않습니다.`); continue;
      }
      try {
        for (const { cell } of members) cell.formula = translateFormulaReferences(master.formula, cell.row - master.cell.row, cell.column - master.cell.column);
        sharedFormulaGroupRecords.push({ sharedIndex, members: members.map(({ cell, formulaRange, attributes }) => ({ address: cell.address, row: cell.row, column: cell.column, formula: cell.formula, formulaRange, attributes: { ...attributes } })) });
      } catch {
        markUnsupported(`공유 수식 그룹 ${sharedIndex}의 상대 참조를 펼칠 수 없습니다.`);
      }
    }
    if (ranges.size !== cells.size) fail(`시트 XML 셀 인덱스 검증에 실패했습니다: ${name} (${[...ranges.keys()].join(',')}/${[...cells.keys()].join(',')})`);
    const outlineHiddenColumns = new Set(outlineMetadata.outlineHiddenColumns || []);
    const manuallyHiddenColumns = columnProperties.nativeHiddenColumns.filter(column => {
      const explicit = columnProperties.manualVisibility?.[column];
      return explicit === undefined ? !outlineHiddenColumns.has(column) : explicit;
    });
    sheetEntries.push({ name, path: rel.path, tableRelationshipPath: tables.length ? path.join(path.dirname(rel.path), '_rels', path.basename(rel.path) + '.rels') : null, cells, ranges, sharedFormulaGroups: sharedFormulaGroupRecords, printAreas, printAreaUnsupported, printTitles, printTitlesUnsupported, pageBreaks: pageBreakState.value, pageBreaksUnsupported: pageBreakState.unsupported, headerFooter: headerFooterState.settings, headerFooterUnsupported: headerFooterState.unsupported, printSettings: readPrintSettings(root), rowHeights: rowDimensions.heights, customRowHeights: rowDimensions.customRows || [], defaultRowHeight: rowDimensions.defaultHeight, defaultColumnWidth: columnProperties.defaultWidth, defaultColumnWidthCharacters: columnProperties.defaultWidthCharacters, columnWidths, hiddenColumns: manuallyHiddenColumns, nativeHiddenColumns: columnProperties.nativeHiddenColumns, columnVisibilityPrefix: Object.keys(root.attributes).find(key => key.startsWith('xmlns:') && root.attributes[key] === COLUMN_VISIBILITY_NS)?.slice(6) || null, rowOutlineLevels: outlineMetadata.rowOutlineLevels, columnOutlineLevels: outlineMetadata.columnOutlineLevels, rowGroups: outlineMetadata.rowGroups, columnGroups: outlineMetadata.columnGroups, outlineHiddenRows: outlineMetadata.outlineHiddenRows, outlineHiddenColumns: outlineMetadata.outlineHiddenColumns, summaryBelow: outlineMetadata.summaryBelow, summaryRight: outlineMetadata.summaryRight, freezePane: paneResult.freezePane, zoomScale, autoFilter, hiddenRows, nativeHiddenRows: [...hiddenRows], rowVisibilityPrefix: Object.keys(root.attributes).find(key => key.startsWith('xmlns:') && root.attributes[key] === ROW_VISIBILITY_NS)?.slice(6) || null, filterMode: ['1', 'true', 'on'].includes(child(root, 'sheetPr')?.attributes.filterMode), conditionalFormats, dataValidations, mergedCells, hyperlinks: hyperlinkData.references, tables, charts, chartPlacements: chartData.chartPlacements || [], chartDrawingPath: chartData.drawingPath || null, chartAnchors: chartData.chartAnchors || [], chartPrintUnsupported: Boolean(chartData.printUnsupported) || chartData.chartPaths.length !== charts.length || childrenByLocalName(root, 'drawing').length > 0 && !chartData.chartPaths.length, chartSourceRanges: chartData.sourceRanges, chartPaths: chartData.chartPaths, hasTables: tables.length > 0, hasColumnDefinitions, sheetDataStructureValid, rowStructureValid, xml: parsed.xml, visible: sheetRef.attributes.state !== 'hidden' && sheetRef.attributes.state !== 'veryHidden' });
    root = null; parsed.root = null; sheetData = null;
  }
  for (const sheet of sheetEntries) for (const validation of sheet.dataValidations || []) if (validation.source?.sheetName && !sheetEntries.some(item => item.name.toLocaleLowerCase() === validation.source.sheetName.toLocaleLowerCase())) markUnsupported(`데이터 유효성 목록 원본 시트를 찾을 수 없습니다: ${validation.source.sheetName}`);
  for (const name of records.keys()) {
    if (name.endsWith('/')) continue;
    if (!known.has(name) && name !== '[Content_Types].xml' && name !== '_rels/.rels' && name !== 'xl/workbook.xml' && name !== workbookRelsPath) markUnsupported(`지원하지 않는 OOXML 구성요소가 있습니다: ${name}`);
    if (/\.(xml|rels)$/i.test(name) && (known.has(name) || name === '[Content_Types].xml' || name === '_rels/.rels' || name === workbookRelsPath)) parseXml(files[name], name);
  }
  if (/\.xlsm$/i.test(extension)) markUnsupported('매크로 통합문서는 JinSheet에서 수정할 수 없습니다.');
  const contentTypes = files['[Content_Types].xml'];
  if (!contentTypes) fail('OOXML 콘텐츠 형식 목록이 없습니다.');
  const types = parseXml(contentTypes, '[Content_Types].xml').root;
  for (const entry of children(types, 'Override')) {
    const type = entry.attributes.ContentType || '';
    if (/macroEnabled|vbaProject|vbaData|activeX|oleObject|pivot|comments|connections|queryTable|slicer|signature|customXml|externalLink/i.test(type)) markUnsupported(`지원하지 않는 콘텐츠 형식이 있습니다: ${type}`);
  }
  const chartSourceRanges = sheetEntries.flatMap((sheet, sheetIndex) => sheet.chartSourceRanges.map(range => ({ ...range, sheetIndex: range.sheetName ? sheetEntries.findIndex(entry => entry.name.toLocaleLowerCase() === range.sheetName.toLocaleLowerCase()) : sheetIndex })));
  if (chartSourceRanges.some(range => range.sheetIndex < 0)) markUnsupported('차트 원본 수식이 없는 워크시트를 참조합니다.');
  const chartPaths = [...new Map(sheetEntries.flatMap((sheet, sheetIndex) => sheet.chartPaths.map(chartPath => [chartPath, { path: chartPath, sheetIndex }]))).values()];
  classifyFilterHiddenRows(sheetEntries, definedNames, date1904);
  return { bytes, files, records, sheets: sheetEntries, normalFont, chartSourceRanges, chartPaths, hasCharts: sheetEntries.some(sheet => sheet.charts.length > 0), sharedStrings, styleFormats, colorContext, definedNames, date1904, stylesPath: stylesPath || null, writable: unsupported.length === 0, readOnlyReason: unsupported.join(' '), unsupported, formulas: anyFormula, extension };
}

function filterRowsForModel(model, options = {}) {
  const calculator = createFormulaCalculator({ now: options.now, state: { ...model, formulas: true }, getCell: (row, column, sheetIndex = 0) => model.sheets[sheetIndex]?.cells.get(columnLetters(column + 1) + (row + 1)) || null });
  return model.sheets.map((sheet, sheetIndex) => {
    const hidden = new Set();
    for (const filter of [sheet.autoFilter, ...(sheet.tables || []).map(table => table.autoFilter)].filter(item => item?.columns.length)) {
      const rules = filter.columns.map(rule => ({ column: rule.column, matches: compileAutoFilterColumn(rule, () => filterColumnCells(filter, rule.column, (row, column) => calculator.cell(sheetIndex, row, column)), { date1904: model.date1904, now: calculator.now, getStyle: cell => calculator.filterStyle(sheetIndex, cell.row, cell.column) }) }));
      for (let row = filter.top + 1; row <= filter.bottom; row++) if (rules.some(rule => !rule.matches({ ...calculator.cell(sheetIndex, row, rule.column), row, column: rule.column }))) hidden.add(row);
    }
    return hidden;
  });
}

function classifyFilterHiddenRows(sheets, definedNames, date1904) {
  // Native OOXML has one hidden bit for both manual hiding and filtering.
  // Our ignorable metadata preserves overlap; external files require inference.
  const external = [];
  for (const [index, sheet] of sheets.entries()) {
    const outlineHidden = new Set(sheet.outlineHiddenRows || []);
    if (!sheet.filterMode && !sheet.rowVisibilityPrefix && !outlineHidden.size) continue;
    const root = parseXml(new TextEncoder().encode(sheet.xml), sheet.path).root;
    const prefix = sheet.rowVisibilityPrefix;
    if (prefix && root.attributes[prefix + ':version'] === '1') {
      const rows = children(child(root, 'sheetData'), 'row'), native = new Set(sheet.nativeHiddenRows);
      sheet.hiddenRows = rows.filter(row => row.attributes[prefix + ':manualHidden'] === '1' && native.has(Number(row.attributes.r) - 1)).map(row => Number(row.attributes.r) - 1);
    } else {
      sheet.hiddenRows = sheet.hiddenRows.filter(row => !outlineHidden.has(row));
      if (sheet.filterMode) external.push(index);
    }
  }
  if (!external.length) return;
  const inferred = filterRowsForModel({ sheets, definedNames, date1904 });
  for (const index of external) sheets[index].hiddenRows = sheets[index].hiddenRows.filter(row => !inferred[index].has(row) && !(sheets[index].outlineHiddenRows || []).includes(row));
}

function preserveImportedFilterVisibility(model) {
  return model.sheets.some(sheet => sheet.filterMode && !sheet.rowVisibilityPrefix) ? patchFilterVisibility(model).model : model;
}

function patchFilterVisibility(model, options = {}) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  const now = options.now instanceof Date ? new Date(options.now.getTime()) : new Date();
  const filtered = filterRowsForModel(model, { now }), files = { ...model.files }, modifiedPaths = new Set();
  for (const sheet of model.sheets) for (const table of sheet.tables || []) {
    if (!table.filterRangeNeedsNormalization || !table.autoFilter) continue;
    const xml = new TextDecoder().decode(files[table.path]), element = findElementRanges(xml, 'autoFilter').get(''), range = tableFilterRange(table);
    if (!element || !range) fail('표 자동 필터 범위를 정규화할 수 없습니다.');
    const opening = xml.slice(element.start, element.startTagEnd), reference = columnLetters(range.left + 1) + (range.top + 1) + ':' + columnLetters(range.right + 1) + (range.bottom + 1);
    files[table.path] = new TextEncoder().encode(xml.slice(0, element.start) + setTagAttribute(opening, 'ref', reference) + xml.slice(element.startTagEnd));
    modifiedPaths.add(table.path);
  }
  for (const [sheetIndex, sheet] of model.sheets.entries()) {
    const active = [sheet.autoFilter, ...(sheet.tables || []).map(table => table.autoFilter)].some(filter => filter?.columns.length);
    const outlineHidden = new Set(sheet.outlineHiddenRows || []);
    const ensureMetadata = options.ensureVisibilityMetadata === true || Array.isArray(options.ensureVisibilityMetadata) && options.ensureVisibilityMetadata.includes(sheetIndex);
    if (!active && !sheet.filterMode && !sheet.rowVisibilityPrefix && !outlineHidden.size && !ensureMetadata) continue;
    const manual = new Set(sheet.hiddenRows), wanted = new Set([...manual, ...filtered[sheetIndex], ...outlineHidden]);
    const old = new Set(sheet.nativeHiddenRows || []);
    const show = [...old].filter(row => !wanted.has(row)).sort((a,b) => a-b);
    const hide = [...wanted].filter(row => !old.has(row)).sort((a,b) => a-b);
    let current = model;
    if (show.length) current = patchHiddenRows(current, [{ sheetIndex, rows: show, hidden: false }], false).model;
    if (hide.length) current = patchHiddenRows(current, [{ sheetIndex, rows: hide, hidden: true }], false).model;
    let xml = current.sheets[sheetIndex].xml;
    const opening = /<worksheet\b[^>]*>/.exec(xml);
    if (!opening) fail('행 표시 메타데이터를 추가할 수 없습니다.');
    let tag = opening[0], attrs = parseTagAttributes(tag);
    const namespacePrefix = (uri, preferred) => {
      const existing = Object.keys(attrs).find(key => key.startsWith('xmlns:') && attrs[key] === uri);
      if (existing) return existing.slice(6);
      let name = preferred, n = 1; while (attrs['xmlns:' + name] !== undefined) name = preferred + n++;
      tag = setTagAttribute(tag, 'xmlns:' + name, uri); attrs['xmlns:' + name] = uri; return name;
    };
    const prefix = namespacePrefix(ROW_VISIBILITY_NS, 'jsv'), mc = namespacePrefix(MC_NS, 'mc');
    const merge = (name, values) => { tag = setTagAttribute(tag, name, [...new Set([...(attrs[name] || '').split(/\s+/).filter(Boolean), ...values])].join(' ')); };
    merge(mc + ':Ignorable', [prefix]);
    merge(mc + ':PreserveAttributes', [prefix + ':version', prefix + ':manualHidden']);
    tag = setTagAttribute(tag, prefix + ':version', '1');
    xml = xml.slice(0, opening.index) + tag + xml.slice(opening.index + opening[0].length);
    xml = xml.replace(/<row\b[^>]*>/g, row => {
      const attributes = parseTagAttributes(row), index = Number(attributes.r) - 1;
      return manual.has(index) || attributes[prefix + ':manualHidden'] !== undefined ? setTagAttribute(row, prefix + ':manualHidden', manual.has(index) ? '1' : '0') : row;
    });
    const properties = findElementRanges(xml, 'sheetPr').get('');
    if (properties) {
      const before = xml.slice(properties.start, properties.startTagEnd);
      xml = xml.slice(0, properties.start) + setTagAttribute(before, 'filterMode', active ? '1' : '0') + xml.slice(properties.startTagEnd);
    } else {
      const start = /<worksheet\b[^>]*>/.exec(xml); const position = start.index + start[0].length;
      xml = xml.slice(0, position) + '<sheetPr filterMode="' + (active ? '1' : '0') + '"/>' + xml.slice(position);
    }
    if (xml !== sheet.xml) { files[sheet.path] = new TextEncoder().encode(xml); modifiedPaths.add(sheet.path); }
  }
  const calculator = createFormulaCalculator({ now, state: { ...model, formulas: true }, getCell: (row, column, sheetIndex = 0) => model.sheets[sheetIndex]?.cells.get(columnLetters(column + 1) + (row + 1)) || null });
  for (const [sheetIndex, sheet] of model.sheets.entries()) for (const owner of [{ path: sheet.path, autoFilter: sheet.autoFilter }, ...(sheet.tables || [])]) {
    const filter = owner.autoFilter;
    if (!filter?.columns.some(column => column.top10 || column.dynamic)) continue;
    let xml = new TextDecoder().decode(files[owner.path]), original = xml;
    for (const range of [...findElementRanges(xml, 'filterColumn', undefined, true).values()].sort((a, b) => b.start - a.start)) {
      const content = xml.slice(range.start, range.end), offset = Number(parseTagAttributes(content.slice(0, content.indexOf('>') + 1)).colId), column = filter.columns.find(item => item.column === filter.left + offset);
      if (column?.dynamic) {
        const criteria = dynamicFilterCriteria(column.dynamic, () => filterColumnCells(filter, column.column, (row, col) => calculator.cell(sheetIndex, row, col)), { now, date1904: model.date1904 });
        const element = findElementRanges(content, 'dynamicFilter').get('');
        if (!element) fail('동적 필터 기준값을 저장할 요소가 없습니다.');
        const tag = '<dynamicFilter type="' + column.dynamic.type + '"' + (Number.isFinite(criteria.value) ? ' val="' + criteria.value + '"' : '') + (Number.isFinite(criteria.maxValue) ? ' maxVal="' + criteria.maxValue + '"' : '') + '/>';
        const updated = content.slice(0, element.start) + tag + content.slice(element.end);
        xml = xml.slice(0, range.start) + updated + xml.slice(range.end); continue;
      }
      if (!column?.top10) continue;
      const threshold = topFilterThreshold(column.top10, filterColumnCells(filter, column.column, (row, col) => calculator.cell(sheetIndex, row, col)));
      const top = findElementRanges(content, 'top10').get('');
      if (!top) fail('순위 필터 기준값을 저장할 요소가 없습니다.');
      const opening = content.slice(top.start, top.startTagEnd), next = threshold === null ? opening.replace(/\s+filterVal\s*=\s*(?:"[^"]*"|'[^']*')/, '') : setTagAttribute(opening, 'filterVal', threshold);
      const updated = content.slice(0, top.start) + next + content.slice(top.startTagEnd);
      xml = xml.slice(0, range.start) + updated + xml.slice(range.end);
    }
    if (xml !== original) { files[owner.path] = new TextEncoder().encode(xml); modifiedPaths.add(owner.path); }
  }
  if (!modifiedPaths.size) return { bytes: model.bytes, model };
  updateDimensionDrawings(model, files, modifiedPaths);
  if (model.formulas) {
    files['xl/workbook.xml'] = new TextEncoder().encode(enableFullCalculationOnLoad(new TextDecoder().decode(files['xl/workbook.xml'])));
    modifiedPaths.add('xl/workbook.xml');
  }
  const bytes = zipSync(files, { level: 0 }), checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('필터 행 표시 저장 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  return { bytes, model: checked };
}

function patchCell(model, sheetIndex, address, value) {
  return patchCells(model, [{ sheetIndex, address, value }]);
}

function findElementRanges(xml, targetName, keyAttribute, allowDuplicates = false) {
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
      if (name === targetName) { const key = allowDuplicates ? String(ranges.size) : entry.key; if (ranges.has(key)) fail('XML 요소 키가 중복되었습니다: ' + key); ranges.set(key, { start: entry.start, startTagEnd: entry.startTagEnd, closingStart: lt, end: end + 1, selfClosing: false }); }
    } else {
      const attrs = parseTagAttributes(raw);
      const key = name === targetName ? String(keyAttribute ? (attrs[keyAttribute] ?? '') : '') : null;
      if (name === targetName && selfClosing) { const rangeKey = allowDuplicates ? String(ranges.size) : key; if (ranges.has(rangeKey)) fail('XML 요소 키가 중복되었습니다: ' + rangeKey); ranges.set(rangeKey, { start: lt, startTagEnd: end + 1, closingStart: end + 1, end: end + 1, selfClosing: true }); }
      else if (!selfClosing) stack.push({ name, start: lt, startTagEnd: end + 1, key });
    }
    i = end + 1;
  }
  if (stack.length) fail('닫히지 않은 XML 요소가 있습니다.');
  return ranges;
}

function chartSourceContains(source, cell) {
  return source.sheetIndex === cell.sheetIndex && cell.row >= source.top && cell.row <= source.bottom && cell.column >= source.left && cell.column <= source.right;
}

function chartHasFormulaSources(model) {
  return model.sheets.some((sheet, sheetIndex) => [...sheet.cells.values()].some(cell => cell.formula && model.chartSourceRanges?.some(source => chartSourceContains(source, { sheetIndex, row: cell.row, column: cell.column }))));
}

function chartRangeValues(model, sourceSheetIndex, source, numeric, calculator, getCell) {
  if ((source.bottom - source.top + 1) * (source.right - source.left + 1) > 10000) fail('차트 원본 범위가 캐시 제한을 초과했습니다.');
  const values = [];
  for (let row = source.top; row <= source.bottom; row++) for (let column = source.left; column <= source.right; column++) {
    const cell = getCell ? getCell(row, column, sourceSheetIndex) : model.sheets[sourceSheetIndex].cells.get(columnLetters(column + 1) + (row + 1));
    const cellValue = cell?.formula ? calculator.value(sourceSheetIndex, row, column) : cell?.value;
    if (Array.isArray(cellValue)) fail('동적 배열 수식을 차트 원본 캐시에 넣을 수 없습니다.');
    if (!cell || cell.kind === 'blank' && !cell.formula || cellValue === null || cellValue === undefined || typeof cellValue === 'string' && cellValue.startsWith('#')) { values.push(null); continue; }
    if (numeric) {
      const value = !cell.formula && typeof cellValue === 'string' && !cellValue.trim() ? null : cell.kind === 'number' || cell.formula ? Number(cellValue) : null;
      values.push(Number.isFinite(value) ? value : null);
    } else values.push(typeof cellValue === 'boolean' ? cellValue ? 'TRUE' : 'FALSE' : String(cellValue));
  }
  return values;
}

function calculateChartPreviews(model, sheetIndex, calculator, getCell, changes) {
  const sheet = model.sheets[sheetIndex];
  if (!sheet?.charts?.length || !changes.length) return sheet?.charts || [];
  const refreshAll = chartHasFormulaSources(model);
  if (!refreshAll && !model.chartSourceRanges?.some(source => changes.some(cell => chartSourceContains(source, cell)))) return sheet.charts;
  const resolveReference = reference => {
    const formula = firstDescendant(reference, ['f'])?.text.trim();
    if (!formula || !['strRef', 'numRef'].includes(localName(reference))) return undefined;
    const source = parseRangeReference(formula, sheet.name);
    const sourceSheetIndex = source?.sheetName ? model.sheets.findIndex(item => item.name.toLocaleLowerCase() === source.sheetName.toLocaleLowerCase()) : sheetIndex;
    if (!source || !model.sheets[sourceSheetIndex]) fail('차트 원본 워크시트를 찾을 수 없습니다.');
    if (!refreshAll && !changes.some(cell => chartSourceContains({ ...source, sheetIndex: sourceSheetIndex }, cell))) return undefined;
    return chartRangeValues(model, sourceSheetIndex, source, localName(reference) === 'numRef', calculator, getCell);
  };
  return sheet.chartPaths.map((partName, index) => {
    const chart = readChartPreview(model.files[partName], partName, sheet.name, resolveReference);
    if (!chart) return sheet.charts[index];
    return { type: chart.type, direction: chart.direction, title: chart.title, series: chart.series };
  });
}

function patchChartCaches(chartBytes, chartPath, chartSheetIndex, model, changedSourceCells) {
  const { xml, root } = parseXml(chartBytes, chartPath);
  const prefix = root.name.includes(':') ? root.name.slice(0, root.name.indexOf(':') + 1) : '';
  const replacements = [];
  const formulaCalculator = createFormulaCalculator({
    state: { formulas: true, sheets: model.sheets, definedNames: model.definedNames, date1904: model.date1904 },
    getCell: (row, column, sheetIndex = 0) => model.sheets[sheetIndex]?.cells.get(columnLetters(column + 1) + (row + 1)) || null
  });
  for (const referenceName of ['numRef', 'strRef']) {
    const referenceRanges = findElementRanges(xml, prefix + referenceName, undefined, true);
    for (const referenceRange of referenceRanges.values()) {
      const referenceXml = xml.slice(referenceRange.start, referenceRange.end);
      const referenceRoot = parseXml(new TextEncoder().encode(referenceXml), chartPath).root;
      const formula = firstDescendant(referenceRoot, ['f'])?.text.trim();
      if (!formula) continue;
      const source = parseRangeReference(formula, model.sheets[chartSheetIndex]?.name);
      if (!source || source.top !== source.bottom && source.left !== source.right) fail('차트 데이터 범위를 안전하게 갱신할 수 없습니다.');
      const sourceSheetIndex = source.sheetName ? model.sheets.findIndex(sheet => sheet.name.toLocaleLowerCase() === source.sheetName.toLocaleLowerCase()) : chartSheetIndex;
      const sourceSheet = model.sheets[sourceSheetIndex];
      if (!sourceSheet) fail('차트 원본 워크시트를 찾을 수 없습니다.');
      if (changedSourceCells && !changedSourceCells.some(cell => chartSourceContains({ ...source, sheetIndex: sourceSheetIndex }, cell))) continue;
      const cacheName = referenceName === 'numRef' ? 'numCache' : 'strCache';
      const cacheRange = findElementRanges(referenceXml, prefix + cacheName).values().next().value;
      if (!cacheRange) fail('차트 캐시 구조를 안전하게 갱신할 수 없습니다.');
      const cacheXml = referenceXml.slice(cacheRange.start, cacheRange.end);
      const cacheRoot = parseXml(new TextEncoder().encode(cacheXml), chartPath).root;
      if (cacheRoot.children.some(item => !['formatCode', 'ptCount', 'pt'].includes(localName(item)))) fail('확장 데이터가 포함된 차트 캐시는 수정할 수 없습니다.');
      const formatRange = findElementRanges(cacheXml, prefix + 'formatCode').values().next().value;
      const formatXml = formatRange ? cacheXml.slice(formatRange.start, formatRange.end) : '';
      const values = chartRangeValues(model, sourceSheetIndex, source, referenceName === 'numRef', formulaCalculator);
      const points = values.map((value, index) => value === null ? '' : '<' + prefix + 'pt idx="' + index + '"><' + prefix + 'v>' + xmlEscape(value) + '</' + prefix + 'v></' + prefix + 'pt>').join('');
      let opening = cacheXml.slice(0, cacheRange.startTagEnd - cacheRange.start);
      if (cacheRange.selfClosing) opening = opening.replace(/\/\s*>$/, '>');
      const closing = '</' + prefix + cacheName + '>';
      const updatedCache = opening + formatXml + '<' + prefix + 'ptCount val="' + values.length + '"/>' + points + closing;
      replacements.push({ start: referenceRange.start + cacheRange.start, end: referenceRange.start + cacheRange.end, value: updatedCache });
    }
  }
  replacements.sort((left, right) => right.start - left.start);
  let updated = xml;
  for (const replacement of replacements) updated = updated.slice(0, replacement.start) + replacement.value + updated.slice(replacement.end);
  return new TextEncoder().encode(updated);
}

function serializeCellContent(value, kind, keep, formula) {
  if (formula !== undefined) {
    if (typeof formula !== 'string' || !formula.length || formula.length > 100000 || formula.startsWith('=')) fail('셀 수식이 유효하지 않습니다.');
    return '<c' + keep + '><f>' + xmlEscape(formula) + '</f><v></v></c>';
  }
  const text = String(value);
  if (kind === 'blank' && text === '') return '<c' + keep + '/>';
  const numeric = /^\s*[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?\s*$/.test(text) && Number.isFinite(Number(text.trim()));
  if (kind === 'error' && /^#[A-Z0-9/?!]+$/i.test(text)) return '<c' + keep + ' t="e"><v>' + xmlEscape(text) + '</v></c>';
  if ((kind === 'number' || kind === 'blank') && numeric) return '<c' + keep + '><v>' + text.trim() + '</v></c>';
  if ((kind === 'boolean' || kind === 'blank') && /^(?:TRUE|FALSE)$/i.test(text)) return '<c' + keep + ' t="b"><v>' + (text.toUpperCase() === 'TRUE' ? '1' : '0') + '</v></c>';
  const space = /^\s|\s$/.test(text) ? ' xml:space="preserve"' : '';
  return '<c' + keep + ' t="inlineStr"><is><t' + space + '>' + xmlEscape(text) + '</t></is></c>';
}

function serializeUnsharedFormula(xml, range, formula) {
  if (!range || typeof formula !== 'string' || !formula) fail('공유 수식 그룹의 구성식을 보존할 수 없습니다.');
  const opening = xml.slice(range.start, range.startTagEnd), name = /^<\s*([A-Za-z_][\w.:-]*)/.exec(opening)?.[1];
  if (!name) fail('공유 수식 요소 이름이 유효하지 않습니다.');
  const attributes = parseTagAttributes(opening);
  delete attributes.t; delete attributes.si; delete attributes.ref;
  const serialized = Object.entries(attributes).map(([key, value]) => ' ' + key + '="' + xmlEscape(value) + '"').join('');
  return '<' + name + serialized + '>' + xmlEscape(formula) + '</' + name + '>';
}

function flattenSharedFormulaGroups(model) {
  const outputFiles = Object.assign({}, model.files);
  for (const sheet of model.sheets) {
    const patches = (sheet.sharedFormulaGroups || []).flatMap(group => group.members.map(member => ({
      ...member.formulaRange,
      replacement: serializeUnsharedFormula(sheet.xml, member.formulaRange, member.formula)
    })));
    if (!patches.length) continue;
    patches.sort((left, right) => right.start - left.start);
    let xml = sheet.xml, previousStart = xml.length + 1;
    for (const patch of patches) {
      if (patch.end > previousStart) fail('공유 수식 요소 범위가 겹쳐 구조 변경을 중단했습니다.');
      xml = xml.slice(0, patch.start) + patch.replacement + xml.slice(patch.end);
      previousStart = patch.start;
    }
    outputFiles[sheet.path] = new TextEncoder().encode(xml);
  }
  const checked = analyze(zipSync(outputFiles, { level: 0 }), model.extension);
  if (!checked.writable) fail('공유 수식을 독립 수식으로 펼친 뒤 검사에 실패했습니다: ' + checked.readOnlyReason);
  if (checked.sheets.some(sheet => (sheet.sharedFormulaGroups || []).length)) fail('공유 수식 그룹을 모두 펼치지 못해 구조 변경을 중단했습니다.');
  return checked;
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
  if (!sheetData) fail('워크시트 데이터 영역을 안전하게 찾지 못했습니다.');
  const grouped = new Map();
  const styleAttribute = cell => {
    if (cell.style === undefined || cell.style === null) return '';
    if (typeof cell.style !== 'string' || !/^(?:0|[1-9]\d*)$/.test(cell.style)) fail('셀 서식 인덱스가 유효하지 않습니다.');
    return ' s="' + cell.style + '"';
  };
  for (const cell of additions) {
    const row = rowNumber(cell.address);
    if (!grouped.has(row)) grouped.set(row, []);
    grouped.get(row).push(cell);
  }
  if (sheetData.selfClosing) {
    const opening = sheet.xml.slice(sheetData.start, sheetData.startTagEnd).replace(/\/\s*>$/, '>');
    const content = [...grouped].sort((a, b) => a[0] - b[0]).map(([row, cells]) => '<row r="' + row + '">' + cells.sort((a, b) => columnNumber(a.address) - columnNumber(b.address)).map(cell => serializeCellContent(cell.value, cell.kind || 'blank', ' r="' + cell.address + '"' + styleAttribute(cell), cell.formula)).join('') + '</row>').join('');
    operations.push({ start: sheetData.start, end: sheetData.end, replacement: opening + content + '</sheetData>' });
    return;
  }
  for (const [rowNumberValue, cells] of grouped) {
    cells.sort((a, b) => columnNumber(a.address) - columnNumber(b.address));
    const row = rows.get(String(rowNumberValue));
    if (row && row.selfClosing) {
      const opening = sheet.xml.slice(row.start, row.startTagEnd).replace(/\s+spans\s*=\s*(["'])[^"']*\1/, '').replace(/\/\s*>$/, '>');
      const content = cells.map(function (cell) { return serializeCellContent(cell.value, cell.kind || 'blank', ' r="' + cell.address + '"' + styleAttribute(cell), cell.formula); }).join('');
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
        operations.push({ start: position, end: position, replacement: serializeCellContent(cell.value, cell.kind || 'blank', ' r="' + cell.address + '"' + styleAttribute(cell), cell.formula), column });
      }
      continue;
    }
    const laterRow = [...rows.entries()].filter(function (entry) { return Number(entry[0]) > rowNumberValue; }).sort(function (a, b) { return Number(a[0]) - Number(b[0]); })[0];
    const position = laterRow ? laterRow[1].start : sheetData.closingStart;
    const cellXml = cells.map(function (cell) { return serializeCellContent(cell.value, cell.kind || 'blank', ' r="' + cell.address + '"' + styleAttribute(cell), cell.formula); }).join('');
    operations.push({ start: position, end: position, replacement: '<row r="' + rowNumberValue + '">' + cellXml + '</row>', column: rowNumberValue });
  }
}

function prepareNumberStyles(model, formatByBase) {
  if (!formatByBase.size) return null;
  const stylesPath = model.stylesPath;
  if (!stylesPath || !model.files[stylesPath]) fail('통합문서에 숫자 표시 형식을 추가할 스타일 정의가 없습니다.');
  let xml = new TextDecoder().decode(model.files[stylesPath]);
  const cellXfs = /<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/.exec(xml);
  if (!cellXfs) fail('셀 서식 목록에 숫자 표시 형식을 추가할 수 없습니다.');
  const start = cellXfs.index, openEnd = xml.indexOf('>', start) + 1, closeStart = start + cellXfs[0].lastIndexOf('</cellXfs>');
  const opening = xml.slice(start, openEnd), body = xml.slice(openEnd, closeStart), attributes = parseTagAttributes(opening);
  const formats = [];
  for (const match of body.matchAll(/<xf(?=\s|>)/g)) {
    const tagStart = match.index, tagEnd = body.indexOf('>', tagStart);
    if (tagEnd < 0) fail('셀 서식 태그가 닫히지 않았습니다.');
    const openingTag = body.slice(tagStart, tagEnd + 1);
    if (/\/\s*>$/.test(openingTag)) formats.push(openingTag);
    else {
      const closing = body.indexOf('</xf>', tagEnd + 1);
      if (closing < 0) fail('셀 서식 항목이 닫히지 않았습니다.');
      formats.push(body.slice(tagStart, closing + 5));
    }
  }
  const count = Number(attributes.count);
  if (!Number.isInteger(count) || count !== formats.length || count !== model.styleFormats.length) fail('셀 서식 인덱스와 XML 목록이 일치하지 않습니다.');
  const normalized = value => value
    .replace(/\s+numFmtId\s*=\s*(["'])\d+\1/, ' numFmtId="0"')
    .replace(/\s+applyNumberFormat\s*=\s*(["'])[^"']*\1/, '')
    .replace(/\s*\/>$/, '>')
    .replace(/\s*>$/, '>');
  const byBase = new Map(), additions = [], customFormats = new Map();
  const numFmtsMatch = /<numFmts\b[^>]*>[\s\S]*?<\/numFmts>/.exec(xml);
  if (numFmtsMatch) for (const match of numFmtsMatch[0].matchAll(/<numFmt\b[^>]*\/?\s*>/g)) {
    const attributes = parseTagAttributes(match[0]);
    if (/^\d+$/.test(attributes.numFmtId || '') && typeof attributes.formatCode === 'string') customFormats.set(Number(attributes.numFmtId), attributes.formatCode);
  }
  for (const [formatKey, { baseStyle, numFmtId, formatCode }] of formatByBase) {
    const base = formats[baseStyle];
    if (!base) fail('숫자 표시를 적용할 기본 셀 서식이 없습니다.');
    if (formatCode) {
      const existingCode = customFormats.get(numFmtId);
      if (existingCode && existingCode !== formatCode) fail('숫자 표시 형식 ID가 기존 형식과 충돌합니다.');
      customFormats.set(numFmtId, formatCode);
    }
    let existing = -1;
    for (let index = 0; index < formats.length; index++) {
      if (model.styleFormats[index]?.numFmtId === numFmtId && normalized(formats[index]) === normalized(base)) { existing = index; break; }
    }
    if (existing >= 0) { byBase.set(formatKey, String(existing)); continue; }
    const newIndex = count + additions.length;
    let clone = base;
    if (/\bnumFmtId\s*=\s*(["'])[^"']*\1/.test(clone)) clone = clone.replace(/\bnumFmtId\s*=\s*(["'])[^"']*\1/, 'numFmtId="' + numFmtId + '"');
    else clone = clone.replace(/^<xf\b/, '<xf numFmtId="' + numFmtId + '"');
    const cloneTagEnd = clone.indexOf('>'), cloneOpening = clone.slice(0, cloneTagEnd + 1), cloneRemainder = clone.slice(cloneTagEnd + 1);
    const updatedCloneOpening = /\bapplyNumberFormat\s*=\s*(["'])[^"']*\1/.test(cloneOpening)
      ? cloneOpening.replace(/\bapplyNumberFormat\s*=\s*(["'])[^"']*\1/, 'applyNumberFormat="1"')
      : /\/\s*>$/.test(cloneOpening) ? cloneOpening.replace(/\s*\/\s*>$/, ' applyNumberFormat="1"/>') : cloneOpening.replace(/>$/, ' applyNumberFormat="1">');
    clone = updatedCloneOpening + cloneRemainder;
    additions.push(clone); byBase.set(formatKey, String(newIndex));
  }
  if (!additions.length && !formatByBase.size) return { stylesPath, byBase, xml: null };
  let updated = xml;
  if (additions.length) {
    const newOpening = opening.replace(/\bcount\s*=\s*(["'])\d+\1/, function (_all, quote) { return 'count=' + quote + (count + additions.length) + quote; });
    if (newOpening === opening) fail('셀 서식 개수를 갱신할 수 없습니다.');
    updated = xml.slice(0, start) + newOpening + body + additions.join('') + xml.slice(closeStart);
  }
  if (formatByBase.size && [...formatByBase.values()].some(item => item.formatCode)) {
    const oldCustomIds = new Set(numFmtsMatch ? [...numFmtsMatch[0].matchAll(/<numFmt\b[^>]*\bnumFmtId="(\d+)"/g)].map(match => Number(match[1])) : []);
    const customXml = [...customFormats].filter(([id]) => !oldCustomIds.has(id)).map(([id, code]) => '<numFmt numFmtId="' + id + '" formatCode="' + xmlEscape(code) + '"/>').join('');
    if (numFmtsMatch) {
      const replacement = numFmtsMatch[0].replace(/\bcount="\d+"/, 'count="' + customFormats.size + '"').replace('</numFmts>', () => customXml + '</numFmts>');
      updated = updated.replace(numFmtsMatch[0], () => replacement);
    } else {
      const cellStyleXfs = updated.indexOf('<cellStyleXfs');
      if (cellStyleXfs < 0) fail('사용자 지정 숫자 표시 형식을 추가할 위치를 찾을 수 없습니다.');
      updated = updated.slice(0, cellStyleXfs) + '<numFmts count="' + customFormats.size + '">' + customXml + '</numFmts>' + updated.slice(cellStyleXfs);
    }
  }
  return { stylesPath, byBase, xml: updated };
}

function fontStyleKey(style, font) {
  const color = Object.hasOwn(font, 'color') ? font.color === null ? 'none' : font.color.toUpperCase() : '-';
  const underline = Object.hasOwn(font, 'underline') ? font.underline ? '1' : '0' : '-';
  const size = Object.hasOwn(font, 'fontSize') ? font.fontSize === null ? 'none' : String(font.fontSize) : '-';
  const strike = Object.hasOwn(font, 'strike') ? font.strike ? '1' : '0' : '-';
  const family = Object.hasOwn(font, 'fontFamily') ? font.fontFamily === null ? 'none' : font.fontFamily : '-';
  return Number(style ?? 0) + ':' + (font.bold ? '1' : '0') + ':' + (font.italic ? '1' : '0') + ':' + color + ':' + underline + ':' + size + ':' + strike + ':' + family;
}

function prepareFontStyles(stylesPath, sourceXml, fontByBase) {
  if (!fontByBase.size) return null;
  let xml = sourceXml;
  const fontsMatch = /<fonts\b[^>]*>[\s\S]*?<\/fonts>/.exec(xml);
  const xfsMatch = /<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/.exec(xml);
  if (!fontsMatch || !xfsMatch) fail('글꼴 서식을 추가할 스타일 정의를 찾을 수 없습니다.');
  const readItems = (container, tag) => {
    const items = [];
    let position = container.indexOf('>') + 1;
    while (position > 0 && position < container.length) {
      const start = container.indexOf(`<${tag}`, position);
      if (start < 0) break;
      const tagEnd = container.indexOf('>', start);
      if (tagEnd < 0) fail('셀 서식 항목이 닫히지 않았습니다.');
      if (/\/\s*>$/.test(container.slice(start, tagEnd + 1))) { items.push(container.slice(start, tagEnd + 1)); position = tagEnd + 1; continue; }
      const closing = container.indexOf(`</${tag}>`, tagEnd + 1);
      if (closing < 0) fail('셀 서식 항목이 닫히지 않았습니다.');
      items.push(container.slice(start, closing + tag.length + 3));
      position = closing + tag.length + 3;
    }
    return items;
  };
  const fonts = readItems(fontsMatch[0], 'font');
  const xfs = readItems(xfsMatch[0], 'xf');
  const getBool = (font, tag) => {
    const match = new RegExp(`<${tag}(?:\\s+[^>]*)?(?:\\s*/>|>(?:[\\s\\S]*?)<\\/${tag}>)`).exec(font);
    if (!match) return false;
    const value = /\bval\s*=\s*(["'])(.*?)\1/.exec(match[0])?.[2];
    return value === undefined || !['0', 'false'].includes(value.toLowerCase());
  };
  const setBool = (font, tag, enabled) => {
    const expression = new RegExp(`<${tag}(?:\\s+[^>]*)?(?:\\s*/>|>(?:[\\s\\S]*?)<\\/${tag}>)`);
    const match = expression.exec(font);
    if (enabled && match) return font;
    if (!enabled && !match) return font;
    if (!enabled) return font.slice(0, match.index) + font.slice(match.index + match[0].length);
    const selfClosingRoot = /^<font\b[^>]*\/>/.exec(font);
    if (selfClosingRoot) return selfClosingRoot[0].replace(/\s*\/\s*>$/, `><${tag}/></font>`) + font.slice(selfClosingRoot[0].length);
    const first = font.indexOf('>');
    if (first < 0) fail('글꼴 서식 태그가 유효하지 않습니다.');
    return font.slice(0, first + 1) + `<${tag}/>` + font.slice(first + 1);
  };
  const setSize = (font, size) => {
    if (size === undefined) return font;
    const expression = /<sz(?:\s+[^>]*)?(?:\s*\/\s*>|>(?:[\s\S]*?)<\/sz>)/;
    const match = expression.exec(font);
    if (size === null) return match ? font.slice(0, match.index) + font.slice(match.index + match[0].length) : font;
    const element = '<sz val="' + size + '"/>';
    if (match) return font.slice(0, match.index) + element + font.slice(match.index + match[0].length);
    const selfClosingRoot = /^<font\b[^>]*\/>/.exec(font);
    if (selfClosingRoot) return selfClosingRoot[0].replace(/\s*\/\s*>$/, '>' + element + '</font>') + font.slice(selfClosingRoot[0].length);
    const insertion = /<(?:u|vertAlign|scheme)\b|<\/font>/.exec(font);
    if (!insertion) fail('글꼴 크기를 추가할 위치를 찾을 수 없습니다.');
    return font.slice(0, insertion.index) + element + font.slice(insertion.index);
  };
  const setColor = (font, color) => {
    if (color === undefined) return font;
    const expression = /<color(?:\s+[^>]*)?(?:\s*\/\s*>|>(?:[\s\S]*?)<\/color>)/;
    const match = expression.exec(font);
    if (color === null) return match ? font.slice(0, match.index) + font.slice(match.index + match[0].length) : font;
    const element = '<color rgb="FF' + color.slice(1).toUpperCase() + '"/>';
    if (match) return font.slice(0, match.index) + element + font.slice(match.index + match[0].length);
    const selfClosingRoot = /^<font\b[^>]*\/>/.exec(font);
    if (selfClosingRoot) return selfClosingRoot[0].replace(/\s*\/\s*>$/, '>' + element + '</font>') + font.slice(selfClosingRoot[0].length);
    const insertion = /<(?:sz|u|vertAlign|scheme)\b|<\/font>/.exec(font);
    if (!insertion) fail('글꼴 색을 추가할 위치를 찾을 수 없습니다.');
    return font.slice(0, insertion.index) + element + font.slice(insertion.index);
  };
  const setFamily = (font, family) => {
    if (family === undefined) return font;
    const expression = /<name(?:\s+[^>]*)?(?:\s*\/\s*>|>(?:[\s\S]*?)<\/name>)/;
    const match = expression.exec(font);
    if (family === null) return match ? font.slice(0, match.index) + font.slice(match.index + match[0].length) : font;
    const element = '<name val="' + xmlEscape(family) + '"/>';
    if (match) return font.slice(0, match.index) + element + font.slice(match.index + match[0].length);
    const selfClosingRoot = /^<font\b[^>]*\/>/.exec(font);
    if (selfClosingRoot) return selfClosingRoot[0].replace(/\s*\/\s*>$/, '>' + element + '</font>') + font.slice(selfClosingRoot[0].length);
    const insertion = /<(?:b|i|u|strike|sz|color|vertAlign|scheme)\b|<\/font>/.exec(font);
    if (!insertion) fail('글꼴 계열을 추가할 위치를 찾을 수 없습니다.');
    return font.slice(0, insertion.index) + element + font.slice(insertion.index);
  };
  const getFontId = xf => Number(parseTagAttributes(xf).fontId) || 0;
  const normalizedXf = xf => xf.replace(/\s+fontId\s*=\s*(["'])\d+\1/, ' fontId="0"').replace(/\s*\/>$/, '>').replace(/\s*>$/, '>');
  const fontCount = Number(parseTagAttributes(fontsMatch[0].slice(0, fontsMatch[0].indexOf('>') + 1)).count);
  const xfCount = Number(parseTagAttributes(xfsMatch[0].slice(0, xfsMatch[0].indexOf('>') + 1)).count);
  if (!Number.isInteger(fontCount) || fontCount !== fonts.length || !Number.isInteger(xfCount) || xfCount !== xfs.length) fail('글꼴·셀 서식 목록의 개수가 일치하지 않습니다.');
  const byBase = new Map(), fontAdditions = [], xfAdditions = [];
  for (const [key, { baseStyle, bold, italic, color, underline, fontSize, fontFamily, strike }] of fontByBase) {
    const base = xfs[baseStyle];
    if (!base) fail('글꼴 표시를 적용할 셀 서식이 없습니다.');
    const baseFontId = getFontId(base), baseFont = fonts[baseFontId];
    if (!baseFont) fail('글꼴 표시를 적용할 글꼴 정의가 없습니다.');
    let updatedFont = setBool(setBool(baseFont, 'b', bold), 'i', italic);
    if (underline !== undefined) updatedFont = setBool(updatedFont, 'u', underline);
    if (strike !== undefined) updatedFont = setBool(updatedFont, 'strike', strike);
    updatedFont = setFamily(setColor(setSize(updatedFont, fontSize), color), fontFamily);
    let fontId = fonts.indexOf(updatedFont);
    if (fontId < 0) {
      const pendingFont = fontAdditions.indexOf(updatedFont);
      if (pendingFont >= 0) fontId = fontCount + pendingFont;
      else { fontId = fontCount + fontAdditions.length; fontAdditions.push(updatedFont); }
    }
    let existing = -1;
    for (let index = 0; index < xfs.length + xfAdditions.length; index++) {
      const candidate = xfs[index] ?? xfAdditions[index - xfs.length];
      if (normalizedXf(candidate) !== normalizedXf(base)) continue;
      const candidateFontId = getFontId(candidate);
      const candidateFont = fonts[candidateFontId] ?? fontAdditions[candidateFontId - fontCount];
      if (candidateFont === updatedFont) { existing = index; break; }
    }
    if (existing >= 0) { byBase.set(key, String(existing)); continue; }
    let clone = base;
    if (/\bfontId\s*=\s*(["'])[^"']*\1/.test(clone)) clone = clone.replace(/\bfontId\s*=\s*(["'])[^"']*\1/, 'fontId="' + fontId + '"');
    else clone = clone.replace(/^<xf\b/, '<xf fontId="' + fontId + '"');
    const cloneTagEnd = clone.indexOf('>'), cloneOpening = clone.slice(0, cloneTagEnd + 1), cloneRemainder = clone.slice(cloneTagEnd + 1);
    const updatedCloneOpening = /\bapplyFont\s*=\s*(["'])[^"']*\1/.test(cloneOpening)
      ? cloneOpening.replace(/\bapplyFont\s*=\s*(["'])[^"']*\1/, 'applyFont="1"')
      : /\/\s*>$/.test(cloneOpening) ? cloneOpening.replace(/\s*\/\s*>$/, ' applyFont="1"/>') : cloneOpening.replace(/>$/, ' applyFont="1">');
    clone = updatedCloneOpening + cloneRemainder;
    const cloneIndex = xfCount + xfAdditions.length;
    xfAdditions.push(clone);
    byBase.set(key, String(cloneIndex));
  }
  const replaceContainer = (source, match, tag, count, additions) => {
    if (!additions.length) return source;
    const container = match[0], start = match.index;
    const openEnd = container.indexOf('>') + 1, closeStart = container.lastIndexOf(`</${tag}>`);
    const opening = container.slice(0, openEnd).replace(/\bcount\s*=\s*(["'])\d+\1/, (_all, quote) => 'count=' + quote + (count + additions.length) + quote);
    return source.slice(0, start) + opening + container.slice(openEnd, closeStart) + additions.join('') + container.slice(closeStart) + source.slice(start + container.length);
  };
  xml = replaceContainer(xml, fontsMatch, 'fonts', fontCount, fontAdditions);
  const updatedXfsMatch = /<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/.exec(xml);
  xml = replaceContainer(xml, updatedXfsMatch, 'cellXfs', xfCount, xfAdditions);
  return { stylesPath, byBase, xml };
}

function prepareFillStyles(stylesPath, sourceXml, fillByBase) {
  if (!fillByBase.size) return null;
  let xml = sourceXml;
  const fillsMatch = /<fills\b[^>]*>[\s\S]*?<\/fills>/.exec(xml);
  const xfsMatch = /<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/.exec(xml);
  if (!fillsMatch || !xfsMatch) fail('채우기 서식을 추가할 스타일 정의를 찾을 수 없습니다.');
  const readItems = (container, tag) => {
    const items = [];
    let position = container.indexOf('>') + 1;
    while (position > 0 && position < container.length) {
      const start = container.indexOf(`<${tag}`, position);
      if (start < 0) break;
      const tagEnd = container.indexOf('>', start);
      if (tagEnd < 0) fail('셀 서식 항목이 닫히지 않았습니다.');
      if (/\/\s*>$/.test(container.slice(start, tagEnd + 1))) { items.push(container.slice(start, tagEnd + 1)); position = tagEnd + 1; continue; }
      const closing = container.indexOf(`</${tag}>`, tagEnd + 1);
      if (closing < 0) fail('셀 서식 항목이 닫히지 않았습니다.');
      items.push(container.slice(start, closing + tag.length + 3));
      position = closing + tag.length + 3;
    }
    return items;
  };
  const fills = readItems(fillsMatch[0], 'fill'), xfs = readItems(xfsMatch[0], 'xf');
  const fillsOpening = fillsMatch[0].slice(0, fillsMatch[0].indexOf('>') + 1), xfsOpening = xfsMatch[0].slice(0, xfsMatch[0].indexOf('>') + 1);
  const fillCount = Number(parseTagAttributes(fillsOpening).count), xfCount = Number(parseTagAttributes(xfsOpening).count);
  if (!Number.isInteger(fillCount) || fillCount !== fills.length || !Number.isInteger(xfCount) || xfCount !== xfs.length) fail('채우기·셀 서식 목록의 개수가 일치하지 않습니다.');
  const getFillId = xf => Number(parseTagAttributes(xf).fillId) || 0;
  const normalizedXf = xf => xf.replace(/\s+fillId\s*=\s*(["'])\d+\1/, ' fillId="0"').replace(/\s*\/>$/, '>').replace(/\s*>$/, '>');
  const fillColor = fill => {
    if (!/<patternFill\b[^>]*\bpatternType\s*=\s*(["'])solid\1/.test(fill)) return null;
    const color = /<fgColor\b[^>]*\brgb\s*=\s*(["'])(?:FF)?([0-9A-F]{6})\1/i.exec(fill)?.[2];
    return color ? '#' + color.toUpperCase() : null;
  };
  const byBase = new Map(), fillAdditions = [], xfAdditions = [];
  for (const [key, { baseStyle, color }] of fillByBase) {
    const base = xfs[baseStyle];
    if (!base) fail('채우기 표시를 적용할 셀 서식이 없습니다.');
    let fillId = color === null ? 0 : fills.findIndex(fill => fillColor(fill) === color);
    if (fillId < 0) {
      const pending = fillAdditions.findIndex(fill => fillColor(fill) === color);
      if (pending >= 0) fillId = fillCount + pending;
      else {
        fillId = fillCount + fillAdditions.length;
        fillAdditions.push('<fill><patternFill patternType="solid"><fgColor rgb="FF' + color.slice(1) + '"/><bgColor indexed="64"/></patternFill></fill>');
      }
    }
    let existing = -1;
    for (let index = 0; index < xfs.length + xfAdditions.length; index++) {
      const candidate = xfs[index] ?? xfAdditions[index - xfs.length];
      if (normalizedXf(candidate) !== normalizedXf(base) || getFillId(candidate) !== fillId) continue;
      existing = index; break;
    }
    if (existing >= 0) { byBase.set(key, String(existing)); continue; }
    let clone = base;
    if (/\bfillId\s*=\s*(["'])[^"']*\1/.test(clone)) clone = clone.replace(/\bfillId\s*=\s*(["'])[^"']*\1/, 'fillId="' + fillId + '"');
    else clone = clone.replace(/^<xf\b/, '<xf fillId="' + fillId + '"');
    const cloneTagEnd = clone.indexOf('>'), cloneOpening = clone.slice(0, cloneTagEnd + 1), cloneRemainder = clone.slice(cloneTagEnd + 1);
    const updatedCloneOpening = /\bapplyFill\s*=\s*(["'])[^"']*\1/.test(cloneOpening)
      ? cloneOpening.replace(/\bapplyFill\s*=\s*(["'])[^"']*\1/, 'applyFill="1"')
      : /\/\s*>$/.test(cloneOpening) ? cloneOpening.replace(/\s*\/\s*>$/, ' applyFill="1"/>') : cloneOpening.replace(/>$/, ' applyFill="1">');
    xfAdditions.push(updatedCloneOpening + cloneRemainder);
    byBase.set(key, String(xfCount + xfAdditions.length - 1));
  }
  const replaceContainer = (source, match, tag, count, additions) => {
    if (!additions.length) return source;
    const container = match[0], start = match.index;
    const openEnd = container.indexOf('>') + 1, closeStart = container.lastIndexOf(`</${tag}>`);
    const opening = container.slice(0, openEnd).replace(/\bcount\s*=\s*(["'])\d+\1/, (_all, quote) => 'count=' + quote + (count + additions.length) + quote);
    return source.slice(0, start) + opening + container.slice(openEnd, closeStart) + additions.join('') + container.slice(closeStart) + source.slice(start + container.length);
  };
  xml = replaceContainer(xml, fillsMatch, 'fills', fillCount, fillAdditions);
  const updatedXfsMatch = /<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/.exec(xml);
  xml = replaceContainer(xml, updatedXfsMatch, 'cellXfs', xfCount, xfAdditions);
  return { stylesPath, byBase, xml };
}

function prepareAlignmentStyles(stylesPath, sourceXml, alignmentByBase) {
  if (!alignmentByBase.size) return null;
  let xml = sourceXml;
  const xfsMatch = /<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/.exec(xml);
  if (!xfsMatch) fail('셀 맞춤을 추가할 셀 서식 목록을 찾을 수 없습니다.');
  const xfs = [];
  let position = xfsMatch[0].indexOf('>') + 1;
  while (position > 0 && position < xfsMatch[0].length) {
    const start = xfsMatch[0].indexOf('<xf', position);
    if (start < 0) break;
    const tagEnd = xfsMatch[0].indexOf('>', start);
    if (tagEnd < 0) fail('셀 맞춤 XF 태그가 닫히지 않았습니다.');
    if (/\/\s*>$/.test(xfsMatch[0].slice(start, tagEnd + 1))) { xfs.push(xfsMatch[0].slice(start, tagEnd + 1)); position = tagEnd + 1; continue; }
    const closing = xfsMatch[0].indexOf('</xf>', tagEnd + 1);
    if (closing < 0) fail('셀 맞춤 XF 태그가 닫히지 않았습니다.');
    xfs.push(xfsMatch[0].slice(start, closing + 5));
    position = closing + 5;
  }
  const opening = xfsMatch[0].slice(0, xfsMatch[0].indexOf('>') + 1), xfCount = Number(parseTagAttributes(opening).count);
  if (!Number.isInteger(xfCount) || xfCount !== xfs.length) fail('셀 맞춤 서식 목록의 개수가 일치하지 않습니다.');
  const setAlignment = (xf, axis, value) => {
    const current = /<alignment\b[^>]*(?:\/\s*>|>[\s\S]*?<\/alignment>)/.exec(xf);
    let element;
    if (current) {
      const end = current[0].indexOf('>'), attrs = parseTagAttributes(current[0].slice(0, end + 1));
      if (current[0].slice(end + 1).replace(/<\/alignment>$/, '').trim()) fail('셀 맞춤 요소에 지원하지 않는 내부 데이터가 있습니다.');
      attrs[axis] = value;
      element = '<alignment' + Object.entries(attrs).map(([name, item]) => ' ' + name + '="' + xmlEscape(item) + '"').join('') + '/>';
      xf = xf.slice(0, current.index) + element + xf.slice(current.index + current[0].length);
    } else {
      element = '<alignment ' + axis + '="' + value + '"/>';
      const protection = xf.indexOf('<protection');
      const closing = xf.lastIndexOf('</xf>');
      if (closing < 0) {
        const selfClosing = /\/>$/.exec(xf);
        if (!selfClosing) fail('셀 맞춤 XF 태그가 유효하지 않습니다.');
        xf = xf.slice(0, selfClosing.index) + '>' + element + '</xf>';
      } else {
        const insertion = protection >= 0 ? protection : closing;
        xf = xf.slice(0, insertion) + element + xf.slice(insertion);
      }
    }
    const tagEnd = xf.indexOf('>'), xfOpening = xf.slice(0, tagEnd + 1);
    if (/\bapplyAlignment\s*=\s*(["'])[^"']*\1/.test(xfOpening)) return xfOpening.replace(/\bapplyAlignment\s*=\s*(["'])[^"']*\1/, 'applyAlignment="1"') + xf.slice(tagEnd + 1);
    if (/\/\s*>$/.test(xfOpening)) return xfOpening.replace(/\s*\/\s*>$/, ' applyAlignment="1"/>') + xf.slice(tagEnd + 1);
    return xfOpening.replace(/>$/, ' applyAlignment="1">') + xf.slice(tagEnd + 1);
  };
  const normalized = xf => xf.replace(/\s+applyAlignment\s*=\s*(["'])[^"']*\1/, '').replace(/<alignment\b[^>]*(?:\/\s*>|>[\s\S]*?<\/alignment>)/, '').replace(/\s*\/>$/, '>').replace(/\s*>$/, '>');
  const byBase = new Map(), additions = [];
  const getStyle = index => xfs[index] ?? additions[index - xfs.length];
  const getOrAdd = (baseIndex, axis, value) => {
    const base = getStyle(baseIndex);
    if (!base) fail('맞춤을 적용할 셀 서식이 없습니다.');
    const target = setAlignment(base, axis, value), normalizedBase = normalized(base);
    for (let index = 0; index < xfs.length + additions.length; index++) {
      const candidate = getStyle(index);
      if (normalized(candidate) === normalizedBase && setAlignment(candidate, axis, value) === candidate) return index;
    }
    const pending = additions.indexOf(target);
    const index = pending >= 0 ? xfCount + pending : xfCount + additions.length;
    if (pending < 0) additions.push(target);
    return index;
  };
  for (const [key, { baseStyle, alignment }] of alignmentByBase) {
    let style = baseStyle;
    for (const [axis, value] of Object.entries(alignment)) style = getOrAdd(style, axis, value);
    byBase.set(key, String(style));
  }
  if (additions.length) {
    const container = xfsMatch[0], start = xfsMatch.index, openEnd = container.indexOf('>') + 1, closeStart = container.lastIndexOf('</cellXfs>');
    const updatedOpening = container.slice(0, openEnd).replace(/\bcount\s*=\s*(["'])\d+\1/, (_all, quote) => 'count=' + quote + (xfCount + additions.length) + quote);
    xml = xml.slice(0, start) + updatedOpening + container.slice(openEnd, closeStart) + additions.join('') + container.slice(closeStart) + xml.slice(start + container.length);
  }
  return { stylesPath, byBase, xml };
}

function prepareBorderStyles(stylesPath, sourceXml, borderByBase) {
  if (!borderByBase.size) return null;
  let xml = sourceXml;
  const bordersMatch = /<borders\b[^>]*>[\s\S]*?<\/borders>/.exec(xml), xfsMatch = /<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/.exec(xml);
  if (!bordersMatch || !xfsMatch) fail('셀 테두리를 추가할 스타일 목록을 찾을 수 없습니다.');
  const readItems = (container, tag) => {
    const items = [];
    let position = container.indexOf('>') + 1;
    while (position > 0 && position < container.length) {
      const start = container.indexOf('<' + tag, position);
      if (start < 0) break;
      const tagEnd = container.indexOf('>', start);
      if (tagEnd < 0) fail('테두리 서식 항목이 닫히지 않았습니다.');
      if (/\/\s*>$/.test(container.slice(start, tagEnd + 1))) { items.push(container.slice(start, tagEnd + 1)); position = tagEnd + 1; continue; }
      const closing = container.indexOf('</' + tag + '>', tagEnd + 1);
      if (closing < 0) fail('테두리 서식 항목이 닫히지 않았습니다.');
      items.push(container.slice(start, closing + tag.length + 3));
      position = closing + tag.length + 3;
    }
    return items;
  };
  const borders = readItems(bordersMatch[0], 'border'), xfs = readItems(xfsMatch[0], 'xf');
  const borderOpening = bordersMatch[0].slice(0, bordersMatch[0].indexOf('>') + 1), xfOpening = xfsMatch[0].slice(0, xfsMatch[0].indexOf('>') + 1);
  const borderCount = Number(parseTagAttributes(borderOpening).count), xfCount = Number(parseTagAttributes(xfOpening).count);
  if (!Number.isInteger(borderCount) || borderCount !== borders.length || !Number.isInteger(xfCount) || xfCount !== xfs.length) fail('테두리·셀 서식 목록의 개수가 일치하지 않습니다.');
  const sideOrder = ['left', 'right', 'top', 'bottom', 'diagonal', 'vertical', 'horizontal', 'start', 'end'];
  const setBorder = (border, changes) => {
    const rootEnd = border.indexOf('>');
    if (rootEnd < 0) fail('테두리 요소가 유효하지 않습니다.');
    const rootOpen = border.slice(0, rootEnd + 1), rootAttrs = parseTagAttributes(rootOpen);
    let body = /\/\s*>$/.test(rootOpen) ? '' : border.slice(rootEnd + 1, border.lastIndexOf('</border>'));
    for (const [side, edge] of Object.entries(changes)) {
      const expression = new RegExp(`<${side}\\b[^>]*(?:\\/\\s*>|>[\\s\\S]*?<\\/${side}>)`), match = expression.exec(body);
      if (edge === null) {
        if (match) body = body.slice(0, match.index) + body.slice(match.index + match[0].length);
        continue;
      }
      const color = edge.color || '#000000';
      const element = '<' + side + ' style="' + edge.line + '"><color rgb="FF' + color.slice(1).toUpperCase() + '"/></' + side + '>';
      if (match) body = body.slice(0, match.index) + element + body.slice(match.index + match[0].length);
      else {
        const next = sideOrder.indexOf(side);
        let insertion = body.length;
        for (const candidateSide of sideOrder.slice(next + 1)) {
          const candidate = body.indexOf('<' + candidateSide);
          if (candidate >= 0) { insertion = candidate; break; }
        }
        body = body.slice(0, insertion) + element + body.slice(insertion);
      }
    }
    const attrs = Object.entries(rootAttrs).map(([name, value]) => ' ' + name + '="' + xmlEscape(value) + '"').join('');
    return '<border' + attrs + '>' + body + '</border>';
  };
  const normalizedXf = xf => xf.replace(/\s+borderId\s*=\s*(["'])\d+\1/, ' borderId="0"').replace(/\s+applyBorder\s*=\s*(["'])[^"']*\1/, '').replace(/\s*\/>$/, '>').replace(/\s*>$/, '>');
  const getBorderId = xf => Number(parseTagAttributes(xf).borderId) || 0;
  const borderAdditions = [], xfAdditions = [], byBase = new Map();
  for (const [key, { baseStyle, changes }] of borderByBase) {
    const base = xfs[baseStyle];
    if (!base) fail('테두리를 적용할 셀 서식이 없습니다.');
    const baseBorderId = getBorderId(base), baseBorder = borders[baseBorderId];
    if (!baseBorder) fail('테두리 정의를 찾을 수 없습니다.');
    const updatedBorder = setBorder(baseBorder, changes);
    let borderId = borders.indexOf(updatedBorder);
    if (borderId < 0) {
      const pending = borderAdditions.indexOf(updatedBorder);
      borderId = pending >= 0 ? borderCount + pending : borderCount + borderAdditions.length;
      if (pending < 0) borderAdditions.push(updatedBorder);
    }
    let existing = -1;
    for (let index = 0; index < xfs.length + xfAdditions.length; index++) {
      const candidate = xfs[index] ?? xfAdditions[index - xfs.length];
      if (normalizedXf(candidate) !== normalizedXf(base) || getBorderId(candidate) !== borderId) continue;
      existing = index; break;
    }
    if (existing >= 0) { byBase.set(key, String(existing)); continue; }
    let clone = base;
    if (/\bborderId\s*=\s*(["'])[^"']*\1/.test(clone)) clone = clone.replace(/\bborderId\s*=\s*(["'])[^"']*\1/, 'borderId="' + borderId + '"');
    else clone = clone.replace(/^<xf\b/, '<xf borderId="' + borderId + '"');
    const tagEnd = clone.indexOf('>'), opening = clone.slice(0, tagEnd + 1), remainder = clone.slice(tagEnd + 1);
    const updatedOpening = /\bapplyBorder\s*=\s*(["'])[^"']*\1/.test(opening)
      ? opening.replace(/\bapplyBorder\s*=\s*(["'])[^"']*\1/, 'applyBorder="1"')
      : /\/\s*>$/.test(opening) ? opening.replace(/\s*\/\s*>$/, ' applyBorder="1"/>') : opening.replace(/>$/, ' applyBorder="1">');
    xfAdditions.push(updatedOpening + remainder);
    byBase.set(key, String(xfCount + xfAdditions.length - 1));
  }
  const replace = (source, match, tag, count, additions) => {
    if (!additions.length) return source;
    const container = match[0], start = match.index, openEnd = container.indexOf('>') + 1, closeStart = container.lastIndexOf('</' + tag + '>');
    const updatedOpening = container.slice(0, openEnd).replace(/\bcount\s*=\s*(["'])\d+\1/, (_all, quote) => 'count=' + quote + (count + additions.length) + quote);
    return source.slice(0, start) + updatedOpening + container.slice(openEnd, closeStart) + additions.join('') + container.slice(closeStart) + source.slice(start + container.length);
  };
  xml = replace(xml, bordersMatch, 'borders', borderCount, borderAdditions);
  const updatedXfsMatch = /<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/.exec(xml);
  xml = replace(xml, updatedXfsMatch, 'cellXfs', xfCount, xfAdditions);
  return { stylesPath, byBase, xml };
}

function ensureWorkbookStyles(model) {
  if (model.stylesPath) return model;
  const files = { ...model.files }, relPath = 'xl/_rels/workbook.xml.rels';
  const relXml = new TextDecoder().decode(files[relPath]), relations = parseRelationships(parseXml(files[relPath], relPath).root, 'xl');
  let number = 1; while (relations.has('rId' + number)) number++;
  let stylesNumber = 1, stylesPath = 'xl/styles.xml';
  while (files[stylesPath]) stylesPath = 'xl/styles' + ++stylesNumber + '.xml';
  files[stylesPath] = new TextEncoder().encode('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>');
  files[relPath] = new TextEncoder().encode(appendRelationship(relXml, '<Relationship Id="rId' + number + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="' + path.basename(stylesPath) + '"/>', relPath));
  const types = new TextDecoder().decode(files['[Content_Types].xml']);
  files['[Content_Types].xml'] = new TextEncoder().encode(types.replace('</Types>', '<Override PartName="/' + stylesPath + '" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'));
  const styled = analyze(zipSync(files, { level: 0 }), model.extension);
  if (!styled.writable || !styled.stylesPath) fail('기본 셀 서식을 추가할 수 없습니다.');
  return styled;
}

function patchCells(model, edits) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  model = preserveImportedFilterVisibility(model);
  if (!Array.isArray(edits) || edits.length > LIMITS.cells * 2) fail('셀 수정 개수가 유효하지 않습니다.');
  if (!model.stylesPath && edits.some(edit => edit && ['alignment', 'font', 'fillColor', 'border', 'numberFormatId'].some(key => Object.hasOwn(edit, key)))) {
    return patchCells(ensureWorkbookStyles(model), edits);
  }
  const formatsByBase = new Map();
  for (const edit of edits) if (edit?.numberFormatId !== undefined) {
    const sheet = model.sheets[edit.sheetIndex], address = edit.address, cell = sheet?.cells.get(address);
    const style = Object.hasOwn(edit, 'style') ? edit.style : cell?.style;
      const baseStyle = style === null || style === undefined ? 0 : Number(style);
    if (!Number.isInteger(baseStyle) || baseStyle < 0 || baseStyle >= model.styleFormats.length) fail('숫자 표시 서식 기준 인덱스가 유효하지 않습니다.');
    if (model.styleFormats[baseStyle]?.numFmtId === 0 || edit.forceNumberFormat === true) formatsByBase.set(baseStyle + ':' + edit.numberFormatId, { baseStyle, numFmtId: edit.numberFormatId, formatCode: edit.numberFormatCode });
  }
  const numberStyles = prepareNumberStyles(model, formatsByBase);
  const requestedFonts = new Map();
  for (const edit of edits) if (edit?.font !== undefined) {
    const sheet = model.sheets[edit.sheetIndex], address = edit.address, cell = sheet?.cells.get(address);
    const baseStyle = Object.hasOwn(edit, 'style') ? edit.style : cell?.style;
    let style = baseStyle === null || baseStyle === undefined ? 0 : Number(baseStyle);
    if (!Number.isInteger(style) || style < 0 || !edit.font || typeof edit.font !== 'object' || Object.keys(edit.font).some(key => !['bold', 'italic', 'color', 'underline', 'fontSize', 'fontFamily', 'strike'].includes(key)) || typeof edit.font.bold !== 'boolean' || typeof edit.font.italic !== 'boolean' || (Object.hasOwn(edit.font, 'color') && edit.font.color !== null && (typeof edit.font.color !== 'string' || !/^#[0-9A-F]{6}$/i.test(edit.font.color))) || (Object.hasOwn(edit.font, 'underline') && typeof edit.font.underline !== 'boolean') || (Object.hasOwn(edit.font, 'strike') && typeof edit.font.strike !== 'boolean') || (Object.hasOwn(edit.font, 'fontSize') && edit.font.fontSize !== null && (typeof edit.font.fontSize !== 'number' || !Number.isFinite(edit.font.fontSize) || edit.font.fontSize < 1 || edit.font.fontSize > 409 || edit.font.fontSize * 2 % 1 !== 0)) || (Object.hasOwn(edit.font, 'fontFamily') && edit.font.fontFamily !== null && (typeof edit.font.fontFamily !== 'string' || !edit.font.fontFamily.trim() || edit.font.fontFamily.length > 64))) fail('글꼴 서식 요청이 유효하지 않습니다.');
    if (edit.numberFormatId !== undefined) style = numberStyles?.byBase.get(style + ':' + edit.numberFormatId) ?? style;
    const key = fontStyleKey(style, edit.font);
    requestedFonts.set(key, { baseStyle: style, bold: edit.font.bold, italic: edit.font.italic, ...(Object.hasOwn(edit.font, 'color') ? { color: edit.font.color === null ? null : edit.font.color.toUpperCase() } : {}), ...(Object.hasOwn(edit.font, 'underline') ? { underline: edit.font.underline } : {}), ...(Object.hasOwn(edit.font, 'fontSize') ? { fontSize: edit.font.fontSize } : {}), ...(Object.hasOwn(edit.font, 'fontFamily') ? { fontFamily: edit.font.fontFamily === null ? null : edit.font.fontFamily.trim() } : {}), ...(Object.hasOwn(edit.font, 'strike') ? { strike: edit.font.strike } : {}) });
  }
  const fontStylesXml = numberStyles?.xml || (model.stylesPath && model.files[model.stylesPath] ? new TextDecoder().decode(model.files[model.stylesPath]) : null);
  if (requestedFonts.size && !fontStylesXml) fail('글꼴 서식을 추가할 스타일 정의를 찾을 수 없습니다.');
  const fontStyles = requestedFonts.size ? prepareFontStyles(model.stylesPath, fontStylesXml, requestedFonts) : null;
  const requestedFills = new Map();
  for (const edit of edits) if (Object.hasOwn(edit || {}, 'fillColor')) {
    const sheet = model.sheets[edit.sheetIndex], address = edit.address, cell = sheet?.cells.get(address);
    const baseStyle = Object.hasOwn(edit, 'style') ? edit.style : cell?.style;
    let style = baseStyle === null || baseStyle === undefined ? 0 : Number(baseStyle);
    const color = edit.fillColor;
    if (!Number.isInteger(style) || style < 0 || (color !== null && (typeof color !== 'string' || !/^#[0-9A-F]{6}$/i.test(color)))) fail('채우기 서식 요청이 유효하지 않습니다.');
    if (edit.numberFormatId !== undefined) style = numberStyles?.byBase.get(style + ':' + edit.numberFormatId) ?? style;
    if (edit.font !== undefined) {
      style = fontStyles?.byBase.get(fontStyleKey(style, edit.font)) ?? style;
    }
    const key = style + ':' + (color === null ? 'none' : color.toUpperCase());
    requestedFills.set(key, { baseStyle: style, color: color === null ? null : color.toUpperCase() });
  }
  const fillStylesXml = fontStyles?.xml || numberStyles?.xml || (model.stylesPath && model.files[model.stylesPath] ? new TextDecoder().decode(model.files[model.stylesPath]) : null);
  if (requestedFills.size && !fillStylesXml) fail('채우기 서식을 추가할 스타일 정의를 찾을 수 없습니다.');
  const fillStyles = requestedFills.size ? prepareFillStyles(model.stylesPath, fillStylesXml, requestedFills) : null;
  const alignmentStyleKey = (style, alignment) => Number(style ?? 0) + ':' + Object.keys(alignment).sort().map(axis => axis + '=' + alignment[axis]).join(',');
  const requestedAlignments = new Map();
  for (const edit of edits) if (edit?.alignment !== undefined) {
    const sheet = model.sheets[edit.sheetIndex], address = edit.address, cell = sheet?.cells.get(address);
    const baseStyle = Object.hasOwn(edit, 'style') ? edit.style : cell?.style;
    let style = baseStyle === null || baseStyle === undefined ? 0 : Number(baseStyle);
    const alignment = edit.alignment;
    if (!Number.isInteger(style) || style < 0 || !alignment || typeof alignment !== 'object' || Object.keys(alignment).length < 1 || Object.keys(alignment).length > 3 || Object.keys(alignment).some(axis => !['horizontal', 'vertical', 'wrapText'].includes(axis)) || (Object.hasOwn(alignment, 'horizontal') && !['left', 'center', 'right'].includes(alignment.horizontal)) || (Object.hasOwn(alignment, 'vertical') && !['top', 'center', 'bottom'].includes(alignment.vertical)) || (Object.hasOwn(alignment, 'wrapText') && typeof alignment.wrapText !== 'boolean')) fail('셀 맞춤 요청이 유효하지 않습니다.');
    if (edit.numberFormatId !== undefined) style = numberStyles?.byBase.get(style + ':' + edit.numberFormatId) ?? style;
    if (edit.font !== undefined) {
      style = fontStyles?.byBase.get(fontStyleKey(style, edit.font)) ?? style;
    }
    if (Object.hasOwn(edit, 'fillColor')) style = fillStyles?.byBase.get(style + ':' + (edit.fillColor === null ? 'none' : String(edit.fillColor).toUpperCase())) ?? style;
    const key = alignmentStyleKey(style, alignment);
    requestedAlignments.set(key, { baseStyle: style, alignment });
  }
  const alignmentStylesXml = fillStyles?.xml || fontStyles?.xml || numberStyles?.xml || (model.stylesPath && model.files[model.stylesPath] ? new TextDecoder().decode(model.files[model.stylesPath]) : null);
  if (requestedAlignments.size && !alignmentStylesXml) fail('셀 맞춤을 추가할 스타일 정의를 찾을 수 없습니다.');
  const alignmentStyles = requestedAlignments.size ? prepareAlignmentStyles(model.stylesPath, alignmentStylesXml, requestedAlignments) : null;
  const borderStyleKey = (style, changes) => Number(style ?? 0) + ':' + Object.keys(changes).sort().map(side => side + '=' + JSON.stringify(changes[side])).join(';');
  const requestedBorders = new Map();
  for (const edit of edits) if (edit?.border !== undefined) {
    const sheet = model.sheets[edit.sheetIndex], address = edit.address, cell = sheet?.cells.get(address);
    const baseStyle = Object.hasOwn(edit, 'style') ? edit.style : cell?.style;
    let style = baseStyle === null || baseStyle === undefined ? 0 : Number(baseStyle);
    const changes = edit.border;
    const invalidBorder = !Number.isInteger(style) || style < 0 || !changes || typeof changes !== 'object' || !Object.keys(changes).length || Object.keys(changes).some(side => !['left', 'right', 'top', 'bottom'].includes(side)) || Object.values(changes).some(edge => edge !== null && (!edge || typeof edge !== 'object' || !['hair', 'thin', 'medium', 'thick', 'dashed', 'dotted', 'double'].includes(edge.line) || (edge.color !== undefined && (typeof edge.color !== 'string' || !/^#[0-9A-F]{6}$/i.test(edge.color)))));
    if (invalidBorder) fail('셀 테두리 요청이 유효하지 않습니다.');
    if (edit.numberFormatId !== undefined) style = numberStyles?.byBase.get(style + ':' + edit.numberFormatId) ?? style;
    if (edit.font !== undefined) {
      style = fontStyles?.byBase.get(fontStyleKey(style, edit.font)) ?? style;
    }
    if (Object.hasOwn(edit, 'fillColor')) style = fillStyles?.byBase.get(style + ':' + (edit.fillColor === null ? 'none' : String(edit.fillColor).toUpperCase())) ?? style;
    if (edit.alignment !== undefined) style = alignmentStyles?.byBase.get(alignmentStyleKey(style, edit.alignment)) ?? style;
    const key = borderStyleKey(style, changes);
    requestedBorders.set(key, { baseStyle: style, changes });
  }
  const borderStylesXml = alignmentStyles?.xml || fillStyles?.xml || fontStyles?.xml || numberStyles?.xml || (model.stylesPath && model.files[model.stylesPath] ? new TextDecoder().decode(model.files[model.stylesPath]) : null);
  if (requestedBorders.size && !borderStylesXml) fail('셀 테두리를 추가할 스타일 정의를 찾을 수 없습니다.');
  const borderStyles = requestedBorders.size ? prepareBorderStyles(model.stylesPath, borderStylesXml, requestedBorders) : null;
  const bySheet = new Map(), additions = new Map(), seen = new Set();
  let writesFormula = false, chartSourceEdited = false, chartDataEdited = false;
  const changedChartSourceCells = [];
  for (const edit of edits) {
    const sheetIndex = edit.sheetIndex, address = edit.address, value = edit.value, sheet = model.sheets[sheetIndex];
    const hasFormula = Object.hasOwn(edit, 'formula'), formula = edit.formula, kind = edit.kind;
    const numberFormatId = edit.numberFormatId;
    writesFormula ||= hasFormula && formula !== null;
    const row = rowNumber(address), column = columnNumber(address);
    if (!sheet || !row || row > LIMITS.rows || !column || column > LIMITS.columns || typeof value !== 'string' || value.length > 100000) fail('셀 수정 요청이 유효하지 않습니다.');
    if (hasFormula && formula !== null && (typeof formula !== 'string' || !formula.length || formula.length > 100000 || formula.startsWith('='))) fail('셀 수식이 유효하지 않습니다.');
    if (Object.hasOwn(edit, 'kind') && !['blank', 'string', 'number', 'boolean', 'error'].includes(kind)) fail('셀 값 종류가 유효하지 않습니다.');
    if (Object.hasOwn(edit, 'numberFormatId') && (!([4, 9, 10, 14, 20, 21, 22].includes(numberFormatId) || Number.isInteger(numberFormatId) && numberFormatId >= 164 && numberFormatId <= 65535 && typeof edit.numberFormatCode === 'string' && edit.numberFormatCode.length <= 255) || kind !== 'number' || (Object.hasOwn(edit, 'forceNumberFormat') && typeof edit.forceNumberFormat !== 'boolean'))) fail('숫자 표시 서식 요청이 유효하지 않습니다.');
    const key = sheetIndex + ':' + address;
    if (seen.has(key)) fail('같은 셀을 중복 수정할 수 없습니다.');
    seen.add(key);
    const cell = sheet.cells.get(address), range = sheet.ranges.get(address);
    if (model.hasCharts) {
      const previousValue = cell?.value === null || cell?.value === undefined ? '' : String(cell.value);
      const previousKind = cell?.kind || 'blank';
      const previousFormula = cell?.formula ?? null;
      const nextKind = kind ?? previousKind;
      const nextFormula = hasFormula ? formula : previousFormula;
      if (value !== previousValue || nextKind !== previousKind || nextFormula !== previousFormula) {
        chartDataEdited = true;
        const coordinate = { sheetIndex, row: row - 1, column: column - 1 };
        if (model.chartSourceRanges?.some(source => chartSourceContains(source, coordinate))) {
          chartSourceEdited = true;
          changedChartSourceCells.push(coordinate);
        }
      }
    }
    let style = Object.hasOwn(edit, 'style') ? edit.style : cell?.style;
    let applyNumberStyle = false;
    if (numberFormatId !== undefined) {
      const baseStyle = style === null || style === undefined ? Number(cell?.style) || 0 : Number(style);
      const baseFormat = model.styleFormats[baseStyle];
      if (baseFormat?.numFmtId === 0 || edit.forceNumberFormat === true) { style = numberStyles.byBase.get(baseStyle + ':' + numberFormatId); applyNumberStyle = true; }
    }
    if (edit.font !== undefined) {
      const fontKey = fontStyleKey(style, edit.font);
      style = fontStyles?.byBase.get(fontKey);
      if (style === undefined) fail('글꼴 셀 서식 인덱스를 만들지 못했습니다.');
    }
    if (Object.hasOwn(edit, 'fillColor')) {
      const fillKey = Number(style ?? 0) + ':' + (edit.fillColor === null ? 'none' : String(edit.fillColor).toUpperCase());
      style = fillStyles?.byBase.get(fillKey);
      if (style === undefined) fail('채우기 셀 서식 인덱스를 만들지 못했습니다.');
    }
    if (edit.alignment !== undefined) {
      const alignmentKey = alignmentStyleKey(style, edit.alignment);
      style = alignmentStyles?.byBase.get(alignmentKey);
      if (style === undefined) fail('맞춤 셀 서식 인덱스를 만들지 못했습니다.');
    }
    if (edit.border !== undefined) {
      const borderKey = borderStyleKey(style, edit.border);
      style = borderStyles?.byBase.get(borderKey);
      if (style === undefined) fail('테두리 셀 서식 인덱스를 만들지 못했습니다.');
    }
    if (!cell && !range) {
      if (!additions.has(sheetIndex)) additions.set(sheetIndex, []);
      additions.get(sheetIndex).push({ address, value, style, kind, ...(hasFormula && formula !== null ? { formula } : {}) });
      if (!bySheet.has(sheetIndex)) bySheet.set(sheetIndex, []);
      continue;
    }
    if (cell?.formulaType && cell.formulaType !== 'normal' && cell.formulaType !== 'shared') fail('배열 수식 셀은 배열 범위를 보존하기 위해 수정할 수 없습니다.');
    if (!cell || !range || ((cell.hasFormula || cell.formula || cell.kind === 'error') && !hasFormula)) fail('기존의 단순 값 셀만 수정할 수 있습니다.');
    const originalOpen = sheet.xml.slice(range.start, range.startTagEnd);
    if (/\b[\w.-]+:/.test(originalOpen.replace(/^<[^\s>]+/, ''))) fail('접두사가 있는 셀 속성은 수정할 수 없습니다.');
    const attributes = parseTagAttributes(originalOpen);
    if (Object.hasOwn(edit, 'style') || applyNumberStyle || edit.font !== undefined || Object.hasOwn(edit, 'fillColor') || edit.alignment !== undefined || edit.border !== undefined) {
      if (style === null) delete attributes.s;
      else if (typeof style === 'string' && /^(?:0|[1-9]\d*)$/.test(style)) attributes.s = style;
      else fail('셀 서식 인덱스가 유효하지 않습니다.');
    }
    const keep = Object.entries(attributes).filter(function (entry) { return entry[0] !== 't'; }).map(function (entry) { return ' ' + entry[0] + '="' + xmlEscape(entry[1]) + '"'; }).join('');
    const preserveIsoDate = cell.isoDate !== undefined && !cell.hasFormula && !formula && (kind ?? cell.kind) === 'number' && /^\s*[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?\s*$/.test(value) && Number(value) === Number(cell.value);
    const replacement = preserveIsoDate ? '<c' + keep + ' t="d"><v>' + xmlEscape(cell.isoDate) + '</v></c>' : serializeCellContent(value, kind ?? cell.kind, keep, formula === null ? undefined : formula);
    if (!bySheet.has(sheetIndex)) bySheet.set(sheetIndex, []);
    bySheet.get(sheetIndex).push(Object.assign({}, range, { replacement, column }));
  }
  const editedAddresses = new Map();
  for (const edit of edits) {
    if (!editedAddresses.has(edit.sheetIndex)) editedAddresses.set(edit.sheetIndex, new Set());
    editedAddresses.get(edit.sheetIndex).add(edit.address);
  }
  for (const [sheetIndex, addresses] of editedAddresses) {
    const sheet = model.sheets[sheetIndex];
    for (const group of sheet.sharedFormulaGroups || []) {
      if (!group.members.some(member => addresses.has(member.address))) continue;
      let patches = bySheet.get(sheetIndex);
      if (!patches) { patches = []; bySheet.set(sheetIndex, patches); }
      for (const member of group.members) {
        if (addresses.has(member.address)) continue;
        patches.push({ ...member.formulaRange, replacement: serializeUnsharedFormula(sheet.xml, member.formulaRange, member.formula), column: member.column });
      }
    }
  }
  for (const [sheetIndex, newCells] of additions) addBlankCellOperations(model.sheets[sheetIndex], newCells, bySheet.get(sheetIndex));
  const outputFiles = Object.assign({}, model.files), modifiedPaths = new Set();
  if (borderStyles?.xml || alignmentStyles?.xml || fillStyles?.xml || fontStyles?.xml || numberStyles?.xml) { outputFiles[model.stylesPath] = new TextEncoder().encode(borderStyles?.xml || alignmentStyles?.xml || fillStyles?.xml || fontStyles?.xml || numberStyles.xml); modifiedPaths.add(model.stylesPath); }
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
  if ((model.formulas || writesFormula) && bySheet.size) {
    const workbookXml = new TextDecoder().decode(outputFiles['xl/workbook.xml']);
    outputFiles['xl/workbook.xml'] = new TextEncoder().encode(enableFullCalculationOnLoad(workbookXml));
    modifiedPaths.add('xl/workbook.xml');
  }
  const refreshAllChartSources = chartDataEdited && chartHasFormulaSources(model);
  if (chartSourceEdited || refreshAllChartSources) {
    const editedModel = analyze(zipSync(outputFiles, { level: 0 }), model.extension);
    if (!editedModel.writable) fail('차트 원본 편집 후 패키지 검사에서 거절되었습니다: ' + editedModel.readOnlyReason);
    for (const chart of model.chartPaths || []) {
      outputFiles[chart.path] = patchChartCaches(outputFiles[chart.path], chart.path, chart.sheetIndex, editedModel, refreshAllChartSources ? null : changedChartSourceCells);
      modifiedPaths.add(chart.path);
    }
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

function enableFullCalculationOnLoad(xml) {
  const calcRanges = findElementRanges(xml, 'calcPr');
  if (calcRanges.size > 1) fail('통합문서 계산 설정이 중복되었습니다.');
  if (calcRanges.size === 1) {
    const range = calcRanges.values().next().value;
    const opening = xml.slice(range.start, range.startTagEnd);
    const updated = setTagAttribute(setTagAttribute(opening, 'fullCalcOnLoad', '1'), 'forceFullCalc', '1');
    return xml.slice(0, range.start) + updated + xml.slice(range.startTagEnd);
  }
  const following = ['oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'extLst'];
  const insertionPoints = following.flatMap(name => [...findElementRanges(xml, name).values()].map(range => range.start));
  const closing = xml.lastIndexOf('</workbook>');
  if (closing < 0) fail('통합문서 끝 태그를 찾을 수 없습니다.');
  const at = insertionPoints.length ? Math.min(...insertionPoints) : closing;
  const calculation = '<calcPr fullCalcOnLoad="1" forceFullCalc="1"/>';
  return xml.slice(0, at) + calculation + xml.slice(at);
}

function dimensionView(xml, name) {
  const root = parseXml(new TextEncoder().encode(xml), name).root, data = child(root, 'sheetData');
  const columns = readColumnWidths(root), rows = readRowDimensions(root, data);
  return { columnWidths: columns.widths, defaultColumnWidth: columns.defaultWidth, hiddenColumns: columns.hiddenColumns, rowHeights: rows.heights, defaultRowHeight: rows.defaultHeight, hiddenRows: readHiddenRows(data) };
}

function replaceDrawingMarker(xml, markerNode, marker) {
  const range = findElementRanges(xml, markerNode.name).get('');
  if (!range) fail('차트 위치 마커를 찾을 수 없습니다.');
  let inner = xml.slice(range.start, range.end);
  for (const [name, value] of Object.entries({ col: marker.column, row: marker.row, colOff: Math.round(marker.x * 9525), rowOff: Math.round(marker.y * 9525) })) {
    const node = childrenByLocalName(markerNode, name)[0], leaf = node && findElementRanges(inner, node.name).get('');
    if (!leaf || leaf.selfClosing) fail('차트 위치 좌표를 안전하게 수정할 수 없습니다.');
    inner = inner.slice(0, leaf.startTagEnd) + value + inner.slice(leaf.closingStart);
  }
  return xml.slice(0, range.start) + inner + xml.slice(range.end);
}

// Fixed-size two-cell anchors must be rewritten together with dimension changes.
// A shared drawing is copied before mutation so other worksheets keep their layout.
function updateDimensionDrawings(model, outputFiles, modifiedPaths) {
  for (const sheet of model.sheets) {
    const drawingPath = sheet.chartDrawingPath;
    if (!drawingPath || !modifiedPaths.has(sheet.path)) continue;
    const parsed = parseXml(model.files[drawingPath], drawingPath);
    if (!parsed.root.children.some(node => localName(node) === 'twoCellAnchor' && ['oneCell', 'absolute'].includes(node.attributes.editAs))) continue;
    const beforeAxes = drawingAxes({ ...sheet, hiddenRows: sheet.nativeHiddenRows || sheet.hiddenRows }), afterView = dimensionView(new TextDecoder().decode(outputFiles[sheet.path]), sheet.path), afterAxes = drawingAxes(afterView);
    const operations = [];
    for (const anchor of parsed.root.children) {
      if (localName(anchor) !== 'twoCellAnchor' || !['oneCell', 'absolute'].includes(anchor.attributes.editAs)) continue;
      const geometry = readChartAnchor(anchor, true);
      if (!geometry) fail('고정 차트의 위치를 해석할 수 없어 셀 크기를 변경할 수 없습니다.');
      const from = drawingPoint(geometry.from, beforeAxes), to = drawingPoint(geometry.to, beforeAxes);
      if (!(to.x > from.x && to.y > from.y)) fail('고정 차트의 원본 크기가 유효하지 않아 셀 크기를 변경할 수 없습니다.');
      const naturalFrom = drawingPoint(geometry.from, afterAxes), naturalTo = drawingPoint(geometry.to, afterAxes);
      const nextFrom = geometry.editAs === 'absolute' ? from : naturalFrom, nextTo = { x: nextFrom.x + to.x - from.x, y: nextFrom.y + to.y - from.y };
      if (Math.max(Math.abs(nextFrom.x - naturalFrom.x), Math.abs(nextFrom.y - naturalFrom.y), Math.abs(nextTo.x - naturalTo.x), Math.abs(nextTo.y - naturalTo.y), Math.abs(nextFrom.x - from.x), Math.abs(nextFrom.y - from.y)) < .000001) continue;
      const siblings = parsed.root.children.filter(node => node.name === anchor.name), range = [...findElementRanges(parsed.xml, anchor.name, undefined, true).values()][siblings.indexOf(anchor)];
      if (!range) fail('고정 차트의 XML 위치를 찾을 수 없습니다.');
      let xml = parsed.xml.slice(range.start, range.end);
      if (geometry.editAs === 'absolute') xml = replaceDrawingMarker(xml, childrenByLocalName(anchor, 'from')[0], drawingMarker(nextFrom, afterAxes));
      xml = replaceDrawingMarker(xml, childrenByLocalName(anchor, 'to')[0], drawingMarker(nextTo, afterAxes));
      const transform = firstDescendant(anchor, ['xfrm']), extent = transform && childrenByLocalName(transform, 'ext')[0];
      if (extent && Number(extent.attributes.cx) > 0 && Number(extent.attributes.cy) > 0) {
        for (const [name, values] of [['off', { x: nextFrom.x, y: nextFrom.y }], ['ext', { cx: nextTo.x - nextFrom.x, cy: nextTo.y - nextFrom.y }]]) {
          const node = childrenByLocalName(transform, name)[0], part = node && findElementRanges(xml, node.name).get('');
          if (!part) continue;
          let tag = xml.slice(part.start, part.startTagEnd);
          for (const [key, value] of Object.entries(values)) tag = setTagAttribute(tag, key, Math.round(value * 9525));
          xml = xml.slice(0, part.start) + tag + xml.slice(part.startTagEnd);
        }
      }
      operations.push({ start: range.start, end: range.end, value: xml });
    }
    if (!operations.length) continue;
    let drawingXml = parsed.xml;
    for (const operation of operations.sort((a, b) => b.start - a.start)) drawingXml = drawingXml.slice(0, operation.start) + operation.value + drawingXml.slice(operation.end);
    writeWorksheetDrawing(sheet, drawingPath, outputFiles, modifiedPaths, drawingXml);
  }
}

function writeWorksheetDrawing(sheet, drawingPath, outputFiles, modifiedPaths, drawingXml) {
  const relPath = path.join(path.dirname(sheet.path), '_rels', path.basename(sheet.path) + '.rels'), relXml = new TextDecoder().decode(outputFiles[relPath]);
  const root = parseXml(outputFiles[sheet.path], sheet.path).root, drawingId = childrenByLocalName(root, 'drawing')[0]?.attributes['r:id'];
  const write = (name, value) => { outputFiles[name] = typeof value === 'string' ? new TextEncoder().encode(value) : value; modifiedPaths.add(name); };
  if (packageReferences(outputFiles, drawingPath).some(reference => reference.name !== relPath || reference.id !== drawingId)) {
    let number = 1, copyPath;
    do { copyPath = path.join(path.dirname(drawingPath), 'drawing' + number++ + '.xml'); } while (outputFiles[copyPath]);
    const drawingRelPath = path.join(path.dirname(drawingPath), '_rels', path.basename(drawingPath) + '.rels');
    write(copyPath, drawingXml); write(path.join(path.dirname(copyPath), '_rels', path.basename(copyPath) + '.rels'), outputFiles[drawingRelPath]);
    const range = findElementRanges(relXml, 'Relationship', 'Id').get(drawingId);
    if (!range) fail('공유 차트의 워크시트 관계를 찾을 수 없습니다.');
    const tag = setTagAttribute(relXml.slice(range.start, range.startTagEnd), 'Target', path.relative(path.dirname(sheet.path), copyPath));
    write(relPath, relXml.slice(0, range.start) + tag + relXml.slice(range.startTagEnd));
    const types = parseXml(outputFiles['[Content_Types].xml'], '[Content_Types].xml'), typeRange = findElementRanges(types.xml, types.root.name).get('');
    if (localName(types.root) !== 'Types' || !typeRange || typeRange.selfClosing) fail('공유 차트의 콘텐츠 형식을 추가할 수 없습니다.');
    const prefix = types.root.name.includes(':') ? types.root.name.slice(0, types.root.name.indexOf(':') + 1) : '';
    const entry = '<' + prefix + 'Override PartName="/' + copyPath + '" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>';
    write('[Content_Types].xml', types.xml.slice(0, typeRange.closingStart) + entry + types.xml.slice(typeRange.closingStart));
  } else write(drawingPath, drawingXml);
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
  updateDimensionDrawings(model, outputFiles, modifiedPaths);
  const output = zipSync(outputFiles, { level: 0 });
  const checked = analyze(output, model.extension);
  if (!checked.writable) fail('열 너비 저장 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  for (const [name, before] of Object.entries(model.files)) {
    if (modifiedPaths.has(name)) continue;
    if (!checked.files[name] || Buffer.compare(Buffer.from(before), Buffer.from(checked.files[name])) !== 0) fail('열 너비 저장 중 비수정 패키지 항목이 달라졌습니다: ' + name);
  }
  return { bytes: output, model: checked };
}

function patchAutoFilters(model, edits, options = {}) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  if (!Array.isArray(edits) || edits.length > LIMITS.sheets) fail('자동 필터 수정 요청이 유효하지 않습니다.');
  if (!edits.length) return { bytes: model.bytes, model };
  const canonical = filter => filter === null ? null : { ...filter, columns: [...(filter?.columns || [])].sort((a, b) => a.column - b.column) };
  if (new Set(edits.map(edit => edit?.sheetIndex)).size === edits.length && edits.every(edit => {
    const sheet = Number.isInteger(edit?.sheetIndex) && model.sheets[edit.sheetIndex];
    if (!sheet) return false;
    const table = tableForFilter(sheet, edit.filter);
    if (table?.filterRangeNeedsNormalization) return false;
    return JSON.stringify(canonical(table ? table.autoFilter : sheet.autoFilter)) === JSON.stringify(canonical(edit.filter));
  })) return { bytes: model.bytes, model };
  if (edits.some(edit => edit?.filter?.columns?.some(column => column?.colorFilter))) model = ensureWorkbookStyles(model);
  model = patchFilterVisibility(model, options).model;
  const outputFiles = Object.assign({}, model.files), modifiedPaths = new Set(), seen = new Set();
  const attribute = value => xmlEscape(value).replaceAll('\r', '&#13;').replaceAll('\n', '&#10;').replaceAll('\t', '&#9;');
  for (const edit of edits) {
    const { sheetIndex, filter } = edit || {}, sheet = model.sheets[sheetIndex];
    if (!Number.isInteger(sheetIndex) || !sheet || seen.has(sheetIndex)) fail('자동 필터 시트가 유효하지 않습니다.');
    seen.add(sheetIndex);
    let replacement = '';
    if (filter !== null) {
      const { top, bottom, left, right, columns } = filter || {};
      const area = (bottom - top + 1) * (right - left + 1);
      if (![top, bottom, left, right].every(Number.isInteger) || top < 0 || bottom < top || bottom === top && !tableForFilter(sheet, filter) || left < 0 || right < left || bottom >= LIMITS.rows || right >= LIMITS.columns || !Number.isSafeInteger(area) || area > LIMITS.cells || !Array.isArray(columns) || columns.length > right - left + 1) fail('자동 필터 범위가 유효하지 않거나 크기 제한을 넘었습니다.');
      const unique = new Set();
      for (const item of columns) {
        if (!Number.isInteger(item?.column) || item.column < left || item.column > right || unique.has(item.column)) fail('자동 필터 조건이 유효하지 않습니다.');
        if (Object.hasOwn(item, 'iconFilter')) {
          if (Object.keys(item).some(key => !['column', 'iconFilter'].includes(key))) fail('아이콘 필터는 같은 열의 다른 조건과 함께 지정할 수 없습니다.');
          validateIconFilter(item.iconFilter);
        } else if (Object.hasOwn(item, 'colorFilter')) {
          if (Object.keys(item).some(key => !['column', 'colorFilter'].includes(key))) fail('색상 필터는 같은 열의 다른 조건과 함께 지정할 수 없습니다.');
          validateColorFilter(item.colorFilter);
        } else if (Object.hasOwn(item, 'dynamic')) {
          if (Object.keys(item).some(key => !['column', 'dynamic'].includes(key))) fail('동적 필터는 같은 열의 다른 조건과 함께 지정할 수 없습니다.');
          validateDynamicFilter(item.dynamic);
        } else if (Object.hasOwn(item, 'top10')) {
          if (Object.keys(item).some(key => !['column', 'top10'].includes(key))) fail('한 열의 순위 필터에 다른 종류의 조건을 함께 지정할 수 없습니다.');
          validateTopFilter(item.top10);
        } else if (item.custom) {
          if (Object.keys(item).some(key => !['column', 'custom'].includes(key))) fail('값 목록과 조건 필터를 함께 지정할 수 없습니다.');
          validateCustomFilter(item.custom); compileAutoFilterColumn(item);
        } else {
          if (!Array.isArray(item.values) || item.values.length > LIMITS.cells || item.values.some(value => typeof value !== 'string' || value.length > 32767) || typeof item.blanks !== 'boolean') fail('자동 필터 조건이 유효하지 않습니다.');
          if (Object.hasOwn(item, 'dateGroups')) validateDateGroups(item.dateGroups);
        }
        unique.add(item.column);
      }
      const serializedColumns = columns.map(item => {
        if (!item.colorFilter) return item;
        const original = new TextDecoder().decode(outputFiles[model.stylesPath]);
        const styled = ensureFilterColorStyle(original, item.colorFilter, model.colorContext);
        if (styled.xml !== original) { outputFiles[model.stylesPath] = new TextEncoder().encode(styled.xml); modifiedPaths.add(model.stylesPath); }
        return { ...item, colorFilter: { ...item.colorFilter, dxfId: styled.dxfId } };
      });
      const reference = columnLetters(left + 1) + (top + 1) + ':' + columnLetters(right + 1) + (bottom + 1);
      const columnXml = [...serializedColumns].sort((a, b) => a.column - b.column).map(item => '<filterColumn colId="' + (item.column - left) + '">' + (item.iconFilter ? '<iconFilter iconSet="' + item.iconFilter.iconSet + '" iconId="' + item.iconFilter.iconId + '"/>' : item.colorFilter ? '<colorFilter dxfId="' + item.colorFilter.dxfId + '" cellColor="' + (item.colorFilter.cellColor ? '1' : '0') + '"/>' : item.dynamic ? '<dynamicFilter type="' + item.dynamic.type + '"/>' : item.top10 ? '<top10 top="' + (item.top10.top ? '1' : '0') + '" percent="' + (item.top10.percent ? '1' : '0') + '" val="' + item.top10.value + '"/>' : item.custom ? '<customFilters' + (item.custom.and ? ' and="1"' : '') + '>' + item.custom.conditions.map(condition => '<customFilter operator="' + condition.operator + '" val="' + attribute(condition.value) + '"/>').join('') + '</customFilters>' : '<filters' + (item.blanks ? ' blank="1"' : '') + '>' + item.values.map(value => '<filter val="' + attribute(value) + '"/>').join('') + (item.dateGroups || []).map(group => '<dateGroupItem' + Object.entries(group).map(([key, value]) => ' ' + key + '="' + attribute(value) + '"').join('') + '/>').join('') + '</filters>') + '</filterColumn>').join('');
      replacement = '<autoFilter ref="' + reference + '"' + (columnXml ? '>' + columnXml + '</autoFilter>' : '/>');
      const parsed = parseXml(new TextEncoder().encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' + replacement + '</worksheet>'), sheet.path);
      const differentialNodes = model.stylesPath ? children(child(parseXml(outputFiles[model.stylesPath], model.stylesPath).root, 'dxfs') || { children: [] }, 'dxf') : [];
      readAutoFilter(child(parsed.root, 'autoFilter'), differentialNodes, model.colorContext, Boolean(tableForFilter(sheet, filter)));
    }
    const table = tableForFilter(sheet, filter);
    if (filter && !table && sheet.tables?.some(item => filter.top <= item.range.bottom && filter.bottom >= item.range.top && filter.left <= item.range.right && filter.right >= item.range.left)) fail('표 필터 범위는 머리글과 데이터 행만 포함해야 합니다.');
    if (table) {
      const xml = new TextDecoder().decode(outputFiles[table.path]), ranges = [...findElementRanges(xml, 'autoFilter').values()];
      if (ranges.length > 1) fail('표에 중복된 자동 필터 요소가 있습니다.');
      const operations = ranges.length ? [{ start: ranges[0].start, end: ranges[0].end, replacement }] : replacement ? (() => {
        const columns = findElementRanges(xml, 'tableColumns').get('');
        if (!columns) fail('표 열 목록 앞에 자동 필터를 안전하게 추가할 수 없습니다.');
        return [{ start: columns.start, end: columns.start, replacement }];
      })() : [];
      if (!operations.length) continue;
      let updated = xml;
      for (const operation of operations) updated = updated.slice(0, operation.start) + operation.replacement + updated.slice(operation.end);
      if (updated !== xml) { outputFiles[table.path] = new TextEncoder().encode(updated); modifiedPaths.add(table.path); }
      continue;
    }
    const ranges = [...findElementRanges(sheet.xml, 'autoFilter').values()];
    if (ranges.length > 1) fail('중복된 자동 필터 요소를 안전하게 수정할 수 없습니다.');
    const operations = [];
    if (ranges.length) operations.push({ start: ranges[0].start, end: ranges[0].end, replacement });
    else if (replacement) {
      const sheetData = findElementRanges(sheet.xml, 'sheetData').get('');
      if (!sheetData || sheetData.selfClosing) fail('자동 필터를 워크시트에 안전하게 추가할 수 없습니다.');
      operations.push({ start: sheetData.end, end: sheetData.end, replacement });
    }
    if (!operations.length) continue;
    let xml = sheet.xml;
    for (const operation of operations) xml = xml.slice(0, operation.start) + operation.replacement + xml.slice(operation.end);
    outputFiles[sheet.path] = new TextEncoder().encode(xml);
    modifiedPaths.add(sheet.path);
  }
  if (!modifiedPaths.size) return { bytes: model.bytes, model };
  const bytes = zipSync(outputFiles, { level: 0 }), checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('자동 필터 저장 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  for (const [name, before] of Object.entries(model.files)) {
    if (modifiedPaths.has(name)) continue;
    if (!checked.files[name] || Buffer.compare(Buffer.from(before), Buffer.from(checked.files[name])) !== 0) fail('자동 필터 저장 중 비수정 패키지 항목이 달라졌습니다: ' + name);
  }
  return patchFilterVisibility(checked, options);
}

function dataValidationReference(range) {
  const first = columnLetters(range.left + 1) + (range.top + 1);
  const last = columnLetters(range.right + 1) + (range.bottom + 1);
  return first === last ? first : first + ':' + last;
}

function dataValidationFormula(rule, index) {
  if (rule.type === 'list') {
    if (Array.isArray(rule.values)) return '"' + rule.values.map(value => String(value).replaceAll('"', '""')).join(',') + '"';
    if (rule.source) {
      const source = rule.source, reference = dataValidationReference(source);
      const sheetName = source.sheetName;
      return (sheetName ? (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheetName) ? sheetName : "'" + sheetName.replaceAll("'", "''") + "'") + '!' : '') + reference;
    }
    if (typeof rule.sourceFormula === 'string' && rule.sourceFormula) return '=' + rule.sourceFormula.replace(/^=/, '');
    if (rule.sourceFormula === '#REF!') return '#REF!';
    fail('데이터 유효성 목록 원본을 저장할 수 없습니다.');
  }
  const formula = rule.formulas?.[index];
  const value = formula === null || formula === undefined ? rule.values?.[index] : formula;
  if (typeof value !== 'string' && typeof value !== 'number' || typeof value === 'number' && !Number.isFinite(value) || typeof value === 'string' && (!value || value.length > 255)) fail('데이터 유효성 기준값을 저장할 수 없습니다.');
  return String(value).replace(/^=/, '');
}

function serializeDataValidationRule(rule, prefix = '') {
  if (!rule || !['list', 'whole', 'decimal', 'date', 'time', 'textLength'].includes(rule.type) || !Array.isArray(rule.ranges) || !rule.ranges.length || rule.ranges.length > 64) fail('데이터 유효성 규칙이 유효하지 않습니다.');
  const needsSecond = rule.type !== 'list' && ['between', 'notBetween'].includes(rule.operator);
  const first = dataValidationFormula(rule, 0), second = needsSecond ? dataValidationFormula(rule, 1) : null;
  if (rule.type !== 'list' && !['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'].includes(rule.operator)) fail('데이터 유효성 비교 연산자가 유효하지 않습니다.');
  const attrs = {
    type: rule.type,
    errorStyle: rule.errorStyle || 'stop',
    ...(rule.type === 'list' ? {} : { operator: rule.operator }),
    allowBlank: rule.allowBlank ? '1' : '0',
    ...(rule.type === 'list' ? { showDropDown: rule.showDropDown ? '1' : '0' } : {}),
    showInputMessage: rule.showInputMessage === false ? '0' : '1',
    showErrorMessage: rule.showErrorMessage === false ? '0' : '1',
    ...(rule.errorTitle ? { errorTitle: rule.errorTitle } : {}),
    ...(rule.error ? { error: rule.error } : {}),
    ...(rule.promptTitle ? { promptTitle: rule.promptTitle } : {}),
    ...(rule.prompt ? { prompt: rule.prompt } : {}),
    sqref: rule.ranges.map(range => {
      if (![range?.top, range?.bottom, range?.left, range?.right].every(Number.isInteger) || range.top < 0 || range.left < 0 || range.bottom < range.top || range.right < range.left || range.bottom >= LIMITS.rows || range.right >= LIMITS.columns) fail('데이터 유효성 적용 범위가 유효하지 않습니다.');
      return dataValidationReference(range);
    }).join(' ')
  };
  if (!['stop', 'warning', 'information'].includes(attrs.errorStyle) || ['errorTitle', 'promptTitle'].some(name => String(attrs[name] || '').length > 32) || ['error', 'prompt'].some(name => String(attrs[name] || '').length > 255)) fail('데이터 유효성 알림 문구가 제한을 벗어났습니다.');
  const attribute = value => xmlEscape(String(value)).replaceAll('"', '&quot;');
  const name = prefix + 'dataValidation';
  return '<' + name + Object.entries(attrs).map(([key, value]) => ' ' + key + '="' + attribute(value) + '"').join('') + '><' + prefix + 'formula1>' + xmlEscape(first) + '</' + prefix + 'formula1>' + (second === null ? '' : '<' + prefix + 'formula2>' + xmlEscape(second) + '</' + prefix + 'formula2>') + '</' + name + '>';
}

function patchDataValidations(model, sheetIndex, validations) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  const sheet = model.sheets[sheetIndex];
  if (!sheet || !Array.isArray(validations) || validations.length > 256) fail('데이터 유효성 수정 요청이 유효하지 않습니다.');
  if (JSON.stringify(sheet.dataValidations || []) === JSON.stringify(validations)) return { bytes: model.bytes, model };
  const xml = sheet.xml, containers = [...findElementRanges(xml, 'dataValidations').values()];
  if (containers.length > 1) fail('중복된 데이터 유효성 컨테이너를 수정할 수 없습니다.');
  const prefix = /^\s*<([A-Za-z_][\w.-]*:)?worksheet\b/.exec(xml)?.[1] || '';
  const replacement = validations.length ? '<' + prefix + 'dataValidations count="' + validations.length + '">' + validations.map(rule => serializeDataValidationRule(rule, prefix)).join('') + '</' + prefix + 'dataValidations>' : '';
  const outputFiles = Object.assign({}, model.files);
  let updated;
  if (containers.length) {
    const range = containers[0];
    updated = xml.slice(0, range.start) + replacement + xml.slice(range.end);
  } else if (replacement) {
    const laterNames = ['hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];
    const later = laterNames.flatMap(name => [...findElementRanges(xml, name).values()].map(range => range.start));
    const conditional = [...findElementRanges(xml, 'conditionalFormatting', undefined, true).values()].map(range => range.end);
    const sheetData = findElementRanges(xml, 'sheetData').get('');
    const at = conditional.length ? Math.max(...conditional) : later.length ? Math.min(...later) : sheetData?.end;
    if (!Number.isInteger(at)) fail('데이터 유효성 규칙을 워크시트에 안전하게 추가할 수 없습니다.');
    updated = xml.slice(0, at) + replacement + xml.slice(at);
  } else return { bytes: model.bytes, model };
  if (updated === xml) return { bytes: model.bytes, model };
  const bytes = zipSync({ ...outputFiles, [sheet.path]: new TextEncoder().encode(updated) }, { level: 0 });
  const checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('데이터 유효성 저장 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  for (const [name, before] of Object.entries(model.files)) if (name !== sheet.path && Buffer.compare(Buffer.from(before), Buffer.from(checked.files[name])) !== 0) fail('데이터 유효성 저장 중 비수정 패키지 항목이 달라졌습니다: ' + name);
  return { bytes, model: checked };
}

function patchConditionalFormatRule(model, sheetIndex, range, buildRule) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  const sheet = model.sheets[sheetIndex];
  if (!sheet || !range || ![range.top, range.bottom, range.left, range.right].every(Number.isInteger) || range.top < 0 || range.left < 0 || range.bottom < range.top || range.right < range.left || range.bottom >= LIMITS.rows || range.right >= LIMITS.columns) fail('조건부 서식 범위가 유효하지 않습니다.');
  const existing = sheet.conditionalFormats || [];
  if (existing.length >= 256) fail('워크시트에는 256개 이하의 조건부 서식 규칙을 사용할 수 있습니다.');

  let worksheetXml = sheet.xml;
  const ruleRanges = [...findElementRanges(worksheetXml, 'cfRule', undefined, true).values()];
  if (ruleRanges.length !== existing.length) fail('조건부 서식 우선 순위를 안전하게 변경할 수 없습니다.');
  const nextPriority = new Map(existing.map((rule, index) => [String(rule.priority), index + 2]));
  for (const ruleRange of ruleRanges.reverse()) {
    const opening = worksheetXml.slice(ruleRange.start, ruleRange.startTagEnd), priority = parseTagAttributes(opening).priority;
    const shiftedPriority = nextPriority.get(priority);
    if (!shiftedPriority) fail('조건부 서식 우선 순위를 안전하게 변경할 수 없습니다.');
    worksheetXml = worksheetXml.slice(0, ruleRange.start) + setTagAttribute(opening, 'priority', shiftedPriority) + worksheetXml.slice(ruleRange.startTagEnd);
  }

  const prefix = /^\s*<([A-Za-z_][\w.-]*:)?worksheet\b/.exec(worksheetXml)?.[1] || '';
  const cellReference = (row, column) => columnLetters(column + 1) + (row + 1);
  const first = cellReference(range.top, range.left), last = cellReference(range.bottom, range.right);
  const reference = first === last ? first : first + ':' + last;
  const conditionalName = prefix + 'conditionalFormatting';
  const formatting = '<' + conditionalName + ' sqref="' + reference + '">' + buildRule(prefix) + '</' + conditionalName + '>';
  const priorNames = ['sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting'];
  const laterNames = ['dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];
  const prior = priorNames.flatMap(name => [...findElementRanges(worksheetXml, name, undefined, name === 'conditionalFormatting').values()].map(item => item.end));
  const later = laterNames.flatMap(name => [...findElementRanges(worksheetXml, name, undefined, true).values()].map(item => item.start));
  const at = prior.length ? Math.max(...prior) : -1, next = later.length ? Math.min(...later) : worksheetXml.indexOf('</' + prefix + 'worksheet>');
  if (at < 0 || next < 0 || at > next) fail('조건부 서식을 워크시트에 안전하게 추가할 수 없습니다.');
  worksheetXml = worksheetXml.slice(0, at) + formatting + worksheetXml.slice(at);

  const bytes = zipSync({ ...model.files, [sheet.path]: new TextEncoder().encode(worksheetXml) }, { level: 0 });
  const checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('조건부 서식 저장 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  for (const [name, before] of Object.entries(model.files)) if (name !== sheet.path && (!checked.files[name] || Buffer.compare(Buffer.from(before), Buffer.from(checked.files[name])) !== 0)) fail('조건부 서식 저장 중 비수정 패키지 항목이 달라졌습니다: ' + name);
  return { bytes, model: checked };
}

function patchDataBarConditionalFormat(model, sheetIndex, range, color) {
  if (typeof color !== 'string' || !/^#[0-9A-F]{6}$/i.test(color)) fail('조건부 서식 색상이 유효하지 않습니다.');
  return patchConditionalFormatRule(model, sheetIndex, range, prefix => {
    const rule = prefix + 'cfRule', dataBar = prefix + 'dataBar', stop = prefix + 'cfvo', colorTag = prefix + 'color';
    return '<' + rule + ' type="dataBar" priority="1"><' + dataBar + ' minLength="0" maxLength="100" showValue="1" gradient="1"><' + stop + ' type="min"/><' + stop + ' type="max"/><' + colorTag + ' rgb="FF' + color.slice(1).toUpperCase() + '"/></' + dataBar + '></' + rule + '>';
  });
}

function patchCellValueConditionalFormat(model, sheetIndex, range, operator, values, color = '#F4CCCC') {
  const operators = ['equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual', 'between', 'notBetween'];
  const expected = operator === 'between' || operator === 'notBetween' ? 2 : 1;
  if (!operators.includes(operator) || !Array.isArray(values) || values.length !== expected || values.some(value => typeof value !== 'number' || !Number.isFinite(value)) || typeof color !== 'string' || !/^#[0-9A-F]{6}$/i.test(color)) fail('숫자 조건부 서식 설정이 유효하지 않습니다.');
  model = ensureWorkbookStyles(model);
  const stylesXml = new TextDecoder().decode(model.files[model.stylesPath]), styled = ensureFilterColorStyle(stylesXml, { cellColor: true, color: color.toUpperCase() }, model.colorContext);
  if (styled.xml !== stylesXml) {
    const files = { ...model.files, [model.stylesPath]: new TextEncoder().encode(styled.xml) };
    model = analyze(zipSync(files, { level: 0 }), model.extension);
    if (!model.writable) fail('조건부 서식 색상을 추가한 뒤 통합 문서를 안전하게 다시 읽을 수 없습니다.');
  }
  return patchConditionalFormatRule(model, sheetIndex, range, prefix => '<' + prefix + 'cfRule type="cellIs" operator="' + operator + '" dxfId="' + styled.dxfId + '" priority="1">' + values.map(value => '<' + prefix + 'formula>' + xmlEscape(String(value)) + '</' + prefix + 'formula>').join('') + '</' + prefix + 'cfRule>');
}

function patchColorScaleConditionalFormat(model, sheetIndex, range, colors) {
  if (!Array.isArray(colors) || ![2, 3].includes(colors.length) || colors.some(color => typeof color !== 'string' || !/^#[0-9A-F]{6}$/i.test(color))) fail('조건부 서식 색조 색상 구성이 유효하지 않습니다.');
  return patchConditionalFormatRule(model, sheetIndex, range, prefix => {
    const rule = prefix + 'cfRule', scale = prefix + 'colorScale', stop = prefix + 'cfvo', colorTag = prefix + 'color';
    const stops = colors.length === 2 ? ['type="min"', 'type="max"'] : ['type="min"', 'type="percentile" val="50"', 'type="max"'];
    const stopXml = stops.map(attributes => '<' + stop + ' ' + attributes + '/>').join('');
    const colorXml = colors.map(color => '<' + colorTag + ' rgb="FF' + color.slice(1).toUpperCase() + '"/>').join('');
    return '<' + rule + ' type="colorScale" priority="1"><' + scale + '>' + stopXml + colorXml + '</' + scale + '></' + rule + '>';
  });
}

function patchDuplicateValuesConditionalFormat(model, sheetIndex, range, color = '#F4CCCC') {
  if (typeof color !== 'string' || !/^#[0-9A-F]{6}$/i.test(color)) fail('조건부 서식 색상이 유효하지 않습니다.');
  if (!range || ![range.top, range.bottom, range.left, range.right].every(Number.isInteger) || range.top < 0 || range.left < 0 || range.bottom < range.top || range.right < range.left || range.bottom >= LIMITS.rows || range.right >= LIMITS.columns || (range.bottom - range.top + 1) * (range.right - range.left + 1) > 100000) fail('중복 값 강조 범위는 100,000셀 이하의 유효한 범위여야 합니다.');
  model = ensureWorkbookStyles(model);
  const stylesXml = new TextDecoder().decode(model.files[model.stylesPath]), styled = ensureFilterColorStyle(stylesXml, { cellColor: true, color: color.toUpperCase() }, model.colorContext);
  if (styled.xml !== stylesXml) {
    const files = { ...model.files, [model.stylesPath]: new TextEncoder().encode(styled.xml) };
    model = analyze(zipSync(files, { level: 0 }), model.extension);
    if (!model.writable) fail('중복 값 서식을 추가한 뒤 통합 문서를 안전하게 다시 읽을 수 없습니다.');
  }
  const cellReference = (row, column, absolute = false) => (absolute ? '$' : '') + columnLetters(column + 1) + (absolute ? '$' : '') + (row + 1);
  const first = cellReference(range.top, range.left), fixedFirst = cellReference(range.top, range.left, true), fixedLast = cellReference(range.bottom, range.right, true);
  const formula = 'AND(' + first + '<>"",COUNTIF(' + fixedFirst + ':' + fixedLast + ',' + first + ')>1)';
  return patchConditionalFormatRule(model, sheetIndex, range, prefix => '<' + prefix + 'cfRule type="expression" dxfId="' + styled.dxfId + '" priority="1"><' + prefix + 'formula>' + xmlEscape(formula) + '</' + prefix + 'formula></' + prefix + 'cfRule>');
}

function standardIconSetThresholds(count) {
  return Array.from({ length: count }, (_item, index) => ({ type: 'percent', value: Math.round(index * 100 / count), gte: true }));
}

function standardIconSetThresholdXml(tagName, count) {
  return standardIconSetThresholds(count).map(stop => '<' + tagName + ' type="percent" val="' + stop.value + '"/>').join('');
}

function patchIconSetConditionalFormat(model, sheetIndex, range, name) {
  if (typeof name !== 'string' || !Object.hasOwn(ICON_SET_LENGTHS, name)) fail('지원하지 않는 아이콘 집합입니다.');
  const iconCount = ICON_SET_LENGTHS[name];
  return patchConditionalFormatRule(model, sheetIndex, range, prefix => {
    const rule = prefix + 'cfRule', iconSet = prefix + 'iconSet', stop = prefix + 'cfvo';
    const thresholds = standardIconSetThresholdXml(stop, iconCount);
    return '<' + rule + ' type="iconSet" priority="1"><' + iconSet + ' iconSet="' + name + '" showValue="1" percent="1" reverse="0">' + thresholds + '</' + iconSet + '></' + rule + '>';
  });
}

function patchClearConditionalFormats(model, sheetIndex, range) {
  if (!model.writable) fail(model.readOnlyReason || '조건부 서식을 수정할 수 없는 문서입니다.');
  const sheet = model.sheets[sheetIndex];
  if (!sheet || !range || ![range.top, range.bottom, range.left, range.right].every(Number.isInteger) || range.top < 0 || range.left < 0 || range.bottom < range.top || range.right < range.left || range.bottom >= LIMITS.rows || range.right >= LIMITS.columns) fail('조건부 서식 범위가 유효하지 않습니다.');
  const overlaps = source => source.top <= range.bottom && source.bottom >= range.top && source.left <= range.right && source.right >= range.left;
  if (!(sheet.conditionalFormats || []).some(rule => rule.ranges?.some(overlaps))) return { bytes: model.bytes, model };

  let worksheetXml = sheet.xml;
  const blocks = [...findElementRanges(worksheetXml, 'conditionalFormatting', undefined, true).values()];
  if (!blocks.length) fail('조건부 서식 범위를 워크시트에서 찾을 수 없습니다.');
  let overlappedBlock = false;
  const referenceFor = source => {
    const first = columnLetters(source.left + 1) + (source.top + 1), last = columnLetters(source.right + 1) + (source.bottom + 1);
    return first === last ? first : first + ':' + last;
  };
  for (const block of blocks.reverse()) {
    const opening = worksheetXml.slice(block.start, block.startTagEnd), attributes = parseTagAttributes(opening);
    const references = String(attributes.sqref || '').trim().split(/\s+/).filter(Boolean);
    if (!references.length || references.length > 64) fail('조건부 서식 참조가 너무 많아 안전하게 변경할 수 없습니다.');
    const originalRanges = references.map(reference => {
      const parsed = parseRangeReference(reference, null, LIMITS.rows * LIMITS.columns);
      if (!parsed || parsed.sheetName) fail('조건부 서식 범위를 안전하게 해석할 수 없습니다.');
      return parsed;
    });
    if (!originalRanges.some(overlaps)) continue;
    overlappedBlock = true;
    const remaining = originalRanges.flatMap(source => subtractRange(source, range));
    if (remaining.length > 64) fail('선택한 셀을 지우면 조건부 서식 범위가 64개를 초과합니다.');
    if (!remaining.length) {
      worksheetXml = worksheetXml.slice(0, block.start) + worksheetXml.slice(block.end);
      continue;
    }
    const updatedOpening = setTagAttribute(opening, 'sqref', remaining.map(referenceFor).join(' '));
    worksheetXml = worksheetXml.slice(0, block.start) + updatedOpening + worksheetXml.slice(block.startTagEnd);
  }
  if (!overlappedBlock) fail('조건부 서식 범위를 워크시트에서 찾을 수 없습니다.');
  if (worksheetXml === sheet.xml) fail('조건부 서식 범위를 안전하게 변경할 수 없습니다.');
  const outputFiles = Object.assign({}, model.files);
  const bytes = zipSync({ ...outputFiles, [sheet.path]: new TextEncoder().encode(worksheetXml) }, { level: 0 });
  const checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('조건부 서식 지우기 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  for (const [name, before] of Object.entries(model.files)) if (name !== sheet.path && (!checked.files[name] || Buffer.compare(Buffer.from(before), Buffer.from(checked.files[name])) !== 0)) fail('조건부 서식 지우기 중 비수정 패키지 항목이 달라졌습니다: ' + name);
  return { bytes, model: checked };
}

function patchDeleteConditionalFormatRule(model, sheetIndex, priority) {
  if (!model.writable) fail(model.readOnlyReason || '조건부 서식을 수정할 수 없는 문서입니다.');
  const sheet = model.sheets[sheetIndex], rules = sheet?.conditionalFormats || [];
  if (!sheet || !Number.isInteger(priority) || priority < 1 || priority > 100000) fail('조건부 서식 규칙 선택이 유효하지 않습니다.');
  if (!rules.some(rule => rule.priority === priority)) fail('선택한 조건부 서식 규칙을 찾을 수 없습니다.');
  let worksheetXml = sheet.xml;
  const ruleRanges = [...findElementRanges(worksheetXml, 'cfRule', undefined, true).values()];
  if (ruleRanges.length !== rules.length) fail('조건부 서식 규칙을 안전하게 제거할 수 없습니다.');
  const target = ruleRanges.find(range => parseTagAttributes(worksheetXml.slice(range.start, range.startTagEnd)).priority === String(priority));
  const blocks = [...findElementRanges(worksheetXml, 'conditionalFormatting', undefined, true).values()];
  const block = blocks.find(range => target && target.start >= range.start && target.end <= range.end);
  if (!target || !block) fail('선택한 조건부 서식 규칙을 찾을 수 없습니다.');
  const blockXml = worksheetXml.slice(block.start, block.end), localRanges = [...findElementRanges(blockXml, 'cfRule', undefined, true).values()];
  const localTarget = localRanges.find(range => parseTagAttributes(blockXml.slice(range.start, range.startTagEnd)).priority === String(priority));
  if (!localTarget || localRanges.length < 1) fail('조건부 서식 규칙을 안전하게 제거할 수 없습니다.');
  if (localRanges.length === 1) worksheetXml = worksheetXml.slice(0, block.start) + worksheetXml.slice(block.end);
  else {
    const updatedBlock = blockXml.slice(0, localTarget.start) + blockXml.slice(localTarget.end);
    worksheetXml = worksheetXml.slice(0, block.start) + updatedBlock + worksheetXml.slice(block.end);
  }
  const outputFiles = Object.assign({}, model.files);
  const bytes = zipSync({ ...outputFiles, [sheet.path]: new TextEncoder().encode(worksheetXml) }, { level: 0 });
  const checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('조건부 서식 규칙 제거 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  for (const [name, before] of Object.entries(model.files)) if (name !== sheet.path && (!checked.files[name] || Buffer.compare(Buffer.from(before), Buffer.from(checked.files[name])) !== 0)) fail('조건부 서식 규칙 제거 중 비수정 패키지 항목이 달라졌습니다: ' + name);
  return { bytes, model: checked };
}

function patchMoveConditionalFormatRule(model, sheetIndex, priority, direction) {
  if (!model.writable) fail(model.readOnlyReason || '조건부 서식을 수정할 수 없는 문서입니다.');
  const sheet = model.sheets[sheetIndex], rules = sheet?.conditionalFormats || [];
  if (!sheet || !Number.isInteger(priority) || priority < 1 || priority > 100000 || !['up', 'down'].includes(direction)) fail('조건부 서식 규칙 선택이 유효하지 않습니다.');
  const selectedIndex = rules.findIndex(rule => rule.priority === priority);
  if (selectedIndex < 0) fail('선택한 조건부 서식 규칙을 찾을 수 없습니다.');
  const neighbor = rules[selectedIndex + (direction === 'up' ? -1 : 1)];
  if (!neighbor) return { bytes: model.bytes, model };

  let worksheetXml = sheet.xml;
  const ruleRanges = [...findElementRanges(worksheetXml, 'cfRule', undefined, true).values()];
  if (ruleRanges.length !== rules.length) fail('조건부 서식 우선 순위를 안전하게 변경할 수 없습니다.');
  const priorities = new Map([[priority, neighbor.priority], [neighbor.priority, priority]]);
  const patches = [];
  for (const range of ruleRanges) {
    const opening = worksheetXml.slice(range.start, range.startTagEnd), currentPriority = parseTagAttributes(opening).priority;
    if (!priorities.has(Number(currentPriority))) continue;
    patches.push({ start: range.start, end: range.startTagEnd, replacement: setTagAttribute(opening, 'priority', priorities.get(Number(currentPriority))) });
  }
  if (patches.length !== 2) fail('조건부 서식 우선 순위를 안전하게 변경할 수 없습니다.');
  for (const patch of patches.sort((left, right) => right.start - left.start)) worksheetXml = worksheetXml.slice(0, patch.start) + patch.replacement + worksheetXml.slice(patch.end);
  const outputFiles = Object.assign({}, model.files);
  const bytes = zipSync({ ...outputFiles, [sheet.path]: new TextEncoder().encode(worksheetXml) }, { level: 0 });
  const checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('조건부 서식 우선 순위 변경 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  for (const [name, before] of Object.entries(model.files)) if (name !== sheet.path && (!checked.files[name] || Buffer.compare(Buffer.from(before), Buffer.from(checked.files[name])) !== 0)) fail('조건부 서식 우선 순위 변경 중 비수정 패키지 항목이 달라졌습니다: ' + name);
  return { bytes, model: checked };
}

function replaceDifferentialColor(dxfXml, property, color) {
  const colorXml = color ? '<' + (property === 'fontColor' ? 'color' : 'fgColor') + ' rgb="FF' + color.slice(1) + '"/>' : null;
  if (property === 'fontColor') {
    const font = findElementRanges(dxfXml, 'font').get('');
    if (!font) return colorXml ? insertDifferentialComponent(dxfXml, 'font', '<font>' + colorXml + '</font>') : dxfXml;
    let fontXml = dxfXml.slice(font.start, font.end);
    const colorRange = findElementRanges(fontXml, 'color').get('');
    if (colorRange) fontXml = fontXml.slice(0, colorRange.start) + (colorXml || '') + fontXml.slice(colorRange.end);
    else if (colorXml) fontXml = insertDifferentialComponent(fontXml, 'color', colorXml);
    return dxfXml.slice(0, font.start) + fontXml + dxfXml.slice(font.end);
  }
  if (property !== 'fillColor') fail('조건부 서식 색상 항목이 유효하지 않습니다.');
  const fill = findElementRanges(dxfXml, 'fill').get('');
  if (!fill) return colorXml ? insertDifferentialComponent(dxfXml, 'fill', '<fill><patternFill patternType="solid">' + colorXml + '</patternFill></fill>') : dxfXml;
  if (!colorXml) return dxfXml.slice(0, fill.start) + dxfXml.slice(fill.end);
  let fillXml = dxfXml.slice(fill.start, fill.end);
  const pattern = findElementRanges(fillXml, 'patternFill').get('');
  if (!pattern) fail('조건부 서식 채우기 색상을 안전하게 변경할 수 없습니다.');
  let patternXml = fillXml.slice(pattern.start, pattern.end);
  const foreground = findElementRanges(patternXml, 'fgColor').get('');
  if (foreground) patternXml = patternXml.slice(0, foreground.start) + colorXml + patternXml.slice(foreground.end);
  else patternXml = insertDifferentialComponent(patternXml, 'fgColor', colorXml);
  fillXml = fillXml.slice(0, pattern.start) + patternXml + fillXml.slice(pattern.end);
  return dxfXml.slice(0, fill.start) + fillXml + dxfXml.slice(fill.end);
}

function insertDifferentialComponent(xml, component, componentXml) {
  const rootName = /^<\s*([A-Za-z_][\w.:-]*)/.exec(xml)?.[1];
  if (!rootName) fail('조건부 서식 스타일을 안전하게 변경할 수 없습니다.');
  const root = findElementRanges(xml, rootName).get('');
  if (!root) fail('조건부 서식 스타일을 안전하게 변경할 수 없습니다.');
  const insertionOrder = component === 'font' ? ['fill', 'alignment', 'border']
    : component === 'fill' ? ['alignment', 'border'] : [];
  const candidates = insertionOrder.flatMap(name => [...findElementRanges(xml, name, undefined, true).values()].map(range => range.start));
  const position = candidates.length ? Math.min(...candidates) : root.closingStart;
  if (root.selfClosing) {
    const opening = xml.slice(root.start, root.startTagEnd).replace(/\/\s*>$/, '>');
    const close = '</' + rootName + '>';
    return xml.slice(0, root.start) + opening + componentXml + close + xml.slice(root.end);
  }
  return xml.slice(0, position) + componentXml + xml.slice(position);
}

function patchDifferentialFormatColor(stylesXml, dxfId, property, color) {
  const root = parseXml(new TextEncoder().encode(stylesXml), 'conditional-format-styles').root;
  const dxfs = child(root, 'dxfs'), dxfNodes = children(dxfs || { children: [] }, 'dxf');
  const ranges = [...findElementRanges(stylesXml, 'dxf', undefined, true).values()];
  if (!dxfs || !/^(?:0|[1-9]\d*)$/.test(dxfs.attributes.count || '') || Number(dxfs.attributes.count) !== dxfNodes.length || ranges.length !== dxfNodes.length || !Number.isInteger(dxfId) || dxfId < 0 || dxfId >= dxfNodes.length) fail('조건부 서식 차등 스타일을 찾을 수 없습니다.');
  if (dxfNodes.length >= 65536) fail('조건부 서식 차등 스타일 개수 제한을 초과했습니다.');
  const sourceRange = ranges[dxfId], sourceXml = stylesXml.slice(sourceRange.start, sourceRange.end);
  const updatedDxf = replaceDifferentialColor(sourceXml, property, color);
  const identical = ranges.findIndex(range => stylesXml.slice(range.start, range.end) === updatedDxf);
  if (identical >= 0) return { xml: stylesXml, dxfId: identical };
  const nextDxfId = dxfNodes.length, opening = setTagAttribute(stylesXml.slice(findElementRanges(stylesXml, 'dxfs').get('').start, findElementRanges(stylesXml, 'dxfs').get('').startTagEnd), 'count', nextDxfId + 1);
  const dxfsRange = findElementRanges(stylesXml, 'dxfs').get('');
  if (!dxfsRange) fail('조건부 서식 차등 스타일 목록을 수정할 수 없습니다.');
  const updatedXml = stylesXml.slice(0, dxfsRange.start) + opening + stylesXml.slice(dxfsRange.startTagEnd, dxfsRange.closingStart) + updatedDxf + stylesXml.slice(dxfsRange.closingStart);
  return { xml: updatedXml, dxfId: nextDxfId };
}

function patchEditConditionalFormatRule(model, sheetIndex, priority, edit) {
  if (!model.writable) fail(model.readOnlyReason || '조건부 서식을 수정할 수 없는 문서입니다.');
  const sheet = model.sheets[sheetIndex], rules = sheet?.conditionalFormats || [], rule = rules.find(item => item.priority === priority);
  if (!sheet || !Number.isInteger(priority) || priority < 1 || priority > 100000 || !rule) fail('선택한 조건부 서식 규칙이 유효하지 않거나 이미 없어졌습니다.');
  if (rule.type === 'iconSet' && rule.iconSet.customIcons && !['icon-set-display', 'icon-set-mapping', 'icon-set-threshold'].includes(edit?.kind)) fail('사용자 지정 아이콘 집합의 스타일과 기준은 현재 편집할 수 없습니다.');
  const ruleRanges = [...findElementRanges(sheet.xml, 'cfRule', undefined, true).values()];
  if (ruleRanges.length !== rules.length) fail('조건부 서식을 워크시트에서 안전하게 변경할 수 없습니다.');
  const target = ruleRanges.find(range => parseTagAttributes(sheet.xml.slice(range.start, range.startTagEnd)).priority === String(priority));
  if (!target) fail('선택한 조건부 서식 규칙을 찾을 수 없습니다.');
  let updatedRuleXml = sheet.xml.slice(target.start, target.end);
  const ruleOpeningEnd = target.startTagEnd - target.start;
  let updatedStylesXml = null;

  if (edit?.kind === 'condition') {
    if (rule.type === 'top10') {
      const top10 = edit.top10;
      if (!top10 || typeof top10 !== 'object' || Array.isArray(top10) || Object.keys(top10).length !== 3 || Object.keys(top10).some(key => !['rank', 'percent', 'bottom'].includes(key)) || !Number.isInteger(top10.rank) || typeof top10.percent !== 'boolean' || typeof top10.bottom !== 'boolean' || top10.rank < 1 || top10.rank > (top10.percent ? 100 : 1000)) fail('상위·하위 순위 설정이 유효하지 않습니다.');
      if (top10.rank === rule.top10.rank && top10.percent === rule.top10.percent && top10.bottom === rule.top10.bottom) return { bytes: model.bytes, model };
      let opening = updatedRuleXml.slice(0, ruleOpeningEnd);
      opening = setTagAttribute(opening, 'rank', top10.rank);
      opening = setTagAttribute(opening, 'percent', top10.percent ? '1' : '0');
      opening = setTagAttribute(opening, 'bottom', top10.bottom ? '1' : '0');
      updatedRuleXml = opening + updatedRuleXml.slice(ruleOpeningEnd);
    } else if (rule.type === 'cellIs') {
      const operator = edit.operator, expected = operator === 'between' || operator === 'notBetween' ? 2 : 1;
      if (!['equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual', 'between', 'notBetween'].includes(operator) || !Array.isArray(edit.values) || edit.values.length !== expected) fail('조건부 서식 규칙 선택이 유효하지 않습니다.');
      const values = edit.values.map(value => {
        const text = String(value ?? '').trim();
        if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(text) || !Number.isFinite(Number(text))) fail('조건부 서식 비교값이 유효하지 않습니다.');
        return String(Number(text));
      });
      if (operator === rule.operator && values.length === rule.values.length && values.every((value, index) => Number(value) === rule.values[index])) return { bytes: model.bytes, model };
      const opening = updatedRuleXml.slice(0, ruleOpeningEnd);
      updatedRuleXml = setTagAttribute(opening, 'operator', operator) + updatedRuleXml.slice(ruleOpeningEnd);
      let formulaRanges = [...findElementRanges(updatedRuleXml, 'formula', undefined, true).values()];
      if (formulaRanges.length === 1 && expected === 2) {
        const lastFormula = formulaRanges[0], closing = updatedRuleXml.lastIndexOf('</cfRule>'), formulaTag = updatedRuleXml.slice(lastFormula.start, lastFormula.startTagEnd), name = /^<\s*([A-Za-z_][\w.:-]*)/.exec(formulaTag)?.[1];
        if (!name || closing < lastFormula.end) fail('조건부 서식 규칙을 안전하게 변경할 수 없습니다.');
        const trailingWhitespace = /\s*$/.exec(updatedRuleXml.slice(lastFormula.end, closing))[0], insertionPoint = closing - trailingWhitespace.length;
        updatedRuleXml = updatedRuleXml.slice(0, insertionPoint) + '<' + name + '></' + name + '>' + updatedRuleXml.slice(insertionPoint);
        formulaRanges = [...findElementRanges(updatedRuleXml, 'formula', undefined, true).values()];
      } else if (formulaRanges.length === 2 && expected === 1) {
        const lastFormula = formulaRanges[1];
        updatedRuleXml = updatedRuleXml.slice(0, lastFormula.start) + updatedRuleXml.slice(lastFormula.end);
        formulaRanges = [...findElementRanges(updatedRuleXml, 'formula', undefined, true).values()];
      }
      if (formulaRanges.length !== expected) fail('조건부 서식 규칙을 안전하게 변경할 수 없습니다.');
      const formulaPatches = formulaRanges.map((range, index) => ({ range, index })).sort((left, right) => right.range.start - left.range.start);
      for (const { index, range } of formulaPatches) {
        const formulaTag = updatedRuleXml.slice(range.start, range.startTagEnd), name = /^<\s*([A-Za-z_][\w.:-]*)/.exec(formulaTag)?.[1];
        if (!name || updatedRuleXml.slice(range.startTagEnd, range.closingStart).includes('<!--')) fail('조건부 서식 규칙을 안전하게 변경할 수 없습니다.');
        const replacement = range.selfClosing ? '<' + name + '>' + xmlEscape(values[index]) + '</' + name + '>' : updatedRuleXml.slice(range.start, range.startTagEnd) + xmlEscape(values[index]) + updatedRuleXml.slice(range.closingStart, range.end);
        updatedRuleXml = updatedRuleXml.slice(0, range.start) + replacement + updatedRuleXml.slice(range.end);
      }
    } else if (rule.type === 'expression') {
      const formula = String(edit.formula ?? '').trim().replace(/^=/, '');
      if (!formula || formula.length > 256 || /[!\[\]]/.test(formula)) fail('조건부 서식 수식이 유효하지 않거나 지원되지 않습니다.');
      const sample = evaluateFormula(formula, { getCell: () => 1 });
      if (sample === '#NAME?' || sample === '#VALUE!') fail('조건부 서식 수식이 유효하지 않거나 지원되지 않습니다.');
      if (formula === rule.formula) return { bytes: model.bytes, model };
      const formulaRanges = [...findElementRanges(updatedRuleXml, 'formula', undefined, true).values()];
      if (formulaRanges.length !== 1) fail('조건부 서식 규칙을 안전하게 변경할 수 없습니다.');
      const range = formulaRanges[0], formulaTag = updatedRuleXml.slice(range.start, range.startTagEnd), name = /^<\s*([A-Za-z_][\w.:-]*)/.exec(formulaTag)?.[1];
      if (!name || updatedRuleXml.slice(range.startTagEnd, range.closingStart).includes('<!--')) fail('조건부 서식 규칙을 안전하게 변경할 수 없습니다.');
      const replacement = range.selfClosing ? '<' + name + '>' + xmlEscape(formula) + '</' + name + '>' : updatedRuleXml.slice(range.start, range.startTagEnd) + xmlEscape(formula) + updatedRuleXml.slice(range.closingStart, range.end);
      updatedRuleXml = updatedRuleXml.slice(0, range.start) + replacement + updatedRuleXml.slice(range.end);
    } else fail('선택한 조건부 서식 규칙을 편집할 수 없습니다.');
  } else if (edit?.kind === 'colors') {
    const expected = rule.type === 'dataBar' ? 1 : rule.type === 'colorScale' ? rule.colorScale.colors.length : 0;
    if (!expected || !Array.isArray(edit.colors) || edit.colors.length !== expected) fail('선택한 조건부 서식 규칙을 편집할 수 없습니다.');
    const colors = edit.colors.map(value => {
      const color = String(value ?? '').trim();
      if (!/^#[0-9a-f]{6}$/i.test(color)) fail('조건부 서식 색상 값이 유효하지 않습니다.');
      return color.slice(1).toUpperCase();
    });
    const oldColors = rule.type === 'dataBar' ? [rule.dataBar.color] : rule.colorScale.colors;
    if (colors.every((color, index) => '#' + color === oldColors[index].toUpperCase())) return { bytes: model.bytes, model };
    const colorRanges = [...findElementRanges(updatedRuleXml, 'color', undefined, true).values()];
    if (colorRanges.length !== expected) fail('조건부 서식 규칙을 안전하게 변경할 수 없습니다.');
    const colorPatches = colorRanges.map((range, index) => ({ range, index })).sort((left, right) => right.range.start - left.range.start);
    for (const { index, range } of colorPatches) updatedRuleXml = updatedRuleXml.slice(0, range.start) + '<color rgb="FF' + colors[index] + '"/>' + updatedRuleXml.slice(range.end);
  } else if (edit?.kind === 'threshold') {
    if (rule.type !== 'colorScale' || rule.colorScale.stops.length !== 3 || !edit.stop || typeof edit.stop !== 'object' || Array.isArray(edit.stop) || !['num', 'percent', 'percentile'].includes(edit.stop.type)) fail('선택한 조건부 서식 규칙을 편집할 수 없습니다.');
    const valueText = String(edit.stop.value ?? '').trim(), value = Number(valueText);
    if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(valueText) || !Number.isFinite(value) || edit.stop.type !== 'num' && (value < 0 || value > 100)) fail('조건부 서식 기준값이 유효하지 않습니다.');
    const currentStop = rule.colorScale.stops[1];
    if (currentStop.type === edit.stop.type && currentStop.value === value) return { bytes: model.bytes, model };
    const stopRanges = [...findElementRanges(updatedRuleXml, 'cfvo', undefined, true).values()];
    if (stopRanges.length !== 3) fail('조건부 서식 기준을 안전하게 변경할 수 없습니다.');
    const range = stopRanges[1], replacement = '<cfvo type="' + edit.stop.type + '" val="' + xmlEscape(valueText) + '"/>';
    updatedRuleXml = updatedRuleXml.slice(0, range.start) + replacement + updatedRuleXml.slice(range.end);
  } else if (edit?.kind === 'data-bar-display') {
    const display = edit.dataBar, expectedKeys = ['minLength', 'maxLength', 'showValue', 'gradient'];
    if (rule.type !== 'dataBar' || !display || typeof display !== 'object' || Array.isArray(display) || Object.keys(display).length !== expectedKeys.length || Object.keys(display).some(key => !expectedKeys.includes(key)) || !expectedKeys.every(key => Object.hasOwn(display, key)) || ![display.minLength, display.maxLength].every(value => Number.isInteger(value) && value >= 0 && value <= 100) || display.minLength > display.maxLength || typeof display.showValue !== 'boolean' || typeof display.gradient !== 'boolean') fail('데이터 막대 표시 설정이 유효하지 않습니다.');
    const current = rule.dataBar;
    if (display.minLength === current.minLength && display.maxLength === current.maxLength && display.showValue === current.showValue && display.gradient === current.gradient) return { bytes: model.bytes, model };
    const dataBarRanges = [...findElementRanges(updatedRuleXml, 'dataBar', undefined, true).values()];
    if (dataBarRanges.length !== 1) fail('조건부 서식 규칙을 안전하게 변경할 수 없습니다.');
    const range = dataBarRanges[0];
    let opening = updatedRuleXml.slice(range.start, range.startTagEnd);
    for (const [name, value] of Object.entries(display)) opening = setTagAttribute(opening, name, typeof value === 'boolean' ? (value ? '1' : '0') : value);
    updatedRuleXml = updatedRuleXml.slice(0, range.start) + opening + updatedRuleXml.slice(range.startTagEnd);
  } else if (edit?.kind === 'icon-set-display') {
    const display = edit.iconSet, expectedKeys = ['showValue', 'reverse'];
    if (rule.type !== 'iconSet' || !display || typeof display !== 'object' || Array.isArray(display) || Object.keys(display).length !== expectedKeys.length || Object.keys(display).some(key => !expectedKeys.includes(key)) || !expectedKeys.every(key => Object.hasOwn(display, key)) || typeof display.showValue !== 'boolean' || typeof display.reverse !== 'boolean') fail('아이콘 집합 표시 설정이 유효하지 않습니다.');
    const current = rule.iconSet;
    if (display.showValue === current.showValue && display.reverse === current.reverse) return { bytes: model.bytes, model };
    const iconSetRanges = [...findElementRanges(updatedRuleXml, 'iconSet', undefined, true).values()];
    if (iconSetRanges.length !== 1) fail('조건부 서식 아이콘 집합을 안전하게 변경할 수 없습니다.');
    const range = iconSetRanges[0];
    let opening = updatedRuleXml.slice(range.start, range.startTagEnd);
    for (const [name, value] of Object.entries(display)) opening = setTagAttribute(opening, name, value ? '1' : '0');
    updatedRuleXml = updatedRuleXml.slice(0, range.start) + opening + updatedRuleXml.slice(range.startTagEnd);
  } else if (edit?.kind === 'icon-set-mapping') {
    if (rule.type !== 'iconSet' || !Array.isArray(edit.customIcons) || edit.customIcons.length !== rule.iconSet.thresholds.length) fail('사용자 지정 아이콘 매핑이 유효하지 않습니다.');
    const mappings = edit.customIcons.map((mapping, index) => {
      if (!mapping || typeof mapping !== 'object' || Array.isArray(mapping) || Object.keys(mapping).length !== 2 || Object.keys(mapping).some(key => !['iconSet', 'iconId'].includes(key)) || !Object.hasOwn(ICON_SET_LENGTHS, mapping.iconSet) || !Number.isInteger(mapping.iconId) || mapping.iconId < 0 || mapping.iconId >= ICON_SET_LENGTHS[mapping.iconSet]) fail('사용자 지정 아이콘 번호 또는 집합이 유효하지 않습니다.');
      return { iconSet: mapping.iconSet, iconId: mapping.iconId, index };
    });
    const current = rule.iconSet.customIcons || rule.iconSet.thresholds.map((_, index) => ({ iconSet: rule.iconSet.name, iconId: index }));
    if (mappings.every((mapping, index) => mapping.iconSet === current[index].iconSet && mapping.iconId === current[index].iconId)) return { bytes: model.bytes, model };
    const iconSetRanges = [...findElementRanges(updatedRuleXml, 'iconSet', undefined, true).values()];
    if (iconSetRanges.length !== 1) fail('조건부 서식 아이콘 집합을 안전하게 변경할 수 없습니다.');
    const range = iconSetRanges[0], iconSetXml = updatedRuleXml.slice(range.start, range.end);
    const iconRanges = [...findElementRanges(iconSetXml, 'cfIcon', undefined, true).values()];
    if (iconRanges.length && iconRanges.length !== current.length) fail('기존 사용자 지정 아이콘 매핑을 안전하게 변경할 수 없습니다.');
    let tag = iconRanges.length ? /^<\s*([A-Za-z_][\w.:-]*)/.exec(iconSetXml.slice(iconRanges[0].start, iconRanges[0].startTagEnd))?.[1] : /^<\s*([A-Za-z_][\w.:-]*)/.exec(updatedRuleXml.slice(range.start, range.startTagEnd))?.[1]?.replace(/[^:]+$/, 'cfIcon');
    if (!tag || tag.split(':').at(-1) !== 'cfIcon') fail('조건부 서식 아이콘 매핑을 안전하게 변경할 수 없습니다.');
    let inner = iconSetXml.slice(range.startTagEnd - range.start, range.closingStart - range.start);
    const innerStart = range.startTagEnd - range.start;
    const localIcons = iconRanges.map(item => ({ start: item.start - innerStart, end: item.end - innerStart })).sort((left, right) => right.start - left.start);
    for (const item of localIcons) inner = inner.slice(0, item.start) + inner.slice(item.end);
    const lastStop = [...findElementRanges(iconSetXml, 'cfvo', undefined, true).values()].at(-1);
    if (!lastStop) fail('조건부 서식 아이콘 기준이 없어 매핑을 바꿀 수 없습니다.');
    const insertion = lastStop.end - innerStart;
    const iconXml = mappings.map(({ iconSet, iconId }) => '<' + tag + ' iconSet="' + iconSet + '" iconId="' + iconId + '"/>').join('');
    inner = inner.slice(0, insertion) + iconXml + inner.slice(insertion);
    const opening = setTagAttribute(updatedRuleXml.slice(range.start, range.startTagEnd), 'custom', '1');
    updatedRuleXml = updatedRuleXml.slice(0, range.start) + opening + inner + updatedRuleXml.slice(range.closingStart, range.end) + updatedRuleXml.slice(range.end);
  } else if (edit?.kind === 'icon-set-style') {
    if (rule.type !== 'iconSet') fail('선택한 조건부 서식 규칙을 편집할 수 없습니다.');
    if (typeof edit.name !== 'string' || !Object.hasOwn(ICON_SET_LENGTHS, edit.name)) fail('지원하지 않는 아이콘 집합입니다.');
    const current = rule.iconSet;
    if (edit.name === current.name) return { bytes: model.bytes, model };
    const iconCount = ICON_SET_LENGTHS[edit.name], iconSetRanges = [...findElementRanges(updatedRuleXml, 'iconSet', undefined, true).values()];
    if (iconSetRanges.length !== 1) fail('조건부 서식 아이콘 집합을 안전하게 변경할 수 없습니다.');
    const range = iconSetRanges[0];
    let opening = setTagAttribute(updatedRuleXml.slice(range.start, range.startTagEnd), 'iconSet', edit.name);
    if (iconCount === current.thresholds.length) {
      updatedRuleXml = updatedRuleXml.slice(0, range.start) + opening + updatedRuleXml.slice(range.startTagEnd);
    } else {
      const stopRanges = [...findElementRanges(updatedRuleXml, 'cfvo', undefined, true).values()];
      if (stopRanges.length !== current.thresholds.length) fail('아이콘 집합 기준을 안전하게 변경할 수 없습니다.');
      let inner = updatedRuleXml.slice(range.startTagEnd, range.closingStart);
      const localStops = stopRanges.map(stop => ({ start: stop.start - range.startTagEnd, end: stop.end - range.startTagEnd })).sort((left, right) => right.start - left.start);
      for (const stop of localStops) inner = inner.slice(0, stop.start) + inner.slice(stop.end);
      if (inner.trim()) fail('아이콘 집합의 추가 XML을 안전하게 보존할 수 없습니다.');
      const stopTag = /^<([\w.-]+(?::[\w.-]+)?)\b/.exec(updatedRuleXml.slice(stopRanges[0].start, stopRanges[0].startTagEnd))?.[1];
      if (!stopTag || stopTag.split(':').at(-1) !== 'cfvo') fail('아이콘 집합 기준을 안전하게 변경할 수 없습니다.');
      const thresholds = standardIconSetThresholdXml(stopTag, iconCount), closing = updatedRuleXml.slice(range.closingStart, range.end);
      updatedRuleXml = updatedRuleXml.slice(0, range.start) + opening + thresholds + inner + closing + updatedRuleXml.slice(range.end);
    }
  } else if (edit?.kind === 'icon-set-threshold') {
    const stop = edit.stop, allowedStopKeys = ['type', 'value', 'formula', 'gte'];
    if (rule.type !== 'iconSet' || !Number.isInteger(edit.index) || edit.index < 0 || edit.index >= rule.iconSet.thresholds.length || !stop || typeof stop !== 'object' || Array.isArray(stop) || Object.keys(stop).some(key => !allowedStopKeys.includes(key)) || typeof stop.gte !== 'boolean' || !['min', 'max', 'num', 'percent', 'percentile', 'formula'].includes(stop.type)) fail('선택한 아이콘 기준이 유효하지 않습니다.');
    let valueText = null, formula = null;
    if (stop.type === 'min' || stop.type === 'max') {
      if (Object.hasOwn(stop, 'value') || Object.hasOwn(stop, 'formula')) fail('아이콘 최솟값·최댓값 기준은 값을 받지 않습니다.');
    } else if (stop.type === 'formula') {
      formula = String(stop.formula ?? '').trim().replace(/^=/, '');
      if (Object.hasOwn(stop, 'value') || !formula || formula.length > 256 || /[!\[\]]/.test(formula)) fail('아이콘 기준 수식은 256자 이하의 같은 시트 수식이어야 합니다.');
      const sample = evaluateFormula(formula, { getCell: () => 1 });
      if (sample === '#NAME?' || sample === '#VALUE!') fail('아이콘 기준 수식에 지원하지 않는 함수가 있습니다.');
    } else {
      valueText = String(stop.value ?? '').trim();
      const value = Number(valueText);
      if (Object.hasOwn(stop, 'formula') || !/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(valueText) || !Number.isFinite(value) || stop.type !== 'num' && (value < 0 || value > 100)) fail('아이콘 기준값이 유효하지 않습니다.');
    }
    const current = rule.iconSet.thresholds[edit.index];
    if (current.type === stop.type && current.gte === stop.gte && (stop.type === 'min' || stop.type === 'max' || stop.type === 'formula' && current.formula === formula || stop.type !== 'formula' && stop.type !== 'min' && stop.type !== 'max' && current.value === Number(valueText))) return { bytes: model.bytes, model };
    const stopRanges = [...findElementRanges(updatedRuleXml, 'cfvo', undefined, true).values()];
    if (stopRanges.length !== rule.iconSet.thresholds.length) fail('아이콘 집합 기준을 안전하게 변경할 수 없습니다.');
    const attributes = ['type="' + stop.type + '"'];
    if (stop.type === 'formula') attributes.push('val="' + xmlEscape(formula) + '"');
    else if (valueText !== null) attributes.push('val="' + xmlEscape(valueText) + '"');
    if (!stop.gte) attributes.push('gte="0"');
    const range = stopRanges[edit.index], replacement = '<cfvo ' + attributes.join(' ') + '/>';
    updatedRuleXml = updatedRuleXml.slice(0, range.start) + replacement + updatedRuleXml.slice(range.end);
  } else if (edit?.kind === 'differential-color') {
    if (!['cellIs', 'expression'].includes(rule.type) || !['fontColor', 'fillColor'].includes(edit.property) || !model.stylesPath || !model.files[model.stylesPath]) fail('선택한 조건부 서식 규칙을 편집할 수 없습니다.');
    let color = edit.color;
    if (color !== null) {
      color = String(color ?? '').trim();
      if (!/^#[0-9a-f]{6}$/i.test(color)) fail('조건부 서식 색상 값이 유효하지 않습니다.');
      color = color.toUpperCase();
    }
    const oldColor = rule.style?.[edit.property === 'fontColor' ? 'color' : 'fill'] || null;
    if (color === oldColor?.toUpperCase() || color === null && oldColor === null) return { bytes: model.bytes, model };
    const originalStylesXml = new TextDecoder().decode(model.files[model.stylesPath]);
    const styled = patchDifferentialFormatColor(originalStylesXml, rule.dxfId, edit.property, color);
    updatedStylesXml = styled.xml;
    updatedRuleXml = setTagAttribute(updatedRuleXml.slice(0, ruleOpeningEnd), 'dxfId', styled.dxfId) + updatedRuleXml.slice(ruleOpeningEnd);
  } else fail('선택한 조건부 서식 규칙을 편집할 수 없습니다.');

  if (updatedRuleXml === sheet.xml.slice(target.start, target.end)) return { bytes: model.bytes, model };
  const worksheetXml = sheet.xml.slice(0, target.start) + updatedRuleXml + sheet.xml.slice(target.end);
  const outputFiles = { ...model.files, [sheet.path]: new TextEncoder().encode(worksheetXml) };
  if (updatedStylesXml !== null) outputFiles[model.stylesPath] = new TextEncoder().encode(updatedStylesXml);
  const bytes = zipSync(outputFiles, { level: 0 });
  const checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('조건부 서식 규칙 편집 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  const changedPaths = new Set([sheet.path, ...(updatedStylesXml !== null ? [model.stylesPath] : [])]);
  for (const [name, before] of Object.entries(model.files)) if (!changedPaths.has(name) && (!checked.files[name] || Buffer.compare(Buffer.from(before), Buffer.from(checked.files[name])) !== 0)) fail('조건부 서식 규칙 편집 중 비수정 패키지 항목이 달라졌습니다: ' + name);
  return { bytes, model: checked };
}

function subtractRange(source, cut) {
  if (cut.top > source.bottom || cut.bottom < source.top || cut.left > source.right || cut.right < source.left) return [source];
  const top = Math.max(source.top, cut.top), bottom = Math.min(source.bottom, cut.bottom), left = Math.max(source.left, cut.left), right = Math.min(source.right, cut.right);
  return [
    { top: source.top, bottom: top - 1, left: source.left, right: source.right },
    { top: bottom + 1, bottom: source.bottom, left: source.left, right: source.right },
    { top, bottom, left: source.left, right: left - 1 },
    { top, bottom, left: right + 1, right: source.right }
  ].filter(range => range.top <= range.bottom && range.left <= range.right);
}

function hyperlinkCellReference(row, column) {
  return columnLetters(column + 1) + (row + 1);
}

function validateExternalHyperlinkTarget(target) {
  if (typeof target !== 'string' || !target.trim() || target.length > 2048 || /[\u0000-\u001f\u007f\\]/.test(target)) fail('하이퍼링크 주소가 유효하지 않습니다.');
  let uri;
  try { uri = new URL(target); } catch { fail('하이퍼링크 주소가 유효하지 않습니다.'); }
  const scheme = uri.protocol.toLowerCase();
  if (!['http:', 'https:', 'mailto:', 'file:'].includes(scheme) || (scheme === 'http:' || scheme === 'https:') && (!uri.hostname || uri.username || uri.password) || (scheme === 'mailto:' || scheme === 'file:') && !uri.pathname) fail('지원하지 않는 하이퍼링크 주소입니다.');
}

function patchHyperlinks(model, sheetIndex, row, column, request = null) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  const sheet = model.sheets[sheetIndex];
  if (!sheet || !Number.isInteger(row) || row < 0 || row >= LIMITS.rows || !Number.isInteger(column) || column < 0 || column >= LIMITS.columns || request !== null && (typeof request !== 'object' || Array.isArray(request))) fail('하이퍼링크 수정 요청이 유효하지 않습니다.');
  const cellRange = { top: row, bottom: row, left: column, right: column };
  const current = sheet.hyperlinks || [];
  const overlapping = current.filter(link => row >= link.top && row <= link.bottom && column >= link.left && column <= link.right);
  if (request === null && !overlapping.length) return { bytes: model.bytes, model };
  let target = null, location = null, display = null, tooltip = null;
  if (request !== null) {
    target = request.target ?? null;
    location = request.location ?? null;
    display = request.display ?? null;
    tooltip = request.tooltip ?? null;
    if (Boolean(target) === Boolean(location)) fail('외부 주소 또는 내부 위치 중 하나만 지정해야 합니다.');
    if (target) validateExternalHyperlinkTarget(target);
    if (location !== null && (typeof location !== 'string' || !location.trim() || location.length > 512 || /[\u0000-\u001f\u007f]/.test(location))) fail('내부 하이퍼링크 위치가 유효하지 않습니다.');
    if (display !== null && (typeof display !== 'string' || display.length > 255 || /[\u0000-\u001f\u007f]/.test(display))) fail('하이퍼링크 표시 이름은 제어 문자를 제외하고 255자 이하여야 합니다.');
    if (tooltip !== null && (typeof tooltip !== 'string' || tooltip.length > 255 || /[\u0000-\u001f\u007f]/.test(tooltip))) fail('하이퍼링크 도구 설명은 제어 문자를 제외하고 255자 이하여야 합니다.');
    if (overlapping.length === 1) {
      const previous = overlapping[0];
      if (previous.top === row && previous.bottom === row && previous.left === column && previous.right === column && previous.target === target && previous.location === location && previous.display === display && previous.tooltip === tooltip) return { bytes: model.bytes, model };
    }
  }

  let links = [];
  for (const link of current) {
    const remainder = subtractRange(link, cellRange);
    for (const range of remainder) links.push({ ...link, ...range });
  }

  const relsPath = path.join(path.dirname(sheet.path), '_rels', path.basename(sheet.path) + '.rels');
  const originalRelsBytes = model.files[relsPath];
  let relsXml = originalRelsBytes ? new TextDecoder().decode(originalRelsBytes) : null;
  let relsPrefix = relsXml ? /^\s*<([A-Za-z_][\w.-]*:)?Relationships\b/.exec(relsXml)?.[1] || '' : '';
  const relationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
  let newRelationshipId = null;
  if (target) {
    let relationships = new Map();
    if (relsXml) {
      const rootPrefix = /^\s*<([A-Za-z_][\w.-]*:)?Relationships\b/.exec(relsXml)?.[1] || '';
      relsPrefix = rootPrefix;
      const rootName = rootPrefix + 'Relationships';
      const root = parseXml(originalRelsBytes, relsPath).root;
      relationships = parseRelationships(root, path.dirname(sheet.path), ['table', 'hyperlink', 'drawing']);
      const rootRange = findElementRanges(relsXml, rootName).get('');
      if (!rootRange) fail('워크시트 관계 파일 구성이 유효하지 않습니다.');
      if (rootRange.selfClosing) {
        const opening = relsXml.slice(rootRange.start, rootRange.startTagEnd).replace(/\/\s*>$/, '>');
        relsXml = relsXml.slice(0, rootRange.start) + opening + '</' + rootName + '>' + relsXml.slice(rootRange.end);
      }
    } else {
      relsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    }
    const ids = new Set(relationships.keys());
    let nextId = 1;
    while (ids.has('rId' + nextId)) nextId++;
    newRelationshipId = 'rId' + nextId;
  }
  if (request !== null) links.push({ ...cellRange, relationshipId: newRelationshipId, location, display, tooltip, target });
  if (links.length > 10000) fail('워크시트 하이퍼링크 개수 제한을 초과했습니다.');

  const worksheetPrefix = /^\s*<([A-Za-z_][\w.-]*:)?worksheet\b/.exec(sheet.xml)?.[1] || '';
  const elementName = worksheetPrefix + 'hyperlinks';
  const containers = [...findElementRanges(sheet.xml, elementName).values()];
  if (containers.length > 1) fail('중복된 워크시트 하이퍼링크 컨테이너를 수정할 수 없습니다.');
  const attribute = value => xmlEscape(String(value)).replaceAll('"', '&quot;');
  const replacement = links.length ? '<' + elementName + '>' + links.map(link => {
    const attributes = [['ref', hyperlinkCellReference(link.top, link.left) + (link.top === link.bottom && link.left === link.right ? '' : ':' + hyperlinkCellReference(link.bottom, link.right))]];
    if (link.relationshipId) attributes.push(['r:id', link.relationshipId]);
    if (link.location) attributes.push(['location', link.location]);
    if (link.display) attributes.push(['display', link.display]);
    if (link.tooltip) attributes.push(['tooltip', link.tooltip]);
    return '<' + worksheetPrefix + 'hyperlink' + attributes.map(([name, value]) => ' ' + name + '="' + attribute(value) + '"').join('') + '/>';
  }).join('') + '</' + elementName + '>' : '';
  let worksheetXml = sheet.xml;
  if (containers.length) {
    const range = containers[0];
    worksheetXml = worksheetXml.slice(0, range.start) + replacement + worksheetXml.slice(range.end);
  } else if (replacement) {
    const laterNames = ['printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];
    const later = laterNames.flatMap(name => [...findElementRanges(sheet.xml, worksheetPrefix + name).values()].map(range => range.start));
    const precedingNames = ['autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting', 'dataValidations'];
    const preceding = precedingNames.flatMap(name => [...findElementRanges(sheet.xml, worksheetPrefix + name, undefined, name === 'conditionalFormatting').values()].map(range => range.end));
    const sheetData = findElementRanges(sheet.xml, worksheetPrefix + 'sheetData').get('');
    const at = later.length ? Math.min(...later) : preceding.length ? Math.max(...preceding) : sheetData?.end;
    if (!Number.isInteger(at)) fail('하이퍼링크를 워크시트에 안전하게 추가할 수 없습니다.');
    worksheetXml = sheet.xml.slice(0, at) + replacement + sheet.xml.slice(at);
  } else return { bytes: model.bytes, model };

  const outputFiles = Object.assign({}, model.files);
  const modifiedPaths = new Set([sheet.path]);
  if (target) {
    const rootName = relsPrefix + 'Relationships';
    const rootRange = findElementRanges(relsXml, rootName).get('');
    if (!rootRange || rootRange.selfClosing) fail('워크시트 관계 파일에 하이퍼링크를 추가할 수 없습니다.');
    const relation = '<' + relsPrefix + 'Relationship Id="' + newRelationshipId + '" Type="' + relationshipType + '" Target="' + attribute(target) + '" TargetMode="External"/>';
    relsXml = relsXml.slice(0, rootRange.closingStart) + relation + relsXml.slice(rootRange.closingStart);
  }
  if (relsXml) {
    const rootName = relsPrefix + 'Relationships', relationName = relsPrefix + 'Relationship';
    const usedIds = new Set(links.map(link => link.relationshipId).filter(Boolean));
    const ranges = findElementRanges(relsXml, relationName, 'Id');
    const removals = [];
    for (const link of overlapping) {
      const id = link.relationshipId;
      if (!id || usedIds.has(id)) continue;
      const range = ranges.get(id);
      if (range) removals.push({ start: range.start, end: range.end });
    }
    for (const removal of removals.sort((left, right) => right.start - left.start)) relsXml = relsXml.slice(0, removal.start) + relsXml.slice(removal.end);
    const root = findElementRanges(relsXml, rootName).get('');
    if (!root) fail('워크시트 관계 파일 구성이 유효하지 않습니다.');
    outputFiles[relsPath] = new TextEncoder().encode(relsXml);
    modifiedPaths.add(relsPath);
  }

  if (target) {
    const root = findElementRanges(worksheetXml, worksheetPrefix + 'worksheet').get('');
    if (!root) fail('워크시트 루트 요소를 찾을 수 없습니다.');
    const opening = worksheetXml.slice(root.start, root.startTagEnd);
    const namespace = parseTagAttributes(opening)['xmlns:r'];
    if (namespace && namespace !== 'http://schemas.openxmlformats.org/officeDocument/2006/relationships') fail('워크시트 관계 네임스페이스가 유효하지 않습니다.');
    if (!namespace) {
      const updatedOpening = setTagAttribute(opening, 'xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
      worksheetXml = worksheetXml.slice(0, root.start) + updatedOpening + worksheetXml.slice(root.startTagEnd);
    }
  }
  outputFiles[sheet.path] = new TextEncoder().encode(worksheetXml);
  const bytes = zipSync(outputFiles, { level: 0 }), checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('하이퍼링크 저장 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  for (const [name, before] of Object.entries(model.files)) {
    if (modifiedPaths.has(name)) continue;
    if (!checked.files[name] || Buffer.compare(Buffer.from(before), Buffer.from(checked.files[name])) !== 0) fail('하이퍼링크 저장 중 비수정 패키지 항목이 달라졌습니다: ' + name);
  }
  return { bytes, model: checked };
}

function patchHiddenRows(model, edits, updateDrawings = true) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  if (updateDrawings) model = preserveImportedFilterVisibility(model);
  if (!Array.isArray(edits) || edits.length > LIMITS.sheets) fail('숨김 행 수정 요청이 유효하지 않습니다.');
  if (!edits.length) return { bytes: model.bytes, model };
  const outputFiles = Object.assign({}, model.files), modifiedPaths = new Set(), seen = new Set();
  for (const edit of edits) {
    const { sheetIndex, rows, hidden } = edit || {}, sheet = model.sheets[sheetIndex];
    if (!Number.isInteger(sheetIndex) || !sheet || seen.has(sheetIndex) || !Array.isArray(rows) || rows.length > LIMITS.rows || typeof hidden !== 'boolean' || rows.some((row, index) => !Number.isInteger(row) || row < 0 || row >= LIMITS.rows || (index && row <= rows[index - 1]))) fail('숨김 행 수정 요청이 유효하지 않습니다.');
    seen.add(sheetIndex);
    if (!rows.length) continue;
    const rowTags = [...sheet.xml.matchAll(/<row\b[^>]*>/g)], rowTagsByNumber = new Map();
    for (const match of rowTags) {
      const index = match[0].match(/\br\s*=\s*(["'])(\d+)\1/);
      if (index) {
        const number = Number(index[2]);
        if (rowTagsByNumber.has(number)) fail('중복된 행 요소를 안전하게 수정할 수 없습니다.');
        rowTagsByNumber.set(number, match);
      }
    }
    const orderedRowNumbers = [...rowTagsByNumber.keys()].sort((left, right) => left - right);
    const operations = [], insertions = new Map();
    for (const row of rows) {
      const rowNumber = row + 1, matching = rowTagsByNumber.get(rowNumber);
      if (!matching) {
        if (!hidden) continue;
        const data = findElementRanges(sheet.xml, 'sheetData').get('');
        if (!data || data.selfClosing) fail('숨김 행을 워크시트에 안전하게 추가할 수 없습니다.');
        let low = 0, high = orderedRowNumbers.length;
        while (low < high) { const middle = (low + high) >>> 1; if (orderedRowNumbers[middle] <= rowNumber) low = middle + 1; else high = middle; }
        const next = low < orderedRowNumbers.length ? rowTagsByNumber.get(orderedRowNumbers[low]) : null;
        const position = next?.index ?? data.closingStart;
        if (!insertions.has(position)) insertions.set(position, []);
        insertions.get(position).push('<row r="' + rowNumber + '" hidden="1"' + (sheet.rowVisibilityPrefix ? ' ' + sheet.rowVisibilityPrefix + ':manualHidden="1"' : '') + '/>');
        continue;
      }
      const match = matching, tag = match[0], hiddenAttribute = /\s+hidden\s*=\s*(["'])[^"']*\1/i;
      let replacement = tag;
      if (hidden) replacement = hiddenAttribute.test(tag) ? tag.replace(hiddenAttribute, ' hidden="1"') : tag.replace(/\s*\/?\s*>$/, part => ' hidden="1"' + part);
      else {
        replacement = tag.replace(hiddenAttribute, '');
        if (parseTagAttributes(tag).ht !== undefined && Number(parseTagAttributes(tag).ht) === 0) replacement = replacement.replace(/\s+ht\s*=\s*(["'])[^"']*\1/i, '').replace(/\s+customHeight\s*=\s*(["'])[^"']*\1/i, '');
        if (/^<row\s+r\s*=\s*(["'])\d+\1\s*\/>$/i.test(replacement)) replacement = '';
      }
      if (replacement && sheet.rowVisibilityPrefix) replacement = setTagAttribute(replacement, sheet.rowVisibilityPrefix + ':manualHidden', hidden ? '1' : '0');
      if (replacement !== tag) operations.push({ start: match.index, end: match.index + tag.length, replacement });
    }
    for (const [position, replacements] of insertions) operations.push({ start: position, end: position, replacement: replacements.join('') });
    if (!operations.length) continue;
    let xml = sheet.xml;
    for (const operation of operations.sort((a, b) => b.start - a.start)) xml = xml.slice(0, operation.start) + operation.replacement + xml.slice(operation.end);
    outputFiles[sheet.path] = new TextEncoder().encode(xml); modifiedPaths.add(sheet.path);
  }
  if (!modifiedPaths.size) return { bytes: model.bytes, model };
  if (updateDrawings) updateDimensionDrawings(model, outputFiles, modifiedPaths);
  const bytes = zipSync(outputFiles, { level: 0 }), checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('숨김 행 저장 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  for (const [name, before] of Object.entries(model.files)) if (!modifiedPaths.has(name) && (!checked.files[name] || Buffer.compare(Buffer.from(before), Buffer.from(checked.files[name])) !== 0)) fail('숨김 행 저장 중 비수정 패키지 항목이 달라졌습니다: ' + name);
  return { bytes, model: checked };
}

function patchRowOutlineTags(sheet, levelUpdates, collapsedUpdates) {
  const xml = sheet.xml, ranges = findElementRanges(xml, 'row', 'r'), operations = [], insertions = new Map();
  const updates = new Map();
  for (const [row, level] of levelUpdates) updates.set(row, { ...(updates.get(row) || {}), outlineLevel: level > 0 ? String(level) : null });
  for (const [row, collapsed] of collapsedUpdates) updates.set(row, { ...(updates.get(row) || {}), collapsed: collapsed ? '1' : null });
  const ordered = [...ranges].map(([number, range]) => ({ number: Number(number), range })).sort((a, b) => a.number - b.number);
  for (const [row, attributes] of [...updates].sort((a, b) => a[0] - b[0])) {
    const number = row + 1;
    if (!Number.isInteger(number) || number < 1 || number > LIMITS.rows) fail('행 그룹 범위가 워크시트 제한을 벗어났습니다.');
    const range = ranges.get(String(number));
    if (range) {
      const opening = xml.slice(range.start, range.startTagEnd);
      const next = Object.entries(attributes).reduce((tag, [name, value]) => value === null ? removeTagAttribute(tag, name) : setTagAttribute(tag, name, value), opening);
      if (next !== opening) operations.push({ start: range.start, end: range.startTagEnd, replacement: next });
      continue;
    }
    const tag = '<row r="' + number + '"' + Object.entries(attributes).filter(([, value]) => value !== null).map(([name, value]) => ' ' + name + '="' + xmlEscape(value) + '"').join('') + '/>';
    const data = findElementRanges(xml, 'sheetData').get('');
    if (!data || data.selfClosing) fail('행 그룹 메타데이터를 워크시트에 추가할 수 없습니다.');
    const next = ordered.find(item => item.number > number), position = next?.range.start ?? data.closingStart;
    if (!insertions.has(position)) insertions.set(position, []);
    insertions.get(position).push({ number, tag });
  }
  for (const [position, values] of insertions) operations.push({ start: position, end: position, replacement: values.sort((a, b) => a.number - b.number).map(item => item.tag).join('') });
  let result = xml;
  for (const operation of operations.sort((a, b) => b.start - a.start)) result = result.slice(0, operation.start) + operation.replacement + result.slice(operation.end);
  return result;
}

function patchColumnOutlineTags(sheet, updates) {
  const xml = sheet.xml, ranges = [...findElementRanges(xml, 'col', 'min').values()].sort((a, b) => a.start - b.start), columns = new Map();
  for (const range of ranges) {
    const tag = xml.slice(range.start, range.startTagEnd), attributes = parseTagAttributes(tag), min = Number(attributes.min), max = Number(attributes.max);
    if (!range.selfClosing || !Number.isInteger(min) || !Number.isInteger(max) || min < 1 || max < min || max > LIMITS.columns) fail('열 그룹 정의를 안전하게 수정할 수 없습니다.');
    for (let column = min - 1; column < max; column++) {
      if (columns.has(column)) fail('겹치는 열 정의를 안전하게 수정할 수 없습니다.');
      const next = { ...attributes }; delete next.min; delete next.max; columns.set(column, next);
    }
  }
  for (const [column, attributes] of updates) {
    if (!Number.isInteger(column) || column < 0 || column >= LIMITS.columns) fail('열 그룹 범위가 워크시트 제한을 벗어났습니다.');
    const next = { ...(columns.get(column) || {}) };
    for (const [name, value] of Object.entries(attributes)) if (value === null) delete next[name]; else next[name] = String(value);
    columns.set(column, next);
  }
  const definitions = [];
  for (const [column, attributes] of [...columns].sort((a, b) => a[0] - b[0])) {
    if (!Object.keys(attributes).length) continue;
    const key = JSON.stringify(attributes), previous = definitions.at(-1);
    if (previous && previous.max === column && previous.key === key) previous.max = column + 1;
    else definitions.push({ min: column, max: column + 1, key, attributes });
  }
  const tags = definitions.map(item => '<col min="' + (item.min + 1) + '" max="' + item.max + '"' + Object.entries(item.attributes).map(([name, value]) => ' ' + name + '="' + xmlEscape(value) + '"').join('') + '/>').join('');
  const container = findElementRanges(xml, 'cols').get('');
  if (container) {
    if (!tags) return xml.slice(0, container.start) + xml.slice(container.end);
    const opening = xml.slice(container.start, container.startTagEnd).replace(/\/\s*>$/, '>');
    return xml.slice(0, container.start) + opening + tags + '</cols>' + xml.slice(container.end);
  }
  if (!tags) return xml;
  const sheetData = findElementRanges(xml, 'sheetData').get('');
  if (!sheetData || sheetData.selfClosing) fail('열 그룹 정의를 워크시트에 추가할 수 없습니다.');
  return xml.slice(0, sheetData.start) + '<cols>' + tags + '</cols>' + xml.slice(sheetData.start);
}

function patchOutlineGroups(model, edits) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  if (!Array.isArray(edits) || edits.length > LIMITS.sheets) fail('행·열 그룹 수정 요청이 유효하지 않습니다.');
  if (!edits.length) return { bytes: model.bytes, model };
  let current = preserveImportedFilterVisibility(model), changed = false;
  const seen = new Set();
  for (const edit of edits) {
    const { sheetIndex, axis, action, start, end, summary } = edit || {}, sheet = current.sheets[sheetIndex];
    const maximum = axis === 'row' ? LIMITS.rows : axis === 'column' ? LIMITS.columns : 0;
    const summaryAction = action === 'collapse-summary' || action === 'expand-summary';
    const levelAction = action === 'show-level';
    const validRange = summaryAction
      ? Number.isInteger(summary) && summary >= 0 && summary < maximum
      : levelAction ? Number.isInteger(edit.level) && edit.level >= 1 && edit.level <= 8
        : Number.isInteger(start) && Number.isInteger(end) && start >= 0 && end >= start && end < maximum && !(axis === 'row' && end - start + 1 > 10000);
    if (!Number.isInteger(sheetIndex) || !sheet || seen.has(sheetIndex) || !['row', 'column'].includes(axis) || !['group', 'ungroup', 'collapse', 'expand', 'collapse-summary', 'expand-summary', 'show-level'].includes(action) || !validRange) fail('행·열 그룹 수정 범위가 유효하지 않습니다.');
    seen.add(sheetIndex);
    const initialTarget = current.sheets[sheetIndex], initialGroups = axis === 'row' ? initialTarget.rowGroups : initialTarget.columnGroups;
    const actionGroups = summaryAction ? initialGroups.filter(group => group.summary === summary) : levelAction ? initialGroups : [];
    if ((summaryAction || levelAction) && !actionGroups.length) fail(summaryAction ? '선택한 요약 위치에 접거나 펼칠 그룹이 없습니다.' : '이 시트에 표시할 행·열 그룹이 없습니다.');
    const coordinates = summaryAction || levelAction
      ? axis === 'column' ? [...new Set(actionGroups.flatMap(group => Array.from({ length: group.end - group.start + 1 }, (_, offset) => group.start + offset)))] : []
      : Array.from({ length: end - start + 1 }, (_, offset) => start + offset);
    if (axis === 'row') current = patchFilterVisibility(current, { ensureVisibilityMetadata: [sheetIndex] }).model;
    else current = ensureColumnVisibilityMetadata(current, sheetIndex, coordinates);
    const target = current.sheets[sheetIndex], levels = axis === 'row' ? target.rowOutlineLevels : target.columnOutlineLevels;
    const groups = axis === 'row' ? target.rowGroups : target.columnGroups;
    const summaryAfter = axis === 'row' ? target.summaryBelow !== false : target.summaryRight !== false;
    const levelUpdates = new Map(), collapsedUpdates = new Map();
    if (action === 'group') {
      const currentLevels = coordinates.map(index => levels[index] || 0), base = currentLevels[0];
      if (currentLevels.some(level => level !== base)) fail('같은 개요 수준의 연속된 행·열만 새 그룹으로 묶을 수 있습니다.');
      if (base >= 7) fail('행·열 그룹은 7단계까지만 만들 수 있습니다.');
      if (groups.some(group => group.collapsed && group.start <= end && group.end >= start)) fail('접힌 그룹을 펼친 뒤 다시 그룹화하세요.');
      const summary = summaryAfter ? end + 1 : start - 1;
      if (summary < 0 || summary >= maximum) fail('요약 행·열을 둘 위치가 없어 그룹을 만들 수 없습니다.');
      for (const coordinate of coordinates) levelUpdates.set(coordinate, base + 1);
    } else if (action === 'ungroup') {
      const currentLevels = coordinates.map(index => levels[index] || 0);
      if (currentLevels.every(level => level === 0)) fail('선택 범위에 해제할 그룹이 없습니다.');
      if (groups.some(group => group.collapsed && group.start <= end && group.end >= start)) fail('그룹을 펼친 뒤 그룹 해제를 실행하세요.');
      for (const coordinate of coordinates) levelUpdates.set(coordinate, Math.max(0, (levels[coordinate] || 0) - 1));
    } else if (summaryAction) {
      const collapsed = action === 'collapse-summary';
      if (actionGroups.every(group => group.collapsed === collapsed)) fail('선택한 요약 위치의 그룹은 이미 요청한 상태입니다.');
      collapsedUpdates.set(summary, collapsed);
    } else if (levelAction) {
      const groupsBySummary = new Map();
      for (const group of groups) groupsBySummary.set(group.summary, [...(groupsBySummary.get(group.summary) || []), group]);
      for (const [summaryCoordinate, summaryGroups] of groupsBySummary) {
        const collapsed = summaryGroups.some(group => group.level >= edit.level);
        if (summaryGroups.some(group => group.collapsed !== collapsed)) collapsedUpdates.set(summaryCoordinate, collapsed);
      }
    } else {
      const matches = groups.filter(group => group.start === start && group.end === end && group.collapsed === (action === 'expand'));
      if (!matches.length) fail('선택한 범위에 접거나 펼칠 그룹이 없습니다.');
      for (const group of matches) collapsedUpdates.set(group.summary, action === 'collapse');
    }
    let xml = axis === 'row'
      ? patchRowOutlineTags(target, levelUpdates, collapsedUpdates)
      : patchColumnOutlineTags(target, new Map([
        ...[...levelUpdates].map(([column, level]) => [column, { outlineLevel: level > 0 ? String(level) : null }]),
        ...[...collapsedUpdates].map(([column, collapsed]) => [column, { collapsed: collapsed ? '1' : null }])
      ]));
    if (!levelUpdates.size && !collapsedUpdates.size) continue;
    const nextLevelMap = { ...levels };
    for (const [coordinate, level] of levelUpdates) { if (level > 0) nextLevelMap[coordinate] = level; else delete nextLevelMap[coordinate]; }
    const maxLevel = Math.max(0, ...Object.values(nextLevelMap));
    const outlineAttribute = axis === 'row' ? 'outlineLevelRow' : 'outlineLevelCol';
    const formatRange = findElementRanges(xml, 'sheetFormatPr').get('');
    if (formatRange) {
      const opening = xml.slice(formatRange.start, formatRange.startTagEnd), updated = maxLevel ? setTagAttribute(opening, outlineAttribute, String(maxLevel)) : removeTagAttribute(opening, outlineAttribute);
      xml = xml.slice(0, formatRange.start) + updated + xml.slice(formatRange.startTagEnd);
    } else if (maxLevel) {
      const sheetProperties = findElementRanges(xml, 'sheetPr').get(''), columns = findElementRanges(xml, 'cols').get(''), data = findElementRanges(xml, 'sheetData').get('');
      const position = columns?.start ?? data?.start;
      if (!Number.isFinite(position)) fail('행·열 그룹 속성을 넣을 위치를 찾을 수 없습니다.');
      const tag = '<sheetFormatPr ' + outlineAttribute + '="' + maxLevel + '"/>';
      const insertAt = sheetProperties && sheetProperties.end <= position ? position : Math.min(position, data?.start ?? position);
      xml = xml.slice(0, insertAt) + tag + xml.slice(insertAt);
    }
    if (xml !== target.xml) {
      const files = { ...current.files, [target.path]: new TextEncoder().encode(xml) };
      const bytes = zipSync(files, { level: 0 }), checked = analyze(bytes, current.extension);
      if (!checked.writable) fail('행·열 그룹 저장 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
      current = checked;
      changed = true;
    }
    if (axis === 'row') current = patchFilterVisibility(current, { ensureVisibilityMetadata: [sheetIndex] }).model;
    else {
      const updated = current.sheets[sheetIndex], wanted = new Set([...(updated.hiddenColumns || []), ...(updated.outlineHiddenColumns || [])]), old = new Set(updated.nativeHiddenColumns || []);
      const show = [...old].filter(column => !wanted.has(column)).sort((a, b) => a - b), hide = [...wanted].filter(column => !old.has(column)).sort((a, b) => a - b);
      if (show.length) current = patchHiddenColumns(current, [{ sheetIndex, columns: show, hidden: false }], false).model;
      if (hide.length) current = patchHiddenColumns(current, [{ sheetIndex, columns: hide, hidden: true }], false).model;
    }
  }
  return changed ? { bytes: current.bytes, model: current } : { bytes: model.bytes, model };
}

function ensureColumnVisibilityMetadata(model, sheetIndex, columns) {
  const sheet = model.sheets[sheetIndex], prefix = sheet.columnVisibilityPrefix || 'jsc', mcPrefix = Object.keys(parseXml(new TextEncoder().encode(sheet.xml), sheet.path).root.attributes).find(key => key.startsWith('xmlns:') && parseXml(new TextEncoder().encode(sheet.xml), sheet.path).root.attributes[key] === MC_NS)?.slice(6) || 'mc';
  let xml = sheet.xml;
  const worksheet = /<worksheet\b[^>]*>/.exec(xml);
  if (!worksheet) fail('열 그룹 표시 메타데이터를 추가할 수 없습니다.');
  let opening = worksheet[0], attrs = parseTagAttributes(opening);
  const actualPrefix = Object.keys(attrs).find(key => key.startsWith('xmlns:') && attrs[key] === COLUMN_VISIBILITY_NS)?.slice(6) || prefix;
  if (!attrs['xmlns:' + actualPrefix]) opening = setTagAttribute(opening, 'xmlns:' + actualPrefix, COLUMN_VISIBILITY_NS);
  const actualMc = Object.keys(attrs).find(key => key.startsWith('xmlns:') && attrs[key] === MC_NS)?.slice(6) || mcPrefix;
  if (!attrs['xmlns:' + actualMc]) opening = setTagAttribute(opening, 'xmlns:' + actualMc, MC_NS);
  attrs = parseTagAttributes(opening);
  const merge = (name, values) => setTagAttribute(opening, name, [...new Set([...(attrs[name] || '').split(/\s+/).filter(Boolean), ...values])].join(' '));
  opening = merge(actualMc + ':Ignorable', [actualPrefix]);
  opening = merge(actualMc + ':PreserveAttributes', [actualPrefix + ':manualHidden']);
  xml = xml.slice(0, worksheet.index) + opening + xml.slice(worksheet.index + worksheet[0].length);
  const manual = new Set(sheet.hiddenColumns || []), updates = new Map(columns.map(column => [column, { [actualPrefix + ':manualHidden']: manual.has(column) ? '1' : '0' }]));
  const updated = patchColumnOutlineTags({ ...sheet, xml }, updates);
  const bytes = zipSync({ ...model.files, [sheet.path]: new TextEncoder().encode(updated) }, { level: 0 }), checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('열 표시 메타데이터 저장 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  return checked;
}

function printAreaReference(sheetName, range) {
  return "'" + sheetName.replaceAll("'", "''") + "'!$" + columnLetters(range.left + 1) + '$' + (range.top + 1) + ':$' + columnLetters(range.right + 1) + '$' + (range.bottom + 1);
}

function replaceWorkbookPrintDefinition(xml, sheetIndex, reference, definitionName = '_xlnm.Print_Area') {
  let found = false;
  xml = xml.replace(/<definedName\b[^>]*>[\s\S]*?<\/definedName>/g, element => {
    const opening = element.slice(0, element.indexOf('>') + 1), attributes = parseTagAttributes(opening);
    if (attributes.name?.toLocaleLowerCase() !== definitionName.toLocaleLowerCase() || attributes.localSheetId === undefined || Number(attributes.localSheetId) !== sheetIndex) return element;
    if (found) fail('인쇄 영역 이름 정의가 중복되었습니다.');
    found = true;
    return reference === null ? '' : opening + xmlEscape(reference) + '</definedName>';
  });
  if (found || reference === null) return xml;
  const entry = '<definedName name="' + definitionName + '" localSheetId="' + sheetIndex + '">' + xmlEscape(reference) + '</definedName>';
  const containers = [...findElementRanges(xml, 'definedNames').values()];
  if (containers.length > 1) fail('이름 정의 목록이 중복되었습니다.');
  const container = containers[0];
  if (container?.selfClosing) return xml.slice(0, container.start) + xml.slice(container.start, container.startTagEnd).replace(/\/\s*>$/, '>') + entry + '</definedNames>' + xml.slice(container.end);
  if (container) return xml.slice(0, container.closingStart) + entry + xml.slice(container.closingStart);
  const following = ['calcPr', 'oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'extLst'];
  const points = following.flatMap(name => [...findElementRanges(xml, name).values()].map(range => range.start));
  const at = points.length ? Math.min(...points) : xml.lastIndexOf('</workbook>');
  if (at < 0) fail('통합문서에 인쇄 영역을 추가할 수 없습니다.');
  return xml.slice(0, at) + '<definedNames>' + entry + '</definedNames>' + xml.slice(at);
}

function patchPrintArea(model, sheetIndex, range) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  const sheet = model.sheets[sheetIndex];
  if (!Number.isInteger(sheetIndex) || !sheet || range !== null && (!range || !['top', 'bottom', 'left', 'right'].every(key => Number.isInteger(range[key])) || range.top < 0 || range.left < 0 || range.bottom < range.top || range.right < range.left || range.bottom >= LIMITS.rows || range.right >= LIMITS.columns)) fail('인쇄 영역 범위가 유효하지 않습니다.');
  const current = sheet.printAreas || [];
  if (range === null && !current.length || range && current.length === 1 && ['top', 'bottom', 'left', 'right'].every(key => current[0][key] === range[key])) return { bytes: model.bytes, model };
  return patchWorkbookPrintDefinition(model, sheetIndex, range === null ? null : printAreaReference(sheet.name, range));
}

function patchWorkbookPrintDefinition(model, sheetIndex, reference, definitionName = '_xlnm.Print_Area') {
  const workbookPath = 'xl/workbook.xml', xml = new TextDecoder().decode(model.files[workbookPath]);
  const updated = replaceWorkbookPrintDefinition(xml, sheetIndex, reference, definitionName);
  if (updated === xml) return { bytes: model.bytes, model };
  const bytes = zipSync({ ...model.files, [workbookPath]: new TextEncoder().encode(updated) }, { level: 0 }), checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('인쇄 영역 저장 후 패키지 검사에서 거절되었습니다.');
  for (const [name, before] of Object.entries(model.files)) if (name !== workbookPath && Buffer.compare(Buffer.from(before), Buffer.from(checked.files[name])) !== 0) fail('인쇄 영역 저장 중 비수정 패키지 항목이 달라졌습니다: ' + name);
  return { bytes, model: checked };
}

function definedNameScope(attributes) { return attributes.localSheetId === undefined ? null : Number(attributes.localSheetId); }

function replaceWorkbookDefinedName(xml, name, localSheetId, reference) {
  const matches = [...xml.matchAll(/<definedName\b[^>]*>[\s\S]*?<\/definedName>/g)];
  const targets = matches.filter(match => {
    const opening = match[0].slice(0, match[0].indexOf('>') + 1), attributes = parseTagAttributes(opening);
    return attributes.name?.toLocaleLowerCase() === name.toLocaleLowerCase() && definedNameScope(attributes) === localSheetId;
  });
  if (targets.length > 1) fail('정의된 이름이 중복되어 안전하게 수정할 수 없습니다.');
  if (targets.length) {
    const target = targets[0];
    if (reference === null) {
      xml = xml.slice(0, target.index) + xml.slice(target.index + target[0].length);
      const containers = [...findElementRanges(xml, 'definedNames').values()];
      if (containers.length > 1) fail('이름 정의 목록이 중복되었습니다.');
      const container = containers[0];
      if (container && !/<definedName\b/.test(xml.slice(container.start, container.end))) xml = xml.slice(0, container.start) + xml.slice(container.end);
      return xml;
    }
    const openingEnd = target[0].indexOf('>') + 1;
    let opening = setTagAttribute(target[0].slice(0, openingEnd), 'name', name);
    if (localSheetId === null) opening = removeTagAttribute(opening, 'localSheetId');
    else opening = setTagAttribute(opening, 'localSheetId', String(localSheetId));
    const replacement = opening + xmlEscape(reference) + '</definedName>';
    return xml.slice(0, target.index) + replacement + xml.slice(target.index + target[0].length);
  }
  if (reference === null) return xml;
  const entry = '<definedName name="' + xmlEscape(name) + '"' + (localSheetId === null ? '' : ' localSheetId="' + localSheetId + '"') + '>' + xmlEscape(reference) + '</definedName>';
  const containers = [...findElementRanges(xml, 'definedNames').values()];
  if (containers.length > 1) fail('이름 정의 목록이 중복되었습니다.');
  const container = containers[0];
  if (container?.selfClosing) return xml.slice(0, container.start) + xml.slice(container.start, container.startTagEnd).replace(/\/\s*>$/, '>') + entry + '</definedNames>' + xml.slice(container.end);
  if (container) return xml.slice(0, container.closingStart) + entry + xml.slice(container.closingStart);
  const following = ['calcPr', 'oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'extLst'];
  const points = following.flatMap(tag => [...findElementRanges(xml, tag).values()].map(range => range.start));
  const at = points.length ? Math.min(...points) : xml.lastIndexOf('</workbook>');
  if (at < 0) fail('통합문서에 이름 정의를 추가할 수 없습니다.');
  return xml.slice(0, at) + '<definedNames>' + entry + '</definedNames>' + xml.slice(at);
}

function definedNameReference(source) {
  return "'" + source.sheetName.replaceAll("'", "''") + "'!$" + columnLetters(source.left + 1) + '$' + (source.top + 1) + ':$' + columnLetters(source.right + 1) + '$' + (source.bottom + 1);
}

function patchDefinedName(model, request) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  if (!request || !['set', 'delete'].includes(request.action) || typeof request.name !== 'string' || !/^[A-Z_\\][A-Z0-9_.]*$/i.test(request.name) || request.name.length > 255 || /^_xlnm\./i.test(request.name) || /^(?:[A-Z]{1,3}[1-9]\d*|R|C|TRUE|FALSE)$/i.test(request.name)) fail('정의된 이름이나 작업이 유효하지 않습니다.');
  const { name, action } = request, localSheetId = request.localSheetId ?? null;
  if (localSheetId !== null && (!Number.isInteger(localSheetId) || !model.sheets[localSheetId])) fail('정의된 이름의 범위 시트가 유효하지 않습니다.');
  const current = model.definedNames.find(item => item.name.toLocaleLowerCase() === name.toLocaleLowerCase() && item.localSheetId === localSheetId);
  if (action === 'delete' && !current) fail('삭제할 정의된 이름을 찾을 수 없습니다.');
  if (current && (current.printArea || current.printTitles || !current.source)) fail('인쇄 이름이나 셀 범위를 해석할 수 없는 이름은 여기서 수정할 수 없습니다.');
  let reference = null;
  if (action === 'set') {
    const source = request.source, sheetIndex = model.sheets.findIndex(sheet => sheet.name.toLocaleLowerCase() === source?.sheetName?.toLocaleLowerCase());
    if (!source || sheetIndex < 0 || !['top', 'bottom', 'left', 'right'].every(key => Number.isInteger(source[key])) || source.top < 0 || source.left < 0 || source.bottom < source.top || source.right < source.left || source.bottom >= LIMITS.rows || source.right >= LIMITS.columns || (source.bottom - source.top + 1) * (source.right - source.left + 1) > 1000) fail('정의된 이름은 1,000셀 이하의 유효한 셀 범위를 참조해야 합니다.');
    if (model.sheets.some(sheet => (sheet.tables || []).some(table => table.name.toLocaleLowerCase() === name.toLocaleLowerCase()))) fail('정의된 이름이 Excel 표 이름과 충돌합니다.');
    if (model.definedNames.some(item => item !== current && item.localSheetId === localSheetId && item.name.toLocaleLowerCase() === name.toLocaleLowerCase())) fail('같은 범위에 동일한 정의된 이름이 있습니다.');
    reference = definedNameReference({ ...source, sheetName: model.sheets[sheetIndex].name });
    if (current && reference === definedNameReference(current.source)) return { bytes: model.bytes, model };
  }
  const workbookPath = 'xl/workbook.xml', xml = new TextDecoder().decode(model.files[workbookPath]);
  let updated = replaceWorkbookDefinedName(xml, name, localSheetId, reference);
  if (updated === xml) return { bytes: model.bytes, model };
  updated = enableFullCalculationOnLoad(updated);
  const outputFiles = { ...model.files, [workbookPath]: new TextEncoder().encode(updated) }, changedFormulaSheets = new Set();
  for (const sheet of model.sheets) {
    const patches = [];
    for (const cell of sheet.cells.values()) {
      if (!cell.formula) continue;
      const cellRange = sheet.ranges.get(cell.address);
      if (!cellRange) fail('이름 변경 후 수식 캐시를 갱신할 셀을 찾을 수 없습니다.');
      const cellXml = sheet.xml.slice(cellRange.start, cellRange.end), value = findElementRanges(cellXml, 'v').get('');
      if (value && !value.selfClosing) patches.push({ start: cellRange.start + value.startTagEnd, end: cellRange.start + value.closingStart, replacement: '' });
    }
    if (!patches.length) continue;
    let sheetXml = sheet.xml;
    for (const patch of patches.sort((left, right) => right.start - left.start)) sheetXml = sheetXml.slice(0, patch.start) + patch.replacement + sheetXml.slice(patch.end);
    outputFiles[sheet.path] = new TextEncoder().encode(sheetXml); changedFormulaSheets.add(sheet.path);
  }
  const bytes = zipSync(outputFiles, { level: 0 }), checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('정의된 이름을 저장한 뒤 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  for (const [partPath, before] of Object.entries(model.files)) if (partPath !== workbookPath && !changedFormulaSheets.has(partPath) && Buffer.compare(Buffer.from(before), Buffer.from(checked.files[partPath])) !== 0) fail('정의된 이름 수정 중 비수정 패키지 항목이 달라졌습니다: ' + partPath);
  return { bytes, model: checked };
}

function validatePrintTitles(titles) {
  if (!titles || Array.isArray(titles)) fail('반복 제목은 50행·10열 이하의 유효한 범위를 입력하세요.');
  for (const [key, maximum, count] of [['rows', LIMITS.rows, 50], ['columns', LIMITS.columns, 10]]) {
    const range = titles[key];
    if (range === null) continue;
    if (!range || Array.isArray(range) || !Number.isInteger(range.start) || !Number.isInteger(range.end) || range.start < 0 || range.end < range.start || range.end >= maximum || range.end - range.start + 1 > count) fail('반복 제목은 50행·10열 이하의 유효한 범위를 입력하세요.');
  }
  return titles;
}

function printTitlesReference(sheetName, titles) {
  const prefix = "'" + sheetName.replaceAll("'", "''") + "'!";
  return [titles.rows ? prefix + '$' + (titles.rows.start + 1) + ':$' + (titles.rows.end + 1) : '', titles.columns ? prefix + '$' + columnLetters(titles.columns.start + 1) + ':$' + columnLetters(titles.columns.end + 1) : ''].filter(Boolean).join(',');
}

function patchPrintTitles(model, sheetIndex, titles) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  const sheet = model.sheets[sheetIndex];
  if (!Number.isInteger(sheetIndex) || !sheet) fail('반복 제목의 시트가 유효하지 않습니다.');
  validatePrintTitles(titles);
  const current = sheet.printTitles || { rows: null, columns: null };
  if (['rows', 'columns'].every(key => titles[key] === null ? current[key] === null : current[key]?.start === titles[key].start && current[key]?.end === titles[key].end)) return { bytes: model.bytes, model };
  return patchWorkbookPrintDefinition(model, sheetIndex, printTitlesReference(sheet.name, titles) || null, '_xlnm.Print_Titles');
}

function transformPrintTitles(titles, operation) {
  const key = operation.axis === 'row' ? 'rows' : 'columns', range = titles[key];
  if (!range) return titles;
  const rectangle = key === 'rows' ? { top: range.start, bottom: range.end, left: 0, right: LIMITS.columns - 1 } : { top: 0, bottom: LIMITS.rows - 1, left: range.start, right: range.end };
  const moved = transformPrintArea(rectangle, operation);
  return { ...titles, [key]: moved ? { start: moved[key === 'rows' ? 'top' : 'left'], end: moved[key === 'rows' ? 'bottom' : 'right'] } : null };
}

function transformPrintArea(range, operation) {
  const first = operation.axis === 'row' ? 'top' : 'left', last = operation.axis === 'row' ? 'bottom' : 'right';
  const maximum = operation.axis === 'row' ? LIMITS.rows : LIMITS.columns;
  const next = { ...range }, position = operation.index;
  if (next[first] === 0 && next[last] === maximum - 1) return next;
  if (operation.action === 'insert') {
    if (position <= next[first]) { next[first]++; next[last]++; }
    else if (position <= next[last]) next[last]++;
    if (next[last] >= maximum) fail('구조 변경이 인쇄 영역의 행·열 제한을 초과합니다.');
  } else if (position < next[first]) { next[first]--; next[last]--; }
  else if (position <= next[last]) { if (next[first] === next[last]) return null; next[last]--; }
  return next;
}

function readPageBreaks(root) {
  const result = { rows: [], columns: [], hasAny: false }, bool = value => ['1', 'true', 'on'].includes(value);
  let unsupported = false;
  for (const [axis, name, maximum, spanMaximum] of [['rows', 'rowBreaks', LIMITS.rows, LIMITS.columns], ['columns', 'colBreaks', LIMITS.columns, LIMITS.rows]]) {
    const containers = children(root, name); if (!containers.length) continue;
    result.hasAny = true;
    if (containers.length !== 1) unsupported = true;
    const node = containers[0], entries = children(node, 'brk'), seen = new Set();
    if (entries.length > 4096 || node.children.length !== entries.length || Object.keys(node.attributes).some(key => !['count', 'manualBreakCount'].includes(key))) unsupported = true;
    let manualCount = 0;
    for (const entry of entries) {
      const a = entry.attributes, id = Number(a.id), min = Number(a.min ?? 0), max = Number(a.max ?? 0), manual = bool(a.man);
      if (manual) manualCount++;
      if (Object.keys(a).some(key => !['id', 'min', 'max', 'man', 'pt'].includes(key)) || entry.children.length
        || !/^\d+$/.test(a.id || '') || !/^\d+$/.test(a.min ?? '0') || !/^\d+$/.test(a.max ?? '0')
        || id > maximum || min > max || max >= spanMaximum || seen.has(id)
        || ['man', 'pt'].some(key => a[key] !== undefined && !['0', '1', 'false', 'true', 'off', 'on'].includes(a[key])) || bool(a.pt)) unsupported = true;
      seen.add(id);
      if (manual) {
        if (min !== 0 || max !== spanMaximum - 1) unsupported = true;
        if (Number.isInteger(id) && id > 0 && id < maximum) result[axis].push(id);
      }
    }
    for (const [key, expected] of [['count', entries.length], ['manualBreakCount', manualCount]]) if (node.attributes[key] !== undefined && (!/^\d+$/.test(node.attributes[key]) || Number(node.attributes[key]) !== expected)) unsupported = true;
    result[axis] = [...new Set(result[axis])].sort((left, right) => left - right);
  }
  return { value: result, unsupported };
}

function breakRecords(xml, name) {
  const range = findElementRanges(xml, name).get('');
  if (!range) return { range: null, records: [] };
  const content = xml.slice(range.start, range.end);
  const records = [...findElementRanges(content, 'brk', undefined, true).values()].map(item => {
    const raw = content.slice(item.start, item.end), attributes = parseTagAttributes(raw.slice(0, item.startTagEnd - item.start));
    return { xml: raw, id: Number(attributes.id), manual: ['1', 'true', 'on'].includes(attributes.man), attributes };
  });
  return { range, records };
}

function updateBreakRecord(record, updates) {
  const end = record.xml.indexOf('>') + 1;
  let opening = record.xml.slice(0, end);
  for (const [key, value] of Object.entries(updates)) opening = setTagAttribute(opening, key, value);
  const attributes = { ...record.attributes, ...Object.fromEntries(Object.entries(updates).map(([key, value]) => [key, String(value)])) };
  return { xml: opening + record.xml.slice(end), id: Number(attributes.id), manual: ['1', 'true', 'on'].includes(attributes.man), attributes };
}

function replaceBreakRecords(xml, name, records) {
  const range = findElementRanges(xml, name).get('');
  let replacement = '';
  if (records.length) {
    records.sort((left, right) => left.id - right.id);
    let opening = range ? xml.slice(range.start, range.startTagEnd).replace(/\/\s*>$/, '>') : '<' + name + '>';
    opening = setTagAttribute(setTagAttribute(opening, 'count', records.length), 'manualBreakCount', records.filter(item => item.manual).length);
    replacement = opening + records.map(item => item.xml).join('') + '</' + name + '>';
  }
  if (range) return xml.slice(0, range.start) + replacement + xml.slice(range.end);
  if (!replacement) return xml;
  const following = [...(name === 'rowBreaks' ? ['colBreaks'] : []), 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];
  const points = following.flatMap(item => [...findElementRanges(xml, item).values()].map(item => item.start));
  const at = points.length ? Math.min(...points) : xml.lastIndexOf('</worksheet>');
  if (at < 0) fail('페이지 나누기를 추가할 위치가 없습니다.');
  return xml.slice(0, at) + replacement + xml.slice(at);
}

function patchPageBreaks(model, sheetIndex, operation) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  const sheet = model.sheets[sheetIndex], { axis, action, index } = operation || {};
  if (!Number.isInteger(sheetIndex) || !sheet || !['add', 'remove', 'reset'].includes(action)) fail('페이지 나누기 요청이 유효하지 않습니다.');
  let xml = sheet.xml;
  if (action === 'reset') {
    const ranges = ['rowBreaks', 'colBreaks'].flatMap(name => [...findElementRanges(xml, name, undefined, true).values()]);
    for (const range of ranges.sort((left, right) => right.start - left.start)) xml = xml.slice(0, range.start) + xml.slice(range.end);
  } else {
    const maximum = axis === 'row' ? LIMITS.rows : LIMITS.columns, spanMaximum = axis === 'row' ? LIMITS.columns : LIMITS.rows;
    if (!['row', 'column'].includes(axis) || !Number.isInteger(index) || index < 1 || index >= maximum) fail('첫 행·열 뒤의 유효한 위치에 페이지 나누기를 설정하세요.');
    if (sheet.pageBreaksUnsupported) fail('지원하지 않는 페이지 나누기는 모두 지운 뒤 새로 설정하세요.');
    const name = axis === 'row' ? 'rowBreaks' : 'colBreaks', { records } = breakRecords(xml, name);
    const manual = records.find(item => item.id === index && item.manual);
    if (action === 'add' && manual || action === 'remove' && !manual) return { bytes: model.bytes, model };
    let next;
    if (action === 'add') {
      if ((sheet.pageBreaks?.[axis === 'row' ? 'rows' : 'columns'] || []).length >= 1026) fail('수동 페이지 나누기는 축마다 1,026개 이하로 설정하세요.');
      if ((sheet.mergedCells || []).some(range => axis === 'row' ? range.top < index && index <= range.bottom : range.left < index && index <= range.right)) fail('페이지 나누기가 병합 셀을 가르지 않도록 위치를 조정하세요.');
      const cached = records.find(item => item.id === index);
      const entry = cached ? updateBreakRecord(cached, { man: '1', min: 0, max: spanMaximum - 1 }) : { id: index, manual: true, attributes: {}, xml: '<brk id="' + index + '" min="0" max="' + (spanMaximum - 1) + '" man="1"/>' };
      next = [...records.filter(item => item !== cached), entry];
    } else next = records.filter(item => item !== manual);
    xml = replaceBreakRecords(xml, name, next);
  }
  if (xml === sheet.xml) return { bytes: model.bytes, model };
  const bytes = zipSync({ ...model.files, [sheet.path]: new TextEncoder().encode(xml) }, { level: 0 }), checked = analyze(bytes, model.extension);
  if (!checked.writable || checked.sheets[sheetIndex].pageBreaksUnsupported) fail('페이지 나누기 저장 후 검증에 실패했습니다.');
  for (const [name, before] of Object.entries(model.files)) if (name !== sheet.path && Buffer.compare(Buffer.from(before), Buffer.from(checked.files[name])) !== 0) fail('페이지 나누기 저장 중 비수정 패키지 항목이 달라졌습니다: ' + name);
  return { bytes, model: checked };
}

function transformPageBreaksXml(xml, operation) {
  for (const [axis, name, maximum, spanMaximum] of [['row', 'rowBreaks', LIMITS.rows, LIMITS.columns], ['column', 'colBreaks', LIMITS.columns, LIMITS.rows]]) {
    const { range, records } = breakRecords(xml, name); if (!range) continue;
    const next = [];
    for (const record of records) {
      let updated = record;
      if (operation.axis === axis) {
        let id = record.id;
        if (operation.action === 'insert' && operation.index <= id) id++;
        else if (operation.action === 'delete' && operation.index < id) id--;
        if (id <= 0 || id >= maximum) continue;
        if (id !== record.id) updated = updateBreakRecord(record, { id });
      } else {
        let min = Number(record.attributes.min ?? 0), max = Number(record.attributes.max ?? 0);
        if (min !== 0 || max !== spanMaximum - 1) {
          if (operation.action === 'insert') { if (operation.index <= min) { min++; max++; } else if (operation.index <= max) max++; }
          else if (operation.index < min) { min--; max--; }
          else if (operation.index <= max) { if (min === max) continue; max--; }
          if (min >= spanMaximum) continue; max = Math.min(max, spanMaximum - 1);
          updated = updateBreakRecord(record, { min, max });
        }
      }
      next.push(updated);
    }
    const unique = new Map();
    for (const record of next) {
      const previous = unique.get(record.id);
      if (!previous || record.manual && !previous.manual) unique.set(record.id, record);
      else if (!previous.manual && !record.manual) unique.set(record.id, updateBreakRecord(previous, { min: Math.min(Number(previous.attributes.min ?? 0), Number(record.attributes.min ?? 0)), max: Math.max(Number(previous.attributes.max ?? 0), Number(record.attributes.max ?? 0)) }));
    }
    xml = replaceBreakRecords(xml, name, [...unique.values()]);
  }
  return xml;
}

function readHeaderFooter(root) {
  const settings = defaultHeaderFooter(), entries = children(root, 'headerFooter'), node = entries[0];
  if (!node) return { settings, unsupported: false };
  let unsupported = entries.length !== 1 || Object.keys(node.attributes).some(key => !['differentFirst', 'differentOddEven', 'scaleWithDoc', 'alignWithMargins'].includes(key));
  for (const key of ['differentFirst', 'differentOddEven', 'scaleWithDoc', 'alignWithMargins']) if (node.attributes[key] !== undefined) {
    const value = node.attributes[key];
    if (!['0', '1', 'false', 'true', 'off', 'on'].includes(value)) unsupported = true;
    else settings[key] = ['1', 'true', 'on'].includes(value);
  }
  const names = [];
  for (const variant of ['odd', 'even', 'first']) for (const part of ['header', 'footer']) {
    const name = variant + (part === 'header' ? 'Header' : 'Footer'); names.push(name);
    const values = children(node, name); if (!values.length) continue;
    if (values.length !== 1 || values[0].children.length || Object.keys(values[0].attributes).some(key => key !== 'xml:space')) unsupported = true;
    try { settings[variant][part] = parseHeaderFooterText(values[0].text); } catch { unsupported = true; }
  }
  if (node.children.some(item => !names.includes(item.name))) unsupported = true;
  try { validateHeaderFooter(settings); } catch { unsupported = true; }
  return { settings, unsupported };
}

function patchHeaderFooter(model, sheetIndex, settings, clear = false) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  const sheet = model.sheets[sheetIndex];
  if (!Number.isInteger(sheetIndex) || !sheet || typeof clear !== 'boolean') fail('머리글·바닥글 설정이 유효하지 않습니다.');
  if (clear) settings = defaultHeaderFooter();
  validateHeaderFooter(settings);
  if (sheet.headerFooterUnsupported && !clear) fail('지원하지 않는 머리글·바닥글은 지운 뒤 새로 설정하세요.');
  const current = sheet.headerFooter || defaultHeaderFooter();
  const unchanged = ['differentFirst', 'differentOddEven', 'scaleWithDoc', 'alignWithMargins'].every(key => current[key] === settings[key])
    && ['odd', 'even', 'first'].every(variant => ['header', 'footer'].every(part => ['left', 'center', 'right'].every(side => current[variant][part][side] === settings[variant][part][side])));
  if (!clear && unchanged) return { bytes: model.bytes, model };
  let xml = sheet.xml;
  const ranges = [...findElementRanges(xml, 'headerFooter', undefined, true).values()].sort((left, right) => right.start - left.start);
  let replacement = '';
  if (!clear) {
    replacement = '<headerFooter' + ['differentFirst', 'differentOddEven', 'scaleWithDoc', 'alignWithMargins'].map(key => ' ' + key + '="' + (settings[key] ? '1' : '0') + '"').join('') + '>';
    for (const variant of ['odd', 'even', 'first']) for (const part of ['header', 'footer']) {
      const value = encodeHeaderFooterText(settings[variant][part]); if (!value) continue;
      const name = variant + (part === 'header' ? 'Header' : 'Footer');
      replacement += '<' + name + ' xml:space="preserve">' + xmlEscape(value) + '</' + name + '>';
    }
    replacement += '</headerFooter>';
  }
  if (ranges.length) {
    for (const [index, range] of ranges.entries()) xml = xml.slice(0, range.start) + (index === ranges.length - 1 ? replacement : '') + xml.slice(range.end);
  } else if (replacement) {
    const following = ['rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];
    const points = following.flatMap(name => [...findElementRanges(xml, name).values()].map(range => range.start));
    const at = points.length ? Math.min(...points) : xml.lastIndexOf('</worksheet>');
    if (at < 0) fail('머리글·바닥글을 추가할 위치가 없습니다.');
    xml = xml.slice(0, at) + replacement + xml.slice(at);
  }
  if (xml === sheet.xml) return { bytes: model.bytes, model };
  const bytes = zipSync({ ...model.files, [sheet.path]: new TextEncoder().encode(xml) }, { level: 0 }), checked = analyze(bytes, model.extension);
  if (!checked.writable || checked.sheets[sheetIndex].headerFooterUnsupported) fail('머리글·바닥글 저장 후 검증에 실패했습니다.');
  for (const [name, before] of Object.entries(model.files)) if (name !== sheet.path && Buffer.compare(Buffer.from(before), Buffer.from(checked.files[name])) !== 0) fail('머리글·바닥글 저장 중 비수정 패키지 항목이 달라졌습니다: ' + name);
  return { bytes, model: checked };
}

function readPrintSettings(root) {
  const setup = child(root, 'pageSetup')?.attributes || {}, margins = child(root, 'pageMargins')?.attributes || {};
  const options = child(root, 'printOptions')?.attributes || {}, fit = child(child(root, 'sheetPr') || { children: [] }, 'pageSetUpPr')?.attributes || {};
  const number = (value, fallback) => value !== undefined && Number.isFinite(Number(value)) ? Number(value) : fallback;
  const bool = value => ['1', 'true', 'on'].includes(value);
  return {
    orientation: setup.orientation === 'landscape' ? 'landscape' : 'portrait', paperSize: number(setup.paperSize, 9),
    pageOrder: setup.pageOrder === 'overThenDown' ? 'overThenDown' : 'downThenOver', firstPageNumber: number(setup.firstPageNumber, 1), useFirstPageNumber: bool(setup.useFirstPageNumber), scale: number(setup.scale, 100), fitToPage: bool(fit.fitToPage), fitToWidth: number(setup.fitToWidth, 1), fitToHeight: number(setup.fitToHeight, 1),
    gridLines: bool(options.gridLines), headings: bool(options.headings),
    margins: Object.fromEntries(Object.entries({ left: .7, right: .7, top: .75, bottom: .75, header: .3, footer: .3 }).map(([key, fallback]) => [key, number(margins[key], fallback)]))
  };
}

function validatePrintSettings(settings) {
  if (!settings || Array.isArray(settings) || !['portrait', 'landscape'].includes(settings.orientation) || ![1, 8, 9].includes(settings.paperSize)
    || !Number.isInteger(settings.scale) || settings.scale < 10 || settings.scale > 400 || typeof settings.fitToPage !== 'boolean'
    || !['fitToWidth', 'fitToHeight'].every(key => Number.isInteger(settings[key]) && settings[key] >= 0 && settings[key] <= 10)
    || !['gridLines', 'headings'].every(key => typeof settings[key] === 'boolean') || !settings.margins || Array.isArray(settings.margins)
    || !['left', 'right', 'top', 'bottom', 'header', 'footer'].every(key => Number.isFinite(settings.margins[key]) && settings.margins[key] >= 0 && settings.margins[key] <= 5)) fail('페이지 설정 값이 유효하지 않습니다.');
  if (settings.firstPageNumber !== undefined && (!Number.isInteger(settings.firstPageNumber) || settings.firstPageNumber < 0 || settings.firstPageNumber > 32767) || settings.useFirstPageNumber !== undefined && typeof settings.useFirstPageNumber !== 'boolean') fail('페이지 설정 값이 유효하지 않습니다.');
  if (settings.pageOrder !== undefined && !['downThenOver', 'overThenDown'].includes(settings.pageOrder)) fail('페이지 설정 값이 유효하지 않습니다.');
  const sizes = { 1: [8.5, 11], 8: [11.6929, 16.5354], 9: [8.2677, 11.6929] };
  const dimensions = sizes[settings.paperSize], [width, height] = settings.orientation === 'landscape' ? [...dimensions].reverse() : dimensions;
  if (settings.margins.left + settings.margins.right > width - 1 || settings.margins.top + settings.margins.bottom > height - 1) fail('여백이 용지의 인쇄 가능 영역을 초과합니다.');
  return settings;
}

function patchPrintSettings(model, sheetIndex, settings) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  const sheet = model.sheets[sheetIndex];
  if (!Number.isInteger(sheetIndex) || !sheet) fail('페이지 설정의 시트가 유효하지 않습니다.');
  validatePrintSettings(settings);
  let xml = sheet.xml;
  const upsert = (name, attributes, following, parent = 'worksheet') => {
    const existing = [...findElementRanges(xml, name).values()];
    if (existing.length > 1) fail('페이지 설정 요소가 중복되었습니다.');
    if (existing.length) {
      const range = existing[0], tag = xml.slice(range.start, range.startTagEnd);
      const next = Object.entries(attributes).reduce((value, [key, setting]) => setTagAttribute(value, key, setting), tag);
      xml = xml.slice(0, range.start) + next + xml.slice(range.startTagEnd); return;
    }
    const container = findElementRanges(xml, parent).get('');
    if (!container) fail('페이지 설정의 상위 요소가 없습니다.');
    const tag = '<' + name + Object.entries(attributes).map(([key, value]) => ' ' + key + '="' + xmlEscape(value) + '"').join('') + '/>';
    if (container.selfClosing) {
      xml = xml.slice(0, container.start) + xml.slice(container.start, container.startTagEnd).replace(/\/\s*>$/, '>') + tag + '</' + parent + '>' + xml.slice(container.end); return;
    }
    const points = following.flatMap(item => [...findElementRanges(xml, item).values()].map(range => range.start));
    const at = points.length ? Math.min(...points) : container.closingStart;
    xml = xml.slice(0, at) + tag + xml.slice(at);
  };
  if (!findElementRanges(xml, 'sheetPr').size) upsert('sheetPr', {}, ['dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData']);
  upsert('pageSetUpPr', { fitToPage: settings.fitToPage ? '1' : '0' }, [], 'sheetPr');
  const afterOptions = ['pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];
  upsert('printOptions', { gridLines: settings.gridLines ? '1' : '0', headings: settings.headings ? '1' : '0' }, afterOptions);
  upsert('pageMargins', Object.fromEntries(['left', 'right', 'top', 'bottom', 'header', 'footer'].map(key => [key, settings.margins[key]])), afterOptions.slice(1));
  upsert('pageSetup', { orientation: settings.orientation, paperSize: settings.paperSize, pageOrder: settings.pageOrder || 'downThenOver', scale: settings.scale, fitToWidth: settings.fitToWidth, fitToHeight: settings.fitToHeight, firstPageNumber: settings.firstPageNumber ?? 1, useFirstPageNumber: settings.useFirstPageNumber ? '1' : '0' }, afterOptions.slice(2));
  if (xml === sheet.xml) return { bytes: model.bytes, model };
  const bytes = zipSync({ ...model.files, [sheet.path]: new TextEncoder().encode(xml) }, { level: 0 }), checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('페이지 설정 저장 후 패키지 검사에서 거절되었습니다.');
  for (const [name, before] of Object.entries(model.files)) if (name !== sheet.path && Buffer.compare(Buffer.from(before), Buffer.from(checked.files[name])) !== 0) fail('페이지 설정 중 비수정 패키지 항목이 달라졌습니다: ' + name);
  return { bytes, model: checked };
}

function patchRowHeightValues(model, sheetIndex, values) {
  if (!Array.isArray(values) || values.length > 10000 || values.some((item, index) => !item || !Number.isInteger(item.row) || item.row < 0 || item.row >= LIMITS.rows || !Number.isFinite(item.height) || item.height < 0.75 || item.height > 409.5 || index > 0 && item.row <= values[index - 1].row)) fail('행 자동 맞춤 결과가 유효하지 않습니다.');
  if (!values.length) return { bytes: model.bytes, model };
  return patchRowHeights(model, sheetIndex, values[0].row, values.at(-1).row, 1, new Map(values.map(item => [item.row, item.height])));
}

function patchRowHeights(model, sheetIndex, startRow, endRow, height, rowValues = null) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  const sheet = model.sheets[sheetIndex];
  if (!Number.isInteger(sheetIndex) || !sheet || !Number.isInteger(startRow) || !Number.isInteger(endRow) || startRow < 0 || endRow < startRow || endRow >= LIMITS.rows || height !== null && (!Number.isFinite(height) || height < 0.75 || height > 409.5)) fail('행 높이는 0.75~409.5pt 범위여야 하며 행 범위가 유효해야 합니다.');
  const fullSheet = !rowValues && startRow === 0 && endRow === LIMITS.rows - 1;
  if (!rowValues && !fullSheet && endRow - startRow + 1 > 10000) fail('행 높이는 한 번에 10,000행 이하 또는 전체 행에 적용하세요.');
  let xml = sheet.xml;
  if (fullSheet && height !== null) {
    const formatRange = findElementRanges(xml, 'sheetFormatPr').get('');
    if (formatRange) {
      const opening = setTagAttribute(xml.slice(formatRange.start, formatRange.startTagEnd), 'defaultRowHeight', height);
      xml = xml.slice(0, formatRange.start) + opening + xml.slice(formatRange.startTagEnd);
    } else {
      const position = Math.min(...['cols', 'sheetData'].flatMap(name => [...findElementRanges(xml, name).values()].map(range => range.start)));
      if (!Number.isFinite(position)) fail('기본 행 높이를 삽입할 위치를 찾을 수 없습니다.');
      xml = xml.slice(0, position) + '<sheetFormatPr defaultRowHeight="' + height + '"/>' + xml.slice(position);
    }
  }
  const rowRanges = findElementRanges(xml, 'row', 'r'), operations = [], insertions = new Map();
  const existing = [...rowRanges].map(([number, range]) => ({ row: Number(number) - 1, range })).sort((left, right) => left.row - right.row);
  const reset = height === null || fullSheet;
  for (const { row, range } of existing) {
    if (row < startRow || row > endRow || rowValues && !rowValues.has(row)) continue;
    const tag = xml.slice(range.start, range.startTagEnd), attributes = parseTagAttributes(tag);
    let next = tag;
    if (attributes.ht !== undefined && Number(attributes.ht) === 0) next = setTagAttribute(next, 'hidden', '1');
    if (reset) next = next.replace(/\s+ht\s*=\s*(["'])[^"']*\1/i, '').replace(/\s+customHeight\s*=\s*(["'])[^"']*\1/i, '');
    else next = setTagAttribute(setTagAttribute(next, 'ht', rowValues?.get(row) ?? height), 'customHeight', rowValues ? '0' : '1');
    if (next !== tag) operations.push({ start: range.start, end: range.startTagEnd, value: next });
  }
  if (!reset) {
    const data = findElementRanges(xml, 'sheetData').get('');
    if (!data) fail('행 높이를 적용할 시트 데이터가 없습니다.');
    const additions = [];
    const targets = rowValues ? [...rowValues.keys()] : Array.from({ length: endRow - startRow + 1 }, (_, index) => startRow + index);
    for (const row of targets) {
      if (rowRanges.has(String(row + 1))) continue;
      const tag = '<row r="' + (row + 1) + '" ht="' + (rowValues?.get(row) ?? height) + '" customHeight="' + (rowValues ? '0' : '1') + '"/>';
      if (data.selfClosing) { additions.push(tag); continue; }
      let low = 0, high = existing.length;
      while (low < high) { const middle = (low + high) >> 1; if (existing[middle].row <= row) low = middle + 1; else high = middle; }
      const position = existing[low]?.range.start ?? data.closingStart;
      if (!insertions.has(position)) insertions.set(position, []);
      insertions.get(position).push(tag);
    }
    if (additions.length) operations.push({ start: data.start, end: data.end, value: xml.slice(data.start, data.startTagEnd).replace(/\/\s*>$/, '>') + additions.join('') + '</sheetData>' });
    for (const [position, values] of insertions) operations.push({ start: position, end: position, value: values.join('') });
  }
  for (const operation of operations.sort((left, right) => right.start - left.start)) xml = xml.slice(0, operation.start) + operation.value + xml.slice(operation.end);
  if (xml === sheet.xml) return { bytes: model.bytes, model };
  const outputFiles = { ...model.files, [sheet.path]: new TextEncoder().encode(xml) }, modifiedPaths = new Set([sheet.path]);
  updateDimensionDrawings(model, outputFiles, modifiedPaths);
  const bytes = zipSync(outputFiles, { level: 0 }), checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('행 높이 저장 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  for (const [name, before] of Object.entries(model.files)) if (!modifiedPaths.has(name) && Buffer.compare(Buffer.from(before), Buffer.from(checked.files[name])) !== 0) fail('행 높이 저장 중 비수정 패키지 항목이 달라졌습니다: ' + name);
  return { bytes, model: checked };
}

function patchHiddenColumns(model, edits, updateManual = true) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  if (!Array.isArray(edits) || edits.length > LIMITS.sheets) fail('숨김 열 수정 요청이 유효하지 않습니다.');
  if (!edits.length) return { bytes: model.bytes, model };
  const outputFiles = Object.assign({}, model.files), modifiedPaths = new Set(), seen = new Set();
  for (const edit of edits) {
    const { sheetIndex, columns, hidden } = edit || {}, sheet = model.sheets[sheetIndex];
    if (!Number.isInteger(sheetIndex) || !sheet || seen.has(sheetIndex) || !Array.isArray(columns) || columns.length > LIMITS.columns || typeof hidden !== 'boolean' || columns.some((column, index) => !Number.isInteger(column) || column < 0 || column >= LIMITS.columns || (index && column <= columns[index - 1]))) fail('숨김 열 수정 요청이 유효하지 않습니다.');
    seen.add(sheetIndex);
    if (!columns.length) continue;
    const ranges = [...findElementRanges(sheet.xml, 'col', 'min').entries()].map(([key, range]) => {
      const tag = sheet.xml.slice(range.start, range.startTagEnd), attributes = parseTagAttributes(tag);
      const min = Number(attributes.min), max = Number(attributes.max);
      if (!range.selfClosing || !Number.isInteger(min) || !Number.isInteger(max) || min < 1 || max < min) fail('열 정의를 안전하게 수정할 수 없습니다.');
      return { range, tag, attributes, min, max };
    }).sort((a, b) => a.min - b.min);
    for (let index = 1; index < ranges.length; index++) if (ranges[index].min <= ranges[index - 1].max) fail('겹치는 열 정의를 안전하게 수정할 수 없습니다.');
    const operations = [], insertions = [];
    for (const item of ranges) {
      const zeroWidth = item.attributes.width !== undefined && excelColumnWidthPixels(Number(item.attributes.width)) === 0 || item.attributes.width === undefined && sheet.defaultColumnWidth === 0;
      const wasHidden = ['1', 'true', 'on'].includes(String(item.attributes.hidden || '').toLowerCase()) || zeroWidth;
      const targets = columns.filter(column => column + 1 >= item.min && column + 1 <= item.max && hidden !== wasHidden);
      if (!targets.length) continue;
      const pieces = [], selected = new Set(targets.map(column => column + 1));
      let start = item.min;
      while (start <= item.max) {
        const isTarget = selected.has(start), hiddenValue = isTarget ? hidden : ['1', 'true', 'on'].includes(String(item.attributes.hidden || '').toLowerCase());
        let end = start;
        while (end + 1 <= item.max && selected.has(end + 1) === isTarget) end++;
        let tag = setTagAttribute(item.tag, 'min', start); tag = setTagAttribute(tag, 'max', end);
        if (hiddenValue) tag = setTagAttribute(tag, 'hidden', '1'); else {
          tag = removeTagAttribute(tag, 'hidden');
          if (isTarget && zeroWidth) {
            tag = removeTagAttribute(removeTagAttribute(tag, 'width'), 'customWidth');
            if (sheet.defaultColumnWidth === 0) tag = setTagAttribute(setTagAttribute(tag, 'width', '9.140625'), 'customWidth', '1');
          }
        }
        if (updateManual && sheet.columnVisibilityPrefix && isTarget) tag = setTagAttribute(tag, sheet.columnVisibilityPrefix + ':manualHidden', hidden ? '1' : '0');
        pieces.push(tag); start = end + 1;
      }
      operations.push({ start: item.range.start, end: item.range.end, replacement: pieces.join('') });
    }
    for (const column of columns) {
      const coordinate = column + 1, covered = ranges.some(item => coordinate >= item.min && coordinate <= item.max);
      const alreadyHidden = (sheet.hiddenColumns || []).includes(column);
      if (covered || hidden === alreadyHidden) continue;
      if (hidden) insertions.push(`<col min="${coordinate}" max="${coordinate}" hidden="1"${updateManual && sheet.columnVisibilityPrefix ? ' ' + sheet.columnVisibilityPrefix + ':manualHidden="1"' : ''}/>`);
      else if (sheet.defaultColumnWidth === 0) insertions.push(`<col min="${coordinate}" max="${coordinate}" width="9.140625" customWidth="1"${updateManual && sheet.columnVisibilityPrefix ? ' ' + sheet.columnVisibilityPrefix + ':manualHidden="0"' : ''}/>`);
    }
    if (!operations.length && !insertions.length) continue;
    const cols = findElementRanges(sheet.xml, 'cols').get('');
    if (insertions.length) {
      const tags = insertions.join('');
      if (cols?.selfClosing) operations.push({ start: cols.start, end: cols.end, replacement: sheet.xml.slice(cols.start, cols.startTagEnd).replace(/\/\s*>$/, '>') + tags + '</cols>' });
      else if (cols) operations.push({ start: cols.closingStart, end: cols.closingStart, replacement: tags });
      else {
        const sheetData = findElementRanges(sheet.xml, 'sheetData').get('');
        if (!sheetData || sheetData.selfClosing) fail('숨김 열을 워크시트에 안전하게 추가할 수 없습니다.');
        operations.push({ start: sheetData.start, end: sheetData.start, replacement: '<cols>' + tags + '</cols>' });
      }
    }
    operations.sort((a, b) => b.start - a.start || b.end - a.end);
    let xml = sheet.xml, previousStart = xml.length + 1;
    for (const operation of operations) { if (operation.end > previousStart) fail('숨김 열 수정 범위가 겹칩니다.'); xml = xml.slice(0, operation.start) + operation.replacement + xml.slice(operation.end); previousStart = operation.start; }
    outputFiles[sheet.path] = new TextEncoder().encode(xml); modifiedPaths.add(sheet.path);
  }
  if (!modifiedPaths.size) return { bytes: model.bytes, model };
  updateDimensionDrawings(model, outputFiles, modifiedPaths);
  const bytes = zipSync(outputFiles, { level: 0 }), checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('숨김 열 저장 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  for (const [name, before] of Object.entries(model.files)) if (!modifiedPaths.has(name) && (!checked.files[name] || Buffer.compare(Buffer.from(before), Buffer.from(checked.files[name])) !== 0)) fail('숨김 열 저장 중 비수정 패키지 항목이 달라졌습니다: ' + name);
  return { bytes, model: checked };
}

function patchFreezePanes(model, edits) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  if (!Array.isArray(edits) || edits.length > LIMITS.sheets) fail('고정 창 수정 요청이 유효하지 않습니다.');
  if (!edits.length) return { bytes: model.bytes, model };
  const outputFiles = Object.assign({}, model.files), modifiedPaths = new Set(), seen = new Set();
  for (const edit of edits) {
    const { sheetIndex, freezePane } = edit || {}, sheet = model.sheets[sheetIndex], rows = freezePane?.rows, columns = freezePane?.columns;
    if (!Number.isInteger(sheetIndex) || !sheet || seen.has(sheetIndex) || !Number.isInteger(rows) || rows < 0 || rows > 50 || !Number.isInteger(columns) || columns < 0 || columns > 10) fail('고정 창 수정 요청이 유효하지 않습니다.');
    seen.add(sheetIndex);
    const freeze = rows > 0 || columns > 0;
    const pane = freeze ? '<pane' + (columns ? ' xSplit="' + columns + '"' : '') + (rows ? ' ySplit="' + rows + '"' : '') + ' topLeftCell="' + columnLetters(columns + 1) + (rows + 1) + '" activePane="' + (rows && columns ? 'bottomRight' : columns ? 'topRight' : 'bottomLeft') + '" state="frozen"/>' : '';
    const selection = freeze ? '<selection pane="' + (rows && columns ? 'bottomRight' : columns ? 'topRight' : 'bottomLeft') + '" activeCell="' + columnLetters(columns + 1) + (rows + 1) + '" sqref="' + columnLetters(columns + 1) + (rows + 1) + '"/>' : '';
    if (JSON.stringify(sheet.freezePane || { rows: 0, columns: 0 }) === JSON.stringify({ rows, columns })) continue;
    const viewRanges = [...findElementRanges(sheet.xml, 'sheetView').values()];
    if (viewRanges.length > 1) fail('여러 사용자 정의 시트 보기를 안전하게 변경할 수 없습니다.');
    let xml = sheet.xml;
    if (viewRanges.length) {
      const range = viewRanges[0], original = xml.slice(range.start, range.end);
      const paneRanges = [...findElementRanges(original, 'pane').values()];
      if (paneRanges.length > 1) fail('여러 분할 창을 안전하게 변경할 수 없습니다.');
      let view = original.replace(/<selection\b[^>]*(?:\/>|>[\s\S]*?<\/selection>)/g, '');
      const panes = [...findElementRanges(view, 'pane').values()];
      if (panes.length) {
        const old = panes[0];
        view = view.slice(0, old.start) + pane + view.slice(old.end);
      } else if (freeze) {
        const opening = /^<sheetView\b[^>]*?(?:\/?>)/.exec(view);
        if (!opening) fail('시트 보기 태그를 안전하게 수정할 수 없습니다.');
        if (/\/\s*>$/.test(opening[0])) view = view.slice(0, opening.index) + opening[0].replace(/\/\s*>$/, '>') + pane + selection + '</sheetView>' + view.slice(opening.index + opening[0].length);
        else view = view.slice(0, opening.index + opening[0].length) + pane + selection + view.slice(opening.index + opening[0].length);
      }
      if (freeze) {
        const currentPane = [...findElementRanges(view, 'pane').values()][0];
        view = view.slice(0, currentPane.end) + selection + view.slice(currentPane.end);
      }
      xml = xml.slice(0, range.start) + view + xml.slice(range.end);
    } else if (freeze) {
      const views = findElementRanges(xml, 'sheetViews').get('');
      const viewXml = '<sheetView workbookViewId="0">' + pane + selection + '</sheetView>';
      if (views?.selfClosing) xml = xml.slice(0, views.start) + '<sheetViews>' + viewXml + '</sheetViews>' + xml.slice(views.end);
      else if (views) xml = xml.slice(0, views.closingStart) + viewXml + xml.slice(views.closingStart);
      else {
        const following = ['sheetFormatPr', 'cols', 'sheetData'];
        const points = following.flatMap(name => [...findElementRanges(xml, name).values()].map(range => range.start));
        const at = points.length ? Math.min(...points) : xml.indexOf('</worksheet>');
        if (at < 0) fail('고정 창을 워크시트에 안전하게 추가할 수 없습니다.');
        xml = xml.slice(0, at) + '<sheetViews>' + viewXml + '</sheetViews>' + xml.slice(at);
      }
    } else continue;
    outputFiles[sheet.path] = new TextEncoder().encode(xml); modifiedPaths.add(sheet.path);
  }
  if (!modifiedPaths.size) return { bytes: model.bytes, model };
  const bytes = zipSync(outputFiles, { level: 0 }), checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('고정 창 저장 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  for (const [name, before] of Object.entries(model.files)) if (!modifiedPaths.has(name) && (!checked.files[name] || Buffer.compare(Buffer.from(before), Buffer.from(checked.files[name])) !== 0)) fail('고정 창 저장 중 비수정 패키지 항목이 달라졌습니다: ' + name);
  return { bytes, model: checked };
}

function patchMergedCells(model, edits) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  if (!Array.isArray(edits) || edits.length > LIMITS.sheets) fail('병합 셀 수정 요청이 유효하지 않습니다.');
  if (model.hasCharts && edits.some(edit => edit?.ranges?.some(range => model.chartSourceRanges.some(source => source.sheetIndex === edit.sheetIndex && range.top <= source.bottom && range.bottom >= source.top && range.left <= source.right && range.right >= source.left)))) fail('차트 원본 셀이 포함된 범위는 병합할 수 없습니다.');
  if (!edits.length) return { bytes: model.bytes, model };
  const outputFiles = Object.assign({}, model.files), modifiedPaths = new Set(), seen = new Set();
  for (const edit of edits) {
    const { sheetIndex, ranges } = edit || {}, sheet = model.sheets[sheetIndex];
    if (!Number.isInteger(sheetIndex) || !sheet || seen.has(sheetIndex) || !Array.isArray(ranges) || ranges.length > 1000 || ranges.some(range => !range || !Number.isInteger(range.top) || !Number.isInteger(range.bottom) || !Number.isInteger(range.left) || !Number.isInteger(range.right) || range.top < 0 || range.bottom < range.top || range.left < 0 || range.right < range.left || range.top === range.bottom && range.left === range.right || range.bottom >= LIMITS.rows || range.right >= LIMITS.columns)) fail('병합 셀 범위가 유효하지 않습니다.');
    if (ranges.some(range => sheet.hyperlinks?.some(link => range.top <= link.bottom && range.bottom >= link.top && range.left <= link.right && range.right >= link.left))) fail('하이퍼링크가 있는 범위는 병합할 수 없습니다.');
    seen.add(sheetIndex);
    const sorted = [...ranges].sort((left, right) => left.top - right.top || left.left - right.left || left.bottom - right.bottom || left.right - right.right);
    for (let index = 0; index < sorted.length; index++) for (let other = 0; other < index; other++) {
      const range = sorted[index], previous = sorted[other];
      if (range.top <= previous.bottom && range.bottom >= previous.top && range.left <= previous.right && range.right >= previous.left) fail('병합 셀 범위가 서로 겹칩니다.');
    }
    const reference = range => columnLetters(range.left + 1) + (range.top + 1) + ':' + columnLetters(range.right + 1) + (range.bottom + 1);
    const tags = sorted.map(range => '<mergeCell ref="' + reference(range) + '"/>').join('');
    const replacement = sorted.length ? '<mergeCells count="' + sorted.length + '">' + tags + '</mergeCells>' : '';
    const existing = findElementRanges(sheet.xml, 'mergeCells').get('');
    let xml;
    if (existing) xml = sheet.xml.slice(0, existing.start) + replacement + sheet.xml.slice(existing.end);
    else if (replacement) {
      const afterNames = ['sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews'];
      const beforeNames = ['phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];
      const after = afterNames.flatMap(name => [...findElementRanges(sheet.xml, name).values()].map(range => range.end));
      const before = beforeNames.flatMap(name => [...findElementRanges(sheet.xml, name).values()].map(range => range.start));
      const anchor = after.length ? Math.max(...after) : findElementRanges(sheet.xml, 'sheetData').get('')?.end;
      const next = before.length ? Math.min(...before) : sheet.xml.indexOf('</worksheet>');
      if (!Number.isInteger(anchor) || anchor < 0 || next < 0 || anchor > next) fail('병합 셀을 워크시트 XML에 안전하게 추가할 수 없습니다.');
      xml = sheet.xml.slice(0, anchor) + replacement + sheet.xml.slice(anchor);
    } else continue;
    if (xml === sheet.xml) continue;
    outputFiles[sheet.path] = new TextEncoder().encode(xml); modifiedPaths.add(sheet.path);
  }
  if (!modifiedPaths.size) return { bytes: model.bytes, model };
  const bytes = zipSync(outputFiles, { level: 0 }), checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('병합 셀 저장 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  for (const [name, before] of Object.entries(model.files)) if (!modifiedPaths.has(name) && (!checked.files[name] || Buffer.compare(Buffer.from(before), Buffer.from(checked.files[name])) !== 0)) fail('병합 셀 저장 중 비수정 패키지 항목이 달라졌습니다: ' + name);
  return { bytes, model: checked };
}

function tableColumnMatches(xml) {
  return [...findElementRanges(xml, 'tableColumn', undefined, true).values()].map(range => ({ 0: xml.slice(range.start, range.end), index: range.start }));
}

function resizeTableColumnDefinitions(xml, columns, operation = null) {
  const container = findElementRanges(xml, 'tableColumns').get('');
  if (!container || container.selfClosing || !Array.isArray(columns) || !columns.length) fail('Excel 표 열 정의를 안전하게 변경할 수 없습니다.');
  const opening = xml.slice(container.start, container.startTagEnd), content = xml.slice(container.startTagEnd, container.closingStart);
  const matches = tableColumnMatches(content), tags = matches.map(match => match[0]);
  const prefix = content.slice(0, matches[0]?.index || 0), suffix = content.slice(matches.at(-1)?.index + matches.at(-1)?.[0].length || 0);
  if (operation?.action === 'insert') {
    if (!Number.isInteger(operation.index) || operation.index < 0 || operation.index > tags.length || columns.length !== tags.length + 1) fail('Excel 표 열 삽입 위치가 유효하지 않습니다.');
    const ids = tags.map(tag => Number(parseTagAttributes(tag).id)).filter(Number.isInteger), id = Math.max(0, ...ids) + 1;
    tags.splice(operation.index, 0, '<tableColumn id="' + id + '" name="' + xmlEscape(columns[operation.index]).replaceAll('\r','&#13;').replaceAll('\n','&#10;').replaceAll('\t','&#9;') + '"/>');
  } else if (operation?.action === 'delete') {
    if (!Number.isInteger(operation.index) || operation.index < 0 || operation.index >= tags.length || columns.length !== tags.length - 1) fail('Excel 표 열 삭제 위치가 유효하지 않습니다.');
    tags.splice(operation.index, 1);
  } else if (operation?.reconcile) {
    const byName = new Map(tags.map(tag => [parseTagAttributes(tag).name?.toLocaleLowerCase(), tag]));
    const ids = tags.map(tag => Number(parseTagAttributes(tag).id)).filter(Number.isInteger);
    let nextId = Math.max(0, ...ids);
    tags.splice(0, tags.length, ...columns.map(name => byName.get(String(name).toLocaleLowerCase()) || '<tableColumn id="' + (++nextId) + '" name="' + xmlEscape(name).replaceAll('\r','&#13;').replaceAll('\n','&#10;').replaceAll('\t','&#9;') + '"/>'));
  } else if (tags.length > columns.length) {
    tags.splice(columns.length);
  }
  else {
    for (let index = tags.length; index < columns.length; index++) {
      const ids = tags.map(tag => Number(parseTagAttributes(tag).id)).filter(Number.isInteger), id = Math.max(0, ...ids) + 1;
      tags.push('<tableColumn id="' + id + '" name="' + xmlEscape(columns[index]).replaceAll('\r','&#13;').replaceAll('\n','&#10;').replaceAll('\t','&#9;') + '"/>');
    }
  }
  if (tags.length !== columns.length) fail('Excel 표 열 정의 개수가 맞지 않습니다.');
  const nextContent = prefix + tags.join('') + suffix;
  return xml.slice(0, container.start) + replaceTagAttribute(opening, 'count', columns.length) + nextContent + xml.slice(container.closingStart);
}

function transformTableFilterColumns(xml, action, columnIndex) {
  const range = findElementRanges(xml, 'autoFilter').get('');
  if (!range || range.selfClosing) return xml;
  const element = xml.slice(range.start, range.end), body = element.slice(element.indexOf('>') + 1, element.lastIndexOf('</autoFilter>'));
  const updatedBody = body.replace(/<filterColumn\b[^>]*(?:\/\s*>|>[\s\S]*?<\/filterColumn>)/g, tag => {
    const attributes = parseTagAttributes(tag), index = Number(attributes.colId);
    if (action === 'delete' && index === columnIndex) return '';
    const openingEnd = tag.indexOf('>') + 1;
    if (action === 'insert' && index >= columnIndex) return replaceTagAttribute(tag.slice(0, openingEnd), 'colId', index + 1) + tag.slice(openingEnd);
    if (action === 'delete' && index > columnIndex) return replaceTagAttribute(tag.slice(0, openingEnd), 'colId', index - 1) + tag.slice(openingEnd);
    return tag;
  });
  return xml.slice(0, range.start) + element.slice(0, element.indexOf('>') + 1) + updatedBody + '</autoFilter>' + xml.slice(range.end);
}

function transformTableSortStateForStructure(xml, sheetName, operation, table) {
  const stateName = /<((?:[\w.-]+:)?sortState)\b/.exec(xml)?.[1];
  if (!stateName) return xml;
  const stateRange = findElementRanges(xml, stateName).get('');
  if (!stateRange) fail('표 정렬 상태 범위를 안전하게 찾을 수 없습니다.');
  const stateXml = xml.slice(stateRange.start, stateRange.end), stateOpening = stateXml.slice(0, stateRange.startTagEnd - stateRange.start);
  const stateAttributes = parseTagAttributes(stateOpening), patches = [];
  if (stateAttributes.ref) {
    const nextRef = transformFormulaForStructure(stateAttributes.ref, sheetName, sheetName, operation);
    if (nextRef.includes('#REF!')) fail('행·열 변경 후 표 정렬 범위가 유효하지 않습니다.');
    if (nextRef !== stateAttributes.ref) patches.push({ start: 0, end: stateRange.startTagEnd - stateRange.start, replacement: setTagAttribute(stateOpening, 'ref', nextRef) });
  }
  const conditionName = /<((?:[\w.-]+:)?sortCondition)\b/.exec(stateXml)?.[1];
  const conditions = conditionName ? [...findElementRanges(stateXml, conditionName, undefined, true).values()] : [];
  let retainedConditions = 0;
  for (const condition of conditions) {
    const opening = stateXml.slice(condition.start, condition.startTagEnd), attributes = parseTagAttributes(opening);
    if (!attributes.ref) fail('표 정렬 조건에 참조 범위가 없습니다.');
    let nextRef = transformFormulaForStructure(attributes.ref, sheetName, sheetName, operation);
    if (operation.axis === 'row' && operation.action === 'insert') {
      const range = parseRangeReference(attributes.ref, sheetName, LIMITS.rows * LIMITS.columns);
      const dataTop = table.range.top + (table.headerRowCount || 0), dataBottom = table.range.bottom - (table.totalsRowCount || 0);
      if (range && (operation.index === dataTop && range.top === dataTop || table.totalsRowCount && operation.index === dataBottom + 1 && range.bottom === dataBottom)) {
        const expanded = { ...range, bottom: range.bottom + 1 };
        const first = columnLetters(expanded.left + 1) + (expanded.top + 1), last = columnLetters(expanded.right + 1) + (expanded.bottom + 1);
        nextRef = first === last ? first : first + ':' + last;
      }
    }
    if (nextRef.includes('#REF!')) {
      patches.push({ start: condition.start, end: condition.end, replacement: '' });
      continue;
    }
    retainedConditions++;
    if (nextRef !== attributes.ref) patches.push({ start: condition.start, end: condition.startTagEnd, replacement: setTagAttribute(opening, 'ref', nextRef) });
  }
  if (conditions.length && retainedConditions === 0) return xml.slice(0, stateRange.start) + xml.slice(stateRange.end);
  let updatedState = stateXml;
  for (const patch of patches.sort((left, right) => right.start - left.start)) updatedState = updatedState.slice(0, patch.start) + patch.replacement + updatedState.slice(patch.end);
  return xml.slice(0, stateRange.start) + updatedState + xml.slice(stateRange.end);
}

function resizeTableSortState(xml, sheetName, table, nextRange) {
  const stateName = /<((?:[\w.-]+:)?sortState)\b/.exec(xml)?.[1];
  if (!stateName) return xml;
  const stateRange = findElementRanges(xml, stateName).get('');
  if (!stateRange) fail('표 정렬 상태 범위를 안전하게 찾을 수 없습니다.');
  const stateXml = xml.slice(stateRange.start, stateRange.end), stateOpening = stateXml.slice(0, stateRange.startTagEnd - stateRange.start);
  const attributes = parseTagAttributes(stateOpening), patches = [];
  const oldDataBottom = table.range.bottom - (table.totalsRowCount || 0), nextDataBottom = nextRange.bottom - (table.totalsRowCount || 0);
  const formatRange = range => {
    const first = columnLetters(range.left + 1) + (range.top + 1), last = columnLetters(range.right + 1) + (range.bottom + 1);
    return first === last ? first : first + ':' + last;
  };
  const fitRange = (reference, body) => {
    const range = parseRangeReference(reference, sheetName, LIMITS.rows * LIMITS.columns);
    if (!range) fail('표 정렬 참조를 안전하게 변경할 수 없습니다.');
    const next = { ...range };
    if (range.top === table.range.top) next.top = nextRange.top;
    if (range.bottom === (body ? oldDataBottom : table.range.bottom)) next.bottom = body ? nextDataBottom : nextRange.bottom;
    if (!body && range.left === table.range.left) next.left = nextRange.left;
    if (!body && range.right === table.range.right) next.right = nextRange.right;
    next.top = Math.max(next.top, body ? table.range.top + (table.headerRowCount || 0) : nextRange.top);
    next.bottom = Math.min(next.bottom, body ? nextDataBottom : nextRange.bottom);
    next.left = Math.max(next.left, nextRange.left);
    next.right = Math.min(next.right, nextRange.right);
    return next.top <= next.bottom && next.left <= next.right ? next : null;
  };
  if (attributes.ref) {
    const next = fitRange(attributes.ref, false);
    if (!next) fail('표 정렬 범위가 새 표 크기와 겹치지 않습니다.');
    const reference = formatRange(next);
    if (reference !== attributes.ref) patches.push({ start: 0, end: stateRange.startTagEnd - stateRange.start, replacement: setTagAttribute(stateOpening, 'ref', reference) });
  }
  const conditionName = /<((?:[\w.-]+:)?sortCondition)\b/.exec(stateXml)?.[1];
  const conditions = conditionName ? [...findElementRanges(stateXml, conditionName, undefined, true).values()] : [];
  let retained = 0;
  for (const condition of conditions) {
    const opening = stateXml.slice(condition.start, condition.startTagEnd), conditionAttributes = parseTagAttributes(opening);
    if (!conditionAttributes.ref) fail('표 정렬 조건에 참조 범위가 없습니다.');
    const next = fitRange(conditionAttributes.ref, true);
    if (!next) { patches.push({ start: condition.start, end: condition.end, replacement: '' }); continue; }
    retained++;
    const reference = formatRange(next);
    if (reference !== conditionAttributes.ref) patches.push({ start: condition.start, end: condition.startTagEnd, replacement: setTagAttribute(opening, 'ref', reference) });
  }
  if (conditions.length && !retained) return xml.slice(0, stateRange.start) + xml.slice(stateRange.end);
  let updatedState = stateXml;
  for (const patch of patches.sort((left, right) => right.start - left.start)) updatedState = updatedState.slice(0, patch.start) + patch.replacement + updatedState.slice(patch.end);
  return xml.slice(0, stateRange.start) + updatedState + xml.slice(stateRange.end);
}

function remapTableFilterColumns(xml, oldColumns, newColumns) {
  const range = findElementRanges(xml, 'autoFilter').get('');
  if (!range || range.selfClosing) return xml;
  const newIndexes = new Map(newColumns.map((name, index) => [name.toLocaleLowerCase(), index]));
  const element = xml.slice(range.start, range.end), body = element.slice(element.indexOf('>') + 1, element.lastIndexOf('</autoFilter>'));
  const updatedBody = body.replace(/<filterColumn\b[^>]*(?:\/\s*>|>[\s\S]*?<\/filterColumn>)/g, tag => {
    const index = Number(parseTagAttributes(tag).colId), name = oldColumns[index]?.toLocaleLowerCase(), nextIndex = newIndexes.get(name);
    if (!name || nextIndex === undefined) return '';
    const openingEnd = tag.indexOf('>') + 1;
    return nextIndex === index ? tag : replaceTagAttribute(tag.slice(0, openingEnd), 'colId', nextIndex) + tag.slice(openingEnd);
  });
  return xml.slice(0, range.start) + element.slice(0, element.indexOf('>') + 1) + updatedBody + '</autoFilter>' + xml.slice(range.end);
}

function validateTotalsMove(model, sheetIndex, source, target, options = {}) {
  const sheet = model.sheets[sheetIndex];
  if (model.sheets.some(item => [...item.cells.values()].some(cell => cell.formulaType && cell.formulaType !== 'normal'))) fail('공유·배열 수식이 포함된 통합문서는 합계 행을 이동할 수 없습니다.');
  const calculator = createFormulaCalculator({ state: model, getCell: (row, column, index = 0) => model.sheets[index]?.cells.get(columnLetters(column + 1) + (row + 1)) || null });
  const intersects = (left, right) => left.top <= right.bottom && left.bottom >= right.top && left.left <= right.right && left.right >= right.left;
  const attached = [...(sheet.mergedCells || []), ...(sheet.hyperlinks || []), ...(sheet.dataValidations || []).flatMap(item => item.ranges), ...(sheet.conditionalFormats || []).flatMap(item => item.ranges)];
  if (attached.some(range => intersects(range, source) || intersects(range, target))) fail('합계 이동 범위의 병합·유효성·조건부 서식·하이퍼링크를 먼저 해제하세요.');
  if (options.allowOccupied) return;
  for (let column = target.left; column <= target.right; column++) {
    const cell = sheet.cells.get(columnLetters(column + 1) + (target.top + 1));
    if (cell && (cell.formula || cell.value !== '') || calculator.cell(sheetIndex, target.top, column)?.isSpill) fail('이동할 합계 행 위치에 데이터가 있습니다. 빈 행을 선택하거나 데이터 행 삽입을 사용하세요.');
  }
}

function moveTableTotals(model, original, moves) {
  const remap = (formula, ownerSheet) => moves.reduce((value, move) => transformFormulaForStructure(value, ownerSheet, original.sheets[move.sheetIndex].name, move), formula);
  const edits = new Map(), put = edit => edits.set(edit.sheetIndex + ':' + edit.address, edit);
  for (const [sheetIndex, sheet] of model.sheets.entries()) for (const cell of sheet.cells.values()) {
    if (!cell.formula) continue;
    const formula = remap(cell.formula, sheet.name);
    if (formula !== cell.formula) put({ sheetIndex, address: cell.address, value: cell.value, formula });
  }
  for (const move of moves) {
    const sheet = original.sheets[move.sheetIndex];
    const emptyTable = sheet.tables.find(table => table.totalsRowCount && table.range.bottom === move.source.top && table.range.left <= move.source.right && table.range.right >= move.source.left && table.range.bottom - table.totalsRowCount < table.range.top + table.headerRowCount);
    for (let column = move.source.left; column <= move.source.right; column++) {
      const address = columnLetters(column + 1) + (move.source.top + 1), cell = sheet.cells.get(address);
      const dataStyle = emptyTable ? null : sheet.cells.get(columnLetters(column + 1) + move.source.top)?.style ?? null;
      if (cell) put({ sheetIndex: move.sheetIndex, address, value: '', kind: 'blank', formula: null, style: dataStyle });
      if (cell) put({ sheetIndex: move.sheetIndex, address: columnLetters(column + 1) + (move.source.top + move.rowDelta + 1), value: cell.value, kind: cell.kind, formula: cell.formula ? remap(cell.formula, sheet.name) : null, style: cell.style });
    }
  }
  let current = edits.size ? patchCells(model, [...edits.values()]).model : model;
  const files = { ...current.files };
  let changed = false;
  const rewriteElements = (xml, names, ownerSheet, partName) => xml.replace(new RegExp('<(' + names.join('|') + ')(\\b[^>]*>)([\\s\\S]*?)<\\/\\1>', 'g'), (whole, name, opening) => {
    const node = parseXml(new TextEncoder().encode(whole), partName).root;
    if (node.children.length) fail('복잡한 참조 수식 구성은 합계 이동을 지원하지 않습니다.');
    const owner = name === 'definedName' && node.attributes.localSheetId !== undefined ? current.sheets[Number(node.attributes.localSheetId)]?.name || ownerSheet : ownerSheet;
    const formula = remap(node.text, owner);
    return formula === node.text ? whole : '<' + name + opening + xmlEscape(formula) + '</' + name + '>';
  });
  const update = (partName, xml) => {
    if (xml !== new TextDecoder().decode(files[partName])) { files[partName] = new TextEncoder().encode(xml); changed = true; }
  };
  update('xl/workbook.xml', rewriteElements(new TextDecoder().decode(files['xl/workbook.xml']), ['definedName'], '', 'xl/workbook.xml'));
  for (const sheet of current.sheets) {
    let xml = rewriteElements(sheet.xml, ['formula', 'formula1', 'formula2'], sheet.name, sheet.path);
    xml = xml.replace(/<hyperlink\b[^>]*>/g, tag => {
      const location = parseTagAttributes(tag).location;
      if (!location) return tag;
      const next = remap(location, sheet.name);
      return next === location ? tag : setTagAttribute(tag, 'location', next);
    });
    update(sheet.path, xml);
    for (const table of sheet.tables || []) {
      const owner = original.sheets.find(item => item.name === sheet.name)?.tables.find(item => item.name === table.name);
      const emptyPrototype = owner?.totalsRowCount && owner.range.bottom - owner.totalsRowCount < owner.range.top + owner.headerRowCount && moves.some(move => original.sheets[move.sheetIndex].name === sheet.name && move.source.top === owner.range.bottom && move.source.left <= owner.range.right && move.source.right >= owner.range.left);
      // An empty table's calculated formula is a prototype for the future data
      // row, not a formula located in the totals cell that is being moved.
      update(table.path, rewriteElements(new TextDecoder().decode(files[table.path]), emptyPrototype ? ['totalsRowFormula'] : ['calculatedColumnFormula', 'totalsRowFormula'], sheet.name, table.path));
    }
  }
  if (changed) {
    current = analyze(zipSync(files, { level: 0 }), model.extension);
    if (!current.writable) fail('합계 행 참조 이동 후 패키지 검사에서 거절되었습니다: ' + current.readOnlyReason);
  }
  return patchFilterVisibility(current);
}

const TOTALS_FUNCTIONS = Object.freeze({ average: 101, countNums: 102, count: 103, max: 104, min: 105, stdDev: 107, sum: 109, var: 110 });

function totalsFormula(table, column, fn) {
  const name = table.columns[column - table.range.left].replace(/['\[\]#@]/g, value => "'" + value);
  return 'SUBTOTAL(' + TOTALS_FUNCTIONS[fn] + ',[[#Data],[' + name + ']])';
}

function totalsFunctionForCell(table, column, cell) {
  if (!cell?.formula) return 'none';
  const index = column - table.range.left, fn = table.totalColumns[index].function, code = TOTALS_FUNCTIONS[fn];
  if (!code) return 'custom';
  const name = table.columns[index].replace(/['\[\]#@]/g, value => "'" + value);
  const expected = [totalsFormula(table, column, fn), 'SUBTOTAL(' + code + ',[' + name + '])', 'SUBTOTAL(' + code + ',' + table.name + '[' + name + '])'];
  return expected.some(formula => formula.toLocaleLowerCase() === cell.formula.trim().toLocaleLowerCase()) ? fn : 'custom';
}

function replaceTotalsColumn(xml, index, totals) {
  const tags = tableColumnMatches(xml), match = tags[index];
  if (!match) fail('합계 열 정의를 찾을 수 없습니다.');
  const openingEnd = match[0].indexOf('>') + 1, attributes = parseTagAttributes(match[0].slice(0, openingEnd));
  delete attributes.totalsRowFunction; delete attributes.totalsRowLabel;
  attributes.totalsRowFunction = totals.function;
  if (totals.label !== null && totals.label !== undefined) attributes.totalsRowLabel = totals.label;
  const attribute = value => xmlEscape(value).replaceAll('\r', '&#13;').replaceAll('\n', '&#10;').replaceAll('\t', '&#9;');
  const opening = '<tableColumn' + Object.entries(attributes).map(([key, value]) => ' ' + key + '="' + attribute(value) + '"').join('');
  let body = /\/\s*>$/.test(match[0]) ? '' : match[0].slice(openingEnd, match[0].lastIndexOf('</tableColumn>'));
  body = body.replace(/<totalsRowFormula\b[^>]*(?:\/\s*>|>[\s\S]*?<\/totalsRowFormula>)/g, '');
  if (totals.formula) {
    const position = body.search(/<(?:xmlColumnPr|extLst)\b/), formula = '<totalsRowFormula>' + xmlEscape(totals.formula) + '</totalsRowFormula>';
    body = position < 0 ? body + formula : body.slice(0, position) + formula + body.slice(position);
  }
  const replacement = opening + (body ? '>' + body + '</tableColumn>' : '/>');
  return xml.slice(0, match.index) + replacement + xml.slice(match.index + match[0].length);
}

function patchTableTotals(model, request) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  const { sheetIndex, tableName, enabled, column, function: fn, label = 'Total' } = request || {}, sheet = model.sheets[sheetIndex], table = sheet?.tables?.find(item => item.name === tableName);
  if (!Number.isInteger(sheetIndex) || !table || enabled !== undefined && typeof enabled !== 'boolean' || typeof label !== 'string' || label.length > 32767 || enabled === undefined && fn === undefined || fn !== undefined && (!Object.hasOwn(TOTALS_FUNCTIONS, fn) && fn !== 'none' || !Number.isInteger(column) || column < table.range.left || column > table.range.right)) fail('표 합계 요청이 유효하지 않습니다.');
  const active = Boolean(table.totalsRowCount), nextActive = enabled ?? active;
  if (!nextActive && fn !== undefined) fail('집계 함수를 선택하려면 합계 행을 먼저 켜세요.');
  if (enabled === active && fn === undefined) return { bytes: model.bytes, model };
  if (!table.headerRowCount) fail('표 합계에는 표시된 머리글이 필요합니다.');
  const bottom = table.range.bottom + (nextActive ? active ? 0 : 1 : -1);
  if (bottom >= LIMITS.rows || bottom < table.range.top + table.headerRowCount - 1 || sheet.tables.some(other => other !== table && other.range.top <= bottom && other.range.bottom >= table.range.top && other.range.left <= table.range.right && other.range.right >= table.range.left)) fail('합계 행 위치가 시트 범위를 벗어나거나 다른 표와 겹칩니다.');
  const area = { top: nextActive ? bottom : table.range.bottom, bottom: nextActive ? bottom : table.range.bottom, left: table.range.left, right: table.range.right };
  validateTotalsMove(model, sheetIndex, area, area, { allowOccupied: active });
  const columns = table.totalColumns.map(item => ({ ...item })), cellEdits = [];
  if (!active && !columns.some(item => item.configured || item.function !== 'none' || item.label !== null || item.formula)) {
    if (columns.length > 1) columns[0].label = label;
    columns.at(-1).function = 'sum';
  }
  if (active && !nextActive) for (let index = 0; index < columns.length; index++) {
    const cell = sheet.cells.get(columnLetters(table.range.left + index + 1) + (table.range.bottom + 1)), item = columns[index];
    if (cell?.formula) {
      if (totalsFunctionForCell(table, table.range.left + index, cell) === 'custom') Object.assign(item, { function: 'custom', label: null, formula: cell.formula });
    } else if (cell?.value !== '' && cell?.value !== undefined) Object.assign(item, cell.kind === 'string' ? { function: 'none', label: cell.value, formula: null } : { function: 'custom', label: null, formula: cell.kind === 'boolean' ? cell.value.toUpperCase() + '()' : cell.value });
    else Object.assign(item, { function: 'none', label: null, formula: null });
  }
  if (fn !== undefined) Object.assign(columns[column - table.range.left], { function: fn, label: null, formula: null });
  let tableXml = new TextDecoder().decode(model.files[table.path]);
  const changedColumns = active && nextActive && fn !== undefined ? [column - table.range.left] : columns.map((_, index) => index);
  for (const index of changedColumns) {
    const item = columns[index], address = columnLetters(table.range.left + index + 1) + ((nextActive ? bottom : table.range.bottom) + 1), existing = sheet.cells.get(address);
    let formula = null, value = '', kind = 'blank';
    if (nextActive) {
      if (Object.hasOwn(TOTALS_FUNCTIONS, item.function)) formula = totalsFormula(table, table.range.left + index, item.function);
      else if (item.function === 'custom' && item.formula) formula = item.formula;
      else if (item.function !== 'none') fail('지원하지 않는 합계 함수입니다.');
      else if (item.label !== null) { value = item.label; kind = 'string'; }
    }
    const style = existing?.style ?? sheet.cells.get(columnLetters(table.range.left + index + 1) + table.range.bottom)?.style ?? null;
    cellEdits.push({ sheetIndex, address, value, kind, formula, style });
    tableXml = replaceTotalsColumn(tableXml, index, item);
  }
  const root = findElementRanges(tableXml, 'table').get(''), opening = tableXml.slice(root.start, root.startTagEnd), reference = columnLetters(table.range.left + 1) + (table.range.top + 1) + ':' + columnLetters(table.range.right + 1) + (bottom + 1);
  const updated = setTagAttribute(setTagAttribute(setTagAttribute(opening, 'ref', reference), 'totalsRowCount', nextActive ? '1' : '0'), 'totalsRowShown', nextActive ? '1' : '0');
  tableXml = tableXml.slice(0, root.start) + updated + tableXml.slice(root.startTagEnd);
  const files = { ...model.files, [table.path]: new TextEncoder().encode(tableXml) }, prepared = analyze(zipSync(files, { level: 0 }), model.extension);
  if (!prepared.writable) fail('표 합계 저장 검사에서 거절되었습니다: ' + prepared.readOnlyReason);
  const patched = patchCells(prepared, cellEdits);
  return patchFilterVisibility(patched.model);
}

function patchCalculatedColumns(model, edits) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  if (!Array.isArray(edits) || edits.length > LIMITS.cells) fail('계산 열 요청이 유효하지 않습니다.');
  const files = { ...model.files }, seen = new Set(); let changed = false;
  for (const edit of edits) {
    const { sheetIndex, tableName, column, formula } = edit || {}, table = model.sheets[sheetIndex]?.tables?.find(item => item.name === tableName), index = column - (table?.range.left ?? 0);
    if (!Number.isInteger(sheetIndex) || !table || !Number.isInteger(column) || index < 0 || index >= table.columns.length || formula !== null && (typeof formula !== 'string' || !formula.length || formula.length > 100000 || formula.startsWith('=')) || seen.has(sheetIndex + ':' + tableName + ':' + column)) fail('계산 열 요청이 유효하지 않습니다.');
    if (table.calculatedColumnUnsupported?.[index]) fail('공유·배열 형식의 계산 열은 변경할 수 없습니다.');
    seen.add(sheetIndex + ':' + tableName + ':' + column);
    const xml = new TextDecoder().decode(files[table.path]), match = tableColumnMatches(xml)[index], tag = match[0], end = tag.indexOf('>') + 1;
    let body = /\/\s*>$/.test(tag) ? '' : tag.slice(end, tag.lastIndexOf('</tableColumn>'));
    const ranges = [...findElementRanges(body, 'calculatedColumnFormula', undefined, true).values()];
    if (ranges.length > 1) fail('중복된 계산 열 수식을 변경할 수 없습니다.');
    if (ranges.length) body = body.slice(0, ranges[0].start) + body.slice(ranges[0].end);
    if (formula !== null) body = '<calculatedColumnFormula>' + xmlEscape(formula) + '</calculatedColumnFormula>' + body;
    const opening = tag.slice(0, end).replace(/\s*\/\s*>$/, '>'), replacement = body ? opening + body + '</tableColumn>' : opening.replace(/>$/, '/>');
    const updated = xml.slice(0, match.index) + replacement + xml.slice(match.index + tag.length);
    if (updated !== xml) { files[table.path] = new TextEncoder().encode(updated); changed = true; }
  }
  if (!changed) return { bytes: model.bytes, model };
  const bytes = zipSync(files, { level: 0 }), checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('계산 열 저장 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  return { bytes, model: checked };
}

function planCalculatedColumn(model, request, getCell = (row, column, sheetIndex) => model.sheets[sheetIndex]?.cells.get(columnLetters(column + 1) + (row + 1)) || null) {
  const { sheetIndex, row, column, formula, mode = 'preserve', forceOrigin = true } = request || {}, sheet = model.sheets[sheetIndex];
  if (!Number.isInteger(sheetIndex) || !sheet || !Number.isInteger(row) || !Number.isInteger(column) || !['preserve', 'all'].includes(mode) || typeof formula !== 'string' || !formula.length || formula.length > 100000 || formula.startsWith('=')) fail('계산 열 요청이 유효하지 않습니다.');
  const table = sheet.tables?.find(item => row >= item.range.top + item.headerRowCount && row <= item.range.bottom - item.totalsRowCount && column >= item.range.left && column <= item.range.right);
  if (!table) return null;
  const first = table.range.top + table.headerRowCount, last = table.range.bottom - table.totalsRowCount, index = column - table.range.left;
  if (last - first + 1 > LIMITS.cells) fail('계산 열은 한 번에 100,000셀 이하를 적용하세요.');
  if (table.calculatedColumnUnsupported?.[index]) fail('공유·배열 형식의 계산 열은 변경할 수 없습니다.');
  const previous = table.calculatedColumns?.[index] || null, canonical = translateFormulaReferences(formula, first - row, 0), edits = [];
  let preserved = 0;
  for (let target = first; target <= last; target++) {
    const cell = getCell(target, column, sheetIndex), empty = !cell?.hasFormula && !cell?.formula && (cell?.value ?? '') === '';
    const matches = previous && cell?.formula && translateFormulaReferences(cell.formula, first - target, 0) === translateFormulaReferences(previous, 0, 0);
    if (mode === 'all' || target === row && forceOrigin || matches || empty && !previous) edits.push({ row: target, column, value: String(cell?.value ?? ''), formula: translateFormulaReferences(canonical, target - first, 0) });
    else preserved++;
  }
  return { edits, metadata: { sheetIndex, tableName: table.name, column, formula: canonical }, preserved };
}

function fillCalculatedRows(model, requests, options = {}) {
  const edits = [], seen = new Set(), calculator = createFormulaCalculator({ state: model, getCell: (row, column, index = 0) => model.sheets[index]?.cells.get(columnLetters(column + 1) + (row + 1)) || null });
  for (const request of requests) {
    const sheet = model.sheets[request.sheetIndex], table = sheet?.tables?.find(item => item.name === request.tableName);
    if (!table) continue;
    const first = table.range.top + table.headerRowCount, last = table.range.bottom - table.totalsRowCount;
    for (const [index, formula] of (table.calculatedColumns || []).entries()) {
      if (!formula) continue;
      const column = table.range.left + index;
      if (table.calculatedColumnUnsupported?.[index]) fail('공유·배열 형식의 계산 열은 변경할 수 없습니다.');
      if (Math.min(request.bottom, last) - Math.max(request.top, first) + 1 > LIMITS.cells) fail('계산 열은 한 번에 100,000셀 이하를 적용하세요.');
      const sample = [...sheet.cells.values()].find(cell => cell.column === column && cell.row >= first && cell.row <= last);
      for (let row = Math.max(request.top, first); row <= Math.min(request.bottom, last); row++) {
        const address = columnLetters(column + 1) + (row + 1), key = request.sheetIndex + ':' + address, cell = sheet.cells.get(address);
        if (options.skipCells?.has(key) || seen.has(key) || cell?.formula || cell?.hasFormula || (cell?.value ?? '') !== '') continue;
        if (calculator.cell(request.sheetIndex, row, column)?.isSpill) fail('동적 배열 결과의 일부 셀은 편집할 수 없습니다. 배열 수식이 있는 셀을 편집하세요.');
        if (sheet.mergedCells.some(range => row >= range.top && row <= range.bottom && column >= range.left && column <= range.right) || sheet.dataValidations.some(rule => rule.ranges.some(range => row >= range.top && row <= range.bottom && column >= range.left && column <= range.right))) fail('계산 열의 새 셀에 병합 또는 유효성 설정이 있어 자동 채우기를 중단했습니다.');
        seen.add(key);
        const neighbor = (row > first ? sheet.cells.get(columnLetters(column + 1) + row) : null) || (row < last ? sheet.cells.get(columnLetters(column + 1) + (row + 2)) : null) || sample;
        edits.push({ sheetIndex: request.sheetIndex, address, value: '', formula: translateFormulaReferences(formula, row - first, 0), style: cell ? cell.style : neighbor?.style ?? null });
        if (edits.length > LIMITS.cells) fail('계산 열은 한 번에 100,000셀 이하를 적용하세요.');
      }
    }
  }
  return edits.length ? patchCells(model, edits) : { bytes: model.bytes, model };
}

function planTableHeaderEdits(model, sheetIndex, edits) {
  const requests = [], replacements = new Map();
  for (const table of model.sheets[sheetIndex]?.tables || []) {
    if (table.headerRowCount !== 1) continue;
    const incoming = edits.filter(edit => edit && edit.row === table.range.top && edit.column >= table.range.left && edit.column <= table.range.right && typeof edit.value === 'string' && (typeof edit.formula === 'string' || edit.value !== (model.sheets[sheetIndex].cells.get(columnLetters(edit.column + 1) + (edit.row + 1))?.value ?? '')));
    if (!incoming.length) continue;
    const indexes = new Set(incoming.map(edit => edit.column - table.range.left));
    const names = new Set(table.columns.filter((_, index) => !indexes.has(index)).map(name => name.toLocaleLowerCase()));
    const columns = [...table.columns];
    for (const edit of incoming) {
      const index = edit.column - table.range.left, base = (typeof edit.formula === 'string' ? '=' + edit.formula : edit.value).trim() || 'Column' + (index + 1);
      if (base.length > 255) fail('표 머리글은 255자 이하로 입력하세요.');
      let name = base, suffix = 2;
      while (names.has(name.toLocaleLowerCase())) { const tail = String(suffix++); name = base.slice(0, 255 - tail.length) + tail; }
      names.add(name.toLocaleLowerCase()); columns[index] = name;
      replacements.set(edit.row + ':' + edit.column, { ...edit, value: name, kind: 'string', formula: null });
    }
    if (columns.some((name, index) => name !== table.columns[index])) requests.push({ sheetIndex, tableName: table.name, columns });
  }
  return { requests, edits: edits.map(edit => replacements.get(edit?.row + ':' + edit?.column) || edit) };
}

// Rewrite bracket leaves only, preserving strings, selectors and other tables.
function renameStructuredColumns(formula, renames, ownerTable = null) {
  let output = '', index = 0;
  const escape = name => name.replace(/['\[\]#@]/g, value => "'" + value);
  while (index < formula.length) {
    const rest = formula.slice(index), quoted = /^(?:"(?:[^"]|"")*"|'(?:[^']|'')*'!)/.exec(rest);
    if (quoted) { output += quoted[0]; index += quoted[0].length; continue; }
    const prefix = /^[\p{L}_][\p{L}\p{N}_.]*/u.exec(rest)?.[0] || '';
    const start = index + prefix.length;
    if (formula[start] !== '[') { const text = prefix || formula[index]; output += text; index += text.length; continue; }
    let end = start, depth = 0;
    for (; end < formula.length; end++) {
      if (formula[end] === "'" && /['\[\]#@]/.test(formula[end + 1] || '')) { end++; continue; }
      if (formula[end] === '[') depth++;
      else if (formula[end] === ']' && --depth === 0) { end++; break; }
    }
    const spec = formula.slice(start, end), mapping = renames.get((prefix || ownerTable || '').toLocaleLowerCase());
    const external = !prefix && /^[\p{L}\p{N}_.]+!/u.test(formula.slice(end));
    let next = spec;
    if (!depth && mapping && !external) {
      const patches = [], stack = [];
      for (let offset = 0; offset < spec.length; offset++) {
        if (spec[offset] === "'" && /['\[\]#@]/.test(spec[offset + 1] || '')) { offset++; continue; }
        if (spec[offset] === '[') { if (stack.length) stack.at(-1).nested = true; stack.push({ start: offset + 1, nested: false }); }
        else if (spec[offset] === ']') {
          const entry = stack.pop(); if (!entry || entry.nested) continue;
          const raw = spec.slice(entry.start, offset), at = raw.startsWith('@') ? '@' : '', text = raw.slice(at.length);
          if (text.startsWith('#')) continue;
          const name = mapping.get(text.replace(/'(['\[\]#@])/g, '$1').toLocaleLowerCase());
          if (name !== undefined) patches.push({ start: entry.start, end: offset, value: at + escape(name) });
        }
      }
      for (const patch of patches.sort((a, b) => b.start - a.start)) next = next.slice(0, patch.start) + patch.value + next.slice(patch.end);
    }
    output += prefix + next; index = end;
  }
  return output;
}

function patchTableColumnNames(model, requests) {
  if (!model.writable) fail(model.readOnlyReason || '읽기 전용 통합문서입니다.');
  const renames = new Map(), files = { ...model.files };
  for (const request of requests) {
    const table = model.sheets[request.sheetIndex]?.tables.find(item => item.name === request.tableName);
    if (!table || request.columns.length !== table.columns.length || request.columns.some(name => typeof name !== 'string' || !name || name.length > 255) || new Set(request.columns.map(name => name.toLocaleLowerCase())).size !== request.columns.length) fail('표 열 이름 변경 요청이 유효하지 않습니다.');
    renames.set(table.name.toLocaleLowerCase(), new Map(table.columns.flatMap((name, index) => name === request.columns[index] ? [] : [[name.toLocaleLowerCase(), request.columns[index]]])));
    let xml = new TextDecoder().decode(files[table.path]);
    for (const [index, match] of [...tableColumnMatches(xml).entries()].reverse()) {
      const end = match[0].indexOf('>') + 1, value = xmlEscape(request.columns[index]).replaceAll('\r','&#13;').replaceAll('\n','&#10;').replaceAll('\t','&#9;');
      const tag = match[0].slice(0, end).replace(/\bname\s*=\s*(?:"[^"]*"|'[^']*')/, () => 'name="' + value + '"') + match[0].slice(end);
      xml = xml.slice(0, match.index) + tag + xml.slice(match.index + match[0].length);
    }
    files[table.path] = new TextEncoder().encode(xml);
  }
  const rewrite = (xml, names, ownerTable = null) => xml.replace(new RegExp('<(' + names.join('|') + ')(\\b[^>]*?(?<!/)>)([\\s\\S]*?)<\\/\\1>', 'g'), (whole, tag, opening) => {
    const node = parseXml(new TextEncoder().encode(whole), 'table-rename').root;
    const formula = renameStructuredColumns(node.text, renames, ownerTable);
    return formula === node.text ? whole : '<' + tag + opening + xmlEscape(formula) + '</' + tag + '>';
  });
  const edits = [];
  for (const [sheetIndex, sheet] of model.sheets.entries()) {
    for (const cell of sheet.cells.values()) {
      if (!cell.formula) continue;
      const table = sheet.tables.find(item => cell.row >= item.range.top && cell.row <= item.range.bottom && cell.column >= item.range.left && cell.column <= item.range.right);
      const formula = renameStructuredColumns(cell.formula, renames, table?.name);
      if (formula !== cell.formula) {
        if (cell.formulaType && cell.formulaType !== 'normal') fail('공유·배열 수식의 표 머리글 참조 변경은 아직 지원하지 않습니다.');
        edits.push({ sheetIndex, address: cell.address, value: cell.value, formula });
      }
    }
    files[sheet.path] = new TextEncoder().encode(rewrite(sheet.xml, ['formula', 'formula1', 'formula2']));
    for (const table of sheet.tables) files[table.path] = new TextEncoder().encode(rewrite(new TextDecoder().decode(files[table.path]), ['calculatedColumnFormula', 'totalsRowFormula'], table.name));
  }
  files['xl/workbook.xml'] = new TextEncoder().encode(rewrite(new TextDecoder().decode(files['xl/workbook.xml']), ['definedName']));
  for (const part of Object.keys(files).filter(name => /^xl\/charts\/.*\.xml$/.test(name))) files[part] = new TextEncoder().encode(rewrite(new TextDecoder().decode(files[part]), ['c:f', 'f']));
  let current = analyze(zipSync(files, { level: 0 }), model.extension);
  if (!current.writable) fail('표 머리글 변경 후 통합문서 검사에서 거절되었습니다: ' + current.readOnlyReason);
  if (edits.length) current = patchCells(current, edits).model;
  return { bytes: current.bytes, model: current };
}

function planTableExpansion(model, sheetIndex, edits, getCell = (row, column, index) => model.sheets[index]?.cells.get(columnLetters(column + 1) + (row + 1)) || null) {
  const sheet = model.sheets[sheetIndex];
  if (!sheet?.tables?.length || model.hasCharts) return { requests: [], edits };
  const valid = edits.filter(edit => Number.isInteger(edit?.row) && Number.isInteger(edit?.column) && edit.row >= 0 && edit.row < LIMITS.rows && edit.column >= 0 && edit.column < LIMITS.columns && typeof edit.value === 'string');
  const inside = (row, column, range) => row >= range.top && row <= range.bottom && column >= range.left && column <= range.right;
  const meaningful = valid.filter(edit => {
    if (sheet.tables.some(table => inside(edit.row, edit.column, table.range))) return false;
    const cell = getCell(edit.row, edit.column, sheetIndex);
    return typeof edit.formula === 'string' ? edit.formula !== cell?.formula : edit.value !== '' && (edit.value !== (cell?.value ?? '') || edit.formula === null && cell?.hasFormula);
  });
  const contiguous = (set, first, last) => { if (last < first || last - first + 1 > set.size) return false; for (let index = first; index <= last; index++) if (!set.has(index)) return false; return true; };
  const proposals = [];
  for (const table of sheet.tables) {
    const range = { ...table.range }, headers = [];
    const rightEdits = table.headerRowCount === 1 ? meaningful.filter(edit => edit.row >= range.top && edit.row <= range.bottom && edit.column > range.right) : [];
    const newRight = rightEdits.length ? rightEdits.reduce((max, edit) => Math.max(max, edit.column), range.right) : range.right;
    const columns = new Set(valid.filter(edit => edit.row >= range.top && edit.row <= range.bottom && edit.column > range.right).map(edit => edit.column));
    if (newRight > range.right && contiguous(columns, range.right + 1, newRight)) {
      const names = new Set(table.columns.map(name => name.toLocaleLowerCase())); let supported = true;
      for (let column = range.right + 1; column <= newRight; column++) {
        const incoming = valid.find(edit => edit.row === range.top && edit.column === column), cell = getCell(range.top, column, sheetIndex), formula = incoming && Object.hasOwn(incoming, 'formula') ? incoming.formula : cell?.formula;
        if (typeof formula === 'string') { supported = false; break; }
        const base = String(incoming?.value ?? cell?.value ?? '').trim() || 'Column' + (column - range.left + 1);
        let name = base, suffix = 2; while (names.has(name.toLocaleLowerCase())) name = base + suffix++;
        names.add(name.toLocaleLowerCase()); headers.push({ row: range.top, column, value: name });
      }
      if (supported) range.right = newRight; else headers.length = 0;
    }
    if (!table.totalsRowCount) {
      const below = meaningful.filter(edit => edit.row > range.bottom && edit.column >= range.left && edit.column <= range.right), bottom = below.length ? below.reduce((max, edit) => Math.max(max, edit.row), range.bottom) : range.bottom;
      const rows = new Set(valid.filter(edit => edit.row > range.bottom && edit.column >= range.left && edit.column <= range.right).map(edit => edit.row));
      if (bottom > range.bottom && contiguous(rows, range.bottom + 1, bottom)) range.bottom = bottom;
    }
    if (range.right !== table.range.right || range.bottom !== table.range.bottom) proposals.push({ sheetIndex, tableName: table.name, range, headers });
  }
  const overlaps = (left, right) => left.top <= right.bottom && left.bottom >= right.top && left.left <= right.right && left.right >= right.left;
  const requests = proposals.filter(proposal => !sheet.mergedCells.some(range => overlaps(proposal.range, range)) && !sheet.tables.some(table => table.name !== proposal.tableName && overlaps(proposal.range, table.range)) && !proposals.some(other => other !== proposal && overlaps(proposal.range, other.range)));
  const headerNames = new Map(requests.flatMap(request => request.headers).map(header => [header.row + ':' + header.column, header.value]));
  return { requests, edits: edits.map(edit => edit && headerNames.has(edit.row + ':' + edit.column) ? { ...edit, value: headerNames.get(edit.row + ':' + edit.column), kind: 'string', formula: null } : edit) };
}

function prepareTableExpansion(model, requests) {
  const seeds = new Map();
  for (const request of requests) {
    const sheet = model.sheets[request.sheetIndex], table = sheet.tables.find(item => item.name === request.tableName), old = table.range, first = old.top + table.headerRowCount, last = old.bottom - table.totalsRowCount, nextLast = request.range.bottom - table.totalsRowCount;
    if ((request.range.bottom - request.range.top + 1) * (request.range.right - request.range.left + 1) > LIMITS.cells) fail('자동 표 확장은 100,000셀 이하 범위를 사용하세요.');
    for (const header of request.headers) {
      const address = columnLetters(header.column + 1) + (header.row + 1), existing = sheet.cells.get(address), neighbor = sheet.cells.get(columnLetters(old.right + 1) + (old.top + 1));
      seeds.set(request.sheetIndex + ':' + address, { sheetIndex: request.sheetIndex, address, value: header.value, kind: 'string', formula: null, style: existing?.style ?? neighbor?.style ?? null });
    }
    for (let row = first; row <= nextLast; row++) for (let column = old.left; column <= request.range.right; column++) {
      if (row <= last && column <= old.right) continue;
      const address = columnLetters(column + 1) + (row + 1), cell = sheet.cells.get(address);
      if (cell?.hasFormula || cell?.formula || (cell?.value ?? '') !== '' || cell?.style !== null && cell?.style !== undefined) continue;
      const source = last >= first ? sheet.cells.get(columnLetters(Math.min(column, old.right) + 1) + (Math.min(row, last) + 1)) : null;
      if (source?.style !== null && source?.style !== undefined) seeds.set(request.sheetIndex + ':' + address, { sheetIndex: request.sheetIndex, address, value: '', formula: null, kind: 'blank', style: source.style });
    }
  }
  return seeds.size ? patchCells(model, [...seeds.values()]).model : model;
}

function patchAddTable(model, request, options = {}) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  const { sheetIndex, range } = request || {}, sheet = model.sheets[sheetIndex];
  if (!Number.isInteger(sheetIndex) || !sheet || !range || !['top','bottom','left','right'].every(key => Number.isInteger(range[key])) || range.top < 0 || range.left < 0 || range.bottom < range.top || range.right < range.left || range.bottom >= LIMITS.rows || range.right >= LIMITS.columns || (range.bottom-range.top+1)*(range.right-range.left+1) > LIMITS.cells) fail('표를 만들려면 첫 행을 머리글로 사용할 100,000셀 이하 범위를 선택하세요.');
  const tables = model.sheets.flatMap(item => item.tables), names = new Set([...tables.map(table => table.name), ...model.definedNames.map(item => item.name)].map(name => name.toLocaleLowerCase()));
  if (tables.length >= 256) fail('통합문서에는 256개 이하의 표를 만들 수 있습니다.');
  if (request.name !== undefined && typeof request.name !== 'string') fail('표 이름이 유효하지 않거나 이미 사용 중입니다.');
  let name = request.name?.trim() || '';
  if (!name) { let index = 1; while (names.has(('Table'+index).toLocaleLowerCase())) index++; name = 'Table'+index; }
  if (name.length > 255 || !/^[\p{L}_\\][\p{L}\p{N}_.]*$/u.test(name) || /^(?:r|c|r\d+c\d+)$/i.test(name) || /^[A-Z]{1,3}[1-9]\d*$/i.test(name) && columnNumber(name.toUpperCase()) <= LIMITS.columns && rowNumber(name.toUpperCase()) <= LIMITS.rows || names.has(name.toLocaleLowerCase())) fail('표 이름이 유효하지 않거나 이미 사용 중입니다.');
  const overlaps = other => range.top <= other.bottom && range.bottom >= other.top && range.left <= other.right && range.right >= other.left;
  if (sheet.tables.some(table => overlaps(table.range)) || sheet.mergedCells.some(overlaps)) fail('기존 표 또는 병합 셀과 겹치는 범위에는 표를 만들 수 없습니다.');
  const migrateFilter = sheet.autoFilter && ['top','bottom','left','right'].every(key => range[key] === sheet.autoFilter[key]);
  if (sheet.autoFilter && overlaps(sheet.autoFilter) && !migrateFilter) fail('기존 자동 필터와 겹칩니다. 같은 범위를 선택하거나 필터를 먼저 해제하세요.');
  const calculator = createFormulaCalculator({ state: model, now: options.now, getCell: (row, column, index=0) => model.sheets[index]?.cells.get(columnLetters(column+1)+(row+1)) || null });
  if (calculator.spillRanges(sheetIndex).some(overlaps)) fail('동적 배열 결과와 겹치는 범위에는 표를 만들 수 없습니다.');
  const headers = [], headerNames = new Set(), edits = [];
  for (let column=range.left; column<=range.right; column++) {
    const address=columnLetters(column+1)+(range.top+1), cell=sheet.cells.get(address);
    if (cell?.formula || cell?.hasFormula) fail('표 머리글에는 수식 대신 값을 사용하세요.');
    const base=String(cell ? formatCellValue(cell, options.language || 'en') : '').replace(/\r\n?/g,'\n').trim() || 'Column'+(column-range.left+1);
    if (base.length > 255) fail('표 머리글은 255자 이하로 입력하세요.');
    let header=base, suffix=2;
    while (headerNames.has(header.toLocaleLowerCase())) { const ending=String(suffix++); let prefix=base.slice(0,255-ending.length); if (/[\uD800-\uDBFF]$/.test(prefix)) prefix=prefix.slice(0,-1); header=prefix+ending; }
    headerNames.add(header.toLocaleLowerCase()); headers.push(header);
    if (cell?.kind !== 'string' || cell.value !== header) edits.push({ sheetIndex, address, value: header, kind: 'string', formula: null });
  }
  model = preserveImportedFilterVisibility(model);
  if (edits.length) model = patchCells(model, edits).model;
  const currentSheet=model.sheets[sheetIndex], files={...model.files}, encode=value=>new TextEncoder().encode(value), decode=value=>new TextDecoder().decode(value);
  const append = (xml, tag, fragment, count) => {
    const element=findElementRanges(xml,tag).get('');
    if (!element) fail('새 표의 XML 구성요소를 추가할 위치가 없습니다.');
    let opening=xml.slice(element.start,element.startTagEnd).replace(/\/\s*>$/,'>');
    if (count !== undefined) opening=setTagAttribute(opening,'count',count);
    return xml.slice(0,element.start)+opening+(element.selfClosing?'':xml.slice(element.startTagEnd,element.closingStart))+fragment+'</'+tag+'>'+xml.slice(element.end);
  };
  const partNames=new Set(children(parseXml(files['[Content_Types].xml'],'[Content_Types].xml').root,'Override').map(item=>item.attributes.PartName));
  let fileIndex=1; while (files['xl/tables/table'+fileIndex+'.xml'] || partNames.has('/xl/tables/table'+fileIndex+'.xml')) fileIndex++;
  const tablePath='xl/tables/table'+fileIndex+'.xml', usedIds=new Set(tables.map(table=>Number(parseXml(files[table.path],table.path).root.attributes.id)));
  let id=1; while (usedIds.has(id)) id++;
  const reference=columnLetters(range.left+1)+(range.top+1)+':'+columnLetters(range.right+1)+(range.bottom+1);
  let worksheet=currentSheet.xml, filterXml='<autoFilter ref="'+reference+'"/>';
  if (migrateFilter) {
    const filter=findElementRanges(worksheet,'autoFilter').get('');
    if (!filter) fail('이동할 자동 필터를 찾을 수 없습니다.');
    filterXml=worksheet.slice(filter.start,filter.end); worksheet=worksheet.slice(0,filter.start)+worksheet.slice(filter.end);
  }
  const root=parseXml(encode(worksheet),currentSheet.path).root, relationNamespace='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  if (root.attributes['xmlns:r'] && root.attributes['xmlns:r'] !== relationNamespace) fail('워크시트 관계 네임스페이스를 안전하게 추가할 수 없습니다.');
  if (!root.attributes['xmlns:r']) { const element=findElementRanges(worksheet,'worksheet').get(''); worksheet=worksheet.slice(0,element.start)+setTagAttribute(worksheet.slice(element.start,element.startTagEnd),'xmlns:r',relationNamespace)+worksheet.slice(element.startTagEnd); }
  const relPath=path.join(path.dirname(currentSheet.path),'_rels',path.basename(currentSheet.path)+'.rels');
  let relXml=files[relPath]?decode(files[relPath]):'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>';
  const relIds=new Set(children(parseXml(encode(relXml),relPath).root,'Relationship').map(item=>item.attributes.Id));
  let relIndex=1; while (relIds.has('rId'+relIndex)) relIndex++;
  const relId='rId'+relIndex, tablePart='<tablePart r:id="'+relId+'"/>';
  relXml=append(relXml,'Relationships','<Relationship Id="'+relId+'" Type="'+relationNamespace+'/table" Target="'+xmlEscape(path.relative(path.dirname(currentSheet.path),tablePath))+'"/>');
  if (child(root,'tableParts')) worksheet=append(worksheet,'tableParts',tablePart,currentSheet.tables.length+1);
  else worksheet=append(worksheet,'worksheet','<tableParts count="1">'+tablePart+'</tableParts>');
  files[currentSheet.path]=encode(worksheet); files[relPath]=encode(relXml);
  files[tablePath]=encode('<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="'+id+'" name="'+xmlEscape(name)+'" displayName="'+xmlEscape(name)+'" ref="'+reference+'" headerRowCount="1" totalsRowCount="0">'+filterXml+'<tableColumns count="'+headers.length+'">'+headers.map((header,index)=>'<tableColumn id="'+(index+1)+'" name="'+xmlEscape(header).replaceAll('\r','&#13;').replaceAll('\n','&#10;').replaceAll('\t','&#9;')+'"/>').join('')+'</tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>');
  files['[Content_Types].xml']=encode(append(decode(files['[Content_Types].xml']),'Types','<Override PartName="/'+tablePath+'" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>'));
  files['xl/workbook.xml']=encode(enableFullCalculationOnLoad(decode(files['xl/workbook.xml'])));
  const checked=analyze(zipSync(files,{level:0}),model.extension);
  if (!checked.writable) fail('표 생성 후 패키지 검사에서 거절되었습니다: '+checked.readOnlyReason);
  const allowed=new Set([currentSheet.path,relPath,tablePath,'[Content_Types].xml','xl/workbook.xml']);
  for (const [part,before] of Object.entries(model.files)) if (!allowed.has(part) && (!checked.files[part] || Buffer.compare(Buffer.from(before),Buffer.from(checked.files[part])) !== 0)) fail('표 생성 중 다른 패키지 항목이 달라졌습니다: '+part);
  return patchFilterVisibility(checked,options);
}

function patchTableStyles(model, edits) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  if (!Array.isArray(edits) || edits.length > LIMITS.sheets) fail('Excel 표 스타일 수정 요청이 유효하지 않습니다.');
  if (!edits.length) return { bytes: model.bytes, model };
  const files = { ...model.files }, changedPaths = new Set(), seen = new Set();
  for (const edit of edits) {
    const { sheetIndex, tableName, styleInfo } = edit || {}, sheet = model.sheets[sheetIndex], table = sheet?.tables?.find(item => item.name === tableName);
    if (!Number.isInteger(sheetIndex) || !sheet || !table || seen.has(sheetIndex + ':' + tableName) || !styleInfo || typeof styleInfo !== 'object' || Array.isArray(styleInfo)) fail('Excel 표 스타일 수정 요청이 유효하지 않습니다.');
    if (!Object.hasOwn(styleInfo, 'name') || styleInfo.name !== null && (typeof styleInfo.name !== 'string' || !styleInfo.name.length || styleInfo.name.length > 255 || /[\u0000-\u001f]/.test(styleInfo.name)) || ['showFirstColumn', 'showLastColumn', 'showRowStripes', 'showColumnStripes'].some(key => typeof styleInfo[key] !== 'boolean')) fail('Excel 표 스타일 이름 또는 옵션이 유효하지 않습니다.');
    if (styleInfo.name === null && ['showFirstColumn', 'showLastColumn', 'showRowStripes', 'showColumnStripes'].some(key => styleInfo[key])) fail('표 스타일을 지우려면 스타일 옵션도 꺼야 합니다.');
    seen.add(sheetIndex + ':' + tableName);
    const xml = new TextDecoder().decode(files[table.path] || model.files[table.path]), styleRanges = findElementRanges(xml, 'tableStyleInfo', null, true);
    if (styleRanges.size > 1) fail('Excel 표 스타일 정보가 중복되었습니다.');
    const currentStyle = table.styleInfo;
    if (currentStyle === null && styleInfo.name === null) continue;
    if (JSON.stringify(currentStyle) === JSON.stringify(styleInfo)) continue;
    let updated = xml;
    if (styleInfo.name === null) {
      if (styleRanges.size) { const range = styleRanges.values().next().value; updated = xml.slice(0, range.start) + xml.slice(range.end); }
    } else if (styleRanges.size) {
      const range = styleRanges.values().next().value, original = xml.slice(range.start, range.startTagEnd);
      let opening = original;
      for (const [key, value] of [['name', styleInfo.name], ['showFirstColumn', styleInfo.showFirstColumn ? '1' : '0'], ['showLastColumn', styleInfo.showLastColumn ? '1' : '0'], ['showRowStripes', styleInfo.showRowStripes ? '1' : '0'], ['showColumnStripes', styleInfo.showColumnStripes ? '1' : '0']]) opening = setTagAttribute(opening, key, value);
      updated = xml.slice(0, range.start) + opening + xml.slice(range.startTagEnd);
    } else {
      const tableRange = findElementRanges(xml, 'table').get('');
      if (!tableRange || tableRange.selfClosing) fail('Excel 표 스타일을 추가할 위치가 없습니다.');
      const attributes = [['name', styleInfo.name], ['showFirstColumn', styleInfo.showFirstColumn ? '1' : '0'], ['showLastColumn', styleInfo.showLastColumn ? '1' : '0'], ['showRowStripes', styleInfo.showRowStripes ? '1' : '0'], ['showColumnStripes', styleInfo.showColumnStripes ? '1' : '0']];
      const serialized = attributes.map(([key, value]) => ' ' + key + '="' + xmlEscape(value) + '"').join('');
      updated = xml.slice(0, tableRange.closingStart) + '<tableStyleInfo' + serialized + '/>' + xml.slice(tableRange.closingStart);
    }
    files[table.path] = new TextEncoder().encode(updated);
    changedPaths.add(table.path);
  }
  if (!changedPaths.size) return { bytes: model.bytes, model };
  const bytes = zipSync(files, { level: 0 }), checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('Excel 표 스타일 저장 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  for (const [name, before] of Object.entries(model.files)) if (!changedPaths.has(name) && (!checked.files[name] || Buffer.compare(Buffer.from(before), Buffer.from(checked.files[name])) !== 0)) fail('Excel 표 스타일 저장 중 비수정 패키지 항목이 달라졌습니다: ' + name);
  return { bytes, model: checked };
}

function patchTableRanges(model, edits, options = {}) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  model = preserveImportedFilterVisibility(model);
  if (model.hasCharts) fail('차트가 있는 통합문서에서는 표 범위를 변경할 수 없습니다.');
  if (!Array.isArray(edits) || edits.length > LIMITS.sheets) fail('Excel 표 범위 수정 요청이 유효하지 않습니다.');
  if (!edits.length) return { bytes: model.bytes, model };
  const outputFiles = Object.assign({}, model.files), changedPaths = new Set(), seen = new Set(), moves = [];
  for (const edit of edits) {
    const { sheetIndex, tableName, range } = edit || {}, sheet = model.sheets[sheetIndex];
    const table = sheet?.tables?.find(item => item.name === tableName);
    if (!Number.isInteger(sheetIndex) || !sheet || !table || seen.has(sheetIndex + ':' + tableName) || !range || !['top', 'bottom', 'left', 'right'].every(key => Number.isInteger(range[key])) || range.top !== table.range.top || range.left < 0 || range.left > table.range.right || range.right < table.range.left || range.right >= LIMITS.columns || range.bottom < range.top + Math.max(1, (table.headerRowCount || 0) + (table.totalsRowCount || 0)) - 1 || range.bottom >= LIMITS.rows) fail('표 범위는 기존 머리글을 유지하고 유효한 데이터 행과 열을 포함해야 합니다.');
    if (table.totalsRowCount && range.bottom !== table.range.bottom) {
      const source = { top: table.range.bottom, bottom: table.range.bottom, left: Math.max(table.range.left, range.left), right: Math.min(table.range.right, range.right) };
      const target = { ...source, top: range.bottom, bottom: range.bottom };
      validateTotalsMove(model, sheetIndex, source, target);
      moves.push({ sheetIndex, source, action: 'move', rowDelta: range.bottom - table.range.bottom, columnDelta: 0 });
    }
    if (sheet.tables.some(other => other !== table && range.top <= other.range.bottom && range.bottom >= other.range.top && range.left <= other.range.right && range.right >= other.range.left)) fail('표 범위가 다른 Excel 표와 겹칩니다.');
    seen.add(sheetIndex + ':' + tableName);
    const width = range.right - range.left + 1;
    const originalXml = new TextDecoder().decode(outputFiles[table.path] || model.files[table.path]);
    const xmlOriginal = resizeTableSortState(originalXml, sheet.name, table, range);
    const columns = [];
    for (let column = range.left; column <= range.right; column++) {
      const oldIndex = column - table.range.left;
      if (oldIndex >= 0 && oldIndex < table.columns.length) columns.push(table.columns[oldIndex]);
      else {
        const header = sheet.cells.get(columnLetters(column + 1) + (table.range.top + 1));
        const name = String(header?.value ?? '').trim();
        if (!name || header?.formula) fail('추가할 표 열에는 수식이 아닌 고유한 머리글 값이 있어야 합니다.');
        if (columns.some(value => value.toLocaleLowerCase() === name.toLocaleLowerCase())) fail('표 열 머리글은 서로 다른 이름이어야 합니다.');
        columns.push(name);
      }
    }
    if (new Set(columns.map(name => name.toLocaleLowerCase())).size !== columns.length) fail('표 열 머리글은 서로 다른 이름이어야 합니다.');
    const nextColumnSet = new Set(columns.map(name => name.toLocaleLowerCase())), removed = table.columns.filter(name => !nextColumnSet.has(name.toLocaleLowerCase())).map(value => value.toLocaleLowerCase());
    if (removed.length) {
      const refersToRemovedColumn = formula => {
        const normalized = String(formula || '').toLocaleLowerCase();
        return removed.some(name => normalized.includes('[' + name + ']') || normalized.includes('[#headers],[' + name + ']') || normalized.includes('[#data],[' + name + ']') || normalized.includes('[#totals],[' + name + ']'));
      };
      if (model.sheets.some((item, ownerSheetIndex) => [...item.cells.values()].some(cell => {
        if (!cell.formula) return false;
        const formula = String(cell.formula).toLocaleLowerCase(), inThisTable = ownerSheetIndex === sheetIndex && cell.row >= table.range.top && cell.row <= table.range.bottom && cell.column >= table.range.left && cell.column <= table.range.right;
        return formula.includes(tableName.toLocaleLowerCase() + '[') && refersToRemovedColumn(formula) || inThisTable && removed.some(name => formula.includes('[@' + name + ']'));
      }))) fail('삭제할 표 열을 참조하는 수식이 있어 표 범위를 줄일 수 없습니다.');
      const tableFormulaTags = [...xmlOriginal.matchAll(/<(?:calculatedColumnFormula|totalsRowFormula)\b[^>]*>([\s\S]*?)<\/(?:calculatedColumnFormula|totalsRowFormula)>/g)];
      if (tableFormulaTags.some(match => refersToRemovedColumn(match[1].replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')))) fail('삭제할 표 열을 참조하는 표 수식이 있어 표 범위를 줄일 수 없습니다.');
    }
    const columnsChanged = range.left !== table.range.left || range.right !== table.range.right;
    let xml = columnsChanged ? resizeTableColumnDefinitions(xmlOriginal, columns, { reconcile: true }) : xmlOriginal;
    if (columnsChanged) xml = remapTableFilterColumns(xml, table.columns, columns);
    const reference = columnLetters(range.left + 1) + (range.top + 1) + ':' + columnLetters(range.right + 1) + (range.bottom + 1);
    const patches = [];
    for (const tag of ['table', 'autoFilter']) {
      const element = findElementRanges(xml, tag).get('');
      if (!element) { if (tag === 'autoFilter' && !table.autoFilter) continue; fail(`Excel 표 ${tag} 범위를 찾을 수 없습니다.`); }
      const opening = xml.slice(element.start, element.startTagEnd);
      if (!/\bref="[^"]+"/.test(opening)) fail(`Excel 표 ${tag} 범위가 유효하지 않습니다.`);
      patches.push({ start: element.start, end: element.startTagEnd, replacement: replaceTagAttribute(opening, 'ref', tag === 'autoFilter' ? columnLetters(range.left + 1) + (range.top + 1) + ':' + columnLetters(range.right + 1) + (range.bottom + 1 - (table.totalsRowCount || 0)) : reference) });
    }
    patches.sort((left, right) => right.start - left.start);
    let updated = xml;
    for (const patch of patches) updated = updated.slice(0, patch.start) + patch.replacement + updated.slice(patch.end);
    outputFiles[table.path] = new TextEncoder().encode(updated);
    changedPaths.add(table.path);
  }
  outputFiles['xl/workbook.xml'] = new TextEncoder().encode(enableFullCalculationOnLoad(new TextDecoder().decode(outputFiles['xl/workbook.xml'])));
  changedPaths.add('xl/workbook.xml');
  const bytes = zipSync(outputFiles, { level: 0 }), checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('Excel 표 범위 저장 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  for (const [name, before] of Object.entries(model.files)) if (!changedPaths.has(name) && (!checked.files[name] || Buffer.compare(Buffer.from(before), Buffer.from(checked.files[name])) !== 0)) fail('Excel 표 범위 저장 중 비수정 패키지 항목이 달라졌습니다: ' + name);
  let result = moves.length ? moveTableTotals(checked, model, moves) : { bytes: checked.bytes, model: checked };
  const growth = edits.flatMap(edit => {
    const oldTable = model.sheets[edit.sheetIndex].tables.find(table => table.name === edit.tableName), nextTable = result.model.sheets[edit.sheetIndex].tables.find(table => table.name === edit.tableName);
    const top = oldTable.range.bottom - oldTable.totalsRowCount + 1, bottom = nextTable.range.bottom - nextTable.totalsRowCount;
    return bottom >= top ? [{ sheetIndex: edit.sheetIndex, tableName: edit.tableName, top, bottom }] : [];
  });
  if (growth.length) result = fillCalculatedRows(result.model, growth, options);
  return patchFilterVisibility(result.model);
}

function transformAutoFilter(filter, operation) {
  if (!filter) return null;
  const next = { ...filter, columns: filter.columns.map(item => item.iconFilter ? { column: item.column, iconFilter: { ...item.iconFilter } } : item.colorFilter ? { column: item.column, colorFilter: { ...item.colorFilter } } : item.dynamic ? { column: item.column, dynamic: { ...item.dynamic } } : item.top10 ? { column: item.column, top10: { ...item.top10 } } : item.custom ? { column: item.column, custom: { and: item.custom.and, conditions: item.custom.conditions.map(condition => ({ ...condition })) } } : { ...item, values: [...item.values], ...(item.dateGroups ? { dateGroups: item.dateGroups.map(group => ({ ...group })) } : {}) }) };
  const coordinate = operation.index + 1;
  const first = operation.axis === 'row' ? next.top + 1 : next.left + 1;
  const last = operation.axis === 'row' ? next.bottom + 1 : next.right + 1;
  if (operation.action === 'insert') {
    if (coordinate <= first) {
      if (operation.axis === 'row') { next.top++; next.bottom++; }
      else {
        next.left++; next.right++;
        for (const item of next.columns) item.column++;
      }
    } else if (coordinate <= last) {
      if (operation.axis === 'row') next.bottom++;
      else {
        next.right++;
        for (const item of next.columns) if (item.column >= operation.index) item.column++;
      }
    }
  } else if (coordinate < first) {
    if (operation.axis === 'row') { next.top--; next.bottom--; }
    else {
      next.left--; next.right--;
      for (const item of next.columns) item.column--;
    }
  } else if (coordinate <= last) {
    if (operation.axis === 'row') {
      next.bottom--;
      if (next.bottom <= next.top) return null;
    } else {
      next.right--;
      next.columns = next.columns.filter(item => item.column !== operation.index).map(item => ({
        ...item,
        column: item.column > operation.index ? item.column - 1 : item.column
      }));
      if (next.right < next.left) return null;
    }
  }
  return next;
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

function appendRelationship(xml, relationship, partName) {
  const { root } = parseXml(new TextEncoder().encode(xml), partName);
  if (localName(root) !== 'Relationships') fail('OOXML 관계 파일을 수정할 수 없습니다.');
  const rootRange = findElementRanges(xml, root.name).values().next().value;
  if (!rootRange) fail('OOXML 관계 파일 끝 태그를 찾을 수 없습니다.');
  if (rootRange.selfClosing) {
    const opening = xml.slice(rootRange.start, rootRange.startTagEnd).replace(/\/\s*>$/, '>');
    return xml.slice(0, rootRange.start) + opening + relationship + '</' + root.name + '>' + xml.slice(rootRange.end);
  }
  return xml.slice(0, rootRange.closingStart) + relationship + xml.slice(rootRange.closingStart);
}

function insertWorksheetDrawing(xml, relationshipId) {
  const rootRange = findElementRanges(xml, 'worksheet').values().next().value;
  const opening = rootRange && xml.slice(rootRange.start, rootRange.startTagEnd);
  if (!opening) fail('워크시트 루트 태그를 수정할 수 없습니다.');
  if (!/\bxmlns:r\s*=/.test(opening)) xml = xml.replace(opening, opening.replace(/>$/, ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'));
  const after = ['legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];
  const points = after.flatMap(name => [...findElementRanges(xml, name).values()].map(range => range.start));
  const end = xml.lastIndexOf('</worksheet>'), position = points.length ? Math.min(...points) : end;
  if (position < 0) fail('워크시트에 차트 관계를 추가할 위치가 없습니다.');
  return xml.slice(0, position) + '<drawing r:id="' + xmlEscape(relationshipId) + '"/>' + xml.slice(position);
}

function chartCacheElement(type, values) {
  const cacheName = type === 'number' ? 'numCache' : 'strCache';
  const points = values.map((value, index) => value === null ? '' : '<c:pt idx="' + index + '"><c:v>' + xmlEscape(value) + '</c:v></c:pt>').join('');
  return '<c:' + cacheName + '>' + (type === 'number' ? '<c:formatCode>General</c:formatCode>' : '') + '<c:ptCount val="' + values.length + '"/>' + points + '</c:' + cacheName + '>';
}

function buildChartSeries(model, sheetIndex, range, chartType, prefix = 'c:') {
  const sheet = model.sheets[sheetIndex];
  if (!sheet || !range || !['top', 'bottom', 'left', 'right'].every(key => Number.isInteger(range[key])) || range.top < 0 || range.left < 0 || range.bottom <= range.top || range.right <= range.left || range.bottom >= LIMITS.rows || range.right >= LIMITS.columns || (range.bottom - range.top + 1) * (range.right - range.left + 1) > 1000) fail('차트를 만들려면 머리글 행과 범주 열을 포함해 1,000셀 이하의 사각형 범위를 선택하세요.');
  if (sheet.mergedCells?.some(merged => range.top <= merged.bottom && range.bottom >= merged.top && range.left <= merged.right && range.right >= merged.left)) fail('병합 셀이 포함된 범위에서는 차트를 만들 수 없습니다.');
  const calculator = createFormulaCalculator({ state: { formulas: true, sheets: model.sheets, definedNames: model.definedNames, date1904: model.date1904 }, getCell: (row, column, targetSheet = 0) => model.sheets[targetSheet]?.cells.get(columnLetters(column + 1) + (row + 1)) || null });
  const get = (row, column) => {
    const cell = sheet.cells.get(columnLetters(column + 1) + (row + 1));
    if (!cell) return null;
    const value = cell.formula ? calculator.value(sheetIndex, row, column) : cell.value;
    return value === null || value === undefined || typeof value === 'string' && value.startsWith('#') ? null : value;
  };
  const categories = [];
  for (let row = range.top + 1; row <= range.bottom; row++) {
    const cell = sheet.cells.get(columnLetters(range.left + 1) + (row + 1)), value = get(row, range.left);
    if (chartType === 'scatter') {
      const numeric = typeof value === 'number' ? value : value !== null && (cell?.kind === 'number' || cell?.formula) ? Number(value) : null;
      if (value !== null && !Number.isFinite(numeric)) fail('분산형 차트의 첫 번째 열에는 숫자 X 값을 선택하세요.');
      categories.push(numeric);
    } else categories.push(value === null ? null : cell?.kind === 'number' && !cell.formula ? formatCellValue(cell) : typeof value === 'boolean' ? value ? 'TRUE' : 'FALSE' : String(value));
  }
  if (chartType === 'scatter' && !categories.some(value => Number.isFinite(value))) fail('분산형 차트에는 숫자 X 값이 필요합니다.');
  const quotedSheet = "'" + sheet.name.replaceAll("'", "''") + "'!";
  const refCell = (row, column) => '$' + columnLetters(column + 1) + '$' + (row + 1);
  const refRange = (top, left, bottom = top, right = left) => quotedSheet + refCell(top, left) + (top === bottom && left === right ? '' : ':' + refCell(bottom, right));
  const series = [];
  for (let column = range.left + 1; column <= range.right; column++) {
    const rawName = get(range.top, column), name = rawName === null || String(rawName) === '' ? 'Series ' + (column - range.left) : String(rawName);
    const values = [];
    for (let row = range.top + 1; row <= range.bottom; row++) {
      const raw = get(row, column), cell = sheet.cells.get(columnLetters(column + 1) + (row + 1));
      const numeric = typeof raw === 'number' ? raw : (cell?.kind === 'number' || cell?.formula) && raw !== null ? Number(raw) : null;
      values.push(Number.isFinite(numeric) ? numeric : null);
    }
    if (!values.some(value => value !== null)) fail('선택 범위에 숫자 데이터가 있는 계열이 없습니다.');
    const index = series.length;
    const categoryXml = chartType === 'scatter'
      ? '<c:xVal><c:numRef><c:f>' + xmlEscape(refRange(range.top + 1, range.left, range.bottom)) + '</c:f>' + chartCacheElement('number', categories) + '</c:numRef></c:xVal>'
      : '<c:cat><c:strRef><c:f>' + xmlEscape(refRange(range.top + 1, range.left, range.bottom)) + '</c:f>' + chartCacheElement('string', categories) + '</c:strRef></c:cat>';
    const valueTag = chartType === 'scatter' ? 'yVal' : 'val';
    const marker = chartType === 'scatter' ? '<c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>' : '';
    series.push('<c:ser><c:idx val="' + index + '"/><c:order val="' + index + '"/><c:tx><c:strRef><c:f>' + xmlEscape(refRange(range.top, column)) + '</c:f>' + chartCacheElement('string', [name]) + '</c:strRef></c:tx>' + marker + categoryXml + '<c:' + valueTag + '><c:numRef><c:f>' + xmlEscape(refRange(range.top + 1, column, range.bottom)) + '</c:f>' + chartCacheElement('number', values) + '</c:numRef></c:' + valueTag + '></c:ser>');
  }
  if (!series.length || series.length > 32) fail('차트는 1~32개의 숫자 계열을 지원합니다.');
  if (['pie', 'doughnut'].includes(chartType) && series.length !== 1) fail('원형·도넛 차트에는 숫자 계열 하나만 선택하세요.');
  return series.map(xml => xml.replaceAll('<c:', '<' + prefix).replaceAll('</c:', '</' + prefix));
}

function patchAddChart(model, sheetIndex, range, chartType = 'bar') {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  if (!['bar', 'line', 'area', 'scatter', 'pie', 'doughnut'].includes(chartType)) fail('지원하지 않는 차트 유형입니다.');
  const sheet = model.sheets[sheetIndex];
  if (!sheet || !range || !['top', 'bottom', 'left', 'right'].every(key => Number.isInteger(range[key])) || range.top < 0 || range.left < 0 || range.bottom <= range.top || range.right <= range.left || range.bottom >= LIMITS.rows || range.right >= LIMITS.columns || (range.bottom - range.top + 1) * (range.right - range.left + 1) > 1000) fail('차트를 만들려면 머리글 행과 범주 열을 포함해 1,000셀 이하의 사각형 범위를 선택하세요.');
  if (sheet.mergedCells?.some(merged => range.top <= merged.bottom && range.bottom >= merged.top && range.left <= merged.right && range.right >= merged.left)) fail('병합 셀이 포함된 범위에서는 차트를 만들 수 없습니다.');
  if ((sheet.charts || []).length >= 32) fail('워크시트 하나에는 최대 32개의 차트만 추가할 수 있습니다.');
  const series = buildChartSeries(model, sheetIndex, range, chartType);
  let chartNumber = 1; while (model.files['xl/charts/chart' + chartNumber + '.xml']) chartNumber++;
  const chartPath = 'xl/charts/chart' + chartNumber + '.xml';
  const axes = '<c:catAx><c:axId val="48650112"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="48672768"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx><c:valAx><c:axId val="48672768"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="48650112"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>';
  const scatterAxes = [[48650112, 48672768, 'b'], [48672768, 48650112, 'l']].map(([id, crossId, position]) => '<c:valAx><c:axId val="' + id + '"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="' + position + '"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="' + crossId + '"/><c:crosses val="autoZero"/></c:valAx>').join('');
  const chartKind = chartType === 'bar' ? '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>' + series.join('') + '<c:gapWidth val="150"/><c:overlap val="0"/><c:axId val="48650112"/><c:axId val="48672768"/></c:barChart>' + axes
    : chartType === 'line' ? '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>' + series.join('') + '<c:marker val="1"/><c:smooth val="0"/><c:axId val="48650112"/><c:axId val="48672768"/></c:lineChart>' + axes
    : chartType === 'area' ? '<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>' + series.join('') + '<c:axId val="48650112"/><c:axId val="48672768"/></c:areaChart>' + axes
    : chartType === 'scatter' ? '<c:scatterChart><c:scatterStyle val="marker"/><c:varyColors val="0"/>' + series.join('') + '<c:axId val="48650112"/><c:axId val="48672768"/></c:scatterChart>' + scatterAxes
    : chartType === 'pie' ? '<c:pieChart><c:varyColors val="1"/>' + series.join('') + '<c:firstSliceAng val="0"/></c:pieChart>'
    : '<c:doughnutChart><c:varyColors val="1"/>' + series.join('') + '<c:firstSliceAng val="0"/><c:holeSize val="50"/></c:doughnutChart>';
  const chartXml = '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart>' + chartTitleXml('c:', 'Chart') + '<c:plotArea><c:layout/>' + chartKind + '</c:plotArea><c:legend><c:legendPos val="b"/><c:layout/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>';
  const fromColumn = Math.min(range.right + 2, LIMITS.columns - 8), fromRow = range.top, toColumn = Math.min(LIMITS.columns - 1, fromColumn + 7), toRow = Math.min(LIMITS.rows - 1, Math.max(range.bottom + 2, range.top + 10));
  const relPath = path.join(path.dirname(sheet.path), '_rels', path.basename(sheet.path) + '.rels');
  let relXml = model.files[relPath] ? new TextDecoder().decode(model.files[relPath]) : '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  const relations = parseRelationships(parseXml(new TextEncoder().encode(relXml), relPath).root, path.dirname(sheet.path), ['table', 'hyperlink', 'drawing']);
  const worksheetRoot = parseXml(model.files[sheet.path], sheet.path).root, drawingElement = child(worksheetRoot, 'drawing');
  let drawingPath, worksheetXml = sheet.xml, drawingXml, drawingRelXml, addDrawingOverride = false;
  if (drawingElement) {
    const drawingRelation = relations.get(drawingElement.attributes['r:id']);
    if (!drawingRelation?.path || !drawingRelation.Type.endsWith('/drawing')) fail('기존 차트 그리기 관계를 해석할 수 없습니다.');
    drawingPath = drawingRelation.path;
    drawingXml = new TextDecoder().decode(model.files[drawingPath]);
    const drawingRelPath = path.join(path.dirname(drawingPath), '_rels', path.basename(drawingPath) + '.rels');
    if (!model.files[drawingRelPath]) fail('기존 차트 관계 파일을 찾을 수 없습니다.');
    drawingRelXml = new TextDecoder().decode(model.files[drawingRelPath]);
    const chartRelations = parseRelationships(parseXml(model.files[drawingRelPath], drawingRelPath).root, path.dirname(drawingPath), ['chart', 'image']);
    let chartRelationNumber = 1; while (chartRelations.has('rId' + chartRelationNumber)) chartRelationNumber++;
    const chartRelationId = 'rId' + chartRelationNumber;
    drawingRelXml = appendRelationship(drawingRelXml, '<Relationship Id="' + chartRelationId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart' + chartNumber + '.xml"/>', drawingRelPath);
    const drawingRoot = parseXml(model.files[drawingPath], drawingPath).root, rootRange = findElementRanges(drawingXml, drawingRoot.name).values().next().value;
    if (!rootRange || localName(drawingRoot) !== 'wsDr') fail('기존 차트 그리기 XML 구조가 유효하지 않습니다.');
    let shapeId = 1;
    for (const node of (function* walk(item) { if (localName(item) === 'cNvPr') yield item; for (const nested of item.children) yield* walk(nested); })(drawingRoot)) shapeId = Math.max(shapeId, (Number(node.attributes.id) || 0) + 1);
    const anchor = '<xdr:twoCellAnchor><xdr:from><xdr:col>' + fromColumn + '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + fromRow + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>' + toColumn + '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + toRow + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="' + shapeId + '" name="Chart ' + chartNumber + '"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="' + chartRelationId + '"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>';
    drawingXml = drawingXml.slice(0, rootRange.closingStart) + anchor + drawingXml.slice(rootRange.closingStart);
  } else {
    let drawingNumber = 1; while (model.files['xl/drawings/drawing' + drawingNumber + '.xml']) drawingNumber++;
    drawingPath = 'xl/drawings/drawing' + drawingNumber + '.xml';
    const ids = new Set(relations.keys()); let relationNumber = 1; while (ids.has('rId' + relationNumber)) relationNumber++;
    const drawingRelId = 'rId' + relationNumber;
    relXml = appendRelationship(relXml, '<Relationship Id="' + drawingRelId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing' + drawingNumber + '.xml"/>', relPath);
    worksheetXml = insertWorksheetDrawing(sheet.xml, drawingRelId);
    const anchor = '<xdr:twoCellAnchor><xdr:from><xdr:col>' + fromColumn + '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + fromRow + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>' + toColumn + '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + toRow + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart ' + chartNumber + '"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>';
    drawingXml = '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' + anchor + '</xdr:wsDr>';
    const drawingRelPath = path.join(path.dirname(drawingPath), '_rels', path.basename(drawingPath) + '.rels');
    drawingRelXml = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart' + chartNumber + '.xml"/></Relationships>';
    addDrawingOverride = true;
  }
  const drawingRelPath = path.join(path.dirname(drawingPath), '_rels', path.basename(drawingPath) + '.rels');
  const typesPath = '[Content_Types].xml', typesXml = new TextDecoder().decode(model.files[typesPath]);
  if (!typesXml.includes('</Types>')) fail('콘텐츠 형식 목록에 차트 형식을 추가할 수 없습니다.');
  const overrides = (addDrawingOverride ? '<Override PartName="/' + drawingPath + '" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' : '') + '<Override PartName="/' + chartPath + '" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>';
  const outputFiles = Object.assign({}, model.files);
  outputFiles[sheet.path] = new TextEncoder().encode(worksheetXml);
  if (!model.files[relPath] || !drawingElement) outputFiles[relPath] = new TextEncoder().encode(relXml);
  outputFiles[drawingPath] = new TextEncoder().encode(drawingXml);
  outputFiles[drawingRelPath] = new TextEncoder().encode(drawingRelXml);
  outputFiles[chartPath] = new TextEncoder().encode(chartXml);
  outputFiles[typesPath] = new TextEncoder().encode(typesXml.replace('</Types>', overrides + '</Types>'));
  const bytes = zipSync(outputFiles, { level: 0 }), checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('차트 삽입 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  for (const [name, before] of Object.entries(model.files)) {
    if ([sheet.path, relPath, drawingPath, drawingRelPath, typesPath].includes(name)) continue;
    if (!checked.files[name] || Buffer.compare(Buffer.from(before), Buffer.from(checked.files[name])) !== 0) fail('차트 삽입 중 기존 패키지 항목이 달라졌습니다: ' + name);
  }
  return { bytes, model: checked };
}

function packageReferences(files, targetPath) {
  const references = [];
  for (const [name, bytes] of Object.entries(files)) {
    if (!name.endsWith('.rels')) continue;
    const relations = parseRelationships(parseXml(bytes, name).root, path.dirname(path.dirname(name)));
    for (const [id, relation] of relations) if (relation.path === targetPath) references.push({ name, id });
  }
  return references;
}

function patchDeleteChart(model, sheetIndex, chartIndex) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  const sheet = model.sheets[sheetIndex], chartPath = sheet?.chartPaths?.[chartIndex];
  if (!Number.isInteger(chartIndex) || chartIndex < 0 || !chartPath || !sheet.charts[chartIndex]) fail('삭제할 차트를 찾을 수 없습니다.');
  const outputFiles = { ...model.files }, modified = new Set(), removedParts = [];
  const write = (name, xml) => { outputFiles[name] = new TextEncoder().encode(xml); modified.add(name); };
  const remove = name => { delete outputFiles[name]; modified.add(name); removedParts.push(name); };
  const removeElement = (xml, name, attribute, key) => {
    const range = findElementRanges(xml, name, attribute).get(key);
    if (!range) fail('삭제할 OOXML 요소를 찾을 수 없습니다: ' + name);
    return xml.slice(0, range.start) + xml.slice(range.end);
  };
  const sheetRelPath = path.join(path.dirname(sheet.path), '_rels', path.basename(sheet.path) + '.rels');
  const sheetRelXml = new TextDecoder().decode(model.files[sheetRelPath]);
  const sheetRoot = parseXml(model.files[sheet.path], sheet.path).root, drawing = childrenByLocalName(sheetRoot, 'drawing')[0];
  const drawingId = drawing?.attributes['r:id'];
  const drawingPath = parseRelationships(parseXml(model.files[sheetRelPath], sheetRelPath).root, path.dirname(sheet.path)).get(drawingId)?.path;
  if (!drawingPath || !model.files[drawingPath]) fail('삭제할 차트의 그리기를 찾을 수 없습니다.');
  const drawingRelPath = path.join(path.dirname(drawingPath), '_rels', path.basename(drawingPath) + '.rels');
  const { xml: drawingXml, root: drawingRoot } = parseXml(model.files[drawingPath], drawingPath);
  const entries = [];
  const refs = node => [...(function* walk(item) { if (localName(item) === 'chart' && item.attributes['r:id']) yield item; for (const child of item.children) yield* walk(child); })(node)];
  for (const anchor of drawingRoot.children) for (const reference of refs(anchor)) entries.push({ anchor, reference });
  const selected = entries[chartIndex];
  if (!selected || !['oneCellAnchor', 'twoCellAnchor', 'absoluteAnchor'].includes(localName(selected.anchor)) || refs(selected.anchor).length !== 1) fail('차트의 그리기 위치를 안전하게 삭제할 수 없습니다.');
  const relationId = selected.reference.attributes['r:id'];
  const drawingRelations = parseRelationships(parseXml(model.files[drawingRelPath], drawingRelPath).root, path.dirname(drawingPath));
  if (drawingRelations.get(relationId)?.path !== chartPath) fail('차트와 그리기 관계가 일치하지 않습니다.');
  const remaining = entries.filter(entry => entry !== selected);
  const sharedDrawing = packageReferences(model.files, drawingPath).some(item => item.name !== sheetRelPath || item.id !== drawingId);
  let addedDrawingPath = null;
  if (!remaining.length) {
    write(sheet.path, removeElement(sheet.xml, drawing.name, 'r:id', drawingId));
    const nextRelations = removeElement(sheetRelXml, 'Relationship', 'Id', drawingId);
    if (parseXml(new TextEncoder().encode(nextRelations), sheetRelPath).root.children.length) write(sheetRelPath, nextRelations);
    else remove(sheetRelPath);
    if (!sharedDrawing) { remove(drawingPath); remove(drawingRelPath); }
  } else {
    const siblings = drawingRoot.children.filter(node => node.name === selected.anchor.name);
    const anchorRange = [...findElementRanges(drawingXml, selected.anchor.name, undefined, true).values()][siblings.indexOf(selected.anchor)];
    if (!anchorRange) fail('삭제할 차트의 XML 위치를 찾을 수 없습니다.');
    const nextDrawing = drawingXml.slice(0, anchorRange.start) + drawingXml.slice(anchorRange.end);
    let nextRelations = new TextDecoder().decode(model.files[drawingRelPath]);
    if (!remaining.some(entry => entry.reference.attributes['r:id'] === relationId)) nextRelations = removeElement(nextRelations, 'Relationship', 'Id', relationId);
    if (sharedDrawing) {
      let number = 1, copyPath;
      do { copyPath = path.join(path.dirname(drawingPath), 'drawing' + number++ + '.xml'); } while (model.files[copyPath]);
      addedDrawingPath = copyPath;
      write(copyPath, nextDrawing);
      write(path.join(path.dirname(copyPath), '_rels', path.basename(copyPath) + '.rels'), nextRelations);
      const range = findElementRanges(sheetRelXml, 'Relationship', 'Id').get(drawingId);
      const tag = setTagAttribute(sheetRelXml.slice(range.start, range.startTagEnd), 'Target', path.relative(path.dirname(sheet.path), copyPath));
      write(sheetRelPath, sheetRelXml.slice(0, range.start) + tag + sheetRelXml.slice(range.startTagEnd));
    } else { write(drawingPath, nextDrawing); write(drawingRelPath, nextRelations); }
  }
  if (!packageReferences(outputFiles, chartPath).length) remove(chartPath);
  const typesPath = '[Content_Types].xml';
  let typesXml = new TextDecoder().decode(model.files[typesPath]);
  for (const name of removedParts) {
    const range = findElementRanges(typesXml, 'Override', 'PartName').get('/' + name);
    if (range) typesXml = typesXml.slice(0, range.start) + typesXml.slice(range.end);
  }
  if (addedDrawingPath) typesXml = typesXml.replace('</Types>', '<Override PartName="/' + addedDrawingPath + '" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>');
  if (typesXml !== new TextDecoder().decode(model.files[typesPath])) write(typesPath, typesXml);
  const bytes = zipSync(outputFiles, { level: 0 }), checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('차트 삭제 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  const expected = sheet.charts.filter((item, index) => index !== chartIndex);
  if (JSON.stringify(checked.sheets[sheetIndex].charts) !== JSON.stringify(expected)) fail('차트 삭제 후 남은 차트가 달라졌습니다.');
  for (let index = 0; index < model.sheets.length; index++) if (index !== sheetIndex && JSON.stringify(checked.sheets[index].charts) !== JSON.stringify(model.sheets[index].charts)) fail('차트 삭제 중 다른 시트의 차트가 달라졌습니다.');
  for (const [name, before] of Object.entries(model.files)) if (!modified.has(name) && Buffer.compare(Buffer.from(before), Buffer.from(checked.files[name])) !== 0) fail('차트 삭제 중 비수정 패키지 항목이 달라졌습니다: ' + name);
  return { bytes, model: checked };
}

function replaceChartSeriesData(xml, generatedXml, scatter) {
  const generated = parseXml(new TextEncoder().encode(generatedXml), 'chart-series').root;
  for (const name of scatter ? ['tx', 'xVal', 'yVal'] : ['tx', 'cat', 'val']) {
    const root = parseXml(new TextEncoder().encode(xml), 'chart-series').root;
    const existing = childrenByLocalName(root, name)[0], replacementNode = childrenByLocalName(generated, name)[0];
    const replacementRange = findElementRanges(generatedXml, replacementNode.name, undefined, true).values().next().value;
    const replacement = generatedXml.slice(replacementRange.start, replacementRange.end);
    if (existing) {
      const range = findElementRanges(xml, existing.name, undefined, true).values().next().value;
      if (!range) fail('차트 계열 데이터 위치를 찾을 수 없습니다.');
      xml = xml.slice(0, range.start) + replacement + xml.slice(range.end);
    } else {
      const next = root.children.find(node => name === 'tx' ? !['idx', 'order'].includes(localName(node)) : (name === 'cat' || name === 'xVal' ? ['val', 'yVal', 'smooth', 'shape', 'extLst'] : ['smooth', 'shape', 'extLst']).includes(localName(node)));
      const position = next ? findElementRanges(xml, next.name, undefined, true).values().next().value?.start : findElementRanges(xml, root.name, undefined, true).values().next().value?.closingStart;
      if (!Number.isInteger(position)) fail('차트 계열 데이터를 삽입할 위치를 찾을 수 없습니다.');
      xml = xml.slice(0, position) + replacement + xml.slice(position);
    }
  }
  return xml;
}

function patchChartSource(model, sheetIndex, chartIndex, range) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  const sheet = model.sheets[sheetIndex], current = sheet?.charts?.[chartIndex], chartPath = sheet?.chartPaths?.[chartIndex];
  if (!current || !chartPath) fail('수정할 차트를 찾을 수 없습니다.');
  const { xml, root } = parseXml(model.files[chartPath], chartPath);
  const plot = childrenByLocalName(childrenByLocalName(root, 'chart')[0], 'plotArea')[0];
  const chart = plot?.children.find(node => ['barChart', 'lineChart', 'areaChart', 'scatterChart', 'pieChart', 'doughnutChart'].includes(localName(node)));
  const series = childrenByLocalName(chart, 'ser'), prefix = chart?.name.includes(':') ? chart.name.split(':')[0] + ':' : '';
  if (!series.length) fail('차트 계열을 찾을 수 없습니다.');
  const generated = buildChartSeries(model, sheetIndex, range, current.type, prefix);
  const ranges = [...findElementRanges(xml, series[0].name, undefined, true).values()];
  if (ranges.length !== series.length) fail('차트 계열 위치를 안전하게 해석할 수 없습니다.');
  let nextIndex = Math.max(...series.map(node => Number(childrenByLocalName(node, 'idx')[0]?.attributes.val) || 0)) + 1;
  let nextOrder = Math.max(...series.map(node => Number(childrenByLocalName(node, 'order')[0]?.attributes.val) || 0)) + 1;
  const extra = generated.slice(series.length).map((item, index) => item.replace('idx val="' + (series.length + index) + '"', 'idx val="' + nextIndex++ + '"').replace('order val="' + (series.length + index) + '"', 'order val="' + nextOrder++ + '"')).join('');
  let outputXml = xml;
  for (let index = ranges.length - 1; index >= 0; index--) {
    const item = ranges[index];
    const replacement = index < generated.length ? replaceChartSeriesData(xml.slice(item.start, item.end), generated[index], current.type === 'scatter') : '';
    outputXml = outputXml.slice(0, item.start) + replacement + (index === ranges.length - 1 ? extra : '') + outputXml.slice(item.end);
  }
  if (outputXml === xml) return { bytes: model.bytes, model };
  const outputFiles = { ...model.files, [chartPath]: new TextEncoder().encode(outputXml) };
  const bytes = zipSync(outputFiles, { level: 0 }), checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('차트 원본 변경 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  const edited = checked.sheets[sheetIndex].charts[chartIndex];
  if (edited.type !== current.type || edited.title !== current.title || edited.direction !== current.direction || edited.series.length !== generated.length) fail('차트 원본 변경 후 차트 구성이 달라졌습니다.');
  for (const [name, before] of Object.entries(model.files)) if (name !== chartPath && Buffer.compare(Buffer.from(before), Buffer.from(checked.files[name])) !== 0) fail('차트 원본 변경 중 다른 패키지 항목이 달라졌습니다: ' + name);
  return { bytes, model: checked };
}

function chartTitleXml(prefix, title) {
  return '<' + prefix + 'title><' + prefix + 'tx><' + prefix + 'rich xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t xml:space="preserve">' + xmlEscape(title) + '</a:t></a:r></a:p></' + prefix + 'rich></' + prefix + 'tx></' + prefix + 'title>';
}

function patchChartPlacement(model, sheetIndex, chartIndex, placement) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  const sheet = model.sheets[sheetIndex], drawingPath = sheet?.chartDrawingPath;
  if (!Number.isInteger(sheetIndex) || !sheet || !drawingPath || !Number.isInteger(chartIndex) || !sheet.charts[chartIndex] || !['twoCell', 'oneCell', 'absolute'].includes(placement)) fail('차트 배치 요청이 유효하지 않습니다.');
  const parsed = parseXml(model.files[drawingPath], drawingPath), entries = [];
  for (const anchor of parsed.root.children) {
    const refs = [];
    (function walk(node) { if (localName(node) === 'chart') refs.push(node); for (const item of node.children) walk(item); })(anchor);
    for (const ref of refs) entries.push({ anchor, ref, count: refs.length });
  }
  const selected = entries[chartIndex], anchor = selected?.anchor, geometry = anchor && readChartAnchor(anchor, true);
  if (!geometry || selected.count !== 1) fail('차트 배치를 안전하게 해석할 수 없습니다.');
  const current = geometry.type === 'twoCellAnchor' ? geometry.editAs || 'twoCell' : geometry.type === 'oneCellAnchor' ? 'oneCell' : 'absolute';
  if (current === placement) return { bytes: model.bytes, model };
  const axes = drawingAxes({ ...sheet, hiddenRows: sheet.nativeHiddenRows || sheet.hiddenRows });
  const from = geometry.type === 'absoluteAnchor' ? { x: geometry.x, y: geometry.y } : drawingPoint(geometry.from, axes);
  const to = geometry.type === 'twoCellAnchor' ? drawingPoint(geometry.to, axes) : { x: from.x + geometry.width, y: from.y + geometry.height };
  if (!(to.x > from.x && to.y > from.y)) fail('차트 배치를 바꾸려면 유효한 차트 크기가 필요합니다.');
  const siblings = parsed.root.children.filter(node => node.name === anchor.name), range = [...findElementRanges(parsed.xml, anchor.name, undefined, true).values()][siblings.indexOf(anchor)];
  if (!range) fail('차트 배치의 XML 위치를 찾을 수 없습니다.');
  let inner = parsed.xml.slice(range.startTagEnd, range.closingStart);
  for (const node of anchor.children.filter(node => ['from', 'to', 'pos', 'ext'].includes(localName(node)))) {
    const ranges = [...findElementRanges(inner, node.name, undefined, true).values()];
    if (ranges.length !== 1) fail('차트 배치 좌표가 중복되어 변경할 수 없습니다.');
    inner = inner.slice(0, ranges[0].start) + inner.slice(ranges[0].end);
  }
  const prefix = anchor.name.includes(':') ? anchor.name.slice(0, anchor.name.indexOf(':') + 1) : '', name = prefix + 'twoCellAnchor';
  const opening = setTagAttribute(parsed.xml.slice(range.start, range.startTagEnd).replace('<' + anchor.name, '<' + name), 'editAs', placement);
  const marker = (name, point) => { const value = drawingMarker(point, axes); return '<' + prefix + name + '><' + prefix + 'col>' + value.column + '</' + prefix + 'col><' + prefix + 'colOff>' + Math.round(value.x * 9525) + '</' + prefix + 'colOff><' + prefix + 'row>' + value.row + '</' + prefix + 'row><' + prefix + 'rowOff>' + Math.round(value.y * 9525) + '</' + prefix + 'rowOff></' + prefix + name + '>'; };
  const drawingXml = parsed.xml.slice(0, range.start) + opening + marker('from', from) + marker('to', to) + inner + '</' + name + '>' + parsed.xml.slice(range.end);
  const outputFiles = { ...model.files }, modifiedPaths = new Set();
  writeWorksheetDrawing(sheet, drawingPath, outputFiles, modifiedPaths, drawingXml);
  const bytes = zipSync(outputFiles, { level: 0 }), checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('차트 배치 저장 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  for (const [name, before] of Object.entries(model.files)) if (!modifiedPaths.has(name) && Buffer.compare(Buffer.from(before), Buffer.from(checked.files[name])) !== 0) fail('차트 배치 저장 중 비수정 패키지 항목이 달라졌습니다: ' + name);
  return { bytes, model: checked };
}

function patchChartMetadata(model, sheetIndex, chartIndex, changes) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  const sheet = model.sheets[sheetIndex], chartPath = sheet?.chartPaths?.[chartIndex], current = sheet?.charts?.[chartIndex];
  if (!sheet || !chartPath || !current || !changes || typeof changes !== 'object' || Array.isArray(changes)) fail('수정할 차트를 찾을 수 없습니다.');
  if (Object.hasOwn(changes, 'title') && (typeof changes.title !== 'string' || changes.title.length > 512)) fail('차트 제목은 512자 이내의 텍스트여야 합니다.');
  if (Object.hasOwn(changes, 'seriesNames') && (!Array.isArray(changes.seriesNames) || changes.seriesNames.length !== current.series.length || changes.seriesNames.some(name => typeof name !== 'string' || name.length > 256))) fail('계열 이름이 차트 구조와 일치하지 않거나 256자를 초과했습니다.');
  if (!Object.hasOwn(changes, 'title') && !Object.hasOwn(changes, 'seriesNames')) fail('변경할 차트 제목 또는 계열 이름이 없습니다.');
  const titleChanged = Object.hasOwn(changes, 'title') && changes.title !== current.title;
  const namesChanged = Object.hasOwn(changes, 'seriesNames') && changes.seriesNames.some((name, index) => name !== current.series[index].name);
  if (!titleChanged && !namesChanged) return { bytes: model.bytes, model };
  const { xml, root } = parseXml(model.files[chartPath], chartPath), chart = childrenByLocalName(root, 'chart')[0], prefix = chart?.name.includes(':') ? chart.name.slice(0, chart.name.indexOf(':') + 1) : '';
  if (!chart) fail('차트 구조를 수정할 수 없습니다.');
  let outputXml = xml;
  if (titleChanged) {
    const title = childrenByLocalName(chart, 'title')[0], titleRange = findElementRanges(xml, prefix + 'title', undefined, true).values().next().value;
    if (title && !titleRange) fail('차트 제목 위치를 찾을 수 없습니다.');
    if (changes.title === '') {
      if (title) outputXml = outputXml.slice(0, titleRange.start) + outputXml.slice(titleRange.end);
    } else {
      const titleXml = chartTitleXml(prefix, changes.title);
      if (title) outputXml = outputXml.slice(0, titleRange.start) + titleXml + outputXml.slice(titleRange.end);
      else {
        const plot = childrenByLocalName(chart, 'plotArea')[0], plotRange = plot && findElementRanges(outputXml, prefix + 'plotArea', undefined, true).values().next().value;
        if (!plotRange) fail('차트 제목을 삽입할 위치를 찾을 수 없습니다.');
        outputXml = outputXml.slice(0, plotRange.start) + titleXml + outputXml.slice(plotRange.start);
      }
    }
  }
  if (namesChanged) {
    const rootAfterTitle = parseXml(new TextEncoder().encode(outputXml), chartPath).root;
    const plot = childrenByLocalName(childrenByLocalName(rootAfterTitle, 'chart')[0], 'plotArea')[0];
    const chartNode = plot?.children.find(node => ['barChart', 'lineChart', 'areaChart', 'scatterChart', 'pieChart', 'doughnutChart'].includes(localName(node)));
    const series = chartNode ? childrenByLocalName(chartNode, 'ser') : [];
    if (series.length !== changes.seriesNames.length) fail('계열 이름 개수가 차트 계열과 일치하지 않습니다.');
    let seriesRanges = [...findElementRanges(outputXml, prefix + 'ser', undefined, true).values()];
    if (seriesRanges.length !== series.length) fail('차트 계열 위치를 찾을 수 없습니다.');
    for (let index = seriesRanges.length - 1; index >= 0; index--) {
      if (changes.seriesNames[index] === current.series[index].name) continue;
      const itemRange = seriesRanges[index], itemXml = outputXml.slice(itemRange.start, itemRange.end), item = parseXml(new TextEncoder().encode(itemXml), chartPath).root;
      const title = childrenByLocalName(item, 'tx')[0], titleRange = findElementRanges(itemXml, prefix + 'tx', undefined, true).values().next().value;
      const nameXml = '<' + prefix + 'tx><' + prefix + 'v>' + xmlEscape(changes.seriesNames[index]) + '</' + prefix + 'v></' + prefix + 'tx>';
      let nextItem;
      if (title) {
        if (!titleRange) fail('계열 이름 위치를 찾을 수 없습니다.');
        nextItem = itemXml.slice(0, titleRange.start) + nameXml + itemXml.slice(titleRange.end);
      } else {
        const order = childrenByLocalName(item, 'order')[0] || childrenByLocalName(item, 'idx')[0];
        const orderName = order?.name;
        const orderRange = orderName && findElementRanges(itemXml, orderName, undefined, true).values().next().value;
        if (!orderRange) fail('차트 계열 이름을 삽입할 위치를 찾을 수 없습니다.');
        nextItem = itemXml.slice(0, orderRange.end) + nameXml + itemXml.slice(orderRange.end);
      }
      outputXml = outputXml.slice(0, itemRange.start) + nextItem + outputXml.slice(itemRange.end);
    }
  }
  const outputFiles = Object.assign({}, model.files);
  outputFiles[chartPath] = new TextEncoder().encode(outputXml);
  const bytes = zipSync(outputFiles, { level: 0 }), checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('차트 정보 수정 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  const edited = checked.sheets[sheetIndex]?.charts?.[chartIndex];
  if (!edited || Object.hasOwn(changes, 'title') && edited.title !== changes.title || Object.hasOwn(changes, 'seriesNames') && edited.series.some((item, index) => item.name !== (changes.seriesNames[index] || 'Series ' + (index + 1)))) fail('차트 정보 수정 후 재분석 결과가 요청과 다릅니다.');
  for (const [name, before] of Object.entries(model.files)) if (name !== chartPath && (!checked.files[name] || Buffer.compare(Buffer.from(before), Buffer.from(checked.files[name])) !== 0)) fail('차트 정보 수정 중 다른 패키지 항목이 달라졌습니다: ' + name);
  return { bytes, model: checked };
}

function renameFormulaSheetReferences(formula, previousName, nextName) {
  let output = '', index = 0;
  const previous = previousName.toLocaleLowerCase(), quotedName = nextName.replaceAll("'", "''");
  while (index < formula.length) {
    if (formula[index] === '"') {
      const start = index++;
      while (index < formula.length) {
        if (formula.startsWith('""', index)) { index += 2; continue; }
        if (formula[index++] === '"') break;
      }
      output += formula.slice(start, index);
      continue;
    }
    if (formula[index] === '[') {
      const start = index++;
      let depth = 1;
      while (index < formula.length && depth) { if (formula[index] === "'" && /['\[\]#@]/.test(formula[index+1]||'')) { index+=2; continue; } if (formula[index] === '[') depth++; else if (formula[index] === ']') depth--; index++; }
      output += formula.slice(start, index);
      continue;
    }
    if (formula[index] === "'") {
      const start = index++; let name = '';
      while (index < formula.length) {
        if (formula.startsWith("''", index)) { name += "'"; index += 2; continue; }
        if (formula[index] === "'") break;
        name += formula[index++];
      }
      if (formula[index] === "'" && formula[index + 1] === '!') {
        index += 2;
        output += name.toLocaleLowerCase() === previous ? "'" + quotedName + "'!" : formula.slice(start, index);
      } else output += formula.slice(start, index);
      continue;
    }
    const prefix = /^[A-Za-z_][A-Za-z0-9_.]*!/.exec(formula.slice(index));
    const previousChar = formula[index - 1] || '';
    if (prefix && !/[A-Za-z0-9_.]/.test(previousChar) && prefix[0].slice(0, -1).toLocaleLowerCase() === previous) {
      output += "'" + quotedName + "'!";
      index += prefix[0].length;
      continue;
    }
    output += formula[index++];
  }
  return output;
}

function deleteFormulaSheetReferences(formula, targetName) {
  let output = '', index = 0, unsafe = false;
  const target = targetName.toLocaleLowerCase();
  const coordinate = start => /^(?:\$?[A-Za-z]{1,3}\$?[1-9]\d*|\$?[A-Za-z]{1,3}|\$?[1-9]\d*)/.exec(formula.slice(start));
  while (index < formula.length) {
    if (formula[index] === '"') {
      const start = index++;
      while (index < formula.length) {
        if (formula.startsWith('""', index)) { index += 2; continue; }
        if (formula[index++] === '"') break;
      }
      output += formula.slice(start, index);
      continue;
    }
    if (formula[index] === '[') {
      const start = index++;
      let depth = 1;
      while (index < formula.length && depth) { if (formula[index] === "'" && /['\[\]#@]/.test(formula[index+1]||'')) { index+=2; continue; } if (formula[index] === '[') depth++; else if (formula[index] === ']') depth--; index++; }
      output += formula.slice(start, index);
      continue;
    }
    let sheetName = null, referenceStart = index, referenceAt = index;
    if (formula[index] === "'") {
      let name = '', end = index + 1;
      while (end < formula.length) {
        if (formula.startsWith("''", end)) { name += "'"; end += 2; continue; }
        if (formula[end] === "'") break;
        name += formula[end++];
      }
      if (formula[end] === "'" && formula[end + 1] === '!') { sheetName = name; referenceAt = end + 2; }
    } else {
      const prefix = /^[A-Za-z_][A-Za-z0-9_.]*!/.exec(formula.slice(index));
      if (prefix && !/[A-Za-z0-9_.]/.test(formula[index - 1] || '')) { sheetName = prefix[0].slice(0, -1); referenceAt = index + prefix[0].length; }
    }
    if (sheetName?.toLocaleLowerCase() === target) {
      const first = coordinate(referenceAt);
      if (first) {
        let referenceEnd = referenceAt + first[0].length;
        if (formula[referenceEnd] === ':') {
          let secondAt = referenceEnd + 1, secondSheet = null;
          if (formula[secondAt] === "'") {
            let name = '', end = secondAt + 1;
            while (end < formula.length) {
              if (formula.startsWith("''", end)) { name += "'"; end += 2; continue; }
              if (formula[end] === "'") break;
              name += formula[end++];
            }
            if (formula[end] === "'" && formula[end + 1] === '!') { secondSheet = name; secondAt = end + 2; }
          } else {
            const prefix = /^[A-Za-z_][A-Za-z0-9_.]*!/.exec(formula.slice(secondAt));
            if (prefix) { secondSheet = prefix[0].slice(0, -1); secondAt += prefix[0].length; }
          }
          const second = coordinate(secondAt);
          if (second && (!secondSheet || secondSheet.toLocaleLowerCase() === target)) referenceEnd = secondAt + second[0].length;
          else if (secondSheet && secondSheet.toLocaleLowerCase() !== target) unsafe = true;
        }
        const next = formula[referenceEnd] || '';
        if (/[A-Za-z0-9_.]/.test(next) || next === '(') { output += formula[index++]; continue; }
        output += '#REF!';
        index = referenceEnd;
        continue;
      }
    }
    output += formula[index++];
  }
  return { formula: output, unsafe };
}

function patchFormulaTextElements(xml, previousName, nextName) {
  const decode = value => value.replaceAll('&quot;', '"').replaceAll('&apos;', "'").replaceAll('&lt;', '<').replaceAll('&gt;', '>').replaceAll('&amp;', '&');
  return xml.replace(/<((?:[\w.-]+:)?(?:f|formula|formula1|formula2|definedName))(\s[^>]*)?>([\s\S]*?)<\/\1>/g, (element, name, attributes = '', body) => {
    const updated = renameFormulaSheetReferences(decode(body), previousName, nextName);
    return updated === decode(body) ? element : '<' + name + attributes + '>' + xmlEscape(updated) + '</' + name + '>';
  });
}

function patchSheetVisibility(model, sheetIndex, visible) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  if (!Number.isInteger(sheetIndex) || sheetIndex < 0 || sheetIndex >= model.sheets.length || typeof visible !== 'boolean') fail('워크시트 표시 상태 변경 요청이 유효하지 않습니다.');
  const sheet = model.sheets[sheetIndex];
  if (sheet.visible === visible) return { bytes: model.bytes, model };
  if (!visible && model.sheets.filter(item => item.visible).length <= 1) fail('통합문서에는 표시된 워크시트가 하나 이상 있어야 합니다.');

  const workbookPath = 'xl/workbook.xml';
  const workbookXml = new TextDecoder().decode(model.files[workbookPath]);
  const sheetTags = [...workbookXml.matchAll(/<sheet\b[^>]*>/g)];
  if (sheetTags.length !== model.sheets.length) fail('워크시트 표시 상태를 안전하게 변경할 수 없습니다.');
  const match = sheetTags[sheetIndex];
  const opening = match[0];
  const updatedTag = replaceTagAttribute(opening, 'state', visible ? 'visible' : 'hidden');
  const updatedXml = workbookXml.slice(0, match.index) + updatedTag + workbookXml.slice(match.index + opening.length);
  const outputFiles = { ...model.files, [workbookPath]: new TextEncoder().encode(updatedXml) };
  const bytes = zipSync(outputFiles, { level: 0 }), checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('워크시트 표시 상태 변경 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  if (checked.sheets[sheetIndex]?.visible !== visible) fail('워크시트 표시 상태를 검증하지 못했습니다.');
  return { bytes, model: checked };
}

function patchSheetZoom(model, sheetIndex, zoomScale) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  if (!Number.isInteger(sheetIndex) || sheetIndex < 0 || sheetIndex >= model.sheets.length || !Number.isInteger(zoomScale) || zoomScale < 10 || zoomScale > 400) fail('워크시트 확대 비율 변경 요청이 유효하지 않습니다.');
  const sheet = model.sheets[sheetIndex];
  if (sheet.zoomScale === zoomScale) return { bytes: model.bytes, model };

  const worksheetPath = sheet.path;
  const worksheetXml = new TextDecoder().decode(model.files[worksheetPath]);
  const updateZoomAttribute = tag => {
    const attributes = parseTagAttributes(tag);
    if (Object.hasOwn(attributes, 'zoomScale')) return replaceTagAttribute(tag, 'zoomScale', zoomScale);
    const element = /^<([^\s/>]+)/.exec(tag);
    if (!element) fail('OOXML 시트 뷰 속성을 수정할 수 없습니다.');
    return tag.replace(/^<([^\s/>]+)/, '<$1 zoomScale="' + zoomScale + '"');
  };
  let updatedXml = null;
  const sheetViews = /<(?:[\w.-]+:)?sheetViews\b[^>]*>[\s\S]*?<\/(?:[\w.-]+:)?sheetViews\s*>/.exec(worksheetXml);
  if (sheetViews) {
    const view = /<(?:[\w.-]+:)?sheetView\b[^>]*?\/>|<(?:[\w.-]+:)?sheetView\b[^>]*>/.exec(sheetViews[0]);
    if (view) {
      const start = sheetViews.index + view.index;
      updatedXml = worksheetXml.slice(0, start) + updateZoomAttribute(view[0]) + worksheetXml.slice(start + view[0].length);
    } else {
      const containerName = /^<((?:[\w.-]+:)?sheetViews)\b/.exec(sheetViews[0])?.[1];
      if (!containerName) fail('OOXML 시트 뷰를 수정할 수 없습니다.');
      const prefix = containerName.includes(':') ? containerName.slice(0, containerName.indexOf(':') + 1) : '';
      const closeOffset = sheetViews[0].lastIndexOf('</');
      if (closeOffset < 0) fail('OOXML 시트 뷰를 수정할 수 없습니다.');
      const viewXml = '<' + prefix + 'sheetView workbookViewId="0" zoomScale="' + zoomScale + '"/>';
      const start = sheetViews.index + closeOffset;
      updatedXml = worksheetXml.slice(0, start) + viewXml + worksheetXml.slice(start);
    }
  } else {
    const selfClosing = /<(?:[\w.-]+:)?sheetViews\b[^>]*\/\s*>/.exec(worksheetXml);
    if (selfClosing) {
      const containerName = /^<((?:[\w.-]+:)?sheetViews)\b/.exec(selfClosing[0])?.[1];
      if (!containerName) fail('OOXML 시트 뷰를 수정할 수 없습니다.');
      const prefix = containerName.includes(':') ? containerName.slice(0, containerName.indexOf(':') + 1) : '';
      const replacement = '<' + containerName + '><' + prefix + 'sheetView workbookViewId="0" zoomScale="' + zoomScale + '"/></' + containerName + '>';
      updatedXml = worksheetXml.slice(0, selfClosing.index) + replacement + worksheetXml.slice(selfClosing.index + selfClosing[0].length);
    } else {
      const rootName = /<((?:[\w.-]+:)?worksheet)\b/.exec(worksheetXml)?.[1];
      if (!rootName) fail('OOXML 워크시트 루트를 찾을 수 없습니다.');
      const prefix = rootName.includes(':') ? rootName.slice(0, rootName.indexOf(':') + 1) : '';
      const viewXml = '<' + prefix + 'sheetViews><' + prefix + 'sheetView workbookViewId="0" zoomScale="' + zoomScale + '"/></' + prefix + 'sheetViews>';
      const nextElement = /<(?:[\w.-]+:)?(?:sheetFormatPr|cols|sheetData)\b/.exec(worksheetXml);
      const worksheetClose = /<\/(?:[\w.-]+:)?worksheet\s*>/.exec(worksheetXml);
      const start = nextElement?.index ?? worksheetClose?.index;
      if (start === undefined) fail('OOXML 워크시트 뷰를 안전하게 추가할 수 없습니다.');
      updatedXml = worksheetXml.slice(0, start) + viewXml + worksheetXml.slice(start);
    }
  }

  const outputFiles = { ...model.files, [worksheetPath]: new TextEncoder().encode(updatedXml) };
  const bytes = zipSync(outputFiles, { level: 0 }), checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('워크시트 확대 비율 변경 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  if (checked.sheets[sheetIndex]?.zoomScale !== zoomScale) fail('워크시트 확대 비율을 검증하지 못했습니다.');
  return { bytes, model: checked };
}

function patchSheetName(model, sheetIndex, name) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  if (!Number.isInteger(sheetIndex) || !model.sheets[sheetIndex] || typeof name !== 'string') fail('시트 이름 변경 요청이 유효하지 않습니다.');
  const nextName = name.trim(), sheet = model.sheets[sheetIndex];
  if (model.sheets.some(item => item.hyperlinks?.some(link => link.location))) fail('내부 하이퍼링크가 있는 통합문서에서는 시트 이름을 변경할 수 없습니다.');
  if (!nextName || Array.from(nextName).length > 31 || /[\\/?*:[\]]/.test(nextName) || /^'|'$/.test(nextName) || /[\x00-\x1f]/.test(nextName)) fail('시트 이름은 1~31자이며 \u005c\u005c / ? * [ ] : 문자를 사용할 수 없습니다.');
  if (model.sheets.some((item, index) => index !== sheetIndex && item.name.toLocaleLowerCase() === nextName.toLocaleLowerCase())) fail('같은 이름의 시트가 이미 있습니다.');
  if (nextName === sheet.name) return { bytes: model.bytes, model };
  const outputFiles = Object.assign({}, model.files), modifiedPaths = new Set();
  let workbookXml = new TextDecoder().decode(outputFiles['xl/workbook.xml']);
  const sheetTags = [...workbookXml.matchAll(/<sheet\b[^>]*\/?\s*>/g)];
  if (sheetTags.length !== model.sheets.length || !sheetTags[sheetIndex] || parseTagAttributes(sheetTags[sheetIndex][0]).name !== sheet.name) fail('통합문서 시트 목록을 안전하게 갱신할 수 없습니다.');
  const selectedTag = sheetTags[sheetIndex][0];
  workbookXml = workbookXml.replace(selectedTag, setTagAttribute(selectedTag, 'name', nextName));
  workbookXml = patchFormulaTextElements(workbookXml, sheet.name, nextName);
  if (sheet.printAreas?.length) workbookXml = replaceWorkbookPrintDefinition(workbookXml, sheetIndex, sheet.printAreas.map(range => printAreaReference(nextName, range)).join(','));
  if (sheet.printTitles?.rows || sheet.printTitles?.columns) workbookXml = replaceWorkbookPrintDefinition(workbookXml, sheetIndex, printTitlesReference(nextName, sheet.printTitles), '_xlnm.Print_Titles');
  outputFiles['xl/workbook.xml'] = new TextEncoder().encode(enableFullCalculationOnLoad(workbookXml));
  modifiedPaths.add('xl/workbook.xml');
  for (const item of model.sheets) {
    const original = new TextDecoder().decode(outputFiles[item.path]);
    const updated = patchFormulaTextElements(original, sheet.name, nextName);
    if (updated !== original) { outputFiles[item.path] = new TextEncoder().encode(updated); modifiedPaths.add(item.path); }
  }
  for (const chart of model.chartPaths || []) {
    const original = new TextDecoder().decode(outputFiles[chart.path]);
    const updated = patchFormulaTextElements(original, sheet.name, nextName);
    if (updated !== original) { outputFiles[chart.path] = new TextEncoder().encode(updated); modifiedPaths.add(chart.path); }
  }
  const bytes = zipSync(outputFiles, { level: 0 }), checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('시트 이름 변경 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  for (const [path, before] of Object.entries(model.files)) {
    if (modifiedPaths.has(path)) continue;
    if (!checked.files[path] || Buffer.compare(Buffer.from(before), Buffer.from(checked.files[path])) !== 0) fail('시트 이름 변경 중 비수정 패키지 항목이 달라졌습니다: ' + path);
  }
  return { bytes, model: checked };
}

function patchAddSheet(model) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  if (model.sheets.length >= LIMITS.sheets) fail('워크시트 개수 제한에 도달했습니다.');
  const existingNames = new Set(model.sheets.map(sheet => sheet.name.toLocaleLowerCase()));
  let number = model.sheets.length + 1, name = `Sheet${number}`;
  while (existingNames.has(name.toLocaleLowerCase())) name = `Sheet${++number}`;
  const outputFiles = Object.assign({}, model.files), modifiedPaths = new Set();
  const workbookXml = new TextDecoder().decode(outputFiles['xl/workbook.xml']);
  const sheetList = findElementRanges(workbookXml, 'sheets').get('');
  if (!sheetList || sheetList.selfClosing) fail('통합 문서 시트 목록을 안전하게 갱신할 수 없습니다.');
  const sheetTags = [...workbookXml.matchAll(/<sheet\b[^>]*\/?\s*>/g)], sheetIds = sheetTags.map(match => Number(parseTagAttributes(match[0]).sheetId));
  if (sheetTags.length !== model.sheets.length || sheetIds.some(id => !Number.isInteger(id) || id < 1) || new Set(sheetIds).size !== sheetIds.length) fail('워크시트 ID 목록이 유효하지 않아 새 시트를 추가할 수 없습니다.');
  const sheetId = Math.max(0, ...sheetIds) + 1;
  const relsPath = 'xl/_rels/workbook.xml.rels', relationshipsXml = new TextDecoder().decode(outputFiles[relsPath]);
  const relationshipTags = [...relationshipsXml.matchAll(/<Relationship\b[^>]*\/?\s*>/g)], relationshipIds = new Set(relationshipTags.map(match => parseTagAttributes(match[0]).Id));
  let relationshipNumber = 1, relationshipId = `rId${relationshipNumber}`;
  while (relationshipIds.has(relationshipId)) relationshipId = `rId${++relationshipNumber}`;
  let fileNumber = 1, sheetPath = `xl/worksheets/sheet${fileNumber}.xml`;
  while (Object.hasOwn(outputFiles, sheetPath)) sheetPath = `xl/worksheets/sheet${++fileNumber}.xml`;
  const target = sheetPath.slice('xl/'.length);
  const closingSheetList = workbookXml.indexOf('</sheets>', sheetList.closingStart);
  if (closingSheetList < 0) fail('워크시트 목록 끝 태그를 찾을 수 없습니다.');
  const sheetTag = '<sheet name="' + xmlEscape(name) + '" sheetId="' + sheetId + '" r:id="' + relationshipId + '"/>';
  let updatedWorkbook = workbookXml.slice(0, closingSheetList) + sheetTag + workbookXml.slice(closingSheetList);
  const workbookViews = [...findElementRanges(updatedWorkbook, 'workbookView').values()];
  if (workbookViews.length) {
    if (workbookViews.length > 1) fail('여러 통합 문서 보기를 안전하게 변경할 수 없습니다.');
    const view = workbookViews[0], opening = updatedWorkbook.slice(view.start, view.startTagEnd);
    updatedWorkbook = updatedWorkbook.slice(0, view.start) + setTagAttribute(opening, 'activeTab', model.sheets.length) + updatedWorkbook.slice(view.startTagEnd);
  } else {
    const bookViews = '<bookViews><workbookView activeTab="' + model.sheets.length + '"/></bookViews>';
    updatedWorkbook = updatedWorkbook.slice(0, sheetList.start) + bookViews + updatedWorkbook.slice(sheetList.start);
  }
  const relClosing = relationshipsXml.lastIndexOf('</Relationships>');
  if (relClosing < 0) fail('통합문서 관계 목록 끝 태그를 찾을 수 없습니다.');
  const relationshipTag = '<Relationship Id="' + relationshipId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="' + target + '"/>';
  const updatedRelationships = relationshipsXml.slice(0, relClosing) + relationshipTag + relationshipsXml.slice(relClosing);
  const contentTypesXml = new TextDecoder().decode(outputFiles['[Content_Types].xml']), contentTypesClosing = contentTypesXml.lastIndexOf('</Types>');
  if (contentTypesClosing < 0) fail('콘텐츠 형식 목록 끝 태그를 찾을 수 없습니다.');
  if (contentTypesXml.includes('PartName="/' + sheetPath + '"')) fail('새 워크시트의 콘텐츠 형식 항목이 이미 있습니다.');
  const contentType = '<Override PartName="/' + sheetPath + '" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
  outputFiles['xl/workbook.xml'] = new TextEncoder().encode(enableFullCalculationOnLoad(updatedWorkbook));
  outputFiles[relsPath] = new TextEncoder().encode(updatedRelationships);
  outputFiles['[Content_Types].xml'] = new TextEncoder().encode(contentTypesXml.slice(0, contentTypesClosing) + contentType + contentTypesXml.slice(contentTypesClosing));
  outputFiles[sheetPath] = new TextEncoder().encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>');
  modifiedPaths.add('xl/workbook.xml'); modifiedPaths.add(relsPath); modifiedPaths.add('[Content_Types].xml'); modifiedPaths.add(sheetPath);
  const bytes = zipSync(outputFiles, { level: 0 }), checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('새 시트 추가 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  for (const [path, before] of Object.entries(model.files)) {
    if (modifiedPaths.has(path)) continue;
    if (!checked.files[path] || Buffer.compare(Buffer.from(before), Buffer.from(checked.files[path])) !== 0) fail('새 시트 추가 중 비수정 패키지 항목이 달라졌습니다: ' + path);
  }
  return { bytes, model: checked, sheetIndex: checked.sheets.length - 1 };
}

function patchDeleteSheet(model, sheetIndex) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  model = preserveImportedFilterVisibility(model);
  if (model.hasCharts) fail('차트가 있는 통합문서에서는 시트를 삭제할 수 없습니다.');
  if (!Number.isInteger(sheetIndex) || !model.sheets[sheetIndex]) fail('삭제할 워크시트가 유효하지 않습니다.');
  if (model.sheets.length <= 1) fail('통합 문서의 마지막 시트는 삭제할 수 없습니다.');
  const sheet = model.sheets[sheetIndex];
  if (model.sheets.some(item => item.hyperlinks?.some(link => link.location))) fail('내부 하이퍼링크가 있는 통합문서에서는 시트를 삭제할 수 없습니다.');
  let workbookXml = new TextDecoder().decode(model.files['xl/workbook.xml']);
  const outputFiles = Object.assign({}, model.files), modifiedPaths = new Set();
  const decode = value => value.replaceAll('&quot;', '"').replaceAll('&apos;', "'").replaceAll('&lt;', '<').replaceAll('&gt;', '>').replaceAll('&amp;', '&');
  const references = formula => deleteFormulaSheetReferences(formula, sheet.name);
  const hasReference = formula => references(formula).formula !== formula;
  const rewriteXml = (xml, path, { rejectNonCell = false, workbook = false } = {}) => xml.replace(/<(f|formula|formula1|formula2|definedName)(\s[^>]*)?>([\s\S]*?)<\/\1>/g, (element, tag, attributes = '', body) => {
    const formula = decode(body), result = references(formula);
    if (result.unsafe || (rejectNonCell && tag !== 'f' && hasReference(formula))) fail('삭제할 시트를 참조하는 데이터 유효성·조건부 서식은 안전하게 갱신할 수 없습니다.');
    if (workbook && tag === 'definedName' && hasReference(formula)) return '';
    if (!hasReference(formula)) return element;
    return '<' + tag + attributes + '>' + xmlEscape(result.formula) + '</' + tag + '>';
  });
  workbookXml = rewriteXml(workbookXml, 'xl/workbook.xml', { workbook: true });
  outputFiles['xl/workbook.xml'] = new TextEncoder().encode(workbookXml);
  modifiedPaths.add('xl/workbook.xml');
  for (const item of model.sheets) {
    if (item.index === sheetIndex) continue;
    const rewritten = rewriteXml(new TextDecoder().decode(outputFiles[item.path]), item.path, { rejectNonCell: true });
    if (rewritten !== new TextDecoder().decode(outputFiles[item.path])) {
      outputFiles[item.path] = new TextEncoder().encode(rewritten);
      modifiedPaths.add(item.path);
    }
  }
  const sheetTags = [...workbookXml.matchAll(/<sheet\b[^>]*\/?\s*>/g)], selectedTag = sheetTags[sheetIndex]?.[0];
  if (sheetTags.length !== model.sheets.length || !selectedTag || parseTagAttributes(selectedTag).name !== sheet.name) fail('통합문서 시트 목록을 안전하게 갱신할 수 없습니다.');
  const relationshipId = parseTagAttributes(selectedTag)['r:id'];
  if (!relationshipId) fail('시트 관계 ID를 찾을 수 없습니다.');
  let updatedWorkbook = workbookXml.replace(selectedTag, '');
  updatedWorkbook = updatedWorkbook.replace(/<definedName\b[^>]*>([\s\S]*?)<\/definedName>/g, (element, body) => {
    const opening = element.slice(0, element.indexOf('>') + 1), attributes = parseTagAttributes(opening);
    if (attributes.localSheetId === undefined) return element;
    const localIndex = Number(attributes.localSheetId);
    if (localIndex === sheetIndex) return '';
    return localIndex > sheetIndex ? setTagAttribute(opening, 'localSheetId', localIndex - 1) + body + '</definedName>' : element;
  });
  updatedWorkbook = updatedWorkbook.replace(/<workbookView\b[^>]*\/?\s*>/g, tag => {
    const attributes = parseTagAttributes(tag), active = Number(attributes.activeTab);
    if (!Number.isInteger(active)) return tag;
    const next = active === sheetIndex ? Math.min(sheetIndex, model.sheets.length - 2) : active > sheetIndex ? active - 1 : active;
    return setTagAttribute(tag, 'activeTab', next);
  });
  outputFiles['xl/workbook.xml'] = new TextEncoder().encode(enableFullCalculationOnLoad(updatedWorkbook));
  modifiedPaths.add('xl/workbook.xml');
  const relsPath = 'xl/_rels/workbook.xml.rels', relationshipsXml = new TextDecoder().decode(outputFiles[relsPath]);
  const relation = [...relationshipsXml.matchAll(/<Relationship\b[^>]*\/?\s*>/g)].find(match => parseTagAttributes(match[0]).Id === relationshipId);
  if (!relation || parseTagAttributes(relation[0]).Target !== sheet.path.slice('xl/'.length)) fail('삭제할 시트의 관계 항목을 안전하게 찾을 수 없습니다.');
  outputFiles[relsPath] = new TextEncoder().encode(relationshipsXml.replace(relation[0], ''));
  modifiedPaths.add(relsPath);
  const contentPath = '[Content_Types].xml';
  let contentTypesXml = new TextDecoder().decode(outputFiles[contentPath]);
  for (const partPath of [sheet.path, ...(sheet.tables || []).map(table => table.path)]) {
    const contentTag = [...contentTypesXml.matchAll(/<Override\b[^>]*\/?\s*>/g)].find(match => parseTagAttributes(match[0]).PartName === '/' + partPath);
    if (contentTag) contentTypesXml = contentTypesXml.replace(contentTag[0], '');
  }
  if (contentTypesXml !== new TextDecoder().decode(outputFiles[contentPath])) {
    outputFiles[contentPath] = new TextEncoder().encode(contentTypesXml);
    modifiedPaths.add(contentPath);
  }
  delete outputFiles[sheet.path];
  modifiedPaths.add(sheet.path);
  if (sheet.hasTables) {
    delete outputFiles[sheet.tableRelationshipPath];
    modifiedPaths.add(sheet.tableRelationshipPath);
    for (const table of sheet.tables) { delete outputFiles[table.path]; modifiedPaths.add(table.path); }
  }
  const bytes = zipSync(outputFiles, { level: 0 }), checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('시트 삭제 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  for (const [path, before] of Object.entries(model.files)) {
    if (modifiedPaths.has(path)) continue;
    if (!checked.files[path] || Buffer.compare(Buffer.from(before), Buffer.from(checked.files[path])) !== 0) fail('시트 삭제 중 비수정 패키지 항목이 달라졌습니다: ' + path);
  }
  return { bytes, model: checked, sheetIndex: Math.min(sheetIndex, checked.sheets.length - 1) };
}

function patchMoveSheet(model, sheetIndex, targetIndex) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  if (!Number.isInteger(sheetIndex) || !model.sheets[sheetIndex] || !Number.isInteger(targetIndex) || !model.sheets[targetIndex]) fail('시트 이동 위치가 유효하지 않습니다.');
  if (sheetIndex === targetIndex) return { bytes: model.bytes, model, sheetIndex };
  const workbookPath = 'xl/workbook.xml', workbookXml = new TextDecoder().decode(model.files[workbookPath]);
  const tags = [...workbookXml.matchAll(/<sheet\b[^>]*\/?\s*>/g)];
  if (tags.length !== model.sheets.length || tags.some((tag, index) => parseTagAttributes(tag[0]).name !== model.sheets[index].name)) fail('통합문서 시트 목록을 안전하게 재정렬할 수 없습니다.');
  const order = model.sheets.map((_, index) => index), [moved] = order.splice(sheetIndex, 1);
  order.splice(targetIndex, 0, moved);
  const oldToNew = new Map(order.map((oldIndex, newIndex) => [oldIndex, newIndex]));
  let nextWorkbook = workbookXml;
  for (const { 0: tag } of tags) nextWorkbook = nextWorkbook.replace(tag, '\u0000');
  const reorderedTags = order.map(index => tags[index][0]);
  let cursor = 0;
  nextWorkbook = nextWorkbook.replace(/\u0000/g, () => reorderedTags[cursor++]);
  nextWorkbook = nextWorkbook.replace(/<definedName\b[^>]*>/g, opening => {
    const attributes = parseTagAttributes(opening);
    if (attributes.localSheetId === undefined) return opening;
    const current = Number(attributes.localSheetId), updated = oldToNew.get(current);
    if (!Number.isInteger(updated)) fail('시트 범위를 지정하는 정의된 이름을 안전하게 갱신할 수 없습니다.');
    return updated === current ? opening : setTagAttribute(opening, 'localSheetId', updated);
  });
  nextWorkbook = nextWorkbook.replace(/<workbookView\b[^>]*\/?\s*>/g, tag => {
    const active = Number(parseTagAttributes(tag).activeTab), updated = oldToNew.get(active);
    return Number.isInteger(updated) && updated !== active ? setTagAttribute(tag, 'activeTab', updated) : tag;
  });
  const outputFiles = Object.assign({}, model.files);
  outputFiles[workbookPath] = new TextEncoder().encode(enableFullCalculationOnLoad(nextWorkbook));
  const bytes = zipSync(outputFiles, { level: 0 }), checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('시트 재정렬 후 패키지 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  for (const [path, before] of Object.entries(model.files)) if (path !== workbookPath && (!checked.files[path] || Buffer.compare(Buffer.from(before), Buffer.from(checked.files[path])) !== 0)) fail('시트 재정렬 중 시트 순서 외의 패키지 항목이 달라졌습니다: ' + path);
  return { bytes, model: checked, sheetIndex: targetIndex };
}

function transformFormulaForStructure(formula, formulaSheet, targetSheet, operation, emptyTablePrototype = null) {
  let output = '', index = 0;
  const isName = character => Boolean(character && /[\p{L}\p{N}_.]/u.test(character));
  const readReference = start => {
    let cursor = start, sheetName = null, sheetPrefix = '';
    if (formula[cursor] === "'") {
      let name = '', end = cursor + 1;
      while (end < formula.length) {
        if (formula.startsWith("''", end)) { name += "'"; end += 2; continue; }
        if (formula[end] === "'") break;
        name += formula[end++];
      }
      if (formula[end] === "'" && formula[end + 1] === '!') { sheetName = name; cursor = end + 2; sheetPrefix = formula.slice(start, cursor); }
      else return null;
    } else {
      const prefix = /^([\p{L}_][\p{L}\p{N}_.]*)!/u.exec(formula.slice(cursor));
      if (prefix) { sheetName = prefix[1]; cursor += prefix[0].length; sheetPrefix = prefix[0]; }
    }
    const cell = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)/.exec(formula.slice(cursor));
    if (!cell) return null;
    const end = cursor + cell[0].length, previous = formula[start - 1] || '', next = formula[end] || '';
    if ((!sheetPrefix && isName(previous)) || isName(next) || next === '(') return null;
    const column = columnNumber(cell[2].toUpperCase() + '1'), row = Number(cell[4]);
    if (!column || column > LIMITS.columns || row > LIMITS.rows) return null;
    return { start, end, sheetName, sheetPrefix, column, row, absColumn: cell[1], columnLetters: cell[2], absRow: cell[3], rowText: cell[4] };
  };
  const formatReference = (reference, row, column) => reference.sheetPrefix + reference.absColumn + columnLetters(column) + reference.absRow + row;
  const movePoint = (row, column) => {
    const coordinate = operation.index + 1;
    let value = operation.axis === 'row' ? row : column;
    if (operation.action === 'insert' && value >= coordinate) value++;
    else if (operation.action === 'delete') {
      if (value === coordinate) return { deleted: true, row, column };
      if (value > coordinate) value--;
    }
    return operation.axis === 'row' ? { row: value, column } : { row, column: value };
  };
  while (index < formula.length) {
    if (formula[index] === '"') {
      const start = index++;
      while (index < formula.length) {
        if (formula.startsWith('""', index)) { index += 2; continue; }
        if (formula[index++] === '"') break;
      }
      output += formula.slice(start, index);
      continue;
    }
    if (formula[index] === '[') {
      const start = index++;
      let depth = 1;
      while (index < formula.length && depth) { if (formula[index] === "'" && /['\[\]#@]/.test(formula[index+1]||'')) { index+=2; continue; } if (formula[index] === '[') depth++; else if (formula[index] === ']') depth--; index++; }
      output += formula.slice(start, index);
      continue;
    }
    if (operation.action === 'move' && /^(?:[\p{L}_][\p{L}\p{N}_.]*):(?:[\p{L}_][\p{L}\p{N}_.]*)!/u.test(formula.slice(index))) fail('3차원 시트 참조가 있어 합계 행 참조를 안전하게 이동할 수 없습니다.');
    const first = readReference(index);
    if (!first) { output += formula[index++]; continue; }
    const separator = /^\s*:\s*/.exec(formula.slice(first.end));
    const range = separator ? readReference(first.end + separator[0].length) : null;
    const firstSheet = first.sheetName || formulaSheet;
    const secondSheet = range ? range.sheetName || first.sheetName || formulaSheet : null;
    if (emptyTablePrototype && operation.axis === 'row' && operation.action === 'delete') {
      const futureDataCell = (reference, owner) => !reference.absRow && reference.row === operation.index + 1 && reference.column - 1 >= emptyTablePrototype.left && reference.column - 1 <= emptyTablePrototype.right && owner.toLocaleLowerCase() === targetSheet.toLocaleLowerCase();
      // Keep relative references to this table's future first data row in the
      // prototype only. Actual worksheet formulas still follow normal deletion.
      if (futureDataCell(first, firstSheet) && (!range || futureDataCell(range, secondSheet))) {
        output += formula.slice(first.start, range?.end ?? first.end);
        index = range?.end ?? first.end;
        continue;
      }
    }
    if (operation.action === 'move') {
      if (first.sheetName?.includes(':') || range?.sheetName?.includes(':')) fail('3차원 시트 참조가 있어 합계 행 참조를 안전하게 이동할 수 없습니다.');
      const inside = reference => reference.row - 1 >= operation.source.top && reference.row - 1 <= operation.source.bottom && reference.column - 1 >= operation.source.left && reference.column - 1 <= operation.source.right;
      const targets = name => name.toLocaleLowerCase() === targetSheet.toLocaleLowerCase();
      const moved = targets(firstSheet) && inside(first) && (!range || targets(secondSheet) && inside(range));
      output += moved ? formatReference(first, first.row + operation.rowDelta, first.column + operation.columnDelta) + (range ? ':' + formatReference(range, range.row + operation.rowDelta, range.column + operation.columnDelta) : '') : formula.slice(first.start, range?.end ?? first.end);
      index = range?.end ?? first.end;
      continue;
    }
    if (range && firstSheet.toLocaleLowerCase() === targetSheet.toLocaleLowerCase() && secondSheet.toLocaleLowerCase() === targetSheet.toLocaleLowerCase()) {
      let firstRow = first.row, firstColumn = first.column, secondRow = range.row, secondColumn = range.column;
      const lower = operation.axis === 'row' ? Math.min(firstRow, secondRow) : Math.min(firstColumn, secondColumn);
      const upper = operation.axis === 'row' ? Math.max(firstRow, secondRow) : Math.max(firstColumn, secondColumn);
      const coordinate = operation.index + 1;
      if (operation.action === 'insert') {
        if (coordinate <= lower) { firstRow += operation.axis === 'row' ? 1 : 0; secondRow += operation.axis === 'row' ? 1 : 0; firstColumn += operation.axis === 'column' ? 1 : 0; secondColumn += operation.axis === 'column' ? 1 : 0; }
        else if (coordinate <= upper) {
          if (operation.axis === 'row') {
            if (firstRow <= secondRow) secondRow++; else firstRow++;
          } else if (firstColumn <= secondColumn) secondColumn++; else firstColumn++;
        }
      } else if (coordinate < lower) {
        firstRow -= operation.axis === 'row' ? 1 : 0; secondRow -= operation.axis === 'row' ? 1 : 0;
        firstColumn -= operation.axis === 'column' ? 1 : 0; secondColumn -= operation.axis === 'column' ? 1 : 0;
      } else if (coordinate <= upper) {
        if (operation.axis === 'row') {
          if (lower === upper) { output += '#REF!'; index = range.end; continue; }
          if (firstRow <= secondRow) secondRow--; else firstRow--;
        } else {
          if (lower === upper) { output += '#REF!'; index = range.end; continue; }
          if (firstColumn <= secondColumn) secondColumn--; else firstColumn--;
        }
      }
      output += formatReference(first, firstRow, firstColumn) + ':' + formatReference(range, secondRow, secondColumn);
      index = range.end;
      continue;
    }
    if (firstSheet.toLocaleLowerCase() === targetSheet.toLocaleLowerCase()) {
      const moved = movePoint(first.row, first.column);
      output += moved.deleted ? '#REF!' : formatReference(first, moved.row, moved.column);
    } else output += formula.slice(first.start, first.end);
    index = first.end;
  }
  return output;
}

function transformDefinedNamesForStructure(xml, model, targetSheet, operation) {
  const definitions = new Map((model.definedNames || []).map(definition => [
    definition.name.toLocaleLowerCase() + ':' + (definition.localSheetId === null ? '*' : definition.localSheetId),
    definition
  ]));
  let changed = false;
  const updatedXml = xml.replace(/<definedName\b[^>]*>[\s\S]*?<\/definedName>/g, element => {
    const openingEnd = element.indexOf('>') + 1, opening = element.slice(0, openingEnd), attributes = parseTagAttributes(opening);
    const localSheetId = attributes.localSheetId === undefined ? null : Number(attributes.localSheetId);
    const definition = definitions.get(attributes.name?.toLocaleLowerCase() + ':' + (localSheetId === null ? '*' : localSheetId));
    if (!definition || definition.printAreas || definition.printTitles) return element;
    const formula = parseXml(new TextEncoder().encode(element), 'xl/workbook.xml').root.text.trim();
    const formulaSheet = localSheetId === null ? definition.source?.sheetName || targetSheet : model.sheets[localSheetId]?.name;
    if (!formulaSheet) fail('이름 정의의 기준 시트를 찾을 수 없어 구조 변경을 중단했습니다.');
    const updated = transformFormulaForStructure(formula, formulaSheet, targetSheet, operation);
    if (updated === formula) return element;
    changed = true;
    return opening + xmlEscape(updated) + '</definedName>';
  });
  return { xml: updatedXml, changed };
}

function transformDataValidationsForStructure(sheet, targetSheet, operation) {
  const container = findElementRanges(sheet.xml, 'dataValidations').get('');
  if (!container || container.selfClosing) return [];
  const entries = [...findElementRanges(sheet.xml, 'dataValidation', undefined, true).values()]
    .filter(range => range.start >= container.start && range.end <= container.end);
  if (!entries.length) return [];
  const patches = [];
  let removed = 0;
  for (const range of entries) {
    const element = sheet.xml.slice(range.start, range.end), openingEnd = element.indexOf('>') + 1;
    let opening = element.slice(0, openingEnd);
    const attributes = parseTagAttributes(opening);
    if (sheet.name.toLocaleLowerCase() === targetSheet.toLocaleLowerCase()) {
      const updatedSqref = transformFormulaForStructure(attributes.sqref, sheet.name, targetSheet, operation)
        .split(/\s+/).filter(reference => reference && reference.toUpperCase() !== '#REF!').join(' ');
      if (!updatedSqref) {
        patches.push({ start: range.start, end: range.end, replacement: '' });
        removed++;
        continue;
      }
      if (updatedSqref !== attributes.sqref) opening = replaceTagAttribute(opening, 'sqref', updatedSqref);
    }
    if (opening !== element.slice(0, openingEnd)) patches.push({ start: range.start, end: range.start + openingEnd, replacement: opening });
    for (const formulaName of ['formula1', 'formula2']) {
      const formulaRange = findElementRanges(element, formulaName).get('');
      if (!formulaRange || formulaRange.selfClosing) continue;
      const formulaElement = element.slice(formulaRange.start, formulaRange.end);
      const formula = parseXml(new TextEncoder().encode(formulaElement), sheet.path).root.text.trim();
      const updated = transformFormulaForStructure(formula, sheet.name, targetSheet, operation);
      if (updated !== formula) patches.push({
        start: range.start + formulaRange.startTagEnd,
        end: range.start + formulaRange.closingStart,
        replacement: xmlEscape(updated)
      });
    }
  }
  if (removed === entries.length) return [{ start: container.start, end: container.end, replacement: '' }];
  if (removed) {
    const opening = sheet.xml.slice(container.start, container.startTagEnd), attributes = parseTagAttributes(opening);
    if (attributes.count !== undefined) patches.push({
      start: container.start,
      end: container.startTagEnd,
      replacement: replaceTagAttribute(opening, 'count', String(entries.length - removed))
    });
  }
  return patches;
}

function transformChartReferencesForStructure(model, targetSheet, operation, outputFiles) {
  const changed = new Set();
  for (const chart of model.chartPaths || []) {
    const original = new TextDecoder().decode(outputFiles[chart.path] || model.files[chart.path]);
    const root = parseXml(new TextEncoder().encode(original), chart.path).root;
    const prefix = root.name.includes(':') ? root.name.slice(0, root.name.indexOf(':') + 1) : '';
    const ranges = [...findElementRanges(original, prefix + 'f', undefined, true).values()], patches = [];
    const chartSheet = model.sheets[chart.sheetIndex];
    if (!chartSheet) fail('차트가 있는 워크시트를 찾을 수 없습니다.');
    for (const range of ranges) {
      if (range.selfClosing) fail('차트 원본 참조를 안전하게 수정할 수 없습니다.');
      const formula = parseXml(new TextEncoder().encode(original.slice(range.start, range.end)), chart.path).root.text.trim();
      if (!formula) fail('차트 원본 참조가 비어 있습니다.');
      const updated = transformFormulaForStructure(formula, chartSheet.name, targetSheet, operation);
      if (updated === formula) continue;
      if (/(?:^|[^\w])#REF!(?:$|[^\w])/.test(updated)) fail('차트 원본 셀을 모두 삭제하면 차트를 보존할 수 없습니다.');
      patches.push({ start: range.start, end: range.end, replacement: original.slice(range.start, range.startTagEnd) + xmlEscape(updated) + original.slice(range.closingStart, range.end) });
    }
    if (!patches.length) continue;
    let xml = original;
    for (const patch of patches.sort((left, right) => right.start - left.start)) xml = xml.slice(0, patch.start) + patch.replacement + xml.slice(patch.end);
    outputFiles[chart.path] = new TextEncoder().encode(xml);
    changed.add(chart.path);
  }
  return changed;
}

function transformDrawingMarker(marker, operation) {
  const coordinate = operation.axis === 'row' ? marker.row : marker.column;
  let next = coordinate;
  if (operation.action === 'insert' && coordinate >= operation.index) next++;
  if (operation.action === 'delete' && coordinate > operation.index) next--;
  if (next < 0 || next >= (operation.axis === 'row' ? LIMITS.rows : LIMITS.columns)) fail('차트 위치가 행·열 한도를 벗어납니다.');
  return operation.axis === 'row' ? { ...marker, row: next } : { ...marker, column: next };
}

function transformChartDrawingForStructure(model, sheet, operation, outputFiles, modifiedPaths) {
  const drawingPath = sheet.chartDrawingPath;
  if (!drawingPath || !sheet.charts?.length) return;
  const original = new TextDecoder().decode(outputFiles[drawingPath] || model.files[drawingPath]);
  const parsed = parseXml(new TextEncoder().encode(original), drawingPath), root = parsed.root;
  if (localName(root) !== 'wsDr') fail('차트 그리기 XML을 안전하게 수정할 수 없습니다.');
  const beforeAxes = drawingAxes({ ...sheet, hiddenRows: sheet.nativeHiddenRows || sheet.hiddenRows });
  const sheetXml = new TextDecoder().decode(outputFiles[sheet.path]);
  const afterAxes = drawingAxes({ ...sheet, ...dimensionView(sheetXml, sheet.path) });
  const anchors = root.children.filter(node => ['twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor'].includes(localName(node)));
  const operations = [];
  for (const anchor of anchors) {
    const geometry = readChartAnchor(anchor, true);
    if (!geometry) fail('차트 위치를 해석할 수 없어 행·열을 변경할 수 없습니다.');
    const siblings = anchors.filter(node => node.name === anchor.name), range = [...findElementRanges(original, anchor.name, undefined, true).values()][siblings.indexOf(anchor)];
    if (!range) fail('차트 위치 XML을 찾을 수 없습니다.');
    let xml = original.slice(range.start, range.end);
    if (geometry.type === 'twoCellAnchor' && geometry.editAs === 'absolute') {
      const from = drawingPoint(geometry.from, beforeAxes), to = drawingPoint(geometry.to, beforeAxes);
      xml = replaceDrawingMarker(xml, childrenByLocalName(anchor, 'from')[0], drawingMarker(from, afterAxes));
      xml = replaceDrawingMarker(xml, childrenByLocalName(anchor, 'to')[0], drawingMarker(to, afterAxes));
      const nextFrom = from, nextTo = to;
      const transform = firstDescendant(anchor, ['xfrm']);
      for (const [name, values] of [['off', { x: nextFrom.x, y: nextFrom.y }], ['ext', { cx: nextTo.x - nextFrom.x, cy: nextTo.y - nextFrom.y }]]) {
        const node = transform && childrenByLocalName(transform, name)[0], part = node && findElementRanges(xml, node.name).get('');
        if (!part) continue;
        let tag = xml.slice(part.start, part.startTagEnd);
        for (const [key, value] of Object.entries(values)) tag = setTagAttribute(tag, key, Math.round(value * 9525));
        xml = xml.slice(0, part.start) + tag + xml.slice(part.startTagEnd);
      }
    } else if (geometry.type === 'twoCellAnchor' && geometry.editAs === 'oneCell') {
      const from = transformDrawingMarker(geometry.from, operation), point = drawingPoint(from, afterAxes), oldFrom = drawingPoint(geometry.from, beforeAxes), oldTo = drawingPoint(geometry.to, beforeAxes);
      const width = oldTo.x - oldFrom.x, height = oldTo.y - oldFrom.y;
      if (!(width > 0 && height > 0)) fail('차트 크기가 유효하지 않아 행·열을 변경할 수 없습니다.');
      const to = drawingMarker({ x: point.x + width, y: point.y + height }, afterAxes);
      xml = replaceDrawingMarker(xml, childrenByLocalName(anchor, 'from')[0], from);
      xml = replaceDrawingMarker(xml, childrenByLocalName(anchor, 'to')[0], to);
      const transform = firstDescendant(anchor, ['xfrm']);
      for (const [name, values] of [['off', { x: point.x, y: point.y }], ['ext', { cx: width, cy: height }]]) {
        const node = transform && childrenByLocalName(transform, name)[0], part = node && findElementRanges(xml, node.name).get('');
        if (!part) continue;
        let tag = xml.slice(part.start, part.startTagEnd);
        for (const [key, value] of Object.entries(values)) tag = setTagAttribute(tag, key, Math.round(value * 9525));
        xml = xml.slice(0, part.start) + tag + xml.slice(part.startTagEnd);
      }
    } else if (geometry.type === 'twoCellAnchor') {
      const from = transformDrawingMarker(geometry.from, operation), to = transformDrawingMarker(geometry.to, operation);
      xml = replaceDrawingMarker(xml, childrenByLocalName(anchor, 'from')[0], from);
      xml = replaceDrawingMarker(xml, childrenByLocalName(anchor, 'to')[0], to);
      const fromPoint = drawingPoint(from, afterAxes), toPoint = drawingPoint(to, afterAxes);
      if (!(toPoint.x > fromPoint.x && toPoint.y > fromPoint.y)) fail('행·열 삭제로 차트의 위치나 크기가 사라집니다.');
    } else if (geometry.type === 'oneCellAnchor') {
      const from = transformDrawingMarker(geometry.from, operation);
      xml = replaceDrawingMarker(xml, childrenByLocalName(anchor, 'from')[0], from);
    }
    if (xml !== original.slice(range.start, range.end)) operations.push({ start: range.start, end: range.end, replacement: xml });
  }
  if (!operations.length) return;
  let updated = original;
  for (const item of operations.sort((left, right) => right.start - left.start)) updated = updated.slice(0, item.start) + item.replacement + updated.slice(item.end);
  writeWorksheetDrawing(sheet, drawingPath, outputFiles, modifiedPaths, updated);
}

function transformMergedCellsForStructure(sheet, operation) {
  const ranges = sheet.mergedCells || [];
  if (!ranges.length) return [];
  const references = findElementRanges(sheet.xml, 'mergeCell', 'ref'), patches = [];
  for (const range of ranges) {
    const start = operation.axis === 'row' ? range.top : range.left;
    const end = operation.axis === 'row' ? range.bottom : range.right;
    const { index, action } = operation;
    if (action === 'insert' && index > start && index <= end || action === 'delete' && index >= start && index <= end) {
      fail('병합 범위를 가로지르는 행·열 삽입·삭제는 지원하지 않습니다.');
    }
    const shift = action === 'insert' ? index <= start : index < start;
    const delta = shift ? (action === 'insert' ? 1 : -1) : 0;
    if (!delta) continue;
    const moved = operation.axis === 'row'
      ? { ...range, top: range.top + delta, bottom: range.bottom + delta }
      : { ...range, left: range.left + delta, right: range.right + delta };
    if (moved.top < 0 || moved.left < 0 || moved.bottom >= LIMITS.rows || moved.right >= LIMITS.columns) fail('병합 셀 범위가 행·열 한도를 벗어납니다.');
    const original = references.get(range.ref);
    if (!original) fail('병합 셀 참조를 워크시트에서 찾을 수 없습니다.');
    const originalReference = /^(\$?)[A-Z]{1,3}(\$?)[1-9]\d*:(\$?)[A-Z]{1,3}(\$?)[1-9]\d*$/i.exec(range.ref);
    if (!originalReference) fail('병합 셀 참조를 안전하게 이동할 수 없습니다.');
    const reference = originalReference[1] + columnLetters(moved.left + 1) + originalReference[2] + (moved.top + 1) + ':'
      + originalReference[3] + columnLetters(moved.right + 1) + originalReference[4] + (moved.bottom + 1);
    const opening = sheet.xml.slice(original.start, original.startTagEnd);
    patches.push({ start: original.start, end: original.startTagEnd, replacement: setTagAttribute(opening, 'ref', reference) });
  }
  return patches;
}

function transformConditionalFormatsForStructure(sheet, operation) {
  const containers = [...findElementRanges(sheet.xml, 'conditionalFormatting', undefined, true).values()];
  if (!containers.length) return [];
  const patches = [], rules = sheet.conditionalFormats || [];
  const transformRange = source => {
    const start = operation.axis === 'row' ? source.top : source.left;
    const end = operation.axis === 'row' ? source.bottom : source.right;
    let nextStart = start, nextEnd = end;
    if (operation.action === 'insert') {
      if (operation.index <= start) { nextStart++; nextEnd++; }
      else if (operation.index <= end) nextEnd++;
    } else if (operation.index < start) { nextStart--; nextEnd--; }
    else if (operation.index <= end) nextEnd--;
    if (nextStart < 0 || nextEnd < nextStart) return null;
    const moved = operation.axis === 'row' ? { ...source, top: nextStart, bottom: nextEnd } : { ...source, left: nextStart, right: nextEnd };
    if (moved.top < 0 || moved.left < 0 || moved.bottom >= LIMITS.rows || moved.right >= LIMITS.columns) fail('조건부 서식 범위가 행·열 제한을 벗어납니다.');
    return moved;
  };
  const formatReference = (range, original) => {
    const parts = original.replaceAll('$', '').split(':');
    const formatCell = (row, column, reference) => {
      const match = /^(\$?)([A-Z]{1,3})(\$?)[1-9]\d*$/i.exec(reference);
      if (!match) fail('조건부 서식 범위 참조를 안전하게 이동할 수 없습니다.');
      return match[1] + columnLetters(column + 1) + match[3] + (row + 1);
    };
    const first = formatCell(range.top, range.left, parts[0]);
    if (parts.length === 1) return first;
    return first + ':' + formatCell(range.bottom, range.right, parts[1]);
  };
  for (const container of containers) {
    const opening = sheet.xml.slice(container.start, container.startTagEnd), attributes = parseTagAttributes(opening);
    const references = String(attributes.sqref || '').trim().split(/\s+/).filter(Boolean);
    if (!references.length) fail('조건부 서식 범위가 비어 있습니다.');
    const transformed = references.map(reference => {
      const parts = reference.replaceAll('$', '').split(':');
      if (parts.length > 2) fail('조건부 서식 범위 참조가 유효하지 않습니다.');
      const source = { top: rowNumber(parts[0]) - 1, bottom: rowNumber(parts.at(-1)) - 1, left: columnNumber(parts[0]) - 1, right: columnNumber(parts.at(-1)) - 1 };
      if (![source.top, source.bottom, source.left, source.right].every(Number.isInteger) || source.top < 0 || source.left < 0) fail('조건부 서식 범위 참조가 유효하지 않습니다.');
      const next = transformRange(source);
      return next ? formatReference(next, reference) : null;
    }).filter(Boolean);
    if (!transformed.length) {
      patches.push({ start: container.start, end: container.end, replacement: '' });
      continue;
    }
    const nextSqref = transformed.join(' ');
    if (nextSqref !== attributes.sqref) patches.push({ start: container.start, end: container.startTagEnd, replacement: setTagAttribute(opening, 'sqref', nextSqref) });
    const ruleRanges = [...findElementRanges(sheet.xml, 'cfRule', undefined, true).values()].filter(range => range.start >= container.start && range.end <= container.end);
    for (const ruleRange of ruleRanges) {
      const ruleOpening = sheet.xml.slice(ruleRange.start, ruleRange.startTagEnd), ruleAttributes = parseTagAttributes(ruleOpening);
      const rule = rules.find(item => String(item.priority) === ruleAttributes.priority);
      if (rule?.type !== 'expression') continue;
      const formulaRanges = [...findElementRanges(sheet.xml, 'formula', undefined, true).values()].filter(range => range.start >= ruleRange.start && range.end <= ruleRange.end);
      if (formulaRanges.length !== 1) fail('조건부 서식 수식을 안전하게 이동할 수 없습니다.');
      const formulaRange = formulaRanges[0], formulaElement = sheet.xml.slice(formulaRange.start, formulaRange.end);
      const formula = parseXml(new TextEncoder().encode(formulaElement), sheet.path).root.text.trim();
      const updated = transformFormulaForStructure(formula, sheet.name, sheet.name, operation);
      if (updated !== formula) patches.push({ start: formulaRange.startTagEnd, end: formulaRange.closingStart, replacement: xmlEscape(updated) });
    }
  }
  return patches;
}

function transformColumnDefinitionsForStructure(sheet, operation) {
  if (operation.axis !== 'column') return [];
  const definitions = [...findElementRanges(sheet.xml, 'col', undefined, true).values()];
  if (!definitions.length) return [];
  const patches = [], coordinate = operation.index + 1;
  for (const range of definitions) {
    const opening = sheet.xml.slice(range.start, range.startTagEnd), attributes = parseTagAttributes(opening);
    let min = Number(attributes.min), max = Number(attributes.max);
    if (!Number.isInteger(min) || !Number.isInteger(max) || min < 1 || max < min || max > LIMITS.columns) fail('열 너비·서식 정의 범위를 안전하게 이동할 수 없습니다.');
    if (operation.action === 'insert') {
      if (coordinate <= min) { min++; max++; }
      else if (coordinate <= max) max++;
    } else if (coordinate < min) { min--; max--; }
    else if (coordinate <= max) {
      if (min === max) { patches.push({ start: range.start, end: range.end, replacement: '' }); continue; }
      max--;
    }
    if (max > LIMITS.columns) fail('열 삽입 후 열 너비·서식 정의가 워크시트 제한을 벗어납니다.');
    if (min !== Number(attributes.min) || max !== Number(attributes.max)) {
      let updated = setTagAttribute(opening, 'min', min);
      updated = setTagAttribute(updated, 'max', max);
      patches.push({ start: range.start, end: range.startTagEnd, replacement: updated });
    }
  }
  if (patches.length && definitions.every(range => patches.some(patch => patch.start === range.start && patch.end === range.end && patch.replacement === ''))) {
    const container = findElementRanges(sheet.xml, 'cols').get('');
    if (container && !container.selfClosing) {
      return [{ start: container.start, end: container.end, replacement: '' }];
    }
  }
  return patches;
}

function transformHyperlinksForStructure(model, targetSheet, operation) {
  const patchesBySheet = new Map(), removedRelationshipIdsBySheet = new Map();
  for (const sheet of model.sheets) {
    const container = findElementRanges(sheet.xml, 'hyperlinks').get('');
    if (!container) continue;
    const elements = [...findElementRanges(sheet.xml, 'hyperlink', undefined, true).values()]
      .filter(range => range.start >= container.start && range.end <= container.end);
    if (elements.length !== (sheet.hyperlinks || []).length) fail('워크시트 하이퍼링크 범위를 안전하게 이동할 수 없습니다.');
    const linkPatches = [], removedIds = new Set();
    for (const [index, element] of elements.entries()) {
      const link = sheet.hyperlinks[index], tag = sheet.xml.slice(element.start, element.startTagEnd), attributes = parseTagAttributes(tag);
      let opening = tag, remove = false;
      if (sheet.name.toLocaleLowerCase() === targetSheet.toLocaleLowerCase()) {
        const movedReference = transformFormulaForStructure(attributes.ref, sheet.name, targetSheet, operation);
        if (movedReference.includes('#REF!')) remove = true;
        else if (movedReference !== attributes.ref) opening = setTagAttribute(opening, 'ref', movedReference);
      }
      if (!remove && link.location) {
        const hasHash = link.location.startsWith('#'), location = hasHash ? link.location.slice(1) : link.location;
        const target = parseRangeReference(location, sheet.name, LIMITS.rows * LIMITS.columns);
        if (target?.sheetName?.toLocaleLowerCase() === targetSheet.toLocaleLowerCase()) {
          const movedLocation = transformFormulaForStructure(location, sheet.name, targetSheet, operation);
          if (movedLocation.includes('#REF!')) fail('삭제할 셀을 가리키는 내부 하이퍼링크를 먼저 수정하거나 제거하세요.');
          if (movedLocation !== location) opening = setTagAttribute(opening, 'location', (hasHash ? '#' : '') + movedLocation);
        }
      }
      if (remove) {
        if (link.relationshipId) removedIds.add(link.relationshipId);
        linkPatches.push({ start: element.start, end: element.end, replacement: '' });
      } else if (opening !== tag) linkPatches.push({ start: element.start, end: element.startTagEnd, replacement: opening });
    }
    if (!linkPatches.length) continue;
    const patches = linkPatches.filter(patch => patch.replacement === '').length === elements.length
      ? [{ start: container.start, end: container.end, replacement: '' }]
      : linkPatches;
    patchesBySheet.set(sheet.path, patches);
    if (removedIds.size) removedRelationshipIdsBySheet.set(sheet.path, removedIds);
  }
  return { patchesBySheet, removedRelationshipIdsBySheet };
}

function transformStructure(model, operation) {
  if (!model.writable) fail(model.readOnlyReason || '이 통합문서는 읽기 전용입니다.');
  model = preserveImportedFilterVisibility(model);
  const { sheetIndex, axis, action, index } = operation || {};
  let sheet = model.sheets[sheetIndex];
  if (!sheet || !['row', 'column'].includes(axis) || !['insert', 'delete'].includes(action) || !Number.isInteger(index) || index < 0) fail('XLSX 행·열 구조 변경 요청이 유효하지 않습니다.');
  const coordinate = index + 1;
  const maxCoordinate = axis === 'row' ? LIMITS.rows : LIMITS.columns;
  if (coordinate > maxCoordinate) fail('XLSX 행·열 제한을 벗어났습니다.');
  if (sheet.pageBreaksUnsupported) fail('지원하지 않는 페이지 나누기는 모두 지운 뒤 새로 설정하세요.');
  const tableRangeChanges = [], tableHeaderAdditions = [];
  if (sheet.hasTables) for (const table of sheet.tables) {
    const range = table.range, header = table.headerRowCount || 0, totals = table.totalsRowCount || 0;
    if (axis === 'row') {
      const firstDataRow = range.top + header, lastDataRow = range.bottom - totals;
      if (action === 'insert') {
        if (index <= range.top) tableRangeChanges.push({ table, range: { ...range, top: range.top + 1, bottom: range.bottom + 1 } });
        else if (index <= lastDataRow + (totals ? 1 : 0)) tableRangeChanges.push({ table, range: { ...range, bottom: range.bottom + 1 } });
      } else if (index < range.top) tableRangeChanges.push({ table, range: { ...range, top: range.top - 1, bottom: range.bottom - 1 } });
      else if (index <= range.bottom) {
        if (index < firstDataRow || index > lastDataRow) fail('Excel 표 머리글·합계 행 삭제는 지원하지 않습니다.');
        if (range.bottom > range.top + header + totals || header === 1) tableRangeChanges.push({ table, range: { ...range, bottom: range.bottom - 1 } });
        else fail('마지막 Excel 표 데이터 행은 삭제할 수 없습니다.');
      }
    } else if (action === 'insert') {
      if (index <= range.left) tableRangeChanges.push({ table, range: { ...range, left: range.left + 1, right: range.right + 1 } });
      else if (index <= range.right) {
        if (header !== 1) fail('머리글이 표시되지 않는 Excel 표에는 열을 삽입할 수 없습니다.');
        const columnIndex = index - range.left, names = table.columns.map(name => name.toLocaleLowerCase());
        let name = 'Column' + (columnIndex + 1), suffix = 2;
        while (names.includes(name.toLocaleLowerCase())) name = 'Column' + (columnIndex + 1) + '_' + suffix++;
        tableRangeChanges.push({ table, range: { ...range, right: range.right + 1 }, columnChange: { action, index: columnIndex, name } });
        tableHeaderAdditions.push({ row: range.top, column: index, name });
      }
    } else if (index < range.left) tableRangeChanges.push({ table, range: { ...range, left: range.left - 1, right: range.right - 1 } });
    else if (index <= range.right) {
      if (range.right === range.left) fail('Excel 표의 마지막 열은 삭제할 수 없습니다.');
      const columnIndex = index - range.left, removedName = table.columns[columnIndex].toLocaleLowerCase(), tableXml = new TextDecoder().decode(model.files[table.path]);
      if (model.sheets.some((item, ownerSheetIndex) => [...item.cells.values()].some(cell => {
        if (!cell.formula) return false;
        const formula = String(cell.formula).toLocaleLowerCase(), inTable = ownerSheetIndex === sheetIndex && cell.row >= range.top && cell.row <= range.bottom && cell.column >= range.left && cell.column <= range.right;
        return formula.includes(table.name.toLocaleLowerCase() + '[') && formula.includes('[' + removedName + ']') || inTable && cell.column !== index && formula.includes('[@' + removedName + ']');
      }))) fail('삭제할 표 열을 참조하는 수식이 있어 워크시트 열을 삭제할 수 없습니다.');
      const otherColumnFormulas = tableColumnMatches(tableXml).filter(match => {
        const tag = match[0], columnName = parseTagAttributes(tag).name;
        const formula = tag.toLocaleLowerCase();
        return columnName?.toLocaleLowerCase() !== removedName && /<(?:calculatedColumnFormula|totalsRowFormula)\b/.test(tag) && (formula.includes('[' + removedName + ']') || formula.includes('[@' + removedName + ']'));
      });
      if (otherColumnFormulas.length) fail('삭제할 표 열을 참조하는 표 수식이 있어 워크시트 열을 삭제할 수 없습니다.');
      tableRangeChanges.push({ table, range: { ...range, right: range.right - 1 }, columnChange: { action, index: columnIndex } });
    }
  }
  if (!sheet.sheetDataStructureValid) fail('워크시트 행 구조를 안전하게 이동할 수 없습니다.');
  if (!sheet.rowStructureValid) fail('행 번호 또는 행 내부 기능을 안전하게 이동할 수 없습니다.');
  if (model.sheets.some(item => (item.sharedFormulaGroups || []).length)) {
    model = flattenSharedFormulaGroups(model);
    sheet = model.sheets[sheetIndex];
  }
  if (model.sheets.some(item => [...item.cells.values()].some(cell => cell.formulaType && cell.formulaType !== 'normal'))) fail('배열 수식이 포함된 통합문서는 행·열 구조를 변경할 수 없습니다.');
  const definedNameTransform = transformDefinedNamesForStructure(new TextDecoder().decode(model.files['xl/workbook.xml']), model, sheet.name, operation);
  const rowRanges = findElementRanges(sheet.xml, 'row', 'r');
  const rowsByPosition = [...rowRanges.entries()].sort((a, b) => a[1].start - b[1].start);
  for (let i = 1; i < rowsByPosition.length; i++) {
    if (Number(rowsByPosition[i][0]) <= Number(rowsByPosition[i - 1][0])) fail('워크시트 행 순서가 유효하지 않아 구조를 변경할 수 없습니다.');
  }
  for (const row of new Set([...sheet.cells.values()].map(cell => cell.row))) {
    const columns = [...sheet.cells.values()].filter(cell => cell.row === row).sort((a, b) => sheet.ranges.get(a.address).start - sheet.ranges.get(b.address).start).map(cell => cell.column);
    for (let i = 1; i < columns.length; i++) if (columns[i] <= columns[i - 1]) fail('워크시트 셀 순서가 유효하지 않아 구조를 변경할 수 없습니다.');
  }
  const patchesByPath = new Map([[sheet.path, []]]);
  const patchesFor = formulaSheet => {
    if (!patchesByPath.has(formulaSheet.path)) patchesByPath.set(formulaSheet.path, []);
    return patchesByPath.get(formulaSheet.path);
  };
  const patches = patchesByPath.get(sheet.path);
  const columnDefinitions = transformColumnDefinitionsForStructure(sheet, operation);
  if (columnDefinitions.length) {
    const container = findElementRanges(sheet.xml, 'cols').get('');
    if (!container || container.selfClosing) fail('열 너비·서식 정의 컨테이너를 안전하게 수정할 수 없습니다.');
    let contents = sheet.xml.slice(container.start, container.end);
    for (const patch of columnDefinitions.sort((left, right) => right.start - left.start)) {
      const start = patch.start - container.start, end = patch.end - container.start;
      contents = contents.slice(0, start) + patch.replacement + contents.slice(end);
    }
    patches.push({ start: container.start, end: container.end, replacement: contents });
  }
  patches.push(...transformMergedCellsForStructure(sheet, operation));
  patches.push(...transformConditionalFormatsForStructure(sheet, operation));
  const hyperlinkRelationshipRemovals = new Map(), hyperlinkRelationshipPaths = new Set();
  for (const validationSheet of model.sheets) {
    const validationPatches = transformDataValidationsForStructure(validationSheet, sheet.name, operation);
    if (validationPatches.length) patchesFor(validationSheet).push(...validationPatches);
    const hyperlinkChanges = transformHyperlinksForStructure({ sheets: [validationSheet] }, sheet.name, operation);
    const hyperlinkPatches = hyperlinkChanges.patchesBySheet.get(validationSheet.path) || [];
    if (hyperlinkPatches.length) patchesFor(validationSheet).push(...hyperlinkPatches);
    const removedIds = hyperlinkChanges.removedRelationshipIdsBySheet.get(validationSheet.path);
    if (removedIds) hyperlinkRelationshipRemovals.set(validationSheet.path, removedIds);
  }
  for (const formulaSheet of model.sheets) for (const cell of formulaSheet.cells.values()) {
    if (!cell.formula) continue;
    if (formulaSheet === sheet && action === 'delete' && ((axis === 'row' && cell.row + 1 === coordinate) || (axis === 'column' && cell.column + 1 === coordinate))) continue;
    const updatedFormula = transformFormulaForStructure(cell.formula, formulaSheet.name, sheet.name, operation);
    if (updatedFormula === cell.formula && !definedNameTransform.changed) continue;
    const cellRange = formulaSheet.ranges.get(cell.address);
    if (!cellRange) fail('수식 셀 범위를 찾을 수 없어 구조 변경을 중단했습니다.');
    const cellXml = formulaSheet.xml.slice(cellRange.start, cellRange.end), formulaRange = findElementRanges(cellXml, 'f').get('');
    if (!formulaRange || formulaRange.selfClosing) fail('수식 XML을 안전하게 수정할 수 없어 구조 변경을 중단했습니다.');
    const formulaPatches = patchesFor(formulaSheet);
    if (updatedFormula !== cell.formula) formulaPatches.push({ start: cellRange.start + formulaRange.startTagEnd, end: cellRange.start + formulaRange.closingStart, replacement: xmlEscape(updatedFormula) });
    const valueRange = findElementRanges(cellXml, 'v').get('');
    if (valueRange && !valueRange.selfClosing) formulaPatches.push({ start: cellRange.start + valueRange.startTagEnd, end: cellRange.start + valueRange.closingStart, replacement: '' });
  }
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
  for (const addition of tableHeaderAdditions) {
    const rowNumber = addition.row + 1, rowRange = rowRanges.get(String(rowNumber));
    if (!rowRange || rowRange.selfClosing) fail('Excel 표 머리글 행을 찾을 수 없어 열을 삽입할 수 없습니다.');
    const following = [...sheet.cells.values()].filter(cell => cell.row === addition.row && cell.column >= addition.column).sort((left, right) => left.column - right.column)[0];
    const insertion = following ? sheet.ranges.get(following.address)?.start : rowRange.closingStart;
    if (!Number.isInteger(insertion)) fail('Excel 표 머리글 셀 위치를 안전하게 찾을 수 없습니다.');
    const previous = sheet.cells.get(columnLetters(addition.column) + rowNumber), previousRange = previous && sheet.ranges.get(previous.address);
    const previousXml = previousRange && sheet.xml.slice(previousRange.start, previousRange.startTagEnd), style = previousXml && /\bs=(['"])(\d+)\1/.exec(previousXml);
    const styleAttribute = style ? ' s="' + style[2] + '"' : '';
    const cell = '<c r="' + columnLetters(addition.column + 1) + rowNumber + '"' + styleAttribute + ' t="inlineStr"><is><t>' + xmlEscape(addition.name) + '</t></is></c>';
    patches.push({ start: insertion, end: insertion, replacement: cell });
  }
  const outputFiles = Object.assign({}, model.files);
  for (const [partPath, partPatches] of patchesByPath) {
    const partSheet = model.sheets.find(item => item.path === partPath);
    partPatches.sort((a, b) => b.start - a.start);
    let xml = partSheet.xml, previousStart = xml.length + 1;
    for (const patch of partPatches) {
      if (patch.end > previousStart) fail('XLSX 구조 변경 범위가 겹쳐 저장을 중단했습니다.');
      xml = xml.slice(0, patch.start) + patch.replacement + xml.slice(patch.end);
      previousStart = patch.start;
    }
    if (partPath === sheet.path) { xml = updateDimensionForStructure(xml, { axis, action, index }); xml = transformPageBreaksXml(xml, operation); }
    outputFiles[partPath] = new TextEncoder().encode(xml);
  }
  for (const [worksheetPath, removedIds] of hyperlinkRelationshipRemovals) {
    const relationshipPath = path.join(path.dirname(worksheetPath), '_rels', path.basename(worksheetPath) + '.rels');
    if (!outputFiles[relationshipPath]) fail('삭제한 하이퍼링크의 관계 파일을 찾을 수 없습니다.');
    const worksheetXml = new TextDecoder().decode(outputFiles[worksheetPath]);
    const remainingIds = new Set([...worksheetXml.matchAll(/<(?:[\w.-]+:)?hyperlink\b[^>]*>/g)].map(match => parseTagAttributes(match[0])['r:id']).filter(Boolean));
    const relationshipXml = new TextDecoder().decode(outputFiles[relationshipPath]);
    const removals = [];
    for (const id of removedIds) {
      if (remainingIds.has(id)) continue;
      const relationship = [...relationshipXml.matchAll(/<(?:[\w.-]+:)?Relationship\b[^>]*\/?\s*>/g)].find(match => parseTagAttributes(match[0]).Id === id);
      if (!relationship || !parseTagAttributes(relationship[0]).Type?.endsWith('/hyperlink')) fail('삭제한 하이퍼링크 관계를 안전하게 찾을 수 없습니다.');
      removals.push({ start: relationship.index, end: relationship.index + relationship[0].length });
    }
    if (!removals.length) continue;
    let updated = relationshipXml;
    for (const removal of removals.sort((left, right) => right.start - left.start)) updated = updated.slice(0, removal.start) + updated.slice(removal.end);
    if (/<(?:[\w.-]+:)?Relationship\b/.test(updated)) outputFiles[relationshipPath] = new TextEncoder().encode(updated);
    else delete outputFiles[relationshipPath];
    hyperlinkRelationshipPaths.add(relationshipPath);
  }
  const columnLettersForTable = column => {
    let value = column + 1, label = '';
    while (value) { value--; label = String.fromCharCode(65 + value % 26) + label; value = Math.floor(value / 26); }
    return label;
  };
  const changedTablePaths = new Set();
  for (const { table, range, columnChange } of tableRangeChanges) {
    const path = table.path;
    let xml = new TextDecoder().decode(outputFiles[path] || model.files[path]);
    xml = transformTableSortStateForStructure(xml, sheet.name, operation, table);
    if (columnChange) {
      const columns = [...table.columns];
      if (columnChange.action === 'insert') columns.splice(columnChange.index, 0, columnChange.name);
      else columns.splice(columnChange.index, 1);
      xml = resizeTableColumnDefinitions(xml, columns, columnChange);
      xml = transformTableFilterColumns(xml, columnChange.action, columnChange.index);
    }
    const ref = `${columnLettersForTable(range.left)}${range.top + 1}:${columnLettersForTable(range.right)}${range.bottom + 1}`;
    const filterRef = `${columnLettersForTable(range.left)}${range.top + 1}:${columnLettersForTable(range.right)}${range.bottom + 1 - (table.totalsRowCount || 0)}`;
    const patches = [];
    for (const tag of ['table', 'autoFilter']) {
      const range = findElementRanges(xml, tag).get('');
      if (!range) { if (tag === 'autoFilter' && !table.autoFilter) continue; fail(`Excel 표 ${tag} 범위를 찾을 수 없습니다.`); }
      const opening = xml.slice(range.start, range.startTagEnd);
      if (!/\bref="[^"]+"/.test(opening)) fail(`Excel 표 ${tag} 범위가 유효하지 않습니다.`);
      patches.push({ start: range.start, end: range.startTagEnd, replacement: replaceTagAttribute(opening, 'ref', tag === 'autoFilter' ? filterRef : ref) });
    }
    patches.sort((a, b) => b.start - a.start);
    let updated = xml;
    for (const patch of patches) updated = updated.slice(0, patch.start) + patch.replacement + updated.slice(patch.end);
    outputFiles[path] = new TextEncoder().encode(updated); changedTablePaths.add(path);
  }
  for (const owner of model.sheets) for (const table of owner.tables || []) {
    const originalXml = new TextDecoder().decode(outputFiles[table.path]), replacements = [];
    for (const name of ['calculatedColumnFormula', 'totalsRowFormula']) for (const range of findElementRanges(originalXml, name, undefined, true).values()) {
      if (range.selfClosing) continue;
      const node = parseXml(new TextEncoder().encode(originalXml.slice(range.start, range.end)), table.path).root;
      if (node.children.length) fail('복잡한 표 수식은 구조 변경을 지원하지 않습니다.');
      const formula = transformFormulaForStructure(node.text, owner.name, sheet.name, operation);
      if (formula !== node.text) replacements.push({ start: range.startTagEnd, end: range.closingStart, replacement: xmlEscape(formula) });
    }
    if (!replacements.length) continue;
    let xml = originalXml;
    for (const replacement of replacements.sort((a, b) => b.start - a.start)) xml = xml.slice(0, replacement.start) + replacement.replacement + xml.slice(replacement.end);
    outputFiles[table.path] = new TextEncoder().encode(xml); changedTablePaths.add(table.path);
  }
  const changedChartPaths = transformChartReferencesForStructure(model, sheet.name, operation, outputFiles), changedChartDrawingPaths = new Set();
  transformChartDrawingForStructure(model, sheet, operation, outputFiles, changedChartDrawingPaths);
  let workbookXml = definedNameTransform.xml;
  if (sheet.printAreas?.length) {
    const areas = sheet.printAreas.map(range => transformPrintArea(range, operation)).filter(Boolean);
    workbookXml = replaceWorkbookPrintDefinition(workbookXml, sheetIndex, areas.length ? areas.map(range => printAreaReference(sheet.name, range)).join(',') : null);
  }
  if (sheet.printTitles?.rows || sheet.printTitles?.columns) {
    const titles = transformPrintTitles(sheet.printTitles, operation);
    workbookXml = replaceWorkbookPrintDefinition(workbookXml, sheetIndex, printTitlesReference(sheet.name, titles) || null, '_xlnm.Print_Titles');
  }
  outputFiles['xl/workbook.xml'] = new TextEncoder().encode(enableFullCalculationOnLoad(workbookXml));
  const bytes = zipSync(outputFiles, { level: 0 });
  let checked = analyze(bytes, model.extension);
  if (!checked.writable) fail('구조 변경 후 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  if (checked.hasCharts) {
    for (const chart of checked.chartPaths) {
      outputFiles[chart.path] = patchChartCaches(checked.files[chart.path], chart.path, chart.sheetIndex, checked, null);
      changedChartPaths.add(chart.path);
    }
    checked = analyze(zipSync(outputFiles, { level: 0 }), model.extension);
    if (!checked.writable) fail('차트 캐시 갱신 후 구조 변경 검사에서 거절되었습니다: ' + checked.readOnlyReason);
  }
  if (sheet.autoFilter) {
    const filterPatch = patchAutoFilters(checked, [{ sheetIndex, filter: transformAutoFilter(sheet.autoFilter, operation) }]);
    checked = filterPatch.model;
  }
  for (const [name, before] of Object.entries(model.files)) {
    if (patchesByPath.has(name) || hyperlinkRelationshipPaths.has(name) || changedTablePaths.has(name) || changedChartPaths.has(name) || changedChartDrawingPaths.has(name) || name === 'xl/workbook.xml' && Buffer.compare(Buffer.from(outputFiles[name]), Buffer.from(model.files[name])) !== 0) continue;
    if (!checked.files[name] || Buffer.compare(Buffer.from(before), Buffer.from(checked.files[name])) !== 0) fail('비수정 패키지 항목이 달라졌습니다: ' + name);
  }
  if (axis === 'row' && sheet.tables?.length) {
    const metadata = [];
    for (const originalTable of sheet.tables) {
      const table = checked.sheets[sheetIndex].tables.find(item => item.name === originalTable.name);
      if (!table) continue;
      const first = originalTable.range.top + originalTable.headerRowCount, last = originalTable.range.bottom - originalTable.totalsRowCount, nextFirst = table.range.top + table.headerRowCount;
      const emptied = action === 'delete' && index === first && first === last && table.range.bottom - table.totalsRowCount < nextFirst;
      for (const [columnIndex, formula] of (originalTable.calculatedColumns || []).entries()) {
        if (!formula) continue;
        const anchor = action === 'delete' && index === first && last > first ? first + 1 : first;
        const movedAnchor = anchor + (action === 'insert' && index <= anchor ? 1 : 0) - (action === 'delete' && index < anchor ? 1 : 0);
        const transformed = transformFormulaForStructure(translateFormulaReferences(formula, anchor - first, 0), sheet.name, sheet.name, operation, emptied ? originalTable.range : null);
        metadata.push({ sheetIndex, tableName: table.name, column: table.range.left + columnIndex, formula: translateFormulaReferences(transformed, nextFirst - movedAnchor, 0) });
      }
    }
    if (metadata.length) checked = patchCalculatedColumns(checked, metadata).model;
    if (action === 'insert') checked = fillCalculatedRows(checked, checked.sheets[sheetIndex].tables.map(table => ({ sheetIndex, tableName: table.name, top: index, bottom: index }))).model;
  }
  return patchFilterVisibility(checked);

}

module.exports = { LIMITS, TABLE_STYLE_NAMES, analyze, calculateChartPreviews, children, conditionalStyleForCell, formatCellValue, inspectZip, parseRangeReference, parseXml, patchCell, patchCells, patchColumnWidths, patchRowHeights, patchRowHeightValues, validatePrintSettings, patchPrintSettings, patchPrintArea, patchDefinedName, patchPrintTitles, validatePrintTitles, patchHeaderFooter, patchPageBreaks, tableFilterRange, patchAutoFilters, patchFilterVisibility, patchHiddenRows, patchHiddenColumns, patchOutlineGroups, patchFreezePanes, patchMergedCells, planTableHeaderEdits, patchTableColumnNames, renameStructuredColumns, planTableExpansion, prepareTableExpansion, patchCalculatedColumns, planCalculatedColumn, fillCalculatedRows, patchAddTable, patchTableStyles, patchTableRanges, patchTableTotals, totalsFunctionForCell, patchAddChart, patchChartMetadata, patchChartPlacement, patchChartSource, patchDeleteChart, patchSheetVisibility, patchSheetZoom, patchSheetName, patchAddSheet, patchDeleteSheet, patchMoveSheet, transformStructure, translateFormulaReferences, rowNumber, safeUnzip, textContent };
module.exports.patchDataValidations = patchDataValidations;
module.exports.patchDataBarConditionalFormat = patchDataBarConditionalFormat;
module.exports.patchCellValueConditionalFormat = patchCellValueConditionalFormat;
module.exports.patchColorScaleConditionalFormat = patchColorScaleConditionalFormat;
module.exports.patchDuplicateValuesConditionalFormat = patchDuplicateValuesConditionalFormat;
module.exports.patchIconSetConditionalFormat = patchIconSetConditionalFormat;
module.exports.patchClearConditionalFormats = patchClearConditionalFormats;
module.exports.patchDeleteConditionalFormatRule = patchDeleteConditionalFormatRule;
module.exports.patchMoveConditionalFormatRule = patchMoveConditionalFormatRule;
module.exports.patchEditConditionalFormatRule = patchEditConditionalFormatRule;
module.exports.patchHyperlinks = patchHyperlinks;
