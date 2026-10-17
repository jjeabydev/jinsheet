'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const Module = require('node:module');
const vm = require('node:vm');
const { parseFragment } = require('parse5');
const { unzipSync, zipSync } = require('../vendor/node_modules/fflate');

class EventEmitter {
  constructor() { this.listeners = new Set(); this.event = listener => { this.listeners.add(listener); return { dispose: () => this.listeners.delete(listener) }; }; }
  fire(value) { for (const listener of this.listeners) listener(value); }
  dispose() { this.listeners.clear(); }
}
const originalLoad = Module._load;
Module._load = function(request, parent, isMain) { if (request === 'vscode') return { EventEmitter }; return originalLoad.call(this, request, parent, isMain); };
const { autoFilterRange, autoFilterOptions, updateAutoFilterColumn, updateAutoFilterCustomColumn, currentRegion, chartPreviewHtml, findMatches, buildReplaceEdits, buildScopedReplaceEdits, getRange, isValidSheetRequest, resolveInternalHyperlink, matchesSelectedSheet, estimateDocumentColumnWidth, navigateCell, selectionSummary, parseClipboard, inferClipboardCellKind, parseClipboardDate, parseClipboardTime, parseClipboardPercentage, parseClipboardCurrency, buildDateTimeShortcutEdit, buildAutoSumFormula, buildClipboardKindMatrix, buildClipboardFormulaMatrix, isSameWorkbookClipboard, buildClearEdits, buildMoveEdits, buildSparseColumnMoveEdits, buildFillEdits, buildFillRangeEdits, buildSortEdits, planTextToColumns, buildTextToColumnsEdits, planRemoveDuplicates, removeDuplicateRows, buildNumberFormatEdits, buildFontFormatEdits, buildFontColorEdits, buildAlignmentFormatEdits, buildWrapTextEdits, buildAutomaticRowFitRequest, buildRowFitRequest, validateRowFitResult, buildBorderFormatEdits, buildFillFormatEdits, buildPasteEdits, buildSparseColumnPasteEdits, serializeClipboard, buildCopyPayload, JinSheetDocument, JinSheetProvider, buildPrintDocument, parsePrintTitleInputs, renderHtml, viewModel } = require('../src/extension');
const { load } = require('../src/csv');
const { analyze, formatCellValue } = require('../src/xlsx');
const { messagesFor } = require('../src/i18n');
const { chartPrintGeometry } = require('../src/chart-print');
const { buildDataValidationRule } = require('../src/extension');
let activeDocument = null;

test('클립보드 기록은 요청 순서로 완료되고 실패 뒤에도 다음 작업이 진행된다', async () => {
  const queue = Object.assign(Object.create(JinSheetProvider.prototype), { clipboardWrite: Promise.resolve() }), started = [], completed = []; let release;
  const writer = async value => { started.push(value); if (value === 'first') await new Promise(resolve => { release = resolve; }); if (value === 'fail') throw new Error('clipboard failed'); completed.push(value); };
  const first = queue.writeClipboardText('first', writer), read = queue.readClipboardText(() => completed.at(-1)), second = queue.writeClipboardText('second', writer);
  await Promise.resolve(); assert.deepEqual(started, ['first']); let finished = false; queue.clipboardWrite.then(() => { finished = true; }); await Promise.resolve(); assert.equal(finished, false);
  release(); await Promise.all([first, second]); await queue.clipboardWrite; assert.deepEqual(completed, ['first', 'second']); assert.equal(await read, 'first', 'a later copy cannot overtake an earlier paste request');
  await assert.rejects(JinSheetProvider.prototype.writeClipboardText.call(queue, 'fail', writer), /clipboard failed/);
  await JinSheetProvider.prototype.writeClipboardText.call(queue, 'after', writer); assert.equal(completed.at(-1), 'after');
});

function isWithin(node, ancestor) { for (let current = node; current; current = current.parentElement) if (current === ancestor) return true; return false; }

function minimalXlsx(sheetViews = '', rows = '<row r="1"><c r="A1" t="inlineStr"><is><t>old</t></is></c></row>', autoFilter = '') {
  const xml = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="s1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="s1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    'xl/worksheets/sheet1.xml': `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${sheetViews}<sheetData>${rows}</sheetData>${autoFilter}</worksheet>`
  };
  return zipSync(Object.fromEntries(Object.entries(xml).map(([name, value]) => [name, new TextEncoder().encode(value)])));
}
function tableXlsx() {
  const files = Object.fromEntries(Object.entries(unzipSync(minimalXlsx('', '<row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Alpha</t></is></c><c r="B2"><v>10</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>Beta</t></is></c><c r="B3"><v>20</v></c></row><row r="4"><c r="B4"><f>SUBTOTAL(9,B2:B3)</f><v>30</v></c><c r="D4"><f>SUM(Sales[Amount])</f><v>0</v></c></row>'))).map(([name, bytes]) => [name, new TextDecoder().decode(bytes)]));
  files['[Content_Types].xml'] = files['[Content_Types].xml'].replace('</Types>', '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>');
  files['xl/worksheets/sheet1.xml'] = files['xl/worksheets/sheet1.xml'].replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ').replace('</worksheet>', '<tableParts count="1"><tablePart r:id="t1"/></tableParts></worksheet>');
  files['xl/worksheets/_rels/sheet1.xml.rels'] = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="t1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>';
  files['xl/tables/table1.xml'] = '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:B3"><autoFilter ref="A1:B3"/><tableColumns count="2"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Amount"/></tableColumns></table>';
  return zipSync(Object.fromEntries(Object.entries(files).map(([name, value]) => [name, new TextEncoder().encode(value)])));
}
function twoSheetValidationXlsx(useDefinedName = false) {
  const files = Object.fromEntries(Object.entries(unzipSync(minimalXlsx())).map(([name, value]) => [name, new TextDecoder().decode(value)]));
  files['xl/workbook.xml'] = files['xl/workbook.xml'].replace('</sheets>', '<sheet name="Lists" sheetId="2" r:id="rId2"/></sheets>');
  if (useDefinedName) files['xl/workbook.xml'] = files['xl/workbook.xml'].replace('</workbook>', '<definedNames><definedName name="ColorList">Lists!$A$1:$A$2</definedName></definedNames></workbook>');
  files['xl/_rels/workbook.xml.rels'] = files['xl/_rels/workbook.xml.rels'].replace('</Relationships>', '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>');
  files['[Content_Types].xml'] = files['[Content_Types].xml'].replace('</Types>', '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  files['xl/worksheets/sheet1.xml'] = files['xl/worksheets/sheet1.xml'].replace('</worksheet>', `<dataValidations count="1"><dataValidation type="list" sqref="B1"><formula1>${useDefinedName ? 'ColorList' : 'Lists!$A$1:$A$2'}</formula1></dataValidation></dataValidations></worksheet>`);
  files['xl/worksheets/sheet2.xml'] = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Red</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Blue</t></is></c></row></sheetData></worksheet>';
  return zipSync(Object.fromEntries(Object.entries(files).map(([name, value]) => [name, new TextEncoder().encode(value)])));
}
function twoSheetSearchXlsx(validation = false) {
  const files = Object.fromEntries(Object.entries(unzipSync(twoSheetValidationXlsx())).map(([name, value]) => [name, new TextDecoder().decode(value)]));
  files['xl/worksheets/sheet1.xml'] = files['xl/worksheets/sheet1.xml'].replace('<t>old</t>', '<t>Red north</t>');
  files['xl/worksheets/sheet2.xml'] = files['xl/worksheets/sheet2.xml'].replace('<t>Red</t>', '<t>Red south</t>');
  if (validation) files['xl/worksheets/sheet2.xml'] = files['xl/worksheets/sheet2.xml'].replace('</worksheet>', '<dataValidations count="1"><dataValidation type="list" errorStyle="warning" sqref="A1"><formula1>"Red south"</formula1></dataValidation></dataValidations></worksheet>');
  return zipSync(Object.fromEntries(Object.entries(files).map(([name, value]) => [name, new TextEncoder().encode(value)])));
}
function uri(fsPath) { return { fsPath, scheme: 'file', toString: () => fsPath }; }

test('차트 미리보기는 범주·제목 데이터를 HTML에서 이스케이프하고 읽기 전용 시트에 표시한다', () => {
  const markup = chartPreviewHtml([{ type: 'bar', title: '<script>alert(1)</script>', series: [{ name: '<img src=x onerror=1>', categories: ['<svg/onload=1>'], values: [5] }] }]);
  assert.match(markup, /&lt;script&gt;alert\(1\)&lt;\/script&gt;/);
  assert.doesNotMatch(markup, /<script|<img|<svg\//i);
  const grid = runGridWebview(renderHtml(), { kind: 'xlsx', rowCount: 1, columnCount: 1 });
  grid.dispatch({ type: 'model', kind: 'xlsx', readOnly: true, reason: '차트 구조는 읽기 전용', rowCount: 1, columnCount: 1, sheetIndex: 0, chartPreviewHtml: markup });
  assert.equal(grid.elements.get('chartPreview').innerHTML, markup);
  assert.equal(grid.elements.get('chartPreview').hidden, false);
});

test('편집 가능한 차트 미리보기는 제목·계열 입력을 이스케이프하고 편집 메시지를 보낸다', () => {
  const chart = { type: 'bar', title: '<Quarterly>', series: [{ name: 'Sales & costs', categories: ['Q1'], values: [10] }] };
  const markup = chartPreviewHtml([chart], 'ko', true);
  assert.match(markup, /data-chart-index="0"/);
  assert.match(markup, /data-chart-title maxlength="512" value="&lt;Quarterly&gt;"/);
  assert.match(markup, /data-chart-series-name maxlength="256" value="Sales &amp; costs"/);
  assert.match(markup, /차트 정보 적용/);
  assert.doesNotMatch(chartPreviewHtml([chart], 'ko', false), /data-chart-edit/);
  const manySeries = { type: 'bar', title: '', series: Array.from({ length: 9 }, (_, index) => ({ name: `Series ${index + 1}`, categories: ['Q1'], values: [index + 1] })) };
  assert.equal((chartPreviewHtml([manySeries], 'ko', true).match(/data-chart-series-name/g) || []).length, 9);
  const { elements, messages } = runGridWebview(renderHtml(), { kind: 'xlsx' });
  const title = { value: 'Quarterly report' }, seriesName = { value: 'Net sales' };
  const card = { dataset: { chartIndex: '0' }, querySelector: selector => selector === '[data-chart-title]' ? title : null, querySelectorAll: () => [seriesName] };
  const button = { closest: selector => selector === '[data-chart-index]' ? card : null };
  const target = { closest: selector => selector === '[data-chart-edit]' ? button : null };
  elements.get('chartPreview').listeners.get('click')({ target });
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'edit-chart', chartIndex: 0, title: 'Quarterly report', seriesNames: ['Net sales'], sheetIndex: 0 });
});

test('차트 원본 변경 버튼은 선택한 셀 범위와 차트 인덱스를 전송한다', () => {
  const { document, documentListeners, elements, messages } = runGridWebview(renderHtml(), { kind: 'xlsx', rowCount: 3, columnCount: 3 });
  const card = { dataset: { chartIndex: '1' } }, button = { closest: () => card }, target = { closest: selector => selector === '[data-chart-source]' ? button : null };
  const count = messages.length;
  elements.get('chartPreview').listeners.get('click')({ target });
  assert.equal(messages.length, count, 'a single cell cannot be used as a chart source');
  const first = document.querySelector('[data-row="0"][data-column="0"]'), last = document.querySelector('[data-row="2"][data-column="2"]');
  first.onmousedown({ target: first, preventDefault() {}, clientX: 10, clientY: 40, shiftKey: false });
  last.onmouseenter(); documentListeners.get('mouseup')();
  elements.get('chartPreview').listeners.get('click')({ target });
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'chart-source', chartIndex: 1, range: { top: 0, bottom: 2, left: 0, right: 2 }, sheetIndex: 0 });
});

test('차트 삭제 버튼은 범위 선택 없이 해당 차트 인덱스를 전송한다', () => {
  const grid = runGridWebview(renderHtml(), { kind: 'xlsx' });
  const card = { dataset: { chartIndex: '2' } }, button = { closest: () => card }, target = { closest: selector => selector === '[data-chart-delete]' ? button : null };
  grid.elements.get('chartPreview').listeners.get('click')({ target });
  assert.deepEqual(JSON.parse(JSON.stringify(grid.messages.at(-1))), { type: 'delete-chart', chartIndex: 2, sheetIndex: 0 });
  grid.dispatch({ type: 'model', kind: 'xlsx', readOnly: true, rowCount: 2, columnCount: 2, sheetIndex: 0 });
  const count = grid.messages.length;
  grid.elements.get('chartPreview').listeners.get('click')({ target });
  assert.equal(grid.messages.length, count);
});

test('차트 미리보기는 저장 전 값·수식 의존·머리글 편집과 undo·redo를 반영한다', async t => {
  const rows = '<row r="1"><c r="A1" t="inlineStr"><is><t>Month</t></is></c><c r="B1" t="inlineStr"><is><t>Sales</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Jan</t></is></c><c r="B2"><f>D2*2</f><v>10</v></c><c r="D2"><v>5</v></c></row>';
  const bytes = minimalXlsx('', rows), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-live-chart-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'chart.xlsx'); await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  document.applyAddChart(0, { top: 0, bottom: 1, left: 0, right: 1 }); await document.save();
  const beforeBytes = Buffer.from(document.state.files['xl/charts/chart1.xml']);
  let operation; document.onDidChange(change => { operation = change; });
  document.applyCellEdit(1, 3, '7');
  assert.deepEqual(viewModel(document).charts[0].series[0].values, [14]);
  assert.deepEqual(document.state.sheets[0].charts[0].series[0].values, [10]);
  assert.deepEqual(Buffer.from(document.state.files['xl/charts/chart1.xml']), beforeBytes);
  operation.undo(); assert.deepEqual(viewModel(document).charts[0].series[0].values, [10]);
  operation.redo(); assert.deepEqual(viewModel(document).charts[0].series[0].values, [14]);
  document.applyCellEdit(0, 1, 'Revenue'); document.applyCellEdit(1, 0, 'Mar');
  const live = viewModel(document).charts[0];
  assert.equal(live.series[0].name, 'Revenue'); assert.deepEqual(live.series[0].categories, ['Mar']);
  await document.save();
  assert.deepEqual(analyze(await fs.readFile(file)).sheets[0].charts[0], live);
  document.dispose();
});

test('차트 인쇄는 사용 범위 확장·라이브 값·인쇄 영역·제외를 반영하고 원본을 바꾸지 않는다', () => {
  const bytes = minimalXlsx('', '<row r="1"><c r="A1" t="inlineStr"><is><t>X</t></is></c><c r="B1" t="inlineStr"><is><t>Y</t></is></c></row><row r="2"><c r="A2"><v>2</v></c><c r="B2"><v>4</v></c></row>');
  const document = new JinSheetDocument(uri('/tmp/chart-print.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  document.applyAddChart(0, { top: 0, bottom: 1, left: 0, right: 1 }, 'scatter');
  document.applyCellEdit(1, 0, '5');
  const original = Buffer.from(document.state.bytes), dirty = document.dirty;
  let html = buildPrintDocument(document); assert.match(html, /data-print-chart=/); assert.match(html, /<svg/); assert.match(html, /data-print-row="9"/);
  assert.ok(html.includes(chartPreviewHtml(viewModel(document).charts, 'en', false, true)));
  assert.deepEqual(Buffer.from(document.state.bytes), original); assert.equal(document.dirty, dirty);
  html = buildPrintDocument(document, 0, null, 'en', undefined, false); assert.doesNotMatch(html, /data-print-chart=/); assert.doesNotMatch(html, /data-print-row="9"/); assert.match(html, /charts are not included/);
  document.applyPrintArea(0, { top: 0, bottom: 1, left: 0, right: 1 }); assert.doesNotMatch(buildPrintDocument(document), /data-print-chart=/);
  document.state.sheets[0].chartAnchors[0] = null; assert.throws(() => buildPrintDocument(document), /チャート|차트 위치/);
  assert.doesNotThrow(() => buildPrintDocument(document, 0, null, 'en', undefined, false)); document.dispose();
});

test('차트 인쇄는 표시 한계를 넘는 데이터를 잘라내지 않고 차트 제외를 안내한다', () => {
  const chart = { type: 'line', title: 'Full', series: [{ name: 'Many', categories: [], values: Array.from({ length: 10001 }, (_, i) => i) }] };
  assert.throws(() => chartPreviewHtml([chart], 'ko', false, true), /10,000/);
  assert.match(chartPreviewHtml([chart]), /<svg/);
  chart.series[0].values = [1, 2]; chart.type = 'doughnut'; chart.series.push({ ...chart.series[0] });
  assert.throws(() => chartPreviewHtml([chart], 'ko', false, true), /1계열/);
  const grid = runGridWebview(renderHtml(), { kind: 'xlsx' }); grid.elements.get('printIncludeCharts').checked = true;
  grid.elements.get('printSheet').onclick(); assert.equal(grid.messages.at(-1).includeCharts, true);
  grid.elements.get('printIncludeCharts').checked = false; grid.elements.get('printSheet').onclick(); assert.equal(grid.messages.at(-1).includeCharts, false);
});

test('인쇄 차트는 32계열과 10,000개 점을 표시하고 범례 높이를 늘린다', () => {
  const chart = { type: 'bar', title: 'Complete', series: Array.from({ length: 32 }, (_, series) => ({ name: `Series-${series}`, categories: [], values: Array.from({ length: 100 }, (_, point) => point + 1) })) };
  const html = chartPreviewHtml([chart], 'en', false, true);
  assert.equal((html.match(/data-chart-point=/g) || []).length, 3200);
  assert.equal((html.match(/data-print-legend=/g) || []).length, 32); assert.match(html, /Series-31/);
  const colors = [...html.matchAll(/data-print-legend="\d+"><rect [^>]*fill="([^"]+)"/g)].map(match => match[1]);
  assert.equal(new Set(colors).size, 32);
  assert.match(html, /viewBox="0 0 640 316"/);
  chart.type = 'line'; chart.series = [{ name: 'Long', categories: [], values: Array.from({ length: 10000 }, (_, index) => index) }];
  const long = chartPreviewHtml([chart], 'en', false, true), points = /<polyline points="([^"]+)"/.exec(long)[1].split(' ');
  assert.equal(points.length, 10000); assert.doesNotMatch(long, /NaN|Infinity/);
  const preview = chartPreviewHtml([chart]); assert.equal(/<polyline points="([^"]+)"/.exec(preview)[1].split(' ').length, 60);
  chart.series = Array.from({ length: 5 }, (_, index) => ({ ...chart.series[0], name: 'Limit-' + index }));
  const maximum = chartPreviewHtml([chart], 'en', false, true);
  assert.equal([...maximum.matchAll(/<polyline points="([^"]+)"/g)].reduce((sum, match) => sum + match[1].split(' ').length, 0), 50000);
  chart.series = Array.from({ length: 32 }, () => ({ name: 'Too much', categories: [], values: Array(2000).fill(1) }));
  assert.throws(() => chartPreviewHtml([chart], 'en', false, true), /50,000/);
  chart.type = 'bar'; chart.series = [{ name: 'Large markup', categories: Array(10000).fill('&'.repeat(256)), values: Array(10000).fill(1) }];
  assert.throws(() => chartPreviewHtml([chart], 'en', false, true), /출력이 너무 큽니다/);
  assert.throws(() => chartPreviewHtml(Array(33).fill(chart), 'en', false, true), /32개 차트/);
});

test('원형 차트는 모든 조각을 인쇄하고 많은 범주의 범례 요약을 명시한다', () => {
  const chart = { type: 'pie', title: 'Categories', series: [{ name: 'All', categories: Array.from({ length: 80 }, (_, index) => `Category-${index}`), values: Array(80).fill(1) }] };
  const html = chartPreviewHtml([chart], 'en', false, true);
  assert.equal((html.match(/data-chart-point=/g) || []).length, 80); assert.match(html, /data-legend-summary/); assert.match(html, /All 80 slices shown; 68 legend labels omitted/); assert.match(html, /Category-79/);
  assert.equal((html.match(/data-print-legend=/g) || []).length, 12);
  assert.doesNotMatch(html, /NaN|Infinity/);
});

test('문서 인쇄 경로는 9계열·80개 원본 값과 마지막 미저장 편집을 끝까지 출력한다', () => {
  const rows = Array.from({ length: 81 }, (_, row) => '<row r="' + (row + 1) + '">' + Array.from({ length: 10 }, (_, column) => {
    const address = String.fromCharCode(65 + column) + (row + 1);
    return !row || !column ? `<c r="${address}" t="inlineStr"><is><t>${row ? 'Item-' + row : 'S' + column}</t></is></c>` : `<c r="${address}"><v>${row * column}</v></c>`;
  }).join('') + '</row>').join('');
  const bytes = minimalXlsx('', rows), document = new JinSheetDocument(uri('/tmp/full-series-print.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  document.applyAddChart(0, { top: 0, bottom: 80, left: 0, right: 9 });
  document.applyPrintSettings(0, { ...document.state.sheets[0].printSettings, fitToPage: true, fitToWidth: 1, fitToHeight: 1 });
  let operation; document.onDidChange(change => { operation = change; }); document.applyCellEdit(80, 9, '999');
  const snapshot = Buffer.from(document.state.bytes), dirty = document.dirty, html = buildPrintDocument(document);
  assert.equal((html.match(/data-chart-point=/g) || []).length, 720); assert.equal((html.match(/data-print-legend=/g) || []).length, 9);
  assert.match(html, /aria-label="S9: Item-80: 999"/); assert.equal(document.state.sheets[0].charts[0].series[8].values[79], 720);
  assert.deepEqual(Buffer.from(document.state.bytes), snapshot); assert.equal(document.dirty, dirty);
  operation.undo(); assert.match(buildPrintDocument(document), /aria-label="S9: Item-80: 720"/);
  operation.redo(); assert.match(buildPrintDocument(document), /aria-label="S9: Item-80: 999"/); document.dispose();
});

test('차트 배치는 미저장 열 너비·저장 전후 Undo·행 높이 변경에서도 크기를 유지한다', async t => {
  const rows = '<row r="1"><c r="A1" t="inlineStr"><is><t>Label</t></is></c><c r="B1" t="inlineStr"><is><t>Value</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>First</t></is></c><c r="B2"><v>10</v></c></row>';
  const bytes = minimalXlsx('', rows), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-chart-placement-')); t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'placement.xlsx'); await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, ''); document.applyAddChart(0, { top: 0, bottom: 1, left: 0, right: 1 });
  const box = () => chartPrintGeometry(viewModel(document), document.state.sheets[0].chartAnchors, document.state.sheets[0]).boxes[0];
  document.applyColumnWidth(5, 300); const visible = box();
  let operation; document.onDidChange(change => { operation = change; }); document.applyChartPlacement(0, 0, 'oneCell'); const placementOperation = operation;
  assert.equal(document.columnWidthDrafts.size, 0); assert.ok(Math.abs(box().width - visible.width) < .001);
  placementOperation.undo(); assert.equal(document.columnWidthDrafts.size, 1); assert.equal(document.state.sheets[0].chartPlacements[0], 'twoCell');
  placementOperation.redo(); assert.equal(document.state.sheets[0].chartPlacements[0], 'oneCell');
  document.applyColumnWidth(5, 400); assert.ok(Math.abs(box().width - visible.width) < .001); assert.match(buildPrintDocument(document), /data-print-chart=/);
  await document.save(); assert.ok(Math.abs(box().width - visible.width) < .001);
  placementOperation.undo(); assert.equal(document.state.sheets[0].chartPlacements[0], 'twoCell'); placementOperation.redo(); assert.equal(document.state.sheets[0].chartPlacements[0], 'oneCell');
  document.applyColumnWidth(0, 200); const moved = box(); document.applyChartPlacement(0, 0, 'absolute');
  document.applyColumnWidth(0, 250); document.applyRowHeights(0, 3, 3, 39);
  for (const key of ['x', 'y', 'width', 'height']) assert.ok(Math.abs(box()[key] - moved[key]) < .001, key);
  await document.save(); const saved = analyze(await fs.readFile(file)), savedBox = chartPrintGeometry(saved.sheets[0], saved.sheets[0].chartAnchors).boxes[0];
  for (const key of ['x', 'y', 'width', 'height']) assert.ok(Math.abs(savedBox[key] - moved[key]) < .001, key); document.dispose();
});

test('차트 배치 UI는 현재 모드를 표시하고 선택한 모드만 메시지로 보낸다', () => {
  const chart = { type: 'line', series: [{ name: 'S', categories: [], values: [1, 2] }] };
  assert.match(chartPreviewHtml([chart], 'ko', true, false, ['absolute']), /value="absolute" selected/);
  assert.doesNotMatch(chartPreviewHtml([chart], 'ko', false, false, ['absolute']), /data-chart-placement/);
  const grid = runGridWebview(renderHtml(), { kind: 'xlsx' });
  const card = { dataset: { chartIndex: '0' }, querySelector: selector => selector === '[data-chart-placement]' ? { value: 'oneCell' } : null }, button = { closest: () => card }, target = { closest: selector => selector === '[data-chart-placement-apply]' ? button : null };
  grid.elements.get('chartPreview').listeners.get('click')({ target });
  assert.deepEqual(JSON.parse(JSON.stringify(grid.messages.at(-1))), { type: 'chart-placement', chartIndex: 0, placement: 'oneCell', sheetIndex: 0 });
});

test('기본 셀 치수와 1px·1785px 열은 그리드·인쇄·저장 좌표에 그대로 반영된다', () => {
  const bytes = minimalXlsx(), document = new JinSheetDocument(uri('/tmp/native-dimensions.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  document.applyColumnWidth(1, 1); document.applyColumnWidth(2, 1785);
  const model = viewModel(document), grid = runGridWebview(renderHtml(), { kind: 'xlsx', rowCount: 2, columnCount: 3, columnWidths: model.columnWidths, defaultColumnWidth: model.defaultColumnWidth, defaultRowHeight: model.defaultRowHeight });
  const cell = column => grid.elements.get('canvas').children.find(item => item.classList.contains('cell') && item.dataset.row === '0' && item.dataset.column === String(column));
  assert.equal(cell(0).style.width, '64px'); assert.equal(cell(0).style.height, '20px'); assert.equal(cell(1).style.width, '1px'); assert.equal(cell(2).style.width, '1785px');
  const html = buildPrintDocument(document, 0, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 }); assert.match(html, /width:1px/); assert.match(html, /width:1785px/);
  const saved = analyze(document.createOutput()); assert.equal(saved.sheets[0].columnWidths[1], .14); assert.equal(saved.sheets[0].columnWidths[2], 255);
  assert.equal(saved.sheets[0].defaultColumnWidth, 64); assert.equal(saved.sheets[0].defaultRowHeight, 20); document.dispose();
});

test('Excel 기본 치수의 고정 차트는 좁고 넓은 열 변경과 재해석 뒤에도 같은 물리 좌표를 유지한다', () => {
  const bytes = minimalXlsx('', '<row r="1"><c r="A1" t="inlineStr"><is><t>X</t></is></c><c r="B1" t="inlineStr"><is><t>Y</t></is></c></row><row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c></row>');
  const document = new JinSheetDocument(uri('/tmp/native-chart-dimensions.xlsx'), 'xlsx', analyze(bytes), bytes, true, ''); document.applyAddChart(0, { top: 0, bottom: 1, left: 0, right: 1 });
  document.applyChartPlacement(0, 0, 'oneCell');
  const box = () => chartPrintGeometry(viewModel(document), document.state.sheets[0].chartAnchors, document.state.sheets[0]).boxes[0];
  assert.deepEqual([box().x, box().width, box().height], [192, 448, 200]);
  document.applyColumnWidth(5, 1785); assert.equal(box().width, 448); document.applyColumnWidth(0, 1); assert.equal(box().x, 129);
  document.applyChartPlacement(0, 0, 'absolute'); document.applyColumnWidth(0, 1785); document.applyRowHeights(0, 3, 3, 39);
  const saved = analyze(document.createOutput()), actual = chartPrintGeometry(saved.sheets[0], saved.sheets[0].chartAnchors).boxes[0];
  for (const [key, value] of Object.entries({ x: 129, y: 0, width: 448, height: 200 })) assert.ok(Math.abs(actual[key] - value) < .001, key); document.dispose();
});

test('차트 인쇄는 선 사이 빈 값·원형 단일 값과 음수를 빠뜨리지 않고 출력한다', () => {
  const chart = { type: 'line', title: '<unsafe>', series: [{ name: 'Gap', categories: ['A', 'B', 'C', 'D', 'E'], values: [1, 2, null, 4, 5] }] };
  const line = chartPreviewHtml([chart], 'en', false, true);
  assert.equal((line.match(/<polyline /g) || []).length, 2); assert.match(line, /&lt;unsafe&gt;/); assert.doesNotMatch(line, /<unsafe>/);
  chart.type = 'pie'; chart.series[0].values = [-5];
  const single = chartPreviewHtml([chart], 'en', false, true); assert.match(single, /<circle cx="320" cy="91" r="68"/); assert.doesNotMatch(single, /chart-empty/);
  chart.series[0].values = [-5, 3]; assert.equal((chartPreviewHtml([chart], 'en', false, true).match(/<path /g) || []).length, 2);
  chart.series[0].values = [1e308, 1e308]; assert.throws(() => chartPreviewHtml([chart], 'en', false, true), /수치 합계/);
  chart.type = 'line'; chart.series[0].values = [-1e308, 1e308]; assert.throws(() => chartPreviewHtml([chart], 'en', false, true), /수치 범위/);
});

test('분산형 차트 X·Y 초안을 표시하고 모든 값이 없어도 편집 UI를 유지한다', () => {
  const rows = '<row r="1"><c r="A1" t="inlineStr"><is><t>X</t></is></c><c r="B1" t="inlineStr"><is><t>Y</t></is></c></row><row r="2"><c r="A2"><v>2</v></c><c r="B2"><v>4</v></c></row>';
  const bytes = minimalXlsx('', rows), document = new JinSheetDocument(uri('/tmp/live-scatter.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  document.applyAddChart(0, { top: 0, bottom: 1, left: 0, right: 1 }, 'scatter');
  document.applyCellEdit(1, 0, '5'); assert.deepEqual(viewModel(document).charts[0].series[0].xValues, [5]);
  document.applyCellEdit(1, 1, '');
  const model = viewModel(document);
  assert.deepEqual(model.charts[0].series[0].values, [null]);
  assert.match(model.chartPreviewHtml, /No chart data to display/); assert.match(model.chartPreviewHtml, /data-chart-delete/);
  document.dispose();
});

test('원형·도넛 차트 미리보기는 양수 조각과 범례를 SVG로 렌더링한다', () => {
  const markup = chartPreviewHtml([{ type: 'doughnut', title: 'Share', series: [{ name: 'Sales', categories: ['A&B', 'Other'], values: [3, 7] }] }]);
  assert.match(markup, /<path d="M 320 91 L/);
  assert.match(markup, /<circle cx="320" cy="91" r="34"/);
  assert.match(markup, /A&amp;B/);
  assert.equal((markup.match(/<path /g) || []).length, 2);
});

test('영역형 차트 미리보기는 기준선에 닫힌 채움 영역과 계열선을 렌더링한다', () => {
  const markup = chartPreviewHtml([{ type: 'area', title: 'Trend', series: [{ name: 'Total', categories: ['A', 'B', 'C'], values: [2, 5, 3] }] }]);
  assert.match(markup, /<polygon points="[^\"]+" fill="#4f81bd" opacity="\.22"\/>/);
  assert.match(markup, /<polyline points="[^\"]+" fill="none" stroke="#4f81bd"/);
});

test('분산형 차트 미리보기는 X·Y 값을 독립 축의 점으로 표시한다', () => {
  const markup = chartPreviewHtml([{ type: 'scatter', title: 'Correlation', series: [{ name: 'Sample', xValues: [1, 2, 3], values: [4, 6, 5] }] }]);
  assert.equal((markup.match(/<circle cx=/g) || []).length, 4);
  assert.match(markup, /<line x1="48" y1="12" x2="48" y2="154"/);
});

test('하이퍼링크 대상은 셀에 전달되고 내부 셀·범위 주소를 안전하게 해석한다', () => {
  const files = Object.fromEntries(Object.entries(unzipSync(twoSheetValidationXlsx())).map(([name, value]) => [name, new TextDecoder().decode(value)]));
  files['xl/worksheets/sheet1.xml'] = files['xl/worksheets/sheet1.xml'].replace('</worksheet>', '<hyperlinks><hyperlink ref="A1" location="\'Lists\'!$A$1:$A$2" display="Jump" tooltip="Open list"/></hyperlinks></worksheet>');
  const bytes = zipSync(Object.fromEntries(Object.entries(files).map(([name, value]) => [name, new TextEncoder().encode(value)])));
  const document = new JinSheetDocument(uri('/tmp/hyperlink.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  const range = getRange(document, { requestId: 1, sheetIndex: 0, rowStart: 0, columnStart: 0, rowCount: 1, columnCount: 1 });
  assert.deepEqual(range.rows[0][0].hyperlink, { display: 'Jump', tooltip: 'Open list' });
  assert.deepEqual(resolveInternalHyperlink(document, 0, "'Lists'!$A$1:$A$2"), { sheetIndex: 1, row: 0, column: 0, range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 } });
  assert.deepEqual(resolveInternalHyperlink(document, 1, '$B$2'), { sheetIndex: 1, row: 1, column: 1, range: { startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 } });
  assert.equal(resolveInternalHyperlink(document, 0, "'Missing'!A1"), null);
  assert.equal(resolveInternalHyperlink(document, 0, 'XFE1'), null);
  document.dispose();
});

test('하이퍼링크가 있는 범위는 정렬·이동·채우기를 차단한다', () => {
  const document = { kind: 'xlsx', state: { sheets: [{ hyperlinks: [{ top: 0, bottom: 0, left: 0, right: 0 }], cells: new Map() }] }, getCell: () => null };
  const range = { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 };
  assert.throws(() => buildSortEdits(document, range, 0, 'ascending', false, false), /하이퍼링크/);
  assert.throws(() => buildMoveEdits(document, { ...range, endRow: 0, endColumn: 0 }, 2, 0), /하이퍼링크/);
  assert.throws(() => buildFillEdits(document, range, 'down'), /하이퍼링크/);
  assert.throws(() => buildFillRangeEdits(document, { ...range, endRow: 0, endColumn: 0 }, range), /하이퍼링크/);
});

test('XLSX 인라인 목록 데이터 유효성은 제안하고 허용되지 않은 셀 입력을 거부한다', () => {
  const validation = '<dataValidations count="2"><dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="A1"><formula1>"Open,In Progress,Closed"</formula1></dataValidation><dataValidation type="list" showErrorMessage="1" sqref="A2:A3"><formula1>"Open,In Progress,Closed"</formula1></dataValidation></dataValidations>';
  const bytes = minimalXlsx('', '<row r="1"><c r="A1" t="inlineStr"><is><t>Open</t></is></c></row>', validation);
  const document = new JinSheetDocument(uri('/tmp/data-validation.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  assert.equal(document.writable, true, document.state.readOnlyReason);
  assert.deepEqual(document.state.sheets[0].dataValidations[0].values, ['Open', 'In Progress', 'Closed']);
  assert.throws(() => document.applyCellEdit(0, 0, 'Pending'), /허용 목록/);
  document.applyCellEdit(0, 0, 'Closed');
  const range = getRange(document, { requestId: 1, rowStart: 0, columnStart: 0, rowCount: 1, columnCount: 1 });
  assert.deepEqual(range.rows[0][0].validation, { type: 'list', values: ['Open', 'In Progress', 'Closed'], operator: null, limits: null, allowBlank: true, showDropDown: false, showInputMessage: true, showErrorMessage: true, errorTitle: '', error: '', promptTitle: '', prompt: '' });
  assert.throws(() => document.applyCellEdit(0, 0, '="Pending"'), /허용 목록/, 'formula results are checked against the list');
  document.applyCellEdit(0, 0, '="Open"');
  assert.equal(getRange(document, { requestId: 2, rowStart: 0, columnStart: 0, rowCount: 1, columnCount: 1 }).rows[0][0].value, 'Open');
  assert.throws(() => document.applyCellEdit(1, 0, ''), /허용 목록/, 'blank values are rejected when allowBlank is false');
  assert.throws(() => document.applyCellEdits([{ row: 1, column: 0, value: 'Pending' }]), /허용 목록/, 'paste cannot bypass list validation');
});

test('데이터 유효성 규칙은 목록과 숫자·날짜·시간 조건을 생성한다', () => {
  const bytes = minimalXlsx(), document = new JinSheetDocument(uri('/tmp/data-validation-create.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  assert.deepEqual(buildDataValidationRule(document, 0, { type: 'list', source: '"New, York",London' }).values, ['New, York', 'London']);
  assert.deepEqual(buildDataValidationRule(document, 0, { type: 'list', source: '=A1:A2' }).source, { top: 0, bottom: 1, left: 0, right: 0 });
  assert.deepEqual(buildDataValidationRule(document, 0, { type: 'whole', operator: 'between', first: '2', second: '=A1+3' }), { type: 'whole', operator: 'between', ranges: [], values: [2, null], formulas: [null, 'A1+3'], errorStyle: 'stop', allowBlank: true, showDropDown: false, showInputMessage: false, showErrorMessage: true, errorTitle: '', error: '', promptTitle: '', prompt: '' });
  assert.equal(buildDataValidationRule(document, 0, { type: 'date', operator: 'greaterThan', first: '1900-02-29' }).values[0], 60);
  assert.equal(buildDataValidationRule(document, 0, { type: 'time', operator: 'lessThanOrEqual', first: '12:30' }).values[0], 0.5208333333333334);
  assert.throws(() => buildDataValidationRule(document, 0, { type: 'whole', operator: 'equal', first: '1.5' }), /숫자 형식/);
  document.state.date1904 = true;
  assert.throws(() => buildDataValidationRule(document, 0, { type: 'date', operator: 'equal', first: '1900-02-29' }), /1900-02-29/);
  document.dispose();
});

test('데이터 유효성 설정은 겹친 규칙 범위만 교체하고 저장·실행 취소·다시 실행을 지원한다', () => {
  const original = '<dataValidations count="1"><dataValidation type="list" sqref="A1:C3"><formula1>"old,legacy"</formula1></dataValidation></dataValidations>';
  const bytes = minimalXlsx('', '', original), document = new JinSheetDocument(uri('/tmp/data-validation-overlap.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  assert.equal(document.writable, true, document.state.readOnlyReason);
  let operation; document.onDidChange(change => { operation = change; });
  const replacement = buildDataValidationRule(document, 0, { type: 'list', source: 'fresh,done' });
  document.applyDataValidation(0, { top: 1, bottom: 1, left: 1, right: 1 }, replacement);
  const validationOperation = operation;
  assert.equal(document.state.sheets[0].dataValidations.length, 2);
  assert.equal(document.state.sheets[0].dataValidations[0].ranges.length, 4);
  assert.throws(() => document.applyCellEdit(1, 1, 'old'), /허용 목록/);
  validationOperation.undo();
  assert.equal(document.state.sheets[0].dataValidations.length, 1);
  assert.throws(() => document.applyCellEdit(1, 1, 'fresh'), /허용 목록/);
  validationOperation.redo();
  assert.equal(document.state.sheets[0].dataValidations.length, 2);
  document.applyCellEdit(1, 1, 'fresh');
  document.applyCellEdit(0, 0, 'legacy');
  const saved = analyze(document.createOutput());
  assert.equal(saved.writable, true, saved.readOnlyReason);
  assert.deepEqual(saved.sheets[0].dataValidations.map(rule => rule.ranges.length), [4, 1]);
  document.applyDataValidation(0, { top: 1, bottom: 1, left: 1, right: 1 }, null);
  assert.equal(document.state.sheets[0].dataValidations.length, 1);
  assert.equal(document.state.sheets[0].dataValidations[0].ranges.length, 4);
  document.dispose();
});

test('조건부 서식 지우기는 선택 셀만 제외하고 저장·실행 취소·다시 실행을 지원한다', () => {
  const rule = '<conditionalFormatting sqref="A1:C3"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF4472C4"/></dataBar></cfRule></conditionalFormatting>';
  const bytes = minimalXlsx('', '', rule), document = new JinSheetDocument(uri('/tmp/conditional-format-clear.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  assert.equal(document.writable, true, document.state.readOnlyReason);
  let operation; document.onDidChange(change => { operation = change; });
  document.clearConditionalFormats(0, { top: 1, bottom: 1, left: 1, right: 1 });
  assert.equal(document.state.sheets[0].conditionalFormats[0].ranges.length, 4);
  operation.undo();
  assert.deepEqual(document.state.sheets[0].conditionalFormats[0].ranges, [{ top: 0, bottom: 2, left: 0, right: 2 }]);
  operation.redo();
  assert.equal(document.state.sheets[0].conditionalFormats[0].ranges.length, 4);
  const saved = analyze(document.createOutput());
  assert.equal(saved.writable, true, saved.readOnlyReason);
  assert.equal(saved.sheets[0].conditionalFormats[0].ranges.length, 4);
  document.dispose();
});

test('조건부 서식 규칙 관리는 선택한 규칙을 저장·실행 취소·다시 실행한다', () => {
  const rules = '<conditionalFormatting sqref="A1:C3"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF4472C4"/></dataBar></cfRule><cfRule type="colorScale" priority="2"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FF00FF00"/></colorScale></cfRule></conditionalFormatting>';
  const bytes = minimalXlsx('', '', rules), document = new JinSheetDocument(uri('/tmp/conditional-format-rule-delete.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  assert.equal(document.writable, true, document.state.readOnlyReason);
  let operation; document.onDidChange(change => { operation = change; });
  document.deleteConditionalFormatRule(0, 1);
  assert.deepEqual(document.state.sheets[0].conditionalFormats.map(rule => rule.priority), [2]);
  operation.undo();
  assert.deepEqual(document.state.sheets[0].conditionalFormats.map(rule => rule.priority), [1, 2]);
  operation.redo();
  assert.deepEqual(document.state.sheets[0].conditionalFormats.map(rule => rule.priority), [2]);
  const saved = analyze(document.createOutput());
  assert.equal(saved.writable, true, saved.readOnlyReason);
  assert.deepEqual(saved.sheets[0].conditionalFormats.map(rule => rule.priority), [2]);
  document.dispose();
});

test('조건부 서식 우선순위 이동은 저장·실행 취소·다시 실행을 지원한다', () => {
  const rules = '<conditionalFormatting sqref="A1:C3"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF4472C4"/></dataBar></cfRule><cfRule type="colorScale" priority="2"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FF00FF00"/></colorScale></cfRule></conditionalFormatting>';
  const bytes = minimalXlsx('', '', rules), document = new JinSheetDocument(uri('/tmp/conditional-format-rule-priority.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  assert.equal(document.writable, true, document.state.readOnlyReason);
  let operation; document.onDidChange(change => { operation = change; });
  assert.equal(document.moveConditionalFormatRule(0, 2, 'up'), true);
  assert.equal(document.state.sheets[0].conditionalFormats.find(rule => rule.type === 'colorScale').priority, 1);
  assert.equal(document.state.sheets[0].conditionalFormats.find(rule => rule.type === 'dataBar').priority, 2);
  assert.equal(operation.label, 'Raise conditional-formatting rule priority');
  operation.undo();
  assert.equal(document.state.sheets[0].conditionalFormats.find(rule => rule.type === 'colorScale').priority, 2);
  operation.redo();
  assert.equal(document.state.sheets[0].conditionalFormats.find(rule => rule.type === 'colorScale').priority, 1);
  const saved = analyze(document.createOutput());
  assert.equal(saved.writable, true, saved.readOnlyReason);
  assert.equal(saved.sheets[0].conditionalFormats.find(rule => rule.type === 'colorScale').priority, 1);
  document.dispose();
});

test('조건부 서식 규칙 편집은 색상 변경을 저장·실행 취소·다시 실행한다', () => {
  const rules = '<conditionalFormatting sqref="A1:C3"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FFFFFF00"/><color rgb="FF00FF00"/></colorScale></cfRule></conditionalFormatting>';
  const bytes = minimalXlsx('', '', rules), document = new JinSheetDocument(uri('/tmp/conditional-format-rule-edit.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  assert.equal(document.writable, true, document.state.readOnlyReason);
  let operation; document.onDidChange(change => { operation = change; });
  assert.equal(document.editConditionalFormatRule(0, 1, { kind: 'colors', colors: ['#112233', '#AABBCC', '#334455'] }), true);
  assert.deepEqual(document.state.sheets[0].conditionalFormats[0].colorScale.colors, ['#112233', '#AABBCC', '#334455']);
  assert.equal(operation.label, 'Edit conditional-formatting rule');
  operation.undo();
  assert.deepEqual(document.state.sheets[0].conditionalFormats[0].colorScale.colors, ['#FF0000', '#FFFF00', '#00FF00']);
  operation.redo();
  assert.deepEqual(document.state.sheets[0].conditionalFormats[0].colorScale.colors, ['#112233', '#AABBCC', '#334455']);
  assert.equal(document.editConditionalFormatRule(0, 1, { kind: 'threshold', stop: { type: 'num', value: '25' } }), true);
  assert.deepEqual(document.state.sheets[0].conditionalFormats[0].colorScale.stops[1], { type: 'num', value: 25 });
  operation.undo();
  assert.deepEqual(document.state.sheets[0].conditionalFormats[0].colorScale.stops[1], { type: 'percentile', value: 50 });
  operation.redo();
  assert.deepEqual(document.state.sheets[0].conditionalFormats[0].colorScale.stops[1], { type: 'num', value: 25 });
  const saved = analyze(document.createOutput());
  assert.equal(saved.writable, true, saved.readOnlyReason);
  assert.deepEqual(saved.sheets[0].conditionalFormats[0].colorScale.colors, ['#112233', '#AABBCC', '#334455']);
  assert.deepEqual(saved.sheets[0].conditionalFormats[0].colorScale.stops[1], { type: 'num', value: 25 });
  document.dispose();
});

test('데이터 막대 표시 설정은 저장·실행 취소·다시 실행 후 복원된다', () => {
  const rules = '<conditionalFormatting sqref="A1:A3"><cfRule type="dataBar" priority="1"><dataBar minLength="5" maxLength="90" showValue="0" gradient="0"><cfvo type="min"/><cfvo type="max"/><color rgb="FF4472C4"/></dataBar></cfRule></conditionalFormatting>';
  const bytes = minimalXlsx('', '', rules), document = new JinSheetDocument(uri('/tmp/data-bar-display-edit.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  assert.equal(document.writable, true, document.state.readOnlyReason);
  let operation; document.onDidChange(change => { operation = change; });
  assert.equal(document.editConditionalFormatRule(0, 1, { kind: 'data-bar-display', dataBar: { minLength: 15, maxLength: 75, showValue: true, gradient: true } }), true);
  assert.deepEqual(document.state.sheets[0].conditionalFormats[0].dataBar, { color: '#4472C4', minLength: 15, maxLength: 75, showValue: true, gradient: true });
  operation.undo();
  assert.deepEqual(document.state.sheets[0].conditionalFormats[0].dataBar, { color: '#4472C4', minLength: 5, maxLength: 90, showValue: false, gradient: false });
  operation.redo();
  assert.deepEqual(document.state.sheets[0].conditionalFormats[0].dataBar, { color: '#4472C4', minLength: 15, maxLength: 75, showValue: true, gradient: true });
  const saved = analyze(document.createOutput());
  assert.equal(saved.writable, true, saved.readOnlyReason);
  assert.deepEqual(saved.sheets[0].conditionalFormats[0].dataBar, { color: '#4472C4', minLength: 15, maxLength: 75, showValue: true, gradient: true });
  document.dispose();
});

test('아이콘 집합 표시·기준 편집은 저장·실행 취소·다시 실행 후 복원된다', () => {
  const rules = '<conditionalFormatting sqref="A1:A3"><cfRule type="iconSet" priority="1"><iconSet iconSet="3TrafficLights1" showValue="0"><cfvo type="min"/><cfvo type="num" val="4" gte="0"/><cfvo type="max"/></iconSet></cfRule></conditionalFormatting>';
  const bytes = minimalXlsx('', '', rules), document = new JinSheetDocument(uri('/tmp/icon-set-edit.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  assert.equal(document.writable, true, document.state.readOnlyReason);
  let operation; document.onDidChange(change => { operation = change; });
  assert.equal(document.editConditionalFormatRule(0, 1, { kind: 'icon-set-display', iconSet: { showValue: true, reverse: true } }), true);
  assert.deepEqual(document.state.sheets[0].conditionalFormats[0].iconSet, { name: '3TrafficLights1', showValue: true, reverse: true, thresholds: [{ type: 'min', value: null, gte: true }, { type: 'num', value: 4, gte: false }, { type: 'max', value: null, gte: true }] });
  operation.undo();
  assert.equal(document.state.sheets[0].conditionalFormats[0].iconSet.showValue, false);
  assert.equal(document.state.sheets[0].conditionalFormats[0].iconSet.reverse, false);
  operation.redo();
  assert.equal(document.state.sheets[0].conditionalFormats[0].iconSet.showValue, true);
  assert.equal(document.state.sheets[0].conditionalFormats[0].iconSet.reverse, true);
  assert.equal(document.editConditionalFormatRule(0, 1, { kind: 'icon-set-threshold', index: 1, stop: { type: 'percentile', value: '65', gte: true } }), true);
  assert.deepEqual(document.state.sheets[0].conditionalFormats[0].iconSet.thresholds[1], { type: 'percentile', value: 65, gte: true });
  operation.undo();
  assert.deepEqual(document.state.sheets[0].conditionalFormats[0].iconSet.thresholds[1], { type: 'num', value: 4, gte: false });
  operation.redo();
  assert.deepEqual(document.state.sheets[0].conditionalFormats[0].iconSet.thresholds[1], { type: 'percentile', value: 65, gte: true });
  const saved = analyze(document.createOutput());
  assert.equal(saved.writable, true, saved.readOnlyReason);
  assert.deepEqual(saved.sheets[0].conditionalFormats[0].iconSet, document.state.sheets[0].conditionalFormats[0].iconSet);
  document.dispose();
});

test('아이콘 집합 조건부 서식 생성은 저장·실행 취소·다시 실행을 지원한다', () => {
  const bytes = minimalXlsx(), document = new JinSheetDocument(uri('/tmp/icon-set-create.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  assert.equal(document.writable, true, document.state.readOnlyReason);
  let operation; document.onDidChange(change => { operation = change; });
  assert.equal(document.applyIconSet(0, { top: 0, bottom: 2, left: 0, right: 0 }, '5Quarters'), true);
  const rule = document.state.sheets[0].conditionalFormats[0];
  assert.equal(rule.type, 'iconSet');
  assert.deepEqual(rule.ranges, [{ top: 0, bottom: 2, left: 0, right: 0 }]);
  assert.deepEqual(rule.iconSet, { name: '5Quarters', showValue: true, reverse: false, thresholds: [0, 20, 40, 60, 80].map(value => ({ type: 'percent', value, gte: true })) });
  assert.equal(operation.label, 'Apply icon set');
  operation.undo();
  assert.equal(document.state.sheets[0].conditionalFormats.length, 0);
  operation.redo();
  const saved = analyze(document.createOutput());
  assert.equal(saved.writable, true, saved.readOnlyReason);
  assert.deepEqual(saved.sheets[0].conditionalFormats[0].iconSet, rule.iconSet);
  document.dispose();
});

test('차등 글꼴색 편집은 실행 취소·다시 실행에서 공유 스타일 참조를 복원한다', () => {
  const files = Object.fromEntries(Object.entries(unzipSync(minimalXlsx())).map(([name, value]) => [name, new TextDecoder().decode(value)]));
  files['[Content_Types].xml'] = files['[Content_Types].xml'].replace('</Types>', '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
  files['xl/_rels/workbook.xml.rels'] = files['xl/_rels/workbook.xml.rels'].replace('</Relationships>', '<Relationship Id="style1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  files['xl/styles.xml'] = '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs><dxfs count="1"><dxf><font><b/><color rgb="FF008000"/></font></dxf></dxfs></styleSheet>';
  files['xl/worksheets/sheet1.xml'] = files['xl/worksheets/sheet1.xml'].replace('</worksheet>', '<conditionalFormatting sqref="A1"><cfRule type="cellIs" operator="greaterThan" dxfId="0" priority="1"><formula>10</formula></cfRule></conditionalFormatting></worksheet>');
  const bytes = zipSync(Object.fromEntries(Object.entries(files).map(([name, value]) => [name, new TextEncoder().encode(value)])));
  const document = new JinSheetDocument(uri('/tmp/conditional-format-differential-color.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  assert.equal(document.writable, true, document.state.readOnlyReason);
  let operation; document.onDidChange(change => { operation = change; });
  assert.equal(document.editConditionalFormatRule(0, 1, { kind: 'differential-color', property: 'fontColor', color: '#123456' }), true);
  assert.deepEqual(document.state.sheets[0].conditionalFormats[0].style, { bold: true, color: '#123456' });
  operation.undo();
  assert.deepEqual(document.state.sheets[0].conditionalFormats[0].style, { bold: true, color: '#008000' });
  assert.equal(document.state.sheets[0].conditionalFormats[0].dxfId, 0);
  operation.redo();
  assert.deepEqual(document.state.sheets[0].conditionalFormats[0].style, { bold: true, color: '#123456' });
  const saved = analyze(document.createOutput());
  assert.equal(saved.writable, true, saved.readOnlyReason);
  assert.deepEqual(saved.sheets[0].conditionalFormats[0].style, { bold: true, color: '#123456' });
  document.dispose();
});

test('데이터 유효성 컨테이너가 없는 XLSX에도 규칙을 안전하게 추가한다', () => {
  const bytes = minimalXlsx(), document = new JinSheetDocument(uri('/tmp/data-validation-first.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  const rule = buildDataValidationRule(document, 0, { type: 'decimal', operator: 'greaterThanOrEqual', first: '0.5' });
  document.applyDataValidation(0, { top: 0, bottom: 2, left: 0, right: 0 }, rule);
  const saved = analyze(document.createOutput()), xml = saved.sheets[0].xml;
  assert.equal(saved.writable, true, saved.readOnlyReason);
  assert.match(xml, /<dataValidations count="1"><dataValidation type="decimal" errorStyle="stop" operator="greaterThanOrEqual" allowBlank="1" showInputMessage="0" showErrorMessage="1" sqref="A1:A3"><formula1>0.5<\/formula1><\/dataValidation><\/dataValidations>/);
  assert.ok(xml.indexOf('</sheetData>') < xml.indexOf('<dataValidations'));
  document.dispose();
});

test('데이터 유효성 버튼은 XLSX 선택 범위를 메시지로 전달한다', () => {
  const html = renderHtml({ kind: 'xlsx' });
  assert.match(html, /id="dataValidation"/);
  assert.match(html, /type:'manage-data-validation',range,sheetIndex:model\.sheetIndex/);
  assert.doesNotMatch(renderHtml({ kind: 'csv' }), /id="dataValidation"/);
});

test('조건부 서식 규칙 관리 버튼은 영어·한국어로 표시하고 XLSX에서만 동작한다', () => {
  const english = renderHtml({ kind: 'xlsx' }, undefined, undefined, 'en');
  const korean = renderHtml({ kind: 'xlsx' }, undefined, undefined, 'ko-KR');
  assert.match(english, /id="manageConditionalFormats"[^>]*>Manage rules/);
  assert.match(english, /type:'manage-conditional-format-rules',sheetIndex:model\.sheetIndex/);
  assert.match(korean, /id="manageConditionalFormats"[^>]*>규칙 관리/);
  assert.doesNotMatch(renderHtml({ kind: 'csv' }), /id="manageConditionalFormats"/);
});

test('XLSX 병합 셀 범위를 뷰 모델에 전달하고 왼쪽 위 셀 외의 수정을 거부한다', () => {
  const files = Object.fromEntries(Object.entries(unzipSync(minimalXlsx())).map(([name, value]) => [name, new TextDecoder().decode(value)]));
  files['xl/worksheets/sheet1.xml'] = files['xl/worksheets/sheet1.xml'].replace('</worksheet>', '<mergeCells count="1"><mergeCell ref="A1:B2"/></mergeCells></worksheet>');
  const bytes = zipSync(Object.fromEntries(Object.entries(files).map(([name, value]) => [name, new TextEncoder().encode(value)])));
  const document = new JinSheetDocument(uri('/tmp/merged-cells.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  assert.equal(document.writable, true, document.state.readOnlyReason);
  assert.deepEqual(viewModel(document).mergedCells, [{ top: 0, bottom: 1, left: 0, right: 1, ref: 'A1:B2' }]);
  const range = getRange(document, { requestId: 1, rowStart: 0, columnStart: 0, rowCount: 2, columnCount: 2 });
  assert.deepEqual(range.mergedCells, viewModel(document).mergedCells);
  document.applyCellEdit(0, 0, 'merged value');
  assert.equal(document.getCell(0, 0).value, 'merged value');
  assert.throws(() => document.applyCellEdit(1, 1, 'hidden value'), /왼쪽 위 셀/);
});

test('XLSX 데이터 유효성 사용자 오류 메시지와 오류 표시 비활성 설정을 적용한다', () => {
  const validation = '<dataValidations count="2"><dataValidation type="list" errorTitle="Invalid choice" error="Choose Yes or No" promptTitle="Answer" prompt="Select Yes or No" sqref="A1"><formula1>"Yes,No"</formula1></dataValidation><dataValidation type="list" showErrorMessage="0" sqref="B1"><formula1>"Yes,No"</formula1></dataValidation></dataValidations>';
  const bytes = minimalXlsx('', '', validation), document = new JinSheetDocument(uri('/tmp/data-validation-messages.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  assert.equal(document.writable, true, document.state.readOnlyReason);
  assert.throws(() => document.applyCellEdit(0, 0, 'Maybe'), error => error.message === 'Invalid choice: Choose Yes or No' && error.userMessage === true);
  document.applyCellEdit(0, 1, 'Maybe');
  const cell = getRange(document, { requestId: 1, rowStart: 0, columnStart: 0, rowCount: 1, columnCount: 2 }).rows[0];
  assert.deepEqual([cell[0].validation.promptTitle, cell[0].validation.prompt, cell[0].validation.errorTitle], ['Answer', 'Select Yes or No', 'Invalid choice']);
  assert.equal(cell[1].validation.showErrorMessage, false);
});

test('XLSX 경고·정보형 데이터 유효성은 확인 후 입력을 적용하고 취소 가능한 대기 상태를 유지한다', () => {
  const validation = '<dataValidations count="2"><dataValidation type="list" errorStyle="warning" errorTitle="Check entry" error="This value is outside the list" sqref="A1:B1"><formula1>"Yes,No"</formula1></dataValidation><dataValidation type="list" errorStyle="information" sqref="C1"><formula1>"Yes,No"</formula1></dataValidation></dataValidations>';
  const bytes = minimalXlsx('', '', validation), document = new JinSheetDocument(uri('/tmp/data-validation-alerts.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  assert.equal(document.writable, true, document.state.readOnlyReason);
  assert.equal(document.state.sheets[0].dataValidations[0].errorStyle, 'warning');
  assert.equal(document.state.sheets[0].dataValidations[1].errorStyle, 'information');
  let warning;
  assert.throws(() => document.applyCellEdits([{ row: 0, column: 0, value: 'Maybe' }, { row: 0, column: 1, value: 'Perhaps' }]), error => {
    warning = error;
    return error.validationPrompt?.style === 'warning' && error.message === 'Check entry: This value is outside the list';
  });
  assert.equal(document.getCell(0, 0).value, '');
  assert.throws(() => warning.retry(), error => {
    warning = error;
    return error.validationPrompt?.style === 'warning';
  }, 'each invalid cell in a paste asks before it is accepted');
  warning.retry();
  assert.deepEqual([document.getCell(0, 0).value, document.getCell(0, 1).value], ['Maybe', 'Perhaps']);
  let information;
  assert.throws(() => document.applyCellEdit(0, 2, 'Other'), error => {
    information = error;
    return error.validationPrompt?.style === 'information';
  });
  information.retry();
  assert.equal(document.getCell(0, 2).value, 'Other');
});

test('XLSX 셀 범위 목록 유효성은 원본 셀 값에 따라 선택지를 갱신한다', () => {
  const rows = '<row r="1"><c r="A1" t="inlineStr"><is><t>Open</t></is></c><c r="B1" t="inlineStr"><is><t>Red</t></is></c></row><row r="2"><c r="B2" t="inlineStr"><is><t>Blue</t></is></c></row><row r="3"><c r="B3" t="inlineStr"><is><t>Green</t></is></c></row>';
  const validation = '<dataValidations count="1"><dataValidation type="list" allowBlank="1" sqref="A1"><formula1>=$B$1:$B$3</formula1></dataValidation></dataValidations>';
  const bytes = minimalXlsx('', rows, validation), document = new JinSheetDocument(uri('/tmp/range-validation.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  assert.equal(document.writable, true, document.state.readOnlyReason);
  assert.deepEqual(getRange(document, { requestId: 1, rowStart: 0, columnStart: 0, rowCount: 1, columnCount: 1 }).rows[0][0].validation.values, ['Red', 'Blue', 'Green']);
  assert.throws(() => document.applyCellEdit(0, 0, 'Pending'), /허용 목록/);
  document.applyCellEdit(2, 1, 'Yellow');
  assert.deepEqual(getRange(document, { requestId: 2, rowStart: 0, columnStart: 0, rowCount: 1, columnCount: 1 }).rows[0][0].validation.values, ['Red', 'Blue', 'Yellow']);
  document.applyCellEdit(0, 0, 'Yellow');
  assert.equal(document.getCell(0, 0).value, 'Yellow');
});

test('XLSX 다른 시트 범위 목록은 참조 시트 값을 제안하고 변경을 반영한다', () => {
  const document = new JinSheetDocument(uri('/tmp/cross-sheet-validation.xlsx'), 'xlsx', analyze(twoSheetValidationXlsx()), Buffer.alloc(0), true, '');
  assert.equal(document.writable, true, document.state.readOnlyReason);
  const getChoices = () => getRange(document, { requestId: 1, rowStart: 0, columnStart: 1, rowCount: 1, columnCount: 1, sheetIndex: 0 }).rows[0][0].validation.values;
  assert.deepEqual(getChoices(), ['Red', 'Blue']);
  document.applyCellEdit(0, 0, 'Green', 1);
  assert.deepEqual(getChoices(), ['Green', 'Blue']);
  document.applyStructureChange({ axis: 'row', action: 'insert', index: 0, sheetIndex: 1 });
  assert.deepEqual(getChoices(), ['Green', 'Blue']);
  assert.match(document.state.sheets[0].xml, /<formula1>Lists!\$A\$2:\$A\$3<\/formula1>/);
});

test('XLSX INDIRECT 목록은 참조 셀의 범위와 이름을 따라 선택지를 갱신한다', () => {
  const files = Object.fromEntries(Object.entries(unzipSync(twoSheetValidationXlsx(true))).map(([name, value]) => [name, new TextDecoder().decode(value)]));
  files['xl/worksheets/sheet1.xml'] = files['xl/worksheets/sheet1.xml'].replace('<formula1>ColorList</formula1>', '<formula1>INDIRECT(C1)</formula1>').replace('</row></sheetData>', '<c r="C1" t="inlineStr"><is><t>ColorList</t></is></c></row></sheetData>');
  const bytes = zipSync(Object.fromEntries(Object.entries(files).map(([name, value]) => [name, new TextEncoder().encode(value)])));
  const document = new JinSheetDocument(uri('/tmp/indirect-list-validation.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  assert.equal(document.writable, true, document.state.readOnlyReason);
  assert.deepEqual(getRange(document, { requestId: 1, rowStart: 0, columnStart: 1, rowCount: 1, columnCount: 1 }).rows[0][0].validation.values, ['Red', 'Blue']);
  assert.throws(() => document.applyCellEdit(0, 1, 'Green'), /허용 목록/);
  document.applyCellEdit(0, 2, 'ColorList');
  assert.deepEqual(getRange(document, { requestId: 2, rowStart: 0, columnStart: 1, rowCount: 1, columnCount: 1 }).rows[0][0].validation.values, ['Red', 'Blue']);
});

test('XLSX 이름 정의 목록 원본은 시트 범위를 찾아 편집 선택지로 쓴다', () => {
  const document = new JinSheetDocument(uri('/tmp/defined-name-validation.xlsx'), 'xlsx', analyze(twoSheetValidationXlsx(true)), Buffer.alloc(0), true, '');
  assert.equal(document.writable, true, document.state.readOnlyReason);
  assert.deepEqual(getRange(document, { requestId: 1, rowStart: 0, columnStart: 1, rowCount: 1, columnCount: 1, sheetIndex: 0 }).rows[0][0].validation.values, ['Red', 'Blue']);
  document.applyStructureChange({ axis: 'row', action: 'insert', index: 0, sheetIndex: 1 });
  assert.deepEqual(getRange(document, { requestId: 2, rowStart: 0, columnStart: 1, rowCount: 1, columnCount: 1, sheetIndex: 0 }).rows[0][0].validation.values, ['Red', 'Blue']);
  const workbook = new TextDecoder().decode(document.state.files['xl/workbook.xml']);
  assert.match(workbook, /name="ColorList">Lists!\$A\$2:\$A\$3<\/definedName>/);
});

test('XLSX 숫자·날짜·시간·텍스트 길이 유효성은 편집 값을 제한한다', () => {
  const rows = '<row r="1"><c r="A1" t="inlineStr"><is><t>5</t></is></c><c r="B1" t="inlineStr"><is><t>1.5</t></is></c><c r="C1" t="inlineStr"><is><t>2024-01-01</t></is></c><c r="D1" t="inlineStr"><is><t>12:00</t></is></c><c r="E1" t="inlineStr"><is><t>cat</t></is></c></row>';
  const validation = '<dataValidations count="5"><dataValidation type="whole" operator="between" allowBlank="1" sqref="A1"><formula1>1</formula1><formula2>10</formula2></dataValidation><dataValidation type="decimal" operator="greaterThan" sqref="B1"><formula1>0</formula1></dataValidation><dataValidation type="date" operator="between" sqref="C1"><formula1>45000</formula1><formula2>46000</formula2></dataValidation><dataValidation type="time" operator="between" sqref="D1"><formula1>0.25</formula1><formula2>0.75</formula2></dataValidation><dataValidation type="textLength" operator="equal" sqref="E1"><formula1>3</formula1></dataValidation></dataValidations>';
  const bytes = minimalXlsx('', rows, validation), document = new JinSheetDocument(uri('/tmp/numeric-validation.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  assert.equal(document.writable, true, document.state.readOnlyReason);
  document.applyCellEdit(0, 0, '10');
  assert.throws(() => document.applyCellEdit(0, 0, '11'), /데이터 유효성/);
  assert.throws(() => document.applyCellEdit(0, 0, '1.5'), /데이터 유효성/);
  document.applyCellEdit(0, 1, '0.1');
  assert.throws(() => document.applyCellEdit(0, 1, '0'), /데이터 유효성/);
  document.applyCellEdit(0, 2, '2024-01-01');
  document.applyCellEdit(0, 2, '45000');
  assert.throws(() => document.applyCellEdit(0, 2, '46001'), /데이터 유효성/);
  assert.throws(() => document.applyCellEdit(0, 2, '2020-01-01'), /데이터 유효성/);
  document.applyCellEdit(0, 3, '12:00');
  assert.throws(() => document.applyCellEdit(0, 3, '05:59'), /데이터 유효성/);
  document.applyCellEdit(0, 4, 'dog');
  assert.throws(() => document.applyCellEdit(0, 4, 'horses'), /데이터 유효성/);
});

test('XLSX 수식 기준 데이터 유효성은 각 적용 행에서 상대 참조를 평가한다', () => {
  const rows = '<row r="1"><c r="A1"><v>0</v></c><c r="B1"><v>3</v></c></row><row r="2"><c r="A2"><v>0</v></c><c r="B2"><v>7</v></c></row>';
  const validation = '<dataValidations count="1"><dataValidation type="whole" operator="between" sqref="A1:A2"><formula1>B1</formula1><formula2>B1+2</formula2></dataValidation></dataValidations>';
  const document = new JinSheetDocument(uri('/tmp/formula-validation.xlsx'), 'xlsx', analyze(minimalXlsx('', rows, validation)), Buffer.alloc(0), true, '');
  assert.equal(document.writable, true, document.state.readOnlyReason);
  document.applyCellEdit(0, 0, '4');
  assert.throws(() => document.applyCellEdit(0, 0, '6'), /데이터 유효성/);
  document.applyCellEdit(1, 0, '8');
  assert.throws(() => document.applyCellEdit(1, 0, '10'), /데이터 유효성/);
  document.applyCellEdit(0, 1, '5');
  document.applyCellEdit(0, 0, '6');
});

test('XLSX 단순 AutoFilter 값 목록은 편집 후 보이는 행을 다시 계산하고 XML을 보존한다', async t => {
  const rows = [
    '<row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Value</t></is></c></row>',
    '<row r="2" hidden="1"><c r="A2" t="inlineStr"><is><t>apple</t></is></c><c r="B2"><v>1</v></c></row>',
    '<row r="3"><c r="A3" t="inlineStr"><is><t>pear</t></is></c><c r="B3"><v>2</v></c></row>',
    '<row r="4"><c r="A4" t="inlineStr"><is><t>apple</t></is></c><c r="B4"><v>3</v></c></row>'
  ].join('');
  const autoFilter = '<autoFilter ref="A1:B4"><filterColumn colId="0"><filters><filter val="apple"/></filters></filterColumn></autoFilter>';
  const bytes = minimalXlsx('', rows, autoFilter);
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-autofilter-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'filtered.xlsx');
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  assert.equal(document.writable, true);
  assert.deepEqual(document.state.sheets[0].hiddenRows, [1]);
  assert.deepEqual(viewModel(document).hiddenRows, [1, 2]);
  let change;
  document.onDidChange(event => { change = event; });
  document.applyCellEdit(1, 0, 'pear', 0);
  assert.deepEqual(viewModel(document).hiddenRows, [1, 2]);
  change.undo();
  assert.deepEqual(viewModel(document).hiddenRows, [1, 2]);
  change.redo();
  await document.save();
  const reopenedBytes = await fs.readFile(file), reopened = analyze(reopenedBytes);
  assert.equal(reopened.writable, true);
  assert.deepEqual(viewModel(new JinSheetDocument(uri(file), 'xlsx', reopened, reopenedBytes, true, '')).hiddenRows, [1, 2]);
  assert.match(new TextDecoder().decode(reopened.files['xl/worksheets/sheet1.xml']), /<autoFilter ref="A1:B4"><filterColumn colId="0"><filters><filter val="apple"\/><\/filters><\/filterColumn><\/autoFilter>/);
  const inserted = require('../src/xlsx').transformStructure(reopened, { sheetIndex: 0, axis: 'row', action: 'insert', index: 1 });
  assert.equal(inserted.model.sheets[0].autoFilter.bottom, 4);
  assert.match(inserted.model.sheets[0].xml, /<autoFilter ref="A1:B5"/);
  document.applyStructureChange({ axis: 'row', action: 'insert', index: 1, sheetIndex: 0 });
  const structureChange = change;
  assert.equal(document.state.sheets[0].autoFilter.bottom, 4);
  await document.save();
  structureChange.undo();
  assert.equal(document.dirty, true);
  assert.equal(document.state.sheets[0].autoFilter.bottom, 3);
  await document.save();
  structureChange.redo();
  await document.save();
  assert.equal(analyze(await fs.readFile(file)).sheets[0].autoFilter.bottom, 4);
  document.dispose();
});

test('AutoFilter 변경은 문서 undo·redo·저장 후에도 조건과 보이는 행을 일치시킨다', async t => {
  const rows = '<row r="1"><c r="A1" t="inlineStr"><is><t>Kind</t></is></c></row>' +
    '<row r="2"><c r="A2" t="inlineStr"><is><t>apple</t></is></c></row>' +
    '<row r="3"><c r="A3" t="inlineStr"><is><t>pear</t></is></c></row>' +
    '<row r="4"><c r="A4" t="inlineStr"><is><t>apple</t></is></c></row>';
  const bytes = minimalXlsx('', rows, '<autoFilter ref="A1:A4"><filterColumn colId="0"><filters><filter val="apple"/></filters></filterColumn></autoFilter>');
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-filter-undo-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'filter.xlsx');
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  let operation;
  document.onDidChange(change => { operation = change; });
  const changed = updateAutoFilterColumn(document, 0, 0, ['pear'], false);
  document.applyAutoFilter(0, changed);
  assert.deepEqual(viewModel(document).hiddenRows, [1, 3]);
  operation.undo();
  assert.deepEqual(viewModel(document).hiddenRows, [2]);
  operation.redo();
  await document.save();
  assert.equal(document.dirty, false);
  const savedBytes = await fs.readFile(file), saved = analyze(savedBytes);
  assert.deepEqual(saved.sheets[0].autoFilter.columns, [{ column: 0, values: ['pear'], blanks: false }]);
  assert.deepEqual(viewModel(new JinSheetDocument(uri(file), 'xlsx', saved, savedBytes, true, '')).hiddenRows, [1, 3]);
  operation.undo();
  assert.equal(document.dirty, true, 'undo after save returns the document to the previous filter condition');
  await document.save();
  assert.deepEqual(analyze(await fs.readFile(file)).sheets[0].autoFilter.columns, [{ column: 0, values: ['apple'], blanks: false }]);
  document.dispose();
});

test('일반 XLSX에 새 AutoFilter를 만들면 undo·redo와 저장에 반영한다', async t => {
  const rows = '<row r="1"><c r="A1" t="inlineStr"><is><t>Color</t></is></c></row>' +
    '<row r="2"><c r="A2" t="inlineStr"><is><t>red</t></is></c></row>' +
    '<row r="3"><c r="A3" t="inlineStr"><is><t>blue</t></is></c></row>';
  const bytes = minimalXlsx('', rows), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-filter-create-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'plain.xlsx');
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  const filter = { top: 0, bottom: 2, left: 0, right: 0, columns: [{ column: 0, values: ['red'], blanks: false }] };
  let operation;
  document.onDidChange(change => { operation = change; });
  document.applyAutoFilter(0, filter);
  assert.deepEqual(viewModel(document).hiddenRows, [2]);
  operation.undo();
  assert.equal(viewModel(document).hiddenRows, undefined);
  operation.redo();
  await document.save();
  assert.deepEqual(analyze(await fs.readFile(file)).sheets[0].autoFilter, filter);
  document.dispose();
});

test('수동 행 숨김은 undo·redo와 XLSX 저장·다시 열기에 반영한다', async t => {
  const rows = ['<row r="1"><c r="A1" t="inlineStr"><is><t>Header</t></is></c></row>', '<row r="2"><c r="A2" t="inlineStr"><is><t>value</t></is></c></row>'].join('');
  const bytes = minimalXlsx('', rows), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-hidden-row-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'rows.xlsx');
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  let operation;
  document.onDidChange(change => { operation = change; });
  document.applyHiddenRows(0, [1], true);
  assert.deepEqual(viewModel(document).hiddenRows, [1]);
  operation.undo();
  assert.equal(viewModel(document).hiddenRows, undefined);
  operation.redo();
  await document.save();
  assert.deepEqual(analyze(await fs.readFile(file)).sheets[0].hiddenRows, [1]);
  document.applyHiddenRows(0, [1], false);
  await document.save();
  assert.deepEqual(analyze(await fs.readFile(file)).sheets[0].hiddenRows, []);
  document.dispose();
});

test('수동 열 숨김은 undo·redo와 XLSX 저장·다시 열기에 반영한다', async t => {
  const rows = '<row r="1"><c r="A1" t="inlineStr"><is><t>Header</t></is></c><c r="B1" t="inlineStr"><is><t>value</t></is></c></row>';
  const files = require('../vendor/node_modules/fflate').unzipSync(minimalXlsx('', rows));
  const sheetPath = 'xl/worksheets/sheet1.xml';
  files[sheetPath] = new TextEncoder().encode(new TextDecoder().decode(files[sheetPath]).replace('<sheetData>', '<cols><col min="1" max="2" width="12" customWidth="1"/></cols><sheetData>'));
  const bytes = require('../vendor/node_modules/fflate').zipSync(files), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-hidden-column-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'columns.xlsx');
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  let operation;
  document.onDidChange(change => { operation = change; });
  document.applyHiddenColumns(0, [1], true);
  assert.deepEqual(viewModel(document).hiddenColumns, [1]);
  operation.undo();
  assert.deepEqual(viewModel(document).hiddenColumns, []);
  operation.redo();
  await document.save();
  assert.deepEqual(analyze(await fs.readFile(file)).sheets[0].hiddenColumns, [1]);
  document.applyHiddenColumns(0, [1], false);
  await document.save();
  assert.deepEqual(analyze(await fs.readFile(file)).sheets[0].hiddenColumns, []);
  document.dispose();
});

test('행 그룹은 문서 undo·redo와 XLSX 저장·다시 열기에 반영한다', async t => {
  const rows = ['<row r="1"><c r="A1" t="inlineStr"><is><t>Header</t></is></c></row>', '<row r="2"><c r="A2" t="inlineStr"><is><t>one</t></is></c></row>', '<row r="3"><c r="A3" t="inlineStr"><is><t>two</t></is></c></row>'].join('');
  const bytes = minimalXlsx('', rows), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-outline-group-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'outline.xlsx');
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  let operation;
  document.onDidChange(change => { operation = change; });
  document.applyOutlineGroups(0, 'row', 'group', 1, 2);
  assert.deepEqual(document.state.sheets[0].rowGroups, [{ start: 1, end: 2, level: 1, summary: 3, collapsed: false }]);
  operation.undo();
  assert.deepEqual(document.state.sheets[0].rowGroups, []);
  operation.redo();
  document.applyOutlineGroups(0, 'row', 'collapse', 1, 2);
  assert.deepEqual(viewModel(document).hiddenRows, [1, 2]);
  operation.undo();
  assert.deepEqual(viewModel(document).hiddenRows, undefined);
  operation.redo();
  await document.save();
  const reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.writable, true, reopened.readOnlyReason);
  assert.deepEqual(reopened.sheets[0].rowGroups, [{ start: 1, end: 2, level: 1, summary: 3, collapsed: true }]);
  assert.deepEqual(viewModel(new JinSheetDocument(uri(file), 'xlsx', reopened, await fs.readFile(file), true, '')).hiddenRows, [1, 2]);
  document.dispose();
});

test('요약 머리글의 그룹 토글은 같은 위치의 중첩 그룹 전체를 undo·redo와 저장에 반영한다', async t => {
  const rows = ['<row r="1"><c r="A1" t="inlineStr"><is><t>Header</t></is></c></row>', '<row r="2"><c r="A2" t="inlineStr"><is><t>one</t></is></c></row>', '<row r="3"><c r="A3" t="inlineStr"><is><t>two</t></is></c></row>', '<row r="4"><c r="A4" t="inlineStr"><is><t>summary</t></is></c></row>'].join('');
  const bytes = minimalXlsx('', rows), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-outline-summary-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'outline-summary.xlsx');
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  t.after(() => document.dispose());
  let operation;
  document.onDidChange(change => { operation = change; });
  document.applyOutlineGroups(0, 'row', 'group', 1, 2);
  document.applyOutlineGroups(0, 'row', 'group', 2, 2);
  assert.deepEqual(document.state.sheets[0].rowGroups.map(({ start, end, level, summary, collapsed }) => ({ start, end, level, summary, collapsed })), [
    { start: 1, end: 2, level: 1, summary: 3, collapsed: false },
    { start: 2, end: 2, level: 2, summary: 3, collapsed: false }
  ]);
  document.applyOutlineGroups(0, 'row', 'collapse-summary', undefined, undefined, 3);
  assert.deepEqual(document.state.sheets[0].rowGroups.map(group => group.collapsed), [true, true]);
  assert.deepEqual(viewModel(document).hiddenRows, [1, 2]);
  operation.undo();
  assert.deepEqual(document.state.sheets[0].rowGroups.map(group => group.collapsed), [false, false]);
  operation.redo();
  await document.save();
  const reopened = analyze(await fs.readFile(file));
  assert.deepEqual(reopened.sheets[0].rowGroups.map(group => group.collapsed), [true, true]);
  document.applyOutlineGroups(0, 'row', 'expand-summary', undefined, undefined, 3);
  assert.deepEqual(document.state.sheets[0].rowGroups.map(group => group.collapsed), [false, false]);
  assert.deepEqual(viewModel(document).hiddenRows, undefined);
});

test('틀 고정은 undo·redo와 XLSX 저장·다시 열기에 반영한다', async t => {
  const bytes = minimalXlsx(), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-freeze-pane-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'freeze.xlsx');
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  let operation;
  document.onDidChange(change => { operation = change; });
  document.applyFreezePane(0, { rows: 1, columns: 1 });
  assert.deepEqual(viewModel(document).freezePane, { rows: 1, columns: 1 });
  operation.undo();
  assert.deepEqual(viewModel(document).freezePane, { rows: 0, columns: 0 });
  operation.redo();
  await document.save();
  assert.deepEqual(analyze(await fs.readFile(file)).sheets[0].freezePane, { rows: 1, columns: 1 });
  document.applyFreezePane(0, { rows: 0, columns: 0 });
  await document.save();
  assert.deepEqual(analyze(await fs.readFile(file)).sheets[0].freezePane, { rows: 0, columns: 0 });
  document.dispose();
});

test('셀 병합은 데이터 손실을 확인하고 undo·redo·저장·부분 선택 병합 해제를 지원한다', async t => {
  const rows = '<row r="1"><c r="A1" t="inlineStr"><is><t>keep</t></is></c><c r="B1" t="inlineStr"><is><t>discard</t></is></c></row><row r="2"><c r="B2"><f>1+1</f><v>2</v></c></row>';
  const bytes = minimalXlsx('', rows), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-merge-cells-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'merged.xlsx');
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  let operation;
  document.onDidChange(change => { operation = change; });
  const range = { top: 0, bottom: 1, left: 0, right: 1 };
  assert.throws(() => document.applyMergedCells(0, range, 'merge'), error => error.needsConfirmation === true);
  assert.equal(document.getCell(0, 1).value, 'discard', 'confirmation is required before values are cleared');
  document.applyMergedCells(0, range, 'merge', true);
  assert.deepEqual(viewModel(document).mergedCells, [Object.assign({ ref: 'A1:B2' }, range)]);
  assert.equal(document.getCell(0, 0).value, 'keep');
  assert.equal(document.getCell(0, 1).value, '');
  assert.equal(document.getCell(1, 1).formula, null);
  operation.undo();
  assert.equal(document.getCell(0, 1).value, 'discard');
  assert.equal(document.getCell(1, 1).formula, '1+1');
  operation.redo();
  await document.save();
  const reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].mergedCells[0].ref, 'A1:B2');
  assert.equal(reopened.sheets[0].cells.get('B1').value, '');
  assert.equal(reopened.sheets[0].cells.get('B2').formula, null);
  document.applyMergedCells(0, { top: 0, bottom: 0, left: 1, right: 1 }, 'unmerge');
  assert.equal(document.state.sheets[0].mergedCells.length, 0);
  await document.save();
  assert.equal(analyze(await fs.readFile(file)).sheets[0].mergedCells.length, 0);
  document.dispose();
});

test('시트 이름 변경은 참조·데이터 유효성·undo/redo와 저장에 반영한다', async t => {
  const bytes = twoSheetValidationXlsx(), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-rename-sheet-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'rename.xlsx');
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  let operation;
  document.onDidChange(change => { operation = change; });
  document.applyCellEdit(0, 0, 'updated');
  document.applySheetName(1, 'Options');
  assert.equal(viewModel(document, 1).sheets[1].name, 'Options');
  assert.equal(document.state.sheets[0].dataValidations[0].source.sheetName, 'Options');
  operation.undo();
  assert.equal(viewModel(document, 1).sheets[1].name, 'Lists');
  operation.redo();
  await document.save();
  const saved = analyze(await fs.readFile(file));
  assert.equal(saved.sheets[1].name, 'Options');
  assert.equal(saved.sheets[0].dataValidations[0].source.sheetName, 'Options');
  assert.equal(saved.sheets[0].cells.get('A1').value, 'updated');
  document.dispose();
});

test('새 워크시트 추가는 활성화·undo/redo와 저장에 반영한다', async t => {
  const bytes = minimalXlsx(), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-add-sheet-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'add-sheet.xlsx');
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  let operation;
  document.onDidChange(change => { operation = change; });
  document.applyCellEdit(0, 0, 'kept');
  assert.equal(document.applyAddSheet(), 1);
  assert.equal(viewModel(document, 1).sheets[1].name, 'Sheet2');
  operation.undo();
  assert.equal(document.state.sheets.length, 1);
  operation.redo();
  await document.save();
  const saved = analyze(await fs.readFile(file));
  assert.equal(saved.sheets.length, 2);
  assert.equal(saved.sheets[1].name, 'Sheet2');
  assert.equal(saved.sheets[0].cells.get('A1').value, 'kept');
  document.dispose();
});

test('차트 삽입은 미리보기에 반영되고 undo/redo와 저장에 반영한다', async t => {
  const rows = '<row r="1"><c r="A1" t="inlineStr"><is><t>Month</t></is></c><c r="B1" t="inlineStr"><is><t>Sales</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Jan</t></is></c><c r="B2"><v>10</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>Feb</t></is></c><c r="B3"><v>20</v></c></row>';
  const bytes = minimalXlsx('', rows), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-add-chart-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'chart.xlsx');
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  let operation;
  document.onDidChange(change => { operation = change; });
  document.applyAddChart(0, { top: 0, bottom: 2, left: 0, right: 1 });
  assert.equal(document.state.sheets[0].charts.length, 1);
  assert.match(viewModel(document, 0).chartPreviewHtml, /Sales/);
  operation.undo();
  assert.equal(document.state.hasCharts, false);
  operation.redo();
  await document.save();
  assert.equal(analyze(await fs.readFile(file)).sheets[0].charts[0].series[0].values[1], 20);
  document.dispose();
});

test('참조되지 않는 시트 삭제는 선택 이동·undo/redo와 저장에 반영한다', async t => {
  const bytes = twoSheetValidationXlsx(), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-delete-sheet-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'delete-sheet.xlsx');
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  let operation;
  document.onDidChange(change => { operation = change; });
  assert.equal(document.applyDeleteSheet(0), 0);
  assert.equal(document.state.sheets.length, 1);
  assert.equal(document.state.sheets[0].name, 'Lists');
  operation.undo();
  assert.equal(document.state.sheets.length, 2);
  operation.redo();
  await document.save();
  const saved = analyze(await fs.readFile(file));
  assert.equal(saved.sheets.length, 1);
  assert.equal(saved.sheets[0].name, 'Lists');
  document.dispose();
});

test('워크시트 순서 변경은 undo/redo와 저장에 반영한다', async t => {
  const bytes = minimalXlsx(), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-move-sheet-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'move-sheet.xlsx');
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  let operation;
  document.onDidChange(change => { operation = change; });
  document.applyAddSheet();
  document.applyMoveSheet(0, 1);
  assert.deepEqual(document.state.sheets.map(sheet => sheet.name), ['Sheet2', 'Data']);
  operation.undo();
  assert.deepEqual(document.state.sheets.map(sheet => sheet.name), ['Data', 'Sheet2']);
  operation.redo();
  await document.save();
  const saved = analyze(await fs.readFile(file));
  assert.deepEqual(saved.sheets.map(sheet => sheet.name), ['Sheet2', 'Data']);
  document.dispose();
});

test('XLSX 혼합 편집 이력은 저장 지점을 지나도 역순 undo와 정방향 redo를 보존한다', async t => {
  const rows = [
    '<row r="1"><c r="A1" t="inlineStr"><is><t>Kind</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c></row>',
    '<row r="2"><c r="A2" t="inlineStr"><is><t>apple</t></is></c><c r="B2"><v>1</v></c></row>',
    '<row r="3"><c r="A3" t="inlineStr"><is><t>pear</t></is></c><c r="B3"><v>2</v></c></row>',
    '<row r="4"><c r="A4" t="inlineStr"><is><t>apple</t></is></c><c r="B4"><v>3</v></c></row>'
  ].join('');
  const bytes = minimalXlsx('', rows), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-mixed-undo-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'mixed.xlsx');
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, ''), edits = [];
  document.onDidChange(change => edits.push(change));

  document.applyCellEdit(1, 1, '10');
  await document.save();
  document.applyColumnWidth(1, 190);
  document.applyAutoFilter(0, { top: 0, bottom: 3, left: 0, right: 1, columns: [{ column: 0, values: ['apple'], blanks: false }] });
  document.applyStructureChange({ axis: 'row', action: 'insert', index: 2 });
  await document.save();
  const mixed = analyze(await fs.readFile(file));
  assert.equal(mixed.sheets[0].cells.get('B2')?.value, '10');
  assert.ok(mixed.sheets[0].columnWidths[1] > 0);
  assert.deepEqual(document.state.sheets[0].autoFilter.columns, [{ column: 0, values: ['apple'], blanks: false }]);
  assert.equal(document.state.sheets[0].cells.get('A4')?.value, 'pear', 'insert shifts the previous row 3 data to row 4');

  for (const edit of edits.slice().reverse()) {
    edit.undo();
    await document.save();
  }
  const undone = analyze(await fs.readFile(file));
  assert.equal(undone.sheets[0].cells.get('B2')?.value, '1');
  assert.equal(undone.sheets[0].cells.get('A3')?.value, 'pear');
  assert.equal(undone.sheets[0].autoFilter, null);
  assert.deepEqual(undone.sheets[0].columnWidths, {});

  for (const edit of edits) {
    edit.redo();
    await document.save();
  }
  const redone = analyze(await fs.readFile(file));
  assert.equal(redone.sheets[0].cells.get('B2')?.value, '10');
  assert.equal(redone.sheets[0].cells.get('A4')?.value, 'pear');
  assert.deepEqual(redone.sheets[0].autoFilter.columns, [{ column: 0, values: ['apple'], blanks: false }]);
  assert.ok(redone.sheets[0].columnWidths[1] > 0);
  document.dispose();
});

test('XLSX AutoFilter 범위가 셀 수 제한을 넘으면 읽기 전용으로 연다', () => {
  const autoFilter = '<autoFilter ref="A1:DX1000"><filterColumn colId="0"><filters><filter val="apple"/></filters></filterColumn></autoFilter>';
  const result = analyze(minimalXlsx('', '', autoFilter));
  assert.equal(result.writable, false);
  assert.match(result.readOnlyReason, /AutoFilter/);
});

test('AutoFilter 값 목록은 필터 범위의 고유 값과 현재 선택을 제공하고 조건을 갱신한다', () => {
  const rows = [['Kind'], ['pear'], ['apple'], ['pear'], ['']].map(([value]) => [{ value, kind: 'string' }]);
  const filter = { top: 0, bottom: 4, left: 0, right: 0, columns: [{ column: 0, values: ['pear'], blanks: false }] };
  const document = { kind: 'xlsx', state: { sheets: [{ name: 'Data', autoFilter: filter }] }, getCell: row => rows[row]?.[0] };
  assert.deepEqual(autoFilterOptions(document, 0, 0), { values: ['apple', 'pear'], blanks: true, colors: { fills: [{ color: null, count: 4 }], fonts: [{ color: null, count: 4 }], fillCount: 1, fontCount: 1 }, selectedValues: ['pear'], includeBlanks: false });
  const changed = updateAutoFilterColumn(document, 0, 0, ['apple'], true);
  assert.deepEqual(changed.columns, [{ column: 0, values: ['apple'], blanks: true }]);
  assert.deepEqual(updateAutoFilterColumn(document, 0, 0, [], false).columns, [{ column: 0, values: [], blanks: false }], 'clearing all checkboxes represents a zero-result filter');
  assert.deepEqual(updateAutoFilterColumn(document, 0, 0, [], false, true).columns, []);
  assert.throws(() => updateAutoFilterColumn(document, 0, 0, ['missing'], false), /유효하지 않습니다/);
  const manyRows = Array.from({ length: 1102 }, (_, row) => [{ value: 'value-' + row, kind: 'string' }]);
  const manyDocument = { kind: 'xlsx', state: { sheets: [{ name: 'Many', autoFilter: { top: 0, bottom: 1101, left: 0, right: 0, columns: [] } }] }, getCell: row => manyRows[row]?.[0] };
  assert.equal(autoFilterOptions(manyDocument, 0, 0).values.length, 1101, 'filter options are not capped at 1,000 distinct values');
});

test('조건 필터는 수식 결과·다른 열 조건·SUBTOTAL·저장 전후 Undo를 함께 반영한다', async t => {
  const rows = '<row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Score</t></is></c></row>' +
    [['alpha', 5], ['beta', 10], ['alphabet', 20], ['gamma', 15]].map(([name, value], index) => `<row r="${index + 2}"><c r="A${index + 2}" t="inlineStr"><is><t>${name}</t></is></c><c r="B${index + 2}">${index === 3 ? '<f>D5*2</f><v>15</v>' : '<v>' + value + '</v>'}</c>${index === 3 ? '<c r="D5"><v>7.5</v></c>' : ''}</row>`).join('') + '<row r="7"><c r="B7"><f>SUBTOTAL(9,B2:B5)</f><v>50</v></c></row>';
  const bytes = minimalXlsx('', rows), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-condition-filter-')); t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'conditions.xlsx'); await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, ''); let operation; document.onDidChange(change => { operation = change; });
  const base = { top: 0, bottom: 4, left: 0, right: 1, columns: [] }, custom = { and: true, conditions: [{ operator: 'greaterThanOrEqual', value: '10' }, { operator: 'lessThan', value: '20' }] };
  document.applyAutoFilter(0, updateAutoFilterCustomColumn(document, 0, 1, custom, base)); const filterOperation = operation;
  assert.deepEqual(viewModel(document).hiddenRows, [1, 3]); assert.equal(getRange(document, { requestId: 1, rowStart: 6, columnStart: 1, rowCount: 1, columnCount: 1 }).rows[0][0].cachedValue, 25);
  const options = autoFilterOptions(document, 0, 1); assert.deepEqual(options.selectedValues, ['10', '15']); assert.deepEqual(options.custom, custom);
  document.applyCellEdit(4, 3, '12'); assert.deepEqual(viewModel(document).hiddenRows, [1, 3, 4]);
  operation.undo(); assert.deepEqual(viewModel(document).hiddenRows, [1, 3]); await document.save(); filterOperation.undo(); assert.equal(viewModel(document).hiddenRows, undefined);
  filterOperation.redo(); assert.deepEqual(viewModel(document).hiddenRows, [1, 3]); await document.save(); assert.deepEqual(analyze(await fs.readFile(file)).sheets[0].autoFilter.columns[0].custom, custom);
  document.applyAutoFilter(0, updateAutoFilterCustomColumn(document, 0, 0, { and: false, conditions: [{ operator: 'equal', value: '*beta*' }] })); assert.deepEqual(viewModel(document).hiddenRows, [1, 3, 4]);
  document.applyAutoFilter(0, updateAutoFilterColumn(document, 0, 1, [], false, true)); assert.equal(document.state.sheets[0].autoFilter.columns.length, 1); document.dispose();
});

test('표의 조건 필터도 표 XML·계산 결과·다시 값 목록으로 변경을 지원한다', () => {
  const bytes = tableXlsx(), document = new JinSheetDocument(uri('/tmp/table-custom-filter.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  const filter = updateAutoFilterCustomColumn(document, 0, 1, { and: false, conditions: [{ operator: 'greaterThan', value: '10' }] }, autoFilterRange(document, 0, 0, 1)); document.applyAutoFilter(0, filter);
  assert.deepEqual(viewModel(document).hiddenRows, [1]); assert.match(new TextDecoder().decode(document.state.files['xl/tables/table1.xml']), /<customFilters>/); assert.equal(document.state.sheets[0].autoFilter, null);
  assert.equal(getRange(document, { requestId: 2, rowStart: 3, columnStart: 1, rowCount: 1, columnCount: 1 }).rows[0][0].cachedValue, 20);
  document.applyAutoFilter(0, updateAutoFilterColumn(document, 0, 1, ['10'], false, false, filter)); assert.deepEqual(viewModel(document).hiddenRows, [2]); document.dispose();
});

test('Boolean 값과 Boolean 수식의 필터는 화면과 SUBTOTAL에서 같은 행을 고른다', () => {
  const rows = '<row r="1"><c r="A1" t="inlineStr"><is><t>Flag</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c></row><row r="2"><c r="A2" t="b"><v>1</v></c><c r="B2"><v>10</v></c></row><row r="3"><c r="A3" t="b"><v>0</v></c><c r="B3"><v>20</v></c></row><row r="4"><c r="A4"><f>1=1</f><v>1</v></c><c r="B4"><v>30</v></c></row><row r="6"><c r="B6"><f>SUBTOTAL(9,B2:B4)</f><v>60</v></c></row>';
  const bytes = minimalXlsx('', rows), document = new JinSheetDocument(uri('/tmp/boolean-filter.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  const filter = { top: 0, bottom: 3, left: 0, right: 1, columns: [{ column: 0, values: ['TRUE'], blanks: false }] };
  document.applyAutoFilter(0, filter);
  const check = () => { assert.deepEqual(viewModel(document).hiddenRows, [2]); assert.equal(getRange(document, { requestId: 3, rowStart: 5, columnStart: 1, rowCount: 1, columnCount: 1 }).rows[0][0].cachedValue, 40); };
  check(); document.applyAutoFilter(0, updateAutoFilterCustomColumn(document, 0, 0, { and: false, conditions: [{ operator: 'equal', value: 'TRUE' }] })); check(); document.dispose();
});

test('같은 조건을 열 순서만 바꿔 적용해도 dirty·Undo 항목을 만들지 않는다', () => {
  const bytes = minimalXlsx('', '<row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1"><v>10</v></c></row>', '<autoFilter ref="A1:B3"><filterColumn colId="0"><customFilters><customFilter operator="equal" val="*a*"/></customFilters></filterColumn><filterColumn colId="1"><customFilters><customFilter operator="greaterThan" val="10"/></customFilters></filterColumn></autoFilter>');
  const document = new JinSheetDocument(uri('/tmp/unchanged-filter.xlsx'), 'xlsx', analyze(bytes), bytes, true, ''); const original = document.state; let events = 0; document.onDidChange(() => events++);
  const filter = original.sheets[0].autoFilter; document.applyAutoFilter(0, { ...filter, columns: [...filter.columns].reverse() });
  assert.equal(document.state, original); assert.equal(document.dirty, false); assert.equal(events, 0); assert.deepEqual(Buffer.from(document.state.bytes), Buffer.from(bytes)); document.dispose();
});

test('Excel 표 필터는 고유 값·숨김 행을 계산하고 표 XML에 저장한다', () => {
  const bytes = tableXlsx(), document = new JinSheetDocument(uri('/tmp/table-filter.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  const filter = autoFilterRange(document, 0, 0, 1);
  assert.deepEqual(filter, { top: 0, bottom: 2, left: 0, right: 1, columns: [] });
  assert.deepEqual(autoFilterOptions(document, 0, 1, filter), { values: ['10', '20'], blanks: false, colors: { fills: [{ color: null, count: 2 }], fonts: [{ color: null, count: 2 }], fillCount: 1, fontCount: 1 }, selectedValues: ['10', '20'], includeBlanks: false });
  let operation;
  document.onDidChange(change => { operation = change; });
  const filtered = updateAutoFilterColumn(document, 0, 1, ['20'], false, false, filter);
  document.applyAutoFilter(0, filtered);
  assert.deepEqual(document.state.sheets[0].tables[0].autoFilter.columns, [{ column: 1, values: ['20'], blanks: false }]);
  assert.deepEqual(viewModel(document).hiddenRows, [1]);
  assert.equal(getRange(document, { requestId: 9, rowStart: 3, columnStart: 1, rowCount: 1, columnCount: 1 }).rows[0][0].cachedValue, 20, 'SUBTOTAL excludes the row filtered by the table');
  assert.equal(getRange(document, { requestId: 10, rowStart: 3, columnStart: 3, rowCount: 1, columnCount: 1 }).rows[0][0].cachedValue, 30, 'a formula can aggregate a table column through a structured reference');
  assert.match(new TextDecoder().decode(document.state.files['xl/tables/table1.xml']), /filter val="20"/);
  operation.undo();
  assert.deepEqual(viewModel(document).hiddenRows, undefined);
  operation.redo();
  assert.deepEqual(viewModel(document).hiddenRows, [1]);
  const cleared = updateAutoFilterColumn(document, 0, 1, [], false, true, filtered);
  document.applyAutoFilter(0, cleared);
  assert.deepEqual(document.state.sheets[0].tables[0].autoFilter.columns, []);
  assert.deepEqual(viewModel(document).hiddenRows, undefined);
  document.dispose();
});

test('Excel 표 범위 조정은 표와 필터 참조를 함께 갱신하고 undo·redo·저장을 지원한다', async t => {
  const base = tableXlsx(), files = Object.fromEntries(Object.entries(unzipSync(base)).map(([name, value]) => [name, new TextDecoder().decode(value)]));
  files['xl/worksheets/sheet1.xml'] = files['xl/worksheets/sheet1.xml'].replace('</row>', '<c r="C1" t="inlineStr"><is><t>Category</t></is></c></row>');
  const bytes = zipSync(Object.fromEntries(Object.entries(files).map(([name, value]) => [name, new TextEncoder().encode(value)]))), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-table-resize-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'table.xlsx'); await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  let change;
  document.onDidChange(event => { change = event; });
  document.applyTableResize(0, 'Sales', { top: 0, bottom: 3, left: 0, right: 2 });
  assert.equal(document.state.sheets[0].tables[0].range.bottom, 3);
  assert.deepEqual(document.state.sheets[0].tables[0].columns, ['Name', 'Amount', 'Category']);
  let tableXml = new TextDecoder().decode(document.state.files['xl/tables/table1.xml']);
  assert.match(tableXml, /<table[^>]*ref="A1:C4"/);
  assert.match(tableXml, /<autoFilter[^>]*ref="A1:C4"/);
  assert.match(tableXml, /<tableColumns count="3"><tableColumn id="1" name="Name"\/><tableColumn id="2" name="Amount"\/><tableColumn id="3" name="Category"\/><\/tableColumns>/);
  change.undo();
  assert.equal(document.state.sheets[0].tables[0].range.bottom, 2);
  assert.deepEqual(document.state.sheets[0].tables[0].columns, ['Name', 'Amount']);
  change.redo();
  await document.save();
  assert.equal(analyze(await fs.readFile(file)).sheets[0].tables[0].range.right, 2);
  document.applyCellEdit(3, 3, '=SUM(Sales[Category])', 0, new Set(), { skipExpansion: true });
  assert.throws(() => document.applyTableResize(0, 'Sales', { top: 0, bottom: 2, left: 0, right: 1 }), /참조하는 수식/);
  document.applyCellEdit(3, 3, '=SUM(Sales[Amount])', 0, new Set(), { skipExpansion: true });
  document.applyTableResize(0, 'Sales', { top: 0, bottom: 2, left: 0, right: 1 });
  await document.save();
  tableXml = new TextDecoder().decode(analyze(await fs.readFile(file)).files['xl/tables/table1.xml']);
  assert.match(tableXml, /<table[^>]*ref="A1:B3"/);
  assert.match(tableXml, /<autoFilter[^>]*ref="A1:B3"/);
  assert.match(tableXml, /<tableColumns count="2"><tableColumn id="1" name="Name"\/><tableColumn id="2" name="Amount"\/><\/tableColumns>/);
  assert.throws(() => document.applyTableResize(0, 'Sales', { top: 0, bottom: 3, left: 0, right: 3 }), /표 열/);
  document.dispose();
});

test('Excel 표 데이터 행 삽입·삭제는 표 범위와 undo·redo·저장을 동기화한다', async t => {
  const bytes = tableXlsx(), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-table-rows-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'table.xlsx'); await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  let change;
  document.onDidChange(event => { change = event; });
  document.applyStructureChange({ axis: 'row', action: 'insert', index: 2, sheetIndex: 0 });
  assert.equal(document.state.sheets[0].tables[0].range.bottom, 3);
  assert.match(new TextDecoder().decode(document.state.files['xl/tables/table1.xml']), /ref="A1:B4"/);
  change.undo();
  assert.equal(document.state.sheets[0].tables[0].range.bottom, 2);
  change.redo();
  await document.save();
  const reopenedBytes = await fs.readFile(file), reopened = analyze(reopenedBytes);
  assert.equal(reopened.writable, true, reopened.readOnlyReason);
  assert.equal(reopened.sheets[0].tables[0].range.bottom, 3);
  document.applyStructureChange({ axis: 'row', action: 'delete', index: 2, sheetIndex: 0 });
  assert.equal(document.state.sheets[0].tables[0].range.bottom, 2);
  await document.save();
  assert.equal(analyze(await fs.readFile(file)).sheets[0].tables[0].range.bottom, 2);
  document.dispose();
});

test('Excel 표 범위 조정은 왼쪽 열도 추가·제거하고 기존 셀 값과 필터 인덱스를 보존한다', () => {
  const files = Object.fromEntries(Object.entries(unzipSync(tableXlsx())).map(([name, value]) => [name, new TextDecoder().decode(value)]));
  files['xl/worksheets/sheet1.xml'] = files['xl/worksheets/sheet1.xml'].replace(/r="([A-Z]+)(\d+)"/g, (_match, letters, row) => 'r="' + ({ A: 'B', B: 'C', D: 'E' }[letters] || letters) + row + '"').replace(/<f>SUBTOTAL\(9,B2:B3\)<\/f>/, '<f>SUBTOTAL(9,C2:C3)</f>').replace('<row r="1">', '<row r="1"><c r="A1" t="inlineStr"><is><t>Region</t></is></c>');
  files['xl/tables/table1.xml'] = files['xl/tables/table1.xml'].replaceAll('A1:B3', 'B1:C3');
  const bytes = zipSync(Object.fromEntries(Object.entries(files).map(([name, value]) => [name, new TextEncoder().encode(value)]))), document = new JinSheetDocument(uri('/tmp/table-left-resize.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  assert.equal(document.writable, true, document.state.readOnlyReason);
  document.applyAutoFilter(0, { top: 0, bottom: 2, left: 1, right: 2, columns: [{ column: 2, values: ['10'], blanks: false }] });
  document.applyTableResize(0, 'Sales', { top: 0, bottom: 2, left: 0, right: 2 });
  assert.deepEqual(document.state.sheets[0].tables[0].columns, ['Region', 'Name', 'Amount']);
  assert.deepEqual([document.state.sheets[0].tables[0].range.left, document.state.sheets[0].tables[0].range.right], [0, 2]);
  assert.equal(document.state.sheets[0].tables[0].autoFilter.columns[0].column, 2);
  assert.equal(document.state.sheets[0].cells.get('A1').value, 'Region');
  document.applyTableResize(0, 'Sales', { top: 0, bottom: 2, left: 2, right: 2 });
  assert.deepEqual(document.state.sheets[0].tables[0].columns, ['Amount']);
  assert.deepEqual([document.state.sheets[0].tables[0].range.left, document.state.sheets[0].tables[0].range.right], [2, 2]);
  assert.equal(document.state.sheets[0].tables[0].autoFilter.columns[0].column, 2);
  assert.equal(document.state.sheets[0].cells.get('B1').value, 'Name', 'shrinking a table keeps the excluded worksheet cells');
  document.dispose();
});

test('Excel 표 열 삽입·삭제는 머리글·정의·필터 범위를 갱신하고 참조 수식 열 삭제를 막는다', async t => {
  const bytes = tableXlsx(), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-table-columns-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'table.xlsx'); await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  let change;
  document.onDidChange(event => { change = event; });
  document.applyAutoFilter(0, { top: 0, bottom: 2, left: 0, right: 1, columns: [{ column: 1, values: ['10'], blanks: false }] });
  document.applyStructureChange({ axis: 'column', action: 'insert', index: 1, sheetIndex: 0 });
  let sheet = document.state.sheets[0], table = sheet.tables[0];
  assert.deepEqual(table.columns, ['Name', 'Column2', 'Amount']);
  assert.deepEqual([table.range.left, table.range.right], [0, 2]);
  assert.equal(sheet.cells.get('B1').value, 'Column2');
  assert.equal(sheet.cells.get('C1').value, 'Amount');
  assert.equal(table.autoFilter.columns[0].column, 2);
  assert.match(new TextDecoder().decode(document.state.files['xl/tables/table1.xml']), /<table[^>]*ref="A1:C3"/);
  change.undo();
  assert.deepEqual(document.state.sheets[0].tables[0].columns, ['Name', 'Amount']);
  change.redo();
  await document.save();
  assert.equal(analyze(await fs.readFile(file)).sheets[0].tables[0].columns[1], 'Column2');
  document.applyStructureChange({ axis: 'column', action: 'delete', index: 1, sheetIndex: 0 });
  table = document.state.sheets[0].tables[0];
  assert.deepEqual(table.columns, ['Name', 'Amount']);
  assert.deepEqual([table.range.left, table.range.right], [0, 1]);
  assert.equal(document.state.sheets[0].cells.get('B1').value, 'Amount');
  assert.equal(table.autoFilter.columns[0].column, 1);
  assert.throws(() => document.applyStructureChange({ axis: 'column', action: 'delete', index: 1, sheetIndex: 0 }), /참조하는 수식/);
  await document.save();
  assert.equal(analyze(await fs.readFile(file)).sheets[0].tables[0].range.right, 1);
  document.dispose();
});

test('동적 배열 수식 결과를 그리드에 표시하고 spill 셀 직접 편집을 막는다', () => {
  const rows = '<row r="1"><c r="A1"><f>SEQUENCE(2,2)</f><v>1</v></c><c r="C1"><f>SUM(A1#)</f><v>10</v></c><c r="D1"><f>A1#</f><v>1</v></c></row>';
  const bytes = minimalXlsx('', rows), document = new JinSheetDocument(uri('/tmp/dynamic-array.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  assert.equal(viewModel(document).rowCount, 2);
  assert.equal(viewModel(document).columnCount, 5);
  const range = getRange(document, { requestId: 1, rowStart: 0, columnStart: 0, rowCount: 2, columnCount: 5 });
  assert.deepEqual(range.rows.map(row => row.map(cell => cell?.value ?? null)), [['1', '2', '10', '1', '2'], ['3', '4', '', '3', '4']]);
  assert.equal(range.rows[1][1].hasFormula, false);
  assert.equal(range.rows[1][3].isSpill, true);
  assert.throws(() => document.applyCellEdit(1, 1, '99', 0), /동적 배열 결과의 일부 셀은 편집할 수 없습니다/);
  assert.throws(() => document.applyCellEdit(1, 3, '99', 0), /동적 배열 결과의 일부 셀은 편집할 수 없습니다/);
  document.dispose();
});

test('그리드 범위에 XLSX 데이터 막대의 색상과 상대 길이를 전달한다', () => {
  const rows = '<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row><row r="3"><c r="A3"><v>3</v></c></row>';
  const dataBar = '<conditionalFormatting sqref="A1:A3"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule></conditionalFormatting>';
  const bytes = minimalXlsx('', rows, dataBar), document = new JinSheetDocument(uri('/tmp/data-bars.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  const range = getRange(document, { requestId: 1, rowStart: 0, columnStart: 0, rowCount: 3, columnCount: 1 });
  assert.deepEqual(range.rows.map(row => row[0].style.dataBar.width), [0, 50, 100]);
  assert.deepEqual(range.rows.map(row => row[0].style.dataBar.color), ['#638EC6', '#638EC6', '#638EC6']);
  document.dispose();
});

test('그리드 범위에 XLSX 색조의 보간 색을 전달한다', () => {
  const rows = '<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row><row r="3"><c r="A3"><v>3</v></c></row>';
  const colorScale = '<conditionalFormatting sqref="A1:A3"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FFFFFF00"/><color rgb="FF00FF00"/></colorScale></cfRule></conditionalFormatting>';
  const bytes = minimalXlsx('', rows, colorScale), document = new JinSheetDocument(uri('/tmp/color-scales.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  const range = getRange(document, { requestId: 1, rowStart: 0, columnStart: 0, rowCount: 3, columnCount: 1 });
  assert.deepEqual(range.rows.map(row => row[0].style.fill), ['#FF0000', '#FFFF00', '#00FF00']);
  document.dispose();
});

test('FILTER 결과를 그리드에 표시하고 선택 행을 동적으로 갱신한다', () => {
  const rows = '<row r="1"><c r="A1"><v>10</v></c><c r="B1" t="b"><v>1</v></c><c r="D1"><f>FILTER(A1:A3,B1:B3)</f><v>10</v></c></row>'
    + '<row r="2"><c r="A2"><v>20</v></c><c r="B2" t="b"><v>0</v></c></row>'
    + '<row r="3"><c r="A3"><v>30</v></c><c r="B3" t="b"><v>1</v></c></row>';
  const bytes = minimalXlsx('', rows), document = new JinSheetDocument(uri('/tmp/dynamic-filter.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  const range = getRange(document, { requestId: 2, rowStart: 0, columnStart: 3, rowCount: 2, columnCount: 1 });
  assert.deepEqual(range.rows.map(row => row[0].value), ['10', '30']);
  assert.equal(range.rows[1][0].isSpill, true);
  assert.throws(() => document.applyCellEdit(1, 3, '99', 0), /동적 배열 결과의 일부 셀은 편집할 수 없습니다/);
  document.dispose();
});

test('필터된 XLSX 범위 지우기는 숨김 행을 보존하고 undo·redo·저장한다', async t => {
  const rows = '<row r="1"><c r="A1" t="inlineStr"><is><t>Kind</t></is></c></row>' +
    '<row r="2"><c r="A2" t="inlineStr"><is><t>apple</t></is></c></row>' +
    '<row r="3"><c r="A3" t="inlineStr"><is><t>pear</t></is></c></row>' +
    '<row r="4"><c r="A4" t="inlineStr"><is><t>apple</t></is></c></row>';
  const bytes = minimalXlsx('', rows, '<autoFilter ref="A1:A4"><filterColumn colId="0"><filters><filter val="apple"/></filters></filterColumn></autoFilter>');
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-filtered-clear-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'filtered.xlsx');
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  assert.deepEqual(viewModel(document).hiddenRows, [2]);
  const edits = buildClearEdits(document, 1, 0, { startRow: 1, endRow: 3, startColumn: 0, endColumn: 0 }, 0, viewModel(document).hiddenRows);
  assert.deepEqual(edits.map(edit => edit.row), [1, 3]);
  let operation;
  document.onDidChange(change => { operation = change; });
  document.applyCellEdits(edits, 0, '필터된 범위 지우기');
  assert.equal(document.getCell(2, 0).value, 'pear');
  operation.undo();
  assert.equal(document.getCell(1, 0).value, 'apple');
  operation.redo();
  await document.save();
  const saved = analyze(await fs.readFile(file));
  assert.equal(saved.sheets[0].cells.get('A3').value, 'pear');
  assert.equal(saved.sheets[0].cells.get('A2').value, '');
  document.dispose();
});

test('Ctrl+A는 연속 데이터 영역 다음 전체 시트를 선택하고 빈 셀도 안전하게 처리한다', () => {
  const rows = Array.from({ length: 5 }, (_, row) => Array.from({ length: 5 }, (_, column) => ({ value: '' })));
  for (const [row, column, value] of [[1, 1, 'a'], [1, 2, 'b'], [2, 1, 'c'], [2, 2, 'd'], [4, 4, 'z']]) rows[row][column] = { value };
  const document = { kind: 'csv', state: { rows }, drafts: new Map() };
  assert.deepEqual(currentRegion(document, 1, 1), { top: 1, bottom: 2, left: 1, right: 2 });
  assert.deepEqual(currentRegion(document, 1, 1, 0, true), { top: 0, bottom: 4, left: 0, right: 4 });
  assert.deepEqual(currentRegion(document, 0, 0), { top: 0, bottom: 0, left: 0, right: 0 });
  rows[0][4] = { value: 'solo' };
  document.drafts.set('0:0:4', '');
  assert.deepEqual(currentRegion(document, 0, 4), { top: 0, bottom: 0, left: 4, right: 4 }, 'draft로 지운 셀은 현재 데이터 영역에서 제외한다');
  const xlsx = { kind: 'xlsx', state: { sheets: [{ cells: new Map() }] }, drafts: new Map() };
  assert.deepEqual(currentRegion(xlsx, 0, 0, 0, true), { top: 0, bottom: 1048575, left: 0, right: 16383 }, 'XLSX 전체 선택은 Excel 워크시트 크기를 따른다');
});

class FakeElement {
  constructor(tagName = 'DIV') {
    this.tagName = tagName;
    this.children = [];
    this.parentElement = null;
    this.style = {};
    this.dataset = {};
    this.attributes = {};
    this.listeners = new Map();
    this.classes = new Set();
    Object.defineProperty(this, 'className', { get: () => [...this.classes].join(' '), set: value => { this.classes = new Set(String(value).split(/\s+/).filter(Boolean)); } });
    this.classList = {
      add: name => this.classes.add(name),
      remove: name => this.classes.delete(name),
      contains: name => this.classes.has(name),
      toggle: (name, force) => { const enabled = force ?? !this.classes.has(name); if (enabled) this.classes.add(name); else this.classes.delete(name); return enabled; }
    };
    this._textContent = '';
    this.value = '';
  }
  get textContent() { return this._textContent + this.children.map(child => child.textContent).join(''); }
  set textContent(value) { this.children = []; this._textContent = String(value ?? ''); }
  get innerHTML() { return this._innerHTML ?? (this.children.map(child => child.outerHTML || escapeHtml(child.textContent)).join('') || escapeHtml(this._textContent)); }
  set innerHTML(value) { const html = String(value ?? ''); this.replaceChildren(...fakeNodes(parseFragment(html).childNodes)); this._textContent = ''; this._innerHTML = html; }
  append(child) { child.parentElement = this; this._innerHTML = undefined; this.children.push(child); }
  insertBefore(child, reference) { if (child.parentElement) child.parentElement.children = child.parentElement.children.filter(item => item !== child); const index = this.children.indexOf(reference); child.parentElement = this; this._innerHTML = undefined; if (index < 0) this.children.push(child); else this.children.splice(index, 0, child); }
  replaceChildren(...children) { for (const child of this.children) { if (activeDocument && isWithin(activeDocument.activeElement, child)) activeDocument.activeElement = activeDocument.body; child.parentElement = null; } this.children = []; this._innerHTML = undefined; for (const child of children) this.append(child); }
  remove() { if (activeDocument && isWithin(activeDocument.activeElement, this)) activeDocument.activeElement = activeDocument.body; if (this.parentElement) this.parentElement.children = this.parentElement.children.filter(child => child !== this); this.parentElement = null; }
  setAttribute(name, value) { const text = String(value); this.attributes[name] = text; if (name === 'id') this.id = text; else if (name === 'class') this.className = text; else if (name.startsWith('data-')) this.dataset[name.slice(5).replace(/-([a-z])/g, (_, char) => char.toUpperCase())] = text; else if (name === 'value') this.value = text; else if (name === 'checked') this.checked = true; else if (name === 'disabled') this.disabled = true; else if (name === 'hidden') this.hidden = true; else if (name === 'type') this.type = text; else if (name === 'title') this.title = text; }
  addEventListener(name, callback) { this.listeners.set(name, callback); }
  closest(selector) { let current = this; while (current) { if (matchesFakeSelector(current, selector)) return current; current = current.parentElement; } return null; }
  contains(target) { return isWithin(target, this); }
  focus() { if (activeDocument) activeDocument.activeElement = this; }
  blur() { if (activeDocument && activeDocument.activeElement === this) activeDocument.activeElement = activeDocument.body; this.onblur?.(); }
  getBoundingClientRect() { const { left = 0, top = 0, width = this.clientWidth || 0, height = this.clientHeight || 0 } = this.rect || {}; return { left, top, width, height, right: left + width, bottom: top + height }; }
  select() { this.selectionStart = 0; this.selectionEnd = String(this.value).length; }
  setSelectionRange(start, end) { this.selectionStart = start; this.selectionEnd = end; }
  setRangeText(text, start = this.selectionStart, end = this.selectionEnd, mode = 'preserve') { this.value = this.value.slice(0, start) + text + this.value.slice(end); const cursor = start + text.length; this.selectionStart = mode === 'select' ? start : cursor; this.selectionEnd = mode === 'select' ? cursor : cursor; }
}

function escapeHtml(value) { return String(value).replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;'); }
function matchesFakeSelector(element, selector) {
  if (!selector || !element) return false;
  const parts = selector.trim().split(/\s+/);
  const matchesSimple = (node, simple) => {
    const tag = /^([a-z][\w-]*)/i.exec(simple)?.[1];
    if (tag && node.tagName.toLowerCase() !== tag.toLowerCase()) return false;
    for (const match of simple.matchAll(/#([\w-]+)/g)) if (node.id !== match[1]) return false;
    for (const match of simple.matchAll(/\.([\w-]+)/g)) if (!node.classList.contains(match[1])) return false;
    for (const match of simple.matchAll(/\[([\w-]+)(?:=["']?([^\]"']+)["']?)?\]/g)) {
      const name = match[1], expected = match[2];
      const actual = name.startsWith('data-') ? node.dataset[name.slice(5).replace(/-([a-z])/g, (_, char) => char.toUpperCase())] : node.attributes[name];
      if (actual === undefined || expected !== undefined && actual !== expected) return false;
    }
    return true;
  };
  let node = element;
  if (!matchesSimple(node, parts.at(-1))) return false;
  for (let index = parts.length - 2; index >= 0; index--) {
    node = node.parentElement;
    while (node && !matchesSimple(node, parts[index])) node = node.parentElement;
    if (!node) return false;
  }
  return true;
}
function fakeNodes(nodes) {
  return (nodes || []).filter(node => node.tagName || node.nodeName === '#text').map(node => {
    if (node.nodeName === '#text') { const text = new FakeElement('#TEXT'); text.textContent = node.value; return text; }
    const element = new FakeElement(node.tagName.toUpperCase());
    for (const attribute of node.attrs || []) element.setAttribute(attribute.name, attribute.value);
    element.children = fakeNodes(node.childNodes);
    for (const child of element.children) child.parentElement = element;
    if (!element.children.length) element._textContent = (node.childNodes || []).map(child => child.value || '').join('');
    if (element.tagName === 'INPUT') { element.checked ||= false; element.disabled ||= false; }
    return element;
  });
}

function runGridWebview(html, { rowCount = 2, columnCount = 2, style = null, kind = 'csv', columnWidths = undefined, defaultColumnWidth = undefined, freezePane = undefined, cellData = {}, visibleRows = undefined, hiddenRows = undefined, hiddenColumns = undefined, manuallyHiddenRows = undefined, manuallyHiddenColumns = undefined, rowGroups = undefined, columnGroups = undefined, autoFilter = undefined, tables = undefined, chartInsertionAvailable = true, rowHeights = undefined, defaultRowHeight = undefined, printAreas = undefined, printSettings = undefined, printTitles = undefined, headerFooter = undefined, pageBreaks = undefined, hasDynamicDateFilters = undefined } = {}) {
  const elements = new Map(), documentListeners = new Map(), windowListeners = new Map(), messages = [], allElements = [];
  const addListener = (registry, name, callback) => { let dispatch = registry.get(name); if (!dispatch) { dispatch = (...args) => { for (const listener of [...dispatch.listeners]) listener(...args); }; dispatch.listeners = []; registry.set(name, dispatch); } dispatch.listeners.push(callback); };
  const parsed = fakeNodes(parseFragment(html).childNodes);
  const indexElements = nodes => { for (const node of nodes) { if (node.id) elements.set(node.id, node); indexElements(node.children); } };
  indexElements(parsed); allElements.push(...parsed);
  const get = id => { if (!elements.has(id)) { const element = new FakeElement(['search', 'replaceValue', 'formula', 'address'].includes(id) ? 'INPUT' : 'DIV'); element.id = id; element.setAttribute('id', id); elements.set(id, element); allElements.push(element); } return elements.get(id); };
  const findAll = (selector, roots = allElements) => roots.flatMap(node => [ ...(matchesFakeSelector(node, selector) ? [node] : []), ...findAll(selector, node.children) ]);
  const document = {
    body: parsed.find(node => node.tagName === 'BODY') || get('body'),
    getElementById: get,
    createElement: name => { const element = new FakeElement(name.toUpperCase()); allElements.push(element); return element; },
    createElementNS: (_namespace, name) => new FakeElement(name.toUpperCase()),
    addEventListener: (name, callback) => addListener(documentListeners, name, callback),
    querySelector: selector => findAll(selector).find(Boolean) || null,
    querySelectorAll: selector => findAll(selector),
    querySelectorAllNS: (_namespace, selector) => findAll(selector)
  };
  activeDocument = document;
  document.activeElement = document.body;
  const grid = get('grid'), canvas = get('canvas');
  grid.clientHeight = 300; grid.clientWidth = 500; grid.scrollTop = 0; grid.scrollLeft = 0; grid.scrollHeight = 30 + rowCount * 26; grid.scrollWidth = 52 + columnCount * 144; grid.rect = { left: 0, top: 0, width: 500, height: 300 }; canvas.parentElement = grid;
  const vscode = { postMessage: message => messages.push(message) };
  const window = { addEventListener: (name, callback) => addListener(windowListeners, name, callback) };
  const script = [...html.matchAll(/<script nonce="[^"]+">([\s\S]*?)<\/script>/g)].at(-1)[1];
  new vm.Script(script).runInNewContext({ acquireVsCodeApi: () => vscode, document, window, Date, setTimeout, clearTimeout, setInterval, clearInterval });
  const dispatch = data => windowListeners.get('message')({ data });
  const respondToRange = () => {
    const request = messages.filter(message => message.type === 'range').at(-1);
    if (!request) return false;
    const rows = Array.from({ length: request.rowCount }, (_, r) => Array.from({ length: request.columnCount }, (_, c) => {
      const row = request.viewRows?.[r] ?? (request.rowStart + r), column = request.columnIndexes?.[c] ?? (request.columnStart + c);
      const values = { '0:0': 'a', '0:1': 'b', '1:0': '1', '1:1': '2' };
      return { value: values[row + ':' + column] ?? `${row + 1}:${column + 1}`, style: row === 1 && column === 1 ? style : null, ...(cellData[row + ':' + column] || {}) };
    }));
    dispatch({ type: 'range', requestId: request.requestId, rowStart: request.rowStart, viewRows: request.viewRows, viewIndexes: request.viewIndexes, rowIndexes: request.viewRows, columnStart: request.columnStart, columnIndexes: request.columnIndexes, rowCount: request.rowCount, columnCount: request.columnCount, sheetIndex: request.sheetIndex, rows });
    return true;
  };
  dispatch({ type: 'model', kind, readOnly: false, rowCount, columnCount, columnWidths, defaultColumnWidth, freezePane, visibleRows, hiddenRows, hiddenColumns, manuallyHiddenRows, manuallyHiddenColumns, rowGroups, columnGroups, autoFilter, tables, chartInsertionAvailable, rowHeights, defaultRowHeight, printAreas, printSettings, printTitles, headerFooter, pageBreaks, hasDynamicDateFilters, sheetIndex: 0, dirty: false });
  respondToRange();
  return { document, documentListeners, elements, messages, windowListeners, respondToRange, dispatch };
}

test('CSV 편집 undo/redo 후 저장하면 수정 값이 기록되고 dirty가 해제된다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-doc-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'sample.csv'), original = Buffer.from('name,value\r\nalpha,1\n');
  await fs.writeFile(file, original);
  const state = load(original, '.csv');
  const document = new JinSheetDocument(uri(file), 'csv', state, original, true, '');
  document.setInitialCsvState(state);
  let edit;
  document.onDidChange(value => { edit = value; });
  document.applyCellEdit(1, 1, '2');
  assert.equal(document.dirty, true);
  edit.undo(); assert.equal(document.state.rows[1][1].value, '1'); assert.equal(document.dirty, false);
  edit.redo(); assert.equal(document.state.rows[1][1].value, '2'); assert.equal(document.dirty, true);
  await document.save();
  assert.equal((await fs.readFile(file, 'utf8')), 'name,value\r\nalpha,2\n');
  assert.equal(document.dirty, false);
  document.dispose();
});

test('CSV 셀의 Alt+Enter 줄바꿈은 인용된 필드로 저장하고 다시 읽는다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-cell-newline-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'multiline.csv'), original = Buffer.from('name,value\r\nalpha,1\r\n');
  await fs.writeFile(file, original);
  const document = new JinSheetDocument(uri(file), 'csv', load(original, '.csv'), original, true, '');
  document.setInitialCsvState(document.state);
  document.applyCellEdit(1, 0, 'alpha\nbeta');
  await document.save();
  const reopened = load(await fs.readFile(file), '.csv');
  assert.equal(reopened.rows[1][0].value, 'alpha\nbeta');
  assert.equal(reopened.rows[1][1].value, '1');
  document.dispose();
});

test('XLSX 열 너비 조절은 dirty·undo/redo·저장과 재열기에 반영된다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-width-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'width.xlsx'), original = minimalXlsx();
  await fs.writeFile(file, original);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(original), original, true, '');
  let operation;
  document.onDidChange(change => { operation = change; });
  document.applyColumnWidth(1, 160);
  assert.equal(document.dirty, true);
  assert.equal(viewModel(document).columnWidths[1], 22.86);
  operation.undo();
  assert.equal(document.dirty, false);
  operation.redo();
  assert.equal(document.dirty, true);
  await document.save();
  assert.equal(document.dirty, false);
  const reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].columnWidths[1], 22.86);
  assert.deepEqual(Buffer.from(reopened.files['xl/workbook.xml']), Buffer.from(document.state.files['xl/workbook.xml']));
  document.dispose();
});

test('CSV 행·열 추가와 삭제가 수정·undo/redo·저장에 반영된다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-structure-csv-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'structure.csv'), original = Buffer.from('name,value\nalpha,1\nbeta,2\n');
  await fs.writeFile(file, original);
  const state = load(original, '.csv'), document = new JinSheetDocument(uri(file), 'csv', state, original, true, '');
  document.setInitialCsvState(state);
  document.applyCellEdit(1, 1, 'edited');
  document.applyStructureChange({ axis: 'row', action: 'insert', index: 1 });
  assert.deepEqual(document.state.rows.map(row => row.map(cell => cell.value)), [['name', 'value'], ['', ''], ['alpha', 'edited'], ['beta', '2']]);
  assert.equal(document.dirty, true);
  document.applyStructureChange({ axis: 'column', action: 'insert', index: 1 });
  assert.deepEqual(document.state.rows[2].map(cell => cell.value), ['alpha', '', 'edited']);
  document.applyStructureChange({ axis: 'row', action: 'delete', index: 1 });
  assert.deepEqual(document.state.rows.map(row => row.map(cell => cell.value)), [['name', '', 'value'], ['alpha', '', 'edited'], ['beta', '', '2']]);
  let latestChange;
  document.onDidChange(change => { latestChange = change; });
  document.applyStructureChange({ axis: 'column', action: 'delete', index: 1 });
  latestChange.undo();
  assert.equal(document.state.rows[1].length, 3);
  latestChange.redo();
  assert.equal(document.state.rows[1].length, 2);
  await document.save();
  assert.equal((await fs.readFile(file, 'utf8')), 'name,value\nalpha,edited\nbeta,2\n');
  assert.equal(document.dirty, false);
  document.dispose();
});

test('선택한 행·열 개수만큼 구조를 한 번에 삽입·삭제하고 단일 undo 단위로 되돌린다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-structure-count-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));

  const csvFile = path.join(directory, 'range.csv'), csvBytes = Buffer.from('a,1\nb,2\nc,3\nd,4\n');
  await fs.writeFile(csvFile, csvBytes);
  const csv = new JinSheetDocument(uri(csvFile), 'csv', load(csvBytes, '.csv'), csvBytes, true, '');
  csv.setInitialCsvState(csv.state);
  let operation;
  csv.onDidChange(change => { operation = change; });
  csv.applyStructureChange({ axis: 'row', action: 'insert', index: 1, count: 2 });
  assert.deepEqual(csv.state.rows.map(row => row[0]?.value ?? ''), ['a', '', '', 'b', 'c', 'd']);
  operation.undo();
  assert.deepEqual(csv.state.rows.map(row => row[0]?.value ?? ''), ['a', 'b', 'c', 'd']);
  operation.redo();
  csv.applyStructureChange({ axis: 'row', action: 'delete', index: 1, count: 2 });
  assert.deepEqual(csv.state.rows.map(row => row[0]?.value ?? ''), ['a', 'b', 'c', 'd']);
  await csv.save();
  assert.deepEqual(await fs.readFile(csvFile), csvBytes);
  csv.dispose();

  const xlsxFile = path.join(directory, 'range.xlsx'), xlsxBytes = minimalXlsx();
  await fs.writeFile(xlsxFile, xlsxBytes);
  const xlsx = new JinSheetDocument(uri(xlsxFile), 'xlsx', analyze(xlsxBytes), xlsxBytes, true, '');
  xlsx.onDidChange(change => { operation = change; });
  xlsx.applyStructureChange({ axis: 'row', action: 'insert', index: 0, count: 2 });
  assert.equal(xlsx.getCell(2, 0).value, 'old');
  operation.undo();
  assert.equal(xlsx.getCell(0, 0).value, 'old');
  operation.redo();
  xlsx.applyStructureChange({ axis: 'row', action: 'delete', index: 0, count: 2 });
  assert.equal(xlsx.getCell(0, 0).value, 'old');
  await xlsx.save();
  assert.equal(analyze(await fs.readFile(xlsxFile)).sheets[0].cells.get('A1').value, 'old');
  xlsx.dispose();
});

test('빈 TSV에 삽입한 첫 행을 저장하고 다시 열 수 있다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-empty-tsv-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'empty.tsv'), original = Buffer.alloc(0);
  await fs.writeFile(file, original);
  const state = load(original, '.tsv'), document = new JinSheetDocument(uri(file), 'csv', state, original, true, '');
  document.setInitialCsvState(state);
  document.applyStructureChange({ axis: 'row', action: 'insert', index: 0 });
  assert.deepEqual(document.state.rows.map(row => row.map(cell => cell.value)), [['']]);
  await document.save();
  assert.equal((await fs.readFile(file, 'utf8')), '""');
  assert.equal(document.dirty, false);
  document.dispose();
});

test('CSV 데이터 끝의 빈 셀은 조회·편집·저장·undo/redo할 수 있다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-csv-blank-grid-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'blank-grid.csv'), bytes = Buffer.from('a,b\n1,2');
  await fs.writeFile(file, bytes);
  const state = load(bytes, '.csv'), document = new JinSheetDocument(uri(file), 'csv', state, bytes, true, '');
  document.setInitialCsvState(state);
  const range = getRange(document, { rowStart: 9999, columnStart: 999, rowCount: 1, columnCount: 1, requestId: 1 });
  assert.deepEqual(range.rows, [[null]]);
  let edit;
  document.onDidChange(change => { edit = change; });
  document.applyCellEdit(4, 3, 'x');
  assert.equal(document.getCell(4, 3).value, 'x');
  await document.save();
  assert.equal(load(await fs.readFile(file), '.csv').rows[4][3].value, 'x');
  edit.undo();
  assert.equal(document.dirty, true);
  assert.equal(document.getCell(4, 3)?.value ?? '', '');
  await document.save();
  assert.deepEqual(await fs.readFile(file), bytes);
  edit.redo();
  assert.equal(document.getCell(4, 3).value, 'x');
  await document.save();
  assert.equal(load(await fs.readFile(file), '.csv').rows[4][3].value, 'x');
  document.dispose();
});

test('저장 뒤 CSV 구조 변경을 undo·redo하면 dirty와 저장 내용이 새 기준 파일에 맞는다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-post-save-undo-csv-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'after-save.csv'), original = Buffer.from('a,b\n1,2\n');
  await fs.writeFile(file, original);
  const state = load(original, '.csv'), document = new JinSheetDocument(uri(file), 'csv', state, original, true, '');
  document.setInitialCsvState(state);
  let edit;
  document.onDidChange(change => { edit = change; });
  document.applyStructureChange({ axis: 'row', action: 'insert', index: 1 });
  await document.save();
  edit.undo();
  assert.equal(document.dirty, true);
  await document.save();
  assert.deepEqual(await fs.readFile(file), original);
  edit.redo();
  assert.equal(document.dirty, true);
  await document.save();
  assert.equal((await fs.readFile(file, 'utf8')), 'a,b\n,\n1,2\n');
  document.dispose();
});

test('잘라내기 범위를 오른쪽으로 이동하면 겹치는 셀을 안전하게 덮고 원본 시작 셀만 비운다', () => {
  const values = new Map([["0:0", 'A'], ["0:1", 'B'], ["0:2", 'C']]);
  const document = { getCell: (row, column) => ({ value: values.get(row + ':' + column) ?? '', kind: 'text' }) };
  const edits = buildMoveEdits(document, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 }, 0, 1);
  assert.deepEqual(edits, [
    { row: 0, column: 1, value: 'A' },
    { row: 0, column: 2, value: 'B' },
    { row: 0, column: 0, value: '' }
  ]);
});

test('필터된 범위 잘라내기 이동은 보이는 원본 행만 연속된 보이는 대상 행으로 옮긴다', () => {
  const rows = [['a'], ['hidden'], ['b'], ['x'], ['y']].map(([value]) => [{ value, kind: 'text' }]);
  const document = { kind: 'csv', state: { rows }, getCell: (row, column) => rows[row]?.[column] };
  const edits = buildMoveEdits(document, { startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 }, 3, 0, 0, [1]);
  assert.deepEqual(edits.map(({ row, value }) => [row, value]), [[3, 'a'], [4, 'b'], [0, ''], [2, '']]);
  assert.equal(edits.some(edit => edit.row === 1), false, 'the hidden source row is neither copied nor cleared');
});

test('선택 범위 채우기는 값·서식을 복제하고 저장 후 한 번 undo·redo한다', async t => {
  const bytes = await fs.readFile(path.join(__dirname, 'fixtures', 'JinSheet-Styled-Sample.xlsx'));
  const state = analyze(bytes), source = state.sheets[0].cells.get('C2');
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-fill-range-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'fill-range.xlsx');
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', state, bytes, true, '');
  let operation, operationCount = 0;
  document.onDidChange(change => { operation = change; operationCount++; });
  const edits = buildFillEdits(document, { startRow: 1, endRow: 1, startColumn: 2, endColumn: 4 }, 'right');
  assert.deepEqual(edits.map(edit => [edit.row, edit.column, edit.value, edit.styleIndex]), [
    [1, 3, source.value, source.style], [1, 4, source.value, source.style]
  ]);
  document.applyCellEdits(edits, 0, '선택 범위 채우기');
  assert.equal(operationCount, 1);
  await document.save();
  const filled = analyze(await fs.readFile(file)).sheets[0];
  assert.equal(filled.cells.get('D2').value, source.value);
  assert.equal(filled.cells.get('E2').value, source.value);
  assert.equal(filled.cells.get('E2').style, source.style);
  operation.undo();
  await document.save();
  const undone = analyze(await fs.readFile(file)).sheets[0];
  assert.equal(undone.cells.get('D2').value, '18000');
  assert.equal(undone.cells.has('E2'), false);
  operation.redo();
  await document.save();
  const redone = analyze(await fs.readFile(file)).sheets[0];
  assert.equal(redone.cells.get('D2').value, source.value);
  assert.equal(redone.cells.get('E2').style, source.style);
  document.dispose();
});

test('필터 선택 범위의 지우기·채우기·정렬은 숨김 행을 건너뛴다', () => {
  const rows = [['header'], ['c'], ['hidden'], ['a']].map(([value]) => [{ value, kind: 'text' }]);
  const document = { kind: 'csv', state: { rows }, drafts: new Map(), getCell: row => rows[row]?.[0] };
  const range = { startRow: 1, endRow: 3, startColumn: 0, endColumn: 0 };
  const hiddenRows = [2];

  assert.deepEqual(buildClearEdits(document, 1, 0, range, 0, hiddenRows).map(edit => edit.row), [1, 3]);
  assert.deepEqual(buildFillEdits(document, range, 'down', 0, hiddenRows).map(edit => edit.row), [3]);
  const sortEdits = buildSortEdits(document, range, 0, 'ascending', false, false, 0, hiddenRows);
  assert.deepEqual(sortEdits.map(edit => edit.row), [1, 3]);
  assert.deepEqual(sortEdits.map(edit => edit.value), ['a', 'c']);
});

test('다중 기준 정렬은 기본 기준과 보조 기준의 방향을 각각 적용한다', () => {
  const rows = [
    ['Department', 'Priority', 'Item'],
    ['Sales', '2', 'zeta'],
    ['Sales', '1', 'alpha'],
    ['IT', '3', 'delta'],
    ['IT', '1', 'charlie']
  ].map(row => row.map(value => ({ value, kind: /^\d+$/.test(value) ? 'number' : 'text' })));
  const document = { kind: 'csv', state: { rows }, getCell: (row, column) => rows[row]?.[column] };
  const range = { startRow: 0, endRow: 4, startColumn: 0, endColumn: 2 };
  const edits = buildSortEdits(document, range, [
    { column: 0, direction: 'ascending' },
    { column: 1, direction: 'descending' }
  ], undefined, true, false);
  const result = rows.map(row => row.map(cell => cell.value));
  for (const edit of edits) result[edit.row][edit.column] = edit.value;
  assert.deepEqual(result.slice(1), [
    ['IT', '3', 'delta'], ['IT', '1', 'charlie'],
    ['Sales', '2', 'zeta'], ['Sales', '1', 'alpha']
  ]);
  assert.throws(() => buildSortEdits(document, range, [
    { column: 0, direction: 'ascending' }, { column: 0, direction: 'descending' }
  ], undefined, true, false), /정렬 기준이 유효하지 않습니다/);
});

test('채우기 핸들은 선택한 값·서식 패턴을 원본 범위 바깥으로 반복한다', () => {
  const cells = new Map([
    ['0:0', { value: 'one', style: 4, styleProperties: { bold: true } }],
    ['1:0', { value: 'two', style: 7, styleProperties: { italic: true } }]
  ]);
  const document = { kind: 'xlsx', getCell: (row, column) => cells.get(row + ':' + column) };
  const edits = buildFillRangeEdits(document,
    { startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 },
    { startRow: 0, endRow: 3, startColumn: 0, endColumn: 0 });
  assert.deepEqual(edits.map(edit => [edit.row, edit.value, edit.styleIndex, edit.styleProperties]), [
    [2, 'one', 4, { bold: true }], [3, 'two', 7, { italic: true }]
  ]);
  assert.throws(() => buildFillRangeEdits(document,
    { startRow: 1, endRow: 2, startColumn: 1, endColumn: 2 },
    { startRow: 0, endRow: 1, startColumn: 1, endColumn: 2 }), /원본 선택 범위를 포함/);

  const pattern = new Map([
    ['1:1', { value: 'A' }], ['1:2', { value: 'B' }],
    ['2:1', { value: 'C' }], ['2:2', { value: 'D' }]
  ]);
  const expanded = buildFillRangeEdits({ kind: 'csv', getCell: (row, column) => pattern.get(row + ':' + column) },
    { startRow: 1, endRow: 2, startColumn: 1, endColumn: 2 },
    { startRow: 0, endRow: 3, startColumn: 0, endColumn: 3 });
  assert.equal(expanded.find(edit => edit.row === 0 && edit.column === 0).value, 'D', 'upward and leftward fill repeats the source pattern');
  assert.equal(expanded.find(edit => edit.row === 3 && edit.column === 3).value, 'A');
  const numbers = new Map([['0:0', { value: '1' }], ['0:1', { value: '3' }]]);
  const numericDocument = { kind: 'csv', getCell: (row, column) => numbers.get(row + ':' + column) };
  const numericFill = buildFillRangeEdits(numericDocument,
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 4 });
  assert.deepEqual(numericFill.map(edit => edit.value), ['5', '7', '9'], 'a numeric seed pair continues its arithmetic series');
  const decimals = new Map([['1:0', { value: '1.5' }], ['2:0', { value: '2.5' }]]);
  const verticalFill = buildFillRangeEdits({ kind: 'csv', getCell: (row, column) => decimals.get(row + ':' + column) },
    { startRow: 1, endRow: 2, startColumn: 0, endColumn: 0 },
    { startRow: 0, endRow: 4, startColumn: 0, endColumn: 0 });
  assert.deepEqual(verticalFill.map(edit => [edit.row, edit.value]), [[0, '0.5'], [3, '3.5'], [4, '4.5']]);
  const leadingZeros = new Map([['0:0', { value: '001' }], ['0:1', { value: '002' }]]);
  const textFill = buildFillRangeEdits({ kind: 'csv', getCell: (row, column) => leadingZeros.get(row + ':' + column) },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 3 });
  assert.deepEqual(textFill.map(edit => edit.value), ['001', '002'], 'leading-zero text remains a repeated pattern');
  const labels = new Map([['0:0', { value: '항목1' }], ['0:1', { value: '항목2' }]]);
  const labelFill = buildFillRangeEdits({ kind: 'csv', getCell: (row, column) => labels.get(row + ':' + column) },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 3 });
  assert.deepEqual(labelFill.map(edit => edit.value), ['항목3', '항목4']);
  const weekdayFill = buildFillRangeEdits({ kind: 'csv', getCell: (row, column) => row === 0 && column === 0 ? { value: 'monday' } : undefined },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 3 });
  assert.deepEqual(weekdayFill.map(edit => edit.value), ['tuesday', 'wednesday', 'thursday'], 'a recognized weekday continues the built-in series and preserves lowercase');
  const weekdayPair = new Map([['0:0', { value: 'Monday' }], ['0:1', { value: 'Tuesday' }]]);
  const weekdayPairFill = buildFillRangeEdits({ kind: 'csv', getCell: (row, column) => weekdayPair.get(row + ':' + column) },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 4 });
  assert.deepEqual(weekdayPairFill.map(edit => edit.value), ['Wednesday', 'Thursday', 'Friday']);
  const reverseWeekdayPair = new Map([['0:1', { value: 'Tuesday' }], ['0:2', { value: 'Monday' }]]);
  const reverseWeekdayFill = buildFillRangeEdits({ kind: 'csv', getCell: (row, column) => reverseWeekdayPair.get(row + ':' + column) },
    { startRow: 0, endRow: 0, startColumn: 1, endColumn: 2 },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 3 });
  assert.deepEqual(reverseWeekdayFill.map(edit => [edit.column, edit.value]), [[0, 'Wednesday'], [3, 'Sunday']]);
  const monthListFill = buildFillRangeEdits({ kind: 'csv', getCell: (row, column) => row === 0 && column === 0 ? { value: 'Dec' } : undefined },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 });
  assert.deepEqual(monthListFill.map(edit => edit.value), ['Jan', 'Feb'], 'abbreviated months wrap at year end');
  const koreanWeekdayFill = buildFillRangeEdits({ kind: 'csv', getCell: (row, column) => row === 0 && column === 0 ? { value: '월요일' } : undefined },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 3 });
  assert.deepEqual(koreanWeekdayFill.map(edit => edit.value), ['화요일', '수요일', '목요일']);
  const koreanWeekdayPair = new Map([['0:0', { value: '월요일' }], ['0:1', { value: '화요일' }]]);
  const koreanWeekdayPairFill = buildFillRangeEdits({ kind: 'csv', getCell: (row, column) => koreanWeekdayPair.get(row + ':' + column) },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 3 });
  assert.deepEqual(koreanWeekdayPairFill.map(edit => edit.value), ['수요일', '목요일']);
  const koreanMonthFill = buildFillRangeEdits({ kind: 'csv', getCell: (row, column) => row === 0 && column === 0 ? { value: '12월' } : undefined },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 });
  assert.deepEqual(koreanMonthFill.map(edit => edit.value), ['1월', '2월']);
  const dates = new Map([['0:0', { value: '2026-01-30' }], ['0:1', { value: '2026-01-31' }]]);
  const dateFill = buildFillRangeEdits({ kind: 'csv', getCell: (row, column) => dates.get(row + ':' + column) },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 3 });
  assert.deepEqual(dateFill.map(edit => edit.value), ['2026-02-01', '2026-02-02']);
  const oneDateFill = buildFillRangeEdits({ kind: 'csv', getCell: (row, column) => row === 0 && column === 0 ? { value: '2026-02-28' } : null },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 });
  assert.deepEqual(oneDateFill.map(edit => edit.value), ['2026-03-01', '2026-03-02'], 'a single ISO date continues one day at a time');
  const koreanDateFill = buildFillRangeEdits({ kind: 'csv', getCell: (row, column) => row === 0 && column === 0 ? { value: '2026.10.15' } : null },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 });
  assert.deepEqual(koreanDateFill.map(edit => edit.value), ['2026.10.16', '2026.10.17'], 'a single Korean-style date continues while preserving its separator');
  const compactKoreanDateFill = buildFillRangeEdits({ kind: 'csv', getCell: (row, column) => row === 0 && column === 0 ? { value: '2026.1.2' } : null },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 });
  assert.deepEqual(compactKoreanDateFill.map(edit => edit.value), ['2026.1.3', '2026.1.4'], 'unpadded month and day formatting remains unpadded');
  const slashDateFill = buildFillRangeEdits({ kind: 'csv', getCell: (row, column) => row === 0 ? [{ value: '2026/10/30' }, { value: '2026/10/31' }][column] : null },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 3 });
  assert.deepEqual(slashDateFill.map(edit => edit.value), ['2026/11/01', '2026/11/02'], 'slash dates continue across a month boundary with their original format');
  const oneExcelDateFill = buildFillRangeEdits({ kind: 'xlsx', getCell: (row, column) => row === 0 && column === 0 ? { value: '46000', kind: 'number', numFmtId: 14, formatCode: 'm/d/yy', style: 1 } : null },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 });
  assert.deepEqual(oneExcelDateFill.map(edit => [edit.value, edit.styleIndex]), [['46001', 1], ['46002', 1]], 'an XLSX date serial increments and keeps its date style');
  const formulaCells = new Map([
    ['0:0', { value: '2', formula: 'A1+$B$1', kind: 'number', style: '2', styleProperties: { bold: true } }],
    ['1:0', null], ['2:0', { value: '4', formula: 'A3+$B$1', kind: 'number', style: '2' }]
  ]);
  const formulaFill = buildFillRangeEdits({ kind: 'xlsx', getCell: (row, column) => formulaCells.get(row + ':' + column) },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    { startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 });
  assert.deepEqual(formulaFill.map(edit => [edit.row, edit.formula, edit.value, edit.styleIndex]), [
    [1, 'A2+$B$1', '', '2'], [2, 'A3+$B$1', '', '2']
  ], 'formula fill shifts relative references, keeps absolute references and carries formatting');
  const overwriteFormula = buildFillRangeEdits({ kind: 'xlsx', getCell: (row, column) => row === 0 ? { value: 'plain' } : { value: '2', formula: 'A1', kind: 'number' } },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    { startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 });
  assert.equal(overwriteFormula[0].formula, null, 'plain values can replace normal formulas in an XLSX fill target');
  assert.throws(() => buildFillRangeEdits({ kind: 'xlsx', getCell: (row, column) => row === 0 ? { value: 'plain' } : { value: '2', formula: 'A1', formulaType: 'shared', kind: 'number' } },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    { startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 }), /공유·배열 수식/);
  const monthEnds = new Map([['0:0', { value: '2026-01-31' }], ['0:1', { value: '2026-02-28' }]]);
  const monthFill = buildFillRangeEdits({ kind: 'csv', getCell: (row, column) => monthEnds.get(row + ':' + column) },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 3 });
  assert.deepEqual(monthFill.map(edit => edit.value), ['2026-03-31', '2026-04-30'], 'month-end seeds continue on calendar month boundaries');
  const annual = new Map([['0:0', { value: '2025-05-15' }], ['0:1', { value: '2026-05-15' }]]);
  const annualFill = buildFillRangeEdits({ kind: 'csv', getCell: (row, column) => annual.get(row + ':' + column) },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 3 });
  assert.deepEqual(annualFill.map(edit => edit.value), ['2027-05-15', '2028-05-15']);
  const protectedTarget = { kind: 'csv', getCell: (row, column) => row === 1 && column === 0 ? { value: 'source' } : row === 2 && column === 0 ? { formula: 'A1+1' } : null };
  assert.throws(() => buildFillRangeEdits(protectedTarget,
    { startRow: 1, endRow: 1, startColumn: 0, endColumn: 0 },
    { startRow: 1, endRow: 2, startColumn: 0, endColumn: 0 }), /수식 또는 일반 오류 셀이 포함/);
});

test('XLSX 숫자 채우기는 계열 값을 숫자로 저장하고 한 번에 undo·redo한다', async t => {
  const bytes = minimalXlsx('', '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>3</v></c></row>');
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-fill-series-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'series.xlsx');
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  let operation, operationCount = 0;
  document.onDidChange(change => { operation = change; operationCount++; });
  const edits = buildFillRangeEdits(document,
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 3 });
  document.applyCellEdits(edits, 0, '선택 범위 채우기');
  assert.equal(operationCount, 1);
  await document.save();
  let sheet = analyze(await fs.readFile(file)).sheets[0];
  assert.equal(sheet.cells.get('C1').value, '5');
  assert.equal(sheet.cells.get('C1').kind, 'number');
  assert.equal(sheet.cells.get('D1').value, '7');
  operation.undo();
  await document.save();
  sheet = analyze(await fs.readFile(file)).sheets[0];
  assert.equal(sheet.cells.has('C1'), false);
  operation.redo();
  await document.save();
  sheet = analyze(await fs.readFile(file)).sheets[0];
  assert.equal(sheet.cells.get('D1').value, '7');
  document.dispose();
});

test('XLSX 키보드 범위 채우기는 수식 참조를 이동하고 저장·undo·redo한다', async t => {
  const bytes = minimalXlsx('', '<row r="1"><c r="A1"><f>B1+1</f><v>3</v></c><c r="B1"><v>2</v></c></row>');
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-fill-formula-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'formula-fill.xlsx');
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  let operation;
  document.onDidChange(change => { operation = change; });
  const edits = buildFillEdits(document, { startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 }, 'down');
  document.applyCellEdits(edits, 0, '선택 범위 채우기');
  assert.equal(document.getCell(1, 0).formula, 'B2+1');
  assert.equal(document.getCell(2, 0).formula, 'B3+1');
  operation.undo();
  assert.equal(document.getCell(1, 0).formula, null);
  assert.equal(document.getCell(1, 0).value, '');
  operation.redo();
  await document.save();
  const reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('A2').formula, 'B2+1');
  assert.equal(reopened.sheets[0].cells.get('A3').formula, 'B3+1');
  const rightEdits = buildFillEdits(document, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 }, 'right');
  document.applyCellEdits(rightEdits, 0, '선택 범위 오른쪽 채우기');
  assert.equal(document.getCell(0, 1).formula, 'C1+1');
  assert.equal(document.getCell(0, 2).formula, 'D1+1');
  operation.undo();
  assert.equal(document.getCell(0, 1).formula, null);
  operation.redo();
  await document.save();
  const filledRight = analyze(await fs.readFile(file));
  assert.equal(filledRight.sheets[0].cells.get('B1').formula, 'C1+1');
  assert.equal(filledRight.sheets[0].cells.get('C1').formula, 'D1+1');
  document.dispose();
});

test('Webview Ctrl+D와 Ctrl+R은 선택 범위 채우기 방향을 전달한다', () => {
  const { document, documentListeners, messages } = runGridWebview(renderHtml());
  const first = document.querySelector('[data-row="0"][data-column="0"]');
  const last = document.querySelector('[data-row="1"][data-column="1"]');
  first.onmousedown({ shiftKey: false, clientX: 80, clientY: 60, preventDefault() {} });
  last.onmouseenter();
  documentListeners.get('mouseup')();
  const keydown = documentListeners.get('keydown');
  keydown({ target: last, key: 'd', ctrlKey: true, preventDefault() {} });
  assert.equal(messages.at(-1).type, 'fill-selection');
  assert.equal(messages.at(-1).direction, 'down');
  keydown({ target: last, key: 'r', metaKey: true, preventDefault() {} });
  assert.equal(messages.at(-1).type, 'fill-selection');
  assert.equal(messages.at(-1).direction, 'right');
});

test('채우기 핸들을 드래그하면 선택을 미리 확장하고 단일 범위 채우기를 요청한다', () => {
  const { document, documentListeners, elements, messages } = runGridWebview(renderHtml(), { rowCount: 4, columnCount: 3 });
  const cell = document.querySelector('[data-row="0"][data-column="0"]');
  const handle = cell.children.find(child => child.className === 'fill-handle');
  assert.ok(handle, 'editable selection exposes its fill handle');
  handle.onmousedown({ clientX: 190, clientY: 55, preventDefault() {}, stopPropagation() {} });
  documentListeners.get('mousemove')({ clientX: 220, clientY: 100 });
  assert.equal(elements.get('address').value, 'B3');
  documentListeners.get('mouseup')();
  assert.deepEqual(JSON.parse(JSON.stringify(messages.filter(message => message.type === 'fill-handle').at(-1))), {
    type: 'fill-handle',
    sourceRange: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    targetRange: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
    sheetIndex: 0
  });
});

test('채우기 핸들을 원본 범위 안으로 되돌리면 채우기를 취소한다', () => {
  const { document, documentListeners, messages } = runGridWebview(renderHtml(), { rowCount: 4, columnCount: 3 });
  const handle = document.querySelector('[data-row="0"][data-column="0"]').children.find(child => child.className === 'fill-handle');
  handle.onmousedown({ clientX: 190, clientY: 55, preventDefault() {}, stopPropagation() {} });
  const move = documentListeners.get('mousemove');
  move({ clientX: 220, clientY: 100 });
  move({ clientX: 70, clientY: 40 });
  documentListeners.get('mouseup')();
  assert.equal(messages.some(message => message.type === 'fill-handle'), false);
});

test('선택한 CSV 표를 머리글 보존·숫자 오름차순으로 한 번에 정렬하고 undo한다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-sort-csv-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'sort.csv'), original = Buffer.from('Item,Qty\nzeta,10\nalpha,2\nBeta,1\nempty,\n');
  await fs.writeFile(file, original);
  const state = load(original, '.csv'), document = new JinSheetDocument(uri(file), 'csv', state, original, true, '');
  document.setInitialCsvState(state);
  let operation, changes = 0;
  document.onDidChange(change => { operation = change; changes++; });
  const range = { startRow: 2, endRow: 2, startColumn: 1, endColumn: 1 };
  const edits = buildSortEdits(document, range, 1, 'ascending', true, true);
  assert.ok(edits.some(edit => edit.column === 0), 'Expand selection includes the adjacent item column');
  document.applyCellEdits(edits, 0, '선택 범위 정렬');
  assert.equal(changes, 1);
  assert.deepEqual(document.state.rows.map(row => row.map(cell => cell.value)), [['Item', 'Qty'], ['Beta', '1'], ['alpha', '2'], ['zeta', '10'], ['empty', '']]);
  await document.save();
  assert.match((await fs.readFile(file, 'utf8')), /^Item,Qty\nBeta,1\nalpha,2\nzeta,10\nempty,\n$/);
  const descending = buildSortEdits(document, range, 1, 'descending', true, true);
  assert.deepEqual(descending.map(edit => edit.value), ['zeta', '10', 'Beta', '1']);
  operation.undo();
  assert.deepEqual(document.state.rows.map(row => row.map(cell => cell.value)), [['Item', 'Qty'], ['zeta', '10'], ['alpha', '2'], ['Beta', '1'], ['empty', '']]);
  document.dispose();
});

test('정렬 선택 영역 확장이 10,000셀 제한을 넘으면 거부해 부분 재배치를 막는다', () => {
  const document = { kind: 'csv', state: { rows: Array.from({ length: 5001 }, () => []) }, getCell: (_row, column) => ({ value: column === 0 ? 'key' : 'adjacent', kind: 'text' }) };
  assert.throws(() => buildSortEdits(document, { startRow: 0, endRow: 5000, startColumn: 0, endColumn: 0 }, 0, 'ascending', false, true), /10,000개/);
});

test('XLSX 행 정렬은 각 셀의 서식 인덱스를 행 데이터와 함께 이동한다', async () => {
  const bytes = await fs.readFile(path.join(__dirname, 'fixtures', 'JinSheet-Styled-Sample.xlsx'));
  const state = analyze(bytes);
  state.sheets[0].cells.get('A2').style = '2';
  state.sheets[0].cells.get('A3').style = '3';
  const document = new JinSheetDocument(uri('/tmp/sort-styles.xlsx'), 'xlsx', state, bytes, true, '');
  const edits = buildSortEdits(document, { startRow: 1, endRow: 2, startColumn: 0, endColumn: 3 }, 0, 'descending', false, false);
  document.applyCellEdits(edits, 0, '선택 범위 정렬');
  const sorted = analyze(document.createOutput()).sheets[0];
  assert.equal(sorted.cells.get('A2').value, 'Pears');
  assert.equal(sorted.cells.get('A2').style, '3');
  assert.equal(sorted.cells.get('A3').value, 'Apples');
  assert.equal(sorted.cells.get('A3').style, '2');
  document.dispose();
});

test('XLSX 정렬은 계산값으로 수식을 정렬하고 수식 참조·undo·redo·저장을 유지한다', async t => {
  const bytes = minimalXlsx('', '<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1*10</f><v>20</v></c></row><row r="2"><c r="A2"><v>1</v></c><c r="B2"><f>A2*10</f><v>10</v></c></row>');
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-sort-formula-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'sort-formula.xlsx');
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  let operation;
  document.onDidChange(change => { operation = change; });
  const edits = buildSortEdits(document, { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 }, 1, 'ascending', false, false);
  document.applyCellEdits(edits, 0, '선택 범위 정렬');
  assert.equal(document.getCell(0, 0).value, '1');
  assert.equal(document.getCell(0, 1).formula, 'A1*10');
  assert.equal(document.getCell(1, 0).value, '2');
  assert.equal(document.getCell(1, 1).formula, 'A2*10');
  await document.save();
  let sorted = analyze(await fs.readFile(file)).sheets[0];
  assert.equal(sorted.cells.get('A1').value, '1');
  assert.equal(sorted.cells.get('B1').formula, 'A1*10');
  operation.undo();
  assert.equal(document.getCell(0, 0).value, '2');
  assert.equal(document.getCell(0, 1).formula, 'A1*10');
  operation.redo();
  await document.save();
  sorted = analyze(await fs.readFile(file)).sheets[0];
  assert.equal(sorted.cells.get('B2').formula, 'A2*10');
  document.dispose();
});

test('Webview 정렬 버튼은 선택 범위·활성 열·머리글 옵션을 전달한다', () => {
  const { document, documentListeners, elements, messages } = runGridWebview(renderHtml());
  const first = document.querySelector('[data-row="0"][data-column="0"]');
  const last = document.querySelector('[data-row="1"][data-column="1"]');
  first.onmousedown({ shiftKey: false, clientX: 80, clientY: 60, preventDefault() {} });
  last.onmouseenter();
  documentListeners.get('mouseup')();
  document.querySelector('[data-row="1"][data-column="1"]').onclick();
  document.getElementById('sortHasHeader').checked = true;
  document.getElementById('sortExpandSelection').checked = true;
  document.getElementById('sortAscending').onclick();
  assert.equal(messages.at(-1).type, 'sort-selection');
  assert.equal(messages.at(-1).direction, 'ascending');
  assert.equal(messages.at(-1).keyColumn, 1);
  assert.equal(messages.at(-1).hasHeader, true);
  assert.equal(messages.at(-1).expandSelection, true);
  document.getElementById('sortExpandSelection').checked = false;
  document.getElementById('sortDescending').onclick();
  assert.equal(messages.at(-1).expandSelection, false);
});

test('사용자 지정 정렬은 선택한 여러 기준과 각 방향을 전달하고 범위 확장 제한을 안내한다', () => {
  const { document, documentListeners, elements, messages } = runGridWebview(renderHtml(), { rowCount: 3, columnCount: 2, language: 'en' });
  const first = document.querySelector('[data-row="0"][data-column="0"]');
  const last = document.querySelector('[data-row="1"][data-column="1"]');
  first.onmousedown({ shiftKey: false, clientX: 80, clientY: 60, preventDefault() {} });
  last.onmouseenter();
  documentListeners.get('mouseup')();
  const customSort = elements.get('toolbar').children.find(item => item.id === 'customSort');
  assert.ok(customSort, 'custom sort button is available in the toolbar');
  customSort.onclick();
  const panel = elements.get('body').children.find(item => item.id === 'sortPanel');
  const levels = panel.children[1], add = panel.children[2], actions = panel.children[5];
  assert.match(panel.children[3].children[2].textContent, /10,000 cells/);
  assert.match(panel.children[3].children[2].textContent, /turn this off/);
  assert.equal(levels.children.length, 1);
  add.onclick();
  assert.equal(levels.children.length, 2);
  const options = panel.children[3];
  options.children[0].children[0].checked = true;
  options.children[1].children[0].checked = false;
  levels.children[0].children[1].value = '0';
  levels.children[0].children[2].value = 'ascending';
  levels.children[1].children[1].value = '1';
  levels.children[1].children[2].value = 'descending';
  actions.children[0].onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1).sortKeys)), [
    { column: 0, direction: 'ascending' }, { column: 1, direction: 'descending' }
  ]);
  assert.equal(messages.at(-1).hasHeader, true);
  assert.equal(messages.at(-1).expandSelection, false);
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1).range)), { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 });
  assert.equal(panel.hidden, true);
});

test('한국어 정렬 대화상자는 확장 셀 한도와 큰 범위 대안을 안내한다', () => {
  const { document, documentListeners, elements } = runGridWebview(renderHtml(null, null, null, 'ko-KR'), { rowCount: 3, columnCount: 2 });
  const first = document.querySelector('[data-row="0"][data-column="0"]'), last = document.querySelector('[data-row="1"][data-column="1"]');
  first.onmousedown({ shiftKey: false, clientX: 80, clientY: 60, preventDefault() {} }); last.onmouseenter(); documentListeners.get('mouseup')();
  elements.get('toolbar').children.find(item => item.id === 'customSort').onclick();
  const panel = elements.get('body').children.find(item => item.id === 'sortPanel');
  assert.match(panel.children[3].children[2].textContent, /10,000셀/);
  assert.match(panel.children[3].children[2].textContent, /이 옵션을 끄세요/);
});

test('선택 범위 숫자 서식 버튼은 날짜·시간·백분율 형식을 XLSX 작업으로 요청한다', () => {
  const { elements, messages } = runGridWebview(renderHtml(), { kind: 'xlsx', readOnly: false });
  for (const [id, format] of [['formatNumber', 'number'], ['formatCurrency', 'currency'], ['formatPercent', 'percent'], ['formatDate', 'date'], ['formatTime', 'time']]) {
    assert.equal(elements.get(id).disabled, false);
    elements.get(id).onclick();
    assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'number-format', range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, format, sheetIndex: 0 });
  }
  const csv = runGridWebview(renderHtml(), { kind: 'csv', readOnly: false });
  assert.equal(csv.elements.get('formatPercent').disabled, true);
  assert.equal(csv.elements.get('applyCustomNumberFormat').disabled, true);
});

test('병합 도구 모음은 XLSX 선택 범위로 병합과 해제를 요청한다', () => {
  const { document, documentListeners, elements, messages } = runGridWebview(renderHtml(), { kind: 'xlsx', readOnly: false, rowCount: 3, columnCount: 3 });
  assert.equal(elements.get('mergeCells').disabled, true);
  assert.equal(elements.get('unmergeCells').disabled, true);
  const first = document.querySelector('[data-row="0"][data-column="0"]'), last = document.querySelector('[data-row="1"][data-column="1"]');
  first.onmousedown({ target: first, preventDefault() {}, clientX: 10, clientY: 40, shiftKey: false });
  last.onmouseenter();
  documentListeners.get('mouseup')();
  assert.equal(elements.get('mergeCells').disabled, false);
  elements.get('mergeCells').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'merge-cells', action: 'merge', range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 }, sheetIndex: 0 });
});

test('표 크기 조정 도구는 머리글부터 표에 포함할 마지막 열까지 선택한 범위를 요청한다', () => {
  const table = { name: 'Sales', range: { top: 0, bottom: 2, left: 0, right: 1 }, headerRowCount: 1, totalsRowCount: 0 };
  const { document, documentListeners, elements, messages } = runGridWebview(renderHtml(), { kind: 'xlsx', readOnly: false, rowCount: 4, columnCount: 3, tables: [table] });
  assert.equal(elements.get('resizeTable').disabled, false);
  documentListeners.get('click')({ target: elements.get('resizeTable') });
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'resize-table', tableName: 'Sales', range: { top: 0, bottom: 0, left: 0, right: 0 }, sheetIndex: 0 });
  const first = document.querySelector('[data-row="0"][data-column="0"]'), last = document.querySelector('[data-row="3"][data-column="2"]');
  first.onmousedown({ target: first, preventDefault() {}, clientX: 10, clientY: 40, shiftKey: false });
  last.onmouseenter();
  documentListeners.get('mouseup')();
  assert.equal(elements.get('resizeTable').disabled, false);
  documentListeners.get('click')({ target: elements.get('resizeTable') });
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'resize-table', tableName: 'Sales', range: { top: 0, bottom: 3, left: 0, right: 2 }, sheetIndex: 0 });
});

test('차트 삽입 도구는 선택 범위를 보내고 기존 차트 통합 문서에서는 비활성화된다', () => {
  const { document, documentListeners, elements, messages } = runGridWebview(renderHtml(), { kind: 'xlsx', readOnly: false, rowCount: 3, columnCount: 2 });
  const first = document.querySelector('[data-row="0"][data-column="0"]'), last = document.querySelector('[data-row="1"][data-column="1"]');
  first.onmousedown({ target: first, preventDefault() {}, clientX: 10, clientY: 40, shiftKey: false });
  last.onmouseenter(); documentListeners.get('mouseup')();
  assert.equal(elements.get('insertChart').disabled, false);
  elements.get('chartType').value = 'area';
  elements.get('insertChart').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'insert-chart', range: { top: 0, bottom: 1, left: 0, right: 1 }, chartType: 'area', sheetIndex: 0 });
  assert.match(renderHtml(), /<option value="scatter">Scatter \(X Y\)<\/option>/);
  elements.get('chartType').value = 'scatter'; elements.get('insertChart').onclick();
  assert.equal(messages.at(-1).chartType, 'scatter');
  const existing = runGridWebview(renderHtml(), { kind: 'xlsx', readOnly: false, chartInsertionAvailable: false });
  assert.equal(existing.elements.get('insertChart').disabled, true);
  assert.equal(existing.elements.get('chartType').disabled, true);
});

test('분산형 차트 삽입은 undo·redo·저장과 숫자 X·Y 미리보기에 반영한다', async t => {
  const rows = '<row r="1"><c r="A1" t="inlineStr"><is><t>X</t></is></c><c r="B1" t="inlineStr"><is><t>Y</t></is></c></row><row r="2"><c r="A2"><v>2</v></c><c r="B2"><v>8</v></c></row>';
  const bytes = minimalXlsx('', rows), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-scatter-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'scatter.xlsx'); await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  let operation; document.onDidChange(change => { operation = change; });
  document.applyAddChart(0, { top: 0, bottom: 1, left: 0, right: 1 }, 'scatter');
  assert.match(viewModel(document).chartPreviewHtml, /<circle cx=/);
  operation.undo(); assert.equal(document.state.sheets[0].charts.length, 0);
  operation.redo(); await document.save();
  const saved = analyze(await fs.readFile(file)).sheets[0].charts[0];
  assert.equal(saved.type, 'scatter'); assert.deepEqual(saved.series[0].xValues, [2]); assert.deepEqual(saved.series[0].values, [8]);
  document.dispose();
});

test('차트 제목·계열 이름 편집은 undo·redo와 XLSX 저장에 반영한다', async t => {
  const rows = '<row r="1"><c r="A1" t="inlineStr"><is><t>Month</t></is></c><c r="B1" t="inlineStr"><is><t>Sales</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Jan</t></is></c><c r="B2"><v>10</v></c></row>';
  const bytes = minimalXlsx('', rows), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-chart-metadata-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'chart.xlsx');
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  document.applyAddChart(0, { top: 0, bottom: 1, left: 0, right: 1 });
  let operation;
  document.onDidChange(change => { operation = change; });
  document.applyChartMetadata(0, 0, { title: 'Q2 <Net>', seriesNames: ['Net sales'] });
  assert.equal(document.state.sheets[0].charts[0].title, 'Q2 <Net>');
  assert.equal(document.state.sheets[0].charts[0].series[0].name, 'Net sales');
  assert.match(viewModel(document, 0, 'ko').chartPreviewHtml, /value="Q2 &lt;Net&gt;"/);
  operation.undo();
  assert.equal(document.state.sheets[0].charts[0].title, 'Chart');
  assert.equal(document.state.sheets[0].charts[0].series[0].name, 'Sales');
  operation.redo();
  await document.save();
  const saved = analyze(await fs.readFile(file));
  assert.equal(saved.sheets[0].charts[0].title, 'Q2 <Net>');
  assert.equal(saved.sheets[0].charts[0].series[0].name, 'Net sales');
  const lastOperation = operation;
  document.applyChartMetadata(0, 0, { title: 'Q2 <Net>', seriesNames: ['Net sales'] });
  assert.equal(document.dirty, false);
  assert.equal(operation, lastOperation);
  document.dispose();
});

test('차트 원본 변경은 저장 전 셀 초안과 함께 undo·redo하고 저장한다', async t => {
  const rows = '<row r="1"><c r="A1" t="inlineStr"><is><t>Month</t></is></c><c r="B1" t="inlineStr"><is><t>Sales</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Jan</t></is></c><c r="B2"><v>10</v></c></row>';
  const bytes = minimalXlsx('', rows), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-chart-source-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'chart.xlsx'); await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  document.applyAddChart(0, { top: 0, bottom: 1, left: 0, right: 1 });
  document.applyCellEdit(0, 2, 'Costs'); document.applyCellEdit(1, 2, '4');
  let operation; document.onDidChange(change => { operation = change; });
  document.applyChartSource(0, 0, { top: 0, bottom: 1, left: 0, right: 2 });
  assert.deepEqual(document.state.sheets[0].charts[0].series.map(item => item.values), [[10], [4]]);
  operation.undo();
  assert.equal(document.state.sheets[0].charts[0].series.length, 1);
  assert.equal(document.getCell(1, 2).value, '4');
  operation.redo(); await document.save();
  assert.equal(analyze(await fs.readFile(file)).sheets[0].charts[0].series[1].name, 'Costs');
  assert.equal(document.dirty, false); document.dispose();
});

test('차트 삭제는 저장 전 셀 편집을 보존하고 undo·redo·저장에 반영한다', async t => {
  const rows = '<row r="1"><c r="A1" t="inlineStr"><is><t>Month</t></is></c><c r="B1" t="inlineStr"><is><t>Sales</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Jan</t></is></c><c r="B2"><v>10</v></c></row>';
  const bytes = minimalXlsx('', rows), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-delete-chart-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'chart.xlsx'); await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  document.applyAddChart(0, { top: 0, bottom: 1, left: 0, right: 1 });
  document.applyCellEdit(1, 1, '20');
  let operation; document.onDidChange(change => { operation = change; });
  document.applyDeleteChart(0, 0);
  assert.equal(document.state.hasCharts, false); assert.equal(document.getCell(1, 1).value, '20');
  await document.save(); assert.equal(analyze(await fs.readFile(file)).hasCharts, false);
  operation.undo();
  assert.equal(document.state.sheets[0].charts.length, 1); assert.equal(document.getCell(1, 1).value, '20');
  await document.save(); assert.deepEqual(analyze(await fs.readFile(file)).sheets[0].charts[0].series[0].values, [20]);
  operation.redo(); await document.save();
  assert.equal(analyze(await fs.readFile(file)).hasCharts, false); assert.equal(document.dirty, false);
  document.dispose();
});

test('표 크기 조정 도구는 기존 표의 왼쪽 바깥 열을 포함한 선택도 허용한다', () => {
  const table = { name: 'Sales', range: { top: 0, bottom: 2, left: 1, right: 2 }, headerRowCount: 1, totalsRowCount: 0 };
  const { document, documentListeners, elements, messages } = runGridWebview(renderHtml(), { kind: 'xlsx', readOnly: false, rowCount: 4, columnCount: 3, tables: [table] });
  const first = document.querySelector('[data-row="0"][data-column="0"]'), last = document.querySelector('[data-row="3"][data-column="2"]');
  first.onmousedown({ target: first, preventDefault() {}, clientX: 10, clientY: 40, shiftKey: false });
  last.onmouseenter();
  documentListeners.get('mouseup')();
  assert.equal(elements.get('resizeTable').disabled, false);
  documentListeners.get('click')({ target: elements.get('resizeTable') });
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'resize-table', tableName: 'Sales', range: { top: 0, bottom: 3, left: 0, right: 2 }, sheetIndex: 0 });
});

test('사용자 지정 숫자 형식 입력은 선택 범위와 형식 코드를 XLSX 작업으로 보낸다', () => {
  const { elements, messages } = runGridWebview(renderHtml(), { kind: 'xlsx', readOnly: false });
  elements.get('customNumberFormat').value = '#,##0.000';
  elements.get('applyCustomNumberFormat').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'number-format-custom', range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, formatCode: '#,##0.000', sheetIndex: 0 });
});

test('선택 범위 굵게·기울임·밑줄·취소선 버튼은 XLSX 글꼴 서식 토글을 요청하고 CSV에서는 비활성화된다', () => {
  const { elements, messages } = runGridWebview(renderHtml(), { kind: 'xlsx', readOnly: false });
  for (const [id, property] of [['formatBold', 'bold'], ['formatItalic', 'italic'], ['formatUnderline', 'underline'], ['formatStrike', 'strike']]) {
    assert.equal(elements.get(id).disabled, false);
    elements.get(id).onclick();
    assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'font-format', range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, property, sheetIndex: 0 });
  }
  const csv = runGridWebview(renderHtml(), { kind: 'csv', readOnly: false });
  assert.equal(csv.elements.get('formatBold').disabled, true);
  assert.equal(csv.elements.get('formatItalic').disabled, true);
  assert.equal(csv.elements.get('formatUnderline').disabled, true);
  assert.equal(csv.elements.get('formatStrike').disabled, true);
  assert.equal(csv.elements.get('fontSize').disabled, true);
  assert.equal(csv.elements.get('applyFontSize').disabled, true);
  assert.equal(csv.elements.get('clearFontSize').disabled, true);
});

test('글꼴 크기 적용과 기본 크기 복원은 선택 범위를 요청한다', () => {
  const { elements, messages } = runGridWebview(renderHtml(), { kind: 'xlsx', readOnly: false });
  elements.get('fontSize').value = '18';
  elements.get('applyFontSize').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'font-size', range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, size: 18, sheetIndex: 0 });
  elements.get('clearFontSize').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'font-size', range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, size: null, sheetIndex: 0 });
});

test('글꼴 계열 적용과 기본 글꼴 복원은 XLSX 선택 범위를 요청한다', () => {
  const { elements, messages } = runGridWebview(renderHtml(), { kind: 'xlsx', readOnly: false });
  elements.set('fontFamily', new FakeElement('INPUT'));
  elements.get('fontFamily').value = 'Aptos';
  elements.get('applyFontFamily').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'font-family', range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, family: 'Aptos', sheetIndex: 0 });
  elements.get('clearFontFamily').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'font-family', range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, family: null, sheetIndex: 0 });
});

test('글꼴 색 선택과 자동 색상은 XLSX 선택 범위를 요청하고 CSV에서는 비활성화된다', () => {
  const { elements, messages } = runGridWebview(renderHtml(), { kind: 'xlsx', readOnly: false });
  elements.get('fontColor').value = '#1234AB';
  elements.get('fontColor').onchange({ target: elements.get('fontColor') });
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'font-color', range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, color: '#1234AB', sheetIndex: 0 });
  elements.get('clearFontColor').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'font-color', range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, color: null, sheetIndex: 0 });
  const csv = runGridWebview(renderHtml(), { kind: 'csv', readOnly: false });
  assert.equal(csv.elements.get('fontColor').disabled, true);
  assert.equal(csv.elements.get('clearFontColor').disabled, true);
});

test('가로·세로 맞춤 버튼은 축별 맞춤을 요청하고 CSV에서는 비활성화된다', () => {
  const { elements, messages } = runGridWebview(renderHtml(), { kind: 'xlsx', readOnly: false });
  for (const [id, axis, value] of [['alignLeft', 'horizontal', 'left'], ['alignCenter', 'horizontal', 'center'], ['alignRight', 'horizontal', 'right'], ['alignTop', 'vertical', 'top'], ['alignMiddle', 'vertical', 'center'], ['alignBottom', 'vertical', 'bottom']]) {
    elements.get(id).onclick();
    assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'alignment-format', range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, axis, value, sheetIndex: 0 });
  }
  const csv = runGridWebview(renderHtml(), { kind: 'csv', readOnly: false });
  assert.equal(csv.elements.get('alignCenter').disabled, true);
  assert.equal(csv.elements.get('alignMiddle').disabled, true);
});

test('글꼴 색 변경은 기존 속성과 작은 범위의 빈 셀을 보존한다', () => {
  const cells = new Map([['A1', { row: 0, column: 0, value: 'value', kind: 'string', style: '2', styleProperties: { bold: true, fill: '#FFEEAA' } }]]);
  const document = {
    kind: 'xlsx', state: { sheets: [{ cells }] }, drafts: new Map(), formulaDrafts: new Map(), styleDrafts: new Map(),
    getCell(row, column) { return row === 0 && column === 1 ? { row, column, value: '', kind: 'blank' } : [...cells.values()].find(cell => cell.row === row && cell.column === column) || null; }
  };
  const range = { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 };
  const edits = buildFontColorEdits(document, range, '#1234ab');
  assert.deepEqual(edits.map(edit => [edit.font, edit.styleProperties]), [
    [{ bold: true, italic: false, color: '#1234AB', underline: false, fontSize: null }, { bold: true, fill: '#FFEEAA', color: '#1234AB' }],
    [{ bold: false, italic: false, color: '#1234AB', underline: false, fontSize: null }, { color: '#1234AB' }]
  ]);
  assert.deepEqual(buildFontColorEdits(document, range, null).map(edit => edit.font.color), [null, null]);
  assert.throws(() => buildFontColorEdits(document, range, 'blue'), /글꼴 색/);
});

test('맞춤 변경은 지정한 축과 다른 셀 속성만 갱신한다', () => {
  const cells = new Map([['A1', { row: 0, column: 0, value: 'value', kind: 'string', style: '2', styleProperties: { color: '#123456', fill: '#FFEEAA', vertical: 'top' } }]]);
  const document = {
    kind: 'xlsx', state: { sheets: [{ cells }] }, drafts: new Map(), formulaDrafts: new Map(), styleDrafts: new Map(),
    getCell(row, column) { return row === 0 && column === 1 ? { row, column, value: '', kind: 'blank' } : [...cells.values()].find(cell => cell.row === row && cell.column === column) || null; }
  };
  const range = { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 };
  const edits = buildAlignmentFormatEdits(document, range, 'horizontal', 'center');
  assert.deepEqual(edits.map(edit => [edit.alignment, edit.styleProperties]), [
    [{ horizontal: 'center' }, { color: '#123456', fill: '#FFEEAA', vertical: 'top', horizontal: 'center' }],
    [{ horizontal: 'center' }, { horizontal: 'center' }]
  ]);
  assert.throws(() => buildAlignmentFormatEdits(document, range, 'vertical', 'justify'), /맞춤/);
});

test('테두리 프리셋은 XLSX 범위에만 적용하고 CSV에서는 비활성화된다', () => {
  const { elements, messages } = runGridWebview(renderHtml(), { kind: 'xlsx', readOnly: false });
  elements.get('borderPreset').value = 'outside';
  elements.get('borderLine').value = 'double';
  elements.get('borderColor').value = '#1234AB';
  elements.get('applyBorderPreset').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'border-format', range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, preset: 'outside', line: 'double', color: '#1234AB', sheetIndex: 0 });
  const csv = runGridWebview(renderHtml(), { kind: 'csv', readOnly: false });
  assert.equal(csv.elements.get('borderPreset').disabled, true);
  assert.equal(csv.elements.get('borderLine').disabled, true);
  assert.equal(csv.elements.get('borderColor').disabled, true);
  assert.equal(csv.elements.get('applyBorderPreset').disabled, true);
});

test('테두리 편집은 기존 속성을 보존하고 빈 셀과 테두리 제거를 처리한다', () => {
  const cells = new Map([['A1', { row: 0, column: 0, value: 'value', kind: 'string', style: '2', styleProperties: { bold: true, borders: { left: { line: 'double', color: '#123456' } } } }]]);
  const document = { kind: 'xlsx', state: { sheets: [{ cells }] }, drafts: new Map(), formulaDrafts: new Map(), styleDrafts: new Map(), getCell(row, column) { return (row === 0 && column === 1 || row === 1 && column <= 1) ? { row, column, value: '', kind: 'blank' } : [...cells.values()].find(cell => cell.row === row && cell.column === column) || null; } };
  const range = { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 };
  const edits = buildBorderFormatEdits(document, range, 'outside');
  assert.deepEqual(edits.map(edit => edit.border), [{ left: { line: 'thin', color: '#000000' }, top: { line: 'thin', color: '#000000' }, bottom: { line: 'thin', color: '#000000' } }, { right: { line: 'thin', color: '#000000' }, top: { line: 'thin', color: '#000000' }, bottom: { line: 'thin', color: '#000000' } }]);
  assert.deepEqual(edits[0].styleProperties, { bold: true, borders: { left: { line: 'thin', color: '#000000' }, top: { line: 'thin', color: '#000000' }, bottom: { line: 'thin', color: '#000000' } } });
  assert.deepEqual(buildBorderFormatEdits(document, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, 'none')[0].border, { left: null, right: null, top: null, bottom: null });
  const custom = buildBorderFormatEdits(document, range, 'all', 0, 'double', '#1234ab');
  assert.deepEqual(custom[0].border.left, { line: 'double', color: '#1234AB' });
  assert.deepEqual(buildBorderFormatEdits(document, { ...range, endRow: 1 }, 'inside-horizontal').map(edit => edit.border), [{ bottom: { line: 'thin', color: '#000000' } }, { bottom: { line: 'thin', color: '#000000' } }]);
  assert.deepEqual(buildBorderFormatEdits(document, range, 'inside-vertical').map(edit => edit.border), [{ right: { line: 'thin', color: '#000000' } }]);
  assert.throws(() => buildBorderFormatEdits(document, range, 'inside'), /테두리/);
  assert.throws(() => buildBorderFormatEdits(document, range, 'all', 0, 'invalid', '#123456'), /테두리/);
  assert.throws(() => buildBorderFormatEdits(document, range, 'all', 0, 'thin', 'blue'), /테두리/);
});

test('채우기 색과 채우기 없음은 선택 범위를 요청하고 CSV에서는 비활성화된다', () => {
  const { elements, messages } = runGridWebview(renderHtml(), { kind: 'xlsx', readOnly: false });
  elements.get('formatFillColor').value = '#A1B2C3';
  elements.get('formatFillColor').onchange({ target: elements.get('formatFillColor') });
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'fill-format', range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, color: '#A1B2C3', sheetIndex: 0 });
  elements.get('clearFillColor').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'fill-format', range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, color: null, sheetIndex: 0 });
  const csv = runGridWebview(renderHtml(), { kind: 'csv', readOnly: false });
  assert.equal(csv.elements.get('formatFillColor').disabled, true);
  assert.equal(csv.elements.get('clearFillColor').disabled, true);
});

test('채우기 색 변경은 기존 셀·초안과 작은 범위의 빈 셀만 대상으로 한다', () => {
  const cells = new Map([
    ['A1', { row: 0, column: 0, value: 'value', kind: 'string', style: '2', styleProperties: { bold: true, borders: { left: { line: 'thin' } } } }],
    ['B1', { row: 0, column: 1, value: 'error', kind: 'error' }]
  ]);
  const document = {
    kind: 'xlsx', state: { sheets: [{ cells }] }, drafts: new Map(), formulaDrafts: new Map(), styleDrafts: new Map(),
    getCell(row, column) { return row === 0 && column === 2 ? { row, column, value: '', kind: 'blank' } : [...cells.values()].find(cell => cell.row === row && cell.column === column) || null; }
  };
  const range = { startRow: 0, endRow: 0, startColumn: 0, endColumn: 3 };
  const edits = buildFillFormatEdits(document, range, '#a1b2c3');
  assert.deepEqual(edits.map(({ row, column, fillColor, styleProperties }) => [row, column, fillColor, styleProperties]), [
    [0, 0, '#A1B2C3', { bold: true, borders: { left: { line: 'thin' } }, fill: '#A1B2C3' }],
    [0, 2, '#A1B2C3', { fill: '#A1B2C3' }]
  ]);
  assert.deepEqual(buildFillFormatEdits(document, range, null).map(edit => edit.fillColor), [null, null]);
  assert.throws(() => buildFillFormatEdits(document, range, 'red'), /채우기 서식/);
});

test('글꼴 서식은 선택 셀의 현재 상태를 토글하고 수식과 기존 속성을 유지한다', () => {
  const cell = { row: 0, column: 0, value: '0.25', kind: 'number', style: '2', styleProperties: { italic: true, fill: '#FFEEAA' }, formula: '1/4', hasFormula: true };
  const cells = new Map([['A1', cell]]);
  const document = { kind: 'xlsx', state: { sheets: [{ cells }] }, drafts: new Map(), formulaDrafts: new Map(), styleDrafts: new Map(), getCell() { return cell; } };
  const bold = buildFontFormatEdits(document, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, 'bold');
  assert.deepEqual(bold, [{ row: 0, column: 0, value: '0.25', kind: 'number', styleIndex: '2', styleProperties: { italic: true, fill: '#FFEEAA', bold: true }, font: { bold: true, italic: true, underline: false, fontSize: null }, formula: '1/4' }]);
  const underline = buildFontFormatEdits(document, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, 'underline');
  assert.equal(underline[0].styleProperties.underline, true);
  assert.equal(underline[0].font.underline, true);
  const size = buildFontFormatEdits(document, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, 'fontSize', 0, 18);
  assert.equal(size[0].styleProperties.fontSize, 18);
  assert.equal(size[0].font.fontSize, 18);
  assert.equal(buildFontFormatEdits(document, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, 'fontSize', 0, null)[0].styleProperties.fontSize, undefined);
  assert.throws(() => buildFontFormatEdits(document, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, 'fontSize', 0, 0), /글꼴/);
  const family = buildFontFormatEdits(document, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, 'fontFamily', 0, 'Aptos');
  assert.equal(family[0].styleProperties.fontFamily, 'Aptos');
  assert.equal(family[0].font.fontFamily, 'Aptos');
  assert.equal(buildFontFormatEdits(document, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, 'fontFamily', 0, null)[0].styleProperties.fontFamily, undefined);
  assert.throws(() => buildFontFormatEdits(document, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, 'fontFamily', 0, ' '), /글꼴/);
});

test('숫자 서식은 희소한 전체 열 선택에서 숫자와 일반 수식만 대상으로 한다', () => {
  const cells = new Map([
    ['A1', { row: 0, column: 0, value: '0.25', kind: 'number', address: 'A1' }],
    ['A3', { row: 2, column: 0, value: '0.5', kind: 'number', address: 'A3', formula: '1/2', hasFormula: true }],
    ['A4', { row: 3, column: 0, value: 'text', kind: 'string', address: 'A4' }],
    ['A5', { row: 4, column: 0, value: '0.75', kind: 'number', address: 'A5', formula: '3/4', formulaType: 'shared', hasFormula: true }]
  ]);
  const document = { kind: 'xlsx', state: { sheets: [{ cells }] }, drafts: new Map(), formulaDrafts: new Map(), getCell(row, column) { return [...cells.values()].find(cell => cell.row === row && cell.column === column) || null; } };
  const edits = buildNumberFormatEdits(document, { startRow: 0, endRow: 1048575, startColumn: 0, endColumn: 0 }, 9);
  assert.deepEqual(edits, [
    { row: 0, column: 0, value: '0.25', kind: 'number', numberFormatId: 9, forceNumberFormat: true },
    { row: 2, column: 0, value: '0.5', kind: 'number', numberFormatId: 9, forceNumberFormat: true, formula: '1/2' }
  ]);
});

test('XLSX 숫자 표시 형식 변경은 값·수식을 유지하고 저장 후 undo·redo한다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-number-format-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'format.xlsx');
  const files = unzipSync(minimalXlsx('', '<row r="1"><c r="A1"><v>0.25</v></c><c r="B1"><f>1/4</f><v>0.25</v></c><c r="C1" t="inlineStr"><is><t>text</t></is></c></row>'));
  files['xl/_rels/workbook.xml.rels'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/_rels/workbook.xml.rels']).replace('</Relationships>', '<Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'));
  files['[Content_Types].xml'] = new TextEncoder().encode(new TextDecoder().decode(files['[Content_Types].xml']).replace('</Types>', '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'));
  files['xl/styles.xml'] = new TextEncoder().encode('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>');
  const bytes = zipSync(files);
  await fs.writeFile(file, bytes);
  const state = require('../src/xlsx').analyze(bytes), document = new JinSheetDocument(uri(file), 'xlsx', state, bytes, true, '');
  let operation;
  document.onDidChange(edit => { operation = edit; });
  const edits = buildNumberFormatEdits(document, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 }, 9);
  assert.deepEqual(edits.map(edit => edit.row + ':' + edit.column), ['0:0', '0:1']);
  document.applyCellEdits(edits, 0, '숫자 표시 형식 변경');
  assert.equal(formatCellValue(document.getCell(0, 0), 'en'), '25%');
  assert.equal(formatCellValue(document.getCell(0, 1), 'en'), '25%', 'formula values receive the selected number format');
  assert.equal(document.getCell(0, 2).value, 'text', 'text cells remain unchanged');
  await document.save();
  let reopened = require('../src/xlsx').analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('A1').numFmtId, 9);
  assert.equal(reopened.sheets[0].cells.get('B1').formula, '1/4');
  operation.undo();
  await document.save();
  reopened = require('../src/xlsx').analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('A1').numFmtId, 0);
  operation.redo();
  await document.save();
  reopened = require('../src/xlsx').analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('B1').numFmtId, 9);
  document.dispose();
});

test('XLSX 사용자 지정 숫자 형식은 유효한 패턴을 렌더링하고 저장·undo·redo한다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-custom-number-format-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'custom-number-format.xlsx');
  const files = unzipSync(minimalXlsx('', '<row r="1"><c r="A1"><v>1234.5</v></c></row>'));
  files['xl/_rels/workbook.xml.rels'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/_rels/workbook.xml.rels']).replace('</Relationships>', '<Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'));
  files['[Content_Types].xml'] = new TextEncoder().encode(new TextDecoder().decode(files['[Content_Types].xml']).replace('</Types>', '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'));
  files['xl/styles.xml'] = new TextEncoder().encode('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>');
  const bytes = zipSync(files);
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  let operation;
  document.onDidChange(edit => { operation = edit; });
  const edits = buildNumberFormatEdits(document, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, { numberFormatId: 164, numberFormatCode: '#,##0.000' });
  document.applyCellEdits(edits, 0, '사용자 지정 숫자 표시 형식 변경');
  assert.equal(formatCellValue(document.getCell(0, 0), 'en'), '1,234.500');
  await document.save();
  let reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('A1').formatCode, '#,##0.000');
  assert.equal(reopened.sheets[0].cells.get('A1').numFmtId, 164);
  operation.undo();
  await document.save();
  reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('A1').numFmtId, 0);
  operation.redo();
  await document.save();
  reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('A1').formatCode, '#,##0.000');
  document.dispose();
  assert.equal(formatCellValue({ value: '1234.5', kind: 'number', numFmtId: 164, formatCode: '₩#,##0' }, 'ko'), '₩1,235');
});

test('XLSX 굵게·기울임은 기존 셀 속성과 숫자 형식을 유지하고 저장 후 undo·redo한다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-font-format-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'font-format.xlsx');
  const files = unzipSync(minimalXlsx('', '<row r="1"><c r="A1" t="inlineStr"><is><t>label</t></is></c><c r="B1"><v>0.25</v></c><c r="D1" s="1"><v>45292</v></c></row>'));
  files['xl/_rels/workbook.xml.rels'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/_rels/workbook.xml.rels']).replace('</Relationships>', '<Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'));
  files['[Content_Types].xml'] = new TextEncoder().encode(new TextDecoder().decode(files['[Content_Types].xml']).replace('</Types>', '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'));
  files['xl/styles.xml'] = new TextEncoder().encode('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font/><font><b/><i/><color rgb="FF123456"/><sz val="14"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFEEAA"/></patternFill></fill></fills><borders count="2"><border/><border><left style="thin"><color rgb="FFABCDEF"/></left></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1"/></cellXfs></styleSheet>');
  const bytes = zipSync(files);
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  const plain = document.getCell(0, 1), styled = document.getCell(0, 3);
  const blank = buildFontFormatEdits(document, { startRow: 0, endRow: 0, startColumn: 2, endColumn: 2 }, 'italic')[0];
  let operation;
  document.onDidChange(edit => { operation = edit; });
  document.applyCellEdits([
    { row: 0, column: 1, value: plain.value, kind: plain.kind, styleIndex: plain.style ?? null, styleProperties: { bold: true }, font: { bold: true, italic: false }, numberFormatId: 9, forceNumberFormat: true },
    { row: 0, column: 3, value: styled.value, kind: styled.kind, styleIndex: styled.style, styleProperties: { ...styled.styleProperties, bold: false }, font: { bold: false, italic: true }, ...(styled.formula ? { formula: styled.formula } : {}) },
    blank
  ], 0, '글꼴 서식 변경');
  assert.equal(document.getCell(0, 1).styleProperties.bold, true);
  assert.equal(document.getCell(0, 3).styleProperties.bold, false);
  await document.save();
  let reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('B1').numFmtId, 9);
  assert.equal(reopened.sheets[0].cells.get('B1').styleProperties.bold, true);
  assert.equal(reopened.sheets[0].cells.get('D1').styleProperties.bold, undefined);
  assert.equal(reopened.sheets[0].cells.get('D1').styleProperties.italic, true);
  assert.equal(reopened.sheets[0].cells.get('D1').styleProperties.fill, '#FFEEAA');
  assert.equal(reopened.sheets[0].cells.get('C1').styleProperties.italic, true, 'formatting an empty selected cell persists for future editing');
  operation.undo();
  await document.save();
  reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('B1').styleProperties?.bold, undefined);
  assert.equal(reopened.sheets[0].cells.get('B1').numFmtId, 0);
  assert.equal(reopened.sheets[0].cells.has('C1'), false);
  operation.redo();
  await document.save();
  reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('B1').styleProperties.bold, true);
  assert.equal(reopened.sheets[0].cells.get('C1').styleProperties.italic, true);
  assert.equal(reopened.sheets[0].cells.get('D1').styleProperties.fill, '#FFEEAA');
  const styledCell = document.getCell(0, 3);
  document.applyCellEdits(buildFontFormatEdits(document, { startRow: 0, endRow: 0, startColumn: 3, endColumn: 3 }, 'underline'), 0, '밑줄 적용');
  await document.save();
  reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('D1').styleProperties.underline, true);
  assert.equal(reopened.sheets[0].cells.get('D1').styleProperties.fontSize, styledCell.styleProperties.fontSize);
  document.applyCellEdits(buildFontFormatEdits(document, { startRow: 0, endRow: 0, startColumn: 3, endColumn: 3 }, 'fontSize', 0, 18), 0, '글꼴 크기 변경');
  await document.save();
  reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('D1').styleProperties.underline, true);
  assert.equal(reopened.sheets[0].cells.get('D1').styleProperties.fontSize, 18);
operation.undo();
  await document.save();
  reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('D1').styleProperties.underline, true);
  assert.equal(reopened.sheets[0].cells.get('D1').styleProperties.fontSize, styledCell.styleProperties.fontSize);
  operation.redo();
  await document.save();
  reopened = analyze(await fs.readFile(file));
assert.equal(reopened.sheets[0].cells.get('D1').styleProperties.fontSize, 18);
document.applyCellEdits(buildFontFormatEdits(document, { startRow: 0, endRow: 0, startColumn: 3, endColumn: 3 }, 'fontFamily', 0, 'Aptos & Co'), 0, '글꼴 계열 변경');
await document.save();
reopened = analyze(await fs.readFile(file));
assert.equal(reopened.sheets[0].cells.get('D1').styleProperties.fontFamily, 'Aptos & Co');
assert.match(new TextDecoder().decode(unzipSync(await fs.readFile(file))['xl/styles.xml']), /<name val="Aptos &amp; Co"\/>/);
document.applyCellEdits(buildFontFormatEdits(document, { startRow: 0, endRow: 0, startColumn: 3, endColumn: 3 }, 'fontFamily', 0, null), 0, '기본 글꼴 복원');
await document.save();
reopened = analyze(await fs.readFile(file));
assert.equal(reopened.sheets[0].cells.get('D1').styleProperties.fontFamily, undefined);
document.applyCellEdits(buildFontFormatEdits(document, { startRow: 0, endRow: 0, startColumn: 3, endColumn: 3 }, 'strike'), 0, '취소선 적용');
await document.save();
reopened = analyze(await fs.readFile(file));
assert.equal(reopened.sheets[0].cells.get('D1').styleProperties.strike, true);
document.applyCellEdits(buildFontFormatEdits(document, { startRow: 0, endRow: 0, startColumn: 3, endColumn: 3 }, 'strike'), 0, '취소선 해제');
await document.save();
reopened = analyze(await fs.readFile(file));
assert.equal(reopened.sheets[0].cells.get('D1').styleProperties.strike, undefined);
});

test('XLSX 채우기 색은 다른 서식을 보존하고 저장·재열기·undo·redo한다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-fill-format-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'fill-format.xlsx');
  const files = unzipSync(minimalXlsx('', '<row r="1"><c r="A1" t="inlineStr"><is><t>label</t></is></c><c r="B1"><v>0.25</v></c><c r="D1" s="1"><v>45292</v></c></row>'));
  files['xl/_rels/workbook.xml.rels'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/_rels/workbook.xml.rels']).replace('</Relationships>', '<Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'));
  files['[Content_Types].xml'] = new TextEncoder().encode(new TextDecoder().decode(files['[Content_Types].xml']).replace('</Types>', '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'));
  files['xl/styles.xml'] = new TextEncoder().encode('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font/><font><b/><color rgb="FF123456"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFEEAA"/></patternFill></fill></fills><borders count="2"><border/><border><left style="thin"><color rgb="FFABCDEF"/></left></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="14" fontId="1" fillId="2" borderId="1"/></cellXfs></styleSheet>');
  const bytes = zipSync(files);
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  const initialStyled = document.getCell(0, 3).styleProperties;
  const operations = [];
  document.onDidChange(edit => { operations.push(edit); });
  const edits = buildFillFormatEdits(document, { startRow: 0, endRow: 0, startColumn: 1, endColumn: 2 }, '#FFCC00');
  assert.deepEqual(edits.map(edit => edit.column), [1, 2], 'the selection includes a previously empty cell');
  document.applyCellEdits(edits, 0, '채우기 색 변경');
  await document.save();
  let reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('B1').styleProperties.fill, '#FFCC00');
  assert.equal(reopened.sheets[0].cells.get('B1').numFmtId, 0);
  assert.equal(reopened.sheets[0].cells.get('C1').styleProperties.fill, '#FFCC00');
  assert.equal(reopened.sheets[0].cells.get('C1').value, '');
  assert.equal(reopened.sheets[0].cells.get('D1').styleProperties.fill, initialStyled.fill, 'an untouched cell keeps its original fill');
  const clear = buildFillFormatEdits(document, { startRow: 0, endRow: 0, startColumn: 3, endColumn: 3 }, null);
  document.applyCellEdits(clear, 0, '채우기 색 제거');
  await document.save();
  reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('D1').styleProperties.fill, undefined);
  assert.equal(reopened.sheets[0].cells.get('D1').numFmtId, 14, 'removing fill preserves the date number format');
  assert.equal(reopened.sheets[0].cells.get('D1').styleProperties.bold, true);
  assert.deepEqual(reopened.sheets[0].cells.get('D1').styleProperties.borders, initialStyled.borders);
  operations[1].undo();
  await document.save();
  reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('D1').styleProperties.fill, '#FFEEAA');
  operations[0].undo();
  await document.save();
  reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('B1').styleProperties?.fill, undefined);
  assert.equal(reopened.sheets[0].cells.has('C1'), false);
  operations[0].redo();
  operations[1].redo();
  await document.save();
  reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('B1').styleProperties.fill, '#FFCC00');
  assert.equal(reopened.sheets[0].cells.get('C1').styleProperties.fill, '#FFCC00');
  assert.equal(reopened.sheets[0].cells.get('D1').styleProperties.fill, undefined);
  assert.equal(reopened.sheets[0].cells.get('D1').numFmtId, 14);
  document.dispose();
});

test('XLSX 글꼴 색은 기존 글꼴·숫자·채우기·테두리를 보존하고 저장·undo·redo한다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-font-color-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'font-color.xlsx');
  const files = unzipSync(minimalXlsx('', '<row r="1"><c r="A1" s="1"><v>45292</v></c></row>'));
  files['xl/_rels/workbook.xml.rels'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/_rels/workbook.xml.rels']).replace('</Relationships>', '<Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'));
  files['[Content_Types].xml'] = new TextEncoder().encode(new TextDecoder().decode(files['[Content_Types].xml']).replace('</Types>', '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'));
  files['xl/styles.xml'] = new TextEncoder().encode('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font/><font><b/><color rgb="FF112233"/><sz val="12"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFEEAA"/></patternFill></fill></fills><borders count="2"><border/><border><left style="thin"><color rgb="FFABCDEF"/></left></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="14" fontId="1" fillId="2" borderId="1"/></cellXfs></styleSheet>');
  const bytes = zipSync(files);
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  const original = document.getCell(0, 0).styleProperties, operations = [];
  document.onDidChange(edit => { operations.push(edit); });
  const apply = buildFontColorEdits(document, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 }, '#0055CC');
  document.applyCellEdits(apply, 0, '글꼴 색 변경');
  await document.save();
  let reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('A1').styleProperties.color, '#0055CC');
  assert.equal(reopened.sheets[0].cells.get('A1').styleProperties.bold, true);
  assert.equal(reopened.sheets[0].cells.get('A1').styleProperties.fill, '#FFEEAA');
  assert.deepEqual(reopened.sheets[0].cells.get('A1').styleProperties.borders, original.borders);
  assert.equal(reopened.sheets[0].cells.get('A1').numFmtId, 14);
  assert.equal(reopened.sheets[0].cells.get('B1').styleProperties.color, '#0055CC');
  document.applyCellEdits(buildFontColorEdits(document, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, null), 0, '글꼴 색 제거');
  await document.save();
  reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('A1').styleProperties.color, undefined);
  assert.equal(reopened.sheets[0].cells.get('A1').styleProperties.bold, true);
  operations[1].undo();
  await document.save();
  reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('A1').styleProperties.color, '#0055CC');
  operations[0].undo();
  await document.save();
  reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('A1').styleProperties.color, original.color);
  assert.equal(reopened.sheets[0].cells.has('B1'), false);
  operations[0].redo();
  operations[1].redo();
  await document.save();
  reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('A1').styleProperties.color, undefined);
  assert.equal(reopened.sheets[0].cells.get('B1').styleProperties.color, '#0055CC');
  document.dispose();
});

test('XLSX 가로·세로 맞춤은 다른 서식을 유지하고 빈 셀·저장·undo·redo를 지원한다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-alignment-format-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'alignment.xlsx');
  const files = unzipSync(minimalXlsx('', '<row r="1"><c r="A1" s="1"><v>45292</v></c></row>'));
  files['xl/_rels/workbook.xml.rels'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/_rels/workbook.xml.rels']).replace('</Relationships>', '<Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'));
  files['[Content_Types].xml'] = new TextEncoder().encode(new TextDecoder().decode(files['[Content_Types].xml']).replace('</Types>', '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'));
  files['xl/styles.xml'] = new TextEncoder().encode('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font/><font><b/><color rgb="FF123456"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFEEAA"/></patternFill></fill></fills><borders count="2"><border/><border><left style="thin"><color rgb="FFABCDEF"/></left></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="14" fontId="1" fillId="2" borderId="1"/></cellXfs></styleSheet>');
  const bytes = zipSync(files);
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  const original = document.getCell(0, 0).styleProperties, operations = [];
  document.onDidChange(edit => { operations.push(edit); });
  document.applyCellEdits(buildAlignmentFormatEdits(document, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 }, 'horizontal', 'center'), 0, '가로 맞춤 변경');
  await document.save();
  let reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('A1').styleProperties.horizontal, 'center');
  assert.equal(reopened.sheets[0].cells.get('B1').styleProperties.horizontal, 'center');
  assert.equal(reopened.sheets[0].cells.get('A1').numFmtId, 14);
  assert.deepEqual(reopened.sheets[0].cells.get('A1').styleProperties.borders, original.borders);
  document.applyCellEdits(buildAlignmentFormatEdits(document, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, 'vertical', 'bottom'), 0, '세로 맞춤 변경');
  await document.save();
  reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('A1').styleProperties.horizontal, 'center');
  assert.equal(reopened.sheets[0].cells.get('A1').styleProperties.vertical, 'bottom');
  assert.equal(reopened.sheets[0].cells.get('B1').styleProperties.vertical, undefined);
  operations[1].undo();
  await document.save();
  reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('A1').styleProperties.vertical, undefined);
  operations[0].undo();
  await document.save();
  reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('A1').styleProperties.horizontal, undefined);
  assert.equal(reopened.sheets[0].cells.has('B1'), false);
  operations[0].redo();
  operations[1].redo();
  await document.save();
  reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('A1').styleProperties.horizontal, 'center');
  assert.equal(reopened.sheets[0].cells.get('A1').styleProperties.vertical, 'bottom');
  assert.equal(reopened.sheets[0].cells.get('B1').styleProperties.horizontal, 'center');
  document.dispose();
});

test('CSV 셀 범위를 탭·줄바꿈 clipboard 형식으로 복사하고 한 번에 붙여넣기·되돌리기한다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-range-csv-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'range.csv'), original = Buffer.from('name,value\nalpha,one\nbeta,two\n');
  await fs.writeFile(file, original);
  const state = load(original, '.csv'), document = new JinSheetDocument(uri(file), 'csv', state, original, true, '');
  document.setInitialCsvState(state);
  let edit;
  document.onDidChange(change => { edit = change; });
  assert.equal(serializeClipboard(document, { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 }), 'name\tvalue\r\nalpha\tone');
  document.applyCellEdits([{ row: 1, column: 0, value: 'A\tA' }, { row: 1, column: 1, value: 'line 1\nline 2' }]);
  assert.equal(document.state.rows[1][0].value, 'A\tA');
  assert.equal(document.state.rows[1][1].value, 'line 1\nline 2');
  assert.equal(serializeClipboard(document, { startRow: 1, endRow: 1, startColumn: 0, endColumn: 1 }), '"A\tA"\t"line 1\nline 2"');
  edit.undo();
  assert.deepEqual(document.state.rows[1].map(cell => cell.value), ['alpha', 'one']);
  edit.redo();
  await document.save();
  const reopened = load(await fs.readFile(file), '.csv');
  assert.deepEqual(reopened.rows[1].map(cell => cell.value), ['A\tA', 'line 1\nline 2']);
  document.dispose();
});

test('필터된 선택 범위 복사는 숨겨진 행을 건너뛰고 원래 셀 주소를 보존한다', () => {
  const values = [['a', '1'], ['hidden', '2'], ['c', '3'], ['d', '4']];
  const document = { kind: 'csv', state: { rows: values }, getCell: (row, column) => ({ value: values[row][column] }) };
  const range = { startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 };
  const payload = buildCopyPayload(document, range, 0, 'en', [0, 2, 3]);
  assert.equal(payload.text, 'a\t1\r\nc\t3\r\nd\t4');
  assert.deepEqual(payload.cells.map(line => line.map(cell => cell.row)), [[0, 0], [2, 2], [3, 3]]);
  assert.throws(() => buildCopyPayload(document, range, 0, 'en', [0, 2, 2]), /표시 행 인덱스/);
  const hiddenPayload = buildCopyPayload(document, range, 0, 'en', undefined, [1]);
  assert.equal(hiddenPayload.text, payload.text);
  assert.deepEqual(hiddenPayload.cells.map(line => line.map(cell => cell.row)), [[0, 0], [2, 2], [3, 3]]);
});

test('내부 셀 복사 버퍼는 평문 클립보드와 함께 셀 서식 패턴도 보존한다', () => {
  const styles = new Map([
    ['0:0', { value: 'A', style: '2', styleProperties: { bold: true } }],
    ['0:1', { value: 'B', style: '5', styleProperties: { italic: true } }]
  ]);
  const document = { kind: 'csv', getCell: (row, column) => styles.get(row + ':' + column) };
  const payload = buildCopyPayload(document, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 });
  assert.equal(payload.text, 'A\tB', 'other applications still receive standard tab-separated text');
  assert.deepEqual(payload.cells[0].map(cell => [cell.value, cell.styleIndex, cell.styleProperties]), [
    ['A', '2', { bold: true }], ['B', '5', { italic: true }]
  ]);
  const styleMatrix = payload.cells.map(line => line.map(({ styleIndex, styleProperties, kind }) => ({ styleIndex, styleProperties, kind })));
  const edits = buildPasteEdits([['A', 'B']], 2, 0,
    { startRow: 2, endRow: 3, startColumn: 0, endColumn: 1 }, styleMatrix);
  assert.deepEqual(edits.map(edit => [edit.row, edit.column, edit.value, edit.styleIndex, edit.styleProperties]), [
    [2, 0, 'A', '2', { bold: true }], [2, 1, 'B', '5', { italic: true }],
    [3, 0, 'A', '2', { bold: true }], [3, 1, 'B', '5', { italic: true }]
  ]);
  const filled = buildPasteEdits([['A']], 0, 0, { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 }, [[{ styleIndex: '2', styleProperties: { bold: true } }]]);
  assert.deepEqual(filled.map(edit => [edit.value, edit.styleIndex]), [['A', '2'], ['A', '2'], ['A', '2'], ['A', '2']]);
});

test('XLSX 수식 붙여넣기는 목적지마다 상대 참조를 조정하고 절대 참조는 고정한다', () => {
  const formulaMatrix = [[{ formula: 'A1+$B2+C$3+$D$4', row: 1, column: 1 }]];
  const single = buildPasteEdits([['2']], 3, 3, undefined, undefined, formulaMatrix);
  assert.equal(single[0].formula, 'C3+$B4+E$3+$D$4');
  const tiled = buildPasteEdits([['2']], 3, 3,
    { startRow: 3, endRow: 4, startColumn: 3, endColumn: 4 }, undefined, formulaMatrix);
  assert.deepEqual(tiled.map(edit => edit.formula), ['C3+$B4+E$3+$D$4', 'D3+$B4+F$3+$D$4', 'C4+$B5+E$3+$D$4', 'D4+$B5+F$3+$D$4']);
  const externalValues = parseClipboard('=A1+1\r\n=SUM(B1:B3)');
  const externalFormulas = buildClipboardFormulaMatrix(null, 'xlsx', 4, 2, externalValues);
  assert.deepEqual(externalFormulas, [[{ formula: 'A1+1', row: 4, column: 2 }], [{ formula: 'SUM(B1:B3)', row: 5, column: 2 }]]);
  const arrayFormula = buildCopyPayload({ kind: 'xlsx', getCell: () => ({ value: '6', formula: 'SUM(A1:A2)', formulaType: 'array' }) },
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 });
  assert.equal(arrayFormula.cells[0][0].formula, null, 'array formulas fall back to cached values instead of losing their array semantics');
});

test('XLSX 내부 클립보드의 수식·서식 메타데이터는 원본 통합문서에서만 재사용한다', () => {
  const sourceUri = 'file:///work/source.xlsx';
  const copied = { kind: 'xlsx', sourceUri, cells: [[{ value: '42', formula: 'A1+1', row: 0, column: 0, styleIndex: '7', styleProperties: { bold: true } }]] };
  const sourceDocument = { uri: { toString: () => sourceUri } };
  const otherDocument = { uri: { toString: () => 'file:///work/other.xlsx' } };
  assert.equal(isSameWorkbookClipboard(copied, sourceDocument), true);
  assert.equal(isSameWorkbookClipboard(copied, otherDocument), false);
  const matrix = copied.cells.map(line => line.map(cell => cell.value));
  assert.equal(buildClipboardFormulaMatrix(copied, 'xlsx', 0, 1, matrix)[0][0].formula, 'A1+1', 'same-workbook paste retains the formula');
  assert.equal(buildClipboardFormulaMatrix(null, 'xlsx', 0, 1, matrix)[0][0], null, 'cross-workbook fallback pastes the cached value without foreign references');
});

test('XLSX 범위 채우기는 일반 값으로 기존 수식을 덮어쓰고 되돌릴 수 있다', async t => {
  const bytes = minimalXlsx('', '<row r="1"><c r="A1" t="inlineStr"><is><t>copied</t></is></c></row><row r="2"><c r="A2"><f>A1+1</f><v>2</v></c></row>');
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-fill-over-formula-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'fill-over-formula.xlsx');
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  let operation;
  document.onDidChange(change => { operation = change; });
  const edits = buildFillEdits(document, { startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 }, 'down');
  assert.equal(edits[0].formula, null);
  document.applyCellEdits(edits, 0, '선택 범위 채우기');
  assert.equal(document.getCell(1, 0).value, 'copied');
  assert.equal(document.getCell(1, 0).formula, null);
  await document.save();
  assert.equal(analyze(await fs.readFile(file)).sheets[0].cells.get('A2').formula, null);
  operation.undo();
  assert.equal(document.getCell(1, 0).formula, 'A1+1');
  document.dispose();
});

test('Ctrl+Enter는 활성 셀을 수식 원점으로 삼아 역방향 선택 범위에도 수식을 저장한다', async t => {
  const bytes = minimalXlsx();
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-fill-range-formula-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'fill-formula.xlsx');
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  const matrix = [['=B1+$C$1']], formulas = buildClipboardFormulaMatrix(null, document.kind, 0, 0, matrix);
  const edits = buildPasteEdits(matrix, 0, 0,
    { startRow: 1, endRow: 0, startColumn: 1, endColumn: 0 }, undefined, formulas);
  let operation;
  document.onDidChange(change => { operation = change; });
  document.applyCellEdits(edits, 0, '선택 범위 값 채우기');
  assert.equal(document.getCell(0, 0).formula, 'B1+$C$1');
  assert.equal(document.getCell(0, 1).formula, 'C1+$C$1');
  assert.equal(document.getCell(1, 0).formula, 'B2+$C$1');
  assert.equal(document.getCell(1, 1).formula, 'C2+$C$1');
  await document.save();
  const saved = analyze(await fs.readFile(file));
  assert.equal(saved.sheets[0].cells.get('A1').formula, 'B1+$C$1');
  assert.equal(saved.sheets[0].cells.get('B1').formula, 'C1+$C$1');
  assert.equal(saved.sheets[0].cells.get('A2').formula, 'B2+$C$1');
  assert.equal(saved.sheets[0].cells.get('B2').formula, 'C2+$C$1');
  operation.undo();
  assert.equal(document.getCell(0, 0).formula, null);
  operation.redo();
  await document.save();
  assert.equal(analyze(await fs.readFile(file)).sheets[0].cells.get('B2').formula, 'C2+$C$1');
  document.dispose();
});

test('XLSX 내부 복사·붙여넣기는 셀 값을 서식과 함께 저장한다', async t => {
  const bytes = await fs.readFile(path.join(__dirname, 'fixtures', 'JinSheet-Styled-Sample.xlsx'));
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-styled-clipboard-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'styled-copy.xlsx');
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  let operation, changes = 0;
  document.onDidChange(change => { operation = change; changes++; });
  const source = document.getCell(1, 2), payload = buildCopyPayload(document,
    { startRow: 1, endRow: 1, startColumn: 2, endColumn: 2 });
  const styleMatrix = payload.cells.map(line => line.map(({ styleIndex, styleProperties, kind }) => ({ styleIndex, styleProperties, kind })));
  const edits = buildPasteEdits([[source.value]], 1, 4, undefined, styleMatrix);
  document.applyCellEdits(edits);
  assert.equal(changes, 1);
  const pendingTarget = document.getCell(1, 4);
  assert.equal(pendingTarget.numFmtId, source.numFmtId, 'the unsaved style draft keeps the source number format');
  assert.equal(pendingTarget.formatCode, source.formatCode);
  assert.equal(formatCellValue(pendingTarget), formatCellValue(source), 'the unsaved pasted value is displayed using the copied number format');
  await document.save();
  let copied = analyze(await fs.readFile(file)).sheets[0].cells.get('E2');
  assert.equal(copied.value, source.value);
  assert.equal(copied.style, source.style);
  assert.deepEqual(copied.styleProperties, source.styleProperties);
  operation.undo();
  await document.save();
  assert.equal(analyze(await fs.readFile(file)).sheets[0].cells.has('E2'), false);
  operation.redo();
  await document.save();
  copied = analyze(await fs.readFile(file)).sheets[0].cells.get('E2');
  assert.equal(copied.style, source.style);
  document.dispose();
});

test('XLSX 내부 복사·붙여넣기는 숫자처럼 보이는 문자열 종류를 저장·undo·redo로 보존한다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-string-kind-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'string-kind.xlsx');
  const original = minimalXlsx('', '<row r="1"><c r="A1" t="inlineStr"><is><t>00123</t></is></c></row>');
  await fs.writeFile(file, original);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(original), original, true, '');
  let operation;
  document.onDidChange(change => { operation = change; });
  const copied = buildCopyPayload(document, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 });
  const styles = copied.cells.map(line => line.map(({ styleIndex, styleProperties, kind }) => ({ styleIndex, styleProperties, kind })));
  document.applyCellEdits(buildPasteEdits([[copied.cells[0][0].value]], 0, 1, undefined, styles));
  assert.equal(document.getCell(0, 1).value, '00123');
  assert.equal(document.getCell(0, 1).kind, 'string');
  await document.save();
  let pasted = analyze(await fs.readFile(file)).sheets[0].cells.get('B1');
  assert.equal(pasted.value, '00123');
  assert.equal(pasted.kind, 'string');
  operation.undo();
  await document.save();
  assert.equal(analyze(await fs.readFile(file)).sheets[0].cells.has('B1'), false);
  operation.redo();
  await document.save();
  pasted = analyze(await fs.readFile(file)).sheets[0].cells.get('B1');
  assert.equal(pasted.value, '00123');
  assert.equal(pasted.kind, 'string');
  document.dispose();
});

test('XLSX 내부 수식 복사·붙여넣기는 상대 참조를 옮기고 저장 후 undo·redo한다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-formula-clipboard-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'formula-copy.xlsx');
  const original = minimalXlsx('', '<row r="1"><c r="A1" t="inlineStr"><is><t>5</t></is></c><c r="B1"><f>A1+1</f><v>6</v></c></row>');
  await fs.writeFile(file, original);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(original), original, true, '');
  let operation;
  document.onDidChange(change => { operation = change; });
  const payload = buildCopyPayload(document, { startRow: 0, endRow: 0, startColumn: 1, endColumn: 1 });
  const styles = payload.cells.map(line => line.map(({ styleIndex, styleProperties }) => ({ styleIndex, styleProperties })));
  const formulas = buildClipboardFormulaMatrix(payload, 'xlsx', 0, 3, [[payload.cells[0][0].value]]);
  document.applyCellEdits(buildPasteEdits([[payload.cells[0][0].value]], 0, 3, undefined, styles, formulas));
  assert.equal(document.getCell(0, 3).formula, 'C1+1');
  assert.equal(document.getCell(0, 3).value, '', 'formula results without a current calculation are not copied from the source cache');
  await document.save();
  let saved = analyze(await fs.readFile(file));
  assert.equal(saved.sheets[0].cells.get('D1').formula, 'C1+1');
  assert.equal(saved.sheets[0].cells.get('D1').cachedValue, '');
  operation.undo();
  await document.save();
  saved = analyze(await fs.readFile(file));
  assert.equal(saved.sheets[0].cells.has('D1'), false);
  operation.redo();
  await document.save();
  saved = analyze(await fs.readFile(file));
  assert.equal(saved.sheets[0].cells.get('D1').formula, 'C1+1');
  document.dispose();
});

test('CSV 범위 붙여넣기는 행·열을 확장하고 저장 후에도 undo·redo할 수 있다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-paste-expand-csv-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'expand.csv'), original = Buffer.from('name,value\r\nalpha,1\r\n');
  await fs.writeFile(file, original);
  const state = load(original, '.csv'), document = new JinSheetDocument(uri(file), 'csv', state, original, true, '');
  document.setInitialCsvState(state);
  let operation;
  document.onDidChange(change => { operation = change; });

  assert.throws(() => document.applyCellEdits([{ row: 2, column: 2, value: 'partial' }, { row: 10000, column: 0, value: 'outside' }]), /CSV 행·열 제한/);
  assert.equal(document.dirty, false, 'an out-of-bounds paste does not partially extend the file');
  document.applyCellEdits([{ row: 0, column: 3, value: 'extra' }, { row: 2, column: 2, value: 'tail' }]);
  assert.deepEqual(document.state.rows.map(row => row.map(cell => cell.value)), [['name', 'value', '', 'extra'], ['alpha', '1'], ['', '', 'tail']]);
  operation.undo();
  assert.deepEqual(document.state.rows.map(row => row.map(cell => cell.value)), [['name', 'value'], ['alpha', '1']]);
  assert.equal(document.dirty, false);
  operation.redo();
  await document.save();
  assert.equal(await fs.readFile(file, 'utf8'), 'name,value,,extra\r\nalpha,1\r\n,,tail');

  operation.undo();
  assert.equal(document.dirty, true);
  await document.save();
  assert.deepEqual(await fs.readFile(file), original);
  operation.redo();
  await document.save();
  assert.equal(await fs.readFile(file, 'utf8'), 'name,value,,extra\r\nalpha,1\r\n,,tail');
  document.dispose();
});

test('클립보드 TSV 인용·셀 수·셀 크기 제한을 검사한다', () => {
  assert.deepEqual(parseClipboard('one\t"two\tparts"\r\n"line 1\nline 2"\t"say ""hi"""'), [['one', 'two\tparts'], ['line 1\nline 2', 'say "hi"']]);
  assert.deepEqual(parseClipboard('a"b'), [['a"b']]);
  assert.deepEqual(parseClipboard(''), []);
  assert.throws(() => parseClipboard('a\t"broken'), /닫히지 않은/);
  assert.throws(() => parseClipboard('x'.repeat(100001)), /100,000자/);
  assert.throws(() => parseClipboard('x'.repeat(1000001)), /1,000,000자/);
  assert.deepEqual(buildClipboardKindMatrix(null, 'xlsx', [['12.5', 'TRUE', '00123', 'text']]), [['number', 'boolean', 'number', 'string']]);
  assert.deepEqual(buildClipboardKindMatrix({ kind: 'csv', cells: [[{ value: '12', kind: 'blank' }]] }, 'xlsx', [['12']]), [['number']], 'CSV-sourced clipboard text is inferred instead of trusting its empty XLSX type marker');
  assert.deepEqual(buildClipboardKindMatrix({ kind: 'xlsx', cells: [[{ value: '00123', kind: 'string' }, { value: 'TRUE', kind: 'boolean' }]] }, 'xlsx', [['00123', 'TRUE']]), [['string', 'boolean']], 'an XLSX internal copy retains explicit source types');
});

test('XLSX 외부 TSV 붙여넣기는 숫자·논리값 종류를 저장·undo·redo에서 유지한다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-external-typed-paste-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'typed-paste.xlsx');
  const original = minimalXlsx('', '<row r="2"><c r="A2" t="inlineStr"><is><t>old-a</t></is></c><c r="B2" t="inlineStr"><is><t>old-b</t></is></c><c r="C2" t="inlineStr"><is><t>old-c</t></is></c></row>');
  await fs.writeFile(file, original);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(original), original, true, '');
  let operation;
  document.onDidChange(change => { operation = change; });
  const matrix = parseClipboard('12.5\tTRUE\t');
  const kinds = matrix.map(line => line.map(inferClipboardCellKind));
  assert.deepEqual(kinds, [['number', 'boolean', 'blank']]);
  document.applyCellEdits(buildPasteEdits(matrix, 1, 0, undefined, undefined, undefined, [], kinds));
  assert.equal(document.getCell(1, 0).kind, 'number');
  assert.equal(document.getCell(1, 1).kind, 'boolean');
  assert.equal(document.getCell(1, 2).kind, 'blank');
  await document.save();
  let cells = analyze(await fs.readFile(file)).sheets[0].cells;
  assert.deepEqual([cells.get('A2').value, cells.get('A2').kind, cells.get('B2').value, cells.get('B2').kind, cells.get('C2').value, cells.get('C2').kind], ['12.5', 'number', 'TRUE', 'boolean', '', 'blank']);
  operation.undo();
  await document.save();
  cells = analyze(await fs.readFile(file)).sheets[0].cells;
  assert.deepEqual([cells.get('A2').value, cells.get('A2').kind, cells.get('B2').value, cells.get('B2').kind, cells.get('C2').value, cells.get('C2').kind], ['old-a', 'string', 'old-b', 'string', 'old-c', 'string']);
  operation.redo();
  await document.save();
  cells = analyze(await fs.readFile(file)).sheets[0].cells;
  assert.deepEqual([cells.get('A2').value, cells.get('A2').kind, cells.get('B2').value, cells.get('B2').kind, cells.get('C2').value, cells.get('C2').kind], ['12.5', 'number', 'TRUE', 'boolean', '', 'blank']);
  document.dispose();
});

test('ISO 날짜 TSV 붙여넣기는 날짜 serial과 표시 서식을 저장·undo·redo한다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-external-date-paste-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'date-paste.xlsx'), bytes = await fs.readFile(path.join(__dirname, 'fixtures', 'JinSheet-Styled-Sample.xlsx'));
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  const originalStyle = document.getCell(2, 0).styleProperties;
  let operation;
  document.onDidChange(change => { operation = change; });
  assert.deepEqual(parseClipboardDate('2024-02-29'), { serial: 45351, numFmtId: 14 });
  assert.deepEqual(parseClipboardDate('2024. 2. 29.', 'ko'), { serial: 45351, numFmtId: 14 });
  assert.deepEqual(parseClipboardDate('2024년 2월 29일', 'ko'), { serial: 45351, numFmtId: 14 });
  assert.deepEqual(parseClipboardDate('2/29/2024', 'en'), { serial: 45351, numFmtId: 14 });
  const timedDate = parseClipboardDate('2024-02-29 12:30', 'en');
  assert.ok(Math.abs(timedDate.serial - (45351 + 12.5 / 24)) < 1e-10);
  assert.equal(timedDate.numFmtId, 22);
assert.equal(parseClipboardDate('2023-02-29'), null, 'invalid calendar dates stay text');
assert.equal(parseClipboardDate('1899-12-31'), null, 'dates before Excel 1900 date system stay text');
assert.deepEqual(parseClipboardTime('13:05'), { serial: 785 / 1440, numFmtId: 20 });
assert.deepEqual(parseClipboardTime('1:05 PM'), { serial: 785 / 1440, numFmtId: 20 });
assert.deepEqual(parseClipboardTime('12:30:15'), { serial: (12 * 3600 + 30 * 60 + 15) / 86400, numFmtId: 21 });
assert.deepEqual(parseClipboardTime('12:05 오전'), { serial: 5 / 1440, numFmtId: 20 });
assert.equal(parseClipboardTime('24:00'), null, 'out of range times stay text');
assert.equal(parseClipboardTime('12:60'), null, 'invalid minutes stay text');
  assert.deepEqual(parseClipboardPercentage('12.5%'), { serial: 0.125, numFmtId: 10 });
assert.deepEqual(parseClipboardPercentage('1,200%'), { serial: 12, numFmtId: 9 });
assert.equal(parseClipboardPercentage('1.2.3%'), null);
const currencyModel = { stylesPath: 'xl/styles.xml', files: { 'xl/styles.xml': new TextEncoder().encode('<styleSheet><numFmts count="0"></numFmts></styleSheet>') } };
assert.deepEqual(parseClipboardCurrency('$1,234.50', 'en', currencyModel), { serial: 1234.5, numFmtId: 164, formatCode: '"$"#,##0.00' });
assert.equal(parseClipboardCurrency('-€1.234,50', 'en', currencyModel)?.serial, -1234.5);
assert.deepEqual(parseClipboardCurrency('₩1,234', 'ko', currencyModel), { serial: 1234, numFmtId: 164, formatCode: '"₩"#,##0' });
assert.equal(parseClipboardCurrency('$12,34.50', 'en', currencyModel), null, 'malformed grouping is rejected');
assert.equal(parseClipboardCurrency('1234.50', 'en', currencyModel), null, 'unmarked numbers are not inferred as currency');
assert.equal(parseClipboardCurrency('hello USD', 'en', currencyModel), null, 'ordinary text remains text');
const matrix = [['2024-02-29']], parsedDate = parseClipboardDate(matrix[0][0]);
  matrix[0][0] = String(parsedDate.serial);
  const edits = [
    ...buildPasteEdits(matrix, 1, 4, undefined, undefined, undefined, [], [['number']], [[14]]),
    ...buildPasteEdits(matrix, 2, 0, undefined, undefined, undefined, [], [['number']], [[14]]),
    ...buildPasteEdits([[String(parseClipboardPercentage('12.5%').serial)]], 1, 5, undefined, undefined, undefined, [], [['number']], [[10]]),
    ...buildPasteEdits([[String(parseClipboardCurrency('€1.234,50', 'en', document.state).serial)]], 3, 0, undefined, undefined, undefined, [], [['number']], [[parseClipboardCurrency('€1.234,50', 'en', document.state)]]),
    ...buildPasteEdits([[String(parseClipboardTime('1:05 PM').serial)]], 4, 4, undefined, undefined, undefined, [], [['number']], [[parseClipboardTime('1:05 PM')]])
  ];
  document.applyCellEdits(edits);
  assert.equal(document.getCell(1, 4).numFmtId, 14, 'the unsaved draft displays with the default Excel date format');
  assert.equal(formatCellValue(document.getCell(1, 4)), '2/29/24');
  assert.equal(document.getCell(2, 0).numFmtId, 14);
  assert.equal(document.getCell(1, 5).numFmtId, 10);
  assert.equal(formatCellValue(document.getCell(1, 5)), '12.50%');
assert.equal(formatCellValue(document.getCell(3, 0)), '€1,234.50', 'currency formatting is shown before save');
assert.equal(document.getCell(4, 4).numFmtId, 20);
assert.equal(formatCellValue(document.getCell(4, 4)), '13:05');
  await document.save();
  let saved = analyze(await fs.readFile(file));
  let cell = saved.sheets[0].cells.get('E2');
  assert.deepEqual([cell.value, cell.kind, cell.numFmtId, cell.formatCode], ['45351', 'number', 14, 'm/d/yy']);
  assert.deepEqual([saved.sheets[0].cells.get('F2').value, saved.sheets[0].cells.get('F2').kind, saved.sheets[0].cells.get('F2').numFmtId, saved.sheets[0].cells.get('F2').formatCode], ['0.125', 'number', 10, '0.00%']);
  assert.equal(saved.sheets[0].cells.get('A4').formatCode, '"€"#,##0.00');
assert.equal(formatCellValue(saved.sheets[0].cells.get('A4')), '€1,234.50');
assert.deepEqual([saved.sheets[0].cells.get('E5').value, saved.sheets[0].cells.get('E5').numFmtId, saved.sheets[0].cells.get('E5').formatCode], [String(785 / 1440), 20, 'h:mm']);
  assert.deepEqual(saved.sheets[0].cells.get('A3').styleProperties, originalStyle, 'adding a date number format retains the existing cell appearance');
  assert.equal(saved.sheets[0].cells.get('A3').numFmtId, 14);
  operation.undo();
  await document.save();
  assert.equal(analyze(await fs.readFile(file)).sheets[0].cells.has('E2'), false);
  assert.equal(analyze(await fs.readFile(file)).sheets[0].cells.has('F2'), false);
assert.notEqual(analyze(await fs.readFile(file)).sheets[0].cells.get('A4')?.formatCode, '"€"#,##0.00', 'undo removes the pasted currency number format');
assert.notEqual(analyze(await fs.readFile(file)).sheets[0].cells.get('E5')?.numFmtId, 20, 'undo restores the prior cell time format');
  assert.equal(analyze(await fs.readFile(file)).sheets[0].cells.get('A3').value, 'Pears');
  operation.redo();
  await document.save();
  saved = analyze(await fs.readFile(file));
  cell = saved.sheets[0].cells.get('E2');
  assert.deepEqual([cell.value, cell.kind, cell.numFmtId], ['45351', 'number', 14]);
  assert.deepEqual([saved.sheets[0].cells.get('F2').value, saved.sheets[0].cells.get('F2').numFmtId], ['0.125', 10]);
assert.equal(saved.sheets[0].cells.get('A4').formatCode, '"€"#,##0.00', 'redo restores the custom currency format after save');
assert.equal(saved.sheets[0].cells.get('E5').numFmtId, 20, 'redo restores the pasted time number format');
  assert.deepEqual(saved.sheets[0].cells.get('A3').styleProperties, originalStyle);
  document.dispose();
});

test('날짜·시간 단축키 값은 XLSX 숫자 서식으로 저장되고 undo·redo된다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-datetime-shortcut-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'shortcut.xlsx'), bytes = await fs.readFile(path.join(__dirname, 'fixtures', 'JinSheet-Styled-Sample.xlsx'));
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  const now = new Date(2024, 1, 29, 13, 5, 59);
  const edits = [buildDateTimeShortcutEdit(document, 4, 0, 'date', now), buildDateTimeShortcutEdit(document, 4, 1, 'time', now)];
  let operation;
  document.onDidChange(change => { operation = change; });
  document.applyCellEdits(edits, 0, '날짜·시간 입력');
  assert.equal(document.getCell(4, 0).numFmtId, 14);
  assert.equal(formatCellValue(document.getCell(4, 0)), '2/29/24');
  assert.equal(document.getCell(4, 1).numFmtId, 20);
  assert.equal(formatCellValue(document.getCell(4, 1)), '13:05');
  await document.save();
  let saved = analyze(await fs.readFile(file));
  assert.deepEqual([saved.sheets[0].cells.get('A5').value, saved.sheets[0].cells.get('A5').numFmtId], ['45351', 14]);
  assert.deepEqual([saved.sheets[0].cells.get('B5').value, saved.sheets[0].cells.get('B5').numFmtId], [String(785 / 1440), 20]);
  operation.undo();
  await document.save();
  saved = analyze(await fs.readFile(file));
  assert.equal(saved.sheets[0].cells.has('A5'), false);
  assert.equal(saved.sheets[0].cells.has('B5'), false);
  operation.redo();
  await document.save();
  saved = analyze(await fs.readFile(file));
  assert.equal(saved.sheets[0].cells.get('A5').numFmtId, 14);
  assert.equal(saved.sheets[0].cells.get('B5').numFmtId, 20);
  document.dispose();
});

test('자동 합계는 위쪽 또는 왼쪽 연속 숫자 범위를 선택하고 필터 중에는 SUBTOTAL을 쓴다', () => {
  const rows = [
    '<row r="1"><c r="A1" t="inlineStr"><is><t>Header</t></is></c><c r="B1" t="inlineStr"><is><t>Other</t></is></c></row>',
    '<row r="2"><c r="A2"><v>2</v></c><c r="B2"><v>5</v></c></row>',
    '<row r="3"><c r="A3"><v>3</v></c><c r="B3"><v>7</v></c></row>'
  ].join('');
  const bytes = minimalXlsx('', rows), document = new JinSheetDocument(uri('/tmp/auto-sum.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  assert.equal(buildAutoSumFormula(document, 3, 0), '=SUM(A2:A3)');
  assert.equal(buildAutoSumFormula(document, 1, 2), '=SUM(A2:B2)');
  assert.equal(buildAutoSumFormula(document, 0, 0), null);
  const filteredBytes = minimalXlsx('', rows, '<autoFilter ref="A1:B3"/>');
  const filtered = new JinSheetDocument(uri('/tmp/auto-sum-filtered.xlsx'), 'xlsx', analyze(filteredBytes), filteredBytes, true, '');
  assert.equal(buildAutoSumFormula(filtered, 3, 0), '=SUBTOTAL(9,A2:A3)');
  filtered.dispose(); document.dispose();
});

test('단일 셀 붙여넣기는 선택 영역 전체를 채우고 호환되는 범위 클립보드를 붙여넣는다', () => {
  const filled = buildPasteEdits([['value']], 1, 1, { startRow: 1, endRow: 0, startColumn: 1, endColumn: 0 });
  assert.deepEqual(filled, [
    { row: 0, column: 0, value: 'value' }, { row: 0, column: 1, value: 'value' },
    { row: 1, column: 0, value: 'value' }, { row: 1, column: 1, value: 'value' }
  ]);
  assert.throws(() => buildPasteEdits([['a', 'b'], ['c', 'd']], 3, 4, { startRow: 3, endRow: 5, startColumn: 4, endColumn: 6 }), /클립보드 범위와 선택 영역의 크기가 맞지 않습니다/, 'a mismatched multi-cell selection is rejected instead of partially filled');
  const tiled = buildPasteEdits([['a', 'b'], ['c', 'd']], 0, 0, { startRow: 0, endRow: 3, startColumn: 0, endColumn: 3 });
  assert.deepEqual(tiled.map(edit => edit.value), ['a', 'b', 'a', 'b', 'c', 'd', 'c', 'd', 'a', 'b', 'a', 'b', 'c', 'd', 'c', 'd']);
  assert.deepEqual(buildPasteEdits([['value']], 0, 0, { startRow: 0, endRow: 3, startColumn: 0, endColumn: 0 }, undefined, undefined, [1, 3]).map(edit => edit.row), [0, 2], 'single-cell fill visits only visible filtered rows');
  assert.deepEqual(buildPasteEdits([['a'], ['b']], 0, 0, { startRow: 0, endRow: 3, startColumn: 0, endColumn: 0 }, undefined, undefined, [1, 3]).map(edit => [edit.row, edit.value]), [[0, 'a'], [2, 'b']], 'multi-cell paste maps copied rows onto visible target rows');
  assert.throws(() => buildPasteEdits([['a', 'b'], ['c', 'd']], 0, 0, { startRow: 0, endRow: 2, startColumn: 0, endColumn: 3 }), /클립보드 범위와 선택 영역의 크기가 맞지 않습니다/, 'an incompatible multi-cell target must not be partially overwritten');
  assert.throws(() => buildPasteEdits([['a', 'b'], ['c', 'd']], 0, 0, { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 }, undefined, undefined, [1]), /클립보드 범위와 선택 영역의 크기가 맞지 않습니다/, 'filtered visible target dimensions must also match before paste');
  assert.throws(() => buildPasteEdits([['x']], 0, 0, { startRow: 0, endRow: 150, startColumn: 0, endColumn: 150 }), /붙여넣을 수 있는 셀은 20,000개/);

  const bytes = Buffer.from('a,b\nc,d\n'), document = new JinSheetDocument(uri('/tmp/fill-selection.csv'), 'csv', load(bytes, '.csv'), bytes, true, '');
  document.setInitialCsvState(load(bytes, '.csv'));
  let operation;
  document.onDidChange(change => { operation = change; });
  document.applyCellEdits(filled);
  assert.deepEqual(document.state.rows.map(line => line.map(cell => cell.value)), [['value', 'value'], ['value', 'value']]);
  operation.undo();
  assert.deepEqual(document.state.rows.map(line => line.map(cell => cell.value)), [['a', 'b'], ['c', 'd']]);
  const tiledBytes = Buffer.from('1,2,3,4\n5,6,7,8\n9,10,11,12\n13,14,15,16\n');
  const tiledDocument = new JinSheetDocument(uri('/tmp/paste-tiled.csv'), 'csv', load(tiledBytes, '.csv'), tiledBytes, true, '');
  tiledDocument.setInitialCsvState(tiledDocument.state);
  tiledDocument.applyCellEdits(tiled);
  assert.deepEqual(tiledDocument.state.rows.map(line => line.map(cell => cell.value)), [['a', 'b', 'a', 'b'], ['c', 'd', 'c', 'd'], ['a', 'b', 'a', 'b'], ['c', 'd', 'c', 'd']]);
  tiledDocument.dispose();
  document.applyCellEdits(filled);
  document.applyCellEdits(buildPasteEdits([['']], 0, 0, { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 }), 0, '선택 범위 값 지우기');
  assert.deepEqual(document.state.rows.map(line => line.map(cell => cell.value)), [['', ''], ['', '']]);
  operation.undo();
  assert.deepEqual(document.state.rows.map(line => line.map(cell => cell.value)), [['value', 'value'], ['value', 'value']]);
  document.dispose();
});

test('XLSX 대량 이동 한도 확장은 CSV 셀 수정 한도를 넓히지 않는다', () => {
  const csvBytes = Buffer.from('value,other\n'), csv = new JinSheetDocument(uri('/tmp/csv-edit-limit.csv'), 'csv', load(csvBytes, '.csv'), csvBytes, true, '');
  csv.setInitialCsvState(load(csvBytes, '.csv'));
  const repeatedEdit = { row: 0, column: 0, value: 'changed' };
  assert.throws(() => csv.applyCellEdits(Array(20001).fill(repeatedEdit)), /셀 범위 수정 요청이 유효하지 않습니다/);
  csv.dispose();

  const xlsxBytes = minimalXlsx(), xlsx = new JinSheetDocument(uri('/tmp/xlsx-edit-limit.xlsx'), 'xlsx', analyze(xlsxBytes), xlsxBytes, true, '');
  assert.throws(() => xlsx.applyCellEdits(Array(20001).fill({ row: 0, column: 1, value: 'changed' })), /같은 셀을 중복 수정/);
  xlsx.dispose();
});

test('저장 직전 외부 변경을 감지하면 외부 내용을 덮어쓰지 않는다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-conflict-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'conflict.csv'), original = Buffer.from('a,b\n1,2\n');
  await fs.writeFile(file, original);
  const state = load(original, '.csv'), document = new JinSheetDocument(uri(file), 'csv', state, original, true, '');
  document.setInitialCsvState(state); document.applyCellEdit(1, 1, '3');
  const external = Buffer.from('external,change\n'); await fs.writeFile(file, external);
  await assert.rejects(document.save(), /외부에서 변경/);
  assert.deepEqual(await fs.readFile(file), external);
  assert.equal(document.dirty, true);
  document.dispose();
});

test('VS Code backup과 Save As가 원본 편집 상태를 보존한다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-backup-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'source.csv'), destination = path.join(directory, 'copy.csv');
  const original = Buffer.from('a,b\n1,2\n'); await fs.writeFile(file, original);
  const state = load(original, '.csv'), document = new JinSheetDocument(uri(file), 'csv', state, original, true, '');
  document.setInitialCsvState(state); document.applyCellEdit(1, 1, 'backup-value');
  const backup = await document.backup({ destination: directory });
  assert.equal((await fs.readFile(backup.id, 'utf8')), 'a,b\n1,backup-value\n');
  await document.saveAs(uri(destination));
  assert.equal((await fs.readFile(destination, 'utf8')), 'a,b\n1,backup-value\n');
  assert.equal((await fs.readFile(file, 'utf8')), 'a,b\n1,2\n');
  await backup.delete(); await assert.rejects(fs.stat(backup.id));
  document.dispose();
});

test('XLSX 편집 후 저장된 파일을 다시 검사해 값과 비타깃 파트를 확인한다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-xlsx-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'sample.xlsx'), original = minimalXlsx(); await fs.writeFile(file, original);
  const state = analyze(original), document = new JinSheetDocument(uri(file), 'xlsx', state, original, true, '');
  document.applyCellEdit(0, 0, 'saved'); await document.save();
  const after = analyze(await fs.readFile(file));
  assert.equal(after.sheets[0].cells.get('A1').value, 'saved');
  assert.equal(document.dirty, false);
  document.dispose();
});

test('XLSX에서 수식을 입력·수정·값으로 바꾸고 undo·redo한다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-xlsx-formula-edit-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'formulas.xlsx');
  const original = minimalXlsx('', '<row r="1"><c r="A1" t="inlineStr"><is><t>old</t></is></c><c r="B1"><f>A1+1</f><v>2</v></c></row>');
  await fs.writeFile(file, original);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(original), original, true, '');
  let edit;
  document.onDidChange(change => { edit = change; });
  assert.equal(document.writable, true);
  assert.equal(document.getCell(0, 1).formula, 'A1+1');
  document.applyCellEdit(0, 0, 'changed');
  await document.save();
  let reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('A1').value, 'changed');
  assert.equal(reopened.sheets[0].cells.get('B1').formula, 'A1+1');
  assert.equal(reopened.sheets[0].cells.get('B1').cachedValue, '2');
  assert.equal(reopened.formulas, true);
  assert.match(new TextDecoder().decode(reopened.files['xl/workbook.xml']), /fullCalcOnLoad="1"/);
  edit.undo();
  await document.save();
  reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('A1').value, 'old');
  assert.equal(reopened.sheets[0].cells.get('B1').formula, 'A1+1');
  edit.redo();
  await document.save();
  reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('A1').value, 'changed');
  assert.equal(reopened.sheets[0].cells.get('B1').formula, 'A1+1');

  document.applyCellEdit(0, 2, '=2+3');
  await document.save();
  const formulaBytes = await fs.readFile(file);
  reopened = analyze(formulaBytes);
  assert.equal(reopened.sheets[0].cells.get('C1').formula, '2+3');
  assert.equal(reopened.formulas, true);
  const reopenedDocument = new JinSheetDocument(uri(file), 'xlsx', reopened, formulaBytes, true, '');
  assert.equal(getRange(reopenedDocument, { rowStart: 0, columnStart: 2, rowCount: 1, columnCount: 1, requestId: 4, sheetIndex: 0 }).rows[0][0].value, '5', 'the saved formula recalculates when reopened');
  reopenedDocument.dispose();
  edit.undo();
  await document.save();
  assert.equal(analyze(await fs.readFile(file)).sheets[0].cells.has('C1'), false);
  edit.redo();
  await document.save();

  document.applyCellEdit(0, 1, 'blocked');
  await document.save();
  reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('B1').formula, null);
  assert.equal(reopened.sheets[0].cells.get('B1').value, 'blocked');
  document.dispose();
});

test('XLSX 일반 붙여넣기는 기존 수식을 값으로 바꾸고 공유 수식 그룹은 셀별로 펼쳐 저장한다', async t => {
  const bytes = minimalXlsx('', '<row r="1"><c r="A1"><f>B1+1</f><v>3</v></c><c r="B1"><v>2</v></c></row>');
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-paste-over-formula-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'paste-over-formula.xlsx');
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  let operation;
  document.onDidChange(change => { operation = change; });
  const edits = buildPasteEdits([['pasted value']], 0, 0);
  edits[0].formula = null;
  document.applyCellEdits(edits, 0, '셀 범위 붙여넣기');
  assert.equal(document.getCell(0, 0).value, 'pasted value');
  assert.equal(document.getCell(0, 0).formula, null);
  await document.save();
  assert.equal(analyze(await fs.readFile(file)).sheets[0].cells.get('A1').formula, null);
  operation.undo();
  assert.equal(document.getCell(0, 0).formula, 'B1+1');
  operation.redo();
  await document.save();
  assert.equal(analyze(await fs.readFile(file)).sheets[0].cells.get('A1').value, 'pasted value');

  const protectedBytes = minimalXlsx('', '<row r="1"><c r="A1" t="inlineStr"><is><t>old</t></is></c><c r="B1"><f t="shared" ref="B1:B2" si="0">A1+1</f><v>2</v></c></row><row r="2"><c r="B2"><f t="shared" si="0"/><v>3</v></c></row>');
  const protectedFile = path.join(directory, 'paste-over-shared-formula.xlsx');
  await fs.writeFile(protectedFile, protectedBytes);
  const protectedDocument = new JinSheetDocument(uri(protectedFile), 'xlsx', analyze(protectedBytes), protectedBytes, true, '');
  let protectedOperation;
  protectedDocument.onDidChange(change => { protectedOperation = change; });
  protectedDocument.applyCellEdits([{ row: 0, column: 0, value: 'must apply' }, { row: 0, column: 1, value: 'replacement', formula: null }]);
  assert.equal(protectedDocument.getCell(0, 0).value, 'must apply');
  assert.equal(protectedDocument.getCell(0, 1).value, 'replacement');
  assert.equal(protectedDocument.getCell(0, 1).formula, null);
  assert.equal(protectedDocument.getCell(1, 1).formula, 'A2+1');
  protectedOperation.undo();
  assert.equal(protectedDocument.getCell(0, 0).value, 'old');
  assert.equal(protectedDocument.getCell(0, 1).formulaType, 'shared');
  protectedOperation.redo();
  await protectedDocument.save();
  const reopenedSharedPaste = analyze(await fs.readFile(protectedFile));
  assert.equal(reopenedSharedPaste.sheets[0].cells.get('A1').value, 'must apply');
  assert.equal(reopenedSharedPaste.sheets[0].cells.get('B1').formula, null);
  assert.equal(reopenedSharedPaste.sheets[0].cells.get('B2').formula, 'A2+1');
  assert.doesNotMatch(reopenedSharedPaste.sheets[0].xml, /t="shared"/);
  protectedDocument.dispose();
  document.dispose();
});

test('마지막 XLSX 수식을 값으로 바꾸거나 되돌리면 구조 변경 가능 상태가 정확히 갱신된다', () => {
  const bytes = minimalXlsx('', '<row r="1"><c r="A1"><f>1+1</f><v>2</v></c></row>');
  const document = new JinSheetDocument(uri('/tmp/structure-after-formula.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  let operation;
  document.onDidChange(change => { operation = change; });
  assert.equal(viewModel(document).formulas, true);
  document.applyCellEdit(0, 0, 'plain value');
  assert.equal(viewModel(document).formulas, false, 'a null formula draft removes the workbook formula restriction');
  operation.undo();
  assert.equal(viewModel(document).formulas, true, 'undo restores the formula restriction');
  operation.redo();
  assert.equal(viewModel(document).formulas, false, 'redo clears the formula restriction again');
  document.dispose();
});

test('XLSX 기본 수식은 다른 셀을 참조해 화면·검색·선택 요약에 계산값을 반영한다', () => {
  const rows = '<row r="1"><c r="A1"><v>2</v></c><c r="B1"><v>3</v></c><c r="C1"><f>A1+B1</f><v>0</v></c><c r="D1"><f>E1</f><v>0</v></c><c r="E1"><f>D1</f><v>0</v></c><c r="F1"><f>SUM(A:A)</f><v>0</v></c><c r="G1"><f>COUNTBLANK(A:A)</f><v>0</v></c></row>' +
    '<row r="2"><c r="A2"><f>SUM(A1:B1)</f><v>0</v></c><c r="B2"><f>A2*2</f><v>0</v></c></row>';
  const bytes = minimalXlsx('', rows), document = new JinSheetDocument(uri('/tmp/jinsheet-formula-calc.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  let range = getRange(document, { rowStart: 0, columnStart: 0, rowCount: 2, columnCount: 7, requestId: 1, sheetIndex: 0 });
  assert.equal(range.rows[0][2].value, '5');
  assert.equal(range.rows[0][2].cachedValue, 5);
  assert.equal(range.rows[0][3].value, '#CIRC!');
  assert.equal(range.rows[0][4].value, '#CIRC!');
  assert.equal(range.rows[1][0].value, '5');
  assert.equal(range.rows[1][1].value, '10');
  assert.equal(range.rows[0][5].value, '7', 'SUM accepts a sparse full-column reference');
  assert.equal(range.rows[0][6].value, '1048574', 'COUNTBLANK counts the remaining blank rows without scanning or allocating them');
  document.applyCellEdit(0, 0, '4');
  range = getRange(document, { rowStart: 0, columnStart: 0, rowCount: 2, columnCount: 7, requestId: 2, sheetIndex: 0 });
  assert.equal(range.rows[0][2].value, '7', 'dependent formula values recalculate after a cell edit');
  assert.equal(range.rows[0][5].value, '11', 'full-column formulas recalculate after an edit');
  assert.equal(range.rows[1][1].value, '14');
  assert.deepEqual(findMatches(document, '14', 0).matches.map(match => match.address), ['B2']);
  assert.deepEqual(selectionSummary(document, { startRow: 0, endRow: 1, startColumn: 0, endColumn: 4, requestId: 3, sheetIndex: 0 }), {
    type: 'selection-summary', requestId: 3, count: 7, numericCount: 5, sum: 35, average: 7
  });
  document.dispose();
});

test('XLSX 셀 안의 Alt+Enter 줄바꿈을 저장하고 다시 읽는다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-xlsx-cell-newline-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'multiline.xlsx'), original = minimalXlsx();
  await fs.writeFile(file, original);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(original), original, true, '');
  document.applyCellEdit(0, 0, 'first line\nsecond line');
  await document.save();
  assert.equal(analyze(await fs.readFile(file)).sheets[0].cells.get('A1').value, 'first line\nsecond line');
  document.dispose();
});

test('XLSX 빈 셀 입력과 값 삭제가 문서 저장·재열기에 반영된다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-xlsx-blank-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'sample.xlsx'), original = minimalXlsx();
  await fs.writeFile(file, original);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(original), original, true, '');
  document.applyCellEdit(1, 1, '12');
  assert.equal(document.getCell(1, 1).value, '12');
  await document.save();
  let reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('B2').kind, 'number');
  assert.equal(reopened.sheets[0].cells.get('B2').value, '12');
  document.applyCellEdit(0, 0, '');
  await document.save();
  reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('A1').value, '');
  document.dispose();
});

test('XLSX 구조 변경이 셀 수정, undo/redo와 저장에 반영된다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-structure-xlsx-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'structure.xlsx'), original = minimalXlsx();
  await fs.writeFile(file, original);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(original), original, true, '');
  document.applyCellEdit(0, 0, 'edited');
  document.applyStructureChange({ axis: 'row', action: 'insert', index: 0 });
  assert.equal(document.getCell(1, 0).value, 'edited');
  assert.equal(document.getCell(0, 0).value, '');
  let latestChange;
  document.onDidChange(change => { latestChange = change; });
  document.applyStructureChange({ axis: 'column', action: 'insert', index: 0 });
  latestChange.undo();
  assert.equal(document.getCell(1, 0).value, 'edited');
  latestChange.redo();
  await document.save();
  const reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('B2').value, 'edited');
  assert.equal(document.dirty, false);
  document.dispose();
});

test('XLSX 범위 붙여넣기는 원자적이며 수식 셀을 포함하면 아무 값도 바꾸지 않는다', () => {
  const original = minimalXlsx(), state = analyze(original);
  state.sheets[0].cells.set('B1', { row: 0, column: 1, address: 'B1', value: '2', kind: 'formula', formula: '1+1' });
  const document = new JinSheetDocument(uri('/tmp/range.xlsx'), 'xlsx', state, original, true, '');
  assert.throws(() => document.applyCellEdits([{ row: 0, column: 0, value: 'changed' }, { row: 0, column: 1, value: '3' }]), /수식·오류 셀/);
  assert.throws(() => document.applyCellEdit(1048576, 0, 'outside'), /XLSX 행·열 제한/);
  assert.equal(document.getCell(0, 0).value, 'old');
  assert.equal(document.dirty, false);
  document.dispose();
});

test('XLSX 다중 셀 붙여넣기는 기존 셀과 빈 셀을 저장하고 한 번에 되돌린다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-range-xlsx-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'range.xlsx'), original = minimalXlsx();
  await fs.writeFile(file, original);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(original), original, true, '');
  let edit;
  document.onDidChange(change => { edit = change; });
  document.applyCellEdits([{ row: 0, column: 0, value: 'updated' }, { row: 0, column: 1, value: 'new cell' }]);
  assert.equal(document.getCell(0, 1).value, 'new cell');
  edit.undo();
  assert.equal(document.getCell(0, 0).value, 'old');
  assert.equal(document.dirty, false);
  edit.redo();
  await document.save();
  const reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('A1').value, 'updated');
  assert.equal(reopened.sheets[0].cells.get('B1').value, 'new cell');
  document.dispose();
});

test('XLSX TSV 복사·붙여넣기는 탭·개행·따옴표 값을 저장 후에도 그대로 보존한다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-clipboard-xlsx-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'clipboard.xlsx'), original = minimalXlsx();
  await fs.writeFile(file, original);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(original), original, true, '');
  let operation;
  document.onDidChange(change => { operation = change; });
  const special = 'Line one\nLine two\t"quoted"';
  document.applyCellEdits([{ row: 0, column: 0, value: special }, { row: 0, column: 1, value: 'second' }]);
  const clipboard = serializeClipboard(document, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 });
  assert.deepEqual(parseClipboard(clipboard), [[special, 'second']]);
  document.applyCellEdits(parseClipboard(clipboard).flatMap((row, rowIndex) => row.map((value, columnIndex) => ({ row: 1 + rowIndex, column: columnIndex, value }))));
  assert.equal(document.getCell(1, 0).value, special);
  assert.equal(document.getCell(1, 1).value, 'second');
  operation.undo();
  assert.equal(document.getCell(1, 0).value, '');
  assert.equal(document.getCell(0, 0).value, special);
  operation.redo();
  await document.save();
  const reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].cells.get('A2').value, special);
  assert.equal(reopened.sheets[0].cells.get('B2').value, 'second');
  document.dispose();
});

test('저장 뒤 XLSX 셀 변경 undo·redo는 저장 기준을 갱신하고 다시 저장한다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-post-save-undo-xlsx-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'after-save.xlsx'), original = minimalXlsx();
  await fs.writeFile(file, original);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(original), original, true, '');
  let edit;
  document.onDidChange(change => { edit = change; });
  document.applyCellEdit(0, 0, 'saved');
  await document.save();
  edit.undo();
  assert.equal(document.dirty, true);
  await document.save();
  assert.equal(analyze(await fs.readFile(file)).sheets[0].cells.get('A1').value, 'old');
  edit.redo();
  assert.equal(document.dirty, true);
  await document.save();
  assert.equal(analyze(await fs.readFile(file)).sheets[0].cells.get('A1').value, 'saved');
  document.dispose();
});

test('XLSX 검색은 저장 전 새 빈 셀의 편집 초안도 찾는다', () => {
  const bytes = minimalXlsx();
  const document = new JinSheetDocument(uri('/tmp/xlsx-search-draft.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  document.applyCellEdit(1, 1, 'pending-search');
  const result = findMatches(document, 'pending-search', 0);
  assert.equal(result.total, 1);
  assert.deepEqual(result.matches, [{ row: 1, column: 1, address: 'B2' }]);
  document.dispose();
});

test('찾기와 바꾸기는 공백만으로 된 검색어도 문자 그대로 처리한다', () => {
  const bytes = Buffer.from('label,value\nx,A B\nx,AB\nx,  \n'), state = load(bytes, '.csv');
  const document = new JinSheetDocument(uri('/tmp/search-whitespace.csv'), 'csv', state, bytes, true, '');
  document.setInitialCsvState(state);
  const found = findMatches(document, ' ', 0);
  assert.equal(found.total, 2);
  assert.deepEqual(found.matches.map(match => match.address), ['B2', 'B4']);
  const replaced = buildReplaceEdits(document, ' ', '_', 0);
  assert.equal(replaced.replaced, 3);
  assert.deepEqual(replaced.edits, [
    { row: 1, column: 1, value: 'A_B' },
    { row: 3, column: 1, value: '__' }
  ]);
  document.dispose();
});

test('찾기 결과는 행 우선 페이지로 100개씩 순환한다', () => {
  const csvBytes = Buffer.from(Array.from({ length: 205 }, (_, row) => `match-${row},x`).join('\n'));
  const csvState = load(csvBytes, '.csv');
  const csv = new JinSheetDocument(uri('/tmp/search-pages.csv'), 'csv', csvState, csvBytes, true, '');
  csv.setInitialCsvState(csvState);
  const first = findMatches(csv, 'match', 0);
  assert.equal(first.total, 205);
  assert.equal(first.matches.length, 100);
  assert.deepEqual(first.matches.slice(0, 2).map(match => match.row), [0, 1]);
  const second = findMatches(csv, 'match', 0, 'sheet', first.matches.at(-1));
  assert.equal(second.total, 205);
  assert.deepEqual([second.matches[0].row, second.matches.at(-1).row], [100, 199]);
  const third = findMatches(csv, 'match', 0, 'sheet', second.matches.at(-1));
  assert.deepEqual(third.matches.map(match => match.row), [200, 201, 202, 203, 204]);
  const wrapped = findMatches(csv, 'match', 0, 'sheet', third.matches.at(-1));
  assert.equal(wrapped.total, 205);
  assert.equal(wrapped.wrapped, true);
  assert.deepEqual(wrapped.matches.slice(0, 2).map(match => match.row), [0, 1]);
  csv.dispose();

  const cells = Array.from({ length: 101 }, (_, index) => {
    let column = 101 - index, letters = '';
    while (column > 0) { column--; letters = String.fromCharCode(65 + column % 26) + letters; column = Math.floor(column / 26); }
    return `<c r="${letters}1" t="inlineStr"><is><t>same</t></is></c>`;
  }).join('');
  const bytes = minimalXlsx('', `<row r="1">${cells}</row>`);
  const workbook = new JinSheetDocument(uri('/tmp/search-pages.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  const sorted = findMatches(workbook, 'same', 0);
  assert.equal(sorted.total, 101);
  assert.equal(sorted.matches[0].address, 'A1', 'page order is independent of the source XML Map insertion order');
  assert.equal(sorted.matches.at(-1).address, 'CV1');
  const next = findMatches(workbook, 'same', 0, 'sheet', sorted.matches.at(-1));
  assert.deepEqual(next.matches.map(match => match.address), ['CW1']);
  workbook.dispose();
});

test('찾기 및 바꾸기는 셀 안의 모든 일치 텍스트를 바꾸고 XLSX 저장 후에도 수식과 서식을 보존한다', async t => {
  const csvBytes = Buffer.from('name,note\nAlpha alpha,alphabet\nBeta,Alpha\n'), csvState = load(csvBytes, '.csv');
  const csv = new JinSheetDocument(uri('/tmp/replace.csv'), 'csv', csvState, csvBytes, true, '');
  csv.setInitialCsvState(csvState);
  const first = buildReplaceEdits(csv, 'alpha', 'Omega', 0, null, false);
  assert.equal(first.replaced, 2);
  assert.deepEqual(first.target, { row: 1, column: 0, address: 'A2' }, 'Replace without a prior Find uses only the first matching cell');
  const current = buildReplaceEdits(csv, 'ALPHA', '$&', 0, { row: 1, column: 0 });
  assert.equal(current.replaced, 2);
  assert.deepEqual(current.edits, [{ row: 1, column: 0, value: '$& $&' }], 'replacement text is literal and covers every match in the selected cell');
  let operation;
  csv.onDidChange(change => { operation = change; });
  csv.applyCellEdits(current.edits, 0, '찾기 및 바꾸기');
  assert.equal(csv.drafts.get('0:1:0'), '$& $&');
  operation.undo();
  assert.equal(csv.drafts.has('0:1:0'), false);
  operation.redo();
  assert.equal(csv.drafts.get('0:1:0'), '$& $&');
  const all = buildReplaceEdits(csv, 'alpha', 'Omega', 0);
  assert.equal(all.replaced, 2);
  assert.equal(all.edits.length, 2);
  csv.dispose();

  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-replace-xlsx-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'replace.xlsx');
  const rows = '<row r="1"><c r="A1" s="1" t="inlineStr"><is><t>Alpha</t></is></c><c r="B1" t="str"><f>A1</f><v>Alpha</v></c></row>';
  const bytes = minimalXlsx('', rows);
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  const style = document.getCell(0, 0, 0).style, formulas = buildReplaceEdits(document, 'alpha', 'Omega', 0);
  assert.equal(formulas.replaced, 1);
  assert.equal(formulas.skippedFormulas, 1);
  assert.deepEqual(formulas.edits, [{ row: 0, column: 0, value: 'Omega', formula: null, kind: 'string' }]);
  document.applyCellEdits(formulas.edits, 0, '찾기 및 바꾸기');
  assert.equal(document.getCell(0, 0, 0).value, 'Omega');
  assert.equal(document.getCell(0, 0, 0).style, style);
  assert.equal(document.getCell(0, 1, 0).formula, 'A1');
  await document.save();
  const saved = analyze(await fs.readFile(file));
  assert.equal(saved.sheets[0].cells.get('A1').value, 'Omega');
  assert.equal(saved.sheets[0].cells.get('A1').style, style);
  assert.equal(saved.sheets[0].cells.get('B1').formula, 'A1');
  document.dispose();
});

test('통합 문서 찾기·바꾸기는 모든 시트에 적용되고 단일 실행 취소·다시 실행으로 복구된다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-workbook-replace-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'workbook.xlsx'), bytes = twoSheetSearchXlsx();
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  const found = findMatches(document, 'red', 0, 'workbook');
  assert.equal(found.total, 2);
  assert.deepEqual(found.matches.map(match => [match.sheetIndex, match.sheetName, match.address]), [[0, 'Data', 'A1'], [1, 'Lists', 'A1']]);
  const firstCurrent = buildScopedReplaceEdits(document, 'Red', 'Green', 1, null, false, 'workbook');
  assert.deepEqual(firstCurrent.target, { row: 0, column: 0, address: 'A1', sheetIndex: 1 }, 'Replace without a prior Find starts in the active sheet');
  const selected = buildScopedReplaceEdits(document, 'Red', 'Green', 0, { row: 0, column: 0, sheetIndex: 1 }, false, 'workbook');
  assert.deepEqual(selected.target, { row: 0, column: 0, address: 'A1', sheetIndex: 1 });
  const replacement = buildScopedReplaceEdits(document, 'Red', 'Green', 0, null, true, 'workbook');
  assert.equal(replacement.replaced, 2);
  assert.equal(replacement.sheetEdits.length, 2);
  let operation, changes = 0;
  document.onDidChange(change => { operation = change; changes++; });
  document.applyCellEditsAcrossSheets(replacement.sheetEdits, '찾기 및 바꾸기');
  assert.equal(changes, 1);
  assert.equal(document.getCell(0, 0, 0).value, 'Green north');
  assert.equal(document.getCell(0, 0, 1).value, 'Green south');
  operation.undo();
  assert.equal(document.getCell(0, 0, 0).value, 'Red north');
  assert.equal(document.getCell(0, 0, 1).value, 'Red south');
  operation.redo();
  await document.save();
  const saved = analyze(await fs.readFile(file));
  assert.equal(saved.sheets[0].cells.get('A1').value, 'Green north');
  assert.equal(saved.sheets[1].cells.get('A1').value, 'Green south');
  document.dispose();
});

test('통합 문서 전체 바꾸기에서 유효성 경고가 나면 일부 시트만 바뀐 상태를 되돌린다', () => {
  const bytes = twoSheetSearchXlsx(true), document = new JinSheetDocument(uri('/tmp/workbook-replace-validation.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  const replacement = buildScopedReplaceEdits(document, 'Red', 'Green', 0, null, true, 'workbook');
  let warning, changes = 0, operation;
  document.onDidChange(change => { operation = change; changes++; });
  try { document.applyCellEditsAcrossSheets(replacement.sheetEdits, '찾기 및 바꾸기'); } catch (error) { warning = error; }
  assert.equal(warning.validationPrompt.style, 'warning');
  assert.equal(document.getCell(0, 0, 0).value, 'Red north');
  assert.equal(document.getCell(0, 0, 1).value, 'Red south');
  assert.equal(document.dirty, false, 'failed multi-sheet changes are rolled back before prompting');
  warning.retry();
  assert.equal(changes, 1, 'accepting warnings creates one history entry');
  assert.equal(document.getCell(0, 0, 0).value, 'Green north');
  assert.equal(document.getCell(0, 0, 1).value, 'Green south');
  operation.undo();
  assert.equal(document.getCell(0, 0, 0).value, 'Red north');
  assert.equal(document.getCell(0, 0, 1).value, 'Red south');
  document.dispose();
});

test('PDF는 읽기 전용 Webview 모델에 원본 바이트를 넘긴다', () => {
  const bytes = Buffer.from('%PDF-1.4\n%%EOF');
  const state = { format: 'pdf', title: 'PDF 문서' };
  const document = new JinSheetDocument(uri('/tmp/sample.pdf'), 'pdf', state, bytes, false, 'PDF 미리보기 전용입니다.');
  const model = viewModel(document);
  assert.equal(model.kind, 'pdf');
  assert.equal(model.readOnly, true);
  assert.equal(model.supported, true);
  assert.deepEqual(Buffer.from(model.sourceBytes), bytes);
  document.dispose();
});

test('Webview 범위 조회·검색은 필요한 셀만 반환하고 요청 크기를 제한한다', () => {
  const bytes = Buffer.from('name,value\nAlpha,1\nBeta,2\n');
  const state = load(bytes, '.csv'), document = new JinSheetDocument(uri('/tmp/range.csv'), 'csv', state, bytes, true, '');
  document.setInitialCsvState(state);
  assert.deepEqual(getRange(document, { requestId: 1, rowStart: 1, columnStart: 0, rowCount: 1, columnCount: 2 }).rows[0].map(cell => cell.value), ['Alpha', '1']);
  assert.throws(() => getRange(document, { requestId: 2, rowStart: 0, columnStart: 0, rowCount: 61, columnCount: 1 }), /범위/);
  assert.equal(findMatches(document, 'beta', 0).matches[0].address, 'A3');
  assert.equal(viewModel(document).rows, undefined);
  document.dispose();
});

test('Webview 범위 조회는 활성 XLSX 조건부 서식을 기본 셀 서식 위에 합친다', () => {
  const cell = { address: 'A1', row: 0, column: 0, value: '42', kind: 'number', style: '1', styleProperties: { fill: '#FFFFFF', italic: true }, numFmtId: 0, formatCode: '0' };
  const document = {
    kind: 'xlsx',
    state: { sheets: [{ name: 'Data', cells: new Map([['A1', cell]]), conditionalFormats: [{ ranges: [{ top: 0, bottom: 0, left: 0, right: 0 }], operator: 'greaterThan', values: [10], priority: 1, stopIfTrue: false, style: { fill: '#FF0000', bold: true } }] }] },
    getCell: () => cell
  };
  const result = getRange(document, { requestId: 1, rowStart: 0, columnStart: 0, rowCount: 1, columnCount: 1 });
  assert.deepEqual(result.rows[0][0].style, { fill: '#FF0000', italic: true, bold: true });
});

test('reader documents accept the initial view request without a spreadsheet sheet list', () => {
  assert.equal(isValidSheetRequest({ kind: 'office' }, 0), true);
  assert.equal(isValidSheetRequest({ kind: 'pdf' }, 0), true);
  assert.equal(isValidSheetRequest({ kind: 'office' }, 1), false);
  assert.equal(isValidSheetRequest({ kind: 'pdf' }, -1), false);
  assert.equal(isValidSheetRequest({ kind: 'csv' }, 0), true);
  assert.equal(isValidSheetRequest({ kind: 'csv' }, 1), false);
  assert.equal(isValidSheetRequest({ kind: 'xlsx', state: { sheets: [{}, {}] } }, 1), true);
  assert.equal(isValidSheetRequest({ kind: 'xlsx', state: { sheets: [{}] } }, 1), false);
});

test('시트 범위 메시지는 선택한 시트의 정수 인덱스만 통과한다', () => {
  assert.equal(matchesSelectedSheet({ sheetIndex: 1 }, 1), true);
  assert.equal(matchesSelectedSheet({ sheetIndex: 0 }, 1), false);
  assert.equal(matchesSelectedSheet({ sheetIndex: '1' }, 1), false);
  assert.equal(matchesSelectedSheet({}, 0), false);
});

test('Ctrl+방향키는 현재 데이터 블록 또는 다음 데이터 영역의 끝으로 이동한다', () => {
  const bytes = Buffer.from('head,value,tail\nleft,,right\nnext,filled,last\n');
  const document = new JinSheetDocument(uri('/tmp/navigation.csv'), 'csv', load(bytes, '.csv'), bytes, true, '');
  document.setInitialCsvState(document.state);
  assert.deepEqual(navigateCell(document, { row: 1, column: 0, direction: 'right', requestId: 1 }), { type: 'navigation', requestId: 1, row: 1, column: 2, extend: false });
  assert.deepEqual(navigateCell(document, { row: 1, column: 1, direction: 'right', requestId: 2 }), { type: 'navigation', requestId: 2, row: 1, column: 2, extend: false });
  assert.deepEqual(navigateCell(document, { row: 0, column: 0, direction: 'down', requestId: 3, extend: true }), { type: 'navigation', requestId: 3, row: 2, column: 0, extend: true });
  assert.throws(() => navigateCell(document, { row: 0, column: 1, direction: 'down', requestId: 4, sheetIndex: 2 }), /첫 번째 시트/);
  document.dispose();
});

test('Ctrl+위아래 화살표는 표시 행에서 데이터 블록을 탐색하고 숨긴 행을 건너뛴다', () => {
  const bytes = Buffer.from('top,x\nx,x\n,x\nx,x\nx,x\nnext,x\nx,x\n,x\n');
  const document = new JinSheetDocument(uri('/tmp/navigation-filtered.csv'), 'csv', load(bytes, '.csv'), bytes, true, '');
  document.setInitialCsvState(document.state);
  const visibleRows = [0, 2, 5, 7];
  assert.equal(navigateCell(document, { row: 0, column: 0, direction: 'down', requestId: 1, visibleRows }).row, 5);
  assert.equal(navigateCell(document, { row: 2, column: 0, direction: 'down', requestId: 2, visibleRows }).row, 5);
  assert.equal(navigateCell(document, { row: 7, column: 0, direction: 'up', requestId: 3, visibleRows }).row, 5);
  assert.throws(() => navigateCell(document, { row: 1, column: 0, direction: 'down', requestId: 4, visibleRows }), /현재 셀이 표시 행에 없습니다/);
  document.dispose();
});

test('Ctrl+방향키는 CSV와 희소 XLSX의 빈 구간·시트 경계를 Excel 방식으로 이동한다', () => {
  const csvBytes = Buffer.from('top,x\nnext,x\n,x\nfar,x\n,x\n');
  const csv = new JinSheetDocument(uri('/tmp/navigation-gaps.csv'), 'csv', load(csvBytes, '.csv'), csvBytes, true, '');
  csv.setInitialCsvState(csv.state);
  const moveCsv = (row, direction, requestId) => navigateCell(csv, { row, column: 0, direction, requestId });
  assert.equal(moveCsv(0, 'down', 1).row, 1, 'a populated cell advances to the end of its contiguous region');
  assert.equal(moveCsv(1, 'down', 2).row, 3, 'the next move crosses a blank gap to the next region');
  assert.equal(moveCsv(2, 'down', 3).row, 3, 'a blank cell moves to the nearest populated cell');
  assert.equal(moveCsv(3, 'down', 4).row, 4, 'the last populated region stops at the bounded CSV edge');
  assert.equal(moveCsv(3, 'up', 5).row, 0, 'moving upward reaches the first cell of the previous contiguous region');
  assert.equal(navigateCell(csv, { row: 0, column: 0, direction: 'down', requestId: 6, hiddenRows: [1, 2] }).row, 3, 'Ctrl+Down skips AutoFilter-hidden rows');
  assert.equal(navigateCell(csv, { row: 3, column: 0, direction: 'up', requestId: 7, hiddenRows: [1, 2] }).row, 0, 'Ctrl+Up skips AutoFilter-hidden rows');
  csv.dispose();

  const cells = new Map([
    ['A1', { address: 'A1', row: 0, column: 0, value: 'first', kind: 'text' }],
    ['A2', { address: 'A2', row: 1, column: 0, value: 'block', kind: 'text' }],
    ['A5', { address: 'A5', row: 4, column: 0, value: 'last', kind: 'text' }],
    ['B2', { address: 'B2', row: 1, column: 1, value: 'base', kind: 'text' }],
    ['B6', { address: 'B6', row: 5, column: 1, value: 'cleared', kind: 'text' }]
  ]);
  const drafts = new Map([['0:B3', 'draft'], ['0:B2', ''], ['0:B6', '']]);
  const xlsx = {
    kind: 'xlsx', state: { sheets: [{ cells }] }, drafts,
    getCell(row, column) {
      const address = String.fromCharCode(65 + column) + (row + 1), cell = cells.get(address);
      return { row, column, address, value: drafts.get('0:' + address) ?? cell?.value ?? '', formula: cell?.formula ?? null };
    }
  };
  const moveXlsx = (row, direction, requestId, column = 0) => navigateCell(xlsx, { row, column, direction, requestId, extend: true });
  assert.deepEqual(moveXlsx(0, 'down', 6), { type: 'navigation', requestId: 6, row: 1, column: 0, extend: true });
  assert.equal(moveXlsx(1, 'down', 7).row, 4, 'sparse XLSX cells skip an empty region to the next occupied block');
  assert.equal(moveXlsx(2, 'down', 8).row, 4, 'the nearest occupied cell is found from an empty starting cell');
  assert.equal(moveXlsx(4, 'down', 9).row, 1048575, 'the final XLSX data region continues to the Excel worksheet edge');
  assert.equal(moveXlsx(0, 'down', 10, 1).row, 2, 'uncommitted edits participate while cleared draft cells leave the region');
  assert.equal(moveXlsx(1, 'down', 11, 1).row, 2, 'the nearest uncommitted occupied cell is reached after the old cell is cleared');
  assert.equal(moveXlsx(2, 'down', 12, 1).row, 1048575, 'a cell already at the end of its block advances to the Excel worksheet edge when no later block exists');
});

test('Ctrl+방향키 탐색은 표시·숨김 행 경로에서도 병합 셀 내부를 앵커로 정규화한다', () => {
  const cells = new Map([['B2', { address: 'B2', row: 1, column: 1, value: 'covered', kind: 'text' }]]);
  const document = {
    kind: 'xlsx',
    state: { sheets: [{ cells, mergedCells: [{ top: 0, bottom: 1, left: 0, right: 1, ref: 'A1:B2' }] }] },
    drafts: new Map(),
    getCell(row, column) { const cell = cells.get(String.fromCharCode(65 + column) + (row + 1)); return { row, column, value: cell?.value || '', formula: null }; }
  };
  const expected = { type: 'navigation', requestId: 1, row: 0, column: 0, extend: false };
  assert.deepEqual(navigateCell(document, { row: 0, column: 1, direction: 'down', requestId: 1 }), expected);
  assert.deepEqual(navigateCell(document, { row: 0, column: 1, direction: 'down', requestId: 1, visibleRows: [0, 1] }), expected);
  assert.deepEqual(navigateCell(document, { row: 0, column: 1, direction: 'down', requestId: 1, hiddenRows: [] }), expected);
});

test('Ctrl+왼쪽·오른쪽은 수동 숨김 열을 건너뛴다', () => {
  const rows = '<row r="1"><c r="A1"><v>first</v></c><c r="E1"><v>last</v></c></row>';
  const bytes = minimalXlsx('', rows), document = new JinSheetDocument(uri('/tmp/navigation-hidden-columns.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  assert.equal(navigateCell(document, { row: 0, column: 0, direction: 'right', requestId: 1, hiddenColumns: [1, 2, 3] }).column, 4);
  assert.equal(navigateCell(document, { row: 0, column: 4, direction: 'left', requestId: 2, hiddenColumns: [1, 2, 3] }).column, 0);
  document.dispose();
});

test('삭제한 XLSX 수식 셀은 Ctrl+A 연속 영역과 Ctrl+방향키 이동에서 빈 셀로 취급한다', () => {
  const bytes = minimalXlsx('', '<row r="3"><c r="A3"><f>1+1</f><v>2</v></c><c r="B3"><v>neighbor</v></c></row><row r="5"><c r="A5"><v>later</v></c></row>');
  const document = new JinSheetDocument(uri('/tmp/navigation-cleared-formula.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  document.applyCellEdits(buildClearEdits(document, 2, 0, { startRow: 2, endRow: 2, startColumn: 0, endColumn: 0 }), 0, '선택 범위 값 지우기');
  assert.deepEqual(currentRegion(document, 2, 1), { top: 2, bottom: 2, left: 1, right: 1 }, 'the cleared formula no longer expands the adjacent cell region');
  assert.equal(navigateCell(document, { row: 1, column: 0, direction: 'down', requestId: 1 }).row, 4, 'Ctrl+Down skips the now-empty former formula cell and reaches the next data block');
  document.dispose();
});

test('Webview uses English by default and Korean for a Korean VS Code locale', () => {
  const english = renderHtml(null, null, null, 'en-US');
  const korean = renderHtml(null, null, null, 'ko-KR');
  assert.match(english, /<html lang="en">/);
  assert.match(english, />\+ Row</);
  assert.match(english, /Find in sheet/);
  assert.doesNotMatch(JSON.stringify(messagesFor('en')), /[가-힣]/);
  assert.match(korean, /<html lang="ko">/);
  assert.match(korean, />\+ 행</);
  assert.match(korean, /시트에서 찾기/);
  assert.match(english, /\.cell\.numeric\{text-align:right\}/);
  assert.match(english, /cell\.classList\.toggle\('numeric',data\?\.kind==='number'\)/);
  const formulaReason = '수식이 포함되어 있어 저장할 수 없습니다. 수식과 캐시값은 보기 전용입니다.';
  const unsupported = { kind: 'xlsx', reason: formulaReason, state: { sheets: [] }, uri: { fsPath: '/tmp/sample.xlsx' } };
  assert.match(viewModel(unsupported, 0, 'en').reason, /contains formulas/);
  assert.equal(viewModel(unsupported, 0, 'ko').reason, formulaReason);
  const formulaWorkbook = { kind: 'xlsx', reason: null, writable: true, dirty: false, state: { formulas: true, sheets: [{ cells: new Map(), columnWidths: {} }] }, drafts: new Map(), columnWidthDrafts: new Map(), uri: { fsPath: '/tmp/formulas.xlsx' } };
  assert.equal(viewModel(formulaWorkbook).formulas, true);
  const grid = runGridWebview(renderHtml(), { kind: 'xlsx' });
  grid.dispatch({ type: 'model', kind: 'xlsx', readOnly: false, formulas: true, rowCount: 2, columnCount: 2, sheetIndex: 0, dirty: false, sheets: [] });
  assert.match(grid.elements.get('notice').textContent, /Cached formula results are shown/);
  for (const id of ['insertRow', 'deleteRow', 'insertColumn', 'deleteColumn']) assert.equal(grid.elements.get(id).disabled, false, 'formula workbooks can start supported structure changes');
  assert.equal(grid.elements.get('sortAscending').disabled, false, 'supported formula sorting remains available');
  grid.dispatch({ type: 'model', kind: 'xlsx', readOnly: false, formulas: false, rowCount: 2, columnCount: 2, sheetIndex: 0, dirty: false, sheets: [] });
  for (const id of ['insertRow', 'deleteRow', 'insertColumn', 'deleteColumn']) assert.equal(grid.elements.get(id).disabled, false, 'formula-free XLSX keeps structure editing available');
});

test('검색은 버튼이나 Enter로 실행하고 셀 서식은 유지한 채 검색 결과를 표시한다', () => {
  const html = renderHtml();
  assert.ok(html.indexOf('id="search"') < html.indexOf('id="findButton"'), 'Find button should be placed to the right of the search input');
  assert.match(html, /<label for="formula">Value<\/label>/);
  assert.match(html, /<label for="search">Search<\/label>/);
  const style = { fill: '#f4cccc', color: '#123456', bold: true };
  const { document, elements, messages, windowListeners } = runGridWebview(html, { style });
  const input = elements.get('search'), onInput = input.listeners.get('input');
  input.value = '배';
  onInput();
  input.value = '배추';
  onInput();
  assert.equal(messages.some(message => message.type === 'search'), false, 'typing alone must not move selection or start a search');
  input.focus();
  input.listeners.get('keydown')({ key: 'Enter', preventDefault() {} });
  const request = messages.find(message => message.type === 'search');
  assert.equal(request.query, '배추');
  windowListeners.get('message')({ data: { type: 'search', requestId: request.requestId, query: request.query, total: 2, matches: [{ row: 1, column: 1 }, { row: 0, column: 1 }] } });
  const first = document.querySelector('[data-row="1"][data-column="1"]');
  assert.equal(document.activeElement, input, 'search results must not transfer keyboard input to a cell');
  assert.equal(first.classList.contains('search-match'), true);
  assert.equal(first.classList.contains('search-current'), true);
  assert.equal(first.style.backgroundColor, '#f4cccc');
  assert.equal(first.style.color, '#123456');
  assert.equal(first.style.fontWeight, '700');
  input.listeners.get('keydown')({ key: 'Enter', preventDefault() {} });
  assert.equal(document.activeElement, input);
  assert.equal(document.querySelector('[data-row="0"][data-column="1"]').classList.contains('search-current'), true);
  assert.equal(first.classList.contains('search-current'), false);
  input.value += 'x';
  onInput();
  assert.equal(messages.some(message => message.type === 'edit'), false, 'typing the next character must not edit the selected match');
  input.value = 'b';
  onInput();
  assert.equal(elements.get('clearSearch').hidden, false, 'Clear should appear when the search field has a query');
  elements.get('findButton').listeners.get('click')();
  assert.equal(messages.filter(message => message.type === 'search').at(-1).query, 'b', 'Find button should run the current query');
  const latestRequest = messages.filter(message => message.type === 'search').at(-1);
  windowListeners.get('message')({ data: { type: 'search', requestId: latestRequest.requestId, query: 'b', total: 1, matches: [{ row: 0, column: 1 }] } });
  elements.get('clearSearch').listeners.get('click')();
  assert.equal(input.value, '');
  assert.equal(elements.get('clearSearch').hidden, true);
  assert.equal(elements.get('searchCount').textContent, '');
  assert.equal(document.activeElement, input);
  assert.equal(document.querySelector('[data-row="0"][data-column="1"]').classList.contains('search-match'), false);
  input.value = 'b';
  onInput();
  input.listeners.get('keydown')({ key: 'Escape', preventDefault() {} });
  assert.equal(input.value, '');
  assert.equal(document.activeElement, input);
});

test('찾기 및 바꾸기 UI는 현재 셀과 전체 시트 요청을 분리하고 결과를 알린다', () => {
  const html = renderHtml();
  assert.ok(html.indexOf('id="findButton"') < html.indexOf('id="replaceGroup"'), 'replace controls follow the adjacent Find control');
  const { document, elements, messages, windowListeners } = runGridWebview(html);
  const search = elements.get('search'), replacement = elements.get('replaceValue');
  search.value = 'alpha'; replacement.value = 'Omega';
  search.listeners.get('input')();
  replacement.listeners.get('input')();
  assert.equal(messages.some(message => message.type === 'replace-current' || message.type === 'replace-all'), false, 'typing alone does not change cells');
  elements.get('replaceCurrentButton').listeners.get('click')();
  let request = messages.filter(message => message.type === 'replace-current').at(-1);
  assert.deepEqual(JSON.parse(JSON.stringify(request)), { type: 'replace-current', query: 'alpha', replacement: 'Omega', requestId: 1, sheetIndex: 0, scope: 'sheet' }, 'Replace starts at the first match when Find has not run');
  windowListeners.get('message')({ data: { type: 'replace-result', requestId: request.requestId, query: request.query, mode: request.type, replaced: 2, skippedFormulas: 1, target: { row: 0, column: 0 }, total: 1, matches: [{ row: 1, column: 0 }] } });
  assert.match(elements.get('replaceStatus').textContent, /2 occurrence\(s\).*1 formula cell\(s\)/);
  assert.equal(document.querySelector('[data-row="1"][data-column="0"]').classList.contains('search-current'), true, 'Replace advances to the next result');
  elements.get('replaceAllButton').listeners.get('click')();
  request = messages.filter(message => message.type === 'replace-all').at(-1);
  assert.equal(request.query, 'alpha');
  assert.equal(request.replacement, 'Omega');
  assert.equal(request.sheetIndex, 0);
  assert.equal(elements.get('replaceAllButton').disabled, true, 'replace controls wait for the host result');
  windowListeners.get('message')({ data: { type: 'replace-result', requestId: request.requestId, query: request.query, mode: request.type, replaced: 0, skippedFormulas: 0, target: null, total: 0, matches: [] } });
  assert.equal(elements.get('replaceAllButton').disabled, false);
});

test('찾기는 100개 결과 뒤 다음 페이지를 요청하고 서버 응답에서 계속 순환한다', () => {
  const { document, elements, messages, windowListeners } = runGridWebview(renderHtml());
  const search = elements.get('search');
  search.value = 'hit';
  search.listeners.get('input')();
  elements.get('findButton').listeners.get('click')();
  const firstRequest = messages.filter(message => message.type === 'search').at(-1);
  const firstPage = Array.from({ length: 100 }, (_, row) => ({ row, column: 0, address: `A${row + 1}` }));
  windowListeners.get('message')({ data: { type: 'search', requestId: firstRequest.requestId, query: 'hit', total: 101, matches: firstPage } });
  for (let index = 0; index < 100; index++) search.listeners.get('keydown')({ key: 'Enter', preventDefault() {} });
  const nextRequest = messages.filter(message => message.type === 'search').at(-1);
  assert.notEqual(nextRequest.requestId, firstRequest.requestId);
  assert.deepEqual(JSON.parse(JSON.stringify(nextRequest.after)), { sheetIndex: 0, row: 99, column: 0 });
  windowListeners.get('message')({ data: { type: 'search', requestId: nextRequest.requestId, query: 'hit', total: 101, matches: [{ row: 100, column: 0, address: 'A101' }] } });
  assert.equal(elements.get('searchCount').textContent.includes('101'), true);
  assert.equal(elements.get('search').value, 'hit');
  search.listeners.get('keydown')({ key: 'Enter', preventDefault() {} });
  const wrapRequest = messages.filter(message => message.type === 'search').at(-1);
  assert.deepEqual(JSON.parse(JSON.stringify(wrapRequest.after)), { sheetIndex: 0, row: 100, column: 0 });
  windowListeners.get('message')({ data: { type: 'search', requestId: wrapRequest.requestId, query: 'hit', total: 101, matches: firstPage.slice(0, 100), wrapped: true } });
  assert.equal(document.querySelector('[data-row="0"][data-column="0"]').classList.contains('search-current'), true);
});

test('통합 문서 찾기는 결과가 있는 시트로 이동하고 바꾸기에 대상 시트를 전달한다', () => {
  const { elements, messages, windowListeners } = runGridWebview(renderHtml(), { kind: 'xlsx' });
  const search = elements.get('search'), scope = elements.get('searchScope');
  scope.value = 'workbook'; search.value = 'alpha'; search.listeners.get('input')();
  elements.get('findButton').listeners.get('click')();
  const request = messages.filter(message => message.type === 'search').at(-1);
  assert.equal(request.scope, 'workbook');
  windowListeners.get('message')({ data: { type: 'search', requestId: request.requestId, query: 'alpha', sheetIndex: 0, total: 2, matches: [{ row: 0, column: 0, sheetIndex: 0, sheetName: 'Data' }, { row: 1, column: 1, sheetIndex: 1, sheetName: 'Lists' }] } });
  search.listeners.get('keydown')({ key: 'Enter', preventDefault() {} });
  const navigation = messages.filter(message => message.type === 'search-navigate').at(-1);
  assert.deepEqual(JSON.parse(JSON.stringify(navigation)), { type: 'search-navigate', scope: 'workbook', requestId: request.requestId, sheetIndex: 1, row: 1, column: 1 });
  elements.get('replaceCurrentButton').listeners.get('click')();
  const replace = messages.filter(message => message.type === 'replace-current').at(-1);
  assert.equal(replace.scope, 'workbook');
  assert.equal(replace.targetSheetIndex, 1);
  assert.equal(replace.row, 1);
  assert.equal(replace.column, 1);
});

test('이름 상자는 셀 주소와 범위로 이동하고 잘못된 참조는 선택을 바꾸지 않는다', () => {
  const { document, elements, documentListeners } = runGridWebview(renderHtml(), { kind: 'csv', rowCount: 6, columnCount: 4 });
  const address = elements.get('address');
  assert.equal(address.value, 'A1');
  address.value = '$B$2:$C$4';
  address.listeners.get('keydown')({ key: 'Enter', preventDefault() {} });
  assert.equal(address.value, 'B2');
  assert.equal(document.querySelector('[data-row="1"][data-column="1"]').classList.contains('selected'), true);
  assert.equal(document.querySelector('[data-row="3"][data-column="2"]').getAttribute?.('aria-selected') ?? document.querySelector('[data-row="3"][data-column="2"]').attributes['aria-selected'], 'true');
  address.value = 'XFD1';
  address.listeners.get('keydown')({ key: 'Enter', preventDefault() {} });
  assert.notEqual(elements.get('goToStatus').textContent, '');
  assert.equal(document.querySelector('[data-row="1"][data-column="1"]').classList.contains('selected'), true);
  documentListeners.get('keydown')({ target: document.querySelector('.cell.selected'), key: 'g', ctrlKey: true, metaKey: false, preventDefault() {} });
  assert.equal(document.activeElement, address, 'Ctrl+G opens the address field');
});

test('XLSX freeze pane metadata reaches the editor model, sparse ranges and saved package', () => {
  const pane = '<sheetViews><sheetView workbookViewId="0"><pane xSplit="2" ySplit="3" topLeftCell="C4" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>';
  const bytes = minimalXlsx(pane);
  const document = new JinSheetDocument(uri('/tmp/frozen.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  assert.deepEqual(viewModel(document).freezePane, { rows: 3, columns: 2 });
  const range = getRange(document, { rowStart: 0, columnStart: 0, rowCount: 2, columnCount: 2, requestId: 1, viewRows: [0, 3], columnIndexes: [0, 2] });
  assert.deepEqual(range.viewRows, [0, 3]);
  assert.deepEqual(range.columnIndexes, [0, 2]);
  document.applyCellEdit(0, 0, 'changed');
  const reopened = analyze(document.createOutput());
  assert.deepEqual(reopened.sheets[0].freezePane, { rows: 3, columns: 2 });
  assert.ok(reopened.sheets[0].xml.includes(pane));
  document.dispose();
});

test('XLSX range exposes localized display values separately from raw cell values', () => {
  const cell = { address: 'A1', row: 0, column: 0, value: '45292', kind: 'number', numFmtId: 14, formatCode: 'm/d/yy' };
  const document = new JinSheetDocument(uri('/tmp/formatted.xlsx'), 'xlsx', { sheets: [{ cells: new Map([['A1', cell]]) }] }, Buffer.alloc(0), false, '');
  const result = getRange(document, { rowStart: 0, columnStart: 0, rowCount: 1, columnCount: 1, requestId: 1 });
  assert.equal(result.rows[0][0].value, '45292');
  assert.equal(result.rows[0][0].displayValue, '1/1/24');
  assert.match(renderHtml(), /cell\.textContent=data\?\.hyperlink\?\.display\?\?data\?\.displayValue\?\?data\?\.value\?\?''/);
});

test('XLSX 클립보드는 표시 숫자와 수식 텍스트를 다른 앱으로 복사한다', async () => {
  const bytes = await fs.readFile(path.join(__dirname, 'fixtures', 'JinSheet-Styled-Sample.xlsx'));
  const document = new JinSheetDocument(uri('/tmp/display-copy.xlsx'), 'xlsx', analyze(bytes), bytes, false, '');
  assert.equal(serializeClipboard(document, { startRow: 1, endRow: 1, startColumn: 0, endColumn: 3 }), 'Apples\t12\t1,500\t18,000');
  document.state.sheets[0].cells.set('E2', { address: 'E2', row: 1, column: 4, value: '18000', cachedValue: '18000', kind: 'number', formula: 'B2*C2', numFmtId: 3, formatCode: '#,##0' });
  assert.equal(serializeClipboard(document, { startRow: 1, endRow: 1, startColumn: 4, endColumn: 4 }), '=B2*C2');
  document.dispose();
});

test('XLSX에서 선택 범위 값을 지워도 셀 서식은 보존된다', async () => {
  const bytes = await fs.readFile(path.join(__dirname, 'fixtures', 'JinSheet-Styled-Sample.xlsx'));
  const state = analyze(bytes), originalCell = state.sheets[0].cells.get('C2');
  const document = new JinSheetDocument(uri('/tmp/clear-preserve-style.xlsx'), 'xlsx', state, bytes, true, '');
  let operation;
  document.onDidChange(change => { operation = change; });
  document.applyCellEdits(buildPasteEdits([['']], 1, 2, { startRow: 1, endRow: 1, startColumn: 2, endColumn: 2 }), 0, '선택 범위 값 지우기');
  const cleared = analyze(document.createOutput()).sheets[0].cells.get('C2');
  assert.equal(cleared.value, '');
  assert.equal(cleared.style, originalCell.style);
  operation.undo();
  assert.equal(document.getCell(1, 2).value, originalCell.value);
  document.dispose();
});

test('Delete로 XLSX 수식 범위를 지우면 수식과 값만 지우고 undo·redo한다', async t => {
  const bytes = minimalXlsx('', '<row r="1"><c r="A1"><f>B1+1</f><v>3</v></c><c r="B1"><v>2</v></c></row>');
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-clear-formula-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'clear-formula.xlsx');
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  let operation;
  document.onDidChange(change => { operation = change; });
  const edits = buildClearEdits(document, 0, 0, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 });
  assert.equal(edits[0].formula, null, 'clearing explicitly removes the protected formula');
  document.applyCellEdits(edits, 0, '선택 범위 값 지우기');
  assert.equal(document.getCell(0, 0).formula, null);
  assert.equal(document.getCell(0, 0).value, '');
  assert.equal(document.getCell(0, 1).value, '');
  await document.save();
  let cleared = analyze(await fs.readFile(file)).sheets[0];
  assert.equal(cleared.cells.get('A1').formula, null);
  assert.equal(cleared.cells.get('A1').value, '');
  operation.undo();
  assert.equal(document.getCell(0, 0).formula, 'B1+1');
  assert.equal(document.getCell(0, 1).value, '2');
  operation.redo();
  await document.save();
  cleared = analyze(await fs.readFile(file)).sheets[0];
  assert.equal(cleared.cells.get('A1').formula, null);
  assert.equal(cleared.cells.get('B1').value, '');
  document.dispose();
});

test('XLSX 범위 잘라내기는 값과 서식을 저장하고 저장 뒤 undo·redo해도 복원한다', async t => {
  const bytes = await fs.readFile(path.join(__dirname, 'fixtures', 'JinSheet-Styled-Sample.xlsx'));
  const state = analyze(bytes), source = state.sheets[0].cells.get('C2');
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-move-preserve-style-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'move-preserve-style.xlsx');
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', state, bytes, true, '');
  let operation;
  document.onDidChange(change => { operation = change; });
  const edits = buildMoveEdits(document, { startRow: 1, endRow: 1, startColumn: 2, endColumn: 2 }, 1, 4);
  document.applyCellEdits(edits, 0, '셀 범위 이동');
  await document.save();
  const moved = analyze(await fs.readFile(file)).sheets[0];
  assert.equal(moved.cells.get('C2').value, '');
  assert.equal(moved.cells.get('C2').style, null);
  assert.equal(moved.cells.get('E2').value, source.value);
  assert.equal(moved.cells.get('E2').style, source.style);
  assert.deepEqual(document.getCell(1, 4).styleProperties, source.styleProperties);
  operation.undo();
  assert.equal(document.getCell(1, 2).value, source.value);
  assert.equal(document.getCell(1, 2).style, source.style);
  assert.equal(document.getCell(1, 4).value, '');
  await document.save();
  const undone = analyze(await fs.readFile(file)).sheets[0];
  assert.equal(undone.cells.get('C2').value, source.value);
  assert.equal(undone.cells.get('C2').style, source.style);
  assert.equal(undone.cells.has('E2'), false);
  operation.redo();
  assert.equal(document.getCell(1, 2).style, null);
  assert.equal(document.getCell(1, 4).style, source.style);
  await document.save();
  const redone = analyze(await fs.readFile(file)).sheets[0];
  assert.equal(redone.cells.get('C2').style, null);
  assert.equal(redone.cells.get('E2').style, source.style);
  document.dispose();
});

test('XLSX 잘라내기는 정상 수식을 이동하고 대상 수식을 덮어쓰며 undo·redo한다', async t => {
  const bytes = minimalXlsx('', '<row r="1"><c r="A1"><f>B1*2</f><v>4</v></c><c r="B1"><v>2</v></c><c r="C1"><f>B1+10</f><v>12</v></c></row>');
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-move-formula-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'move-formula.xlsx');
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  let operation;
  document.onDidChange(change => { operation = change; });
  const edits = buildMoveEdits(document, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, 0, 2);
  assert.equal(edits.find(edit => edit.column === 2).formula, 'B1*2', 'a moved formula keeps its existing references');
  assert.equal(edits.find(edit => edit.column === 0).formula, null, 'the source formula is explicitly removed');
  document.applyCellEdits(edits, 0, '셀 범위 이동');
  await document.save();
  let moved = analyze(await fs.readFile(file)).sheets[0];
  assert.equal(moved.cells.get('A1').formula, null);
  assert.equal(moved.cells.get('C1').formula, 'B1*2');
  operation.undo();
  assert.equal(document.getCell(0, 0).formula, 'B1*2');
  assert.equal(document.getCell(0, 2).formula, 'B1+10');
  operation.redo();
  await document.save();
  moved = analyze(await fs.readFile(file)).sheets[0];
  assert.equal(moved.cells.get('C1').formula, 'B1*2');
  document.dispose();
});

test('Webview HTML의 스크립트 문법을 검사하고 원격 출처를 허용하지 않는다', () => {
  const html = renderHtml();
  const scripts = [...html.matchAll(/<script nonce="[^"]+">([\s\S]*?)<\/script>/g)];
  assert.ok(scripts.length >= 2);
  new vm.Script(scripts.at(-1)[1]);
  assert.match(html, /default-src 'none'/);
  assert.match(html, /style-src 'unsafe-inline'/);
  assert.match(html, /script-src 'nonce-/);
  assert.match(html, /background:#fff;color:#202124/);
  assert.match(html, /background:#f3f3f3;color:#444/);
  assert.match(html, /#reader\{[^}]*overflow:hidden[^}]*background:var\(--vscode-editor-background,#fff\);color:var\(--vscode-foreground,#202124\)/);
  assert.match(html, /#readerCanvas\{[^}]*overflow:auto[^}]*background:#f3f3f3;color:#202124/);
  assert.match(html, /section\.docx\{[^}]*box-shadow:0 1px 4px #0002!important/);
  assert.match(html, /#toolbar\{[^}]*background:var\(--vscode-editorGroupHeader-tabsBackground/);
  assert.match(html, /--vscode-button-secondaryBackground/);
  assert.match(html, /section\.docx\{[^}]*background:#fff!important;color:#202124!important/);
  assert.match(html, /Loading PDF pages/);
  assert.match(html, /body\.document-view\{height:100vh;overflow:hidden/);
  assert.match(html, /body\{display:flex;flex-direction:column;height:100vh;margin:0;overflow:hidden/);
  assert.match(html, /#grid\{[^}]*flex:1 1 auto;min-height:0;overflow:auto/);
  assert.match(html, /\.cell textarea\.cell-editor\{[^}]*height:16px;min-height:16px/);
  assert.match(html, /\.cell textarea\.cell-editor\{[^}]*user-select:text;-webkit-user-select:text/);
  assert.match(html, /#grid,#readerCanvas\{scrollbar-color:var\(--vscode-scrollbarSlider-background/);
  assert.match(html, /scrollbar-width:thin/);
  assert.match(html, /\.rowhead\{width:52px;z-index:1\}/);
  assert.match(html, /\.colhead\{top:0;z-index:1/);
  assert.match(html, /\.outline-toggle\{[^}]*cursor:pointer\}/);
  assert.match(html, /\.outline-bracket\{[^}]*pointer-events:none/);
  assert.match(html, /id="insertRow"/);
  assert.match(html, /id="deleteColumn"/);
  assert.match(html, /type:'structure'/);
  assert.match(html, /selectedRange/);
  assert.match(html, /range:\s*selectedRange/);
  assert.match(html, /cell\.onmouseenter/);
  assert.match(html, /id="formula" type="text"/);
  assert.match(html, /formula\.disabled=Boolean\(model\.readOnly\|\|\(data\?\.kind==='error'&&!data\?\.hasFormula\)\)/);
  assert.match(html, /formulaInput\.addEventListener\('keydown'/);
  assert.match(html, /type: 'navigate'/);
  assert.match(html, /e\.key === 'Home'/);
  assert.match(html, /model.gridRowCount=XLSX_MAX_ROWS/);
  assert.doesNotMatch(html, /type:'sort'|clearSort|sortHint|rowMap/);
  assert.doesNotMatch(html, /type:'filter'/);
  assert.doesNotMatch(html, /id="headerRow"|header-row/);
  assert.match(html, /\.cell\.in-range\{box-shadow:inset 0 0 0 9999px rgba\(0,100,220,\.14\)\}/);
  assert.match(html, /if\(style\.fill\)cell\.style\.backgroundColor=style\.fill/);
  assert.match(html, /const XLSX_MAX_ROWS=1048576;const XLSX_MAX_COLUMNS=16384/);
  assert.match(html, /visibleColumnRange\(columnOffsets,grid\.scrollLeft,grid\.clientWidth\)/);
  assert.match(html, /col-resize-handle/);
  assert.match(html, /aria-valuemin/);
  assert.match(html, /columnOffsets\[cell\.column\]/);
  assert.match(html, /new IntersectionObserver/);
  assert.match(html, /renderVisiblePdfPages\(generation\)/);
  assert.match(html, /readerCanvas'\)\.getBoundingClientRect\(\)/);
  assert.match(html, /container\.scrollTop\+delta/);
  assert.match(html, /Preparing pages/);
  assert.doesNotMatch(html, /이전 페이지|다음 페이지/);
  assert.match(html, /#pdfPage\{display:block;max-width:none/);
  assert.match(html, /pdfCanvasOutputScale\(viewport\.width,viewport\.height,window\.devicePixelRatio\|\|1\)/);
  assert.match(html, /transform:outputScale===1\?null:\[outputScale,0,0,outputScale,0,0\]/);
  assert.match(html, /fit\.textContent=t\('fit-width'\)/);
  assert.match(html, /id="readerCanvas"/);
  assert.match(html, /#readerCanvas\.pptx-view\.pptx-zoomed,#readerCanvas\.docx-view\.docx-zoomed\{overflow-x:auto\}/);
  assert.match(html, /setPptxZoom\(renderedInstance\.zoomPercent\+10,generation\)/);
  assert.match(html, /const width=renderedInstance\.slideDimensions\?\.width;if\(width>0\)setPptxZoom\(Math\.min\(100,canvas\.clientWidth\/width\*100\),generation\)/);
  assert.match(html, /canvas\.style\.padding='0';canvas\.classList\.add\('pptx-view'\)/);
  assert.match(html, /renderedInstance\.setZoom\(zoom\)/);
  assert.match(html, /docx-view\.docx-zoomed/);
  assert.match(html, /setDocxZoom\(docxUserZoom\+10,generation\)/);
  assert.match(html, /wrapper\.style\.zoom=String\(fitScale\*docxUserZoom\/100\)/);
  assert.match(html, /PDF compatibility rendering did not respond within 30 seconds/);
  assert.match(html, /\.docx-wrapper\{background:#f3f3f3!important/);
  assert.match(html, /\.flyfish-pptx-content,#readerCanvas \.slide\{background:#fff/);
  assert.match(html, /connect-src 'self'/);
  assert.match(html, /worker-src blob:/);
  assert.match(html, /self\.setImmediate=self\.setImmediate\|\|function\(callback/);
  assert.match(html, /self\.clearImmediate=self\.clearImmediate\|\|function\(handle/);
  assert.match(html, /renderers\.js/);
  assert.match(html, /pdf\.worker\.min\.mjs/);
  assert.match(html, /pdf\.min\.js/);
  assert.match(html, /globalThis\.pdfjsLib/);
  assert.match(html, /const nativeWorker=window\.Worker;window\.Worker=undefined/);
  assert.match(html, /JINSHEET_PDF_WORKER_TIMEOUT/);
  assert.match(html, /void pdfLoadingTask\.destroy\(\)\.catch/);
  assert.match(html, /Retrying with PDF compatibility mode/);
  assert.doesNotMatch(html, /현재 XLSX는 기존 값이 있는 셀만 수정할 수 있습니다/);
  assert.match(html, /e.key === 'Delete' \|\| e.key === 'Backspace'/);
  assert.match(html, /!e.target.closest\?\.\('#grid'\)/);
  assert.match(html, /Clear selected cells: Delete\/Backspace/);
  assert.match(html, /edit\(cell,row,column,data,e\.key\)/);
  assert.match(html, /input\.setSelectionRange\(input\.value\.length,input\.value\.length\)/);
  assert.match(html, /finally\{window\.Worker=nativeWorker\}/);
  assert.doesNotMatch(html, /type="module"/);
  assert.match(html, /fetch\(window\.jinSheetAssets\.pptxWorker\)/);
  assert.match(html, /workerType:'classic'/);
  assert.match(html, /URL\.createObjectURL\(new Blob/);
  assert.match(html, /URL\.revokeObjectURL\(pptxWorkerBlobUrl\)/);
  assert.doesNotMatch(html, /(?:src|href)\s*=\s*["']https?:\/\//i);
  assert.doesNotMatch(html, /url\(\s*["']?https?:\/\//i);
});

test('그리드 스크롤에서도 열 머리글·행 번호·모서리 머리글은 고정 위치를 유지한다', () => {
  const { elements, respondToRange } = runGridWebview(renderHtml());
  respondToRange();
  const grid = elements.get('grid'), canvas = elements.get('canvas');
  const corner = canvas.children.find(item => item.classList.contains('corner'));
  const column = canvas.children.find(item => item.classList.contains('colhead'));
  const row = canvas.children.find(item => item.classList.contains('rowhead'));
  grid.scrollTop = 260;
  grid.scrollLeft = 432;
  grid.listeners.get('scroll')();
  assert.equal(column.style.transform, 'translateY(260px)');
  assert.equal(row.style.transform, 'translateX(432px)');
  assert.equal(corner.style.transform, 'translate(432px,260px)');
});

test('셀 직접 입력에서 연속 키와 한글 IME 이벤트가 같은 편집기를 유지한다', () => {
  const english = runGridWebview(renderHtml());
  english.respondToRange();
  const englishCell = english.document.querySelector('[data-row="0"][data-column="0"]');
  englishCell.onkeydown({ target: englishCell, key: 'x', ctrlKey: false, metaKey: false, altKey: false, preventDefault() {}, stopPropagation() {} });
  const englishInput = englishCell.children[0];
  assert.equal(englishInput.value, 'x');
  assert.equal(englishInput.className, 'cell-editor');
  englishCell.onkeydown({ target: englishInput, key: 'y', ctrlKey: false, metaKey: false, altKey: false, preventDefault() {}, stopPropagation() {} });
  assert.strictEqual(englishCell.children[0], englishInput);
  englishInput.value += 'y';
  englishCell.onkeydown({ target: englishInput, key: 'z', ctrlKey: false, metaKey: false, altKey: false, preventDefault() {}, stopPropagation() {} });
  assert.strictEqual(englishCell.children[0], englishInput);
  englishInput.value += 'z';
  englishInput.onblur();
  assert.deepEqual(JSON.parse(JSON.stringify(english.messages.at(-1))), { type: 'edit', row: 0, column: 0, value: 'xyz', sheetIndex: 0 });

  const korean = runGridWebview(renderHtml());
  korean.respondToRange();
  const koreanCell = korean.document.querySelector('[data-row="0"][data-column="0"]');
  koreanCell.onkeydown({ target: koreanCell, key: 'Process', isComposing: true, ctrlKey: false, metaKey: false, altKey: false });
  const koreanInput = koreanCell.children[0];
  assert.equal(koreanInput.tagName, 'TEXTAREA');
  assert.equal(koreanInput.value, '');
  assert.equal(korean.document.activeElement, koreanInput);
  koreanCell.onkeydown({ target: koreanInput, key: 'Process', isComposing: true, ctrlKey: false, metaKey: false, altKey: false });
  assert.strictEqual(koreanCell.children[0], koreanInput);
  koreanInput.value = '월간매출';
  koreanInput.onblur();
  assert.deepEqual(JSON.parse(JSON.stringify(korean.messages.at(-1))), { type: 'edit', row: 0, column: 0, value: '월간매출', sheetIndex: 0 });
});

test('더블클릭 셀 편집 입력란은 마우스 텍스트 선택을 유지하고 부모 재편집을 막는다', () => {
  const html = renderHtml();
  assert.match(html, /textarea\.cell-editor\{[^}]*user-select:text;-webkit-user-select:text/);
  const { document, documentListeners } = runGridWebview(html);
  const cell = document.querySelector('[data-row="0"][data-column="0"]');
  cell.onclick();
  document.caretPositionFromPoint = () => ({ offsetNode: cell, offset: 1 });
  cell.ondblclick({ clientX: 10, clientY: 10 });
  const input = cell.children[0];
  assert.equal(input.className, 'cell-editor');
  assert.equal(input.selectionStart, 1, 'double-click editing puts the caret at the clicked text position');
  assert.equal(input.selectionEnd, 1);
  cell.ondblclick({ target: input, clientX: 10, clientY: 10 });
  assert.strictEqual(cell.children[0], input, 'double-clicking editor text must not replace the active editor');

  let prevented = false;
  cell.onmousedown({ target: input, clientX: 10, clientY: 10, preventDefault() { prevented = true; } });
  documentListeners.get('mousemove')({ clientX: 60, clientY: 10 });
  cell.onclick({ target: input });

  assert.equal(prevented, false, 'text selection mousedown must not be cancelled');
  assert.equal(document.activeElement, input, 'clicking selected input text must not move focus back to the grid');
  assert.strictEqual(cell.children[0], input, 'mouse selection must leave the same editor mounted');
});

test('F2는 현재 셀 편집을 시작하고 커서를 값 끝에 둔다', () => {
  const { document } = runGridWebview(renderHtml());
  const cell = document.querySelector('[data-row="0"][data-column="0"]');
  let prevented = false;
  cell.onkeydown({ target: cell, key: 'F2', preventDefault() { prevented = true; }, stopPropagation() {} });
  const input = cell.children[0];
  assert.equal(prevented, true);
  assert.equal(input.value, 'a');
  assert.equal(input.selectionStart, 1);
  assert.equal(input.selectionEnd, 1);
});

test('편집 중이 아닌 Enter와 Shift+Enter는 셀 선택을 아래·위로 이동한다', () => {
  const { document, documentListeners, elements, messages } = runGridWebview(renderHtml(), { rowCount: 3, columnCount: 2 });
  const first = document.querySelector('[data-row="0"][data-column="0"]');
  first.onkeydown({ target: first, key: 'Enter', preventDefault() {}, stopPropagation() {} });
  assert.equal(first.children.some(child => child.className === 'cell-editor'), false, 'Enter on a selected cell does not start editing');
  const keydown = documentListeners.get('keydown');
  keydown({ target: first, key: 'Enter', shiftKey: false, preventDefault() {} });
  assert.equal(elements.get('address').value, 'A2');
  keydown({ target: document.querySelector('[data-row="1"][data-column="0"]'), key: 'Enter', shiftKey: true, preventDefault() {} });
  assert.equal(elements.get('address').value, 'A1');
  assert.equal(messages.some(message => message.type === 'edit'), false);
});

test('XLSX 수식 셀은 F2 인라인 편집에서 수식을 표시하고 수정 내용을 커밋한다', () => {
  const { document, messages } = runGridWebview(renderHtml(), {
    kind: 'xlsx',
    cellData: {
      '0:0': { value: '2', displayValue: '2', formula: '1+1', hasFormula: true, cachedValue: '2' },
      '0:1': { value: '#DIV/0!', displayValue: '#DIV/0!', kind: 'error', formula: '1/0', hasFormula: true, cachedValue: '#DIV/0!' }
    }
  });
  const cell = document.querySelector('[data-row="0"][data-column="0"]');
  cell.onkeydown({ target: cell, key: 'F2', preventDefault() {}, stopPropagation() {} });
  const input = cell.children[0];
  assert.equal(input.value, '=1+1');
  assert.equal(input.selectionStart, input.value.length);
  input.value = '=1+2';
  input.onkeydown({ key: 'Enter', shiftKey: false, ctrlKey: false, metaKey: false, altKey: false, preventDefault() {} });
  assert.deepEqual(JSON.parse(JSON.stringify(messages.filter(message => message.type === 'edit').at(-1))), {
    type: 'edit', row: 0, column: 0, value: '=1+2', sheetIndex: 0
  });
  const errorFormulaCell = document.querySelector('[data-row="0"][data-column="1"]');
  errorFormulaCell.onkeydown({ target: errorFormulaCell, key: 'F2', preventDefault() {}, stopPropagation() {} });
  assert.equal(errorFormulaCell.children[0].value, '=1/0', 'formula errors remain editable');
});

test('Alt+Enter는 셀 편집 중 값에 줄바꿈을 삽입하고 Enter는 편집을 저장한다', () => {
  const { document, documentListeners, messages } = runGridWebview(renderHtml());
  const cell = document.querySelector('[data-row="0"][data-column="0"]');
  cell.onkeydown({ target: cell, key: 'x', isComposing: false, ctrlKey: false, metaKey: false, altKey: false, preventDefault() {}, stopPropagation() {} });
  const editor = cell.children[0];
  const globalKeydown = documentListeners.get('keydown');
  globalKeydown({ target: editor, key: 'ArrowLeft', preventDefault() {} });
  globalKeydown({ target: editor, key: 'c', ctrlKey: true, metaKey: false, preventDefault() {} });
  globalKeydown({ target: editor, key: 'v', ctrlKey: true, metaKey: false, preventDefault() {} });
  assert.equal(messages.some(message => ['navigate', 'copy', 'paste'].includes(message.type)), false, 'editor caret and clipboard shortcuts must not trigger grid commands');
  let prevented = false;
  editor.onkeydown({ key: 'Enter', altKey: true, ctrlKey: false, metaKey: false, preventDefault() { prevented = true; } });
  assert.equal(prevented, true);
  assert.equal(editor.value, 'x\n');
  editor.setRangeText('y', editor.selectionStart, editor.selectionEnd, 'end');
  editor.onkeydown({ key: 'Enter', altKey: false, ctrlKey: false, metaKey: false, isComposing: false, preventDefault() {} });
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'edit', row: 0, column: 0, value: 'x\ny', sheetIndex: 0 });
});

test('셀 편집기와 수식 입력줄의 Ctrl+Enter는 선택 범위 채우기 메시지를 보낸다', () => {
  const { document, documentListeners, elements, messages } = runGridWebview(renderHtml());
  const first = document.querySelector('[data-row="0"][data-column="0"]');
  const last = document.querySelector('[data-row="1"][data-column="1"]');
  first.onmousedown({ shiftKey: false, clientX: 80, clientY: 60, preventDefault() {} });
  last.onmouseenter();
  documentListeners.get('mouseup')();
  last.onkeydown({ target: last, key: 'x', ctrlKey: false, metaKey: false, altKey: false, preventDefault() {}, stopPropagation() {} });
  const editor = last.children[0];
  editor.value = '=B1+$C$1';
  editor.onkeydown({ key: 'Enter', ctrlKey: true, metaKey: false, altKey: false, preventDefault() {} });
  assert.deepEqual(JSON.parse(JSON.stringify(messages.filter(message => message.type === 'fill-range').at(-1))), {
    type: 'fill-range', row: 1, column: 1, value: '=B1+$C$1',
    range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 }, sheetIndex: 0
  });

  const formula = elements.get('formula');
  formula.value = '=B1+$C$1';
  formula.listeners.get('input')();
  formula.listeners.get('keydown')({ target: formula, key: 'Enter', ctrlKey: true, metaKey: false, preventDefault() {} });
  assert.equal(messages.at(-1).type, 'fill-range');
  assert.equal(messages.at(-1).row, 1);
  assert.equal(messages.at(-1).column, 1);
  assert.equal(messages.at(-1).value, '=B1+$C$1');

  const reverse = runGridWebview(renderHtml());
  const reverseFirst = reverse.document.querySelector('[data-row="0"][data-column="0"]');
  const reverseLast = reverse.document.querySelector('[data-row="1"][data-column="1"]');
  reverseLast.onmousedown({ shiftKey: false, clientX: 200, clientY: 90, preventDefault() {} });
  reverseFirst.onmouseenter();
  reverse.documentListeners.get('mouseup')();
  reverseFirst.onkeydown({ target: reverseFirst, key: 'x', ctrlKey: false, metaKey: false, altKey: false, preventDefault() {}, stopPropagation() {} });
  const reverseEditor = reverseFirst.children[0];
  reverseEditor.value = '=B1+$C$1';
  reverseEditor.onkeydown({ key: 'Enter', ctrlKey: true, metaKey: false, altKey: false, preventDefault() {} });
  const reverseFill = reverse.messages.find(message => message.type === 'fill-range');
  assert.equal(reverseFill.row, 0);
  assert.equal(reverseFill.column, 0);
  assert.deepEqual(JSON.parse(JSON.stringify(reverseFill.range)), { startRow: 1, endRow: 0, startColumn: 1, endColumn: 0 });
  const reverseFormula = reverse.elements.get('formula');
  reverseFormula.value = '=B1+$C$1';
  reverseFormula.listeners.get('input')();
  reverseFormula.listeners.get('keydown')({ target: reverseFormula, key: 'Enter', ctrlKey: true, metaKey: false, preventDefault() {} });
  const reverseFormulaFill = reverse.messages.filter(message => message.type === 'fill-range').at(-1);
  assert.equal(reverseFormulaFill.row, 0);
  assert.equal(reverseFormulaFill.column, 0);
  assert.deepEqual(JSON.parse(JSON.stringify(reverseFormulaFill.range)), { startRow: 1, endRow: 0, startColumn: 1, endColumn: 0 });
});

test('Backspace는 단일 셀 편집을 시작하고 Delete는 셀 값을 지운다', () => {
  const single = runGridWebview(renderHtml());
  const cell = single.document.querySelector('[data-row="0"][data-column="0"]');
  single.documentListeners.get('keydown')({ target: cell, key: 'Backspace', preventDefault() {}, stopPropagation() {} });
  assert.equal(cell.children[0].className, 'cell-editor');
  assert.equal(cell.children[0].value, '');
  assert.equal(single.messages.some(message => message.type === 'clear'), false);
  cell.children[0].onkeydown({ key: 'Escape', preventDefault() {} });
  assert.equal(single.messages.some(message => message.type === 'edit'), false, 'Escape restores the original value instead of committing the cleared edit');

  const range = runGridWebview(renderHtml());
  const first = range.document.querySelector('[data-row="0"][data-column="0"]');
  const last = range.document.querySelector('[data-row="1"][data-column="1"]');
  first.onmousedown({ shiftKey: false, clientX: 80, clientY: 60, preventDefault() {} });
  last.onmouseenter();
  range.documentListeners.get('mouseup')();
  range.documentListeners.get('keydown')({ target: last, key: 'Backspace', preventDefault() {}, stopPropagation() {} });
  assert.equal(range.messages.at(-1).type, 'clear', 'Backspace clears a multi-cell selection as one range operation');
  const deletion = runGridWebview(renderHtml());
  const deleteCell = deletion.document.querySelector('[data-row="0"][data-column="0"]');
  deletion.documentListeners.get('keydown')({ target: deleteCell, key: 'Delete', preventDefault() {} });
  assert.equal(deletion.messages.at(-1).type, 'clear');
});

test('수식 입력줄은 포커스를 벗어날 때 값을 커밋하고 Escape는 변경을 취소한다', () => {
  const committed = runGridWebview(renderHtml());
  const formula = committed.elements.get('formula');
  formula.value = 'blurred value';
  formula.listeners.get('input')();
  formula.listeners.get('blur')();
  assert.deepEqual(JSON.parse(JSON.stringify(committed.messages.at(-1))), { type: 'edit', row: 0, column: 0, value: 'blurred value', sheetIndex: 0 });

  const clicked = runGridWebview(renderHtml());
  const clickFormula = clicked.elements.get('formula');
  clickFormula.value = 'commit before selection';
  clickFormula.listeners.get('input')();
  clicked.document.querySelector('[data-row="0"][data-column="1"]').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(clicked.messages.at(-1))), { type: 'edit', row: 0, column: 0, value: 'commit before selection', sheetIndex: 0 });

  const sameCell = runGridWebview(renderHtml());
  const sameFormula = sameCell.elements.get('formula');
  sameFormula.value = 'commit on active cell click';
  sameFormula.listeners.get('input')();
  sameCell.document.querySelector('[data-row="0"][data-column="0"]').onmousedown({ target: sameCell.document.querySelector('[data-row="0"][data-column="0"]'), preventDefault() {} });
  sameCell.documentListeners.get('mouseup')();
  assert.deepEqual(JSON.parse(JSON.stringify(sameCell.messages.at(-1))), { type: 'edit', row: 0, column: 0, value: 'commit on active cell click', sheetIndex: 0 });

  const redrawn = runGridWebview(renderHtml());
  const redrawFormula = redrawn.elements.get('formula');
  redrawFormula.focus();
  redrawFormula.value = 'keep during redraw';
  redrawFormula.listeners.get('input')();
  redrawn.respondToRange();
  assert.equal(redrawFormula.value, 'keep during redraw');
  redrawFormula.listeners.get('blur')();
  assert.deepEqual(JSON.parse(JSON.stringify(redrawn.messages.at(-1))), { type: 'edit', row: 0, column: 0, value: 'keep during redraw', sheetIndex: 0 });

  const cancelled = runGridWebview(renderHtml());
  const cancelFormula = cancelled.elements.get('formula');
  cancelFormula.value = 'discard this';
  cancelFormula.listeners.get('input')();
  cancelFormula.listeners.get('keydown')({ target: cancelFormula, key: 'Escape', preventDefault() {} });
  cancelFormula.listeners.get('blur')();
  assert.equal(cancelled.messages.some(message => message.type === 'edit'), false);
  assert.equal(cancelFormula.value, 'a');
});

test('변경되지 않은 수식 입력줄에서 Enter와 Tab을 눌러도 셀 편집을 만들지 않는다', () => {
  const grid = runGridWebview(renderHtml());
  const formula = grid.elements.get('formula');
  formula.listeners.get('keydown')({ target: formula, key: 'Enter', preventDefault() {} });
  assert.equal(grid.elements.get('address').value, 'A2', 'Enter still advances to the next row');
  assert.equal(grid.messages.some(message => message.type === 'edit'), false);
  formula.listeners.get('keydown')({ target: formula, key: 'Tab', preventDefault() {} });
  assert.equal(grid.elements.get('address').value, 'B2', 'Tab still advances to the next column');
  assert.equal(grid.messages.some(message => message.type === 'edit'), false);
});

test('셀 범위를 편집할 때 Tab과 Enter는 모델 갱신 뒤에도 선택 영역 안에서 이동하고 범위를 유지한다', () => {
  const { document, documentListeners, elements, dispatch, respondToRange } = runGridWebview(renderHtml(), { rowCount: 4, columnCount: 4 });
  let a1 = document.querySelector('[data-row="0"][data-column="0"]');
  let b2 = document.querySelector('[data-row="1"][data-column="1"]');
  a1.onmousedown({ shiftKey: false, clientX: 80, clientY: 60, preventDefault() {} });
  b2.onmouseenter();
  documentListeners.get('mouseup')();
  const beginEdit = cell => {
    cell.onkeydown({ target: cell, key: 'x', isComposing: false, ctrlKey: false, metaKey: false, altKey: false, preventDefault() {}, stopPropagation() {} });
    const input = cell.children[0];
    input.blur = () => input.onblur();
    input.value = 'edited';
    return input;
  };

  beginEdit(b2).onkeydown({ key: 'Tab', shiftKey: false, preventDefault() {}, stopPropagation() {} });
  assert.equal(elements.get('address').value, 'A1', 'Tab from the final selected cell wraps to the selection start');
  assert.equal(a1.classList.contains('in-range'), true, 'the selected range remains visible while the active cell moves');
  dispatch({ type: 'model', kind: 'csv', readOnly: false, rowCount: 4, columnCount: 4, sheetIndex: 0, dirty: true });
  respondToRange();
  a1 = document.querySelector('[data-row="0"][data-column="0"]');
  b2 = document.querySelector('[data-row="1"][data-column="1"]');
  assert.equal(a1.classList.contains('in-range'), true, 'same-sheet model updates preserve the selected range');
  beginEdit(a1).onkeydown({ key: 'Tab', shiftKey: true, preventDefault() {}, stopPropagation() {} });
  assert.equal(elements.get('address').value, 'B2', 'Shift+Tab from the first selected cell wraps to the selection end');

  beginEdit(b2).onkeydown({ key: 'Enter', shiftKey: false, preventDefault() {} });
  assert.equal(elements.get('address').value, 'B1', 'Enter advances vertically and wraps within the selected range');
  const b1 = document.querySelector('[data-row="0"][data-column="1"]');
  beginEdit(b1).onkeydown({ key: 'Enter', shiftKey: true, preventDefault() {} });
  assert.equal(elements.get('address').value, 'B2', 'Shift+Enter moves upward and wraps within the selected range');
});

test('Webview 드래그 범위 선택·복사와 수식 입력줄 Enter가 편집 메시지를 보낸다', () => {
  const { document, documentListeners, elements, messages, windowListeners, dispatch, respondToRange } = runGridWebview(renderHtml());
  let first = document.querySelector('[data-row="0"][data-column="0"]');
  let last = document.querySelector('[data-row="1"][data-column="1"]');
  first.onmousedown({ shiftKey: false, clientX: 80, clientY: 60, preventDefault() {} });
  last.onmouseenter();
  documentListeners.get('mouseup')();
  last.onclick();
  documentListeners.get('keydown')({ target: last, key: 'c', ctrlKey: true, metaKey: false, preventDefault() {} });
  const copy = messages.at(-1);
  assert.equal(copy.type, 'copy');
  assert.deepEqual(JSON.parse(JSON.stringify(copy.range)), { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 });
  const keydown = documentListeners.get('keydown'), address = elements.get('address');
  keydown({ target: last, key: 'Delete', preventDefault() {} });
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'clear', row: 0, column: 0, range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 }, sheetIndex: 0 }, 'Delete clears the whole selected range');
  keydown({ target: last, key: 'v', ctrlKey: true, metaKey: false, preventDefault() {} });
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'paste', row: 0, column: 0, range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 }, sheetIndex: 0 }, 'paste begins at the selected range anchor and carries its dimensions');
  last.onmousedown({ shiftKey: false, clientX: 80, clientY: 60, preventDefault() {} });
  first.onmouseenter();
  documentListeners.get('mouseup')();
  first.onclick();
  keydown({ target: first, key: 'v', ctrlKey: true, metaKey: false, preventDefault() {} });
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'paste', row: 1, column: 1, range: { startRow: 1, endRow: 0, startColumn: 1, endColumn: 0 }, sheetIndex: 0 }, 'reverse selection keeps its original active cell as the paste start');
  keydown({ target: first, key: 'x', ctrlKey: true, metaKey: false, preventDefault() {} });
  assert.equal(messages.at(-1).type, 'cut');
  const cutRange = { startRow: 1, endRow: 0, startColumn: 1, endColumn: 0 };
  dispatch({ type: 'cut-ready', range: cutRange, sheetIndex: 0 });
  keydown({ target: first, key: 'v', ctrlKey: true, metaKey: false, preventDefault() {} });
  assert.equal(messages.at(-1).type, 'move-paste', 'paste after cut requests one internal move operation');
  keydown({ target: first, key: 'x', ctrlKey: true, metaKey: false, preventDefault() {} });
  dispatch({ type: 'cut-ready', range: cutRange, sheetIndex: 0 });
  keydown({ target: first, key: 'Escape', preventDefault() {} });
  keydown({ target: first, key: 'v', ctrlKey: true, metaKey: false, preventDefault() {} });
  assert.equal(messages.at(-1).type, 'paste', 'Escape cancels a pending move and restores normal clipboard paste');
  first.onmousedown({ shiftKey: false, clientX: 80, clientY: 60, preventDefault() {} });
  last.onmouseenter();
  documentListeners.get('mouseup')();
  last.onclick();
  assert.ok(first.closest('#grid'));
  assert.equal(first.tagName, 'DIV');
  keydown({ target: first, key: 'Home', preventDefault() {} });
  assert.equal(address.value, 'A2');
  respondToRange(); first = document.querySelector('[data-row="0"][data-column="0"]');
  let endHandled = false;
  keydown({ target: first, key: 'End', preventDefault() { endHandled = true; } });
  assert.equal(endHandled, true);
  assert.equal(address.value, 'B2');
  respondToRange(); first = document.querySelector('[data-row="0"][data-column="0"]');
  keydown({ target: first, key: 'End', ctrlKey: true, preventDefault() {} });
  assert.equal(address.value, 'B2');
  respondToRange(); first = document.querySelector('[data-row="0"][data-column="0"]');
  keydown({ target: first, key: 'Home', ctrlKey: true, preventDefault() {} });
  assert.equal(address.value, 'A1');
  respondToRange(); first = document.querySelector('[data-row="0"][data-column="0"]');
  keydown({ target: first, key: 'ArrowRight', ctrlKey: true, preventDefault() {} });
  assert.equal(messages.at(-1).type, 'navigate');
  assert.equal(messages.at(-1).direction, 'right');
  windowListeners.get('message')({ data: { type: 'navigation', requestId: messages.at(-1).requestId, row: 0, column: 1, extend: false } });
  assert.equal(address.value, 'B1');
  respondToRange();
  respondToRange(); first = document.querySelector('[data-row="0"][data-column="0"]'); first.onclick();
  keydown({ target: first, key: 'ArrowDown', shiftKey: true, preventDefault() {} });
  assert.equal(address.value, 'A2');
  respondToRange(); last = document.querySelector('[data-row="1"][data-column="0"]');
  keydown({ target: last, key: 'c', ctrlKey: true, metaKey: false, preventDefault() {} });
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1).range)), { startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 });
  const formula = elements.get('formula');
  formula.value = 'edited';
  formula.listeners.get('input')();
  formula.listeners.get('keydown')({ target: formula, key: 'Enter', preventDefault() {} });
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'edit', row: 1, column: 0, value: 'edited', sheetIndex: 0 });
  assert.equal(address.value, 'A1', 'formula bar Enter wraps within the selected range');
  formula.listeners.get('keydown')({ target: formula, key: 'Enter', shiftKey: true, preventDefault() {} });
  assert.equal(address.value, 'A2', 'formula bar Shift+Enter moves up within the selected range');
});

test('저장된 행 높이는 셀·머리글·드래그 좌표와 페이지 이동에 반영된다', () => {
  const grid = runGridWebview(renderHtml(), { kind: 'xlsx', rowCount: 100, rowHeights: { 0: 60, 1: 40, 3: 80 }, freezePane: { rows: 1, columns: 0 }, cellData: { '0:0': { value: 'one\ntwo', style: { wrapText: true } } } });
  const cell = grid.document.querySelector('[data-row="0"][data-column="0"]');
  assert.equal(cell.style.height, '60px'); assert.equal(cell.style.whiteSpace, 'pre-wrap');
  assert.equal(grid.document.querySelector('[data-row="1"][data-column="0"]').style.top, '90px');
  assert.equal(grid.document.querySelector('[data-row="2"][data-column="0"]').style.top, '130px');
  const rowHead = grid.elements.get('canvas').children.find(item => item.classList.contains('rowhead') && item.textContent === '2');
  assert.equal(rowHead.style.height, '40px'); assert.equal(rowHead.style.top, '90px');
  cell.onmousedown({ target: cell, preventDefault() {}, clientX: 80, clientY: 40, shiftKey: false });
  grid.documentListeners.get('mousemove')({ clientX: 80, clientY: 145 });
  grid.documentListeners.get('mouseup')();
  assert.equal(grid.elements.get('address').value, 'A3', 'pixel hit testing respects tall rows');
  cell.onmousedown({ target: cell, preventDefault() {}, clientX: 80, clientY: 40, shiftKey: false }); grid.documentListeners.get('mouseup')();
  grid.documentListeners.get('keydown')({ target: cell, key: 'PageDown', preventDefault() {} });
  assert.equal(grid.elements.get('address').value, 'A5', 'page movement uses viewport pixels after the tall frozen row');
});

test('행 높이 도구는 선택 행의 pt 값·기본 높이 메시지를 보내고 CSV에서 비활성화된다', () => {
  const grid = runGridWebview(renderHtml(), { kind: 'xlsx', rowCount: 4, rowHeights: { 0: 80 } });
  assert.equal(grid.elements.get('rowHeight').value, '60');
  const first = grid.document.querySelector('[data-row="0"][data-column="0"]'), last = grid.document.querySelector('[data-row="2"][data-column="0"]');
  first.onmousedown({ target: first, preventDefault() {}, clientX: 80, clientY: 40, shiftKey: false }); last.onmouseenter(); grid.documentListeners.get('mouseup')();
  grid.elements.get('rowHeight').value = '45'; grid.elements.get('applyRowHeight').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(grid.messages.at(-1))), { type: 'row-height', startRow: 0, endRow: 2, height: 45, sheetIndex: 0 });
  grid.elements.get('resetRowHeight').onclick(); assert.equal(grid.messages.at(-1).height, null);
  const count = grid.messages.length; grid.elements.get('rowHeight').value = '500'; grid.elements.get('applyRowHeight').onclick(); assert.equal(grid.messages.length, count);
  const csv = runGridWebview(renderHtml(), { kind: 'csv' }); assert.equal(csv.elements.get('rowHeight').disabled, true); assert.equal(csv.elements.get('applyRowHeight').disabled, true);
});

test('행 높이 변경과 기본 높이 복원은 셀 초안·저장·undo·redo를 보존한다', async t => {
  const bytes = minimalXlsx(), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-row-height-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'height.xlsx'); await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  document.applyCellEdit(0, 0, 'draft');
  let operation; document.onDidChange(change => { operation = change; });
  document.applyRowHeights(0, 0, 1, 45);
  assert.equal(document.getCell(0, 0).value, 'draft'); assert.deepEqual(viewModel(document).rowHeights, { 0: 60, 1: 60 });
  operation.undo(); assert.deepEqual(viewModel(document).rowHeights, {}); assert.equal(document.getCell(0, 0).value, 'draft');
  operation.redo(); await document.save(); assert.deepEqual(analyze(await fs.readFile(file)).sheets[0].rowHeights, { 0: 60, 1: 60 });
  document.applyRowHeights(0, 0, 1, null); await document.save(); assert.deepEqual(analyze(await fs.readFile(file)).sheets[0].rowHeights, {});
  document.dispose();
});

test('줄 바꿈과 자동 높이 변경은 한 번의 undo·redo·저장으로 함께 반영한다', async t => {
  const bytes = minimalXlsx(), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-wrap-layout-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'wrap.xlsx'); await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  const range = { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 };
  let operation, count = 0; document.onDidChange(change => { operation = change; count++; });
  const edits = buildWrapTextEdits(document, range);
  const fit = buildRowFitRequest(document, 0, 0, 0, edits, true);
  assert.equal(fit.rows[0].cells[0].wrapText, true);
  assert.throws(() => document.applyTextLayout(0, edits, [{ row: 0, height: 500 }]), /자동 맞춤 결과/);
  assert.equal(Boolean(document.getCell(0, 0).styleProperties?.wrapText), false); assert.equal(count, 0);
  document.applyTextLayout(0, edits, [{ row: 0, height: 45 }], '텍스트 줄 바꿈 변경');
  assert.equal(count, 1); assert.equal(document.getCell(0, 0).styleProperties.wrapText, true); assert.equal(viewModel(document).rowHeights[0], 60);
  operation.undo(); assert.equal(Boolean(document.getCell(0, 0).styleProperties?.wrapText), false); assert.deepEqual(viewModel(document).rowHeights, {});
  operation.redo(); await document.save();
  const saved = analyze(await fs.readFile(file)); assert.equal(saved.sheets[0].cells.get('A1').styleProperties.wrapText, true); assert.deepEqual(saved.sheets[0].customRowHeights, []);
  operation.undo(); await document.save(); assert.equal(Boolean(analyze(await fs.readFile(file)).sheets[0].cells.get('A1').styleProperties?.wrapText), false);
  operation.redo(); await document.save(); assert.equal(analyze(await fs.readFile(file)).sheets[0].cells.get('A1').styleProperties.wrapText, true);
  assert.equal(buildWrapTextEdits(document, range)[0].alignment.wrapText, false);
  document.applyRowHeights(0, 0, 0, 75);
  assert.deepEqual(buildRowFitRequest(document, 0, 0, 0, buildWrapTextEdits(document, range), true).rows, [], 'manual height is kept when toggling wrap');
  document.dispose();
});

test('행 자동 맞춤 결과는 오래된 문서 상태와 요청 밖 행을 거부한다', () => {
  const pending = { sheetIndex: 0, revision: 3, rows: [1] }, message = { sheetIndex: 0, heights: [{ row: 1, height: 30 }] };
  assert.doesNotThrow(() => validateRowFitResult(pending, message, 3, 0));
  assert.throws(() => validateRowFitResult(pending, message, 4, 0), /시트/);
  assert.throws(() => validateRowFitResult(pending, { ...message, heights: [{ row: 2, height: 30 }] }, 3, 0), /결과/);
});

test('줄 바꿈·행 자동 맞춤 버튼과 브라우저 측정 응답을 연결한다', () => {
  const grid = runGridWebview(renderHtml(), { kind: 'xlsx' });
  grid.elements.get('wrapText').onclick(); assert.equal(grid.messages.at(-1).type, 'wrap-text');
  grid.elements.get('autoFitRow').onclick(); assert.equal(grid.messages.at(-1).type, 'row-auto-fit');
  grid.dispatch({ type: 'row-fit-request', requestId: 4, sheetIndex: 0, defaultHeight: 26, rows: [{ row: 0, cells: [{ text: '<script>', width: 144, wrapText: true }] }] });
  const response = JSON.parse(JSON.stringify(grid.messages.at(-1)));
  assert.equal(response.type, 'row-fit-result'); assert.equal(response.requestId, 4); assert.deepEqual(response.heights, [{ row: 0, height: 19.5 }]);
});

test('행 경계 드래그는 선택한 여러 행을 미리 보기하고 해제 시 한 번 저장한다', () => {
  const grid = runGridWebview(renderHtml(), { kind: 'xlsx', rowCount: 5 });
  const heads = grid.elements.get('canvas').children.filter(item => item.classList.contains('rowhead'));
  heads[0].onmousedown({ target: heads[0], preventDefault() {}, clientX: 10, clientY: 40, shiftKey: false }); heads[1].onmouseenter(); grid.documentListeners.get('mouseup')();
  const handle = heads[0].children.find(item => item.classList.contains('row-resize-handle'));
  handle.onmousedown({ preventDefault() {}, stopPropagation() {}, clientY: 55 });
  grid.documentListeners.get('mousemove')({ clientX: 10, clientY: 75 });
  assert.equal(grid.document.querySelector('[data-row="0"][data-column="0"]').style.height, '46px');
  assert.equal(grid.document.querySelector('[data-row="1"][data-column="0"]').style.height, '46px');
  assert.equal(grid.document.querySelector('[data-row="2"][data-column="0"]').style.top, '122px');
  assert.equal(grid.messages.filter(message => message.type === 'row-height').length, 0);
  grid.documentListeners.get('mouseup')();
  const heights = grid.messages.filter(message => message.type === 'row-height'); assert.equal(heights.length, 1);
  assert.deepEqual(JSON.parse(JSON.stringify(heights[0])), { type: 'row-height', startRow: 0, endRow: 1, height: 34.5, sheetIndex: 0 });
  handle.ondblclick({ preventDefault() {}, stopPropagation() {} }); assert.equal(grid.messages.at(-1).type, 'row-auto-fit');
});

test('행 경계 드래그는 Escape·창 이탈·새 모델 수신 시 저장하지 않고 취소한다', () => {
  const grid = runGridWebview(renderHtml(), { kind: 'xlsx' });
  const getHandle = () => grid.elements.get('canvas').children.find(item => item.classList.contains('rowhead')).children.find(item => item.classList.contains('row-resize-handle'));
  const start = () => { getHandle().onmousedown({ preventDefault() {}, stopPropagation() {}, clientY: 55 }); grid.documentListeners.get('mousemove')({ clientX: 10, clientY: 80 }); };
  start(); grid.documentListeners.get('keydown')({ key: 'Escape', preventDefault() {} });
  assert.equal(grid.document.querySelector('[data-row="0"][data-column="0"]').style.height, '26px'); assert.equal(grid.document.body.style.cursor, '');
  start(); grid.windowListeners.get('blur')(); assert.equal(grid.document.querySelector('[data-row="0"][data-column="0"]').style.height, '26px');
  start(); grid.dispatch({ type: 'model', kind: 'xlsx', readOnly: false, rowCount: 2, columnCount: 2, sheetIndex: 1 }); grid.documentListeners.get('mouseup')();
  assert.equal(grid.messages.filter(message => message.type === 'row-height').length, 0);
  const csv = runGridWebview(renderHtml(), { kind: 'csv' }); assert.equal(csv.elements.get('canvas').children.some(item => item.children.some(child => child.classList.contains('row-resize-handle'))), false);
});

test('행 경계의 키보드 조절은 pt 단위와 높이 한도를 적용한다', () => {
  const grid = runGridWebview(renderHtml(), { kind: 'xlsx' });
  const handle = grid.elements.get('canvas').children.find(item => item.classList.contains('rowhead')).children.find(item => item.classList.contains('row-resize-handle'));
  handle.focus();
  handle.onkeydown({ key: 'ArrowDown', preventDefault() {}, stopPropagation() {} });
  assert.equal(grid.messages.filter(message => message.type === 'row-height').at(-1).height, 20.5);
  assert.equal(handle.attributes['aria-orientation'], 'horizontal');
  grid.dispatch({ type: 'model', kind: 'xlsx', readOnly: false, rowCount: 2, columnCount: 2, sheetIndex: 0, rowHeights: { 0: 27.33 } }); grid.respondToRange();
  const nextHandle = grid.document.activeElement;
  assert.equal(nextHandle.classList.contains('row-resize-handle'), true, 'keyboard focus survives the model refresh');
  nextHandle.onkeydown({ key: 'ArrowDown', preventDefault() {}, stopPropagation() {} });
  assert.equal(grid.messages.filter(message => message.type === 'row-height').at(-1).height, 21.5);
  grid.dispatch({ type: 'model', kind: 'xlsx', readOnly: false, rowCount: 2, columnCount: 2, sheetIndex: 0, rowHeights: { 0: 60.44 } }); grid.respondToRange();
  grid.document.activeElement.onkeydown({ key: 'ArrowDown', preventDefault() {}, stopPropagation() {} });
  assert.equal(grid.messages.filter(message => message.type === 'row-height').at(-1).height, 46.33, 'a 45.33 pt row grows by exactly one point');
});

test('PageUp/PageDown은 화면 단위로 이동하고 Ctrl+PageUp/PageDown은 시트별 선택을 복원한다', () => {
  const { document, documentListeners, elements, messages, dispatch, respondToRange } = runGridWebview(renderHtml(), { rowCount: 100, columnCount: 4, kind: 'xlsx' });
  const sheets = [
    { name: 'One', index: 0, visible: true },
    { name: 'Hidden', index: 1, visible: false },
    { name: 'Three', index: 2, visible: true }
  ];
  dispatch({ type: 'model', kind: 'xlsx', readOnly: false, rowCount: 100, columnCount: 4, sheetIndex: 0, sheets });
  respondToRange();
  let cell = document.querySelector('[data-row="8"][data-column="1"]');
  cell.onmousedown({ shiftKey: false, clientX: 80, clientY: 60, preventDefault() {} });
  documentListeners.get('mouseup')();
  const keydown = documentListeners.get('keydown');
  keydown({ target: cell, key: 'PageDown', shiftKey: false, preventDefault() {} });
  assert.equal(elements.get('address').value, 'B19', 'PageDown advances by the visible row count');
  keydown({ target: cell, key: 'PageUp', shiftKey: false, preventDefault() {} });
  assert.equal(elements.get('address').value, 'B9', 'PageUp returns by one visible page');

  keydown({ target: cell, key: 'PageDown', shiftKey: true, preventDefault() {} });
  assert.equal(document.querySelector('[data-row="18"][data-column="1"]').classList.contains('in-range'), true, 'Shift+PageDown extends the selection');
  keydown({ target: cell, key: 'PageUp', shiftKey: false, preventDefault() {} });

  keydown({ target: cell, key: 'PageDown', ctrlKey: true, preventDefault() {} });
  assert.deepEqual(JSON.parse(JSON.stringify(messages.filter(message => message.type === 'refresh').at(-1))), { type: 'refresh', sheetIndex: 2 }, 'hidden sheets are skipped');
  dispatch({ type: 'model', kind: 'xlsx', readOnly: false, rowCount: 100, columnCount: 4, sheetIndex: 2, sheets });
  respondToRange();
  cell = document.querySelector('[data-row="4"][data-column="2"]');
  cell.onmousedown({ shiftKey: false, clientX: 80, clientY: 60, preventDefault() {} });
  documentListeners.get('mouseup')();
  keydown({ target: cell, key: 'PageUp', ctrlKey: true, preventDefault() {} });
  assert.deepEqual(JSON.parse(JSON.stringify(messages.filter(message => message.type === 'refresh').at(-1))), { type: 'refresh', sheetIndex: 0 });
  dispatch({ type: 'model', kind: 'xlsx', readOnly: false, rowCount: 100, columnCount: 4, sheetIndex: 0, sheets });
  respondToRange();
  assert.equal(elements.get('address').value, 'B9', 'returning to a sheet restores its previous active cell');
});

test('PageUp/PageDown은 필터로 숨긴 행을 제외한 표시 행 기준으로 이동한다', () => {
  const visibleRows = Array.from({ length: 100 }, (_, index) => index * 2);
  const { document, documentListeners, elements } = runGridWebview(renderHtml(), { rowCount: 200, columnCount: 2, kind: 'xlsx', visibleRows });
  const cell = document.querySelector('[data-row="16"][data-column="0"]');
  cell.onmousedown({ shiftKey: false, clientX: 80, clientY: 60, preventDefault() {} });
  documentListeners.get('mouseup')();
  const keydown = documentListeners.get('keydown');
  keydown({ target: cell, key: 'PageDown', preventDefault() {} });
  assert.equal(elements.get('address').value, 'A37', 'PageDown advances by ten visible rows, skipping hidden rows');
  keydown({ target: cell, key: 'PageUp', preventDefault() {} });
  assert.equal(elements.get('address').value, 'A17', 'PageUp returns by ten visible rows');
});

test('PageUp/PageDown은 고정 행을 제외한 스크롤 가능 화면 크기로 이동한다', () => {
  const { document, documentListeners, elements } = runGridWebview(renderHtml(), { rowCount: 100, columnCount: 2, kind: 'xlsx', freezePane: { rows: 2, columns: 0 } });
  const cell = document.querySelector('[data-row="7"][data-column="0"]');
  cell.onmousedown({ shiftKey: false, clientX: 80, clientY: 60, preventDefault() {} });
  documentListeners.get('mouseup')();
  documentListeners.get('keydown')({ target: cell, key: 'PageDown', preventDefault() {} });
  assert.equal(elements.get('address').value, 'A16', 'the viewport contains eight scrollable rows after two frozen rows');
});

test('XLSX freeze panes request sparse ranges on deep scroll and keep frozen cells visible', async () => {
  const { document, elements, messages, respondToRange } = runGridWebview(renderHtml(), { rowCount: 1000, columnCount: 30, kind: 'xlsx', freezePane: { rows: 2, columns: 2 } });
  const grid = elements.get('grid');
  const initial = messages.filter(message => message.type === 'range').at(-1);
  assert.deepEqual(Array.from(initial.viewRows.slice(0, 2)), [0, 1]);
  assert.deepEqual(Array.from(initial.columnIndexes.slice(0, 2)), [0, 1]);
  grid.scrollTop = 2600;
  grid.scrollLeft = 1200;
  grid.listeners.get('scroll')();
  await new Promise(resolve => setTimeout(resolve, 80));
  const scrolled = messages.filter(message => message.type === 'range').at(-1);
  assert.ok(scrolled.viewRows.includes(0) && scrolled.viewRows.includes(1));
  assert.ok(scrolled.viewRows.includes(100));
  assert.ok(scrolled.columnIndexes.includes(0) && scrolled.columnIndexes.includes(1));
  assert.ok(scrolled.columnIndexes.some(column => column >= 10));
  respondToRange();
  const frozenRow = document.querySelector('[data-row="0"][data-column="0"]');
  const frozenColumn = document.querySelector('[data-row="100"][data-column="0"]');
  const frozenCell = document.querySelector('[data-row="0"][data-column="1"]');
  assert.equal(frozenRow.style.transform, 'translate(1200px,2600px)');
  assert.equal(frozenColumn.style.transform, 'translate(1200px,0px)');
  assert.equal(frozenCell.style.transform, 'translate(1200px,2600px)');
  const count = messages.filter(message => message.type === 'range').length;
  grid.listeners.get('scroll')();
  await new Promise(resolve => setTimeout(resolve, 80));
  assert.equal(messages.filter(message => message.type === 'range').length, count);
});

test('현재 viewport 안의 키보드 이동은 범위 요청을 반복하지 않고 선택 셀 포커스를 유지한다', () => {
  const { document, documentListeners, elements, messages } = runGridWebview(renderHtml());
  let cell = document.querySelector('[data-row="0"][data-column="0"]');
  cell.onclick();
  const initialRangeRequests = messages.filter(message => message.type === 'range').length;
  const keydown = documentListeners.get('keydown');
  keydown({ target: cell, key: 'ArrowRight', preventDefault() {} });
  cell = document.querySelector('[data-row="0"][data-column="1"]');
  assert.equal(elements.get('address').value, 'B1');
  assert.equal(document.activeElement, cell);
  assert.equal(cell.attributes['aria-selected'], 'true');
  assert.equal(cell.tabIndex, 0);
  assert.equal(document.querySelector('[data-row="0"][data-column="0"]').tabIndex, -1);
  keydown({ target: cell, key: 'ArrowDown', preventDefault() {} });
  cell = document.querySelector('[data-row="1"][data-column="1"]');
  assert.equal(elements.get('address').value, 'B2');
  assert.equal(document.activeElement, cell);
  assert.equal(messages.filter(message => message.type === 'range').length, initialRangeRequests);
});

test('현재 viewport 안의 키보드 이동은 수식 입력줄을 새 셀 값으로 갱신한다', () => {
  const { document, documentListeners, elements } = runGridWebview(renderHtml(), { rowCount: 3, columnCount: 3 });
  const active = document.querySelector('[data-row="0"][data-column="0"]');
  documentListeners.get('keydown')({ target: active, key: 'ArrowRight', preventDefault() {} });
  assert.equal(elements.get('address').value, 'B1');
  assert.equal(elements.get('formula').value, 'b', 'an already-rendered target cell supplies its value without another range render');
});

test('Shift+Tab은 범위를 확장하지 않고 이전 셀만 선택한다', () => {
  const { document, documentListeners, elements } = runGridWebview(renderHtml());
  let cell = document.querySelector('[data-row="0"][data-column="1"]');
  cell.onclick();
  documentListeners.get('keydown')({ target: cell, key: 'Tab', shiftKey: true, preventDefault() {} });
  cell = document.querySelector('[data-row="0"][data-column="0"]');
  assert.equal(elements.get('address').value, 'A1');
  assert.equal(document.activeElement, cell);
  assert.equal(cell.classList.contains('in-range'), true);
  assert.equal(document.querySelector('[data-row="0"][data-column="1"]').classList.contains('in-range'), false);
});

test('화살표 이동은 셀을 화면 안에 유지하고 불필요하게 행 맨 위로 점프하지 않는다', () => {
  const { document, documentListeners, elements, messages, respondToRange } = runGridWebview(renderHtml(), { rowCount: 100 });
  let cell = document.querySelector('[data-row="0"][data-column="0"]');
  cell.onclick();
  const initialRangeRequests = messages.filter(message => message.type === 'range').length;
  const keydown = documentListeners.get('keydown');
  for (let row = 1; row <= 14; row++) {
    const previousRangeRequests = messages.filter(message => message.type === 'range').length;
    keydown({ target: cell, key: 'ArrowDown', preventDefault() {} });
    cell = document.querySelector('[data-row="' + row + '"][data-column="0"]');
    if (messages.filter(message => message.type === 'range').length > previousRangeRequests) respondToRange();
    cell = document.querySelector('[data-row="' + row + '"][data-column="0"]');
    assert.equal(document.activeElement, cell);
  }
  assert.equal(elements.get('address').value, 'A15');
  assert.equal(elements.get('grid').scrollTop, 120);
  assert.equal(messages.filter(message => message.type === 'range').length, initialRangeRequests, 'already-rendered rows should not force another range request');
});

test('표시 행 매핑은 숨긴 행을 건너뛰고 실제 행 주소를 유지한다', () => {
  const { document, documentListeners, elements, messages } = runGridWebview(renderHtml(), { rowCount: 8, columnCount: 2, visibleRows: [0, 2, 5, 7] });
  assert.equal(elements.get('canvas').style.height, '134px');
  assert.equal(document.querySelector('[data-row="2"][data-column="0"]').style.top, '56px');
  assert.equal(document.querySelector('[data-row="5"][data-column="0"]').style.top, '82px');
  documentListeners.get('keydown')({ target: elements.get('grid'), key: 'ArrowDown', preventDefault() {} });
  assert.equal(elements.get('address').value, 'A3');
  assert.ok(messages.filter(message => message.type === 'range').at(-1).viewRows.includes(2));
  documentListeners.get('keydown')({ target: elements.get('grid'), key: 'ArrowDown', ctrlKey: true, preventDefault() {} });
  assert.deepEqual(messages.filter(message => message.type === 'navigate').at(-1).visibleRows, [0, 2, 5, 7]);
});

test('XLSX 희소 숨김 행은 압축 렌더링·화살표·복사 메시지에 반영된다', () => {
  const { document, documentListeners, elements, messages } = runGridWebview(renderHtml(), { kind: 'xlsx', rowCount: 8, columnCount: 2, hiddenRows: [1, 3] });
  assert.equal(elements.get('canvas').style.height, String(30 + (1048576 - 2) * 26) + 'px');
  assert.equal(document.querySelector('[data-row="2"][data-column="0"]').style.top, '56px');
  documentListeners.get('keydown')({ target: elements.get('grid'), key: 'ArrowDown', preventDefault() {} });
  assert.equal(elements.get('address').value, 'A3');
  documentListeners.get('keydown')({ target: elements.get('grid'), key: 'c', ctrlKey: true, preventDefault() {} });
  const copyRequest = messages.filter(message => message.type === 'copy').at(-1);
  assert.deepEqual(copyRequest.hiddenRows, [1, 3]);
  assert.equal(copyRequest.visibleRows, undefined);
});

test('필터된 표시 행에서 셀 드래그는 화면 순번을 실제 행 주소로 바꿔 선택한다', () => {
  const { document, documentListeners, elements, messages } = runGridWebview(renderHtml(), { rowCount: 8, columnCount: 2, visibleRows: [0, 2, 5, 7] });
  const first = document.querySelector('[data-row="0"][data-column="0"]');
  first.onmousedown({ target: first, shiftKey: false, clientX: 80, clientY: 60, preventDefault() {} });
  documentListeners.get('mousemove')({ clientX: 80, clientY: 112 });
  assert.equal(elements.get('address').value, 'A8');
  assert.equal(document.querySelector('[data-row="7"][data-column="0"]').classList.contains('in-range'), true);
  documentListeners.get('mouseup')();
  documentListeners.get('keydown')({ target: elements.get('grid'), key: 'c', ctrlKey: true, preventDefault() {} });
  const copyRequest = messages.filter(message => message.type === 'copy').at(-1);
  assert.deepEqual(copyRequest.visibleRows, [0, 2, 5, 7]);
  assert.equal(copyRequest.range.startRow, 0);
  assert.equal(copyRequest.range.endRow, 7);
  let active = document.querySelector('[data-row="7"][data-column="0"]');
  active.ondblclick({ target: active, clientX: 80, clientY: 112 });
  active.children[0].onkeydown({ key: 'Enter', shiftKey: true, preventDefault() {} });
  assert.equal(elements.get('address').value, 'A6', 'Shift+Enter moves to the previous visible row inside the selection');
  active = document.querySelector('[data-row="5"][data-column="0"]');
  active.ondblclick({ target: active, clientX: 80, clientY: 86 });
  active.children[0].onkeydown({ key: 'Tab', shiftKey: false, preventDefault() {}, stopPropagation() {} });
  assert.equal(elements.get('address').value, 'A8', 'Tab advances to the next visible row when the selection has one column');
  active = document.querySelector('[data-row="7"][data-column="0"]');
  active.ondblclick({ target: active, clientX: 80, clientY: 112 });
  active.children[0].onkeydown({ key: 'Tab', shiftKey: false, preventDefault() {}, stopPropagation() {} });
  assert.equal(elements.get('address').value, 'A1', 'Tab wraps to the first visible row in the selection');
});

test('셀 드래그가 그리드 가장자리를 넘으면 스크롤과 선택 범위가 함께 확장된다', () => {
  const { document, documentListeners, elements } = runGridWebview(renderHtml(), { rowCount: 80, columnCount: 12 });
  const first = document.querySelector('[data-row="0"][data-column="0"]');
  first.onmousedown({ shiftKey: false, clientX: 80, clientY: 60, preventDefault() {} });
  documentListeners.get('mousemove')({ clientX: 510, clientY: 310 });
  assert.ok(elements.get('grid').scrollTop > 0);
  assert.ok(elements.get('grid').scrollLeft > 0);
  assert.notEqual(elements.get('address').value, 'A1');
  documentListeners.get('mouseup')();
});

test('Webview 행·열 삽입·삭제 버튼이 선택한 위치의 구조 변경을 요청한다', () => {
  const { document, documentListeners, elements, messages, respondToRange } = runGridWebview(renderHtml());
  const target = document.querySelector('[data-row="1"][data-column="1"]');
  target.onclick();
  for (const [id, expected] of [
    ['insertRow', { axis: 'row', action: 'insert', index: 1 }],
    ['deleteRow', { axis: 'row', action: 'delete', index: 1 }],
    ['insertColumn', { axis: 'column', action: 'insert', index: 1 }],
    ['deleteColumn', { axis: 'column', action: 'delete', index: 1 }]
  ]) {
    elements.get(id).onclick();
    assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'structure', ...expected, sheetIndex: 0 });
    respondToRange();
  }
  const first = document.querySelector('[data-row="0"][data-column="0"]');
  const last = document.querySelector('[data-row="1"][data-column="1"]');
  first.onmousedown({ shiftKey: false, clientX: 80, clientY: 60, preventDefault() {} });
  last.onmouseenter();
  documentListeners.get('mouseup')();
  elements.get('insertRow').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'structure', axis: 'row', action: 'insert', index: 0, count: 2, sheetIndex: 0 });
  elements.get('deleteColumn').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'structure', axis: 'column', action: 'delete', index: 0, count: 2, sheetIndex: 0 });
});

test('행 번호와 열 머리글은 전체 축을 선택하고 Shift+클릭으로 축 범위를 확장한다', () => {
  const { document, elements, messages } = runGridWebview(renderHtml(), { rowCount: 4, columnCount: 4 });
  const rowHeaders = elements.get('canvas').children.filter(item => item.classList.contains('rowhead'));
  rowHeaders[1].onclick({ shiftKey: false });
  assert.equal(document.querySelector('[data-row="1"][data-column="0"]').classList.contains('in-range'), true);
  assert.equal(document.querySelector('[data-row="1"][data-column="3"]').classList.contains('in-range'), true);
  rowHeaders[3].onclick({ shiftKey: true });
  assert.equal(document.querySelector('[data-row="2"][data-column="2"]').classList.contains('in-range'), true);
  assert.equal(document.querySelector('[data-row="0"][data-column="2"]').classList.contains('in-range'), false);
  elements.get('insertRow').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'structure', axis: 'row', action: 'insert', index: 1, count: 3, sheetIndex: 0 });

  const columnHeaders = elements.get('canvas').children.filter(item => item.classList.contains('colhead'));
  columnHeaders[1].onclick({ target: columnHeaders[1], shiftKey: false });
  columnHeaders[2].onclick({ target: columnHeaders[2], shiftKey: true });
  assert.equal(document.querySelector('[data-row="3"][data-column="1"]').classList.contains('in-range'), true);
  assert.equal(document.querySelector('[data-row="3"][data-column="2"]').classList.contains('in-range'), true);
  assert.equal(document.querySelector('[data-row="3"][data-column="0"]').classList.contains('in-range'), false);
  elements.get('deleteColumn').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'structure', axis: 'column', action: 'delete', index: 1, count: 2, sheetIndex: 0 });

  const dragged = runGridWebview(renderHtml(), { rowCount: 4, columnCount: 4 });
  const draggedRows = dragged.elements.get('canvas').children.filter(item => item.classList.contains('rowhead'));
  draggedRows[0].onmousedown({ clientX: 20, clientY: 40, shiftKey: false, preventDefault() {} });
  draggedRows[2].onmouseenter();
  dragged.documentListeners.get('mouseup')();
  draggedRows[2].onclick({ shiftKey: false });
  assert.equal(dragged.document.querySelector('[data-row="1"][data-column="3"]').classList.contains('in-range'), true);
  assert.equal(dragged.document.querySelector('[data-row="3"][data-column="3"]').classList.contains('in-range'), false);

  const draggedColumns = dragged.elements.get('canvas').children.filter(item => item.classList.contains('colhead'));
  draggedColumns[0].onmousedown({ target: draggedColumns[0], clientX: 80, clientY: 20, shiftKey: false, preventDefault() {} });
  draggedColumns[2].onmouseenter({ target: draggedColumns[2] });
  dragged.documentListeners.get('mouseup')();
  draggedColumns[2].onclick({ target: draggedColumns[2], shiftKey: false });
  assert.equal(dragged.document.querySelector('[data-row="3"][data-column="1"]').classList.contains('in-range'), true);
  assert.equal(dragged.document.querySelector('[data-row="3"][data-column="3"]').classList.contains('in-range'), false);
});

test('Shift+Space, Ctrl+Space, Ctrl+Shift+Space는 행·열·전체 시트를 선택한다', () => {
  const { document, documentListeners, elements, messages, dispatch } = runGridWebview(renderHtml(), { rowCount: 4, columnCount: 4 });
  const cell = document.querySelector('[data-row="1"][data-column="2"]');
  cell.onclick();
  const keydown = documentListeners.get('keydown');
  let prevented = false;
  keydown({ target: cell, key: ' ', shiftKey: true, ctrlKey: false, metaKey: false, preventDefault() { prevented = true; } });
  assert.equal(prevented, true);
  assert.equal(elements.get('address').value, 'C2');
  assert.equal(document.querySelector('[data-row="1"][data-column="0"]').classList.contains('in-range'), true);
  assert.equal(document.querySelector('[data-row="1"][data-column="3"]').classList.contains('in-range'), true);
  assert.equal(document.querySelector('[data-row="0"][data-column="2"]').classList.contains('in-range'), false);

  prevented = false;
  keydown({ target: cell, key: ' ', shiftKey: false, ctrlKey: true, metaKey: false, preventDefault() { prevented = true; } });
  assert.equal(prevented, true);
  assert.equal(document.querySelector('[data-row="0"][data-column="2"]').classList.contains('in-range'), true);
  assert.equal(document.querySelector('[data-row="3"][data-column="2"]').classList.contains('in-range'), true);
  assert.equal(document.querySelector('[data-row="1"][data-column="1"]').classList.contains('in-range'), false);

  keydown({ target: cell, key: ' ', shiftKey: true, ctrlKey: true, metaKey: false, preventDefault() {} });
  const sheetRequest = messages.at(-1);
  assert.deepEqual(JSON.parse(JSON.stringify(sheetRequest)), { type: 'current-region', row: 1, column: 2, includeSheet: true, requestId: sheetRequest.requestId, sheetIndex: 0 });
  dispatch({ type: 'current-region', requestId: sheetRequest.requestId, row: 1, column: 2, sheetIndex: 0, range: { top: 0, bottom: 3, left: 0, right: 3 } });
  assert.equal(elements.get('address').value, 'C2', 'whole-sheet selection preserves the active cell');
  assert.equal(document.querySelector('[data-row="0"][data-column="0"]').classList.contains('in-range'), true);
  assert.equal(document.querySelector('[data-row="3"][data-column="3"]').classList.contains('in-range'), true);
});

test('XLSX 행·열 머리글 선택은 사용 영역이 아닌 전체 워크시트 축을 선택하고 Delete를 희소 처리한다', () => {
  const rowGrid = runGridWebview(renderHtml(), { rowCount: 3, columnCount: 3, kind: 'xlsx' });
  const rowCell = rowGrid.document.querySelector('[data-row="1"][data-column="0"]');
  rowCell.onclick();
  rowGrid.documentListeners.get('keydown')({ target: rowCell, key: ' ', shiftKey: true, ctrlKey: false, metaKey: false, preventDefault() {} });
  rowGrid.documentListeners.get('keydown')({ target: rowCell, key: 'Delete', preventDefault() {} });
  assert.deepEqual(JSON.parse(JSON.stringify(rowGrid.messages.at(-1).range)), { startRow: 1, endRow: 1, startColumn: 0, endColumn: 16383 });

  const columnGrid = runGridWebview(renderHtml(), { rowCount: 3, columnCount: 3, kind: 'xlsx' });
  const columnCell = columnGrid.document.querySelector('[data-row="1"][data-column="1"]');
  columnCell.onclick();
  columnGrid.documentListeners.get('keydown')({ target: columnCell, key: ' ', shiftKey: false, ctrlKey: true, metaKey: false, preventDefault() {} });
  columnGrid.documentListeners.get('keydown')({ target: columnCell, key: 'Delete', preventDefault() {} });
  assert.deepEqual(JSON.parse(JSON.stringify(columnGrid.messages.at(-1).range)), { startRow: 0, endRow: 1048575, startColumn: 1, endColumn: 1 });

  const bytes = minimalXlsx('', '<row r="1"><c r="A1"><f>1+1</f><v>2</v></c><c r="C1" t="inlineStr"><is><t>text</t></is></c></row>');
  const document = new JinSheetDocument(uri('/tmp/full-row-clear.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  const edits = buildClearEdits(document, 0, 0, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 16383 });
  assert.deepEqual(edits.map(({ row, column, value, formula }) => ({ row, column, value, formula })), [
    { row: 0, column: 0, value: '', formula: null }, { row: 0, column: 2, value: '', formula: undefined }
  ]);
  let clearOperation;
  document.onDidChange(change => { clearOperation = change; });
  document.applyCellEdits(edits, 0, '전체 행 지우기');
  assert.equal(document.getCell(0, 0).formula, null);
  assert.equal(document.getCell(0, 2).value, '');
  clearOperation.undo();
  assert.equal(document.getCell(0, 0).formula, '1+1');
  assert.equal(document.getCell(0, 2).value, 'text');
  clearOperation.redo();
  assert.equal(document.getCell(0, 0).formula, null);
  const savedRow = analyze(document.createOutput());
  assert.equal(savedRow.sheets[0].cells.get('A1').formula, null);
  assert.equal(savedRow.sheets[0].cells.get('C1').value, '');
  const copiedRow = buildCopyPayload(document, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 16383 });
  assert.equal(copiedRow.cells[0].length, 16384);
  const pastedRow = buildPasteEdits(parseClipboard(copiedRow.text), 1, 0);
  assert.equal(pastedRow.length, 16384, 'a full Excel row remains copyable and pasteable within the cell operation limit');
  assert.equal(buildMoveEdits({ kind: 'xlsx', getCell: () => null }, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 16383 }, 1, 0).length, 32768, 'a full Excel row can be moved with one destination and one source edit per column');
  assert.equal(buildPasteEdits([['fill']], 0, 0, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 16383 }).length, 16384, 'one copied cell fills a selected full row');
  document.dispose();
});

test('XLSX 전체 열 복사·붙여넣기는 희소 셀만 수정하고 빈 칸·수식·저장·undo/redo를 보존한다', () => {
  const rows = '<row r="1"><c r="A1"><f>B1</f><v>stale</v></c><c r="B1" t="inlineStr"><is><t>replace</t></is></c></row>' +
    '<row r="2"><c r="B2" t="inlineStr"><is><t>clear me</t></is></c></row>' +
    '<row r="3"><c r="A3" s="2" t="inlineStr"><is><t>source</t></is></c><c r="B3" s="1" t="inlineStr"><is><t>replace me</t></is></c></row>';
  const bytes = minimalXlsx('', rows), document = new JinSheetDocument(uri('/tmp/full-column-clipboard.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  const fullColumn = { startRow: 0, endRow: 1048575, startColumn: 0, endColumn: 0 };
  const copied = buildCopyPayload(document, fullColumn);
  assert.equal(copied.sparse, true);
  assert.equal(copied.text, null, 'a million row TSV is not materialized in memory or sent to the OS clipboard');
  assert.deepEqual(copied.cells.map(({ rowOffset, value, formula }) => ({ rowOffset, value, formula })), [
    { rowOffset: 0, value: 'stale', formula: 'B1' },
    { rowOffset: 2, value: 'source', formula: null }
  ]);
  const edits = buildSparseColumnPasteEdits(document, copied, 0, 1, { startRow: 0, endRow: 1048575, startColumn: 1, endColumn: 1 });
  assert.equal(edits.length, 3, 'only source cells and occupied destination cells become edits');
  assert.deepEqual(edits.map(({ row, value, formula }) => ({ row, value, formula })), [
    { row: 0, value: '', formula: 'C1' },
    { row: 1, value: '', formula: null },
    { row: 2, value: 'source', formula: null }
  ]);
  assert.equal(edits[2].styleIndex, '2', 'cell formatting travels with the sparse source cell');
  let operation;
  document.onDidChange(change => { operation = change; });
  document.applyCellEdits(edits);
  assert.equal(document.getCell(0, 1).formula, 'C1');
  assert.equal(document.getCell(1, 1).value, '');
  assert.equal(document.getCell(2, 1).value, 'source');
  operation.undo();
  assert.equal(document.getCell(0, 1).value, 'replace');
  assert.equal(document.getCell(1, 1).value, 'clear me');
  operation.redo();
  const saved = analyze(document.createOutput());
  assert.equal(saved.sheets[0].cells.get('B1').formula, 'C1');
  assert.equal(saved.sheets[0].cells.get('B2').value, '');
  assert.equal(saved.sheets[0].cells.get('B3').value, 'source');
  assert.equal(saved.sheets[0].cells.get('B3').style, '2');
  document.dispose();
});

test('XLSX 전체 열 잘라내기 이동은 값·수식·서식을 옮기고 원본과 대상 빈 셀을 처리한다', () => {
  const rows = '<row r="1"><c r="A1"><f>B1</f><v>cached</v></c><c r="B1" t="inlineStr"><is><t>replace</t></is></c></row>' +
    '<row r="2"><c r="B2" t="inlineStr"><is><t>clear</t></is></c></row>' +
    '<row r="3"><c r="A3" s="2" t="inlineStr"><is><t>move</t></is></c><c r="B3" s="1" t="inlineStr"><is><t>replace</t></is></c></row>';
  const bytes = minimalXlsx('', rows), document = new JinSheetDocument(uri('/tmp/full-column-move.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  const sourceRange = { startRow: 0, endRow: 1048575, startColumn: 0, endColumn: 0 };
  assert.deepEqual(buildSparseColumnMoveEdits(document, sourceRange, 0, 0), [], 'moving a whole column back to itself changes nothing');
  const edits = buildSparseColumnMoveEdits(document, sourceRange, 0, 1);
  assert.equal(edits.length, 6);
  assert.deepEqual(edits.slice(0, 2).map(({ row, column, value, formula }) => ({ row, column, value, formula })), [
    { row: 0, column: 1, value: '', formula: 'C1' }, { row: 0, column: 0, value: '', formula: null }
  ]);
  assert.equal(edits[4].value, 'move');
  assert.equal(edits[4].styleIndex, '2');
  let operation;
  document.onDidChange(change => { operation = change; });
  document.applyCellEdits(edits, 0, '전체 열 이동');
  assert.equal(document.getCell(0, 0).formula, null);
  assert.equal(document.getCell(0, 1).formula, 'C1');
  assert.equal(document.getCell(1, 1).value, '');
  assert.equal(document.getCell(2, 1).style, '2');
  operation.undo();
  assert.equal(document.getCell(0, 0).formula, 'B1');
  assert.equal(document.getCell(1, 1).value, 'clear');
  operation.redo();
  const saved = analyze(document.createOutput());
  assert.equal(saved.sheets[0].cells.get('A1').formula, null);
  assert.equal(saved.sheets[0].cells.get('B1').formula, 'C1');
  assert.equal(saved.sheets[0].cells.get('B3').style, '2');
  document.dispose();
});

test('XLSX 전체 행 Ctrl+R 채우기는 16,384개 열까지 첫 셀 값과 서식을 채운다', () => {
  const document = { kind: 'xlsx', getCell: (row, column) => row === 0 && column === 0 ? { value: 'seed', style: '3', styleProperties: { bold: true } } : null };
  const edits = buildFillEdits(document, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 16383 }, 'right');
  assert.equal(edits.length, 16383);
  assert.deepEqual([edits[0].row, edits[0].column, edits[0].value, edits[0].styleIndex], [0, 1, 'seed', '3']);
  assert.deepEqual([edits.at(-1).row, edits.at(-1).column, edits.at(-1).value, edits.at(-1).styleIndex], [0, 16383, 'seed', '3']);
});

test('XLSX 전체 행 이동은 32,768 edit 안에서 저장·undo·redo한다', () => {
  const bytes = minimalXlsx('', '<row r="1"><c r="A1" t="inlineStr"><is><t>source</t></is></c><c r="C1" t="inlineStr"><is><t>destination</t></is></c></row>');
  const document = new JinSheetDocument(uri('/tmp/full-row-move.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  const range = { startRow: 0, endRow: 0, startColumn: 0, endColumn: 16383 }, edits = buildMoveEdits(document, range, 1, 0);
  assert.equal(edits.length, 32768);
  let operation;
  document.onDidChange(change => { operation = change; });
  document.applyCellEdits(edits, 0, '전체 행 이동');
  assert.equal(document.getCell(0, 0).value, '');
  assert.equal(document.getCell(1, 0).value, 'source');
  assert.equal(document.getCell(0, 2).value, '');
  assert.equal(document.getCell(1, 2).value, 'destination');
  operation.undo();
  assert.equal(document.getCell(0, 0).value, 'source');
  assert.equal(document.getCell(0, 2).value, 'destination');
  operation.redo();
  const saved = analyze(document.createOutput());
  assert.equal(saved.sheets[0].cells.get('A1').value, '');
  assert.equal(saved.sheets[0].cells.get('A2').value, 'source');
  assert.equal(saved.sheets[0].cells.get('C2').value, 'destination');
  document.dispose();
});

test('그리드 Tab과 Shift+Tab은 선택 범위 안에서 순환하며 범위를 유지한다', () => {
  const { document, documentListeners, elements } = runGridWebview(renderHtml(), { rowCount: 3, columnCount: 3 });
  const first = document.querySelector('[data-row="0"][data-column="0"]');
  const last = document.querySelector('[data-row="1"][data-column="1"]');
  first.onmousedown({ shiftKey: false, clientX: 80, clientY: 60, preventDefault() {} });
  last.onmouseenter();
  documentListeners.get('mouseup')();
  const keydown = documentListeners.get('keydown');
  keydown({ target: last, key: 'Tab', shiftKey: false, preventDefault() {} });
  assert.equal(elements.get('address').value, 'A1', 'Tab wraps from the final selected cell to the first');
  assert.equal(first.classList.contains('in-range'), true);
  const active = document.querySelector('[data-row="0"][data-column="0"]');
  keydown({ target: active, key: 'Tab', shiftKey: true, preventDefault() {} });
  assert.equal(elements.get('address').value, 'B2', 'Shift+Tab wraps from the first selected cell to the last');
  assert.equal(last.classList.contains('in-range'), true);
  keydown({ target: last, key: 'Enter', shiftKey: false, preventDefault() {} });
  assert.equal(elements.get('address').value, 'B1', 'Enter advances within the selected range and wraps at its bottom edge');
  assert.equal(last.classList.contains('in-range'), true, 'Enter keeps the full selected range');
  const firstColumnTop = document.querySelector('[data-row="0"][data-column="0"]');
  keydown({ target: document.querySelector('[data-row="0"][data-column="1"]'), key: 'Enter', shiftKey: true, preventDefault() {} });
  assert.equal(elements.get('address').value, 'B2', 'Shift+Enter moves upward and wraps at the top edge');
  assert.equal(firstColumnTop.classList.contains('in-range'), true);
  const below = document.querySelector('[data-row="2"][data-column="2"]');
  below.onmousedown({ shiftKey: false, clientX: 80, clientY: 60, preventDefault() {} });
  documentListeners.get('mouseup')();
  assert.equal(elements.get('address').value, 'C3', 'clicking outside a selected range collapses it to the clicked cell');
  keydown({ target: below, key: 'Enter', shiftKey: true, preventDefault() {} });
  assert.equal(elements.get('address').value, 'C2', 'Shift+Enter from one cell moves upward');
});

test('행·열 머리글을 가장자리까지 드래그하면 같은 축으로 자동 스크롤한다', () => {
  const rows = runGridWebview(renderHtml(), { rowCount: 100, columnCount: 4 });
  const rowHeader = rows.elements.get('canvas').children.find(item => item.classList.contains('rowhead'));
  rowHeader.onmousedown({ clientX: 20, clientY: 40, shiftKey: false, preventDefault() {} });
  rows.documentListeners.get('mousemove')({ clientX: 20, clientY: 310 });
  assert.ok(rows.elements.get('grid').scrollTop > 0);
  assert.notEqual(rows.elements.get('address').value, 'A1');
  rows.documentListeners.get('mouseup')();

  const columns = runGridWebview(renderHtml(), { rowCount: 4, columnCount: 100 });
  const columnHeader = columns.elements.get('canvas').children.find(item => item.classList.contains('colhead'));
  columnHeader.onmousedown({ target: columnHeader, clientX: 80, clientY: 20, shiftKey: false, preventDefault() {} });
  columns.documentListeners.get('mousemove')({ clientX: 510, clientY: 20 });
  assert.ok(columns.elements.get('grid').scrollLeft > 0);
  assert.notEqual(columns.elements.get('address').value, 'A1');
  columns.documentListeners.get('mouseup')();
});

test('Webview 열 머리글에는 정렬 동작이 없고 검색 UI는 유지한다', () => {
  const { elements, messages } = runGridWebview(renderHtml());
  assert.equal(elements.has('clearSort'), false);
  assert.equal(elements.has('filterValuesButton'), true, 'existing AutoFilter columns expose a value list control');
  assert.equal(elements.has('filterPanel'), true, 'the toolbar contains an accessible filter value panel');
  assert.doesNotMatch(renderHtml(), /id="clearSort"|sortHint|type:'sort'|type:'clear-sort'|id="filter"/);
  const header = elements.get('canvas').children.find(child => child.classList.contains('colhead'));
  assert.equal(header.textContent, 'A');
  assert.equal(typeof header.onclick, 'function', 'column headers select the used column without sorting it');
  header.onclick({ target: header, shiftKey: false });
  assert.equal(elements.get('canvas').children.find(child => child.classList.contains('cell') && child.dataset.column === '0').classList.contains('in-range'), true);
  assert.equal(messages.some(message => message.type === 'sort' || message.type === 'clear-sort'), false);
});

test('XLSX 도구 모음은 선택 행 숨김과 수동 숨김 행 모두 표시를 요청한다', () => {
  const grid = runGridWebview(renderHtml(), { kind: 'xlsx', rowCount: 3, columnCount: 2, hiddenRows: [1], manuallyHiddenRows: [1] });
  assert.equal(grid.elements.get('hideRows').disabled, false);
  grid.elements.get('hideRows').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(grid.messages.at(-1))), { type: 'hidden-rows', action: 'hide', startRow: 0, endRow: 0, sheetIndex: 0 });
  assert.equal(grid.elements.get('unhideRows').disabled, false);
  grid.elements.get('unhideRows').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(grid.messages.at(-1))), { type: 'hidden-rows', action: 'unhide-all', sheetIndex: 0 });
  const csv = runGridWebview(renderHtml(), { kind: 'csv', rowCount: 3, columnCount: 2, manuallyHiddenRows: [1] });
  assert.equal(csv.elements.get('hideRows').disabled, false);
  csv.elements.get('hideRows').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(csv.messages.at(-1))), { type: 'hidden-rows', action: 'hide', startRow: 0, endRow: 0, sheetIndex: 0 });
  assert.equal(csv.elements.get('unhideRows').disabled, false);
  csv.elements.get('unhideRows').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(csv.messages.at(-1))), { type: 'hidden-rows', action: 'unhide-all', sheetIndex: 0 });
});

test('XLSX 도구 모음은 선택 열 숨김과 수동 숨김 열 모두 표시를 요청한다', () => {
  const grid = runGridWebview(renderHtml(), { kind: 'xlsx', rowCount: 3, columnCount: 3, hiddenColumns: [1], manuallyHiddenColumns: [1] });
  assert.equal(grid.elements.get('hideColumns').disabled, false);
  grid.elements.get('hideColumns').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(grid.messages.at(-1))), { type: 'hidden-columns', action: 'hide', startColumn: 0, endColumn: 0, sheetIndex: 0 });
  assert.equal(grid.elements.get('unhideColumns').disabled, false);
  grid.elements.get('unhideColumns').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(grid.messages.at(-1))), { type: 'hidden-columns', action: 'unhide-all', sheetIndex: 0 });
  const csv = runGridWebview(renderHtml(), { kind: 'csv', rowCount: 3, columnCount: 2, manuallyHiddenColumns: [1] });
  assert.equal(csv.elements.get('hideColumns').disabled, false);
  csv.elements.get('hideColumns').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(csv.messages.at(-1))), { type: 'hidden-columns', action: 'hide', startColumn: 0, endColumn: 0, sheetIndex: 0 });
  assert.equal(csv.elements.get('unhideColumns').disabled, false);
  csv.elements.get('unhideColumns').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(csv.messages.at(-1))), { type: 'hidden-columns', action: 'unhide-all', sheetIndex: 0 });
  const laidOut = runGridWebview(renderHtml(), { kind: 'xlsx', rowCount: 2, columnCount: 3, hiddenColumns: [1] });
  const hiddenHeader = laidOut.elements.get('canvas').children.find(child => child.classList.contains('colhead') && child.dataset.column === '1');
  assert.equal(hiddenHeader.style.width, '0px');
  const visible = runGridWebview(renderHtml(), { kind: 'xlsx', rowCount: 2, columnCount: 3 });
  assert.equal(Number.parseFloat(visible.elements.get('canvas').style.width) - Number.parseFloat(laidOut.elements.get('canvas').style.width), 144);
});

test('XLSX 도구 모음은 틀 고정·첫 행·첫 열·해제를 요청한다', () => {
  const grid = runGridWebview(renderHtml(), { kind: 'xlsx', rowCount: 3, columnCount: 3, freezePane: { rows: 0, columns: 0 } });
  grid.elements.get('freezeTopRow').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(grid.messages.at(-1))), { type: 'freeze-pane', rows: 1, columns: 0, sheetIndex: 0 });
  grid.elements.get('freezeFirstColumn').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(grid.messages.at(-1))), { type: 'freeze-pane', rows: 0, columns: 1, sheetIndex: 0 });
  grid.elements.get('freezePanes').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(grid.messages.at(-1))), { type: 'freeze-pane', rows: 0, columns: 0, sheetIndex: 0 });
  grid.elements.get('unfreezePanes').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(grid.messages.at(-1))), { type: 'freeze-pane', rows: 0, columns: 0, sheetIndex: 0 });
});

test('행·열 그룹 도구는 선택 범위 그룹화·접기·펼치기·해제 요청과 상태를 반영한다', () => {
  const grid = runGridWebview(renderHtml(), { kind: 'xlsx', rowCount: 3, columnCount: 3 });
  const rowGroup = { start: 0, end: 0, level: 1, summary: 1, collapsed: false };
  grid.elements.get('groupRows').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(grid.messages.at(-1))), { type: 'outline-groups', axis: 'row', action: 'group', start: 0, end: 0, sheetIndex: 0 });
  grid.dispatch({ type: 'model', kind: 'xlsx', rowCount: 3, columnCount: 3, rowGroups: [rowGroup], sheetIndex: 0, dirty: true });
  grid.respondToRange();
  assert.deepEqual(grid.elements.get('rowOutlineLevels').children.map(button => button.textContent), ['1', '2']);
  grid.elements.get('rowOutlineLevels').children[0].onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(grid.messages.at(-1))), { type: 'outline-groups', axis: 'row', action: 'show-level', level: 1, sheetIndex: 0 });
  assert.equal(grid.elements.get('collapseRows').disabled, false);
  assert.equal(grid.elements.get('expandRows').disabled, true);
  const rowSummary = grid.elements.get('canvas').children.find(item => item.classList.contains('rowhead') && item.dataset.row === '1');
  const rowToggle = rowSummary.children.find(item => item.classList.contains('outline-toggle'));
  assert.equal(rowToggle.textContent, '−');
  rowToggle.onclick({ preventDefault() {}, stopPropagation() {} });
  assert.deepEqual(JSON.parse(JSON.stringify(grid.messages.at(-1))), { type: 'outline-groups', axis: 'row', action: 'collapse-summary', summary: 1, sheetIndex: 0 });
  grid.elements.get('collapseRows').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(grid.messages.at(-1))), { type: 'outline-groups', axis: 'row', action: 'collapse', start: 0, end: 0, sheetIndex: 0 });
  grid.dispatch({ type: 'model', kind: 'xlsx', rowCount: 3, columnCount: 3, rowGroups: [{ ...rowGroup, collapsed: true }], hiddenRows: [0], manuallyHiddenRows: [], sheetIndex: 0, dirty: true });
  grid.respondToRange();
  assert.equal(grid.elements.get('collapseRows').disabled, true);
  assert.equal(grid.elements.get('expandRows').disabled, false);
  const expandedRowSummary = grid.elements.get('canvas').children.find(item => item.classList.contains('rowhead') && item.dataset.row === '1');
  const expandRowToggle = expandedRowSummary.children.find(item => item.classList.contains('outline-toggle'));
  assert.equal(expandRowToggle.textContent, '+');
  expandRowToggle.onclick({ preventDefault() {}, stopPropagation() {} });
  assert.deepEqual(JSON.parse(JSON.stringify(grid.messages.at(-1))), { type: 'outline-groups', axis: 'row', action: 'expand-summary', summary: 1, sheetIndex: 0 });
  grid.elements.get('expandRows').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(grid.messages.at(-1))), { type: 'outline-groups', axis: 'row', action: 'expand', start: 0, end: 0, sheetIndex: 0 });
  grid.elements.get('ungroupRows').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(grid.messages.at(-1))), { type: 'outline-groups', axis: 'row', action: 'ungroup', start: 0, end: 0, sheetIndex: 0 });

  grid.elements.get('groupColumns').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(grid.messages.at(-1))), { type: 'outline-groups', axis: 'column', action: 'group', start: 0, end: 0, sheetIndex: 0 });
  grid.dispatch({ type: 'model', kind: 'xlsx', rowCount: 3, columnCount: 3, columnGroups: [{ start: 0, end: 0, level: 1, summary: 1, collapsed: false }], sheetIndex: 0, dirty: true });
  grid.respondToRange();
  assert.deepEqual(grid.elements.get('columnOutlineLevels').children.map(button => button.textContent), ['1', '2']);
  grid.elements.get('columnOutlineLevels').children[0].onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(grid.messages.at(-1))), { type: 'outline-groups', axis: 'column', action: 'show-level', level: 1, sheetIndex: 0 });
  assert.equal(grid.elements.get('collapseColumns').disabled, false);
  const columnSummary = grid.elements.get('canvas').children.find(item => item.classList.contains('colhead') && item.dataset.column === '1');
  const columnToggle = columnSummary.children.find(item => item.classList.contains('outline-toggle'));
  assert.equal(columnToggle.textContent, '−');
  columnToggle.onclick({ preventDefault() {}, stopPropagation() {} });
  assert.deepEqual(JSON.parse(JSON.stringify(grid.messages.at(-1))), { type: 'outline-groups', axis: 'column', action: 'collapse-summary', summary: 1, sheetIndex: 0 });
  grid.elements.get('collapseColumns').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(grid.messages.at(-1))), { type: 'outline-groups', axis: 'column', action: 'collapse', start: 0, end: 0, sheetIndex: 0 });
});

test('CSV·TSV 행·열 그룹은 요약 머리글에서 접고 펼칠 수 있고 세션 한정 안내를 표시한다', () => {
  const rowGroup = { start: 0, end: 0, level: 1, summary: 1, collapsed: false };
  const columnGroup = { ...rowGroup };
  const grid = runGridWebview(renderHtml(), { kind: 'csv', rowCount: 3, columnCount: 3, rowGroups: [rowGroup], columnGroups: [columnGroup] });
  assert.equal(grid.elements.get('csvOutlineSessionNote').hidden, false);

  const rowSummary = grid.elements.get('canvas').children.find(item => item.classList.contains('rowhead') && item.dataset.row === '1');
  const columnSummary = grid.elements.get('canvas').children.find(item => item.classList.contains('colhead') && item.dataset.column === '1');
  assert.equal(rowSummary.children.find(item => item.classList.contains('outline-toggle')).textContent, '−');
  assert.equal(columnSummary.children.find(item => item.classList.contains('outline-toggle')).textContent, '−');
  rowSummary.children.find(item => item.classList.contains('outline-toggle')).onclick({ preventDefault() {}, stopPropagation() {} });
  assert.deepEqual(JSON.parse(JSON.stringify(grid.messages.at(-1))), { type: 'outline-groups', axis: 'row', action: 'collapse-summary', summary: 1, sheetIndex: 0 });
  columnSummary.children.find(item => item.classList.contains('outline-toggle')).onclick({ preventDefault() {}, stopPropagation() {} });
  assert.deepEqual(JSON.parse(JSON.stringify(grid.messages.at(-1))), { type: 'outline-groups', axis: 'column', action: 'collapse-summary', summary: 1, sheetIndex: 0 });

  grid.dispatch({ type: 'model', kind: 'csv', rowCount: 3, columnCount: 3, rowGroups: [{ ...rowGroup, collapsed: true }], columnGroups: [{ ...columnGroup, collapsed: true }], hiddenRows: [0], hiddenColumns: [0], sheetIndex: 0, dirty: false });
  grid.respondToRange();
  const expandedRowSummary = grid.elements.get('canvas').children.find(item => item.classList.contains('rowhead') && item.dataset.row === '1');
  const expandedColumnSummary = grid.elements.get('canvas').children.find(item => item.classList.contains('colhead') && item.dataset.column === '1');
  assert.equal(expandedRowSummary.children.find(item => item.classList.contains('outline-toggle')).textContent, '+');
  assert.equal(expandedColumnSummary.children.find(item => item.classList.contains('outline-toggle')).textContent, '+');
});

test('중첩 그룹의 요약 머리글은 같은 위치의 모든 행·열 그룹을 접고 펼친다', () => {
  const groups = [
    { start: 0, end: 1, level: 1, summary: 2, collapsed: false },
    { start: 1, end: 1, level: 2, summary: 2, collapsed: false }
  ];
  const grid = runGridWebview(renderHtml(), { kind: 'xlsx', rowCount: 4, columnCount: 4, rowGroups: groups, columnGroups: groups });
  const rowSummary = grid.elements.get('canvas').children.find(item => item.classList.contains('rowhead') && item.dataset.row === '2');
  const columnSummary = grid.elements.get('canvas').children.find(item => item.classList.contains('colhead') && item.dataset.column === '2');
  const rowToggle = rowSummary.children.find(item => item.classList.contains('outline-toggle'));
  const columnToggle = columnSummary.children.find(item => item.classList.contains('outline-toggle'));
  assert.deepEqual(grid.elements.get('rowOutlineLevels').children.map(button => button.textContent), ['1', '2', '3']);
  assert.deepEqual(grid.elements.get('columnOutlineLevels').children.map(button => button.textContent), ['1', '2', '3']);
  const rowLevelOne = grid.elements.get('rowOutlineLevels').children[0];
  assert.equal(rowToggle.textContent, '−');
  assert.equal(columnToggle.textContent, '−');
  rowToggle.onclick({ preventDefault() {}, stopPropagation() {} });
  assert.deepEqual(JSON.parse(JSON.stringify(grid.messages.at(-1))), { type: 'outline-groups', axis: 'row', action: 'collapse-summary', summary: 2, sheetIndex: 0 });
  columnToggle.onclick({ preventDefault() {}, stopPropagation() {} });
  assert.deepEqual(JSON.parse(JSON.stringify(grid.messages.at(-1))), { type: 'outline-groups', axis: 'column', action: 'collapse-summary', summary: 2, sheetIndex: 0 });

  const collapsed = groups.map(group => ({ ...group, collapsed: true }));
  grid.dispatch({ type: 'model', kind: 'xlsx', rowCount: 4, columnCount: 4, rowGroups: collapsed, columnGroups: collapsed, hiddenRows: [0, 1], hiddenColumns: [0, 1], sheetIndex: 0, dirty: true });
  grid.respondToRange();
  assert.equal(grid.elements.get('rowOutlineLevels').children[0], rowLevelOne, 'redraws preserve the outline control node while the depth is unchanged');
  const collapsedRowSummary = grid.elements.get('canvas').children.find(item => item.classList.contains('rowhead') && item.dataset.row === '2');
  const collapsedColumnSummary = grid.elements.get('canvas').children.find(item => item.classList.contains('colhead') && item.dataset.column === '2');
  collapsedRowSummary.children.find(item => item.classList.contains('outline-toggle')).onclick({ preventDefault() {}, stopPropagation() {} });
  assert.deepEqual(JSON.parse(JSON.stringify(grid.messages.at(-1))), { type: 'outline-groups', axis: 'row', action: 'expand-summary', summary: 2, sheetIndex: 0 });
  collapsedColumnSummary.children.find(item => item.classList.contains('outline-toggle')).onclick({ preventDefault() {}, stopPropagation() {} });
  assert.deepEqual(JSON.parse(JSON.stringify(grid.messages.at(-1))), { type: 'outline-groups', axis: 'column', action: 'expand-summary', summary: 2, sheetIndex: 0 });
});

test('행·열 개요 머리글에 중첩 수준과 그룹 경계가 표시된다', () => {
  const groups = [
    { start: 0, end: 2, level: 1, summary: 3, collapsed: false },
    { start: 1, end: 2, level: 2, summary: 3, collapsed: false }
  ];
  const grid = runGridWebview(renderHtml(), { kind: 'xlsx', rowCount: 4, columnCount: 4, rowGroups: groups, columnGroups: groups });
  const rowHeaders = grid.elements.get('canvas').children.filter(item => item.classList.contains('rowhead'));
  const columnHeaders = grid.elements.get('canvas').children.filter(item => item.classList.contains('colhead'));
  const rowBrackets = row => rowHeaders.find(item => item.dataset.row === String(row)).children.filter(item => item.classList.contains('outline-row-bracket'));
  const columnBrackets = column => columnHeaders.find(item => item.dataset.column === String(column)).children.filter(item => item.classList.contains('outline-column-bracket'));

  assert.equal(rowBrackets(0).length, 1);
  assert.equal(rowBrackets(0)[0].dataset.level, '1');
  assert.equal(rowBrackets(0)[0].classList.contains('outline-bracket-start'), true);
  assert.equal(rowBrackets(1).length, 2);
  assert.equal(rowBrackets(1).find(item => item.dataset.level === '2').classList.contains('outline-bracket-start'), true);
  assert.equal(rowBrackets(2).every(item => item.classList.contains('outline-bracket-end')), true);

  assert.equal(columnBrackets(0).length, 1);
  assert.equal(columnBrackets(0)[0].dataset.level, '1');
  assert.equal(columnBrackets(0)[0].classList.contains('outline-bracket-start'), true);
  assert.equal(columnBrackets(1).length, 2);
  assert.equal(columnBrackets(1).find(item => item.dataset.level === '2').classList.contains('outline-bracket-start'), true);
  assert.equal(columnBrackets(2).every(item => item.classList.contains('outline-bracket-end')), true);
  assert.equal(columnHeaders.find(item => item.dataset.column === '0').classList.contains('has-outline-brackets'), true);
});

test('XLSX 도구 모음은 선택 셀의 자동 합계를 요청하고 CSV에서는 비활성화한다', () => {
  const grid = runGridWebview(renderHtml(), { kind: 'xlsx', rowCount: 3, columnCount: 2 });
  assert.equal(grid.elements.get('autoSum').disabled, false);
  grid.elements.get('autoSum').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(grid.messages.at(-1))), { type: 'auto-sum', row: 0, column: 0, sheetIndex: 0 });
  let prevented = false;
  grid.documentListeners.get('keydown')({ target: grid.elements.get('grid'), altKey: true, ctrlKey: false, metaKey: false, key: '=', preventDefault: () => { prevented = true; } });
  assert.equal(prevented, true);
  assert.deepEqual(JSON.parse(JSON.stringify(grid.messages.at(-1))), { type: 'auto-sum', row: 0, column: 0, sheetIndex: 0 });
  const csv = runGridWebview(renderHtml(), { kind: 'csv', rowCount: 3, columnCount: 2 });
  assert.equal(csv.elements.get('autoSum').disabled, true);
});

test('AutoFilter 값 UI는 필터 목록을 요청하고 선택한 조건을 적용한다', () => {
  const create = runGridWebview(renderHtml(), { kind: 'xlsx', rowCount: 5, columnCount: 2 });
  assert.equal(create.elements.get('filterValuesButton').disabled, false, 'a data region can start a new AutoFilter');
  create.elements.get('filterValuesButton').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(create.messages.at(-1))), { type: 'filter-options', row: 0, column: 0, sheetIndex: 0 });

  const { document, elements, messages, dispatch } = runGridWebview(renderHtml(), {
    kind: 'xlsx', rowCount: 5, columnCount: 1,
    autoFilter: { top: 0, bottom: 4, left: 0, right: 0, columns: [] }
  });
  const button = elements.get('filterValuesButton');
  assert.equal(button.disabled, false);
  button.onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'filter-options', row: 0, column: 0, sheetIndex: 0 });
  const filter = { top: 0, bottom: 4, left: 0, right: 0, columns: [] };
  dispatch({ type: 'filter-options', column: 0, sheetIndex: 0, filter, options: { values: ['apple', 'pear'], blanks: true, selectedValues: ['apple', 'pear'], includeBlanks: true } });
  assert.equal(elements.get('filterPanel').hidden, false);
  const firstOption = elements.get('filterValuesList').children[0].children[0];
  firstOption.checked = false;
  firstOption.onchange();
  elements.get('filterApply').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'auto-filter', row: 0, column: 0, values: ['pear'], blanks: true, sheetIndex: 0, requestId: 1 });
  assert.equal(elements.get('filterPanel').hidden, false, 'the editor stays open until the provider accepts the filter');
  dispatch({ type: 'filter-applied', sheetIndex: 0, column: 0, requestId: 1 });
  assert.equal(elements.get('filterPanel').hidden, true);
});

test('AutoFilter 값 UI는 긴 목록을 페이지로 렌더링하고 검색한 값만 빠르게 선택한다', () => {
  const values = Array.from({ length: 450 }, (_, index) => 'item-' + String(index).padStart(3, '0'));
  const { elements, messages, dispatch } = runGridWebview(renderHtml(), {
    kind: 'xlsx', rowCount: 452, columnCount: 1,
    autoFilter: { top: 0, bottom: 451, left: 0, right: 0, columns: [] }
  });
  elements.get('filterValuesButton').onclick();
  const filter = { top: 0, bottom: 451, left: 0, right: 0, columns: [] };
  dispatch({ type: 'filter-options', column: 0, sheetIndex: 0, filter, options: { values, blanks: false, selectedValues: values, includeBlanks: false } });
  assert.equal(elements.get('filterValuesList').children.length, 200);
  assert.equal(elements.get('filterShowMore').hidden, false);
  elements.get('filterShowMore').onclick();
  assert.equal(elements.get('filterValuesList').children.length, 400);
  elements.get('filterShowMore').onclick();
  assert.equal(elements.get('filterValuesList').children.length, 450);
  assert.equal(elements.get('filterShowMore').hidden, true);
  const search = elements.get('filterSearch');
  search.value = values.at(-1);
  search.oninput();
  assert.equal(elements.get('filterValuesList').children.length, 1);
  assert.equal(elements.get('filterValuesList').children[0].children[1].textContent, 'item-449');
  elements.get('filterSelectAll').checked = false;
  elements.get('filterSelectAll').onchange();
  elements.get('filterApply').onclick();
  const applied = messages.at(-1);
  assert.equal(applied.values.length, 449);
  assert.equal(applied.values.includes('item-449'), false);
});

test('조건 필터 UI는 확정 전에는 적용하지 않고 Enter·AND/OR·실패 복구를 지원한다', () => {
  const grid = runGridWebview(renderHtml(), { kind: 'xlsx', rowCount: 5, columnCount: 2 });
  const { elements, messages, dispatch } = grid; elements.get('filterValuesButton').onclick();
  const filter = { top: 0, bottom: 4, left: 0, right: 1, columns: [] };
  dispatch({ type: 'filter-options', column: 0, sheetIndex: 0, filter, options: { values: ['alpha', 'beta'], blanks: false, selectedValues: ['alpha', 'beta'], includeBlanks: false } });
  elements.get('filterMode').value = 'conditions'; elements.get('filterMode').onchange(); assert.equal(elements.get('filterValueControls').hidden, true);
  const count = messages.length; elements.get('filterOperator1').value = 'contains'; elements.get('filterConditionValue1').value = '한글'; assert.equal(messages.length, count);
  const event = { key: 'Enter', keyCode: 229, isComposing: true, target: { tagName: 'INPUT', type: 'text' }, preventDefault() {}, stopPropagation() {} };
  elements.get('filterPanel').onkeydown(event); assert.equal(messages.length, count);
  elements.get('filterPanel').onkeydown({ ...event, key: 'Escape' }); assert.equal(elements.get('filterPanel').hidden, false);
  const multiline = elements.get('filterConditionValue1'); multiline.value = '한글'; multiline.setSelectionRange(2, 2);
  elements.get('filterPanel').onkeydown({ ...event, keyCode: 13, isComposing: false, altKey: true, target: Object.assign(multiline, { tagName: 'TEXTAREA' }) }); assert.equal(multiline.value, '한글\n'); assert.equal(messages.length, count); multiline.value = '한글';
  elements.get('filterUseSecond').checked = true; elements.get('filterUseSecond').onchange(); elements.get('filterLogic').value = 'or'; elements.get('filterOperator2').value = 'beginsWith'; elements.get('filterConditionValue2').value = 'a~*';
  elements.get('filterPanel').onkeydown({ ...event, keyCode: 13, isComposing: false });
  const applied = JSON.parse(JSON.stringify(messages.at(-1))); assert.deepEqual(applied.custom, { and: false, conditions: [{ operator: 'equal', value: '*한글*' }, { operator: 'equal', value: 'a~**' }] });
  assert.equal(elements.get('filterPanel').hidden, false); assert.equal(elements.get('filterEditor').disabled, true);
  dispatch({ type: 'filter-error', sheetIndex: 0, column: 0, requestId: applied.requestId, message: 'Invalid condition' });
  assert.equal(elements.get('filterPanel').hidden, false); assert.equal(elements.get('filterEditor').disabled, false); assert.equal(elements.get('filterConditionValue1').value, '한글'); assert.equal(elements.get('notice').textContent, 'Invalid condition');
  elements.get('filterApply').onclick(); const retry = messages.at(-1); assert.ok(retry.requestId > applied.requestId);
  dispatch({ type: 'filter-applied', sheetIndex: 0, column: 0, requestId: applied.requestId }); assert.equal(elements.get('filterPanel').hidden, false, 'stale replies cannot close a new submission');
  dispatch({ type: 'filter-applied', sheetIndex: 0, column: 0, requestId: retry.requestId }); assert.equal(elements.get('filterPanel').hidden, true);
});

test('조건 필터 UI는 기존 와일드카드를 재표시하고 취소·이 열 지우기를 구분한다', () => {
  const { elements, messages, dispatch } = runGridWebview(renderHtml(), { kind: 'xlsx', rowCount: 5, columnCount: 1 });
  const filter = { top: 0, bottom: 4, left: 0, right: 0, columns: [] };
  const open = custom => { elements.get('filterValuesButton').onclick(); dispatch({ type: 'filter-options', column: 0, sheetIndex: 0, filter, options: { values: [], blanks: false, selectedValues: [], includeBlanks: false, custom } }); };
  open({ and: true, conditions: [{ operator: 'notEqual', value: '*test*' }, { operator: 'equal', value: '*~?' }] });
  assert.equal(elements.get('filterMode').value, 'conditions'); assert.equal(elements.get('filterOperator1').value, 'notContains'); assert.equal(elements.get('filterConditionValue1').value, 'test'); assert.equal(elements.get('filterOperator2').value, 'endsWith'); assert.equal(elements.get('filterConditionValue2').value, '~?');
  const count = messages.length; elements.get('filterCancel').onclick(); assert.equal(messages.length, count); assert.equal(elements.get('filterPanel').hidden, true);
  open({ and: false, conditions: [{ operator: 'equal', value: 'literal~*' }] }); assert.equal(elements.get('filterOperator1').value, 'equal'); assert.equal(elements.get('filterConditionValue1').value, 'literal~*');
  elements.get('filterClear').onclick(); assert.equal(messages.at(-1).clear, true); assert.equal(messages.at(-1).custom, undefined);
  dispatch({ type: 'model', kind: 'xlsx', readOnly: false, rowCount: 4, columnCount: 1, sheetIndex: 1 }); assert.equal(elements.get('filterPanel').hidden, true);
  const switched = messages.length; elements.get('filterApply').onclick(); assert.equal(messages.length, switched, 'criteria from the previous sheet cannot be applied to the new sheet');
});

test('모델 재렌더링은 그리드 포커스를 유지하고 조건 입력란 포커스를 빼앗지 않는다', () => {
  const grid = runGridWebview(renderHtml(), { kind: 'xlsx', rowCount: 5, columnCount: 2 });
  const model = { type: 'model', kind: 'xlsx', readOnly: false, rowCount: 5, columnCount: 2, sheetIndex: 0 };
  grid.document.querySelector('[data-row="0"][data-column="0"]').focus(); grid.dispatch(model);
  assert.equal(grid.document.activeElement, grid.elements.get('grid')); grid.respondToRange(); assert.equal(grid.document.activeElement.classList.contains('cell'), true);
  grid.elements.get('filterValuesButton').onclick(); grid.dispatch({ type: 'filter-options', column: 0, sheetIndex: 0, filter: { top: 0, bottom: 4, left: 0, right: 1, columns: [] }, options: { values: ['a'], selectedValues: ['a'], blanks: false, includeBlanks: false } });
  const input = grid.elements.get('filterConditionValue1'); input.focus(); input.value = '한글 검색'; grid.dispatch(model); grid.respondToRange();
  assert.equal(grid.document.activeElement, input); assert.equal(input.value, '한글 검색');
});

test('열 경계 드래그와 키보드 조절이 가상 그리드의 셀 위치를 함께 갱신한다', () => {
  const { document, documentListeners, elements, messages } = runGridWebview(renderHtml(), { rowCount: 2, columnCount: 4 });
  const canvas = elements.get('canvas');
  const headerA = canvas.children.find(item => item.classList.contains('colhead') && item.dataset.column === '0');
  const headerB = canvas.children.find(item => item.classList.contains('colhead') && item.dataset.column === '1');
  const cellA = canvas.children.find(item => item.classList.contains('cell') && item.dataset.column === '0');
  const cellB = canvas.children.find(item => item.classList.contains('cell') && item.dataset.column === '1');
  const handle = headerA.children.find(item => item.classList.contains('col-resize-handle'));
  assert.equal(handle.attributes.role, 'separator');
  assert.equal(handle.attributes['aria-valuenow'], '144');
  handle.onmousedown({ clientX: 144, preventDefault() {}, stopPropagation() {} });
  documentListeners.get('mousemove')({ clientX: 204 });
  assert.equal(headerA.style.width, '204px');
  assert.equal(cellA.style.width, '204px');
  assert.equal(headerB.style.left, '256px');
  assert.equal(cellB.style.left, '256px');
  assert.equal(handle.attributes['aria-valuenow'], '204');
  documentListeners.get('mouseup')();
  handle.onkeydown({ key: 'ArrowRight', preventDefault() {}, stopPropagation() {} });
  assert.equal(headerA.style.width, '214px');
  assert.equal(headerB.style.left, '266px');
  assert.equal(handle.attributes['aria-valuenow'], '214');
  assert.equal(messages.some(message => message.type === 'sort' || message.type === 'edit'), false);
});

test('열 너비는 시트별로 분리되고 편집기 탭 안에서 유지된다', () => {
  const { dispatch, documentListeners, elements, respondToRange } = runGridWebview(renderHtml(), { rowCount: 2, columnCount: 4 });
  const firstHeader = () => elements.get('canvas').children.find(item => item.classList.contains('colhead') && item.dataset.column === '0');
  const handle = firstHeader().children.find(item => item.classList.contains('col-resize-handle'));
  handle.onmousedown({ clientX: 144, preventDefault() {}, stopPropagation() {} });
  documentListeners.get('mousemove')({ clientX: 220 });
  documentListeners.get('mouseup')();
  assert.equal(firstHeader().style.width, '220px');
  dispatch({ type: 'model', kind: 'csv', readOnly: false, rowCount: 2, columnCount: 4, sheetIndex: 1, dirty: false });
  respondToRange();
  assert.equal(firstHeader().style.width, '144px');
  dispatch({ type: 'model', kind: 'csv', readOnly: false, rowCount: 2, columnCount: 4, sheetIndex: 0, dirty: false });
  respondToRange();
  assert.equal(firstHeader().style.width, '220px');
});


test('워크북 열 너비가 모델에서 그리드의 열·셀 크기에 반영된다', () => {
  const document = { kind: 'xlsx', reason: '', state: { sheets: [{ name: 'Data', visible: true, cells: new Map(), columnWidths: { 0: 24 } }] }, drafts: new Map(), writable: true, dirty: false, uri: { fsPath: '/tmp/width.xlsx' } };
  assert.deepEqual(viewModel(document).columnWidths, { 0: 24 });
  const { elements, messages } = runGridWebview(renderHtml(), { kind: 'xlsx', columnWidths: { 0: 24 } });
  const header = elements.get('canvas').children.find(item => item.classList.contains('colhead') && item.dataset.column === '0');
  const cell = elements.get('canvas').children.find(item => item.classList.contains('cell') && item.dataset.column === '0');
  assert.equal(header.style.width, '168px');
  assert.equal(cell.style.width, '168px');
  header.children.find(item => item.classList.contains('col-resize-handle')).onkeydown({ key: 'ArrowRight', preventDefault() {}, stopPropagation() {} });
  assert.deepEqual(JSON.parse(JSON.stringify(messages.find(message => message.type === 'column-width'))), { type: 'column-width', column: 0, width: 178, sheetIndex: 0 });
});

test('XLSX basic cell style travels through range data and reaches grid rendering', () => {
  const style = { color: '#123456', fill: '#FFEEAA', fontSize: 14, bold: true, italic: true, horizontal: 'center', vertical: 'bottom', borders: { left: { color: '#ABCDEF', line: 'thin' } } };
  const { elements } = runGridWebview(renderHtml(), { style, kind: 'xlsx' });
  const cell = elements.get('canvas').children.find(item => item.classList.contains('cell') && item.dataset.row === '1' && item.dataset.column === '1');
  assert.equal(cell.style.color, '#123456');
  assert.equal(cell.style.backgroundColor, '#FFEEAA');
  assert.equal(cell.style.fontSize, '14pt');
  assert.equal(cell.style.fontWeight, '700');
  assert.equal(cell.style.fontStyle, 'italic');
  assert.equal(cell.style.textAlign, 'center');
  assert.equal(cell.style.paddingTop, '0px');
  assert.equal(cell.style.paddingBottom, '0px');
  assert.equal(cell.style.justifyContent, 'flex-end');
  assert.equal(cell.style.borderLeft, '1px solid #ABCDEF');
});


test('XLSX range responses include parsed cell style properties', () => {
  const style = { bold: true, fill: '#FFEEAA' };
  const document = { kind: 'xlsx', getCell: () => ({ value: 'value', kind: 'string', styleProperties: style }) };
  const range = getRange(document, { rowStart: 0, columnStart: 0, rowCount: 1, columnCount: 1, requestId: 1 });
  assert.deepEqual(range.rows[0][0].style, style);
});

test('열 자동 맞춤은 CSV와 XLSX 전체 셀·표시값·신규 draft를 검사한다', () => {
  const csvBytes = Buffer.from('name,description\nA,Short\nB,Much longer description\n');
  const csv = new JinSheetDocument(uri('/tmp/fit.csv'), 'csv', load(csvBytes, '.csv'), csvBytes, true, '');
  assert.equal(estimateDocumentColumnWidth(csv, 1), 181);
  const xlsxBytes = minimalXlsx();
  const xlsx = new JinSheetDocument(uri('/tmp/fit.xlsx'), 'xlsx', analyze(xlsxBytes), xlsxBytes, true, '');
  assert.equal(estimateDocumentColumnWidth(xlsx, 0), 48);
  xlsx.applyCellEdit(1, 0, 'A much longer value');
  assert.equal(estimateDocumentColumnWidth(xlsx, 0), 153);
  assert.throws(() => estimateDocumentColumnWidth(csv, -1), /자동 맞춤/);
  csv.dispose(); xlsx.dispose();
});

test('열 자동 맞춤 버튼은 선택 열을 요청하고 유효한 응답만 적용·저장한다', () => {
  const { document, elements, messages, dispatch } = runGridWebview(renderHtml(), { kind: 'xlsx', rowCount: 2, columnCount: 2 });
  document.querySelector('[data-row="0"][data-column="1"]').onclick();
  elements.get('autoFitColumn').onclick();
  const request = messages.filter(message => message.type === 'auto-fit-column').at(-1);
  assert.deepEqual(JSON.parse(JSON.stringify(request)), { type: 'auto-fit-column', requestId: 1, column: 1, sheetIndex: 0 });
  const header = elements.get('canvas').children.find(item => item.classList.contains('colhead') && item.dataset.column === '1');
  dispatch({ type: 'auto-fit-width', requestId: 0, column: 1, width: 240, sheetIndex: 0 });
  assert.equal(header.style.width, '144px');
  dispatch({ type: 'auto-fit-width', requestId: request.requestId, column: 1, width: 240, sheetIndex: 0 });
  assert.equal(header.style.width, '240px');
  assert.deepEqual(JSON.parse(JSON.stringify(messages.filter(message => message.type === 'column-width').at(-1))), { type: 'column-width', column: 1, width: 240, sheetIndex: 0 });
});

test('워크시트 탭을 두 번 클릭해 이름을 편집하고 Enter로 확정한다', () => {
  const { elements, messages, dispatch } = runGridWebview(renderHtml(), { kind: 'xlsx', rowCount: 2, columnCount: 2 });
  dispatch({ type: 'model', kind: 'xlsx', readOnly: false, rowCount: 2, columnCount: 2, sheetIndex: 0, sheets: [{ name: 'Data', index: 0, visible: true }, { name: 'Hidden', index: 1, visible: false }] });
  const tab = elements.get('sheets').children[0];
  let prevented = false;
  tab.ondblclick({ preventDefault() { prevented = true; } });
  const input = tab.children[0];
  input.value = 'Quarter 2';
  input.listeners.get('keydown')({ key: 'Enter', preventDefault() {} });
  assert.equal(prevented, true);
  assert.deepEqual(JSON.parse(JSON.stringify(messages.filter(message => message.type === 'rename-sheet').at(-1))), { type: 'rename-sheet', sheetIndex: 0, name: 'Quarter 2' });
  elements.get('sheets').children.at(-2).onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(messages.filter(message => message.type === 'add-sheet').at(-1))), { type: 'add-sheet', sheetIndex: 0 });
  elements.get('sheets').children.at(-1).onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(messages.filter(message => message.type === 'delete-sheet').at(-1))), { type: 'delete-sheet', sheetIndex: 0 });
});

test('Ctrl+A는 현재 영역, 반복 입력은 시트를 선택하고 오래된 응답은 무시한다', () => {
  const { document, documentListeners, elements, messages, dispatch } = runGridWebview(renderHtml(), { rowCount: 4, columnCount: 4 });
  const keydown = documentListeners.get('keydown'), active = document.querySelector('[data-row="0"][data-column="0"]');
  keydown({ target: active, key: 'a', ctrlKey: true, metaKey: false, preventDefault() {} });
  let request = messages.filter(message => message.type === 'current-region').at(-1);
  assert.equal(request.includeSheet, false);
  dispatch({ type: 'current-region', requestId: request.requestId, row: 0, column: 0, sheetIndex: 0, range: { top: 0, bottom: 1, left: 0, right: 1 } });
  assert.equal(active.classList.contains('in-range'), true);
  keydown({ target: active, key: 'a', ctrlKey: true, metaKey: false, preventDefault() {} });
  request = messages.filter(message => message.type === 'current-region').at(-1);
  assert.equal(request.includeSheet, true);
  dispatch({ type: 'current-region', requestId: request.requestId - 1, row: 0, column: 0, sheetIndex: 0, range: { top: 0, bottom: 3, left: 0, right: 3 } });
  assert.equal(document.querySelector('[data-row="3"][data-column="3"]').classList.contains('in-range'), false);
  dispatch({ type: 'current-region', requestId: request.requestId, row: 0, column: 0, sheetIndex: 0, range: { top: 0, bottom: 3, left: 0, right: 3 } });
  assert.equal(document.querySelector('[data-row="3"][data-column="3"]').classList.contains('in-range'), true);
  assert.equal(elements.get('address').value, 'A1', 'Ctrl+A keeps the active cell');
});

test('자동 행 측정은 줄 바꿈·자동 높이 행을 포함하고 수동 높이는 제외한다', () => {
  const bytes = minimalXlsx(), document = new JinSheetDocument(uri('/tmp/auto-fit.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  const range = { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 };
  document.applyTextLayout(0, buildWrapTextEdits(document, range), [{ row: 0, height: 30 }]);
  let request = buildAutomaticRowFitRequest(document);
  assert.deepEqual(request.rows.map(item => item.row), [0]); assert.equal(request.rows[0].cells[0].wrapText, true);
  document.applyColumnWidth(0, 50, 0);
  request = buildAutomaticRowFitRequest(document); assert.ok(request.rows[0].cells[0].width <= 51);
  document.applyRowHeights(0, 0, 0, 60);
  assert.deepEqual(buildAutomaticRowFitRequest(document).rows, []);
  document.dispose();
});

test('파생 행 높이는 별도 undo 없이 표시·저장되고 편집 취소 시 무효화된다', async t => {
  const bytes = minimalXlsx(), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-auto-row-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'auto.xlsx'); await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  let count = 0, operation; document.onDidChange(change => { count++; operation = change; });
  document.applyCellEdit(0, 0, 'one\ntwo\nthree');
  const revision = document.renderRevision, pending = document.beginAutoRowFit(0, [0]);
  const saving = document.save();
  assert.equal(document.pendingAutoRowFits.size, 1);
  document.completeAutoRowFit(pending, { sheetIndex: 0, heights: [{ row: 0, height: 60 }] });
  assert.equal(document.renderRevision, revision); assert.equal(count, 1); assert.equal(viewModel(document).rowHeights[0], 80);
  await saving;
  assert.equal(analyze(await fs.readFile(file)).sheets[0].rowHeights[0], 80);
  operation.undo(); assert.equal(document.autoRowHeights.size, 0); assert.equal(document.getCell(0, 0).value, 'old');
  operation.redo(); assert.equal(document.getCell(0, 0).value, 'one\ntwo\nthree');
  document.dispose();
});

test('오래된·잘못된 자동 측정과 닫힌 편집기는 저장 대기를 남기지 않는다', async () => {
  const bytes = minimalXlsx(), document = new JinSheetDocument(uri('/tmp/auto-stale.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  const old = document.beginAutoRowFit(0, [0]); document.applyCellEdit(0, 0, 'new');
  assert.equal(document.completeAutoRowFit(old, { sheetIndex: 0, heights: [{ row: 0, height: 90 }] }), false);
  const invalid = document.beginAutoRowFit(0, [0]);
  assert.throws(() => document.completeAutoRowFit(invalid, { sheetIndex: 0, heights: [{ row: 1, height: 90 }] }), /결과/);
  assert.equal(document.pendingAutoRowFits.size, 0); assert.equal(document.autoRowHeights.size, 0);
  const closing = document.beginAutoRowFit(0, [0]); document.dispose(); await closing.promise;
  assert.equal(document.pendingAutoRowFits.size, 0);
});

test('다른 시트의 배경 행 측정은 셀을 선택하거나 현재 시트를 바꾸지 않는다', () => {
  const grid = runGridWebview(renderHtml(), { kind: 'xlsx', sheetIndex: 0 });
  const before = grid.elements.get('address').value;
  grid.dispatch({ type: 'row-fit-request', background: true, requestId: 10, sheetIndex: 1, defaultHeight: 26, rows: [{ row: 3, cells: [{ text: 'line', width: 50, wrapText: true }] }] });
  assert.equal(grid.messages.at(-1).sheetIndex, 1); assert.equal(grid.messages.at(-1).requestId, 10);
  assert.equal(grid.elements.get('address').value, before);
});

test('배경 높이 갱신은 입력 중인 셀과 커서 포커스를 보존한다', () => {
  const grid = runGridWebview(renderHtml(), { kind: 'xlsx', sheetIndex: 0 });
  const cell = grid.document.querySelector('[data-row="0"][data-column="0"]');
  cell.ondblclick({ target: cell, clientX: 80, clientY: 40 });
  const input = cell.children[0]; input.value = 'unfinished edit'; input.focus();
  grid.dispatch({ type: 'row-layout', sheetIndex: 0, rowHeights: { 0: 80 } });
  assert.equal(cell.style.height, '80px'); assert.equal(cell.children[0], input);
  assert.equal(input.value, 'unfinished edit'); assert.equal(grid.document.activeElement, input);
});

test('자동 높이의 측정 응답이 없으면 제한 시간 후 대기를 끝낸다', async t => {
  t.mock.timers.enable({ apis: ['setTimeout'] });
  const bytes = minimalXlsx(), document = new JinSheetDocument(uri('/tmp/auto-timeout.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  const pending = document.beginAutoRowFit(0, [0]);
  t.mock.timers.tick(2000); await document.waitForAutoRowFits();
  assert.equal(document.pendingAutoRowFits.size, 0);
  assert.equal(document.completeAutoRowFit(pending, { sheetIndex: 0, heights: [{ row: 0, height: 90 }] }), false);
  assert.equal(document.dirty, false); document.dispose();
});

test('다른 이름 저장과 백업도 대기 중인 자동 높이를 포함한다', async t => {
  const bytes = minimalXlsx(), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-auto-row-output-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  for (const mode of ['saveAs', 'backup']) {
    const document = new JinSheetDocument(uri(path.join(directory, 'source.xlsx')), 'xlsx', analyze(bytes), bytes, true, '');
    document.applyCellEdit(0, 0, mode);
    const pending = document.beginAutoRowFit(0, [0]), target = path.join(directory, mode + '.xlsx');
    const output = mode === 'saveAs' ? document.saveAs(uri(target)) : document.backup({ destination: uri(target) });
    document.completeAutoRowFit(pending, { sheetIndex: 0, heights: [{ row: 0, height: 45 }] });
    await output;
    const saved = analyze(await fs.readFile(target));
    assert.equal(saved.sheets[0].rowHeights[0], 60); assert.equal(saved.sheets[0].cells.get('A1').value, mode);
    document.dispose();
  }
});

test('페이지 설정은 셀 초안과 저장 후 undo·redo를 보존한다', async t => {
  const bytes = minimalXlsx(), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-page-setup-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'pages.xlsx'); await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  document.applyCellEdit(0, 0, 'unsaved'); let operation; document.onDidChange(change => { operation = change; });
  document.applyPrintSettings(0, { ...document.state.sheets[0].printSettings, orientation: 'landscape', gridLines: true });
  assert.equal(document.getCell(0, 0).value, 'unsaved'); assert.equal(viewModel(document).printSettings.orientation, 'landscape');
  await document.save(); assert.equal(analyze(await fs.readFile(file)).sheets[0].printSettings.gridLines, true);
  operation.undo(); assert.equal(viewModel(document).printSettings.orientation, 'portrait'); assert.equal(document.getCell(0, 0).value, 'unsaved');
  operation.redo(); await document.save(); assert.equal(analyze(await fs.readFile(file)).sheets[0].printSettings.orientation, 'landscape');
  document.dispose();
});

test('인쇄 출력은 미저장 수식·서식과 숨김 상태를 사용하며 원본을 수정하지 않는다', () => {
  const bytes = minimalXlsx('', '<row r="1"><c r="A1"><v>5</v></c><c r="B1"><f>A1*2</f><v>0</v></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>hidden text</t></is></c></row>');
  const document = new JinSheetDocument(uri('/tmp/print.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  document.applyCellEdit(0, 0, '7'); document.applyHiddenRows(0, [1], true);
  const html = buildPrintDocument(document);
  assert.match(html, />14</); assert.doesNotMatch(html, /hidden text/); assert.equal(document.getCell(0, 1).formula, 'A1*2');
  assert.throws(() => buildPrintDocument(document, 0, { startRow: 0, endRow: 2000, startColumn: 0, endColumn: 99 }), /인쇄/);
  const selected = buildPrintDocument(document, 0, { startRow: 0, endRow: 0, startColumn: 1, endColumn: 1 });
  assert.match(selected, />14</); assert.doesNotMatch(selected, />7</); document.dispose();
});

test('페이지 설정과 인쇄 버튼은 시트와 선택 범위를 구분하고 CSV 설정 편집을 막는다', () => {
  const grid = runGridWebview(renderHtml(), { kind: 'xlsx', sheetIndex: 0 });
  grid.elements.get('printOrientation').value = 'landscape'; grid.elements.get('printFit').checked = true; grid.elements.get('printFitWidth').value = '1'; grid.elements.get('printFitHeight').value = '0';
  grid.elements.get('applyPageSettings').onclick(); const config = grid.messages.at(-1);
  assert.equal(config.type, 'page-settings'); assert.equal(config.settings.orientation, 'landscape'); assert.equal(config.settings.fitToHeight, 0);
  grid.elements.get('printSheet').onclick(); assert.equal(grid.messages.at(-1).type, 'export-print'); assert.equal(grid.messages.at(-1).range, null);
  grid.elements.get('printSelection').onclick(); assert.equal(grid.messages.at(-1).range.startRow, 0);
  const csv = runGridWebview(renderHtml(), { kind: 'csv' }); assert.equal(csv.elements.get('applyPageSettings').disabled, true); assert.equal(csv.elements.get('printSheet').disabled, false);
});

test('CSV 인쇄는 선택 범위와 미저장 값을 반영하고 빈 시트는 안내한다', () => {
  const bytes = Buffer.from('name,value\nold,2\n'), document = new JinSheetDocument(uri('/tmp/print.csv'), 'csv', load(bytes, '.csv'), bytes, true, '');
  document.setInitialCsvState(document.state); document.applyCellEdit(1, 0, '<new>');
  const html = buildPrintDocument(document, 0, { startRow: 1, endRow: 1, startColumn: 0, endColumn: 1 }, 'ko');
  assert.match(html, /&lt;new&gt;/); assert.match(html, /인쇄 \/ PDF 저장/); assert.doesNotMatch(html, />name</);
  document.dispose();
  const empty = minimalXlsx('', ''), blank = new JinSheetDocument(uri('/tmp/empty.xlsx'), 'xlsx', analyze(empty), empty, true, '');
  assert.throws(() => buildPrintDocument(blank), /표시 셀/); blank.dispose();
});

test('페이지 맞춤 옵션은 일반 모델 갱신 후에도 배율 입력의 비활성 상태를 유지한다', () => {
  const grid = runGridWebview(renderHtml(), { kind: 'xlsx', sheetIndex: 0 });
  grid.elements.get('printFit').checked = true; grid.elements.get('printFit').onchange();
  assert.equal(grid.elements.get('printScale').disabled, true);
  grid.dispatch({ type: 'model', kind: 'xlsx', sheetIndex: 0, rowCount: 2, columnCount: 2, readOnly: false });
  assert.equal(grid.elements.get('printScale').disabled, true); assert.equal(grid.elements.get('printFitWidth').disabled, false);
});

test('인쇄 영역은 미저장 값과 저장 후 Undo·Redo를 보존하고 선택 인쇄로 대체할 수 있다', async t => {
  const bytes = minimalXlsx('', '<row r="1"><c r="A1" t="inlineStr"><is><t>include</t></is></c><c r="B1" t="inlineStr"><is><t>outside</t></is></c></row>');
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-print-area-')); t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'areas.xlsx'); await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  document.applyCellEdit(0, 0, 'draft include'); let operation; document.onDidChange(change => { operation = change; });
  document.applyPrintArea(0, { top: 0, bottom: 0, left: 0, right: 0 });
  assert.match(buildPrintDocument(document), />draft include</); assert.doesNotMatch(buildPrintDocument(document), />outside</);
  const selection = buildPrintDocument(document, 0, { startRow: 0, endRow: 0, startColumn: 1, endColumn: 1 });
  assert.match(selection, />outside</); assert.doesNotMatch(selection, />draft include</);
  await document.save(); assert.equal(analyze(await fs.readFile(file)).sheets[0].printAreas.length, 1);
  operation.undo(); assert.deepEqual(viewModel(document).printAreas, []); assert.equal(document.getCell(0, 0).value, 'draft include');
  operation.redo(); document.applyStructureChange({ sheetIndex: 0, axis: 'row', action: 'insert', index: 0 });
  assert.equal(viewModel(document).printAreas[0].top, 1); assert.match(buildPrintDocument(document), />draft include</);
  await document.save(); document.dispose();
});

test('여러 인쇄 영역은 각각 출력하며 전부 숨겨진 영역은 건너뛴다', () => {
  const bytes = minimalXlsx('', '<row r="1"><c r="A1" t="inlineStr"><is><t>first area</t></is></c><c r="B1" t="inlineStr"><is><t>not printed</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>second area</t></is></c></row>');
  const document = new JinSheetDocument(uri('/tmp/multiple-areas.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  document.state.sheets[0].printAreas = [{ sheetName: 'Data', top: 0, bottom: 0, left: 0, right: 0 }, { sheetName: 'Data', top: 1, bottom: 1, left: 0, right: 0 }];
  const html = buildPrintDocument(document);
  assert.equal((html.match(/<section class="page"/g) || []).length, 2); assert.match(html, />first area</); assert.match(html, />second area</); assert.doesNotMatch(html, />not printed</);
  document.state.sheets[0].hiddenRows = [1];
  const hidden = buildPrintDocument(document); assert.match(hidden, />first area</); assert.doesNotMatch(hidden, />second area</);
  document.dispose();
});

test('전체 열 인쇄 영역은 사용 행까지만 내보내고 선택 범위와 지우기 버튼을 구분한다', () => {
  const bytes = minimalXlsx(), document = new JinSheetDocument(uri('/tmp/full-column-area.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  document.applyPrintArea(0, { top: 0, bottom: 1048575, left: 0, right: 0 });
  assert.match(buildPrintDocument(document), />old</); document.dispose();
  const grid = runGridWebview(renderHtml(), { kind: 'xlsx', sheetIndex: 0, printAreas: [{ top: 1, bottom: 3, left: 0, right: 1 }] });
  assert.equal(grid.elements.get('printAreaStatus').textContent, 'Print area: A2:B4'); assert.equal(grid.elements.get('clearPrintArea').disabled, false);
  grid.elements.get('setPrintArea').onclick(); assert.equal(grid.messages.at(-1).type, 'print-area'); assert.equal(grid.messages.at(-1).action, 'set');
  grid.elements.get('clearPrintArea').onclick(); assert.equal(grid.messages.at(-1).action, 'clear'); assert.equal(grid.messages.at(-1).range, null);
  const csv = runGridWebview(renderHtml(), { kind: 'csv' }); assert.equal(csv.elements.get('setPrintArea').disabled, true);
});

test('해석하지 못한 인쇄 영역을 시트 전체로 대체하지 않고 명시적 선택 인쇄를 허용한다', () => {
  const bytes = minimalXlsx(), files = unzipSync(bytes);
  files['xl/workbook.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/workbook.xml']).replace('</sheets>', '</sheets><definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">OFFSET(Data!A1,0,0,2,2)</definedName></definedNames>'));
  const packed = zipSync(files), state = analyze(packed);
  assert.equal(state.sheets[0].printAreaUnsupported, true);
  const document = new JinSheetDocument(uri('/tmp/unsupported-area.xlsx'), 'xlsx', state, packed, false, state.readOnlyReason);
  assert.throws(() => buildPrintDocument(document), /직접 선택/);
  const explicit = buildPrintDocument(document, 0, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }); assert.match(explicit, />old</);
  document.dispose();
});

test('반복 제목 입력은 행 번호·열 문자와 빈칸을 정규화하고 범위를 검증한다', () => {
  assert.deepEqual(parsePrintTitleInputs(' $1:$3 ', 'a:b'), { rows: { start: 0, end: 2 }, columns: { start: 0, end: 1 } });
  assert.deepEqual(parsePrintTitleInputs('', 'C'), { rows: null, columns: { start: 2, end: 2 } });
  for (const [rows, columns] of [['0:1', 'A'], ['3:1', 'A'], ['1:51', 'A'], ['', 'A:K'], ['', 'XFE'], ['A1:B2', ''], ['', 'B:A']]) assert.throws(() => parsePrintTitleInputs(rows, columns));
});

test('인쇄 영역 밖 제목을 가져와 저장·Undo·Redo하고 본문과 중복하지 않는다', async t => {
  const bytes = minimalXlsx('', '<row r="1"><c r="A1" t="inlineStr"><is><t>Corner title</t></is></c><c r="B1" t="inlineStr"><is><t>Column title</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Row title</t></is></c><c r="B2"><v>12</v></c></row>');
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-print-titles-')); t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'titles.xlsx'); await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  document.applyPrintArea(0, { top: 1, bottom: 1, left: 1, right: 1 }); let operation; document.onDidChange(change => { operation = change; });
  document.applyPrintTitles(0, { rows: { start: 0, end: 0 }, columns: { start: 0, end: 0 } });
  const html = buildPrintDocument(document); assert.match(html, />Corner title</); assert.match(html, />Column title</); assert.match(html, />Row title</); assert.match(html, />12</);
  await document.save(); assert.ok(analyze(await fs.readFile(file)).sheets[0].printTitles.rows);
  operation.undo(); assert.deepEqual(viewModel(document).printTitles, { rows: null, columns: null }); assert.doesNotMatch(buildPrintDocument(document), />Corner title</);
  operation.redo(); await document.save(); assert.match(buildPrintDocument(document), />Corner title</);
  document.dispose();
});

test('숨겨진 제목과 숨겨진 본문을 제외하고 해석 불가 제목의 인쇄를 막는다', () => {
  const bytes = minimalXlsx('', '<row r="1"><c r="A1" t="inlineStr"><is><t>Hidden title</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Body</t></is></c></row>');
  const document = new JinSheetDocument(uri('/tmp/hidden-titles.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  document.applyPrintTitles(0, { rows: { start: 0, end: 0 }, columns: null }); document.applyHiddenRows(0, [0], true);
  const html = buildPrintDocument(document); assert.match(html, />Body</); assert.doesNotMatch(html, />Hidden title</);
  document.state.sheets[0].hiddenRows = [1]; document.state.sheets[0].printAreas = [{ top: 1, bottom: 1, left: 0, right: 0 }];
  assert.throws(() => buildPrintDocument(document), /표시/);
  document.state.sheets[0].printTitlesUnsupported = true; assert.throws(() => buildPrintDocument(document, 0, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }), /반복 제목/);
  document.dispose();
});

test('반복 제목 도구는 적용·지우기 메시지를 보내고 모델 갱신에서 미적용 입력을 보존한다', () => {
  const grid = runGridWebview(renderHtml(), { kind: 'xlsx', printTitles: { rows: { start: 0, end: 1 }, columns: { start: 0, end: 0 } } });
  assert.equal(grid.elements.get('printTitleRows').value, '1:2'); assert.equal(grid.elements.get('printTitleColumns').value, 'A:A');
  grid.elements.get('printTitleRows').value = '2:3'; grid.elements.get('applyPrintTitles').onclick(); assert.equal(grid.messages.at(-1).type, 'print-titles'); assert.equal(grid.messages.at(-1).rows, '2:3');
  grid.dispatch({ type: 'model', kind: 'xlsx', sheetIndex: 0, rowCount: 2, columnCount: 2, readOnly: false, printTitles: { rows: { start: 0, end: 1 }, columns: { start: 0, end: 0 } } });
  assert.equal(grid.elements.get('printTitleRows').value, '2:3');
  grid.elements.get('clearPrintTitles').onclick(); assert.equal(grid.messages.at(-1).action, 'clear'); assert.equal(grid.elements.get('printTitleRows').value, '');
  const csv = runGridWebview(renderHtml(), { kind: 'csv' }); assert.equal(csv.elements.get('applyPrintTitles').disabled, true);
});

test('인쇄 영역 밖 병합 제목은 전체 범위를 요구하고 완전한 제목을 출력한다', () => {
  const original = minimalXlsx('', '<row r="1"><c r="A1" t="inlineStr"><is><t>Merged header</t></is></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>Body</t></is></c></row>'), files = unzipSync(original);
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</worksheet>', '<mergeCells count="1"><mergeCell ref="A1:B2"/></mergeCells></worksheet>'));
  const bytes = zipSync(files), document = new JinSheetDocument(uri('/tmp/merged-print-titles.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  document.applyPrintArea(0, { top: 2, bottom: 2, left: 0, right: 1 });
  document.applyPrintTitles(0, { rows: { start: 0, end: 0 }, columns: null }); assert.throws(() => buildPrintDocument(document), /병합/);
  document.applyPrintTitles(0, { rows: { start: 0, end: 1 }, columns: null }); const html = buildPrintDocument(document);
  assert.match(html, /rowspan="2" colspan="2"/); assert.match(html, />Merged header</); assert.match(html, />Body</);
  document.dispose();
});

test('머리글 저장·Undo·Redo와 내보내기는 셀 초안·파일명·시트명을 보존한다', async t => {
  const bytes = minimalXlsx(), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-header-footer-')); t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'header.xlsx'); await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, ''); document.applyCellEdit(0, 0, 'draft');
  const config = JSON.parse(JSON.stringify(document.state.sheets[0].headerFooter)); config.odd.header.left = '&F'; config.odd.header.right = '&A'; config.odd.footer.center = '&P / &N';
  let operation; document.onDidChange(change => { operation = change; }); document.applyHeaderFooter(0, config); await document.save();
  assert.equal(analyze(await fs.readFile(file)).sheets[0].headerFooter.odd.header.left, '&F');
  const html = buildPrintDocument(document); assert.match(html, />header.xlsx</); assert.match(html, />Data</); assert.match(html, />draft</);
  operation.undo(); assert.equal(viewModel(document).headerFooter.odd.header.left, ''); assert.equal(document.getCell(0, 0).value, 'draft');
  operation.redo(); await document.save(); assert.match(buildPrintDocument(document), />header.xlsx</); document.dispose();
});

test('머리글 입력은 쪽 구분별 초안을 보존하고 필드 삽입과 모두 지우기를 제공한다', () => {
  const grid = runGridWebview(renderHtml(), { kind: 'xlsx' });
  grid.elements.get('hf-header-center').value = 'Default '; grid.elements.get('hf-header-center').oninput();
  grid.elements.get('hfToken').value = '&P'; grid.elements.get('hfInsert').onclick();
  assert.equal(grid.elements.get('hf-header-center').value, 'Default &P');
  grid.elements.get('hfVariant').value = 'first'; grid.elements.get('hfVariant').onchange(); grid.elements.get('hf-header-center').value = 'First'; grid.elements.get('hf-header-center').oninput();
  grid.elements.get('hfVariant').value = 'odd'; grid.elements.get('hfVariant').onchange(); assert.equal(grid.elements.get('hf-header-center').value, 'Default &P');
  grid.elements.get('hfApply').onclick(); const config = grid.messages.at(-1); assert.equal(config.type, 'header-footer'); assert.equal(config.settings.first.header.center, 'First');
  grid.elements.get('hfClear').onclick(); assert.equal(grid.messages.at(-1).action, 'clear'); assert.equal(grid.elements.get('hf-header-center').value, '');
  const csv = runGridWebview(renderHtml(), { kind: 'csv' }); assert.equal(csv.elements.get('hfApply').disabled, true);
});

test('시작 쪽 번호 지정과 해석 불가 머리글의 편집·지우기 상태를 구분한다', () => {
  const grid = runGridWebview(renderHtml(), { kind: 'xlsx' });
  assert.equal(grid.elements.get('printFirstPage').disabled, true);
  grid.elements.get('printUseFirstPage').checked = true; grid.elements.get('printUseFirstPage').onchange(); assert.equal(grid.elements.get('printFirstPage').disabled, false);
  grid.elements.get('printFirstPage').value = '0'; grid.elements.get('applyPageSettings').onclick();
  assert.equal(grid.messages.at(-1).settings.firstPageNumber, 0); assert.equal(grid.messages.at(-1).settings.useFirstPageNumber, true);
  grid.dispatch({ type: 'model', kind: 'xlsx', sheetIndex: 0, rowCount: 2, columnCount: 2, readOnly: false, headerFooterUnsupported: true });
  assert.equal(grid.elements.get('hfApply').disabled, true); assert.equal(grid.elements.get('hfClear').disabled, false); assert.match(grid.elements.get('hfStatus').textContent, /Clear/);
});

test('저장된 여러 인쇄 영역은 시작 번호 0과 전체 쪽 수를 머리글·바닥글에 연속 적용한다', () => {
  const original = minimalXlsx(), files = unzipSync(original);
  files['xl/workbook.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/workbook.xml']).replace('</sheets>', '</sheets><definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$A$1,Data!$B$1:$B$1</definedName></definedNames>'));
  const bytes = zipSync(files), document = new JinSheetDocument(uri('/tmp/combined-headers.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  document.applyPrintSettings(0, { ...document.state.sheets[0].printSettings, useFirstPageNumber: true, firstPageNumber: 0 });
  const config = JSON.parse(JSON.stringify(document.state.sheets[0].headerFooter)); config.odd.footer.center = '&P / &N'; document.applyHeaderFooter(0, config);
  const html = buildPrintDocument(document), pages = html.match(/<section class="page"[\s\S]*?<\/section>/g);
  assert.equal(pages.length, 2);
  const text = page => page.match(/<!--print-hf-start-->([\s\S]*?)<!--print-hf-end-->/)[1].replace(/<[^>]*>/g, '');
  assert.equal(text(pages[0]), '0 / 2'); assert.equal(text(pages[1]), '1 / 2');
  document.dispose();
});

test('페이지 나누기는 초안·저장 후 Undo·Redo와 인쇄 분할을 보존한다', async t => {
  const bytes = minimalXlsx('', '<row r="1"><c r="A1" t="inlineStr"><is><t>one</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>two</t></is></c></row>');
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-breaks-')); t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'breaks.xlsx'); await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, ''); document.applyCellEdit(0, 0, 'draft');
  let operation; document.onDidChange(change => { operation = change; }); document.applyPageBreaks(0, { axis: 'row', action: 'add', index: 1 });
  assert.equal((buildPrintDocument(document).match(/<section class="page"/g) || []).length, 2);
  await document.save(); assert.deepEqual(analyze(await fs.readFile(file)).sheets[0].pageBreaks.rows, [1]);
  operation.undo(); assert.deepEqual(viewModel(document).pageBreaks.rows, []); assert.equal(document.getCell(0, 0).value, 'draft');
  operation.redo(); document.applyStructureChange({ axis: 'row', action: 'insert', index: 0, sheetIndex: 0 }); assert.deepEqual(viewModel(document).pageBreaks.rows, [2]);
  await document.save(); document.dispose();
});

test('행·열 머리글의 나누기 표시는 서식 변경 없이 도구 상태와 연결된다', () => {
  const grid = runGridWebview(renderHtml(), { kind: 'xlsx', pageBreaks: { rows: [1], columns: [1], hasAny: true } });
  assert.equal(grid.elements.get('breakAddRow').disabled, true);
  const rowHead = grid.elements.get('canvas').children.find(item => item.classList.contains('rowhead') && item.dataset.viewRow === '1');
  const columnHead = grid.elements.get('canvas').children.find(item => item.classList.contains('colhead') && item.dataset.column === '1');
  assert.equal(rowHead.classList.contains('manual-row-break'), true); assert.equal(columnHead.classList.contains('manual-column-break'), true);
  const cell = grid.document.querySelector('[data-row="1"][data-column="1"]'); cell.onclick({ shiftKey: false });
  assert.equal(grid.elements.get('breakRemoveRow').disabled, false); grid.elements.get('breakRemoveRow').onclick();
  assert.equal(grid.messages.at(-1).type, 'page-break'); assert.equal(grid.messages.at(-1).index, 1); assert.equal(grid.messages.at(-1).action, 'remove');
  grid.elements.get('breakReset').onclick(); assert.equal(grid.messages.at(-1).action, 'reset');
  assert.match(grid.elements.get('pageBreakStatus').textContent, /Before rows: 2/);
  grid.elements.get('printPageOrder').value = 'overThenDown'; grid.elements.get('applyPageSettings').onclick(); assert.equal(grid.messages.at(-1).settings.pageOrder, 'overThenDown');
});

test('지원하지 않는 나누기는 기본 인쇄를 막고 페이지 수 맞춤은 보존한 채 무시한다', () => {
  const original = minimalXlsx(), files = unzipSync(original);
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</worksheet>', '<rowBreaks count="1" manualBreakCount="1"><brk id="2" min="0" max="0" man="1"/></rowBreaks></worksheet>'));
  const bytes = zipSync(files), document = new JinSheetDocument(uri('/tmp/unsupported-break.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  assert.throws(() => buildPrintDocument(document), /지운 뒤/);
  document.applyPrintSettings(0, { ...document.state.sheets[0].printSettings, fitToPage: true, fitToWidth: 1, fitToHeight: 1 });
  assert.match(buildPrintDocument(document), />old</); assert.equal(document.state.sheets[0].pageBreaksUnsupported, true);
  document.dispose();
});

test('조건 필터 저장·재개방·해제는 native 행 표시와 수동 숨김을 보존하고 Undo한다', async t => {
  const rows = '<row r="1"><c r="A1"><v>0</v></c></row><row r="2" hidden="1"><c r="A2"><v>5</v></c></row><row r="3"><c r="A3"><v>15</v></c></row><row r="4"><c r="A4"><v>25</v></c></row>';
  const bytes = minimalXlsx('', rows), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-native-filter-')); t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'visibility.xlsx'); await fs.writeFile(file, bytes);
  let document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  const filter = { top: 0, bottom: 3, left: 0, right: 0, columns: [{ column: 0, custom: { and: false, conditions: [{ operator: 'greaterThan', value: '10' }] } }] };
  document.applyAutoFilter(0, filter); document.applyCellEdit(3, 0, '2'); await document.save();
  let saved = analyze(await fs.readFile(file)); assert.deepEqual(saved.sheets[0].nativeHiddenRows, [1, 3]); assert.deepEqual(saved.sheets[0].hiddenRows, [1]); document.dispose();
  document = new JinSheetDocument(uri(file), 'xlsx', saved, Buffer.from(saved.bytes), true, ''); let operation; document.onDidChange(change => { operation = change; });
  document.applyAutoFilter(0, null); const clear = operation; assert.deepEqual(viewModel(document).hiddenRows, [1]); await document.save();
  assert.deepEqual(analyze(await fs.readFile(file)).sheets[0].nativeHiddenRows, [1]);
  clear.undo(); assert.deepEqual(viewModel(document).hiddenRows, [1, 3]); await document.save(); assert.deepEqual(analyze(await fs.readFile(file)).sheets[0].nativeHiddenRows, [1, 3]);
  clear.redo(); assert.deepEqual(viewModel(document).hiddenRows, [1]); document.dispose();
});

function totalsTableXlsx({ legacy = false, noFilter = false, noHeader = false } = {}) {
  const files = unzipSync(tableXlsx());
  let table = new TextDecoder().decode(files['xl/tables/table1.xml']).replace('ref="A1:B3"', 'ref="A1:B4" totalsRowCount="1"' + (noHeader ? ' headerRowCount="0"' : ''));
  if (legacy) table = table.replace('<autoFilter ref="A1:B3"', '<autoFilter ref="A1:B4"');
  if (noFilter) table = table.replace(/<autoFilter\b[^>]*\/>/, '');
  files['xl/tables/table1.xml'] = new TextEncoder().encode(table);
  let sheet = new TextDecoder().decode(files['xl/worksheets/sheet1.xml']);
  sheet = sheet.replace(/<row r="4">[\s\S]*?<\/row>/, '<row r="4"><c r="A4" t="inlineStr"><is><t>Total</t></is></c><c r="B4"><f>SUBTOTAL(109,[Amount])</f><v>30</v></c><c r="C4"><f>SUM(Sales[[#Totals],[Amount]])</f><v>30</v></c><c r="D4"><f>SUM(Sales[Amount])</f><v>30</v></c></row>');
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(sheet);
  return zipSync(files);
}

test('native 총계 표는 총계를 필터 옵션·숨김에서 제외하고 암시적 구조 참조로 합계를 계산한다', async t => {
  const bytes = totalsTableXlsx(), model = analyze(bytes); assert.equal(model.writable, true, model.readOnlyReason);
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-table-totals-')); t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'totals.xlsx'); await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', model, bytes, true, ''); let operation; document.onDidChange(change => { operation = change; });
  const filter = autoFilterRange(document, 0, 3, 1); assert.equal(filter.bottom, 2, 'selecting the total row still targets the data filter');
  assert.deepEqual(autoFilterOptions(document, 0, 0, filter).values, ['Alpha', 'Beta']);
  document.applyAutoFilter(0, updateAutoFilterCustomColumn(document, 0, 1, { and: false, conditions: [{ operator: 'greaterThan', value: '15' }] }, filter));
  const applied = operation; assert.deepEqual(viewModel(document).hiddenRows, [1]);
  const range = () => getRange(document, { requestId: 7, rowStart: 3, columnStart: 1, rowCount: 1, columnCount: 3 }).rows[0].map(cell => cell.cachedValue);
  assert.deepEqual(range(), [20, 20, 30]); await document.save();
  let saved = analyze(await fs.readFile(file)); assert.deepEqual(saved.sheets[0].nativeHiddenRows, [1]); assert.equal(saved.sheets[0].tables[0].autoFilter.bottom, 2); assert.equal(saved.sheets[0].autoFilter, null);
  applied.undo(); assert.deepEqual(range(), [30, 30, 30]); applied.redo(); assert.deepEqual(range(), [20, 20, 30]);
  document.applyAutoFilter(0, updateAutoFilterColumn(document, 0, 1, [], false, false, autoFilterRange(document, 0, 3, 1)));
  assert.deepEqual(viewModel(document).hiddenRows, [1, 2]); assert.deepEqual(range(), [0, 0, 30]); await document.save();
  saved = analyze(await fs.readFile(file)); assert.deepEqual(saved.sheets[0].nativeHiddenRows, [1, 2]); assert.doesNotMatch(saved.sheets[0].xml, /<row r="4"[^>]*hidden="1"/); document.dispose();
  const reopened = new JinSheetDocument(uri(file), 'xlsx', saved, Buffer.from(saved.bytes), true, '');
  reopened.applyAutoFilter(0, updateAutoFilterColumn(reopened, 0, 1, [], false, true, autoFilterRange(reopened, 0, 3, 1)));
  assert.equal(viewModel(reopened).hiddenRows, undefined); assert.equal(getRange(reopened, { requestId: 8, rowStart: 3, columnStart: 1, rowCount: 1, columnCount: 1 }).rows[0][0].cachedValue, 30);
  await reopened.save(); assert.deepEqual(analyze(await fs.readFile(file)).sheets[0].nativeHiddenRows, []); reopened.dispose();

});

test('총계 표에 필터를 새로 만들 때 데이터 범위를 사용하고 머리글 없는 표는 보호한다', () => {
  for (const noHeader of [false, true]) {
    const bytes = totalsTableXlsx({ noFilter: true, noHeader }), model = analyze(bytes); assert.equal(model.writable, true, model.readOnlyReason);
    const document = new JinSheetDocument(uri('/tmp/totals-new-filter.xlsx'), 'xlsx', model, bytes, true, '');
    if (noHeader) assert.throws(() => autoFilterRange(document, 0, 1, 1), /표시된 머리글/);
    else {
      const filter = autoFilterRange(document, 0, 3, 1); assert.equal(filter.bottom, 2);
      document.applyAutoFilter(0, updateAutoFilterColumn(document, 0, 1, ['10'], false, false, filter));
      assert.equal(document.state.sheets[0].autoFilter, null); assert.equal(document.state.sheets[0].tables[0].autoFilter.bottom, 2); assert.deepEqual(viewModel(document).hiddenRows, [2]);
    }
    document.dispose();
  }
});

test('이전 총계 포함 필터는 데이터 범위로 읽고 저장 시 표 XML을 정규화한다', () => {
  const bytes = totalsTableXlsx({ legacy: true }), model = analyze(bytes); assert.equal(model.writable, true, model.readOnlyReason);
  assert.equal(model.sheets[0].tables[0].autoFilter.bottom, 2);
  const document = new JinSheetDocument(uri('/tmp/legacy-totals-filter.xlsx'), 'xlsx', model, bytes, true, '');
  const saved = analyze(document.createOutput()); assert.equal(saved.sheets[0].tables[0].filterRangeNeedsNormalization, false);
  assert.match(new TextDecoder().decode(saved.files['xl/tables/table1.xml']), /<autoFilter ref="A1:B3"/);
  assert.match(new TextDecoder().decode(saved.files['xl/tables/table1.xml']), /<table[^>]*ref="A1:B4"/); document.dispose();
});

test('총계 표의 크기 변경 도구는 총계를 포함한 선택의 행·열 범위를 요청한다', () => {
  const table = { name: 'Sales', range: { top: 0, bottom: 3, left: 0, right: 1 }, headerRowCount: 1, totalsRowCount: 1 };
  const { document, documentListeners, elements, messages } = runGridWebview(renderHtml(), { kind: 'xlsx', rowCount: 6, columnCount: 3, tables: [table] });
  const select = (row, column) => {
    const first = document.querySelector('[data-row="0"][data-column="0"]'), last = document.querySelector('[data-row="' + row + '"][data-column="' + column + '"]');
    first.onmousedown({ target: first, preventDefault() {}, clientX: 10, clientY: 40, shiftKey: false }); last.onmouseenter(); documentListeners.get('mouseup')();
  };
  select(5, 1); assert.equal(elements.get('resizeTable').disabled, false); documentListeners.get('click')({ target: elements.get('resizeTable') }); assert.equal(messages.at(-1).range.bottom, 5);
  select(3, 2); assert.equal(elements.get('resizeTable').disabled, false); documentListeners.get('click')({ target: elements.get('resizeTable') }); assert.equal(messages.at(-1).range.right, 2);
});

test('총계 표 높이 조정은 미저장 값·필터·저장 전후 Undo와 재개방을 유지한다', async t => {
  const bytes = totalsTableXlsx(), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-totals-resize-')); t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'resize.xlsx'); await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, ''); let operation; document.onDidChange(change => { operation = change; });
  document.applyCellEdit(4, 1, '7'); document.applyTableResize(0, 'Sales', { top: 0, bottom: 5, left: 0, right: 1 }); const resize = operation;
  const total = (row, owner = document) => getRange(owner, { requestId: 9, rowStart: row, columnStart: 1, rowCount: 1, columnCount: 1 }).rows[0][0].cachedValue;
  assert.equal(total(5), 37); assert.equal(document.state.sheets[0].cells.get('A6').value, 'Total'); await document.save();
  const grownBytes = await fs.readFile(file); const beforeConflict = document.state;
  assert.throws(() => document.applyTableResize(0, 'Sales', { top: 0, bottom: 2, left: 0, right: 1 }), /이동할 합계 행 위치에 데이터/);
  assert.equal(document.state, beforeConflict); assert.equal(document.dirty, false); assert.deepEqual(await fs.readFile(file), grownBytes);
  resize.undo(); assert.equal(total(3), 30); assert.equal(document.getCell(4, 1).value, '7'); await document.save(); assert.equal(analyze(await fs.readFile(file)).sheets[0].tables[0].range.bottom, 3);
  resize.redo(); assert.equal(total(5), 37); await document.save(); document.dispose();
  const saved = analyze(await fs.readFile(file)), reopened = new JinSheetDocument(uri(file), 'xlsx', saved, Buffer.from(saved.bytes), true, '');
  const filter = autoFilterRange(reopened, 0, 5, 1); reopened.applyAutoFilter(0, updateAutoFilterCustomColumn(reopened, 0, 1, { and: false, conditions: [{ operator: 'greaterThan', value: '10' }] }, filter));
  assert.equal(total(5, reopened), 20); await reopened.save(); assert.deepEqual(analyze(await fs.readFile(file)).sheets[0].nativeHiddenRows, [1, 3, 4]);
  reopened.applyTableResize(0, 'Sales', { top: 0, bottom: 3, left: 0, right: 1 }); assert.equal(total(3, reopened), 20); await reopened.save();
  const shrunk = analyze(await fs.readFile(file)); assert.deepEqual(shrunk.sheets[0].nativeHiddenRows, [1]); assert.equal(shrunk.sheets[0].cells.get('B5').value, '7'); assert.equal(shrunk.sheets[0].tables[0].autoFilter.bottom, 2); reopened.dispose();
});

test('총계 제어는 저장·재개방·집계 변경·필터·Undo와 사용자 수식을 보존한다', async t => {
  const bytes = totalsTableXlsx(), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-totals-controls-')); t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'controls.xlsx'); await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, ''); let operation, events = 0; document.onDidChange(change => { operation = change; events++; });
  const total = () => getRange(document, { requestId: 10, rowStart: 3, columnStart: 1, rowCount: 1, columnCount: 1 }).rows[0][0].cachedValue;
  assert.equal(viewModel(document).tables[0].totalColumns[1].function, 'custom');
  document.applyTableTotals(0, 'Sales', { column: 1, function: 'average' }); const average = operation; assert.equal(total(), 15); await document.save();
  const count = events; document.applyTableTotals(0, 'Sales', { column: 1, function: 'average' }); assert.equal(events, count); assert.equal(document.dirty, false);
  average.undo(); assert.equal(total(), 30); await document.save(); average.redo(); assert.equal(total(), 15); await document.save();
  const filter = autoFilterRange(document, 0, 3, 1); document.applyAutoFilter(0, updateAutoFilterColumn(document, 0, 1, ['20'], false, false, filter)); assert.equal(total(), 20);
  document.applyTableTotals(0, 'Sales', { enabled: false }); assert.equal(document.state.sheets[0].tables[0].totalsRowCount, 0); assert.equal(document.getCell(3, 1).value, ''); await document.save(); document.dispose();
  const saved = analyze(await fs.readFile(file)), reopened = new JinSheetDocument(uri(file), 'xlsx', saved, Buffer.from(saved.bytes), true, '');
  reopened.applyTableTotals(0, 'Sales', { enabled: true }); assert.equal(reopened.state.sheets[0].tables[0].totalColumns[1].function, 'average'); assert.equal(getRange(reopened, { requestId: 11, rowStart: 3, columnStart: 1, rowCount: 1, columnCount: 1 }).rows[0][0].cachedValue, 20);
  await reopened.save(); assert.deepEqual(analyze(await fs.readFile(file)).sheets[0].nativeHiddenRows, [1]); reopened.dispose();
});

test('총계 도구는 요청 완료까지 잠그고 오류 입력·재시도와 오래된 응답을 처리한다', () => {
  const table = { name: 'Sales', range: { top: 0, bottom: 3, left: 0, right: 1 }, headerRowCount: 1, totalsRowCount: 1, columns: ['Name', 'Amount'], totalColumns: [{ function: 'none' }, { function: 'sum' }] };
  const { document, elements, messages, dispatch } = runGridWebview(renderHtml(), { kind: 'xlsx', rowCount: 4, columnCount: 2, tables: [table] });
  const cell = document.querySelector('[data-row="3"][data-column="1"]'); cell.onclick({ target: cell });
  const input = elements.get('tableTotalsFunction'), apply = elements.get('applyTableTotalFunction'), toggle = elements.get('tableTotals');
  assert.equal(toggle.attributes['aria-pressed'], 'true'); assert.equal(input.value, 'sum'); assert.equal(elements.get('tableTotalsColumn').textContent, 'Totals column: Amount');
  input.value = 'average'; input.onchange(); apply.onclick(); const request = messages.at(-1); assert.equal(request.function, 'average'); assert.equal(toggle.disabled, true); assert.equal(input.value, 'average');
  dispatch({ type: 'table-totals-error', requestId: request.requestId + 1, sheetIndex: 0, tableName: 'Sales', message: 'stale' }); assert.equal(toggle.disabled, true);
  dispatch({ type: 'table-totals-error', requestId: request.requestId, sheetIndex: 0, tableName: 'Sales', message: 'retry' }); assert.equal(toggle.disabled, false); assert.equal(input.value, 'average'); assert.equal(document.activeElement, input);
  apply.onclick(); assert.notEqual(messages.at(-1).requestId, request.requestId); const retry = messages.at(-1); dispatch({ type: 'table-totals-applied', requestId: request.requestId, sheetIndex: 0, tableName: 'Sales', row: 3, column: 1 }); assert.equal(toggle.disabled, true);
  dispatch({ type: 'table-totals-error', requestId: retry.requestId, sheetIndex: 0, tableName: 'Sales', message: 'retry' }); toggle.onclick(); assert.equal(messages.at(-1).enabled, false);
});

function calculatedTableXlsx({ native = false, exceptions = {} } = {}) {
  const files = unzipSync(tableXlsx());
  const rows = '<row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c><c r="C1" t="inlineStr"><is><t>Result</t></is></c></row>' + [10, 20, 30].map((value, index) => {
    const row = index + 2, exception = exceptions[row]; let result = native ? '<c r="C' + row + '"><f>B' + row + '*2</f><v>' + value * 2 + '</v></c>' : '';
    if (exception === '') result = '<c r="C' + row + '"/>';
    else if (typeof exception === 'number') result = '<c r="C' + row + '"><v>' + exception + '</v></c>';
    else if (exception?.formula) result = '<c r="C' + row + '"><f>' + exception.formula + '</f><v>0</v></c>';
    return '<row r="' + row + '"><c r="A' + row + '" t="inlineStr"><is><t>item' + index + '</t></is></c><c r="B' + row + '"><v>' + value + '</v></c>' + result + '</row>';
  }).join('') + '<row r="6"><c r="D6"><f>SUBTOTAL(9,Sales[Result])</f><v>0</v></c></row>';
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace(/<sheetData>[\s\S]*?<\/sheetData>/, '<sheetData>' + rows + '</sheetData>'));
  files['xl/tables/table1.xml'] = new TextEncoder().encode('<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:C4"><autoFilter ref="A1:C4"/><tableColumns count="3"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Amount"/><tableColumn id="3" name="Result">' + (native ? '<calculatedColumnFormula>B2*2</calculatedColumnFormula>' : '') + '</tableColumn></tableColumns></table>');
  return zipSync(files);
}

test('수식 입력은 표 계산 열을 상대 참조로 채우고 하나의 Undo·저장·재개방을 지원한다', async t => {
  const bytes = calculatedTableXlsx(), directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-calculated-column-')); t.after(() => fs.rm(directory, { recursive: true, force: true })); const file = path.join(directory, 'calculated.xlsx'); await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, ''); let operation; document.onDidChange(change => { operation = change; });
  document.applyCellEdit(3, 2, '=B4*2'); const applied = operation;
  assert.deepEqual([1,2,3].map(row=>document.getCell(row,2).formula), ['B2*2','B3*2','B4*2']); assert.equal(document.state.sheets[0].tables[0].calculatedColumns[2], 'B2*2');
  assert.deepEqual(getRange(document,{requestId:12,rowStart:1,columnStart:2,rowCount:3,columnCount:1}).rows.map(row=>row[0].cachedValue), [20,40,60]); await document.save();
  applied.undo(); assert.equal(document.getCell(1,2).hasFormula, false); assert.equal(document.getCell(1,2).value, ''); assert.equal(document.state.sheets[0].cells.has('C2'), false); assert.equal(document.state.sheets[0].tables[0].calculatedColumns[2], null); await document.save();
  applied.redo(); await document.save(); const saved=analyze(await fs.readFile(file)); assert.equal(saved.sheets[0].cells.get('C3').formula,'B3*2'); assert.equal(saved.sheets[0].tables[0].calculatedColumns[2],'B2*2'); document.dispose();
});

test('계산 열은 값·수식·빈 셀 예외를 보존하고 명시적 전체 적용과 해제를 지원한다', () => {
  const bytes = calculatedTableXlsx({native:true,exceptions:{3:999,4:{formula:'B4*5'}}}), document = new JinSheetDocument(uri('/tmp/calculated-exceptions.xlsx'),'xlsx',analyze(bytes),bytes,true,'');
  document.applyCellEdit(1,2,'=B2*3'); assert.equal(document.getCell(2,2).value,'999'); assert.equal(document.getCell(3,2).formula,'B4*5'); assert.equal(document.state.sheets[0].tables[0].calculatedColumns[2],'B2*3');
  document.applyCellEdit(1,2,'=B2*4',0,new Set(),{skipCalculated:true}); assert.equal(document.state.sheets[0].tables[0].calculatedColumns[2],'B2*3');
  const kept=document.applyCalculatedColumn(0,1,2,'preserve'); assert.equal(kept.preserved,2); assert.equal(document.getCell(2,2).value,'999');
  document.applyCalculatedColumn(0,1,2,'all'); assert.deepEqual([1,2,3].map(row=>document.getCell(row,2).formula),['B2*4','B3*4','B4*4']);
  document.applyCellEdit(2,2,''); document.applyCellEdit(1,2,'=B2*6'); assert.equal(document.getCell(2,2).formula,null); assert.equal(document.getCell(2,2).value,''); assert.equal(document.getCell(3,2).formula,'B4*6');
  assert.equal(document.applyCalculatedColumn(0,1,2,'preserve').preserved,1); assert.equal(document.getCell(2,2).formula,null);
  document.applyCalculatedColumn(0,1,2,'preserve',true); assert.equal(document.state.sheets[0].tables[0].calculatedColumns[2],null); assert.equal(document.getCell(3,2).formula,'B4*6'); document.dispose();
});

test('빈 열의 구조 참조 자동 채우기는 필터로 숨긴 행도 갱신하고 전체 지우기는 정의를 해제한다', () => {
  const bytes=calculatedTableXlsx(), document=new JinSheetDocument(uri('/tmp/calculated-filter.xlsx'),'xlsx',analyze(bytes),bytes,true,'');
  document.applyAutoFilter(0,updateAutoFilterColumn(document,0,1,['20'],false,false,autoFilterRange(document,0,0,1)));
  document.applyCellEdit(2,2,'=[@Amount]*2'); assert.deepEqual([1,2,3].map(row=>document.getCell(row,2).formula),['[@Amount]*2','[@Amount]*2','[@Amount]*2']);
  assert.equal(getRange(document,{requestId:13,rowStart:5,columnStart:3,rowCount:1,columnCount:1}).rows[0][0].cachedValue,40);
  document.applyCellEdits([1,2,3].map(row=>({row,column:2,value:'',formula:null}))); assert.equal(document.state.sheets[0].tables[0].calculatedColumns[2],null); document.dispose();
});

test('계산 열 자동 전파의 유효성 오류는 전체 변경을 취소하고 경고 재시도도 정의를 보존한다', () => {
  for (const style of ['stop','warning']) {
    const files=unzipSync(calculatedTableXlsx()); files['xl/worksheets/sheet1.xml']=new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</worksheet>','<dataValidations count="1"><dataValidation type="whole" operator="between" errorStyle="'+style+'" sqref="C2:C4"><formula1>0</formula1><formula2>5</formula2></dataValidation></dataValidations></worksheet>'));
    const bytes=zipSync(files), document=new JinSheetDocument(uri('/tmp/calculated-validation.xlsx'),'xlsx',analyze(bytes),bytes,true,''); let pending;
    try {document.applyCellEdit(1,2,'=[@Amount]*2');assert.fail('validation must reject');}catch(error){pending=error}
    assert.equal(document.dirty,false); assert.equal(document.state.sheets[0].tables[0].calculatedColumns[2],null); assert.equal(document.formulaDrafts.size,0);
    if(style==='warning') {
      let prompts=0; while(pending) { assert.equal(pending.validationPrompt.style,'warning'); prompts++;try{pending.retry();pending=null}catch(error){pending=error} }
      assert.equal(prompts,3); assert.equal(document.state.sheets[0].tables[0].calculatedColumns[2],'[@Amount]*2'); assert.equal(document.formulaDrafts.size,3);
    }
    document.dispose();
  }
});

test('계산 열 UI는 자동 채우기 해제·적용 범위와 요청 실패·재시도를 처리한다', () => {
  const table={name:'Sales',range:{top:0,bottom:3,left:0,right:2},headerRowCount:1,totalsRowCount:0,columns:['Name','Amount','Result'],calculatedColumns:[null,null,'B2*2'],totalColumns:[{function:'none'},{function:'none'},{function:'none'}]};
  const {document,elements,messages,dispatch}=runGridWebview(renderHtml(),{kind:'xlsx',rowCount:4,columnCount:3,tables:[table],cellData:{'1:2':{formula:'B2*2',hasFormula:true}}});
  const cell=document.querySelector('[data-row="1"][data-column="2"]');cell.onclick({target:cell}); const apply=elements.get('applyCalculatedColumn'),scope=elements.get('calculatedColumnScope');
  assert.equal(apply.disabled,false);elements.get('autoCalculatedColumns').checked=false; const formula=elements.get('formula'); formula.value='=B2*5';formula.listeners.get('input')();formula.listeners.get('keydown')({target:formula,key:'Enter',preventDefault(){}});assert.equal(messages.findLast(message=>message.type==='edit').tableAutoFill,false);
  scope.value='all';apply.onclick();const request=messages.at(-1);assert.equal(request.type,'calculated-column');assert.equal(request.mode,'all');assert.equal(apply.disabled,true);assert.equal(elements.get('tableTotals').disabled,true);
  dispatch({type:'calculated-column-error',requestId:request.requestId+1,sheetIndex:0,tableName:'Sales',column:2,message:'stale'});assert.equal(apply.disabled,true);
  dispatch({type:'calculated-column-error',requestId:request.requestId,sheetIndex:0,tableName:'Sales',column:2,message:'retry'});assert.equal(apply.disabled,false);assert.equal(scope.value,'all');assert.equal(document.activeElement,scope);
  apply.onclick();assert.notEqual(messages.at(-1).requestId,request.requestId);const retry=messages.at(-1);dispatch({type:'calculated-column-error',requestId:retry.requestId,sheetIndex:0,tableName:'Sales',column:2,message:'retry'});
  elements.get('clearCalculatedColumn').onclick();assert.equal(messages.at(-1).remove,true);
});

test('계산 열은 통합문서 저장 셀 한도를 넘기기 전에 자동 전파 전체를 거부한다', () => {
  const files=unzipSync(calculatedTableXlsx()), base=analyze(zipSync(files)), count=100000-base.sheets.reduce((sum,sheet)=>sum+sheet.cells.size,0);
  files['xl/workbook.xml']=new TextEncoder().encode(new TextDecoder().decode(files['xl/workbook.xml']).replace('</sheets>','<sheet name="Capacity" sheetId="2" r:id="capacity"/></sheets>'));
  files['xl/_rels/workbook.xml.rels']=new TextEncoder().encode(new TextDecoder().decode(files['xl/_rels/workbook.xml.rels']).replace('</Relationships>','<Relationship Id="capacity" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/capacity.xml"/></Relationships>'));
  files['[Content_Types].xml']=new TextEncoder().encode(new TextDecoder().decode(files['[Content_Types].xml']).replace('</Types>','<Override PartName="/xl/worksheets/capacity.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'));
  files['xl/worksheets/capacity.xml']=new TextEncoder().encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'+Array.from({length:count},(_,index)=>'<row r="'+(index+1)+'"><c r="A'+(index+1)+'"><v>1</v></c></row>').join('')+'</sheetData></worksheet>');
  const bytes=zipSync(files,{level:0}), model=analyze(bytes);assert.equal(model.writable,true,model.readOnlyReason);assert.equal(model.sheets.reduce((sum,sheet)=>sum+sheet.cells.size,0),100000);
  const document=new JinSheetDocument(uri('/tmp/calculated-capacity.xlsx'),'xlsx',model,bytes,true,'');assert.throws(()=>document.applyCellEdit(1,2,'=B2*2'),/통합문서 100,000셀/);assert.equal(document.dirty,false);assert.equal(document.formulaDrafts.size,0);assert.equal(document.state.sheets[0].tables[0].calculatedColumns[2],null);document.dispose();
});

test('표 아래 입력은 계산 열·데이터 서식과 함께 확장되고 한 번의 Undo·저장을 지원한다', async t => {
  const {patchCells}=require('../src/xlsx');const original=analyze(calculatedTableXlsx({native:true})),styled=patchCells(original,[{sheetIndex:0,address:'B4',value:'30',font:{bold:true,italic:false}},{sheetIndex:0,address:'C4',value:'60',formula:'B4*2',font:{bold:false,italic:true}}]);
  const directory=await fs.mkdtemp(path.join(os.tmpdir(),'jinsheet-auto-expand-'));t.after(()=>fs.rm(directory,{recursive:true,force:true}));const file=path.join(directory,'expand.xlsx');await fs.writeFile(file,styled.bytes);
  const document=new JinSheetDocument(uri(file),'xlsx',styled.model,Buffer.from(styled.bytes),true,'');let operation,events=0;document.onDidChange(change=>{operation=change;events++});
  document.applyCellEdit(4,1,'7');const expanded=operation;assert.equal(events,1);assert.equal(document.state.sheets[0].tables[0].range.bottom,4);assert.equal(document.getCell(4,1).kind,'number');assert.equal(document.getCell(4,2).formula,'B5*2');
  assert.equal(getRange(document,{requestId:14,rowStart:4,columnStart:2,rowCount:1,columnCount:1}).rows[0][0].cachedValue,14);await document.save();let saved=analyze(await fs.readFile(file));assert.equal(saved.sheets[0].cells.get('B5').styleProperties.bold,true);assert.equal(saved.sheets[0].cells.get('C5').styleProperties.italic,true);
  expanded.undo();assert.equal(document.state.sheets[0].tables[0].range.bottom,3);assert.equal(document.getCell(4,1).value,'');assert.equal(document.getCell(4,2).formula,null);await document.save();
  expanded.redo();await document.save();saved=analyze(await fs.readFile(file));assert.equal(saved.sheets[0].tables[0].autoFilter.bottom,4);assert.equal(saved.sheets[0].cells.get('B5').value,'7');assert.equal(saved.sheets[0].cells.get('D6').formula,'SUBTOTAL(9,Sales[Result])');document.dispose();
});

test('표 경계의 연속 붙여넣기는 행·열을 함께 확장하고 직접 지정한 값을 보존한다', () => {
  const bytes=calculatedTableXlsx({native:true}),document=new JinSheetDocument(uri('/tmp/auto-expand-paste.xlsx'),'xlsx',analyze(bytes),bytes,true,'');let events=0;document.onDidChange(()=>events++);
  document.applyCellEdits([{row:3,column:1,value:'31',formula:null},{row:3,column:3,value:'x',formula:null},{row:4,column:1,value:'7',formula:null},{row:4,column:3,value:'y',formula:null},{row:5,column:1,value:'8',formula:null},{row:5,column:3,value:'z',formula:null}]);
  const table=document.state.sheets[0].tables[0];assert.deepEqual(table.range,{top:0,bottom:5,left:0,right:3});assert.deepEqual(table.columns,['Name','Amount','Result','Column4']);assert.equal(events,1);assert.equal(document.getCell(0,3).value,'Column4');assert.equal(document.getCell(5,3).value,'z');
  assert.deepEqual(getRange(document,{requestId:15,rowStart:3,columnStart:2,rowCount:3,columnCount:1}).rows.map(row=>row[0].cachedValue),[62,14,16]);document.dispose();
});

test('오른쪽 입력은 고유한 머리글·여러 줄 이름과 새 계산 열을 원자적으로 만든다', () => {
  const bytes=calculatedTableXlsx({native:true}),document=new JinSheetDocument(uri('/tmp/auto-expand-columns.xlsx'),'xlsx',analyze(bytes),bytes,true,'');let operation;document.onDidChange(change=>operation=change);
  document.applyCellEdit(0,3,'Amount');assert.equal(document.getCell(0,3).value,'Amount2');assert.equal(document.state.sheets[0].tables[0].columns[3],'Amount2');const header=operation;header.undo();assert.equal(document.state.sheets[0].tables[0].columns.length,3);header.redo();
  document.applyCellEdit(1,4,'=[@Amount]*3');assert.equal(document.state.sheets[0].tables[0].columns[4],'Column5');assert.equal(document.state.sheets[0].tables[0].calculatedColumns[4],'[@Amount]*3');assert.deepEqual(getRange(document,{requestId:16,rowStart:1,columnStart:4,rowCount:3,columnCount:1}).rows.map(row=>row[0].cachedValue),[30,60,90]);
  document.applyCellEdit(0,5,'line\nname');assert.equal(document.state.sheets[0].tables[0].columns[5],'line\nname');assert.equal(document.getCell(0,5).value,'line\nname');assert.match(new TextDecoder().decode(document.state.files['xl/tables/table1.xml']),/line&#10;name/);document.dispose();
});

test('표 자동 확장은 빈 입력·간격·대각선·끄기와 합계 아래 입력을 구분한다', () => {
  let bytes=calculatedTableXlsx({native:true}),document=new JinSheetDocument(uri('/tmp/auto-expand-boundary.xlsx'),'xlsx',analyze(bytes),bytes,true,'');
  document.applyCellEdit(5,1,'9');assert.equal(document.state.sheets[0].tables[0].range.bottom,3);document.applyCellEdit(5,3,'note');assert.equal(document.state.sheets[0].tables[0].range.right,2);
  document.applyCellEdit(4,1,'7',0,new Set(),{skipExpansion:true});assert.equal(document.state.sheets[0].tables[0].range.bottom,3);document.applyCellEdit(4,1,'8');assert.equal(document.state.sheets[0].tables[0].range.bottom,4);assert.equal(document.state.sheets[0].tables[0].range.right,2);assert.equal(document.getCell(5,1).value,'9');document.dispose();
  bytes=totalsTableXlsx();document=new JinSheetDocument(uri('/tmp/auto-expand-totals.xlsx'),'xlsx',analyze(bytes),bytes,true,'');document.applyCellEdit(4,1,'7');assert.equal(document.state.sheets[0].tables[0].range.bottom,3);assert.equal(document.state.sheets[0].tables[0].totalsRowCount,1);assert.equal(document.getCell(3,1).formula,'SUBTOTAL(109,[Amount])');document.dispose();
});

test('자동 확장의 유효성 오류·경고 재시도는 데이터와 표 범위를 함께 원자적으로 처리한다', () => {
  for(const style of ['stop','warning']) {
    const files=unzipSync(calculatedTableXlsx({native:true}));files['xl/worksheets/sheet1.xml']=new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</worksheet>','<dataValidations count="1"><dataValidation type="whole" operator="between" errorStyle="'+style+'" sqref="B5:B6"><formula1>0</formula1><formula2>5</formula2></dataValidation></dataValidations></worksheet>'));
    const bytes=zipSync(files),document=new JinSheetDocument(uri('/tmp/auto-expand-validation.xlsx'),'xlsx',analyze(bytes),bytes,true,'');let pending,events=0;document.onDidChange(()=>events++);
    try{document.applyCellEdits([{row:4,column:1,value:'7'},{row:5,column:1,value:'8'}]);assert.fail('validation must reject')}catch(error){pending=error}
    assert.equal(document.dirty,false);assert.equal(document.state.sheets[0].tables[0].range.bottom,3);assert.equal(document.getCell(4,1).value,'');assert.equal(events,0);
    if(style==='warning'){let count=0;while(pending){assert.equal(pending.validationPrompt.style,'warning');count++;try{pending.retry();pending=null}catch(error){pending=error}}assert.equal(count,2);assert.equal(events,1);assert.equal(document.state.sheets[0].tables[0].range.bottom,5);assert.deepEqual(getRange(document,{requestId:17,rowStart:4,columnStart:2,rowCount:2,columnCount:1}).rows.map(row=>row[0].cachedValue),[14,16]);}
    document.dispose();
  }
});

test('자동 확장 경고 대기 중 다른 편집이 생기면 오래된 계획으로 문서를 덮어쓰지 않는다', () => {
  const files=unzipSync(calculatedTableXlsx({native:true}));files['xl/worksheets/sheet1.xml']=new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</worksheet>','<dataValidations count="1"><dataValidation type="whole" operator="between" errorStyle="warning" sqref="B5"><formula1>0</formula1><formula2>5</formula2></dataValidation></dataValidations></worksheet>'));
  const bytes=zipSync(files),document=new JinSheetDocument(uri('/tmp/auto-expand-stale.xlsx'),'xlsx',analyze(bytes),bytes,true,'');let warning;try{document.applyCellEdit(4,1,'7')}catch(error){warning=error}assert.equal(warning.validationPrompt.style,'warning');
  document.applyCellEdit(1,0,'Changed');assert.throws(()=>warning.retry(),/문서가 변경되어/);assert.equal(document.getCell(1,0).value,'Changed');assert.equal(document.getCell(4,1).value,'');assert.equal(document.state.sheets[0].tables[0].range.bottom,3);document.dispose();
});

test('자동 확장 뒤 필터는 새 값과 계산 열을 반영하고 수동 숨김과 구분한다', async t => {
  const bytes=calculatedTableXlsx({native:true}),directory=await fs.mkdtemp(path.join(os.tmpdir(),'jinsheet-auto-filter-expand-'));t.after(()=>fs.rm(directory,{recursive:true,force:true}));const file=path.join(directory,'filtered.xlsx');await fs.writeFile(file,bytes);const document=new JinSheetDocument(uri(file),'xlsx',analyze(bytes),bytes,true,'');
  document.applyAutoFilter(0,updateAutoFilterColumn(document,0,1,['20'],false,false,autoFilterRange(document,0,0,1)));document.applyCellEdit(4,1,'20');assert.deepEqual(viewModel(document).hiddenRows,[1,3]);assert.equal(getRange(document,{requestId:18,rowStart:5,columnStart:3,rowCount:1,columnCount:1}).rows[0][0].cachedValue,80);
  await document.save();const saved=analyze(await fs.readFile(file));assert.deepEqual(saved.sheets[0].nativeHiddenRows,[1,3]);assert.deepEqual(saved.sheets[0].hiddenRows,[]);assert.equal(saved.sheets[0].tables[0].autoFilter.bottom,4);document.dispose();
});

test('표 확장 설정은 입력과 붙여넣기에 전달하고 차트 문서에서는 사용 불가를 표시한다', () => {
  const table={name:'Sales',range:{top:0,bottom:2,left:0,right:1},columns:['Name','Amount'],headerRowCount:1,totalsRowCount:0,calculatedColumns:[null,null],totalColumns:[{function:'none'},{function:'none'}]};const {elements,messages,documentListeners,dispatch}=runGridWebview(renderHtml(),{kind:'xlsx',rowCount:4,columnCount:3,tables:[table]});
  assert.equal(elements.get('autoExpandTables').checked,true);elements.get('autoExpandTables').checked=false;const formula=elements.get('formula');formula.value='7';formula.listeners.get('input')();formula.listeners.get('keydown')({target:formula,key:'Enter',preventDefault(){}});assert.equal(messages.findLast(message=>message.type==='edit').tableAutoExpand,false);
  documentListeners.get('keydown')({target:elements.get('grid'),key:'v',ctrlKey:true,metaKey:false,preventDefault(){}});assert.equal(messages.findLast(message=>message.type==='paste').tableAutoExpand,false);
  dispatch({type:'model',kind:'xlsx',readOnly:false,sheetIndex:0,rowCount:4,columnCount:3,tables:[table],tableExpansionAvailable:false});assert.equal(elements.get('autoExpandTables').disabled,true);assert.match(elements.get('autoExpandTables').title,/charts/);
});

test('표 머리글 변경은 구조화 참조·계산 열·저장·Undo를 함께 갱신한다', async t => {
  const {patchCells}=require('../src/xlsx');let model=analyze(calculatedTableXlsx({native:true}));
  model=patchCells(model,[2,3,4].map(row=>({sheetIndex:0,address:'C'+row,value:String((row-1)*20),formula:'[@Amount]*2'}))).model;
  const files={...model.files};files['xl/tables/table1.xml']=new TextEncoder().encode(new TextDecoder().decode(files['xl/tables/table1.xml']).replace('B2*2','[@Amount]*2'));
  const bytes=zipSync(files),directory=await fs.mkdtemp(path.join(os.tmpdir(),'jinsheet-header-'));t.after(()=>fs.rm(directory,{recursive:true,force:true}));const file=path.join(directory,'headers.xlsx');await fs.writeFile(file,bytes);
  const document=new JinSheetDocument(uri(file),'xlsx',analyze(bytes),bytes,true,'');let operation,count=0;document.onDidChange(change=>{operation=change;count++});
  document.applyCellEdit(0,1,'가격');assert.equal(count,1);assert.deepEqual(document.state.sheets[0].tables[0].columns,['Name','가격','Result']);assert.equal(document.getCell(1,2).formula,'[@가격]*2');assert.equal(document.state.sheets[0].tables[0].calculatedColumns[2],'[@가격]*2');assert.equal(getRange(document,{requestId:1,rowStart:1,columnStart:2,rowCount:1,columnCount:1}).rows[0][0].cachedValue,20);
  const rename=operation;await document.save();let saved=analyze(await fs.readFile(file));assert.equal(saved.sheets[0].cells.get('B1').value,'가격');assert.equal(saved.sheets[0].cells.get('C4').formula,'[@가격]*2');rename.undo();assert.equal(document.getCell(0,1).value,'Amount');assert.equal(document.getCell(1,2).formula,'[@Amount]*2');rename.redo();assert.equal(document.getCell(1,2).formula,'[@가격]*2');
  document.applyCellEdit(0,2,'Total');assert.equal(document.getCell(5,3).formula,'SUBTOTAL(9,Sales[Total])');await document.save();saved=analyze(await fs.readFile(file));assert.equal(saved.sheets[0].tables[0].columns[2],'Total');document.dispose();
});

test('표 머리글 붙여넣기는 동시 이름 교환·중복·빈 값·수식 텍스트를 처리한다', () => {
  const bytes=calculatedTableXlsx(),document=new JinSheetDocument(uri('/tmp/header-swap.xlsx'),'xlsx',analyze(bytes),bytes,true,'');
  document.applyCellEdits([{row:0,column:0,value:'Amount'},{row:0,column:1,value:'Name'}]);assert.deepEqual(document.state.sheets[0].tables[0].columns,['Amount','Name','Result']);
  document.applyCellEdit(0,2,'name');assert.equal(document.getCell(0,2).value,'name2');document.applyCellEdit(0,2,'');assert.equal(document.getCell(0,2).value,'Column3');
  document.applyCellEdit(0,2,'=B2*2');assert.equal(document.getCell(0,2).value,'=B2*2');assert.equal(document.getCell(0,2).formula,null);assert.equal(document.state.sheets[0].tables[0].columns[2],'=B2*2');
  const before=document.state;assert.throws(()=>document.applyCellEdit(0,2,'x'.repeat(256)),/255/);assert.equal(document.state,before);document.dispose();
});

test('구조화 열 이름 변경은 문자열·다른 표·선택자·특수문자와 동시 교환을 구분한다', () => {
  const {renameStructuredColumns}=require('../src/xlsx');const maps=new Map([['sales',new Map([['amount',"금액]#@'"],['result','Amount']])]]);
  assert.equal(renameStructuredColumns('SUM(Sales[Amount],Other[Amount],"Sales[Amount]",[@Amount],Sales[[#Data],[Amount]:[Result]])',maps,'Sales'),'SUM(Sales[금액\']\'#\'@\'\'],Other[Amount],"Sales[Amount]",[@금액\']\'#\'@\'\'],Sales[[#Data],[금액\']\'#\'@\'\']:[Amount]])');
  assert.equal(renameStructuredColumns('SUM([Amount],Sales[Result])',maps,'Other'),'SUM([Amount],Sales[Amount])');
});

test('머리글 변경은 필터·기존 A1 이름·유효성·조건부 수식과 열 메타데이터를 보존한다', () => {
  const files=unzipSync(calculatedTableXlsx({native:true})),decode=part=>new TextDecoder().decode(files[part]),encode=(part,text)=>files[part]=new TextEncoder().encode(text);
  encode('xl/workbook.xml',decode('xl/workbook.xml').replace('</workbook>','<definedNames><definedName name="SalesValues">Data!$B$2:$B$4</definedName></definedNames></workbook>'));
  encode('xl/worksheets/sheet1.xml',decode('xl/worksheets/sheet1.xml').replace('</worksheet>','<conditionalFormatting sqref="B2:B4"><cfRule type="expression" priority="1" dxfId="0"><formula>B2&gt;0</formula></cfRule></conditionalFormatting><dataValidations count="1"><dataValidation type="decimal" operator="lessThan" sqref="A2:A4"><formula1>100</formula1></dataValidation></dataValidations></worksheet>'));
  encode('xl/tables/table1.xml',decode('xl/tables/table1.xml').replace('<autoFilter ref="A1:C4"/>','<autoFilter ref="A1:C4"><filterColumn colId="1"><customFilters><customFilter operator="greaterThan" val="15"/></customFilters></filterColumn></autoFilter>').replace('B2*2','[@Amount]*2'));
  encode('xl/styles.xml','<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellXfs count="1"><xf fontId="0" fillId="0" borderId="0" numFmtId="0"/></cellXfs><dxfs count="1"><dxf><font><b/></font></dxf></dxfs></styleSheet>');
  encode('xl/_rels/workbook.xml.rels',decode('xl/_rels/workbook.xml.rels').replace('</Relationships>','<Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'));
  encode('[Content_Types].xml',decode('[Content_Types].xml').replace('</Types>','<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'));
  const bytes=zipSync(files),model=analyze(bytes);assert.equal(model.writable,true,model.readOnlyReason);const document=new JinSheetDocument(uri('/tmp/header-metadata.xlsx'),'xlsx',model,bytes,true,'');document.applyCellEdit(0,1,'$&\n금액');
  const table=document.state.sheets[0].tables[0];assert.equal(table.columns[1],'$&\n금액');assert.equal(table.calculatedColumns[2],'[@$&\n금액]*2');assert.equal(table.autoFilter.columns[0].column,1);assert.match(new TextDecoder().decode(document.state.files['xl/workbook.xml']),/Data!\$B\$2:\$B\$4/);assert.match(document.state.sheets[0].xml,/<formula>B2&gt;0<\/formula>/);assert.match(document.state.sheets[0].xml,/<formula1>100<\/formula1>/);assert.deepEqual(viewModel(document).hiddenRows,[1]);document.dispose();
});

test('머리글 변경의 유효성 경고·실패와 미저장 수식은 전체 작업으로 유지된다', () => {
  for(const style of ['stop','warning']) {
    const files=unzipSync(calculatedTableXlsx({native:true}));files['xl/worksheets/sheet1.xml']=new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</worksheet>','<dataValidations count="1"><dataValidation type="textLength" operator="lessThan" errorStyle="'+style+'" sqref="B1"><formula1>3</formula1></dataValidation></dataValidations></worksheet>'));
    const bytes=zipSync(files),document=new JinSheetDocument(uri('/tmp/header-warning.xlsx'),'xlsx',analyze(bytes),bytes,true,'');document.applyCellEdit(1,2,'=SUM(Sales[Amount])');let pending;const before=document.state;
    try{document.applyCellEdit(0,1,'Revenue');assert.fail('must reject')}catch(error){pending=error}
    assert.equal(document.state,before);assert.equal(document.getCell(0,1).value,'Amount');assert.equal(document.getCell(1,2).formula,'SUM(Sales[Amount])');
    if(style==='warning'){pending.retry();assert.equal(document.getCell(1,2).formula,'SUM(Sales[Revenue])');assert.equal(document.getCell(0,1).value,'Revenue')}
    document.dispose();
  }
});

test('머리글 변경과 양 방향 표 확장 붙여넣기는 하나의 Undo로 복구된다', () => {
  const bytes=calculatedTableXlsx(),document=new JinSheetDocument(uri('/tmp/header-expansion.xlsx'),'xlsx',analyze(bytes),bytes,true,'');let operation,count=0;document.onDidChange(change=>{operation=change;count++});
  document.applyCellEdits([{row:0,column:1,value:'Revenue'},{row:0,column:3,value:'Extra'},{row:4,column:1,value:'7'}]);assert.equal(count,1);assert.deepEqual(document.state.sheets[0].tables[0].columns,['Name','Revenue','Result','Extra']);assert.equal(document.state.sheets[0].tables[0].range.bottom,4);operation.undo();assert.deepEqual(document.state.sheets[0].tables[0].columns,['Name','Amount','Result']);assert.equal(document.getCell(4,1)?.value??'','');document.dispose();
});

test('머리글 변경은 다른 시트의 참조·특수문자 이름과 메타데이터 인용부호를 처리한다', () => {
  const {patchCells,patchAddSheet}=require('../src/xlsx');let model=patchAddSheet(analyze(calculatedTableXlsx())).model;model=patchCells(model,[{sheetIndex:1,address:'A1',value:'60',formula:'SUM(Sales[Amount])'},{sheetIndex:1,address:'B1',value:'0',formula:'"Sales[Amount]"'}]).model;
  const files={...model.files};files['xl/tables/table1.xml']=new TextEncoder().encode(new TextDecoder().decode(files['xl/tables/table1.xml']).replace('name="Amount"',"name='Amount'"));const bytes=zipSync(files),document=new JinSheetDocument(uri('/tmp/header-cross-sheet.xlsx'),'xlsx',analyze(bytes),bytes,true,'');
  document.applyCellEdit(0,1,"금액]#@'");assert.equal(document.getCell(0,0,1).formula,"SUM(Sales[금액']'#'@''])");assert.equal(document.getCell(0,1,1).formula,'"Sales[Amount]"');assert.equal(document.state.sheets[0].tables[0].columns[1],"금액]#@'");assert.equal(getRange(document,{requestId:12,sheetIndex:1,rowStart:0,columnStart:0,rowCount:1,columnCount:1}).rows[0][0].cachedValue,60);document.dispose();
});

test('머리글 경고를 기다리는 동안 다른 편집이 있으면 오래된 변경을 적용하지 않는다', () => {
  const files=unzipSync(calculatedTableXlsx());files['xl/worksheets/sheet1.xml']=new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</worksheet>','<dataValidations count="1"><dataValidation type="textLength" operator="lessThan" errorStyle="warning" sqref="B1"><formula1>3</formula1></dataValidation></dataValidations></worksheet>'));const bytes=zipSync(files),document=new JinSheetDocument(uri('/tmp/header-stale.xlsx'),'xlsx',analyze(bytes),bytes,true,'');let pending;
  try{document.applyCellEdit(0,1,'Revenue');assert.fail('must warn')}catch(error){pending=error}document.applyCellEdit(1,0,'keep');assert.throws(()=>pending.retry(),/문서가 변경/);assert.equal(document.getCell(1,0).value,'keep');assert.equal(document.getCell(0,1).value,'Amount');document.dispose();
});

test('새 빈 시트의 첫 입력은 여러 행·수식·서식과 함께 저장되고 Undo할 수 있다', async t => {
  const {patchAddSheet}=require('../src/xlsx'),model=patchAddSheet(analyze(calculatedTableXlsx())).model,directory=await fs.mkdtemp(path.join(os.tmpdir(),'jinsheet-empty-sheet-'));t.after(()=>fs.rm(directory,{recursive:true,force:true}));const file=path.join(directory,'empty.xlsx');await fs.writeFile(file,model.bytes);const document=new JinSheetDocument(uri(file),'xlsx',model,Buffer.from(model.bytes),true,'');let operation;document.onDidChange(change=>operation=change);
  document.applyCellEdits([{row:4,column:2,value:'4',kind:'number'},{row:0,column:1,value:'2',kind:'number'},{row:0,column:0,value:'6',formula:'B1+C5',kind:'number'}],1);const input=operation;await document.save();let saved=analyze(await fs.readFile(file));assert.equal(saved.sheets[1].cells.get('A1').formula,'B1+C5');assert.equal(saved.sheets[1].cells.get('C5').value,'4');assert.equal(saved.sheets[1].cells.size,3);input.undo();await document.save();assert.equal(analyze(await fs.readFile(file)).sheets[1].cells.size,0);input.redo();await document.save();assert.equal(analyze(await fs.readFile(file)).sheets[1].cells.size,3);document.dispose();
});

test('머리글을 참조하는 공유 수식 변경은 셀·메타데이터를 남기지 않고 거부한다', () => {
  const files=unzipSync(calculatedTableXlsx());files['xl/worksheets/sheet1.xml']=new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('<f>SUBTOTAL(9,Sales[Result])</f>','<f t="shared" si="0" ref="D6:D7">SUBTOTAL(9,Sales[Result])</f>'));
  const bytes=zipSync(files),model=analyze(bytes);assert.equal(model.writable,true,model.readOnlyReason);const document=new JinSheetDocument(uri('/tmp/header-shared.xlsx'),'xlsx',model,bytes,true,'');assert.throws(()=>document.applyCellEdit(0,2,'Total'),/공유·배열/);assert.equal(document.state,model);assert.equal(document.dirty,false);assert.equal(document.getCell(0,2).value,'Result');document.dispose();
});

test('표 마지막 셀의 새 행 추가는 입력·계산 열·서식·범위를 단일 Undo와 저장에 묶는다', async t => {
  const bytes=calculatedTableXlsx({native:true}),directory=await fs.mkdtemp(path.join(os.tmpdir(),'jinsheet-table-tab-'));t.after(()=>fs.rm(directory,{recursive:true,force:true}));const file=path.join(directory,'tab.xlsx');await fs.writeFile(file,bytes);const document=new JinSheetDocument(uri(file),'xlsx',analyze(bytes),bytes,true,'');let operation,count=0;document.onDidChange(change=>{operation=change;count++});
  const result=document.appendTableRow(0,'Sales',3,2,'=B4*3');assert.deepEqual(result,{row:4,column:0});assert.equal(count,1);assert.equal(document.state.sheets[0].tables[0].range.bottom,4);assert.equal(document.getCell(3,2).formula,'B4*3');assert.equal(document.getCell(4,2).formula,'B5*3');assert.equal(document.state.sheets[0].tables[0].calculatedColumns[2],'B2*3');const appended=operation;
  await document.save();let saved=analyze(await fs.readFile(file));assert.equal(saved.sheets[0].tables[0].autoFilter.bottom,4);assert.equal(saved.sheets[0].cells.get('C5').formula,'B5*3');appended.undo();assert.equal(document.getCell(3,2).formula,'B4*2');assert.equal(document.state.sheets[0].tables[0].range.bottom,3);appended.redo();assert.equal(document.getCell(4,2).formula,'B5*3');await document.save();document.dispose();
});

test('표 Tab 행 추가는 합계 이동과 기존 참조를 갱신하고 아래 충돌을 원자적으로 거부한다', () => {
  const {patchCells}=require('../src/xlsx'),patched=patchCells(analyze(totalsTableXlsx()),[{sheetIndex:0,address:'B5',value:'7',kind:'number'}]),bytes=patched.bytes,document=new JinSheetDocument(uri('/tmp/tab-totals.xlsx'),'xlsx',patched.model,bytes,true,'');const before=document.state;
  assert.throws(()=>document.appendTableRow(0,'Sales',2,1,'25'),/표 아래에 데이터/);assert.equal(document.state,before);assert.equal(document.dirty,false);assert.equal(document.getCell(2,1).value,'20');
  document.applyCellEdit(4,1,'');const result=document.appendTableRow(0,'Sales',2,1,'25');assert.deepEqual(result,{row:3,column:0});assert.equal(document.state.sheets[0].tables[0].range.bottom,4);assert.equal(document.getCell(3,1)?.value??'','');assert.equal(document.getCell(4,1).formula,'SUBTOTAL(109,[Amount])');assert.equal(document.getCell(3,2).formula,'SUM(Sales[[#Totals],[Amount]])');document.dispose();
});

test('표 Tab 행 추가는 머리글·중간 셀과 병합 영역을 보호한다', () => {
  const bytes=calculatedTableXlsx({native:true}),document=new JinSheetDocument(uri('/tmp/tab-guards.xlsx'),'xlsx',analyze(bytes),bytes,true,'');for(const [row,column] of [[0,2],[2,2],[3,1]])assert.throws(()=>document.appendTableRow(0,'Sales',row,column),/마지막 데이터 셀/);assert.equal(document.dirty,false);document.dispose();
  const files=unzipSync(bytes);files['xl/worksheets/sheet1.xml']=new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</worksheet>','<mergeCells count="1"><mergeCell ref="A5:B5"/></mergeCells></worksheet>'));const merged=zipSync(files),blocked=new JinSheetDocument(uri('/tmp/tab-merged.xlsx'),'xlsx',analyze(merged),merged,true,'');assert.throws(()=>blocked.appendTableRow(0,'Sales',3,2),/병합/);assert.equal(blocked.dirty,false);blocked.dispose();
});

test('표 Tab UI는 성공 응답 뒤에만 이동하고 중복·오래된 응답과 범위 선택을 구분한다', () => {
  const table={name:'Sales',range:{top:0,bottom:3,left:0,right:2},headerRowCount:1,totalsRowCount:0,columns:['Name','Amount','Result'],calculatedColumns:[null,null,null],totalColumns:[{function:'none'},{function:'none'},{function:'none'}]};const {document,elements,messages,dispatch,documentListeners}=runGridWebview(renderHtml(),{kind:'xlsx',rowCount:6,columnCount:5,tables:[table]});
  const click=(row,column)=>{const cell=document.querySelector('[data-row="'+row+'"][data-column="'+column+'"]');cell.onclick({target:cell});return cell};const tab=(cell,shiftKey=false)=>documentListeners.get('keydown')({target:cell,key:'Tab',shiftKey,preventDefault(){}});
  let cell=click(1,2);tab(cell);assert.equal(elements.get('address').value,'A3');cell=click(2,0);tab(cell,true);assert.equal(elements.get('address').value,'C2');cell=click(3,2);tab(cell);const request=messages.at(-1);assert.equal(request.type,'table-tab');assert.equal(elements.get('address').value,'C4');tab(cell);assert.equal(messages.filter(message=>message.type==='table-tab').length,1);
  dispatch({...request,type:'table-tab-applied',requestId:request.requestId+1,row:4,column:0});assert.equal(elements.get('address').value,'C4');dispatch({...request,type:'table-tab-applied',row:4,column:0});assert.equal(elements.get('address').value,'A5');
  cell=click(3,2);elements.get('autoExpandTables').checked=false;tab(cell);assert.equal(elements.get('address').value,'D4');
});

test('표 Tab 입력의 유효성 경고 재시도는 단일 변경이며 다른 편집 후에는 거부한다', () => {
  for(const stale of [false,true]){
    const files=unzipSync(calculatedTableXlsx({native:true}));files['xl/worksheets/sheet1.xml']=new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</worksheet>','<dataValidations count="1"><dataValidation type="whole" operator="lessThan" errorStyle="warning" sqref="C4"><formula1>100</formula1></dataValidation></dataValidations></worksheet>'));const bytes=zipSync(files),document=new JinSheetDocument(uri('/tmp/tab-warning.xlsx'),'xlsx',analyze(bytes),bytes,true,'');let pending,count=0;document.onDidChange(()=>count++);
    try{document.appendTableRow(0,'Sales',3,2,'200');assert.fail('must warn')}catch(error){pending=error}assert.equal(document.dirty,false);assert.equal(count,0);assert.equal(document.getCell(3,2).formula,'B4*2');
    if(stale){document.applyCellEdit(1,0,'keep');assert.throws(()=>pending.retry(),/문서가 변경/);assert.equal(document.getCell(1,0).value,'keep');assert.equal(document.state.sheets[0].tables[0].range.bottom,3)}else{assert.deepEqual(pending.retry(),{row:4,column:0});assert.equal(count,1);assert.equal(document.getCell(3,2).value,'200');assert.equal(document.state.sheets[0].tables[0].range.bottom,4)}document.dispose();
  }
});

test('표 Tab 새 행이 필터로 숨겨지면 입력과 행 생성을 함께 취소한다', () => {
  const files=unzipSync(calculatedTableXlsx({native:true}));files['xl/tables/table1.xml']=new TextEncoder().encode(new TextDecoder().decode(files['xl/tables/table1.xml']).replace('<autoFilter ref="A1:C4"/>','<autoFilter ref="A1:C4"><filterColumn colId="1"><customFilters><customFilter operator="greaterThan" val="0"/></customFilters></filterColumn></autoFilter>'));const bytes=zipSync(files),document=new JinSheetDocument(uri('/tmp/tab-filter.xlsx'),'xlsx',analyze(bytes),bytes,true,'');assert.throws(()=>document.appendTableRow(0,'Sales',3,2,'70'),/필터에 가려/);assert.equal(document.dirty,false);assert.equal(document.getCell(3,2).value,'60');assert.equal(document.state.sheets[0].tables[0].range.bottom,3);document.dispose();
});

test('표 Tab 편집 UI는 셀·입력줄의 변경을 한 요청으로 보내고 실패 시 입력을 복구한다', () => {
  const table={name:'Sales',range:{top:0,bottom:3,left:0,right:2},headerRowCount:1,totalsRowCount:0,columns:['Name','Amount','Result'],calculatedColumns:[null,null,null],totalColumns:[{function:'none'},{function:'none'},{function:'none'}]};const grid=runGridWebview(renderHtml(),{kind:'xlsx',rowCount:6,columnCount:5,tables:[table]});const {document,elements,messages,dispatch}=grid;const cell=document.querySelector('[data-row="3"][data-column="2"]');cell.onclick({target:cell});const formula=elements.get('formula');formula.value='new value';formula.listeners.get('input')();formula.listeners.get('keydown')({target:formula,key:'Tab',preventDefault(){}});const request=messages.at(-1);assert.equal(request.type,'table-tab');assert.equal(request.value,'new value');assert.equal(messages.some(message=>message.type==='edit'),false);dispatch({...request,type:'table-tab-error',message:'blocked'});assert.equal(elements.get('address').value,'C4');assert.equal(formula.value,'new value');assert.equal(document.activeElement,formula);assert.equal(elements.get('notice').textContent,'blocked');
});

test('표 Tab 행 추가는 숨긴 열을 건너뛰고 새 셀의 데이터 서식을 유지한다', () => {
  const {patchCells,patchHiddenColumns}=require('../src/xlsx');let model=patchCells(analyze(calculatedTableXlsx({native:true})),[{sheetIndex:0,address:'B4',value:'30',font:{bold:true,italic:false}},{sheetIndex:0,address:'C4',value:'60',formula:'B4*2',font:{bold:false,italic:true}}]).model;model=patchHiddenColumns(model,[{sheetIndex:0,columns:[0,2],hidden:true}]).model;const document=new JinSheetDocument(uri('/tmp/tab-hidden.xlsx'),'xlsx',model,Buffer.from(model.bytes),true,'');assert.deepEqual(document.appendTableRow(0,'Sales',3,1,'35'),{row:4,column:1});assert.equal(document.getCell(4,1).style,document.getCell(3,1).style);assert.equal(document.getCell(4,2).style,document.getCell(3,2).style);assert.equal(document.getCell(4,2).formula,'B5*2');assert.deepEqual(document.state.sheets[0].hiddenColumns,[0,2]);document.dispose();
});

test('표 범위 안 Tab 순환은 행을 만들지 않고 셀 편집 Tab은 입력을 원자적으로 전달한다', () => {
  const table={name:'Sales',range:{top:0,bottom:3,left:0,right:2},headerRowCount:1,totalsRowCount:0,columns:['Name','Amount','Result'],calculatedColumns:[null,null,null],totalColumns:[{function:'none'},{function:'none'},{function:'none'}]};const {document,elements,messages,documentListeners}=runGridWebview(renderHtml(),{kind:'xlsx',rowCount:6,columnCount:5,tables:[table]});const first=document.querySelector('[data-row="2"][data-column="0"]'),last=document.querySelector('[data-row="3"][data-column="2"]');first.onmousedown({shiftKey:false,clientX:80,clientY:60,preventDefault(){}});last.onmouseenter();documentListeners.get('mouseup')();documentListeners.get('keydown')({target:last,key:'Tab',preventDefault(){}});assert.equal(elements.get('address').value,'A3');assert.equal(messages.some(message=>message.type==='table-tab'),false);
  last.onmousedown({target:last,shiftKey:false,clientX:80,clientY:60,preventDefault(){}});documentListeners.get('mouseup')();last.onclick({target:last});last.onkeydown({target:last,key:'x',isComposing:false,ctrlKey:false,metaKey:false,altKey:false,preventDefault(){},stopPropagation(){}});const input=last.children[0];input.value='8';input.blur=()=>input.onblur();input.onkeydown({key:'Tab',preventDefault(){},stopPropagation(){}});const request=messages.at(-1);assert.equal(request.type,'table-tab');assert.equal(request.value,'8');assert.equal(messages.some(message=>message.type==='edit'),false);
});

test('늦은 표 Tab 응답은 검색 포커스를 가져오지 않고 대기 중 새 입력도 버리지 않는다', () => {
  const table={name:'Sales',range:{top:0,bottom:3,left:0,right:2},headerRowCount:1,totalsRowCount:0,columns:['Name','Amount','Result'],calculatedColumns:[null,null,null],totalColumns:[{function:'none'},{function:'none'},{function:'none'}]};const {document,elements,messages,dispatch,documentListeners}=runGridWebview(renderHtml(),{kind:'xlsx',rowCount:6,columnCount:5,tables:[table]});let cell=document.querySelector('[data-row="3"][data-column="2"]');cell.onclick({target:cell});documentListeners.get('keydown')({target:cell,key:'Tab',preventDefault(){}});let request=messages.at(-1);elements.get('search').focus();dispatch({...request,type:'table-tab-applied',row:4,column:0});assert.equal(elements.get('address').value,'C4');assert.equal(document.activeElement,elements.get('search'));
  cell.onclick({target:cell});const formula=elements.get('formula');formula.value='first';formula.listeners.get('input')();formula.listeners.get('keydown')({target:formula,key:'Tab',preventDefault(){}});request=messages.at(-1);assert.equal(request.type,'table-tab');formula.value='second';formula.listeners.get('input')();formula.listeners.get('keydown')({target:formula,key:'Tab',preventDefault(){}});assert.equal(messages.findLast(message=>message.type==='edit').value,'second');dispatch({...request,type:'table-tab-applied',row:4,column:0});assert.equal(elements.get('address').value,'D4');
});

test('상위 순위 필터는 동점·수식·수동 숨김·native 기준값을 저장과 Undo에 연결한다', async t => {
  const {patchCells}=require('../src/xlsx'),{updateAutoFilterTopColumn}=require('../src/extension');let model=patchCells(analyze(calculatedTableXlsx({native:true})),[{sheetIndex:0,address:'B4',value:'20'},{sheetIndex:0,address:'C4',value:'40',formula:'B4*2'},{sheetIndex:0,address:'E6',value:'0',formula:'SUBTOTAL(109,Sales[Result])'}]).model;
  const directory=await fs.mkdtemp(path.join(os.tmpdir(),'jinsheet-top-filter-'));t.after(()=>fs.rm(directory,{recursive:true,force:true}));const file=path.join(directory,'top.xlsx');await fs.writeFile(file,model.bytes);const document=new JinSheetDocument(uri(file),'xlsx',model,Buffer.from(model.bytes),true,'');let operation;document.onDidChange(change=>operation=change);
  const range=autoFilterRange(document,0,0,1);document.applyHiddenRows(0,[2],true);document.applyAutoFilter(0,updateAutoFilterTopColumn(document,0,1,{top:true,percent:false,value:1},range));const ranked=operation;assert.deepEqual(viewModel(document).hiddenRows,[1,2]);const sums=()=>getRange(document,{requestId:1,rowStart:5,columnStart:3,rowCount:1,columnCount:2}).rows[0].map(cell=>cell.cachedValue);assert.deepEqual(sums(),[80,40]);await document.save();let saved=analyze(await fs.readFile(file));assert.deepEqual(saved.sheets[0].tables[0].autoFilter.columns[0].top10,{top:true,percent:false,value:1,filterValue:20});assert.deepEqual(saved.sheets[0].nativeHiddenRows,[1,2]);assert.deepEqual(saved.sheets[0].hiddenRows,[2]);
  ranked.undo();assert.deepEqual(viewModel(document).hiddenRows,[2]);ranked.redo();assert.deepEqual(sums(),[80,40]);document.applyCellEdit(3,1,'5');assert.deepEqual(sums(),[40,0]);assert.deepEqual(viewModel(document).hiddenRows,[1,2,3]);await document.save();saved=analyze(await fs.readFile(file));assert.equal(saved.sheets[0].tables[0].autoFilter.columns[0].top10.filterValue,20);
  document.applyCellEdit(2,1,'50');await document.save();saved=analyze(await fs.readFile(file));assert.equal(saved.sheets[0].tables[0].autoFilter.columns[0].top10.filterValue,50);const options=autoFilterOptions(document,0,1,autoFilterRange(document,0,0,1));assert.equal(options.top10.value,1);assert.deepEqual(options.selectedValues,['50']);document.dispose();
  const reopened=new JinSheetDocument(uri(file),'xlsx',saved,Buffer.from(saved.bytes),true,'');reopened.applyAutoFilter(0,updateAutoFilterColumn(reopened,0,1,[],false,true,autoFilterRange(reopened,0,0,1)));assert.deepEqual(viewModel(reopened).hiddenRows,[2]);reopened.dispose();
});

test('하위 백분율 필터는 전체 모집단·다른 열 필터와 구조 변경을 유지한다', () => {
  const {updateAutoFilterTopColumn}=require('../src/extension');const bytes=calculatedTableXlsx({native:true}),document=new JinSheetDocument(uri('/tmp/top-percent.xlsx'),'xlsx',analyze(bytes),bytes,true,'');const range=autoFilterRange(document,0,0,1);document.applyAutoFilter(0,updateAutoFilterTopColumn(document,0,1,{top:false,percent:true,value:34},range));assert.deepEqual(viewModel(document).hiddenRows,[3]);
  const current=autoFilterRange(document,0,0,0);document.applyAutoFilter(0,updateAutoFilterColumn(document,0,0,['item1','item2'],false,false,current));assert.deepEqual(viewModel(document).hiddenRows,[1,3]);document.applyStructureChange({sheetIndex:0,axis:'row',action:'insert',index:2});const filter=document.state.sheets[0].tables[0].autoFilter;assert.equal(filter.bottom,4);assert.deepEqual(filter.columns.find(column=>column.top10).top10,{top:false,percent:true,value:34,filterValue:20});document.dispose();
});

test('상위 필터 UI는 명시적 적용·재개방·오류와 열별 해제를 처리한다', () => {
  const {elements,messages,dispatch}=runGridWebview(renderHtml(),{kind:'xlsx',rowCount:5,columnCount:2,autoFilter:{top:0,bottom:4,left:0,right:1,columns:[]}});elements.get('filterValuesButton').onclick();const request=messages.at(-1);dispatch({type:'filter-options',sheetIndex:0,column:request.column,options:{values:['10','20'],selectedValues:['20'],blanks:false,includeBlanks:false,top10:{top:false,percent:true,value:25}}});assert.equal(elements.get('filterMode').value,'top10');assert.equal(elements.get('filterTopDirection').value,'bottom');assert.equal(elements.get('filterTopValue').value,'25');assert.equal(elements.get('filterTopControls').hidden,false);assert.equal(elements.get('filterValueControls').hidden,true);
  elements.get('filterTopValue').value='0';assert.equal(messages.some(message=>message.type==='auto-filter'),false);elements.get('filterApply').onclick();let submitted=messages.at(-1);assert.deepEqual(JSON.parse(JSON.stringify(submitted.top10)),{top:false,percent:true,value:0});dispatch({type:'filter-error',requestId:submitted.requestId,sheetIndex:0,column:request.column,message:'invalid'});assert.equal(elements.get('filterEditor').disabled,false);assert.equal(elements.get('filterTopValue').value,'0');assert.equal(elements.get('filterMode').value,'top10');
  elements.get('filterTopValue').value='20';elements.get('filterApply').onclick();submitted=messages.at(-1);dispatch({type:'filter-applied',requestId:submitted.requestId,sheetIndex:0,column:request.column});assert.equal(elements.get('filterPanel').hidden,true);
});

test('표 순위 필터는 합계를 모집단에서 제외하고 계산 열의 숫자 결과로 순위를 매긴다', () => {
  const {updateAutoFilterTopColumn}=require('../src/extension');const bytes=totalsTableXlsx(),document=new JinSheetDocument(uri('/tmp/rank-totals.xlsx'),'xlsx',analyze(bytes),bytes,true,'');document.applyAutoFilter(0,updateAutoFilterTopColumn(document,0,1,{top:true,percent:false,value:1},autoFilterRange(document,0,0,1)));assert.deepEqual(viewModel(document).hiddenRows,[1]);assert.equal(getRange(document,{requestId:1,rowStart:3,columnStart:1,rowCount:1,columnCount:1}).rows[0][0].cachedValue,20);assert.equal(document.state.sheets[0].tables[0].autoFilter.columns[0].top10.filterValue,20);document.dispose();
  const formulaBytes=calculatedTableXlsx({native:true}),formulas=new JinSheetDocument(uri('/tmp/rank-formulas.xlsx'),'xlsx',analyze(formulaBytes),formulaBytes,true,'');formulas.applyAutoFilter(0,updateAutoFilterTopColumn(formulas,0,2,{top:false,percent:false,value:2},autoFilterRange(formulas,0,0,2)));assert.deepEqual(viewModel(formulas).hiddenRows,[3]);assert.equal(formulas.state.sheets[0].tables[0].autoFilter.columns[0].top10.filterValue,40);formulas.dispose();
});

function dateFilterXlsx(date1904=false) {
  const {patchCells}=require('../src/xlsx');let bytes=calculatedTableXlsx();if(date1904){const files=unzipSync(bytes);files['xl/workbook.xml']=new TextEncoder().encode(new TextDecoder().decode(files['xl/workbook.xml']).replace('<sheets>','<workbookPr date1904="1"/><sheets>'));bytes=zipSync(files)}
  const offset=date1904?1462:0;return patchCells(analyze(bytes),[{sheetIndex:0,address:'B2',value:String(45351.5-offset),kind:'number',numberFormatId:14},{sheetIndex:0,address:'B3',value:String(45352-offset),kind:'number',numberFormatId:14},{sheetIndex:0,address:'B4',value:'0',formula:date1904?'43889.75':'DATE(2024,2,29)+0.75',kind:'number',numberFormatId:14},...[10,20,30].map((value,index)=>({sheetIndex:0,address:'C'+(index+2),value:String(value),kind:'number'}))]).bytes;
}

test('날짜 그룹 필터는 수식·시간·1904 날짜의 표시와 native 저장·Undo를 함께 유지한다', async t => {
  const {updateAutoFilterDateColumn}=require('../src/extension'),{parseDateGroups}=require('../src/auto-filter');const directory=await fs.mkdtemp(path.join(os.tmpdir(),'jinsheet-date-filter-'));t.after(()=>fs.rm(directory,{recursive:true,force:true}));
  for(const date1904 of [false,true]){const bytes=dateFilterXlsx(date1904),file=path.join(directory,date1904?'1904.xlsx':'1900.xlsx');await fs.writeFile(file,bytes);const document=new JinSheetDocument(uri(file),'xlsx',analyze(bytes),bytes,true,'');let operation;document.onDidChange(change=>operation=change);document.applyAutoFilter(0,updateAutoFilterDateColumn(document,0,1,parseDateGroups('2024-02'),[],false,autoFilterRange(document,0,0,1)));const applied=operation;assert.deepEqual(viewModel(document).hiddenRows,[2]);assert.equal(getRange(document,{requestId:1,rowStart:5,columnStart:3,rowCount:1,columnCount:1}).rows[0][0].cachedValue,40);assert.equal(getRange(document,{requestId:2,rowStart:1,columnStart:1,rowCount:1,columnCount:1},'ko').rows[0][0].displayValue,'2024-02-29');await document.save();let saved=analyze(await fs.readFile(file));assert.equal(saved.date1904,date1904);assert.deepEqual(saved.sheets[0].tables[0].autoFilter.columns[0].dateGroups,[{dateTimeGrouping:'month',year:2024,month:2}]);assert.deepEqual(saved.sheets[0].nativeHiddenRows,[2]);applied.undo();assert.equal(viewModel(document).hiddenRows,undefined);applied.redo();assert.deepEqual(viewModel(document).hiddenRows,[2]);document.dispose();
    const reopened=new JinSheetDocument(uri(file),'xlsx',saved,Buffer.from(saved.bytes),true,'');const filter=autoFilterRange(reopened,0,0,1);assert.equal(autoFilterOptions(reopened,0,1,filter).dateGroups[0].month,2);reopened.applyAutoFilter(0,updateAutoFilterColumn(reopened,0,1,[],false,true,filter));assert.equal(viewModel(reopened).hiddenRows,undefined);reopened.dispose();}
});

test('날짜 그룹과 개별 값·빈 셀은 OR이며 다른 열의 조건과는 교차한다', () => {
  const {updateAutoFilterDateColumn}=require('../src/extension'),{parseDateGroups}=require('../src/auto-filter');const bytes=dateFilterXlsx(),document=new JinSheetDocument(uri('/tmp/date-mixed.xlsx'),'xlsx',analyze(bytes),bytes,true,'');document.applyCellEdit(2,1,'2024-02-29');document.applyAutoFilter(0,updateAutoFilterDateColumn(document,0,1,parseDateGroups('2024-02'),['2024-02-29'],false,autoFilterRange(document,0,0,1)));assert.equal(viewModel(document).hiddenRows,undefined);const options=autoFilterOptions(document,0,1,autoFilterRange(document,0,0,1));assert.deepEqual(options.dateValues,['2024-02-29']);document.applyAutoFilter(0,updateAutoFilterColumn(document,0,0,['item0','item1'],false,false,autoFilterRange(document,0,0,0)));assert.deepEqual(viewModel(document).hiddenRows,[3]);document.applyStructureChange({sheetIndex:0,axis:'row',action:'insert',index:2});assert.equal(document.state.sheets[0].tables[0].autoFilter.columns.find(item=>item.dateGroups).dateGroups[0].month,2);assert.equal(document.state.sheets[0].tables[0].autoFilter.bottom,4);document.dispose();
});

test('날짜 그룹 UI는 정밀도·혼합 값·빈 셀을 복원하고 명시적으로 적용한다', () => {
  const {elements,messages,dispatch}=runGridWebview(renderHtml(),{kind:'xlsx',rowCount:5,columnCount:2,autoFilter:{top:0,bottom:4,left:0,right:1,columns:[]}});elements.get('filterValuesButton').onclick();const request=messages.at(-1);dispatch({type:'filter-options',sheetIndex:0,column:request.column,options:{values:['text'],selectedValues:['text'],blanks:true,includeBlanks:true,dateGroups:[{dateTimeGrouping:'minute',year:2024,month:2,day:29,hour:12,minute:30}],dateValues:['text']}});assert.equal(elements.get('filterMode').value,'dates');assert.equal(elements.get('filterDateGroups').value,'2024-02-29T12:30');assert.equal(elements.get('filterDateBlanks').checked,true);assert.equal(elements.get('filterDateControls').hidden,false);assert.equal(elements.get('filterDateKeepLabel').hidden,false);
  elements.get('filterDateGroups').value='2024-02';assert.equal(messages.some(message=>message.type==='auto-filter'),false);elements.get('filterApply').onclick();const submitted=messages.at(-1);assert.equal(submitted.dateGroupText,'2024-02');assert.deepEqual(Array.from(submitted.values),['text']);dispatch({type:'filter-error',sheetIndex:0,column:request.column,requestId:submitted.requestId,message:'bad date'});assert.equal(elements.get('filterEditor').disabled,false);assert.equal(elements.get('filterDateGroups').value,'2024-02');elements.get('filterDateKeepValues').checked=false;elements.get('filterApply').onclick();assert.deepEqual(Array.from(messages.at(-1).values),[]);
});

test('1904 날짜 붙여넣기와 오늘 날짜 입력은 원래 날짜를 저장하고 표시한다', () => {
  const {buildDateTimeShortcutEdit,parseClipboardDate}=require('../src/extension');assert.deepEqual(parseClipboardDate('2024-02-29','en',true),{serial:43889,numFmtId:14});assert.deepEqual(parseClipboardDate('1904-01-01','en',true),{serial:0,numFmtId:14});assert.equal(parseClipboardDate('1900-01-01','en',true),null);assert.equal(parseClipboardDate('2024-02-29 12:30','en',true).serial,43889+12.5/24);
  const bytes=dateFilterXlsx(true),document=new JinSheetDocument(uri('/tmp/1904-input.xlsx'),'xlsx',analyze(bytes),bytes,true,'');const edit=buildDateTimeShortcutEdit(document,1,1,'date',new Date(2024,1,29,12,30));assert.equal(edit.value,'43889');document.applyCellEdits([edit]);assert.equal(getRange(document,{requestId:1,rowStart:1,columnStart:1,rowCount:1,columnCount:1},'ko').rows[0][0].displayValue,'2024-02-29');const time=buildDateTimeShortcutEdit(document,1,1,'time',new Date(2024,1,29,12,30));assert.equal(time.value,String(12.5/24));document.dispose();
});

test('1904 날짜 수식은 셀 참조·배열·필터·저장·Undo에서 native 날짜와 일치한다', async t => {
  const {updateAutoFilterDateColumn}=require('../src/extension'),{parseDateGroups}=require('../src/auto-filter');const bytes=dateFilterXlsx(true),directory=await fs.mkdtemp(path.join(os.tmpdir(),'jinsheet-date-epoch-'));t.after(()=>fs.rm(directory,{recursive:true,force:true}));const file=path.join(directory,'epoch.xlsx');await fs.writeFile(file,bytes);const document=new JinSheetDocument(uri(file),'xlsx',analyze(bytes),bytes,true,'');let operation;document.onDidChange(change=>operation=change);
  document.applyCellEdit(3,1,'=DATE(2024,2,29)+0.75');assert.equal(getRange(document,{requestId:1,rowStart:3,columnStart:1,rowCount:1,columnCount:1},'ko').rows[0][0].displayValue,'2024-02-29');assert.equal(getRange(document,{requestId:2,rowStart:3,columnStart:1,rowCount:1,columnCount:1}).rows[0][0].cachedValue,43889.75);document.applyCellEdit(5,5,'=YEAR(B4)');assert.equal(getRange(document,{requestId:3,rowStart:5,columnStart:5,rowCount:1,columnCount:1}).rows[0][0].cachedValue,2024);
  document.applyAutoFilter(0,updateAutoFilterDateColumn(document,0,1,parseDateGroups('2024-02'),[],false,autoFilterRange(document,0,0,1)));assert.deepEqual(viewModel(document).hiddenRows,[2]);assert.equal(getRange(document,{requestId:4,rowStart:5,columnStart:3,rowCount:1,columnCount:1}).rows[0][0].cachedValue,40);await document.save();const saved=analyze(await fs.readFile(file));assert.equal(saved.sheets[0].cells.get('B4').formula,'DATE(2024,2,29)+0.75');assert.equal(saved.date1904,true);document.dispose();
  const reopened=new JinSheetDocument(uri(file),'xlsx',saved,Buffer.from(saved.bytes),true,'');assert.equal(getRange(reopened,{requestId:5,rowStart:3,columnStart:1,rowCount:1,columnCount:1}).rows[0][0].cachedValue,43889.75);reopened.applyCellEdit(6,5,'=SEQUENCE(2,1,DATE(1904,1,1),1)');assert.deepEqual(getRange(reopened,{requestId:6,rowStart:6,columnStart:5,rowCount:2,columnCount:1}).rows.map(row=>row[0].cachedValue),[0,1]);reopened.dispose();
  const original=new JinSheetDocument(uri('/tmp/epoch-undo.xlsx'),'xlsx',analyze(bytes),bytes,true,'');let change;original.onDidChange(edit=>change=edit);original.applyCellEdit(3,1,'=DATE(2024,3,1)');assert.equal(getRange(original,{requestId:7,rowStart:3,columnStart:1,rowCount:1,columnCount:1}).rows[0][0].cachedValue,43890);change.undo();assert.equal(original.getCell(3,1).formula,'43889.75');change.redo();assert.equal(getRange(original,{requestId:8,rowStart:3,columnStart:1,rowCount:1,columnCount:1}).rows[0][0].cachedValue,43890);original.dispose();
});

test('1904 날짜 유효성과 조건부 서식은 DATE·TEXT 기준 수식을 같은 체계로 평가한다', () => {
  const files=unzipSync(dateFilterXlsx(true)),decode=part=>new TextDecoder().decode(files[part]),encode=(part,text)=>files[part]=new TextEncoder().encode(text);
  encode('xl/worksheets/sheet1.xml',decode('xl/worksheets/sheet1.xml').replace('</worksheet>','<conditionalFormatting sqref="B2:B4"><cfRule type="expression" priority="1" dxfId="0"><formula>B2=DATE(2024,2,29)</formula></cfRule></conditionalFormatting><dataValidations count="1"><dataValidation type="date" operator="between" sqref="B2"><formula1>DATE(2024,2,29)</formula1><formula2>DATE(2024,3,1)</formula2></dataValidation></dataValidations></worksheet>'));
  const styles=decode('xl/styles.xml'),dxfs='<dxfs count="1"><dxf><font><b/></font></dxf></dxfs>';encode('xl/styles.xml',/<dxfs\b/.test(styles)?styles.replace(/<dxfs\b[^>]*(?:\/>|>[\s\S]*?<\/dxfs>)/,dxfs):styles.replace('</styleSheet>',dxfs+'</styleSheet>'));const bytes=zipSync(files),model=analyze(bytes);assert.equal(model.writable,true,model.readOnlyReason);const document=new JinSheetDocument(uri('/tmp/date-epoch-validation.xlsx'),'xlsx',model,bytes,true,'');document.applyCellEdit(1,1,'=DATE(2024,2,29)');let cell=getRange(document,{requestId:1,rowStart:1,columnStart:1,rowCount:1,columnCount:1}).rows[0][0];assert.equal(cell.cachedValue,43889);assert.equal(cell.style.bold,true);const before=document.captureSnapshot();assert.throws(()=>document.applyCellEdit(1,1,'=DATE(2024,3,2)'),/유효성/);assert.equal(document.getCell(1,1).formula,'DATE(2024,2,29)');assert.equal(document.state,before.state);document.dispose();
});

test('1904 INDIRECT 유효성 목록은 날짜 함수로 원본 범위를 선택한다', () => {
  const files=unzipSync(twoSheetValidationXlsx());files['xl/workbook.xml']=new TextEncoder().encode(new TextDecoder().decode(files['xl/workbook.xml']).replace('<sheets>','<workbookPr date1904="1"/><sheets>'));files['xl/worksheets/sheet1.xml']=new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('Lists!$A$1:$A$2','INDIRECT("Lists!A"&amp;MONTH(0))'));const bytes=zipSync(files),model=analyze(bytes);assert.equal(model.writable,true,model.readOnlyReason);const document=new JinSheetDocument(uri('/tmp/date-epoch-list.xlsx'),'xlsx',model,bytes,true,'');document.applyCellEdit(0,1,'Red');assert.equal(document.getCell(0,1).value,'Red');assert.throws(()=>document.applyCellEdit(0,1,'Blue'),/허용 목록/);document.dispose();
});

test('동적 오늘 필터는 시각 변경·TODAY·SUBTOTAL·저장과 Undo를 같은 날짜로 처리한다', async t => {
  const {updateAutoFilterDynamicColumn}=require('../src/extension'),directory=await fs.mkdtemp(path.join(os.tmpdir(),'jinsheet-relative-date-'));t.after(()=>fs.rm(directory,{recursive:true,force:true}));
  for(const epoch of [false,true]){const bytes=dateFilterXlsx(epoch),file=path.join(directory,epoch?'1904.xlsx':'1900.xlsx');await fs.writeFile(file,bytes);const document=new JinSheetDocument(uri(file),'xlsx',analyze(bytes),bytes,true,'');document.now=new Date(2024,1,29,23,59);let operation,count=0;document.onDidChange(change=>{operation=change;count++});document.applyAutoFilter(0,updateAutoFilterDynamicColumn(document,0,1,{type:'today',value:0,maxValue:1},autoFilterRange(document,0,0,1)));const applied=operation;assert.equal(count,1);assert.deepEqual(viewModel(document).hiddenRows,[2]);assert.equal(viewModel(document).hasDynamicDateFilters,true);assert.equal(getRange(document,{requestId:1,rowStart:5,columnStart:3,rowCount:1,columnCount:1}).rows[0][0].cachedValue,40);await document.save();let saved=analyze(await fs.readFile(file));assert.deepEqual(saved.sheets[0].tables[0].autoFilter.columns[0].dynamic,{type:'today',value:45351-(epoch?1462:0),maxValue:45352-(epoch?1462:0)});assert.equal(document.dirty,false);
    document.now=new Date(2024,2,1,0,1);assert.deepEqual(viewModel(document).hiddenRows,[1,3]);assert.equal(document.dirty,false);assert.equal(count,1);assert.equal(getRange(document,{requestId:2,rowStart:5,columnStart:3,rowCount:1,columnCount:1}).rows[0][0].cachedValue,20);await document.save();saved=analyze(await fs.readFile(file));assert.equal(saved.sheets[0].tables[0].autoFilter.columns[0].dynamic.value,45352-(epoch?1462:0));assert.deepEqual(saved.sheets[0].nativeHiddenRows,[1,3]);applied.undo();assert.equal(viewModel(document).hiddenRows,undefined);applied.redo();assert.deepEqual(viewModel(document).hiddenRows,[1,3]);assert.equal(autoFilterOptions(document,0,1,autoFilterRange(document,0,0,1)).dynamic.type,'today');document.dispose();}
});

test('동적 평균과 매년 월 조건은 다른 열 조건·수동 숨김·구조 변경과 함께 유지된다', () => {
  const {updateAutoFilterDynamicColumn}=require('../src/extension'),bytes=calculatedTableXlsx({native:true}),document=new JinSheetDocument(uri('/tmp/dynamic-average.xlsx'),'xlsx',analyze(bytes),bytes,true,'');document.applyAutoFilter(0,updateAutoFilterDynamicColumn(document,0,1,{type:'aboveAverage'},autoFilterRange(document,0,0,1)));assert.deepEqual(viewModel(document).hiddenRows,[1,2]);assert.equal(document.state.sheets[0].tables[0].autoFilter.columns[0].dynamic.value,20);assert.equal(viewModel(document).hasDynamicDateFilters,false);document.applyStructureChange({sheetIndex:0,axis:'row',action:'insert',index:2});assert.equal(document.state.sheets[0].tables[0].autoFilter.columns[0].dynamic.type,'aboveAverage');assert.equal(document.state.sheets[0].tables[0].autoFilter.columns[0].dynamic.value,20);document.dispose();
  const dates=dateFilterXlsx(),months=new JinSheetDocument(uri('/tmp/dynamic-month.xlsx'),'xlsx',analyze(dates),dates,true,'');months.now=new Date(2026,9,16);months.applyAutoFilter(0,updateAutoFilterDynamicColumn(months,0,1,{type:'M2'},autoFilterRange(months,0,0,1)));assert.deepEqual(viewModel(months).hiddenRows,[2]);assert.equal(viewModel(months).hasDynamicDateFilters,false);months.applyAutoFilter(0,updateAutoFilterColumn(months,0,0,['item0'],false,false,autoFilterRange(months,0,0,0)));assert.deepEqual(viewModel(months).hiddenRows,[2,3]);months.dispose();
});

test('계산기의 시각 스냅샷은 TODAY와 동적 필터 사이에 자정 경합을 만들지 않는다', () => {
  const {createFormulaCalculator}=require('../src/formula'),{updateAutoFilterDynamicColumn}=require('../src/extension'),bytes=dateFilterXlsx(),document=new JinSheetDocument(uri('/tmp/dynamic-clock.xlsx'),'xlsx',analyze(bytes),bytes,true,'');document.now=new Date(2024,1,29,23,59);document.applyCellEdit(6,5,'=TODAY()');document.applyAutoFilter(0,updateAutoFilterDynamicColumn(document,0,1,{type:'today'},autoFilterRange(document,0,0,1)));const calculator=createFormulaCalculator(document);document.now.setDate(document.now.getDate()+1);assert.equal(calculator.value(0,6,5),45351);assert.equal(calculator.value(0,5,3),40);const next=createFormulaCalculator(document);assert.equal(next.value(0,6,5),45352);assert.equal(next.value(0,5,3),20);document.dispose();
});

test('동적 필터 UI는 설정 복원·명시적 적용·오류·열별 해제를 처리한다', () => {
  const {elements,messages,dispatch}=runGridWebview(renderHtml(),{kind:'xlsx',rowCount:5,columnCount:2,autoFilter:{top:0,bottom:4,left:0,right:1,columns:[]}});elements.get('filterValuesButton').onclick();const request=messages.at(-1);dispatch({type:'filter-options',sheetIndex:0,column:request.column,options:{values:['1'],selectedValues:[],blanks:false,includeBlanks:false,dynamic:{type:'thisMonth',value:0,maxValue:1}}});assert.equal(elements.get('filterMode').value,'dynamic');assert.equal(elements.get('filterDynamicType').value,'thisMonth');assert.equal(elements.get('filterDynamicControls').hidden,false);elements.get('filterDynamicType').value='nextMonth';assert.equal(messages.some(message=>message.type==='auto-filter'),false);elements.get('filterApply').onclick();const submitted=messages.at(-1);assert.equal(submitted.dynamic.type,'nextMonth');dispatch({type:'filter-error',sheetIndex:0,column:request.column,requestId:submitted.requestId,message:'retry'});assert.equal(elements.get('filterDynamicType').value,'nextMonth');assert.equal(elements.get('filterEditor').disabled,false);elements.get('filterClear').onclick();assert.equal(messages.at(-1).clear,true);assert.equal(messages.at(-1).dynamic,undefined);
});

test('날짜가 바뀌면 안전한 시점에만 동적 필터를 갱신하고 편집 중에는 입력을 유지한다', t => {
  const initial=new Date(2024,1,29,23,59,50).getTime();t.mock.timers.enable({apis:['Date','setTimeout'],now:initial});const grid=runGridWebview(renderHtml(),{kind:'xlsx',hasDynamicDateFilters:true});grid.elements.get('filterPanel').hidden=true;const formula=grid.elements.get('formula');formula.value='unfinished';formula.listeners.get('input')();t.mock.timers.tick(11000);assert.equal(grid.messages.some(message=>message.type==='refresh-dynamic-filters'),false);assert.equal(formula.value,'unfinished');formula.listeners.get('keydown')({key:'Escape',preventDefault(){}});grid.elements.get('fontFamily').tagName='INPUT';grid.elements.get('fontFamily').value='unfinished font';grid.elements.get('fontFamily').focus();t.mock.timers.tick(60000);assert.equal(grid.messages.some(message=>message.type==='refresh-dynamic-filters'),false);assert.equal(grid.elements.get('fontFamily').value,'unfinished font');grid.document.body.focus();t.mock.timers.tick(60000);assert.equal(grid.messages.filter(message=>message.type==='refresh-dynamic-filters').length,1);grid.dispatch({type:'model',kind:'xlsx',readOnly:false,rowCount:2,columnCount:2,sheetIndex:0,hasDynamicDateFilters:false});t.mock.timers.tick(86400000);assert.equal(grid.messages.filter(message=>message.type==='refresh-dynamic-filters').length,1);
});

test('수정 없는 상대 날짜 문서의 저장·다른 이름 저장·백업은 새 기준값을 내보내고 읽기 전용은 보존한다', async t => {
  const {updateAutoFilterDynamicColumn}=require('../src/extension'),directory=await fs.mkdtemp(path.join(os.tmpdir(),'jinsheet-dynamic-export-'));t.after(()=>fs.rm(directory,{recursive:true,force:true}));const bytes=dateFilterXlsx(),file=path.join(directory,'source.xlsx');await fs.writeFile(file,bytes);const document=new JinSheetDocument(uri(file),'xlsx',analyze(bytes),bytes,true,'');document.now=new Date(2024,1,29);document.applyAutoFilter(0,updateAutoFilterDynamicColumn(document,0,1,{type:'today'},autoFilterRange(document,0,0,1)));await document.save();const revision=document.savedRevision;await document.save();assert.equal(document.savedRevision,revision);const original=await fs.readFile(file);document.now=new Date(2024,2,1);const backup=await document.backup({destination:uri(path.join(directory,'backup.xlsx'))});assert.equal(analyze(await fs.readFile(backup.id)).sheets[0].tables[0].autoFilter.columns[0].dynamic.value,45352);assert.equal((await fs.readFile(file)).equals(original),true);await backup.delete();assert.equal(document.dirty,false);
  const copy=path.join(directory,'copy.xlsx');await document.saveAs(uri(copy));assert.equal(analyze(await fs.readFile(copy)).sheets[0].tables[0].autoFilter.columns[0].dynamic.value,45352);assert.equal((await fs.readFile(file)).equals(original),true);document.dispose();
  const readonly=new JinSheetDocument(uri(file),'xlsx',analyze(original),original,false,'readonly');readonly.now=new Date(2024,2,1);assert.equal(readonly.hasRelativeDateFilters(),false);const preserved=path.join(directory,'readonly.xlsx');await readonly.saveAs(uri(preserved));assert.equal((await fs.readFile(preserved)).equals(original),true);readonly.dispose();
});

function coloredTableXlsx(){const {patchCells}=require('../src/xlsx');return patchCells(analyze(calculatedTableXlsx({native:true})),[{sheetIndex:0,address:'B2',value:'10',fillColor:'#FF0000',font:{bold:true,italic:false,color:'#00FF00'}},{sheetIndex:0,address:'B3',value:'20',fillColor:'#0000FF',font:{bold:false,italic:false,color:'#00FF00'}},{sheetIndex:0,address:'B4',value:'30',fillColor:'#FF0000',font:{bold:false,italic:true,color:'#0000FF'}},{sheetIndex:0,address:'E6',value:'0',formula:'SUBTOTAL(109,Sales[Result])'}]).bytes}

test('색상 필터는 셀 서식을 유지하며 native 정의·수동 숨김·SUBTOTAL·Undo·저장을 연결한다', async t => {
  const {updateAutoFilterColorColumn}=require('../src/extension'),bytes=coloredTableXlsx(),directory=await fs.mkdtemp(path.join(os.tmpdir(),'jinsheet-color-filter-'));t.after(()=>fs.rm(directory,{recursive:true,force:true}));const file=path.join(directory,'color.xlsx');await fs.writeFile(file,bytes);const document=new JinSheetDocument(uri(file),'xlsx',analyze(bytes),bytes,true,'');let operation;document.onDidChange(edit=>operation=edit);const cells=['B2','B3','B4'].map(address=>document.state.sheets[0].cells.get(address));document.applyHiddenRows(0,[3],true);const range=autoFilterRange(document,0,0,1),options=autoFilterOptions(document,0,1,range);assert.deepEqual(options.colors.fills,[{color:'#0000FF',count:1},{color:'#FF0000',count:2}]);document.applyAutoFilter(0,updateAutoFilterColorColumn(document,0,1,{cellColor:true,color:'#ff0000'},range));const applied=operation;assert.deepEqual(viewModel(document).hiddenRows,[2,3]);assert.deepEqual(getRange(document,{requestId:1,rowStart:5,columnStart:3,rowCount:1,columnCount:2}).rows[0].map(cell=>cell.cachedValue),[80,20]);assert.deepEqual(['B2','B3','B4'].map(address=>document.state.sheets[0].cells.get(address)),cells);await document.save();let saved=analyze(await fs.readFile(file));assert.equal(saved.writable,true,saved.readOnlyReason);const criterion=saved.sheets[0].tables[0].autoFilter.columns[0].colorFilter;assert.equal(criterion.color,'#FF0000');assert.equal(criterion.cellColor,true);assert.equal(Number.isInteger(criterion.dxfId),true);assert.deepEqual(saved.sheets[0].nativeHiddenRows,[2,3]);applied.undo();assert.deepEqual(viewModel(document).hiddenRows,[3]);applied.redo();assert.deepEqual(viewModel(document).hiddenRows,[2,3]);await document.save();document.dispose();
  const reopened=new JinSheetDocument(uri(file),'xlsx',saved,Buffer.from(saved.bytes),true,'');reopened.applyAutoFilter(0,updateAutoFilterColumn(reopened,0,1,[],false,true,autoFilterRange(reopened,0,0,1)));assert.deepEqual(viewModel(reopened).hiddenRows,[3]);reopened.dispose();
});

test('글자색·채우기 없음 필터와 색상 draft 변경은 표시와 수식에 같은 결과를 준다', () => {
  const {updateAutoFilterColorColumn,buildFillFormatEdits}=require('../src/extension'),bytes=coloredTableXlsx(),document=new JinSheetDocument(uri('/tmp/font-filter.xlsx'),'xlsx',analyze(bytes),bytes,true,'');document.applyAutoFilter(0,updateAutoFilterColorColumn(document,0,1,{cellColor:false,color:'#00FF00'},autoFilterRange(document,0,0,1)));assert.deepEqual(viewModel(document).hiddenRows,[3]);assert.equal(getRange(document,{requestId:1,rowStart:5,columnStart:3,rowCount:1,columnCount:1}).rows[0][0].cachedValue,60);document.applyAutoFilter(0,updateAutoFilterColorColumn(document,0,1,{cellColor:true,color:null},autoFilterRange(document,0,0,1)));assert.deepEqual(viewModel(document).hiddenRows,[1,2,3]);document.applyCellEdits(buildFillFormatEdits(document,{startRow:1,endRow:1,startColumn:1,endColumn:1},null,0));assert.deepEqual(viewModel(document).hiddenRows,[2,3]);assert.equal(getRange(document,{requestId:2,rowStart:5,columnStart:3,rowCount:1,columnCount:1}).rows[0][0].cachedValue,20);const saved=analyze(document.createOutput());assert.deepEqual(saved.sheets[0].nativeHiddenRows,[2,3]);assert.equal(saved.sheets[0].tables[0].autoFilter.columns[0].colorFilter.color,null);document.dispose();
});

test('조건부 서식 색상 필터는 데이터 변경에 따라 표시와 SUBTOTAL을 갱신한다', () => {
  const {updateAutoFilterColorColumn}=require('../src/extension');const files=unzipSync(coloredTableXlsx()),decode=part=>new TextDecoder().decode(files[part]);files['xl/styles.xml']=new TextEncoder().encode(decode('xl/styles.xml').replace('</styleSheet>','<dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs></styleSheet>'));files['xl/worksheets/sheet1.xml']=new TextEncoder().encode(decode('xl/worksheets/sheet1.xml').replace('</worksheet>','<conditionalFormatting sqref="B2:B4"><cfRule type="cellIs" operator="greaterThan" dxfId="0" priority="1"><formula>15</formula></cfRule></conditionalFormatting></worksheet>'));const bytes=zipSync(files),document=new JinSheetDocument(uri('/tmp/cf-color-filter.xlsx'),'xlsx',analyze(bytes),bytes,true,'');document.applyAutoFilter(0,updateAutoFilterColorColumn(document,0,1,{cellColor:true,color:'#FF0000'},autoFilterRange(document,0,0,1)));assert.equal(viewModel(document).hiddenRows,undefined);assert.equal(getRange(document,{requestId:1,rowStart:5,columnStart:3,rowCount:1,columnCount:1}).rows[0][0].cachedValue,120);document.applyCellEdit(2,1,'5');assert.deepEqual(viewModel(document).hiddenRows,[2]);assert.equal(getRange(document,{requestId:2,rowStart:5,columnStart:3,rowCount:1,columnCount:1}).rows[0][0].cachedValue,80);assert.equal(getRange(document,{requestId:3,rowStart:2,columnStart:1,rowCount:1,columnCount:1}).rows[0][0].style.fill,'#0000FF');document.dispose();
});

test('색상 척도 필터는 새 draft 셀을 통계에 포함하고 화면 색·SUBTOTAL과 일치한다', () => {
  const {updateAutoFilterColorColumn}=require('../src/extension'),files=unzipSync(coloredTableXlsx());files['xl/worksheets/sheet1.xml']=new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</worksheet>','<conditionalFormatting sqref="B2:B5"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FF0000FF"/></colorScale></cfRule></conditionalFormatting></worksheet>'));const bytes=zipSync(files),document=new JinSheetDocument(uri('/tmp/scale-filter.xlsx'),'xlsx',analyze(bytes),bytes,true,'');const options=autoFilterOptions(document,0,1,autoFilterRange(document,0,0,1));assert.ok(options.colors.fills.some(item=>item.color==='#800080'));document.applyAutoFilter(0,updateAutoFilterColorColumn(document,0,1,{cellColor:true,color:'#800080'},autoFilterRange(document,0,0,1)));assert.deepEqual(viewModel(document).hiddenRows,[1,3]);assert.equal(getRange(document,{requestId:1,rowStart:5,columnStart:3,rowCount:1,columnCount:1}).rows[0][0].cachedValue,40);document.applyCellEdit(4,1,'50',0,new Set(),{skipExpansion:true});assert.equal(document.state.sheets[0].cells.has('B5'),false);assert.deepEqual(viewModel(document).hiddenRows,[1,2]);assert.equal(getRange(document,{requestId:2,rowStart:5,columnStart:3,rowCount:1,columnCount:1}).rows[0][0].cachedValue,60);assert.equal(getRange(document,{requestId:3,rowStart:2,columnStart:1,rowCount:1,columnCount:1}).rows[0][0].style.fill,'#BF0040');document.dispose();
});

test('색상 필터 UI는 색 목록·미리보기·직접 선택·오류를 처리하고 입력 중 적용하지 않는다', () => {
  const {elements,messages,dispatch}=runGridWebview(renderHtml(),{kind:'xlsx',rowCount:4,columnCount:2,autoFilter:{top:0,bottom:3,left:0,right:1,columns:[]}});elements.get('filterValuesButton').onclick();const request=messages.at(-1);dispatch({type:'filter-options',sheetIndex:0,column:request.column,options:{values:['1'],selectedValues:['1'],blanks:false,includeBlanks:false,colorFilter:{cellColor:false,color:'#00FF00',dxfId:2},colors:{fills:[{color:null,count:1},{color:'#FF0000',count:2}],fonts:[{color:'#00FF00',count:3}],fillCount:2,fontCount:1}}});assert.equal(elements.get('filterMode').value,'color');assert.equal(elements.get('filterColorTarget').value,'font');assert.equal(elements.get('filterColorValue').value,'#00FF00');assert.equal(elements.get('filterColorPreview').style.color,'#00FF00');elements.get('filterColorTarget').value='fill';elements.get('filterColorTarget').onchange();elements.get('filterColorValue').value='custom';elements.get('filterColorValue').onchange();elements.get('filterColorCustom').value='#123456';elements.get('filterColorCustom').oninput();assert.equal(elements.get('filterColorPreview').style.backgroundColor,'#123456');assert.equal(messages.some(message=>message.type==='auto-filter'),false);elements.get('filterApply').onclick();const submitted=messages.at(-1);assert.deepEqual(JSON.parse(JSON.stringify(submitted.colorFilter)),{cellColor:true,color:'#123456'});dispatch({type:'filter-error',sheetIndex:0,column:request.column,requestId:submitted.requestId,message:'retry'});assert.equal(elements.get('filterColorValue').value,'custom');assert.equal(elements.get('filterColorCustom').value,'#123456');elements.get('filterClear').onclick();assert.equal(messages.at(-1).clear,true);assert.equal(messages.at(-1).colorFilter,undefined);
});

test('아이콘 필터 UI는 현재 열의 조건부 아이콘을 선택해 적용하고 실패 시 다시 고를 수 있다', () => {
  const { elements, messages, dispatch } = runGridWebview(renderHtml(), { kind: 'xlsx', rowCount: 4, columnCount: 1, autoFilter: { top: 0, bottom: 3, left: 0, right: 0, columns: [] } });
  elements.get('filterValuesButton').onclick();
  const request = messages.at(-1);
  dispatch({ type: 'filter-options', sheetIndex: 0, column: request.column, options: { values: ['1', '2'], selectedValues: ['1', '2'], blanks: false, includeBlanks: false, iconFilters: [{ iconSet: '3Arrows', iconId: 0, count: 1 }, { iconSet: '3Arrows', iconId: 2, count: 1 }] } });
  elements.get('filterMode').value = 'icons'; elements.get('filterMode').onchange();
  assert.equal(elements.get('filterIconControls').hidden, false);
  elements.get('filterApply').onclick();
  let submitted = messages.at(-1);
  assert.deepEqual(JSON.parse(JSON.stringify(submitted.iconFilter)), { iconSet: '3Arrows', iconId: 0 });
  dispatch({ type: 'filter-error', sheetIndex: 0, column: request.column, requestId: submitted.requestId, message: 'retry' });
  assert.equal(elements.get('filterEditor').disabled, false);
  elements.get('filterIconValue').value = '3Arrows:2';
  elements.get('filterApply').onclick(); submitted = messages.at(-1);
  assert.deepEqual(JSON.parse(JSON.stringify(submitted.iconFilter)), { iconSet: '3Arrows', iconId: 2 });
  dispatch({ type: 'filter-applied', sheetIndex: 0, column: request.column, requestId: submitted.requestId });
  elements.get('filterValuesButton').onclick();
  const reopened = messages.at(-1);
  assert.equal(reopened.type, 'filter-options');
  dispatch({ type: 'filter-options', sheetIndex: 0, column: request.column, options: { values: ['1', '2'], selectedValues: ['1'], blanks: false, includeBlanks: false, iconFilters: [{ iconSet: '3Arrows', iconId: 0, count: 1 }, { iconSet: '3Arrows', iconId: 2, count: 1 }], iconFilter: { iconSet: '3Arrows', iconId: 2 } } });
  elements.get('filterClear').onclick();
  assert.equal(messages.at(-1).clear, true);
});

test('동적 배열 결과의 기존 셀 색은 표시와 색상 필터·SUBTOTAL에서 유지된다', () => {
  const {patchCells}=require('../src/xlsx'),{updateAutoFilterColorColumn}=require('../src/extension');const rows='<row r="1"><c r="A1" t="inlineStr"><is><t>Values</t></is></c></row><row r="2"><c r="A2"><f>SEQUENCE(3,1,10,10)</f><v>10</v></c></row><row r="6"><c r="B6"><f>SUBTOTAL(9,A2:A4)</f><v>0</v></c></row>';const model=patchCells(analyze(minimalXlsx('',rows)),[{sheetIndex:0,address:'A2',value:'10',formula:'SEQUENCE(3,1,10,10)',fillColor:'#FF0000'},{sheetIndex:0,address:'A3',value:'',kind:'blank',fillColor:'#0000FF'},{sheetIndex:0,address:'A4',value:'',kind:'blank',fillColor:'#FF0000'}]).model;const document=new JinSheetDocument(uri('/tmp/spill-color-filter.xlsx'),'xlsx',model,Buffer.from(model.bytes),true,'');const base={top:0,bottom:3,left:0,right:0,columns:[]};document.applyAutoFilter(0,updateAutoFilterColorColumn(document,0,0,{cellColor:true,color:'#FF0000'},base));assert.deepEqual(viewModel(document).hiddenRows,[2]);const cell=getRange(document,{requestId:1,rowStart:2,columnStart:0,rowCount:1,columnCount:1}).rows[0][0];assert.equal(cell.isSpill,true);assert.equal(cell.style.fill,'#0000FF');assert.equal(getRange(document,{requestId:2,rowStart:5,columnStart:1,rowCount:1,columnCount:1}).rows[0][0].cachedValue,40);document.dispose();
});

function filteredAppendXlsx() {
  const files=unzipSync(calculatedTableXlsx({native:true}));files['xl/tables/table1.xml']=new TextEncoder().encode(new TextDecoder().decode(files['xl/tables/table1.xml']).replace('<autoFilter ref="A1:C4"/>','<autoFilter ref="A1:C4"><filterColumn colId="1"><customFilters><customFilter operator="greaterThan" val="0"/></customFilters></filterColumn></autoFilter>'));return zipSync(files);
}

test('필터에 가려지는 표 행은 선택 전 원본을 유지하고 승인 시 필터·입력·행을 한 번에 변경한다',()=>{
  const bytes=filteredAppendXlsx(),document=new JinSheetDocument(uri('/tmp/append-filter-choice.xlsx'),'xlsx',analyze(bytes),bytes,true,'');let operation,count=0,pending;document.onDidChange(change=>{operation=change;count++});
  try{document.appendTableRow(0,'Sales',3,2,'70');assert.fail('expected filter choice')}catch(error){pending=error}
  assert.equal(pending.tableFilterPrompt.tableName,'Sales');assert.equal(typeof pending.retry,'function');assert.equal(document.dirty,false);assert.equal(count,0);assert.equal(document.getCell(3,2).value,'60');
  assert.deepEqual(pending.retry(),{row:4,column:0});assert.equal(count,1);assert.equal(document.getCell(3,2).value,'70');assert.equal(document.state.sheets[0].tables[0].range.bottom,4);assert.deepEqual(document.state.sheets[0].tables[0].autoFilter.columns,[]);assert.equal(viewModel(document).hiddenRows,undefined);assert.equal(document.getCell(4,2).formula,'B5*2');
  operation.undo();assert.equal(document.getCell(3,2).formula,'B4*2');assert.equal(document.state.sheets[0].tables[0].range.bottom,3);assert.equal(document.state.sheets[0].tables[0].autoFilter.columns.length,1);assert.equal(document.dirty,false);
  operation.redo();const reopened=analyze(document.createOutput());assert.equal(reopened.sheets[0].tables[0].range.bottom,4);assert.deepEqual(reopened.sheets[0].tables[0].autoFilter.columns,[]);assert.equal(reopened.sheets[0].cells.get('C4').value,'70');assert.equal(count,1);assert.throws(()=>pending.retry(),/문서가 변경/);document.dispose();
});

test('표 필터 해제 선택을 기다리는 동안 다른 편집·저장·필터 변경이 있으면 오래된 재시도를 거부한다',async()=>{
  for(const action of ['edit','filter','save']){
    const bytes=filteredAppendXlsx(),directory=await fs.mkdtemp(path.join(os.tmpdir(),'jinsheet-append-choice-')),file=path.join(directory,'choice.xlsx');await fs.writeFile(file,bytes);const document=new JinSheetDocument(uri(file),'xlsx',analyze(bytes),bytes,true,'');let pending;
    if(action==='save')document.applyCellEdit(1,0,'draft before prompt');
    try{document.appendTableRow(0,'Sales',3,2,'70')}catch(error){pending=error}
    if(action==='edit')document.applyCellEdit(1,0,'keep');else if(action==='filter')document.applyAutoFilter(0,{...document.state.sheets[0].tables[0].autoFilter,columns:[]});else await document.save();
    assert.throws(()=>pending.retry(),/문서가 변경/,action);assert.equal(document.state.sheets[0].tables[0].range.bottom,3);assert.equal(document.getCell(3,2).value,'60');if(action==='edit')assert.equal(document.getCell(1,0).value,'keep');document.dispose();await fs.rm(directory,{recursive:true,force:true});
  }
});

test('표 행 선택은 데이터 유효성 경고 뒤에도 전체 변경을 원자적으로 유지한다',()=>{
  const files=unzipSync(filteredAppendXlsx());files['xl/worksheets/sheet1.xml']=new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</worksheet>','<dataValidations count="1"><dataValidation type="whole" operator="lessThan" errorStyle="warning" sqref="C4"><formula1>100</formula1></dataValidation></dataValidations></worksheet>'));const bytes=zipSync(files),document=new JinSheetDocument(uri('/tmp/append-choice-validation.xlsx'),'xlsx',analyze(bytes),bytes,true,'');let pending,count=0;document.onDidChange(()=>count++);
  try{document.appendTableRow(0,'Sales',3,2,'200')}catch(error){pending=error}assert.equal(pending.validationPrompt.style,'warning');
  try{pending.retry();assert.fail('expected filter choice')}catch(error){pending=error}assert.equal(pending.tableFilterPrompt.tableName,'Sales');assert.equal(count,0);assert.equal(document.getCell(3,2).value,'60');assert.deepEqual(pending.retry(),{row:4,column:0});assert.equal(count,1);assert.equal(document.getCell(3,2).value,'200');document.dispose();
});

test('수동 숨김 새 행은 필터 해제 선택을 제안하지 않고 수동 숨김을 유지한다',()=>{
  const {patchHiddenRows}=require('../src/xlsx'),original=analyze(filteredAppendXlsx()),model=patchHiddenRows(original,[{sheetIndex:0,rows:[4],hidden:true}]).model,document=new JinSheetDocument(uri('/tmp/append-manual-hidden.xlsx'),'xlsx',model,Buffer.from(model.bytes),true,'');let pending;
  try{document.appendTableRow(0,'Sales',3,2,'70')}catch(error){pending=error}assert.match(pending.message,/수동 숨김/);assert.equal(pending.tableFilterPrompt,undefined);assert.equal(pending.retry,undefined);assert.equal(document.dirty,false);assert.deepEqual(document.state.sheets[0].hiddenRows,[4]);assert.equal(document.state.sheets[0].tables[0].autoFilter.columns.length,1);document.dispose();
});

test('표 필터 해제 후 행 추가는 기존 수동 숨김을 유지하고 일치하는 새 행에는 선택을 요구하지 않는다',()=>{
  const {patchHiddenRows}=require('../src/xlsx'),model=patchHiddenRows(analyze(filteredAppendXlsx()),[{sheetIndex:0,rows:[1],hidden:true}]).model,document=new JinSheetDocument(uri('/tmp/append-keep-manual.xlsx'),'xlsx',model,Buffer.from(model.bytes),true,'');let pending;
  try{document.appendTableRow(0,'Sales',3,2)}catch(error){pending=error}assert.deepEqual(pending.retry(),{row:4,column:0});assert.deepEqual(document.state.sheets[0].hiddenRows,[1]);assert.deepEqual(viewModel(document).hiddenRows,[1]);document.dispose();
  const bytes=filteredAppendXlsx(),matching=new JinSheetDocument(uri('/tmp/append-matching.xlsx'),'xlsx',analyze(bytes),bytes,true,'');const filter=matching.state.sheets[0].tables[0].autoFilter;matching.applyAutoFilter(0,{...filter,columns:[{column:1,values:['10','20','30'],blanks:true}]});assert.deepEqual(matching.appendTableRow(0,'Sales',3,2),{row:4,column:0});assert.equal(matching.state.sheets[0].tables[0].autoFilter.columns.length,1);assert.equal(viewModel(matching).hiddenRows,undefined);matching.dispose();
});

test('다른 필터가 새 표 행을 숨기면 대상 표의 조건도 임의로 해제하지 않는다',()=>{
  const {patchAutoFilters}=require('../src/xlsx'),model=patchAutoFilters(analyze(filteredAppendXlsx()),[{sheetIndex:0,filter:{top:0,bottom:5,left:4,right:4,columns:[{column:4,values:['keep'],blanks:false}]}}]).model,document=new JinSheetDocument(uri('/tmp/append-other-filter.xlsx'),'xlsx',model,Buffer.from(model.bytes),true,'');let pending;
  try{document.appendTableRow(0,'Sales',3,2,'70')}catch(error){pending=error}assert.match(pending.message,/다른 필터/);assert.equal(pending.tableFilterPrompt,undefined);assert.equal(document.dirty,false);assert.equal(document.state.sheets[0].autoFilter.columns.length,1);assert.equal(document.state.sheets[0].tables[0].autoFilter.columns.length,1);assert.equal(document.getCell(3,2).value,'60');document.dispose();
});

function emptyTableXlsx({ single=false, filter=false }={}) {
  const files=unzipSync(calculatedTableXlsx({native:true})),encode=text=>new TextEncoder().encode(text),decode=bytes=>new TextDecoder().decode(bytes);
  files['xl/worksheets/sheet1.xml']=encode(decode(files['xl/worksheets/sheet1.xml']).replace(/<row r="[234]">[\s\S]*?<\/row>/g,''));
  let table=decode(files['xl/tables/table1.xml']).replaceAll('A1:C4',single?'A1':'A1:C1');
  if(single)table=table.replace('<tableColumns count="3">','<tableColumns count="1">').replace(/<tableColumn id="2"[^>]*\/>/,'').replace(/<tableColumn id="3"[\s\S]*?<\/tableColumn>/,'');
  if(filter)table=table.replace(/<autoFilter ref="[^"]*"\/>/,'<autoFilter ref="'+(single?'A1':'A1:C1')+'"><filterColumn colId="0"><customFilters><customFilter operator="notEqual" val=""/></customFilters></filterColumn></autoFilter>');
  files['xl/tables/table1.xml']=encode(table);return zipSync(files);
}

test('머리글뿐인 표는 단일 셀 ref와 빈 필터를 읽고 구조화 집계를 0으로 계산한다',()=>{
  for(const single of [false,true]){
    const bytes=emptyTableXlsx({single}),model=analyze(bytes);assert.equal(model.writable,true,model.readOnlyReason);const document=new JinSheetDocument(uri('/tmp/empty-table.xlsx'),'xlsx',model,bytes,true,'');
    const filter=autoFilterRange(document,0,0,0),options=autoFilterOptions(document,0,0,filter);assert.equal(filter.bottom,0);assert.deepEqual(options.values,[]);assert.equal(options.blanks,false);
    const {createFormulaCalculator}=require('../src/formula');
    if(!single){assert.equal(getRange(document,{requestId:1,rowStart:5,columnStart:3,rowCount:1,columnCount:1}).rows[0][0].cachedValue,0);for(const [formula,value]of [['SUM(Sales[Amount])',0],['COUNT(Sales[Amount])',0],['COUNTA(Sales[Amount])',0],['ROWS(Sales[Amount])',0],['COLUMNS(Sales[Amount])',1],['SUM(Sales[#All])',0],['COUNTBLANK(Sales[Amount])',0]]){document.applyCellEdit(7,4,'='+formula);assert.equal(createFormulaCalculator(document).value(0,7,4),value,formula)}}document.dispose();
  }
});

test('빈 표의 첫 데이터 입력은 계산 열을 채우고 한 번의 Undo로 머리글 표에 돌아간다',()=>{
  const bytes=emptyTableXlsx(),document=new JinSheetDocument(uri('/tmp/empty-table-input.xlsx'),'xlsx',analyze(bytes),bytes,true,'');let change,count=0;document.onDidChange(value=>{change=value;count++});document.applyCellEdit(1,1,'7');assert.equal(count,1);assert.equal(document.state.sheets[0].tables[0].range.bottom,1);assert.equal(document.getCell(1,2).formula,'B2*2');assert.equal(getRange(document,{requestId:1,rowStart:5,columnStart:3,rowCount:1,columnCount:1}).rows[0][0].cachedValue,14);change.undo();assert.equal(document.state.sheets[0].tables[0].range.bottom,0);assert.equal(document.getCell(1,1).value,'');assert.equal(document.state.sheets[0].cells.has('B2'),false);assert.equal(document.dirty,false);change.redo();assert.equal(analyze(document.createOutput()).sheets[0].cells.get('C2').formula,'B2*2');document.dispose();
});

test('빈 표의 마지막 머리글 Tab은 첫 행과 계산 열을 만들고 필터 선택을 재사용한다',()=>{
  for(const filter of [false,true]){
    const bytes=emptyTableXlsx({filter}),document=new JinSheetDocument(uri('/tmp/empty-table-tab.xlsx'),'xlsx',analyze(bytes),bytes,true,'');let count=0;document.onDidChange(()=>count++);
    if(filter){let pending;try{document.appendTableRow(0,'Sales',0,2)}catch(error){pending=error}assert.equal(pending.tableFilterPrompt.tableName,'Sales');assert.equal(document.dirty,false);assert.deepEqual(pending.retry(),{row:1,column:0});assert.deepEqual(document.state.sheets[0].tables[0].autoFilter.columns,[])}else assert.deepEqual(document.appendTableRow(0,'Sales',0,2),{row:1,column:0});
    assert.equal(count,1);assert.equal(document.getCell(1,2).formula,'B2*2');assert.equal(document.state.sheets[0].tables[0].range.bottom,1);document.dispose();
  }
});

test('빈 표의 합계 행은 0을 계산하고 첫 행 추가 시 이동하며 다시 끌 수 있다',()=>{
  const bytes=emptyTableXlsx(),document=new JinSheetDocument(uri('/tmp/empty-table-totals.xlsx'),'xlsx',analyze(bytes),bytes,true,'');document.applyTableTotals(0,'Sales',{enabled:true});assert.equal(document.state.sheets[0].tables[0].totalsRowCount,1);assert.equal(getRange(document,{requestId:1,rowStart:1,columnStart:2,rowCount:1,columnCount:1}).rows[0][0].cachedValue,0);document.applyTableTotals(0,'Sales',{enabled:false});assert.equal(document.state.sheets[0].tables[0].range.bottom,0);document.applyTableTotals(0,'Sales',{enabled:true});assert.deepEqual(document.appendTableRow(0,'Sales',0,2),{row:1,column:0});assert.equal(document.state.sheets[0].tables[0].range.bottom,2);assert.equal(document.getCell(1,2).formula,'B2*2');assert.match(document.getCell(2,2).formula,/SUBTOTAL/);document.dispose();
});

test('표를 머리글 범위로 줄여도 셀 값·계산 열 정의는 유지하고 다시 확대할 수 있다',()=>{
  const bytes=calculatedTableXlsx({native:true}),document=new JinSheetDocument(uri('/tmp/empty-table-resize.xlsx'),'xlsx',analyze(bytes),bytes,true,'');document.applyTableResize(0,'Sales',{top:0,bottom:0,left:0,right:2});assert.equal(document.state.sheets[0].tables[0].range.bottom,0);assert.equal(document.state.sheets[0].tables[0].calculatedColumns[2],'B2*2');assert.equal(document.getCell(1,1).value,'10');assert.equal(getRange(document,{requestId:1,rowStart:5,columnStart:3,rowCount:1,columnCount:1}).rows[0][0].cachedValue,0);document.applyTableResize(0,'Sales',{top:0,bottom:3,left:0,right:2});assert.equal(getRange(document,{requestId:2,rowStart:5,columnStart:3,rowCount:1,columnCount:1}).rows[0][0].cachedValue,120);document.dispose();
});

test('빈 표의 머리글 서식은 첫 데이터 행이나 계산 열에 복사하지 않는다',()=>{
  const {patchCells}=require('../src/xlsx');for(const totals of [false,true]){
    const model=patchCells(analyze(emptyTableXlsx()),[0,1,2].map(column=>({sheetIndex:0,address:String.fromCharCode(65+column)+'1',value:['Name','Amount','Result'][column],font:{bold:true,italic:false},fillColor:'#FF0000'}))).model;
    const document=new JinSheetDocument(uri('/tmp/empty-table-style.xlsx'),'xlsx',model,Buffer.from(model.bytes),true,'');if(totals)document.applyTableTotals(0,'Sales',{enabled:true});document.appendTableRow(0,'Sales',0,2);for(const column of [0,1,2]){const style=document.getCell(1,column)?.styleProperties;assert.equal(style?.bold,undefined);assert.equal(style?.fill,undefined)}assert.equal(document.state.sheets[0].tables[0].calculatedColumns[2],'B2*2');document.dispose();
  }
});

test('빈 표의 조건 필터는 저장·재개방 후 첫 입력에 적용하며 일반 단일 행 필터는 계속 거부한다',()=>{
  const bytes=emptyTableXlsx(),document=new JinSheetDocument(uri('/tmp/empty-table-filter.xlsx'),'xlsx',analyze(bytes),bytes,true,'');const filter=autoFilterRange(document,0,0,1);document.applyAutoFilter(0,updateAutoFilterCustomColumn(document,0,1,{and:false,conditions:[{operator:'greaterThan',value:'0'}]},filter));assert.equal(document.state.sheets[0].tables[0].autoFilter.bottom,0);const saved=analyze(document.createOutput());assert.equal(saved.writable,true,saved.readOnlyReason);assert.equal(saved.sheets[0].tables[0].autoFilter.columns[0].custom.conditions[0].value,'0');document.applyCellEdit(1,1,'7');assert.equal(viewModel(document).hiddenRows,undefined);assert.equal(document.state.sheets[0].tables[0].autoFilter.bottom,1);document.dispose();
  const oneRow=analyze(minimalXlsx('',undefined,'<autoFilter ref="A1:A1"/>'));assert.equal(oneRow.writable,false);assert.match(oneRow.readOnlyReason,/AutoFilter/);
});

test('빈 한 열 표의 UI는 머리글 Tab과 필터 창을 활성화한다',()=>{
  const table={name:'Only',range:{top:0,bottom:0,left:0,right:0},headerRowCount:1,totalsRowCount:0,columns:['Value'],calculatedColumns:[null],totalColumns:[{function:'none'}]},grid=runGridWebview(renderHtml(),{kind:'xlsx',rowCount:1,columnCount:1,tables:[table]});assert.equal(grid.elements.get('filterValuesButton').disabled,false);const cell=grid.document.querySelector('[data-row="0"][data-column="0"]');cell.onclick({target:cell});grid.documentListeners.get('keydown')({target:cell,key:'Tab',preventDefault(){}});assert.equal(grid.messages.at(-1).type,'table-tab');assert.equal(grid.messages.at(-1).tableName,'Only');assert.equal(grid.messages.at(-1).row,0);assert.equal(grid.messages.at(-1).column,0);
});

test('표 필터 유지 후 입력줄에서 Tab을 다시 누르면 일반 편집 없이 행 추가를 재요청한다',()=>{
  const table={name:'Sales',range:{top:0,bottom:3,left:0,right:2},headerRowCount:1,totalsRowCount:0,columns:['Name','Amount','Result'],calculatedColumns:[null,null,'B2*2'],totalColumns:[{function:'none'},{function:'none'},{function:'none'}]},grid=runGridWebview(renderHtml(),{kind:'xlsx',rowCount:6,columnCount:5,tables:[table]});const cell=grid.document.querySelector('[data-row="3"][data-column="2"]');cell.onclick({target:cell});const formula=grid.elements.get('formula');formula.value='=B4*3';formula.listeners.get('input')();const tab=()=>formula.listeners.get('keydown')({target:formula,key:'Tab',preventDefault(){}});tab();const request=grid.messages.at(-1);grid.dispatch({...request,type:'table-tab-error',message:'cancelled'});assert.equal(grid.document.activeElement,formula);assert.equal(formula.value,'=B4*3');tab();const retry=grid.messages.at(-1);assert.equal(retry.type,'table-tab');assert.equal(retry.value,'=B4*3');assert.ok(retry.requestId>request.requestId);assert.equal(grid.messages.some(message=>message.type==='edit'),false);
});

test('표의 마지막 데이터 행 삭제는 빈 표·필터·계산 열을 유지하고 Undo·재입력을 지원한다',()=>{
  const bytes=emptyTableXlsx(),document=new JinSheetDocument(uri('/tmp/table-delete-last.xlsx'),'xlsx',analyze(bytes),bytes,true,'');document.applyCellEdit(1,1,'7');const filter=autoFilterRange(document,0,0,1);document.applyAutoFilter(0,updateAutoFilterCustomColumn(document,0,1,{and:false,conditions:[{operator:'greaterThan',value:'0'}]},filter));let operation,count=0;document.onDidChange(change=>{operation=change;count++});document.applyStructureChange({sheetIndex:0,axis:'row',action:'delete',index:1});assert.equal(count,1);assert.equal(document.state.sheets[0].tables[0].range.bottom,0);assert.equal(document.state.sheets[0].tables[0].autoFilter.bottom,0);assert.equal(document.state.sheets[0].tables[0].autoFilter.columns.length,1);assert.equal(document.state.sheets[0].tables[0].calculatedColumns[2],'B2*2');assert.equal(document.state.sheets[0].cells.has('B2'),false);assert.equal(getRange(document,{requestId:1,rowStart:4,columnStart:3,rowCount:1,columnCount:1}).rows[0][0].cachedValue,0);const deletion=operation;deletion.undo();assert.equal(document.getCell(1,1).value,'7');assert.equal(document.state.sheets[0].tables[0].range.bottom,1);deletion.redo();document.applyCellEdit(1,1,'9');assert.equal(document.getCell(1,2).formula,'B2*2');assert.equal(getRange(document,{requestId:2,rowStart:4,columnStart:3,rowCount:1,columnCount:1}).rows[0][0].cachedValue,18);const reopened=analyze(document.createOutput());assert.equal(reopened.writable,true,reopened.readOnlyReason);assert.equal(reopened.sheets[0].tables[0].range.bottom,1);document.dispose();
});

test('마지막 데이터 행 삭제 후 남은 총계는 빈 범위를 계산하고 첫 행 추가 시 다시 이동한다',()=>{
  const bytes=emptyTableXlsx(),document=new JinSheetDocument(uri('/tmp/table-delete-last-total.xlsx'),'xlsx',analyze(bytes),bytes,true,'');document.applyCellEdit(1,1,'7');document.applyTableTotals(0,'Sales',{enabled:true});document.applyStructureChange({sheetIndex:0,axis:'row',action:'delete',index:1});assert.equal(document.state.sheets[0].tables[0].range.bottom,1);assert.equal(document.state.sheets[0].tables[0].totalsRowCount,1);assert.equal(document.state.sheets[0].tables[0].autoFilter.bottom,0);assert.match(document.getCell(1,2).formula,/SUBTOTAL/);assert.equal(getRange(document,{requestId:1,rowStart:1,columnStart:2,rowCount:1,columnCount:1}).rows[0][0].cachedValue,0);assert.equal(document.state.sheets[0].tables[0].calculatedColumns[2],'B2*2');document.appendTableRow(0,'Sales',0,2);document.applyCellEdit(1,1,'11');assert.equal(document.getCell(1,2).formula,'B2*2');assert.equal(getRange(document,{requestId:2,rowStart:2,columnStart:2,rowCount:1,columnCount:1}).rows[0][0].cachedValue,22);document.dispose();
});

test('빈 표 템플릿은 상대 데이터 참조만 유지하고 절대·외부·머리글·문자열 참조를 구분한다',()=>{
  const bytes=emptyTableXlsx(),document=new JinSheetDocument(uri('/tmp/table-delete-prototype.xlsx'),'xlsx',analyze(bytes),bytes,true,'');document.applyCellEdit(1,1,'7');const {patchCalculatedColumns,patchCells}=require('../src/xlsx');let model=patchCalculatedColumns(analyze(document.createOutput()),[{sheetIndex:0,tableName:'Sales',column:2,formula:'SUM($B2:B2)+B1+B10+$B$2+D2+"B2"+Sales[Amount]'}]).model;model=patchCells(model,[{sheetIndex:0,address:'E6',value:'',formula:'B2'}]).model;document.dispose();const mixed=new JinSheetDocument(uri('/tmp/table-delete-mixed.xlsx'),'xlsx',model,Buffer.from(model.bytes),true,'');mixed.applyStructureChange({sheetIndex:0,axis:'row',action:'delete',index:1});assert.equal(mixed.state.sheets[0].tables[0].calculatedColumns[2],'SUM($B2:B2)+B1+B9+#REF!+#REF!+"B2"+Sales[Amount]');assert.equal(mixed.getCell(4,4).formula,'#REF!');mixed.dispose();
});

test('모든 표 데이터 행을 차례로 삭제해도 머리글과 계산 열이 남고 머리글·총계 삭제는 거부한다',()=>{
  const bytes=calculatedTableXlsx({native:true}),document=new JinSheetDocument(uri('/tmp/table-delete-all.xlsx'),'xlsx',analyze(bytes),bytes,true,'');for(const index of [3,2,1])document.applyStructureChange({sheetIndex:0,axis:'row',action:'delete',index});assert.equal(document.state.sheets[0].tables[0].range.bottom,0);assert.equal(document.state.sheets[0].tables[0].calculatedColumns[2],'B2*2');assert.equal(getRange(document,{requestId:1,rowStart:2,columnStart:3,rowCount:1,columnCount:1}).rows[0][0].cachedValue,0);const before=document.state;assert.throws(()=>document.applyStructureChange({sheetIndex:0,axis:'row',action:'delete',index:0}),/머리글/);assert.equal(document.state,before);document.applyTableTotals(0,'Sales',{enabled:true});const withTotal=document.state;assert.throws(()=>document.applyStructureChange({sheetIndex:0,axis:'row',action:'delete',index:1}),/합계 행/);assert.equal(document.state,withTotal);document.appendTableRow(0,'Sales',0,2);document.applyCellEdit(1,1,'12');assert.equal(document.getCell(1,2).formula,'B2*2');document.dispose();
});

test('표 머리글이 없는 한 행 표의 마지막 행은 잘못된 빈 ref 대신 기존 상태를 보존한다',()=>{
  const files=unzipSync(emptyTableXlsx({single:true}));files['xl/tables/table1.xml']=new TextEncoder().encode(new TextDecoder().decode(files['xl/tables/table1.xml']).replace('<table ','<table headerRowCount="0" ').replace(/<autoFilter[^>]*\/>/,''));const bytes=zipSync(files),model=analyze(bytes);assert.equal(model.writable,true,model.readOnlyReason);const document=new JinSheetDocument(uri('/tmp/table-no-header-last.xlsx'),'xlsx',model,bytes,true,'');assert.throws(()=>document.applyStructureChange({sheetIndex:0,axis:'row',action:'delete',index:0}),/마지막 Excel 표 데이터 행/);assert.equal(document.state,model);assert.equal(document.dirty,false);document.dispose();
});

test('표 스타일은 기본 미리보기와 직접 셀 서식을 유지하고 저장·Undo·Redo에 반영한다',()=>{
  const {patchCells}=require('../src/xlsx'),bytes=minimalXlsx('', '<row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>First</t></is></c><c r="B2"><v>7</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>Second</t></is></c><c r="B3"><v>8</v></c></row>');
  const model=patchCells(analyze(bytes),[{sheetIndex:0,address:'B2',value:'7',fillColor:'#FF0000'}]).model,document=new JinSheetDocument(uri('/tmp/table-style-preview.xlsx'),'xlsx',model,Buffer.from(model.bytes),true,'');
  document.applyCreateTable(0,{top:0,bottom:2,left:0,right:1},'Sales');
  assert.deepEqual(document.state.sheets[0].tables[0].styleInfo,{name:'TableStyleMedium2',showFirstColumn:false,showLastColumn:false,showRowStripes:true,showColumnStripes:false});
  const header=getRange(document,{requestId:1,rowStart:0,columnStart:0,rowCount:1,columnCount:2}).rows[0][0],firstData=getRange(document,{requestId:2,rowStart:1,columnStart:0,rowCount:1,columnCount:2}).rows[0][1],secondData=getRange(document,{requestId:3,rowStart:2,columnStart:0,rowCount:1,columnCount:2}).rows[0][1];
  assert.equal(header.tableStyle.fill,'#4472C4');assert.equal(firstData.style.fill,'#FF0000');assert.equal(firstData.tableStyle.fill,'#D9E2F3');assert.equal(secondData.tableStyle,null);
  let operation;document.onDidChange(change=>operation=change);
  document.applyTableStyle(0,'Sales',{name:'TableStyleLight9',showFirstColumn:true,showLastColumn:true,showRowStripes:false,showColumnStripes:true});
  assert.deepEqual(document.state.sheets[0].tables[0].styleInfo,{name:'TableStyleLight9',showFirstColumn:true,showLastColumn:true,showRowStripes:false,showColumnStripes:true});
  const styled=analyze(document.createOutput());assert.equal(styled.writable,true,styled.readOnlyReason);assert.equal(styled.sheets[0].tables[0].styleInfo.name,'TableStyleLight9');
  operation.undo();assert.equal(document.state.sheets[0].tables[0].styleInfo.name,'TableStyleMedium2');operation.redo();assert.equal(document.state.sheets[0].tables[0].styleInfo.name,'TableStyleLight9');document.dispose();
});

test('표 만들기는 초안을 보존하고 한 번의 Undo·Redo·저장으로 생성과 이름을 복원한다',async()=>{
  const bytes=minimalXlsx('', '<row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>First</t></is></c><c r="B2"><v>7</v></c></row>'),directory=await fs.mkdtemp(path.join(os.tmpdir(),'jinsheet-new-table-')),file=path.join(directory,'new.xlsx');await fs.writeFile(file,bytes);const document=new JinSheetDocument(uri(file),'xlsx',analyze(bytes),bytes,true,'');document.applyCellEdit(1,1,'9');let operation,count=0;document.onDidChange(change=>{operation=change;count++});assert.deepEqual(document.applyCreateTable(0,{top:0,bottom:1,left:0,right:1},'Sales'),{name:'Sales'});assert.equal(count,1);assert.equal(document.getCell(1,1).value,'9');const creation=operation;creation.undo();assert.equal(document.state.sheets[0].tables.length,0);assert.equal(document.getCell(1,1).value,'9');creation.redo();await document.save();assert.equal(analyze(await fs.readFile(file)).sheets[0].tables[0].name,'Sales');creation.undo();await document.save();assert.equal(analyze(await fs.readFile(file)).sheets[0].tables.length,0);assert.equal(analyze(await fs.readFile(file)).sheets[0].cells.get('B2').value,'9');document.dispose();await fs.rm(directory,{recursive:true,force:true});
});

test('표 생성 실패는 초안·dirty·이력을 유지하고 이름 수정 후 재시도할 수 있다',()=>{
  const bytes=minimalXlsx(),document=new JinSheetDocument(uri('/tmp/create-table-retry.xlsx'),'xlsx',analyze(bytes),bytes,true,'');document.applyCellEdit(1,0,'7');const state=document.state;let count=0;document.onDidChange(()=>count++);assert.throws(()=>document.applyCreateTable(0,{top:0,bottom:1,left:0,right:0},'A1'),/표 이름/);assert.equal(document.state,state);assert.equal(document.dirty,true);assert.equal(document.getCell(1,0).value,'7');assert.equal(count,0);assert.deepEqual(document.applyCreateTable(0,{top:0,bottom:1,left:0,right:0},'Sales'),{name:'Sales'});assert.equal(count,1);document.dispose();
});

test('표 만들기 UI는 범위·이름을 명시적으로 보내고 오류·중복·오래된 응답을 처리한다',()=>{
  const grid=runGridWebview(renderHtml(),{kind:'xlsx',rowCount:3,columnCount:3}),{document,elements,messages,dispatch,documentListeners}=grid;const first=document.querySelector('[data-row="0"][data-column="0"]'),last=document.querySelector('[data-row="2"][data-column="1"]');first.onmousedown({target:first,shiftKey:false,clientX:80,clientY:60,preventDefault(){}});last.onmouseenter();documentListeners.get('mouseup')();elements.get('newTableName').value='Sales';elements.get('createTable').onclick();const request=messages.at(-1);assert.equal(request.type,'create-table');assert.deepEqual(JSON.parse(JSON.stringify(request.range)),{top:0,bottom:2,left:0,right:1});assert.equal(request.name,'Sales');assert.equal(elements.get('createTable').disabled,true);elements.get('createTable').onclick();assert.equal(messages.filter(message=>message.type==='create-table').length,1);dispatch({type:'table-create-error',requestId:request.requestId+1,sheetIndex:0,message:'stale'});assert.equal(elements.get('createTable').disabled,true);dispatch({type:'table-create-error',requestId:request.requestId,sheetIndex:0,message:'bad name'});assert.equal(elements.get('newTableName').value,'Sales');assert.equal(elements.get('createTable').disabled,false);assert.equal(document.activeElement,elements.get('newTableName'));elements.get('newTableName').value='Sales2';elements.get('newTableName').listeners.get('keydown')({key:'Enter',preventDefault(){}});const retry=messages.at(-1);assert.equal(retry.type,'create-table');assert.ok(retry.requestId>request.requestId);dispatch({type:'table-created',requestId:request.requestId,sheetIndex:0,name:'Sales'});assert.equal(elements.get('newTableName').value,'Sales2');dispatch({type:'table-created',requestId:retry.requestId,sheetIndex:0,name:'Sales2'});assert.equal(elements.get('newTableName').value,'');assert.match(elements.get('notice').textContent,/Sales2/);
  for(const model of [{kind:'csv'},{kind:'xlsx',readOnly:true},{kind:'xlsx',tables:[{name:'Old',range:{top:0,bottom:1,left:0,right:1},headerRowCount:1,totalsRowCount:0}]}]){const other=runGridWebview(renderHtml(),{rowCount:3,columnCount:3,...model});if(model.readOnly)other.dispatch({type:'model',rowCount:3,columnCount:3,sheetIndex:0,...model});assert.equal(other.elements.get('createTable').disabled,true,JSON.stringify(model))}
});

test('표 스타일 UI는 선택한 기본 스타일을 명시적으로 적용하고 오류 뒤 입력을 보존한다',()=>{
  const table={name:'Sales',range:{top:0,bottom:2,left:0,right:1},headerRowCount:1,totalsRowCount:0,styleInfo:{name:'TableStyleMedium2',showFirstColumn:false,showLastColumn:false,showRowStripes:true,showColumnStripes:false}},grid=runGridWebview(renderHtml(),{kind:'xlsx',rowCount:3,columnCount:2,tables:[table],cellData:{'0:0':{tableStyle:{fill:'#4472C4',color:'#FFFFFF'}},'1:1':{style:{fill:'#FF0000'},tableStyle:{fill:'#D9E2F3'}}}}),{document,elements,messages,dispatch}=grid;
  assert.equal(elements.get('applyTableStyle').disabled,false);assert.equal(elements.get('tableStyleName').value,'TableStyleMedium2');
  assert.equal(document.querySelector('[data-row="0"][data-column="0"]').style.backgroundColor,'#4472C4');assert.equal(document.querySelector('[data-row="1"][data-column="1"]').style.backgroundColor,'#FF0000');
  elements.get('tableStyleName').value='TableStyleLight9';elements.get('tableStyleName').onchange();elements.get('tableStyleFirstColumn').checked=true;elements.get('tableStyleFirstColumn').onchange();elements.get('tableStyleRows').checked=false;elements.get('tableStyleRows').onchange();elements.get('tableStyleColumns').checked=true;elements.get('tableStyleColumns').onchange();elements.get('applyTableStyle').onclick();
  const request=messages.at(-1);assert.equal(request.type,'table-style');assert.equal(request.tableName,'Sales');assert.deepEqual(JSON.parse(JSON.stringify(request.styleInfo)),{name:'TableStyleLight9',showFirstColumn:true,showLastColumn:false,showRowStripes:false,showColumnStripes:true});assert.equal(elements.get('applyTableStyle').disabled,true);
  dispatch({type:'table-style-error',requestId:request.requestId+1,sheetIndex:0,tableName:'Sales',message:'stale'});assert.equal(elements.get('applyTableStyle').disabled,true);
  dispatch({type:'table-style-error',requestId:request.requestId,sheetIndex:0,tableName:'Sales',message:'retry'});assert.equal(elements.get('applyTableStyle').disabled,false);assert.equal(elements.get('tableStyleName').value,'TableStyleLight9');assert.equal(elements.get('tableStyleFirstColumn').checked,true);assert.equal(elements.get('tableStyleRows').checked,false);
});

test('비어 있는 워크시트에도 한 셀 표를 만들고 첫 행을 추가한다',()=>{
  const bytes=minimalXlsx('',''),model=analyze(bytes),document=new JinSheetDocument(uri('/tmp/create-blank-table.xlsx'),'xlsx',model,bytes,true,'');assert.deepEqual(document.applyCreateTable(0,{top:0,bottom:0,left:0,right:0},''),{name:'Table1'});assert.equal(document.getCell(0,0).value,'Column1');assert.deepEqual(document.appendTableRow(0,'Table1',0,0),{row:1,column:0});document.applyCellEdit(1,0,'7');assert.equal(analyze(document.createOutput()).sheets[0].tables[0].range.bottom,1);document.dispose();
});

test('CSV 값 필터는 세션에서 행을 숨기고 편집·지우기에 반응한다', () => {
  const rows = [['이름', '팀'], ['민수', 'A'], ['지우', 'B'], ['서연', 'A'], ['도윤', '']].map(row => row.map(value => ({ value, kind: 'string' })));
  const document = { kind: 'csv', state: { rows }, csvAutoFilter: null, drafts: new Map(), writable: true, reason: '', dirty: false, getCell(row, column) { return rows[row]?.[column] ?? null; } };
  const filter = autoFilterRange(document, 0, 0, 0), options = autoFilterOptions(document, 0, 1, filter);
  assert.deepEqual(filter, { top: 0, bottom: 4, left: 0, right: 1, columns: [] });
  assert.deepEqual(options.values, ['A', 'B']); assert.deepEqual(options.selectedValues, ['A', 'B']); assert.equal(options.blanks, true); assert.equal(options.includeBlanks, true);
  document.csvAutoFilter = updateAutoFilterColumn(document, 0, 1, ['A'], false, false, filter);
  assert.deepEqual(viewModel(document).hiddenRows, [2, 4]); assert.equal(viewModel(document).dirty, false);
  rows[2][1].value = 'A'; assert.deepEqual(viewModel(document).hiddenRows, [4]);
  assert.throws(() => updateAutoFilterColumn(document, 0, 1, ['missing'], false, false, document.csvAutoFilter), /자동 필터 선택 값/);
  document.csvAutoFilter = updateAutoFilterColumn(document, 0, 1, [], false, true, document.csvAutoFilter);
  assert.equal(viewModel(document).hiddenRows, undefined); assert.deepEqual(document.csvAutoFilter.columns, []);
});

test('CSV 필터 UI는 값·조건 모드와 세션 한정 안내를 제공한다', () => {
  const { elements, messages, dispatch } = runGridWebview(renderHtml(), { kind: 'csv', rowCount: 3, columnCount: 2 });
  elements.get('filterValuesButton').onclick();
  const request = messages.at(-1);
  assert.equal(request.type, 'filter-options'); assert.equal(request.sheetIndex, 0);
  dispatch({ type: 'filter-options', sheetIndex: 0, column: request.column, options: { values: ['A', 'B'], selectedValues: ['A', 'B'], blanks: true, includeBlanks: true } });
  assert.equal(elements.get('filterPanel').hidden, false); assert.equal(elements.get('filterSessionNote').hidden, false);
  assert.equal(elements.get('filterMode').value, 'values');
  elements.get('filterApply').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1).values)), ['A', 'B']); assert.equal(messages.at(-1).blanks, true);
  const valueRequest = messages.at(-1);
  dispatch({ type: 'filter-applied', requestId: valueRequest.requestId, sheetIndex: 0, column: valueRequest.column });
  elements.get('filterValuesButton').onclick();
  const conditionRequest = messages.at(-1);
  dispatch({ type: 'filter-options', sheetIndex: 0, column: conditionRequest.column, options: { values: ['A', 'B'], selectedValues: ['B'], blanks: false, includeBlanks: false, custom: { and: false, conditions: [{ operator: 'equal', value: '*B*' }] } } });
  assert.equal(elements.get('filterMode').value, 'conditions'); assert.equal(elements.get('filterConditionControls').hidden, false);
  elements.get('filterApply').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1).custom)), { and: false, conditions: [{ operator: 'equal', value: '*B*' }] });
});

test('CSV/TSV 행·열 그룹은 세션에만 유지되고 접기·개요 수준·Undo를 처리한다', () => {
  const { updateCsvOutlineGroups } = require('../src/extension'), bytes = Buffer.from('Name,Team\nA,red\nB,blue\nC,red\nD,green'), state = load(bytes, '.csv');
  const document = new JinSheetDocument(uri('/tmp/session-outline.csv'), 'csv', state, bytes, true, ''); document.setInitialCsvState(state);
  const changes = []; document.onDidChange(change => changes.push(change));
  assert.equal(updateCsvOutlineGroups(document, 0, 'row', 'group', 1, 3), true);
  assert.deepEqual(viewModel(document).rowGroups, [{ start: 1, end: 3, level: 1, summary: 4, collapsed: false }]);
  updateCsvOutlineGroups(document, 0, 'row', 'group', 2, 3);
  assert.deepEqual(viewModel(document).rowGroups.map(group => [group.start, group.end, group.level, group.summary]), [[1, 3, 1, 4], [2, 3, 2, 4]]);
  updateCsvOutlineGroups(document, 0, 'row', 'collapse-summary', undefined, undefined, 4);
  assert.deepEqual(viewModel(document).hiddenRows, [1, 2, 3]);
  changes.at(-1).undo(); assert.equal(viewModel(document).hiddenRows, undefined);
  changes.at(-1).redo(); assert.deepEqual(viewModel(document).hiddenRows, [1, 2, 3]);
  updateCsvOutlineGroups(document, 0, 'row', 'show-level', undefined, undefined, undefined, 8);
  assert.equal(viewModel(document).hiddenRows, undefined); assert.equal(document.dirty, false);
  updateCsvOutlineGroups(document, 0, 'row', 'ungroup', 2, 3);
  assert.deepEqual(viewModel(document).rowGroups.map(group => [group.start, group.end, group.level]), [[1, 3, 1]]);
  updateCsvOutlineGroups(document, 0, 'column', 'group', 0, 1);
  updateCsvOutlineGroups(document, 0, 'column', 'collapse-summary', undefined, undefined, 2);
  assert.deepEqual(viewModel(document).hiddenColumns, [0, 1]);
  assert.throws(() => updateCsvOutlineGroups(document, 0, 'row', 'group', 9999, 9999), /요약 행·열/);
  document.dispose();
});

test('CSV 조건 필터는 텍스트 조건과 숫자형 텍스트 비교를 결합하고 여러 열에 교차 적용한다', () => {
  const rows = [['이름', '점수'], ['Alpha', '2'], ['Beta', '10'], ['Gamma', '30'], ['Delta', '40']].map(row => row.map(value => ({ value, kind: 'string' })));
  const document = { kind: 'csv', state: { rows }, csvAutoFilter: null, drafts: new Map(), writable: true, reason: '', dirty: false, getCell(row, column) { return rows[row]?.[column] ?? null; } };
  const range = autoFilterRange(document, 0, 0, 0);
  document.csvAutoFilter = updateAutoFilterCustomColumn(document, 0, 1, { and: true, conditions: [{ operator: 'greaterThan', value: '9' }, { operator: 'lessThanOrEqual', value: '30' }] }, range);
  assert.deepEqual(viewModel(document).hiddenRows, [1, 4]);
  const options = autoFilterOptions(document, 0, 1, document.csvAutoFilter);
  assert.deepEqual(options.selectedValues, ['10', '30']); assert.deepEqual(options.custom.conditions.map(condition => condition.operator), ['greaterThan', 'lessThanOrEqual']);
  document.csvAutoFilter = updateAutoFilterCustomColumn(document, 0, 0, { and: true, conditions: [{ operator: 'equal', value: '*a*' }] }, document.csvAutoFilter);
  assert.deepEqual(viewModel(document).hiddenRows, [1, 4]);
  assert.throws(() => updateAutoFilterCustomColumn(document, 0, 0, { and: true, conditions: [{ operator: 'equal', value: 7 }] }, document.csvAutoFilter), /조건 필터/);
});

test('CSV 그룹 UI를 사용할 수 있고 세션 한정 안내를 표시한다', () => {
  const { elements, messages } = runGridWebview(renderHtml(), { kind: 'csv', rowCount: 5, columnCount: 2 });
  assert.equal(elements.get('groupRows').disabled, false); assert.equal(elements.get('csvOutlineSessionNote').hidden, false);
  elements.get('groupRows').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'outline-groups', axis: 'row', action: 'group', start: 0, end: 0, sheetIndex: 0 });
});

test('텍스트 나누기 계획은 행별 필드를 만들고 목적지 충돌을 명시적으로 확인한다', () => {
  const rows = [
    [{ value: 'North,"West, Coast"', kind: 'string' }, { value: '', kind: 'blank' }, { value: '', kind: 'blank' }],
    [{ value: 'South,East', kind: 'string' }, { value: 'keep', kind: 'string' }, { value: '', kind: 'blank' }]
  ];
  const document = { kind: 'csv', writable: true, state: { rows }, getCell(row, column) { return rows[row]?.[column] || null; } };
  const range = { startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 };
  const plan = planTextToColumns(document, range, ',', false);
  assert.deepEqual(plan.preview, [['North', 'West, Coast'], ['South', 'East']]);
  assert.equal(plan.width, 2);
  assert.deepEqual(plan.conflicts, [{ row: 1, column: 1 }]);
  assert.throws(() => buildTextToColumnsEdits(document, range, ','), error => error.needsConfirmation === true && error.conflicts === 1);
  assert.deepEqual(buildTextToColumnsEdits(document, range, ',', false, 0, true), [
    { row: 0, column: 0, value: 'North' }, { row: 0, column: 1, value: 'West, Coast' },
    { row: 1, column: 0, value: 'South' }, { row: 1, column: 1, value: 'East' }
  ]);
});

test('중복 행 제거 계획은 선택 열 기준으로 첫 레코드를 남기고 머리글을 보존한다', () => {
  const rows = [
    ['Name', 'Region', 'Value'],
    ['Ada', 'West', '1'],
    ['ada', 'WEST', '2'],
    ['Lin', 'East', '3'],
    ['Lin', 'East', '4']
  ].map(row => row.map(value => ({ value, kind: 'string' })));
  const document = { kind: 'csv', writable: true, state: { rows }, getCell(row, column) { return rows[row]?.[column] || null; } };
  const range = { startRow: 0, endRow: 4, startColumn: 0, endColumn: 2 };
  assert.deepEqual(planRemoveDuplicates(document, range, [0, 1], true), {
    range: { top: 0, bottom: 4, left: 0, right: 2 }, duplicateRows: [2, 4], removedCount: 2, hasHeader: true, columns: [0, 1]
  });
  assert.deepEqual(planRemoveDuplicates(document, range, [0], false).duplicateRows, [2, 4]);
  assert.throws(() => planRemoveDuplicates(document, range, [], true), /비교 열/);
  assert.throws(() => planRemoveDuplicates(document, range, [3], true), /비교 열/);
});

test('중복 행 제거는 CSV에서 한 번에 undo/redo하고 저장 결과를 유지한다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-remove-duplicates-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'duplicates.csv'), bytes = Buffer.from('Name,Region,Value\nAda,West,1\nada,WEST,2\nLin,East,3\nLin,East,4\n');
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'csv', load(bytes, '.csv'), bytes, true, '');
  document.setInitialCsvState(document.state);
  let operation, changeCount = 0;
  document.onDidChange(change => { operation = change; changeCount++; });
  const plan = planRemoveDuplicates(document, { startRow: 0, endRow: 4, startColumn: 0, endColumn: 2 }, [0, 1], true);
  removeDuplicateRows(document, plan);
  assert.equal(changeCount, 1);
  assert.deepEqual(document.state.rows.map(row => row.map(cell => cell.value)), [['Name', 'Region', 'Value'], ['Ada', 'West', '1'], ['Lin', 'East', '3']]);
  operation.undo();
  assert.equal(document.state.rows.length, 5);
  operation.redo();
  await document.save();
  assert.equal((await fs.readFile(file, 'utf8')), 'Name,Region,Value\nAda,West,1\nLin,East,3\n');
  document.dispose();
});

test('중복 행 제거는 XLSX 구조와 저장된 셀을 갱신하고 undo/redo한다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-remove-duplicates-xlsx-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const rows = '<row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Value</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Ada</t></is></c><c r="B2"><v>1</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>ADA</t></is></c><c r="B3"><v>2</v></c></row><row r="4"><c r="A4" t="inlineStr"><is><t>Lin</t></is></c><c r="B4"><v>3</v></c></row>';
  const file = path.join(directory, 'duplicates.xlsx'), bytes = minimalXlsx('', rows);
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  let operation, changeCount = 0;
  document.onDidChange(change => { operation = change; changeCount++; });
  const plan = planRemoveDuplicates(document, { startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 }, [0], true);
  removeDuplicateRows(document, plan);
  assert.equal(changeCount, 1);
  assert.equal(document.getCell(2, 0).value, 'Lin');
  operation.undo();
  assert.equal(document.getCell(2, 0).value, 'ADA');
  operation.redo();
  await document.save();
  const saved = analyze(await fs.readFile(file));
  assert.equal(saved.sheets[0].cells.get('A3').value, 'Lin');
  assert.equal(saved.sheets[0].cells.has('A4'), false);
  document.dispose();
});

test('중복 행 제거는 Excel 표 데이터 행을 지원하고 첫 행을 보존한다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-remove-duplicates-table-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const files = Object.fromEntries(Object.entries(unzipSync(tableXlsx())).map(([name, bytes]) => [name, new TextDecoder().decode(bytes)]));
  files['xl/worksheets/sheet1.xml'] = files['xl/worksheets/sheet1.xml'].replace('<t>Beta</t>', '<t>ALPHA</t>');
  const bytes = zipSync(Object.fromEntries(Object.entries(files).map(([name, value]) => [name, new TextEncoder().encode(value)]))), file = path.join(directory, 'table.xlsx');
  await fs.writeFile(file, bytes);
  const document = new JinSheetDocument(uri(file), 'xlsx', analyze(bytes), bytes, true, '');
  let operation;
  document.onDidChange(change => { operation = change; });
  const range = { startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 };
  const plan = planRemoveDuplicates(document, range, [0], true);
  assert.deepEqual(plan.duplicateRows, [2]);
  removeDuplicateRows(document, plan);
  assert.equal(document.state.sheets[0].tables[0].range.bottom, 1);
  operation.undo();
  assert.equal(document.state.sheets[0].tables[0].range.bottom, 2);
  assert.deepEqual(planRemoveDuplicates(document, range, [0], false).duplicateRows, [2]);
  document.dispose();
});

test('중복 행 제거 도구는 영어와 한국어 UI에 연결된다', () => {
  const html = renderHtml();
  assert.match(html, /id='removeDuplicates'/);
  assert.match(html, /remove-duplicates-preview/);
  assert.match(html, /type:'remove-duplicates',requestId:id/);
  assert.match(renderHtml(undefined, undefined, undefined, 'ko'), /중복 항목 제거/);
});

test('중복 제거 UI는 선택 범위 미리보기를 확인한 뒤에만 적용 요청을 보낸다', () => {
  const { document, documentListeners, elements, messages, dispatch } = runGridWebview(renderHtml(), { kind: 'csv', rowCount: 4, columnCount: 2 });
  const first = document.querySelector('[data-row="0"][data-column="0"]'), last = document.querySelector('[data-row="3"][data-column="1"]');
  first.onmousedown({ target: first, preventDefault() {}, clientX: 10, clientY: 40, shiftKey: false });
  last.onmouseenter(); documentListeners.get('mouseup')();
  document.querySelector('#removeDuplicates').onclick();
  const preview = messages.at(-1);
  assert.equal(preview.type, 'remove-duplicates-preview');
  assert.deepEqual(JSON.parse(JSON.stringify(preview.range)), { startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 });
  assert.deepEqual(Array.from(preview.columns), [0, 1]);
  assert.equal(preview.hasHeader, true);
  assert.equal(document.querySelector('#removeDuplicatesApply').disabled, true);
  dispatch({ type: 'remove-duplicates-preview', requestId: preview.requestId, sheetIndex: 0, removedCount: 1 });
  assert.equal(document.querySelector('#removeDuplicatesApply').disabled, false);
  document.querySelector('#removeDuplicatesApply').onclick();
  const apply = messages.at(-1);
  assert.equal(apply.type, 'remove-duplicates');
  assert.equal(apply.requestId > preview.requestId, true);
  dispatch({ type: 'remove-duplicates-applied', requestId: apply.requestId, sheetIndex: 0, removedCount: 1 });
  assert.equal(document.querySelector('#removeDuplicatesPanel').hidden, true);
  assert.equal(elements.get('notice').textContent, 'Removed 1 duplicate rows.');
});

test('텍스트 나누기 도구 부트스트랩은 기본 초기화 흐름에 포함된다', () => {
  const html = renderHtml();
  assert.match(html, /button\.id='textToColumns'/);
  assert.match(html, /id='textToColumnsPreview'/);
  assert.match(html, /id='textToColumnsApply'/);
  assert.match(html, /type:'text-to-columns-preview'/);
  assert.match(html, /type:'text-to-columns',requestId:applyRequest/);
});

test('텍스트 나누기는 CSV와 XLSX에서 적용·undo·redo·저장·재열기를 보존한다', async t => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-text-columns-roundtrip-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const range = { startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 };

  const csvBytes = Buffer.from('"North, Coast",\nSouth,East\n');
  const csvFile = path.join(directory, 'split.csv');
  await fs.writeFile(csvFile, csvBytes);
  const csv = new JinSheetDocument(uri(csvFile), 'csv', load(csvBytes, '.csv'), csvBytes, true, '');
  csv.setInitialCsvState(csv.state);
  assert.throws(() => buildTextToColumnsEdits(csv, range, ','), error => error.needsConfirmation === true && error.conflicts === 1);
  const csvEdits = buildTextToColumnsEdits(csv, range, ',', false, 0, true);
  let csvOperation;
  csv.onDidChange(change => { csvOperation = change; });
  csv.applyCellEdits(csvEdits, 0, '텍스트 나누기', new Set(), true, { skipExpansion: true, skipHeaders: true });
  assert.deepEqual(csv.state.rows.map(row => row.map(cell => cell.value)), [['North', ' Coast'], ['South', '']]);
  csvOperation.undo();
  assert.equal(csv.getCell(0, 0).value, 'North, Coast');
  assert.deepEqual(csv.state.rows.map(row => row.map(cell => cell.value)), [['North, Coast', ''], ['South', 'East']]);
  csvOperation.redo();
  await csv.save();
  const csvReopened = load(await fs.readFile(csvFile), '.csv');
  assert.deepEqual(csvReopened.rows.map(row => row.map(cell => cell.value)), [['North', ' Coast'], ['South', '']]);
  csv.dispose();

  const xlsxBytes = minimalXlsx('', '<row r="1"><c r="A1" t="inlineStr"><is><t>North, Coast</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>South,East</t></is></c></row>');
  const xlsxFile = path.join(directory, 'split.xlsx');
  await fs.writeFile(xlsxFile, xlsxBytes);
  const xlsx = new JinSheetDocument(uri(xlsxFile), 'xlsx', analyze(xlsxBytes), xlsxBytes, true, '');
  let xlsxOperation;
  xlsx.onDidChange(change => { xlsxOperation = change; });
  const xlsxEdits = buildTextToColumnsEdits(xlsx, range, ',', false, 0, true);
  xlsx.applyCellEdits(xlsxEdits, 0, '텍스트 나누기', new Set(), true, { skipExpansion: true, skipHeaders: true });
  assert.equal(xlsx.getCell(0, 0).value, 'North');
  assert.equal(xlsx.getCell(0, 1).value, ' Coast');
  xlsxOperation.undo();
  assert.equal(xlsx.getCell(0, 0).value, 'North, Coast');
  xlsxOperation.redo();
  await xlsx.save();
  const xlsxReopened = analyze(await fs.readFile(xlsxFile));
  assert.equal(xlsxReopened.sheets[0].cells.get('A1').value, 'North');
  assert.equal(xlsxReopened.sheets[0].cells.get('B1').value, ' Coast');
  assert.equal(xlsxReopened.sheets[0].cells.get('A2').value, 'South');
  assert.equal(xlsxReopened.sheets[0].cells.get('B2').value, 'East');
  xlsx.dispose();
});

test('텍스트 나누기는 여러 열 선택·수식 셀·시트 경계와 빈 입력을 보호한다', () => {
  const cells = new Map([['0:0', { value: 'a,b', kind: 'string' }], ['1:0', { value: 'x,y', formula: 'A1', hasFormula: true }]]);
  const document = { kind: 'csv', writable: true, state: { rows: [] }, getCell(row, column) { return cells.get(row + ':' + column) || null; } };
  assert.throws(() => buildTextToColumnsEdits(document, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 }, ','), /한 열/);
  assert.throws(() => buildTextToColumnsEdits(document, { startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 }, ','), /수식 셀/);
  assert.throws(() => buildTextToColumnsEdits(document, { startRow: 0, endRow: 0, startColumn: 16383, endColumn: 16383 }, ','), /시트 경계/);
  const empty = { kind: 'csv', writable: true, state: { rows: [] }, getCell() { return { value: '' }; } };
  assert.throws(() => buildTextToColumnsEdits(empty, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, ','), /구분자를 찾지 못했습니다/);
});
