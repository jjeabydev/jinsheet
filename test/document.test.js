'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const Module = require('node:module');
const vm = require('node:vm');
const { zipSync } = require('../vendor/node_modules/fflate');

class EventEmitter {
  constructor() { this.listeners = new Set(); this.event = listener => { this.listeners.add(listener); return { dispose: () => this.listeners.delete(listener) }; }; }
  fire(value) { for (const listener of this.listeners) listener(value); }
  dispose() { this.listeners.clear(); }
}
const originalLoad = Module._load;
Module._load = function(request, parent, isMain) { if (request === 'vscode') return { EventEmitter }; return originalLoad.call(this, request, parent, isMain); };
const { applyFilterAction, applySortAction, buildRowMap, findMatches, getRange, isValidSheetRequest, mapViewRow, matchesSelectedSheet, estimateDocumentColumnWidth, navigateCell, navigateViewCell, parseClipboard, serializeClipboard, JinSheetDocument, renderHtml, viewModel } = require('../src/extension');
const { load } = require('../src/csv');
const { analyze } = require('../src/xlsx');
let activeDocument = null;

function isWithin(node, ancestor) { for (let current = node; current; current = current.parentElement) if (current === ancestor) return true; return false; }

function minimalXlsx(sheetViews = '') {
  const xml = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="s1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="s1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    'xl/worksheets/sheet1.xml': `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${sheetViews}<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>old</t></is></c></row></sheetData></worksheet>`
  };
  return zipSync(Object.fromEntries(Object.entries(xml).map(([name, value]) => [name, new TextEncoder().encode(value)])));
}
function uri(fsPath) { return { fsPath, scheme: 'file', toString: () => fsPath }; }

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
    this.textContent = '';
    this.value = '';
  }
  append(child) { child.parentElement = this; this.children.push(child); }
  replaceChildren(...children) { for (const child of this.children) { if (activeDocument && isWithin(activeDocument.activeElement, child)) activeDocument.activeElement = activeDocument.body; child.parentElement = null; } this.children = []; for (const child of children) this.append(child); }
  remove() { if (activeDocument && isWithin(activeDocument.activeElement, this)) activeDocument.activeElement = activeDocument.body; if (this.parentElement) this.parentElement.children = this.parentElement.children.filter(child => child !== this); this.parentElement = null; }
  setAttribute(name, value) { this.attributes[name] = String(value); }
  addEventListener(name, callback) { this.listeners.set(name, callback); }
  closest(selector) { let current = this; while (current) { if (selector === '#grid' && current.id === 'grid') return current; current = current.parentElement; } return null; }
  focus() { if (activeDocument) activeDocument.activeElement = this; }
  select() {}
  setSelectionRange() {}
}

function runGridWebview(html, { rowCount = 2, columnCount = 2, style = null, kind = 'csv', columnWidths = undefined, freezePane = undefined } = {}) {
  const elements = new Map(), documentListeners = new Map(), windowListeners = new Map(), messages = [];
  const get = id => { if (!elements.has(id)) { const element = new FakeElement(id === 'search' || id === 'formula' ? 'INPUT' : 'DIV'); element.id = id; elements.set(id, element); } return elements.get(id); };
  const document = {
    body: get('body'),
    getElementById: get,
    createElement: name => new FakeElement(name.toUpperCase()),
    addEventListener: (name, callback) => documentListeners.set(name, callback),
    querySelector: selector => {
      const cells = get('canvas').children.filter(child => child.classList.contains('cell'));
      if (selector === '.cell.selected') return cells.find(cell => cell.classList.contains('selected')) || null;
      const match = /^\[data-row="(\d+)"\]\[data-column="(\d+)"\]$/.exec(selector);
      return match ? cells.find(cell => cell.dataset.row === match[1] && cell.dataset.column === match[2]) || null : null;
    },
    querySelectorAll: selector => selector === '.cell' ? get('canvas').children.filter(child => child.classList.contains('cell')) : []
  };
  activeDocument = document;
  document.activeElement = document.body;
  const grid = get('grid'), canvas = get('canvas');
  grid.clientHeight = 300; grid.clientWidth = 500; grid.scrollTop = 0; grid.scrollLeft = 0; canvas.parentElement = grid;
  const vscode = { postMessage: message => messages.push(message) };
  const window = { addEventListener: (name, callback) => windowListeners.set(name, callback) };
  const script = [...html.matchAll(/<script nonce="[^"]+">([\s\S]*?)<\/script>/g)].at(-1)[1];
  new vm.Script(script).runInNewContext({ acquireVsCodeApi: () => vscode, document, window, setTimeout, clearTimeout });
  const dispatch = data => windowListeners.get('message')({ data });
  const respondToRange = () => {
    const request = messages.filter(message => message.type === 'range').at(-1);
    const rows = Array.from({ length: request.rowCount }, (_, r) => Array.from({ length: request.columnCount }, (_, c) => {
      const row = request.viewRows?.[r] ?? (request.rowStart + r), column = request.columnIndexes?.[c] ?? (request.columnStart + c);
      const values = { '0:0': 'a', '0:1': 'b', '1:0': '1', '1:1': '2' };
      return { value: values[row + ':' + column] ?? `${row + 1}:${column + 1}`, style: row === 1 && column === 1 ? style : null };
    }));
    dispatch({ type: 'range', requestId: request.requestId, rowStart: request.rowStart, viewRows: request.viewRows, rowIndexes: request.viewRows, columnStart: request.columnStart, columnIndexes: request.columnIndexes, rowCount: request.rowCount, columnCount: request.columnCount, sheetIndex: request.sheetIndex, rows });
  };
  dispatch({ type: 'model', kind, readOnly: false, rowCount, columnCount, columnWidths, freezePane, sheetIndex: 0, dirty: false });
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
  assert.equal(viewModel(document).columnWidths[1], 22.14);
  operation.undo();
  assert.equal(document.dirty, false);
  operation.redo();
  assert.equal(document.dirty, true);
  await document.save();
  assert.equal(document.dirty, false);
  const reopened = analyze(await fs.readFile(file));
  assert.equal(reopened.sheets[0].columnWidths[1], 22.14);
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

test('클립보드 TSV 인용·셀 수·셀 크기 제한을 검사한다', () => {
  assert.deepEqual(parseClipboard('one\t"two\tparts"\r\n"line 1\nline 2"\t"say ""hi"""'), [['one', 'two\tparts'], ['line 1\nline 2', 'say "hi"']]);
  assert.deepEqual(parseClipboard('a"b'), [['a"b']]);
  assert.deepEqual(parseClipboard(''), []);
  assert.throws(() => parseClipboard('a\t"broken'), /닫히지 않은/);
  assert.throws(() => parseClipboard('x'.repeat(100001)), /100,000자/);
  assert.throws(() => parseClipboard('x'.repeat(1000001)), /1,000,000자/);
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

test('정렬·필터 화면은 원본 행 주소를 유지해 편집·복사·검색 범위를 연결한다', () => {
  const bytes = Buffer.from('product,quantity\npear,12\napple,2\npear,4\n');
  const document = new JinSheetDocument(uri('/tmp/sort-filter.csv'), 'csv', load(bytes, '.csv'), bytes, true, '');
  document.setInitialCsvState(document.state);
  const view = { sortColumn: 1, sortDirection: 'asc', filterColumn: 0, filterText: 'pear', hasHeader: true };
  const rowMap = buildRowMap(document, 0, view);
  assert.deepEqual(rowMap, [0, 3, 1]);
  assert.equal(mapViewRow(1, rowMap), 3);
  assert.throws(() => mapViewRow(3, rowMap), /범위를 벗어났습니다/);
  const range = getRange(document, { requestId: 1, rowStart: 1, columnStart: 0, rowCount: 2, columnCount: 2 }, rowMap);
  assert.deepEqual(range.rowIndexes, [3, 1]);
  assert.deepEqual(range.rows.map(row => row.map(cell => cell.value)), [['pear', '4'], ['pear', '12']]);
  assert.equal(serializeClipboard(document, { startRow: 1, endRow: 2, startColumn: 0, endColumn: 1 }, 0, rowMap), 'pear\t4\r\npear\t12');
  assert.deepEqual(findMatches(document, 'pear', 0, rowMap).matches.map(match => [match.row, match.address]), [[1, 'A4'], [2, 'A2']]);
  assert.deepEqual(navigateViewCell(document, { row: 1, column: 0, direction: 'right', requestId: 2 }, rowMap), { type: 'navigation', requestId: 2, row: 1, column: 1, extend: false });
  assert.deepEqual(viewModel(document, 0, 'en', view).rowMap, rowMap);
  document.applyCellEdit(mapViewRow(1, rowMap), 0, 'pear updated');
  const saved = load(document.createOutput(), '.csv');
  assert.equal(saved.rows[3][0].value, 'pear updated');
  assert.equal(saved.rows[2][0].value, 'apple');
  document.dispose();
});

test('정렬 키 추가·방향 전환·해제·교체가 우선순위를 올바르게 갱신한다', () => {
  const view = { sortKeys: [{ column: 0, direction: 'asc' }] };
  applySortAction(view, { column: 1, mode: 'add', direction: 'asc' });
  assert.deepEqual(view.sortKeys, [{ column: 0, direction: 'asc' }, { column: 1, direction: 'asc' }]);
  applySortAction(view, { column: 1, mode: 'add', direction: 'desc' });
  assert.deepEqual(view.sortKeys, [{ column: 0, direction: 'asc' }, { column: 1, direction: 'desc' }]);
  applySortAction(view, { column: 0, mode: 'remove', direction: 'asc' });
  assert.deepEqual(view.sortKeys, [{ column: 1, direction: 'desc' }]);
  assert.equal(view.sortColumn, 1);
  assert.equal(view.sortDirection, 'desc');
  applySortAction(view, { column: 2, mode: 'replace', direction: 'asc' });
  assert.deepEqual(view.sortKeys, [{ column: 2, direction: 'asc' }]);
});

test('열별 필터 조건은 누적되고 빈 조건은 해당 열에서만 제거된다', () => {
  const view = { filterRules: [] };
  applyFilterAction(view, { column: 0, query: 'pear' });
  applyFilterAction(view, { column: 1, query: '1' });
  assert.deepEqual(view.filterRules, [{ column: 0, query: 'pear' }, { column: 1, query: '1' }]);
  applyFilterAction(view, { column: 1, query: '' });
  assert.deepEqual(view.filterRules, [{ column: 0, query: 'pear' }]);
  applyFilterAction(view, { column: 0, query: ' ' });
  assert.deepEqual(view.filterRules, []);
});

test('다중 열 필터는 모든 조건을 만족하는 행만 남기고 머리글은 유지한다', () => {
  const bytes = Buffer.from('fruit,quantity,color\npear,12,green\napple,12,red\npear,4,green\npear,21,red\n');
  const document = new JinSheetDocument(uri('/tmp/multi-filter.csv'), 'csv', load(bytes, '.csv'), bytes, true, '');
  document.setInitialCsvState(document.state);
  const rowMap = buildRowMap(document, 0, {
    filterRules: [{ column: 0, query: 'pear' }, { column: 1, query: '1' }],
    hasHeader: true
  });
  assert.deepEqual(rowMap, [0, 1, 4]);
  document.dispose();
});

test('여러 열 정렬은 우선순위를 적용하고 같은 값은 원본 순서를 보존한다', () => {
  const bytes = Buffer.from('fruit,quantity,label\npear,12,a\napple,12,b\npear,4,c\napple,2,d\npear,12,e\n');
  const document = new JinSheetDocument(uri('/tmp/multi-sort.csv'), 'csv', load(bytes, '.csv'), bytes, true, '');
  document.setInitialCsvState(document.state);
  const rowMap = buildRowMap(document, 0, {
    sortKeys: [{ column: 0, direction: 'asc' }, { column: 1, direction: 'desc' }],
    hasHeader: true
  });
  assert.deepEqual(rowMap, [0, 2, 4, 1, 5, 3]);
  assert.deepEqual(viewModel(document, 0, 'en', { sortKeys: [{ column: 0, direction: 'asc' }, { column: 1, direction: 'desc' }] }).sortKeys,
    [{ column: 0, direction: 'asc' }, { column: 1, direction: 'desc' }]);
  document.dispose();
});

test('정렬·필터는 행 제한을 넘으면 명확히 거절한다', () => {
  const bytes = Buffer.from('');
  const cells = new Map(Array.from({ length: 100001 }, (_, row) => ['A' + (row + 1), { row, column: 0, address: 'A' + (row + 1), value: String(row) }]));
  const document = new JinSheetDocument(uri('/tmp/large-view.xlsx'), 'xlsx', { sheets: [{ cells }] }, bytes, true, '');
  assert.throws(() => buildRowMap(document, 0, { sortColumn: 0, sortDirection: 'asc' }), /100,000개 이하/);
  document.dispose();
});

test('정렬·필터된 실제 XLSX 편집은 표시 행이 아니라 원본 셀을 저장한다', async t => {
  const bytes = await fs.readFile(path.join(__dirname, 'fixtures', 'JinSheet-Sample.xlsx'));
  const document = new JinSheetDocument(uri('/tmp/sort-filter.xlsx'), 'xlsx', analyze(bytes), bytes, true, '');
  t.after(() => document.dispose());
  const view = { sortColumn: 1, sortDirection: 'asc', filterColumn: 0, filterText: '배', hasHeader: true };
  const rowMap = buildRowMap(document, 0, view);
  assert.deepEqual(rowMap, [0, 2]);
  document.applyCellEdit(mapViewRow(1, rowMap), 1, '8');
  const reopened = analyze(document.createOutput());
  assert.equal(reopened.sheets[0].cells.get('B2').value, '12');
  assert.equal(reopened.sheets[0].cells.get('A3').value, '배');
  assert.equal(reopened.sheets[0].cells.get('B3').value, '8');
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

test('Webview uses English by default and Korean for a Korean VS Code locale', () => {
  const english = renderHtml(null, null, null, 'en-US');
  const korean = renderHtml(null, null, null, 'ko-KR');
  assert.match(english, /<html lang="en">/);
  assert.match(english, />\+ Row</);
  assert.match(english, /Find in sheet/);
  assert.doesNotMatch(english, /[가-힣]/);
  assert.match(korean, /<html lang="ko">/);
  assert.match(korean, />\+ 행</);
  assert.match(korean, /시트에서 찾기/);
  assert.match(english, /\.cell\.numeric\{text-align:right\}/);
  assert.match(english, /cell\.classList\.toggle\('numeric',data\?\.kind==='number'\)/);
  const formulaReason = '수식이 포함되어 있어 저장할 수 없습니다. 수식과 캐시값은 보기 전용입니다.';
  const unsupported = { kind: 'xlsx', reason: formulaReason, state: { sheets: [] }, uri: { fsPath: '/tmp/sample.xlsx' } };
  assert.match(viewModel(unsupported, 0, 'en').reason, /contains formulas/);
  assert.equal(viewModel(unsupported, 0, 'ko').reason, formulaReason);
});

test('Webview search input sends only the latest query after a short pause', async () => {
  const { elements, messages } = runGridWebview(renderHtml());
  const input = elements.get('search'), onInput = input.listeners.get('input');
  onInput({ target: { value: 'alpha' } });
  onInput({ target: { value: 'alpha beta' } });
  await new Promise(resolve => setTimeout(resolve, 140));
  assert.deepEqual(messages.filter(message => message.type === 'search').map(message => message.query), ['alpha beta']);
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
  assert.match(renderHtml(), /cell\.textContent=data\?\.displayValue\?\?data\?\.value\?\?''/);
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
  assert.match(html, /#grid,#readerCanvas\{scrollbar-color:var\(--vscode-scrollbarSlider-background/);
  assert.match(html, /scrollbar-width:thin/);
  assert.match(html, /\.rowhead\{width:52px;z-index:1\}\.colhead\{top:0;z-index:1/);
  assert.match(html, /id="insertRow"/);
  assert.match(html, /id="deleteColumn"/);
  assert.match(html, /type:'structure'/);
  assert.match(html, /selectedRange/);
  assert.match(html, /range:\s*selectedRange/);
  assert.match(html, /cell\.onmouseenter/);
  assert.match(html, /id="formula" type="text"/);
  assert.match(html, /formula\.disabled=Boolean\(model\.readOnly\|\|data\?\.formula/);
  assert.match(html, /getElementById\('formula'\)\.addEventListener\('keydown'/);
  assert.match(html, /type: 'navigate'/);
  assert.match(html, /e\.key === 'Home'/);
  assert.match(html, /model.gridRowCount=model.rowMap\?model.rowMap.length:XLSX_MAX_ROWS/);
  assert.match(html, /type:'sort'/);
  assert.match(html, /type:'filter'/);
  assert.match(html, /id="headerRow"/);
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
  assert.match(html, /Clear selected cell: Delete\/Backspace/);
  assert.match(html, /edit\(cell,row,column,data,e\.key\)/);
  assert.match(html, /input\.setSelectionRange\(input\.value\.length,input\.value\.length\)/);
  assert.match(html, /finally\{window\.Worker=nativeWorker\}/);
  assert.doesNotMatch(html, /type="module"/);
  assert.match(html, /fetch\(window\.jinSheetAssets\.pptxWorker\)/);
  assert.match(html, /workerType:'classic'/);
  assert.match(html, /URL\.createObjectURL\(new Blob/);
  assert.match(html, /URL\.revokeObjectURL\(pptxWorkerBlobUrl\)/);
  assert.doesNotMatch(html, /https?:\/\//);
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

test('Webview 드래그 범위 선택·복사와 수식 입력줄 Enter가 편집 메시지를 보낸다', () => {
  const { document, documentListeners, elements, messages, windowListeners, respondToRange } = runGridWebview(renderHtml());
  let first = document.querySelector('[data-row="0"][data-column="0"]');
  let last = document.querySelector('[data-row="1"][data-column="1"]');
  first.onmousedown({ shiftKey: false });
  last.onmouseenter();
  documentListeners.get('mouseup')();
  last.onclick();
  documentListeners.get('keydown')({ target: last, key: 'c', ctrlKey: true, metaKey: false, preventDefault() {} });
  const copy = messages.at(-1);
  assert.equal(copy.type, 'copy');
  assert.deepEqual(JSON.parse(JSON.stringify(copy.range)), { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 });
  const keydown = documentListeners.get('keydown'), address = elements.get('address');
  assert.ok(first.closest('#grid'));
  assert.equal(first.tagName, 'DIV');
  keydown({ target: first, key: 'Home', preventDefault() {} });
  assert.equal(address.textContent, 'A2');
  respondToRange(); first = document.querySelector('[data-row="0"][data-column="0"]');
  let endHandled = false;
  keydown({ target: first, key: 'End', preventDefault() { endHandled = true; } });
  assert.equal(endHandled, true);
  assert.equal(address.textContent, 'B2');
  respondToRange(); first = document.querySelector('[data-row="0"][data-column="0"]');
  keydown({ target: first, key: 'End', ctrlKey: true, preventDefault() {} });
  assert.equal(address.textContent, 'B2');
  respondToRange(); first = document.querySelector('[data-row="0"][data-column="0"]');
  keydown({ target: first, key: 'Home', ctrlKey: true, preventDefault() {} });
  assert.equal(address.textContent, 'A1');
  respondToRange(); first = document.querySelector('[data-row="0"][data-column="0"]');
  keydown({ target: first, key: 'ArrowRight', ctrlKey: true, preventDefault() {} });
  assert.equal(messages.at(-1).type, 'navigate');
  assert.equal(messages.at(-1).direction, 'right');
  windowListeners.get('message')({ data: { type: 'navigation', requestId: messages.at(-1).requestId, row: 0, column: 1, extend: false } });
  assert.equal(address.textContent, 'B1');
  respondToRange();
  respondToRange(); first = document.querySelector('[data-row="0"][data-column="0"]'); first.onclick();
  keydown({ target: first, key: 'ArrowDown', shiftKey: true, preventDefault() {} });
  assert.equal(address.textContent, 'A2');
  respondToRange(); last = document.querySelector('[data-row="1"][data-column="0"]');
  keydown({ target: last, key: 'c', ctrlKey: true, metaKey: false, preventDefault() {} });
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1).range)), { startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 });
  const formula = elements.get('formula');
  formula.value = 'edited';
  formula.listeners.get('keydown')({ target: formula, key: 'Enter', preventDefault() {} });
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'edit', row: 1, column: 0, value: 'edited', sheetIndex: 0 });
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
  assert.equal(elements.get('address').textContent, 'B1');
  assert.equal(document.activeElement, cell);
  assert.equal(cell.attributes['aria-selected'], 'true');
  assert.equal(cell.tabIndex, 0);
  assert.equal(document.querySelector('[data-row="0"][data-column="0"]').tabIndex, -1);
  keydown({ target: cell, key: 'ArrowDown', preventDefault() {} });
  cell = document.querySelector('[data-row="1"][data-column="1"]');
  assert.equal(elements.get('address').textContent, 'B2');
  assert.equal(document.activeElement, cell);
  assert.equal(messages.filter(message => message.type === 'range').length, initialRangeRequests);
});

test('viewport를 넘는 키보드 이동도 새 범위를 그린 뒤 선택 셀 포커스를 복구한다', () => {
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
  assert.equal(elements.get('address').textContent, 'A15');
  assert.equal(messages.filter(message => message.type === 'range').length, initialRangeRequests + 1);
});

test('Webview 행·열 삽입·삭제 버튼이 선택한 위치의 구조 변경을 요청한다', () => {
  const { document, elements, messages, respondToRange } = runGridWebview(renderHtml());
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
});

test('Webview 열 머리글 정렬·선택 열 필터 UI와 필터 행 주소를 유지한다', () => {
  const { document, elements, messages, dispatch } = runGridWebview(renderHtml());
  const header = elements.get('canvas').children.find(child => child.classList.contains('colhead'));
  header.onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'sort', column: 0, mode: 'replace', direction: 'asc', hasHeader: true, sheetIndex: 0 });
  header.onclick({ shiftKey: true });
  assert.equal(messages.at(-1).mode, 'add');
  assert.equal(messages.at(-1).direction, 'asc');
  header.onclick({ ctrlKey: true });
  assert.equal(messages.at(-1).mode, 'remove');
  const selected = document.querySelector('[data-row="0"][data-column="1"]');
  selected.onclick();
  elements.get('filter').value = '12';
  elements.get('applyFilter').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'filter', column: 1, query: '12', hasHeader: true, sheetIndex: 0 });
  elements.get('clearView').onclick();
  assert.deepEqual(JSON.parse(JSON.stringify(messages.at(-1))), { type: 'clear-view', hasHeader: true, sheetIndex: 0 });

  dispatch({ type: 'model', kind: 'csv', readOnly: false, rowCount: 3, columnCount: 2, rowMap: [0, 2], filterRules: [{ column: 0, query: 'pear' }, { column: 1, query: '12' }], hasHeader: true, sheetIndex: 0, dirty: false });
  assert.equal(elements.get('filter').value, '12');
  const request = messages.filter(message => message.type === 'range').at(-1);
  dispatch({ type: 'range', requestId: request.requestId, rowStart: request.rowStart, rowIndexes: [0, 2], columnStart: request.columnStart, rowCount: 2, columnCount: 2, sheetIndex: 0, rows: [[{ value: 'header' }, { value: 'value' }], [{ value: 'target' }, { value: '12' }]] });
  const filteredHeader = elements.get('canvas').children.find(child => child.classList.contains('colhead') && child.textContent.startsWith('A'));
  assert.equal(filteredHeader.classList.contains('filtered'), true);
  document.querySelector('[data-row="1"][data-column="0"]').onclick();
  assert.equal(elements.get('address').textContent, 'A3');
  assert.equal(elements.get('filter').value, 'pear');
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
  assert.equal(header.style.width, '173px');
  assert.equal(cell.style.width, '173px');
  header.children.find(item => item.classList.contains('col-resize-handle')).onkeydown({ key: 'ArrowRight', preventDefault() {}, stopPropagation() {} });
  assert.deepEqual(JSON.parse(JSON.stringify(messages.find(message => message.type === 'column-width'))), { type: 'column-width', column: 0, width: 183, sheetIndex: 0 });
});

test('XLSX basic cell style travels through range data and reaches grid rendering', () => {
  const style = { color: '#123456', fill: '#FFEEAA', fontSize: 14, bold: true, italic: true, horizontal: 'center', vertical: 'bottom', borders: { left: { color: '#ABCDEF', line: 'thin' } } };
  const { elements } = runGridWebview(renderHtml(), { style });
  const cell = elements.get('canvas').children.find(item => item.classList.contains('cell') && item.dataset.row === '1' && item.dataset.column === '1');
  assert.equal(cell.style.color, '#123456');
  assert.equal(cell.style.backgroundColor, '#FFEEAA');
  assert.equal(cell.style.fontSize, '14pt');
  assert.equal(cell.style.fontWeight, '700');
  assert.equal(cell.style.fontStyle, 'italic');
  assert.equal(cell.style.textAlign, 'center');
  assert.equal(cell.style.paddingTop, '8px');
  assert.equal(cell.style.paddingBottom, '0px');
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
