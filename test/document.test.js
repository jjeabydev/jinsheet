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
const { findMatches, getRange, JinSheetDocument, renderHtml, viewModel } = require('../src/extension');
const { load } = require('../src/csv');
const { analyze } = require('../src/xlsx');

function minimalXlsx() {
  const xml = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="s1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="s1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    'xl/worksheets/sheet1.xml': `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>old</t></is></c></row></sheetData></worksheet>`
  };
  return zipSync(Object.fromEntries(Object.entries(xml).map(([name, value]) => [name, new TextEncoder().encode(value)])));
}
function uri(fsPath) { return { fsPath, scheme: 'file', toString: () => fsPath }; }

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

test('Webview HTML의 스크립트 문법을 검사하고 원격 출처를 허용하지 않는다', () => {
  const html = renderHtml();
  const script = /<script nonce="[^"]+">([\s\S]*?)<\/script>/.exec(html)?.[1];
  assert.ok(script);
  new vm.Script(script);
  assert.match(html, /default-src 'none'/);
  assert.doesNotMatch(html, /https?:\/\//);
});
