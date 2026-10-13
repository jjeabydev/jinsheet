'use strict';

const vscode = require('vscode');
const fs = require('node:fs/promises');
const crypto = require('node:crypto');
const path = require('node:path');
const { applyEdits, encode, load, parse: parseCsv, quote: quoteCsv, transformStructure: transformCsvStructure } = require('./csv');
const { LIMITS: XLSX_LIMITS, analyze, formatCellValue, patchCells, patchColumnWidths, transformStructure: transformXlsxStructure } = require('./xlsx');
const { analyzeOffice } = require('./office');
const { pdfCanvasOutputScale, validatePdfInput } = require('./pdf');
const { DEFAULT_COLUMN_WIDTH, ROW_HEADER_WIDTH, createColumnOffsets, visibleColumnRange, clampColumnWidth, excelColumnWidthPixels, estimateBestFitWidth, browserGridSource } = require('./grid');
const { languageOf, message, messagesFor, translateDiagnostic } = require('./i18n');

const CSV_LIMIT_BYTES = 10 * 1024 * 1024;
const CSV_LIMIT_ROWS = 10000;
const CSV_LIMIT_COLUMNS = 1000;
const MAX_CELL_CHARS = 100000;
const MAX_CLIPBOARD_CHARS = 1000000;

function hash(bytes) { return crypto.createHash('sha256').update(bytes).digest('hex'); }

class JinSheetDocument {
  constructor(uri, kind, state, bytes, writable, reason) {
    this.uri = uri;
    this.kind = kind;
    this.state = state;
    this.originalBytes = Buffer.from(bytes);
    this.originalFingerprint = hash(bytes);
    this.savedRevision = 0;
    this.baselineXlsxFiles = kind === 'xlsx' ? state.files : null;
    this.writable = writable && uri.scheme === 'file';
    this.reason = this.writable ? reason : (reason || '원격·가상 파일 시스템에서는 안전한 저장을 보장할 수 없어 읽기 전용으로 엽니다.');
    this.drafts = new Map();
    this.columnWidthDrafts = new Map();
    this.structureChanged = false;
    this.emitter = new vscode.EventEmitter();
    this.onDidChange = this.emitter.event;
    this.renderEmitter = new vscode.EventEmitter();
    this.onDidRender = this.renderEmitter.event;
    this.saving = null;
  }

  get dirty() { return this.structureChanged || this.drafts.size > 0 || this.columnWidthDrafts.size > 0; }

  getCell(row, column, sheetIndex = 0) {
    if (this.kind === 'csv') return this.state.rows[row]?.[column] ? { ...this.state.rows[row][column], row, column, value: this.state.rows[row][column].value } : null;
    const sheet = this.state.sheets[sheetIndex];
    const address = addressFor(row + 1, column + 1);
    const cell = sheet?.cells.get(address);
    const draft = this.drafts.get(sheetIndex + ':' + address);
    if (!cell) return { row, column, address, value: draft ?? '', kind: 'blank', formula: null };
    return { ...cell, value: draft ?? cell.value };
  }

  applyCellEdit(row, column, value, sheetIndex = 0) {
    this.applyCellEdits([{ row, column, value }], sheetIndex, '셀 값 편집');
  }

  applyCellEdits(edits, sheetIndex = 0, label = '셀 범위 붙여넣기') {
    if (!this.writable) throw new Error(this.reason || '읽기 전용 파일입니다.');
    if (!Array.isArray(edits) || edits.length < 1 || edits.length > 10000 || !Number.isInteger(sheetIndex) || sheetIndex < 0) throw new RangeError('셀 범위 수정 요청이 유효하지 않습니다.');
    if (this.kind === 'csv' && sheetIndex !== 0) throw new RangeError('CSV에는 첫 번째 시트만 있습니다.');
    if (this.kind === 'xlsx' && !this.state.sheets[sheetIndex]) throw new RangeError('수정할 시트를 찾을 수 없습니다.');
    const seen = new Set(), changes = [];
    for (const edit of edits) {
      const { row, column, value } = edit || {};
      if (!Number.isInteger(row) || !Number.isInteger(column) || row < 0 || column < 0 || typeof value !== 'string' || value.length > MAX_CELL_CHARS) throw new RangeError('셀 수정 요청이 유효하지 않습니다.');
      if (this.kind === 'xlsx' && (row >= XLSX_LIMITS.rows || column >= XLSX_LIMITS.columns)) throw new RangeError('XLSX 행·열 제한을 벗어난 셀입니다.');
      const key = row + ':' + column;
      if (seen.has(key)) throw new RangeError('같은 셀을 중복 수정할 수 없습니다.');
      seen.add(key);
      const cell = this.getCell(row, column, sheetIndex);
      if (!cell) throw new RangeError('붙여넣기 범위가 시트 크기를 벗어났습니다.');
      if (this.kind === 'xlsx' && (cell.formula || cell.kind === 'error')) throw new Error('수식·오류 셀이 포함된 범위는 수정할 수 없습니다.');
      if (cell.value !== value) changes.push({ row, column, value });
    }
    if (!changes.length) return;
    const before = this.captureSnapshot();
    for (const edit of changes) this.setDraft(edit.row, edit.column, edit.value, sheetIndex);
    const after = this.captureSnapshot();
    this.emitter.fire({
      label,
      undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); },
      redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); }
    });
    this.renderEmitter.fire();
  }

  applyColumnWidth(column, pixelWidth, sheetIndex = 0) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || 'XLSX 열 너비를 저장할 수 없습니다.');
    if (!Number.isInteger(sheetIndex) || !this.state.sheets[sheetIndex] || !Number.isInteger(column) || column < 0 || column >= XLSX_LIMITS.columns || !Number.isFinite(pixelWidth) || pixelWidth < 48 || pixelWidth > 640) throw new RangeError('열 너비 수정 요청이 유효하지 않습니다.');
    const key = sheetIndex + ':' + column;
    const baseline = this.state.sheets[sheetIndex].columnWidths[column];
    const pixels = clampColumnWidth(pixelWidth);
    const width = Math.round(((pixels - 5) / 7) * 100) / 100;
    const previous = this.columnWidthDrafts.get(key);
    const resetToDefault = baseline === undefined && Math.abs(pixels - DEFAULT_COLUMN_WIDTH) < 1;
    const resetToBaseline = baseline !== undefined && Math.abs(excelColumnWidthPixels(baseline) - pixels) < 1;
    if (previous === width && !resetToDefault && !resetToBaseline) return;
    if (previous === undefined && (resetToDefault || resetToBaseline)) return;
    const before = this.captureSnapshot();
    if (resetToDefault || resetToBaseline) this.columnWidthDrafts.delete(key);
    else this.columnWidthDrafts.set(key, width);
    const after = this.captureSnapshot();
    this.emitter.fire({
      label: '열 너비 변경',
      undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); },
      redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); }
    });
    this.renderEmitter.fire();
  }

  applyStructureChange(operation) {
    if (!this.writable) throw new Error(this.reason || '읽기 전용 파일입니다.');
    const { axis, action, index, sheetIndex = 0 } = operation || {};
    if (!['row', 'column'].includes(axis) || !['insert', 'delete'].includes(action) || !Number.isInteger(index) || index < 0 || !Number.isInteger(sheetIndex)) throw new RangeError('행·열 구조 변경 요청이 유효하지 않습니다.');
    const before = this.captureSnapshot();
    if (this.kind === 'csv') {
      if (sheetIndex !== 0) throw new RangeError('CSV에는 첫 번째 시트만 있습니다.');
      this.state = transformCsvStructure(this.state, path.extname(this.uri.fsPath).toLowerCase(), { axis, action, index });
      this.originalCsvState = this.state;
    } else if (this.kind === 'xlsx') {
      const edits = [...this.drafts].map(([key, value]) => {
        const [sheet, address] = key.split(':');
        return { sheetIndex: Number(sheet), address, value };
      });
      const current = edits.length ? patchCells(this.state, edits).model : this.state;
      this.state = transformXlsxStructure(current, { sheetIndex, axis, action, index }).model;
      if (axis === 'column') {
        const shiftedWidths = new Map();
        for (const [key, width] of this.columnWidthDrafts) {
          const [draftSheet, draftColumn] = key.split(':').map(Number);
          let nextColumn = draftColumn;
          if (draftSheet === sheetIndex && action === 'insert' && draftColumn >= index) nextColumn++;
          else if (draftSheet === sheetIndex && action === 'delete' && draftColumn === index) continue;
          else if (draftSheet === sheetIndex && action === 'delete' && draftColumn > index) nextColumn--;
          shiftedWidths.set(draftSheet + ':' + nextColumn, width);
        }
        this.columnWidthDrafts = shiftedWidths;
      }
    } else throw new Error('이 문서 형식은 행·열 구조를 편집할 수 없습니다.');
    this.drafts.clear();
    this.structureChanged = true;
    const after = this.captureSnapshot();
    this.emitter.fire({
      label: action === 'insert' ? '행·열 삽입' : '행·열 삭제',
      undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); },
      redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); }
    });
    this.renderEmitter.fire();
  }

  captureSnapshot() {
    return { state: this.state, originalCsvState: this.originalCsvState, drafts: new Map(this.drafts), columnWidthDrafts: new Map(this.columnWidthDrafts), structureChanged: this.structureChanged, savedRevision: this.savedRevision };
  }

  restoreSnapshot(snapshot) {
    this.state = snapshot.state;
    this.drafts = new Map(snapshot.drafts);
    this.columnWidthDrafts = new Map(snapshot.columnWidthDrafts || []);
    if (snapshot.savedRevision !== this.savedRevision) {
      if (this.kind === 'csv') {
        this.originalCsvState = this.state;
        this.drafts.clear();
      } else {
        const edits = [...this.drafts].map(([key, value]) => {
          const [sheet, address] = key.split(':');
          return { sheetIndex: Number(sheet), address, value };
        });
        if (edits.length) this.state = patchCells(this.state, edits).model;
        this.drafts.clear();
      }
      this.structureChanged = !this.matchesSavedState();
    } else {
      this.originalCsvState = snapshot.originalCsvState;
      this.structureChanged = snapshot.structureChanged;
    }
  }

  matchesSavedState() {
    if (this.kind === 'csv') return Buffer.compare(encode(this.state.text, this.state.encoding, this.state.bom), this.originalBytes) === 0;
    const current = this.state.files, baseline = this.baselineXlsxFiles;
    const names = Object.keys(baseline || {});
    return names.length === Object.keys(current || {}).length && names.every(name => current[name] && Buffer.compare(Buffer.from(current[name]), Buffer.from(baseline[name])) === 0);
  }

  setDraft(row, column, value, sheetIndex) {
    if (this.kind === 'csv') {
      const key = `0:${row}:${column}`;
      const baseline = this.baselineCsvValue(row, column);
      if (value === baseline) this.drafts.delete(key); else this.drafts.set(key, value);
      this.state = this.makeCsvState();
    } else {
      const address = addressFor(row + 1, column + 1);
      const key = `${sheetIndex}:${address}`;
      const baseline = this.state.sheets[sheetIndex].cells.get(address)?.value ?? '';
      if (value === baseline) this.drafts.delete(key); else this.drafts.set(key, value);
    }
  }

  baselineCsvValue(row, column) { return this.originalCsvState.rows[row]?.[column]?.value ?? ''; }

  makeCsvState() {
    const edits = [...this.drafts].map(([key, value]) => {
      const [, row, column] = key.split(':').map(Number);
      return { row, column, value };
    });
    const text = applyEdits(this.originalCsvState.text, this.originalCsvState.rows, edits, this.originalCsvState.delimiter);
    const state = load(encode(text, this.originalCsvState.encoding, this.originalCsvState.bom), path.extname(this.uri.fsPath), this.originalCsvState.delimiter);
    return state;
  }

  setInitialCsvState(state) { this.originalCsvState = state; this.state = state; }

  async save(cancellationToken) {
    if (!this.dirty) return;
    if (cancellationToken?.isCancellationRequested) throw new Error('저장이 취소되었습니다.');
    if (this.saving) return this.saving;
    this.saving = this.saveCurrent(cancellationToken);
    try { await this.saving; } finally { this.saving = null; }
  }

  async saveCurrent(cancellationToken) {
    const sourcePath = this.uri.fsPath;
    const current = await fs.readFile(sourcePath);
    if (hash(current) !== this.originalFingerprint) throw new Error('파일이 외부에서 변경되어 저장을 중단했습니다. 다시 열거나 다른 이름으로 저장하세요.');
    let output;
    if (this.kind === 'csv') {
      output = encode(this.state.text, this.state.encoding, this.state.bom);
      const verified = load(output, path.extname(sourcePath), this.state.delimiter);
      if (!sameRows(verified.rows, this.state.rows) || verified.delimiter !== this.state.delimiter) throw new Error('저장 전 CSV 재파싱 검증에 실패했습니다. 원본을 유지했습니다.');
    } else output = this.createOutput();
    if (cancellationToken?.isCancellationRequested) throw new Error('저장이 취소되었습니다.');
    const temporary = `${sourcePath}.jinsheet-${crypto.randomBytes(12).toString('hex')}.tmp`;
    let handle;
    try {
      const stat = await fs.stat(sourcePath);
      handle = await fs.open(temporary, 'wx', stat.mode & 0o777);
      await handle.writeFile(output);
      await handle.sync();
      await handle.close(); handle = null;
      if (cancellationToken?.isCancellationRequested) throw new Error('저장이 취소되었습니다.');
      const latest = await fs.readFile(sourcePath);
      if (hash(latest) !== this.originalFingerprint) throw new Error('저장 중 파일이 외부에서 변경되어 원본을 유지했습니다. 임시 저장본을 제거했습니다.');
      await fs.rename(temporary, sourcePath);
      await syncDirectory(path.dirname(sourcePath));
    } catch (error) {
      await handle?.close().catch(() => {});
      await fs.rm(temporary, { force: true }).catch(() => {});
      throw error;
    }
    this.originalBytes = Buffer.from(output);
    this.originalFingerprint = hash(output);
    if (this.kind === 'csv') this.originalCsvState = this.state;
    else { this.state = analyze(output, path.extname(sourcePath)); this.baselineXlsxFiles = this.state.files; }
    this.drafts.clear();
    this.columnWidthDrafts.clear();
    this.structureChanged = false;
    this.savedRevision++;
    this.renderEmitter.fire();
  }

  async saveAs(destination, cancellationToken) {
    if (cancellationToken?.isCancellationRequested) throw new Error('저장이 취소되었습니다.');
    if (destination.scheme !== 'file') throw new Error('현재는 로컬 파일로만 저장할 수 있습니다.');
    const destinationExt = path.extname(destination.fsPath).toLowerCase();
    if ((this.kind === 'csv' && !['.csv', '.tsv'].includes(destinationExt)) || (this.kind === 'xlsx' && destinationExt !== '.xlsx')) throw new Error('Save As는 원본의 파일 형식을 유지해야 합니다.');
    if (this.kind === 'csv' && ((destinationExt === '.tsv') !== (this.state.delimiter === '\t'))) throw new Error('Save As 파일 확장자와 구분자 형식이 일치해야 합니다.');
    const output = this.dirty ? this.createOutput() : this.originalBytes;
    const temp = `${destination.fsPath}.jinsheet-${crypto.randomBytes(12).toString('hex')}.tmp`;
    let handle;
    try {
      handle = await fs.open(temp, 'wx', 0o600);
      await handle.writeFile(output); await handle.sync(); await handle.close(); handle = null;
      if (cancellationToken?.isCancellationRequested) throw new Error('저장이 취소되었습니다.');
      await fs.rename(temp, destination.fsPath);
      await syncDirectory(path.dirname(destination.fsPath));
    } catch (error) {
      await handle?.close().catch(() => {}); await fs.rm(temp, { force: true }).catch(() => {}); throw error;
    }
    this.uri = destination; this.originalBytes = Buffer.from(output); this.originalFingerprint = hash(output);
    if (this.kind === 'csv') this.originalCsvState = this.state; else { this.state = analyze(output, path.extname(destination.fsPath)); this.baselineXlsxFiles = this.state.files; }
    this.drafts.clear();
    this.columnWidthDrafts.clear();
    this.structureChanged = false;
    this.savedRevision++;
    this.renderEmitter.fire();
  }

  createOutput() {
    if (this.kind === 'csv') return encode(this.state.text, this.state.encoding, this.state.bom);
    const patched = patchCells(this.state, [...this.drafts].map(([key, value]) => { const [sheet, address] = key.split(':'); return { sheetIndex: Number(sheet), address, value }; }));
    const widths = [...this.columnWidthDrafts].map(([key, width]) => { const [sheet, column] = key.split(':').map(Number); return { sheetIndex: sheet, column, width }; });
    return widths.length ? patchColumnWidths(patched.model, widths).bytes : patched.bytes;
  }

  async revert() {
    if (this.kind === 'csv') this.setInitialCsvState(load(this.originalBytes, path.extname(this.uri.fsPath)));
    else { this.state = analyze(this.originalBytes, path.extname(this.uri.fsPath)); this.baselineXlsxFiles = this.state.files; }
    this.drafts.clear();
    this.columnWidthDrafts.clear();
    this.structureChanged = false;
    this.savedRevision++;
    this.renderEmitter.fire();
  }

  async backup(context) {
    const backupDir = context.destination;
    const extension = path.extname(this.uri.fsPath) || '.data';
    const id = path.join(backupDir, `${crypto.randomBytes(12).toString('hex')}${extension}`);
    const bytes = this.dirty ? this.createOutput() : this.originalBytes;
    await fs.writeFile(id, bytes, { flag: 'wx', mode: 0o600 });
    return { id, delete: async () => fs.rm(id, { force: true }) };
  }

  dispose() { this.customDocumentSubscription?.dispose(); this.emitter.dispose(); this.renderEmitter.dispose(); }
}

function sameRows(left, right) {
  if (left.length !== right.length) return false;
  for (let r = 0; r < left.length; r++) {
    if (left[r].length !== right[r].length) return false;
    for (let c = 0; c < left[r].length; c++) if (left[r][c].value !== right[r][c].value) return false;
  }
  return true;
}

async function syncDirectory(directory) {
  try { const handle = await fs.open(directory, 'r'); try { await handle.sync(); } finally { await handle.close(); } }
  catch { /* Some platforms do not support syncing directory handles. */ }
}

function addressFor(row, column) {
  let letters = '';
  while (column > 0) { column--; letters = String.fromCharCode(65 + (column % 26)) + letters; column = Math.floor(column / 26); }
  return `${letters}${row}`;
}

function activate(context) {
  const provider = new JinSheetProvider(context);
  context.subscriptions.push(vscode.window.registerCustomEditorProvider('jinsheet.editor', provider, {
    supportsMultipleEditorsPerDocument: true,
    webviewOptions: { retainContextWhenHidden: true }
  }));
  context.subscriptions.push(vscode.commands.registerCommand('jinsheet.open', async uri => {
    const target = uri || vscode.window.activeTextEditor?.document.uri;
    if (target) await vscode.commands.executeCommand('vscode.openWith', target, 'jinsheet.editor');
  }));
  context.subscriptions.push(vscode.commands.registerCommand('jinsheet.editAsText', async () => {
    const activeInput = vscode.window.tabGroups.activeTabGroup.activeTab?.input;
    const uri = activeInput?.uri || vscode.window.activeTextEditor?.document.uri;
    if (uri) await vscode.commands.executeCommand('vscode.openWith', uri, 'default');
  }));
}

class JinSheetProvider {
  constructor(context) {
    this.context = context;
    this.language = languageOf(vscode.env.language);
    this.emitter = new vscode.EventEmitter();
    this.onDidChangeCustomDocument = this.emitter.event;
  }

  async openCustomDocument(uri) {
    const bytes = Buffer.from(await vscode.workspace.fs.readFile(uri));
    const extension = path.extname(uri.fsPath).toLowerCase();
    let document;
    if (extension === '.xlsx' || extension === '.xlsm') {
      let state, reason = '';
      try { state = analyze(bytes, extension); reason = state.readOnlyReason; }
      catch (error) { reason = error.message; state = { sheets: [], files: {}, unsupported: [reason], writable: false, readOnlyReason: reason }; }
      document = new JinSheetDocument(uri, 'xlsx', state, bytes, state.writable, reason);
    } else if (['.docx', '.docm', '.pptx', '.pptm'].includes(extension)) {
      let state, reason = 'Office 문서는 JinSheet 내부 렌더러로 표시하며 읽기 전용입니다.';
      try { state = analyzeOffice(bytes, extension); }
      catch (error) { reason = `문서 미리보기를 열 수 없습니다: ${error.message}`; state = { unsupported: true, format: 'office', title: 'Office 문서', sections: [] }; }
      document = new JinSheetDocument(uri, 'office', state, bytes, false, reason);
    } else if (extension === '.pdf') {
      let state, reason = 'PDF 문서는 내장 렌더러로 표시하며 읽기 전용입니다.';
      try { state = validatePdfInput(bytes); }
      catch (error) { reason = `PDF 미리보기를 열 수 없습니다: ${error.message}`; state = { unsupported: true, format: 'pdf', title: 'PDF 문서' }; }
      document = new JinSheetDocument(uri, 'pdf', state, bytes, false, reason);
    } else {
      let state, reason;
      try {
        if (!['.csv', '.tsv'].includes(extension)) throw new Error(extension === '.doc' || extension === '.ppt' ? '구형 .doc/.ppt 바이너리 형식은 지원하지 않습니다. .docx/.pptx 파일을 사용하세요.' : 'CSV, TSV와 XLSX만 지원합니다.');
        if (bytes.byteLength > CSV_LIMIT_BYTES) throw new Error('초기 버전은 10 MiB를 넘는 CSV/TSV 파일을 열지 않습니다.');
        state = load(bytes, extension);
        if (state.rows.length > CSV_LIMIT_ROWS) throw new Error('초기 버전은 10,000행을 넘는 CSV/TSV 파일을 편집하지 않습니다.');
        if (state.rows.some(row => row.length > CSV_LIMIT_COLUMNS)) throw new Error('초기 버전은 한 행이 1,000열을 넘는 CSV/TSV 파일을 편집하지 않습니다.');
      } catch (error) { reason = error.message; state = { unsupported: true }; }
      const writable = Boolean(state.rows);
      document = new JinSheetDocument(uri, 'csv', state, bytes, writable, reason || '');
      if (writable) document.setInitialCsvState(state);
    }
    const sub = document.onDidChange(change => this.emitter.fire({ document, label: translateDiagnostic(change.label, this.language), undo: change.undo, redo: change.redo }));
    document.customDocumentSubscription = sub;
    return document;
  }

  async resolveCustomEditor(document, panel) {
    let selectedSheet = 0;
    const postModel = () => {
      let model;
      try { model = viewModel(document, selectedSheet, this.language); }
      catch (error) {
        panel.webview.postMessage({ type: 'error', message: translateDiagnostic(error.message || error, this.language) });
        model = viewModel(document, selectedSheet, this.language);
      }
      panel.webview.postMessage(model);
    };
    const mediaRoot = vscode.Uri.joinPath(this.context.extensionUri, 'media');
    panel.webview.options = { enableScripts: true, localResourceRoots: [mediaRoot] };
    panel.webview.html = renderHtml(document, panel.webview, this.context.extensionUri, this.language);
    panel.webview.onDidReceiveMessage(async message => {
      if (!message || typeof message !== 'object' || Array.isArray(message)) return;
      try {
        const messageSheet = message.sheetIndex;
        const currentSheetMessage = matchesSelectedSheet(message, selectedSheet);
        if (message.type === 'edit' && currentSheetMessage && Number.isInteger(message.row) && Number.isInteger(message.column)) {
          document.applyCellEdit(message.row, message.column, message.value, messageSheet);
        } else if (message.type === 'column-width' && currentSheetMessage && Number.isInteger(message.column) && message.column >= 0 && message.column < XLSX_LIMITS.columns && Number.isFinite(message.width)) {
          document.applyColumnWidth(message.column, message.width, messageSheet);
        } else if (message.type === 'auto-fit-column' && currentSheetMessage && Number.isInteger(message.requestId) && message.requestId > 0 && Number.isInteger(message.column) && message.column >= 0 && message.column < (document.kind === 'csv' ? CSV_LIMIT_COLUMNS : XLSX_LIMITS.columns)) {
          const width = estimateDocumentColumnWidth(document, message.column, messageSheet, this.language);
          panel.webview.postMessage({ type: 'auto-fit-width', requestId: message.requestId, column: message.column, width, sheetIndex: messageSheet });
        } else if (message.type === 'structure' && currentSheetMessage && ['row', 'column'].includes(message.axis) && ['insert', 'delete'].includes(message.action) && Number.isInteger(message.index)) {
          document.applyStructureChange({ axis: message.axis, action: message.action, index: message.index, sheetIndex: messageSheet });
        } else if (message.type === 'navigate' && currentSheetMessage && ['up', 'down', 'left', 'right'].includes(message.direction)) {
          panel.webview.postMessage(navigateCell(document, message));
        } else if (message.type === 'range') {
          if (!currentSheetMessage) return;
          panel.webview.postMessage(getRange(document, message, this.language));
        } else if (message.type === 'selection-summary' && currentSheetMessage && Number.isInteger(message.requestId) && message.requestId > 0) {
          panel.webview.postMessage(selectionSummary(document, message));
        } else if (message.type === 'copy' && currentSheetMessage && Number.isInteger(message.row) && Number.isInteger(message.column)) {
          if (message.range) await vscode.env.clipboard.writeText(serializeClipboard(document, message.range, messageSheet));
          else {
            const cell = document.getCell(message.row, message.column, messageSheet);
            await vscode.env.clipboard.writeText(cell?.formula ? `=${cell.formula}` : (cell?.value ?? ''));
          }
        } else if (message.type === 'paste' && currentSheetMessage && Number.isInteger(message.row) && Number.isInteger(message.column)) {
          const targetSheet = messageSheet;
          const value = await vscode.env.clipboard.readText();
          if (selectedSheet !== targetSheet) return;
          const matrix = parseClipboard(value), edits = [];
          for (let row = 0; row < matrix.length; row++) {
            const destinationRow = message.row + row;
            for (let column = 0; column < matrix[row].length; column++) edits.push({ row: destinationRow, column: message.column + column, value: matrix[row][column] });
          }
          if (edits.length) document.applyCellEdits(edits, targetSheet);
        } else if (message.type === 'search' && currentSheetMessage && typeof message.query === 'string' && message.query.length <= 1000) {
          panel.webview.postMessage({ ...findMatches(document, message.query, messageSheet), requestId: message.requestId });
        } else if (message.type === 'ready' || message.type === 'refresh') {
          const requestedSheet = Number.isInteger(message.sheetIndex) ? message.sheetIndex : 0;
          if (!isValidSheetRequest(document, requestedSheet)) return;
          selectedSheet = requestedSheet;
          postModel();
        }
      } catch (error) { panel.webview.postMessage({ type: 'error', message: translateDiagnostic(error.message || error, this.language) }); }
    }, undefined, this.context.subscriptions);
    const sub = document.onDidRender(postModel);
    panel.onDidDispose(() => sub.dispose());
  }

  saveCustomDocument(document, token) { return document.save(token).catch(error => { throw new Error(translateDiagnostic(error.message || error, this.language)); }); }
  saveCustomDocumentAs(document, destination, token) { return document.saveAs(destination, token).catch(error => { throw new Error(translateDiagnostic(error.message || error, this.language)); }); }
  revertCustomDocument(document) { return document.revert().catch(error => { throw new Error(translateDiagnostic(error.message || error, this.language)); }); }
  backupCustomDocument(document, context) { return document.backup(context); }
  dispose() { this.emitter.dispose(); }
}

function matchesSelectedSheet(message, selectedSheet) {
  return Number.isInteger(message?.sheetIndex) && message.sheetIndex === selectedSheet;
}

function isValidSheetRequest(document, requestedSheet) {
  if (!Number.isInteger(requestedSheet) || requestedSheet < 0) return false;
  if (document.kind === 'csv') return requestedSheet === 0;
  if (document.kind === 'xlsx') return Boolean(document.state.sheets?.[requestedSheet]);
  return ['office', 'pdf'].includes(document.kind) && requestedSheet === 0;
}


function viewModel(document, sheetIndex = 0, language = 'en') {
  const localizedReason = translateDiagnostic(document.reason, language);
  if (document.kind === 'office') return { type: 'model', kind: 'office', readOnly: true, reason: localizedReason, dirty: false, title: document.state.title, format: path.extname(document.uri.fsPath).slice(1), supported: !document.state.unsupported, sourceBytes: Uint8Array.from(document.originalBytes) };
  if (document.kind === 'pdf') return { type: 'model', kind: 'pdf', readOnly: true, reason: localizedReason, dirty: false, title: document.state.title, format: 'pdf', supported: !document.state.unsupported, sourceBytes: Uint8Array.from(document.originalBytes) };
  if (document.kind === 'csv') {
    if (!document.state.rows) return { type: 'model', kind: 'csv', readOnly: true, reason: localizedReason, rows: [] };
    let columnCount = 0;
    for (const row of document.state.rows) columnCount = Math.max(columnCount, row.length);
    return { type: 'model', kind: 'csv', readOnly: !document.writable, reason: localizedReason, dirty: document.dirty, rowCount: document.state.rows.length, columnCount, sheetIndex: 0 };
  }
  const sheet = document.state.sheets[sheetIndex];
  if (!sheet) return { type: 'model', kind: 'xlsx', readOnly: true, reason: localizedReason || message('no-data', language), sheets: [] };
  let usedRows = 0, usedColumns = 0;
  for (const cell of sheet.cells.values()) { usedRows = Math.max(usedRows, cell.row + 1); usedColumns = Math.max(usedColumns, cell.column + 1); }
  for (const [key] of document.drafts) {
    const [draftSheet, address] = key.split(':');
    if (Number(draftSheet) !== sheetIndex) continue;
    const coordinate = parseAddress(address);
    usedRows = Math.max(usedRows, coordinate.row + 1);
    usedColumns = Math.max(usedColumns, coordinate.column + 1);
  }
  const columnWidths = { ...(sheet.columnWidths || {}) };
  for (const [key, width] of document.columnWidthDrafts || []) {
    const [draftSheet, column] = key.split(':').map(Number);
    if (draftSheet === sheetIndex) columnWidths[column] = width;
  }
  return {
    type: 'model', kind: 'xlsx', readOnly: !document.writable, reason: localizedReason, dirty: document.dirty,
    rowCount: usedRows, columnCount: usedColumns,
    sheetIndex, columnWidths, freezePane: sheet.freezePane || { rows: 0, columns: 0 }, sheets: document.state.sheets.map((item, index) => ({ name: item.name, index, visible: item.visible }))
  };
}

function estimateDocumentColumnWidth(document, column, sheetIndex = 0, language = 'en') {
  if (!Number.isInteger(column) || column < 0) throw new RangeError('열 너비 자동 맞춤 요청이 유효하지 않습니다.');
  const values = [];
  if (document.kind === 'csv') {
    if (sheetIndex !== 0) throw new RangeError('CSV에는 첫 번째 시트만 있습니다.');
    for (let row = 0; row < document.state.rows.length; row++) {
      const cell = document.getCell(row, column, 0);
      if (cell) values.push(cell.value);
    }
    return estimateBestFitWidth(values);
  }
  if (document.kind !== 'xlsx' || !document.state.sheets[sheetIndex]) throw new RangeError('열을 찾을 수 없습니다.');
  const sheet = document.state.sheets[sheetIndex], seen = new Set();
  for (const cell of sheet.cells.values()) {
    if (cell.column !== column) continue;
    const current = document.getCell(cell.row, column, sheetIndex);
    values.push(formatCellValue(current, language));
    seen.add(cell.address);
  }
  for (const [key] of document.drafts) {
    const [draftSheet, address] = key.split(':');
    if (Number(draftSheet) !== sheetIndex || seen.has(address)) continue;
    const coordinate = parseAddress(address);
    if (coordinate.column !== column) continue;
    values.push(formatCellValue(document.getCell(coordinate.row, column, sheetIndex), language));
  }
  return estimateBestFitWidth(values);
}

function getRange(document, message, language = 'en') {
  const { rowStart, columnStart, rowCount, columnCount, requestId } = message;
  if (![rowStart, columnStart, rowCount, columnCount, requestId].every(Number.isInteger) || rowStart < 0 || columnStart < 0 || rowCount < 1 || rowCount > 60 || columnCount < 1 || columnCount > 30) throw new RangeError('조회 범위가 유효하지 않습니다.');
  const sheetIndex = Number.isInteger(message.sheetIndex) ? message.sheetIndex : 0;
  const viewRows = message.viewRows ?? Array.from({ length: rowCount }, (_, index) => rowStart + index);
  const columnIndexes = message.columnIndexes ?? Array.from({ length: columnCount }, (_, index) => columnStart + index);
  const maximumRows = document.kind === 'csv' ? document.state.rows.length : XLSX_LIMITS.rows;
  const maximumColumns = document.kind === 'csv' ? CSV_LIMIT_COLUMNS : XLSX_LIMITS.columns;
  if (!Array.isArray(viewRows) || viewRows.length !== rowCount || viewRows.some((row, index) => !Number.isInteger(row) || row < 0 || row >= maximumRows || (index && row <= viewRows[index - 1]))) throw new RangeError('요청한 행 인덱스가 유효하지 않습니다.');
  if (!Array.isArray(columnIndexes) || columnIndexes.length !== columnCount || columnIndexes.some((column, index) => !Number.isInteger(column) || column < 0 || column >= maximumColumns || (index && column <= columnIndexes[index - 1]))) throw new RangeError('요청한 열 인덱스가 유효하지 않습니다.');
  const rowIndexes = viewRows;
  const rows = rowIndexes.map(row => columnIndexes.map(column => {
    const cell = document.getCell(row, column, sheetIndex);
    return cell ? { value: cell.value, displayValue: document.kind === 'xlsx' ? formatCellValue(cell, language) : cell.value, formula: cell.formula || null, cachedValue: cell.cachedValue ?? null, address: cell.address || null, kind: cell.kind || null, style: document.kind === 'xlsx' ? cell.styleProperties || null : null } : null;
  }));
  return { type: 'range', requestId, rowStart, viewRows, rowIndexes, columnStart, columnIndexes, rowCount, columnCount, sheetIndex, rows };
}

function selectionSummary(document, message) {
  const { startRow, endRow, startColumn, endColumn, requestId } = message;
  if (![startRow, endRow, startColumn, endColumn, requestId].every(Number.isInteger) ||
      Math.min(startRow, endRow, startColumn, endColumn) < 0 || requestId < 1) {
    throw new RangeError('선택 범위가 유효하지 않습니다.');
  }
  const top = Math.min(startRow, endRow), bottom = Math.max(startRow, endRow);
  const left = Math.min(startColumn, endColumn), right = Math.max(startColumn, endColumn);
  const cellCount = (bottom - top + 1) * (right - left + 1);
  if (!Number.isSafeInteger(cellCount) || cellCount > 10000) {
    return { type: 'selection-summary', requestId, tooLarge: true };
  }
  const sheetIndex = Number.isInteger(message.sheetIndex) ? message.sheetIndex : 0;
  let count = 0, numericCount = 0, sum = 0;
  for (let row = top; row <= bottom; row++) {
    for (let column = left; column <= right; column++) {
      const cell = document.getCell(row, column, sheetIndex);
      if (!cell || cell.value === '') continue;
      count++;
      const raw = String(cell.value).trim();
      if (!/^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(raw)) continue;
      const number = Number(raw);
      if (!Number.isFinite(number)) continue;
      numericCount++;
      sum += number;
    }
  }
  return {
    type: 'selection-summary', requestId, count, numericCount,
    sum: Number.isFinite(sum) ? sum : null,
    average: numericCount && Number.isFinite(sum) ? sum / numericCount : null
  };
}

function navigateCell(document, message) {
  const { row, column, direction, sheetIndex = 0, requestId } = message || {};
  if (![row, column, sheetIndex, requestId].every(Number.isInteger) || row < 0 || column < 0 || sheetIndex < 0 || !['up', 'down', 'left', 'right'].includes(direction)) throw new RangeError('키보드 이동 요청이 유효하지 않습니다.');
  if (document.kind === 'xlsx' && !document.state.sheets[sheetIndex]) throw new RangeError('이동할 시트를 찾을 수 없습니다.');
  if (document.kind === 'csv' && sheetIndex !== 0) throw new RangeError('CSV에는 첫 번째 시트만 있습니다.');
  if (document.kind === 'xlsx' && (row >= XLSX_LIMITS.rows || column >= XLSX_LIMITS.columns)) throw new RangeError('XLSX 행·열 제한을 벗어난 이동 요청입니다.');
  if (document.kind === 'csv' && (row >= document.state.rows.length || column >= document.state.rows.reduce((count, cells) => Math.max(count, cells.length), 0))) throw new RangeError('CSV 범위를 벗어난 이동 요청입니다.');
  const vertical = direction === 'up' || direction === 'down';
  const step = direction === 'up' || direction === 'left' ? -1 : 1;
  const current = vertical ? row : column;
  const edge = vertical
    ? (document.kind === 'xlsx' ? XLSX_LIMITS.rows - 1 : Math.max(0, document.state.rows.length - 1))
    : (document.kind === 'xlsx' ? XLSX_LIMITS.columns - 1 : Math.max(0, Math.max(0, ...document.state.rows.map(cells => cells.length)) - 1));
  const occupied = new Set();
  if (document.kind === 'csv') {
    if (vertical) {
      document.state.rows.forEach((cells, index) => { const cell = cells[column]; if (cell?.value) occupied.add(index); });
    } else {
      for (let index = 0; index < (document.state.rows[row]?.length || 0); index++) if (document.state.rows[row][index].value) occupied.add(index);
    }
  } else {
    for (const cell of document.state.sheets[sheetIndex].cells.values()) {
      if ((vertical ? cell.column : cell.row) !== (vertical ? column : row)) continue;
      const currentCell = document.getCell(cell.row, cell.column, sheetIndex);
      if (cell.formula || currentCell.value !== '') occupied.add(vertical ? cell.row : cell.column);
    }
    for (const [key, value] of document.drafts) {
      const [draftSheet, address] = key.split(':');
      if (Number(draftSheet) !== sheetIndex) continue;
      const coordinate = parseAddress(address);
      if ((vertical ? coordinate.column : coordinate.row) !== (vertical ? column : row)) continue;
      const position = vertical ? coordinate.row : coordinate.column;
      if (value === '') occupied.delete(position); else occupied.add(position);
    }
  }
  const sign = (position) => position * step;
  const positions = [...occupied].filter(position => sign(position) > sign(current)).sort((a, b) => sign(a) - sign(b));
  let target = current;
  if (occupied.has(current)) {
    while (target + step >= 0 && target + step <= edge && occupied.has(target + step)) target += step;
    const nextBlock = positions.find(position => sign(position) > sign(target));
    if (nextBlock !== undefined) {
      target = nextBlock;
      while (target + step >= 0 && target + step <= edge && occupied.has(target + step)) target += step;
    } else target = step > 0 ? edge : 0;
  } else if (positions.length) {
    target = positions[0];
    while (target + step >= 0 && target + step <= edge && occupied.has(target + step)) target += step;
  } else target = step > 0 ? edge : 0;
  return { type: 'navigation', requestId, row: vertical ? target : row, column: vertical ? column : target, extend: Boolean(message.extend) };
}

function parseAddress(address) {
  const match = /^([A-Z]+)([1-9]\d*)$/.exec(address);
  if (!match) throw new RangeError('셀 주소를 키보드 이동에 사용할 수 없습니다.');
  let column = 0;
  for (const letter of match[1]) column = column * 26 + letter.charCodeAt(0) - 64;
  return { row: Number(match[2]) - 1, column: column - 1 };
}

function selectionBounds(range) {
  const { startRow, endRow, startColumn, endColumn } = range || {};
  if (![startRow, endRow, startColumn, endColumn].every(Number.isInteger) || Math.min(startRow, endRow, startColumn, endColumn) < 0) throw new RangeError('복사 범위가 유효하지 않습니다.');
  const bounds = { top: Math.min(startRow, endRow), bottom: Math.max(startRow, endRow), left: Math.min(startColumn, endColumn), right: Math.max(startColumn, endColumn) };
  if ((bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1) > 10000) throw new RangeError('한 번에 복사할 수 있는 셀은 10,000개까지입니다.');
  return bounds;
}

function serializeClipboard(document, range, sheetIndex = 0) {
  const bounds = selectionBounds(range), rows = [];
  for (let row = bounds.top; row <= bounds.bottom; row++) {
    const values = [];
    for (let column = bounds.left; column <= bounds.right; column++) {
      const cell = document.getCell(row, column, sheetIndex);
      values.push(quoteCsv(cell?.formula ? '=' + cell.formula : (cell?.value ?? ''), '\t', false));
    }
    rows.push(values.join('\t'));
  }
  return rows.join('\r\n');
}

function parseClipboard(text) {
  if (typeof text !== 'string' || text.length > MAX_CLIPBOARD_CHARS) throw new RangeError('붙여넣기 내용은 1,000,000자까지 허용합니다.');
  if (!text) return [];
  const values = /[\t\r\n]/.test(text)
    ? parseCsv(text, '\t', 10000).map(row => row.map(cell => cell.value))
    : [[text]];
  if (values.some(row => row.some(value => value.length > MAX_CELL_CHARS))) throw new RangeError('붙여넣기 셀은 100,000자까지 허용합니다.');
  return values;
}

function findMatches(document, query, sheetIndex) {
  const needle = query.trim().toLocaleLowerCase();
  if (!needle) return { type: 'search', query, total: 0, matches: [] };
  const matches = []; let total = 0;
  const visit = (row, column, base) => {
    const value = document.kind === 'csv' ? (document.drafts.get(`0:${row}:${column}`) ?? base.value) : (document.drafts.get(`${sheetIndex}:${base.address}`) ?? base.value);
    if (String(value).toLocaleLowerCase().includes(needle)) {
      total++;
      if (matches.length < 100) matches.push({ row, column, address: base.address || addressFor(row + 1, column + 1) });
    }
  };
  if (document.kind === 'csv') document.state.rows?.forEach((cells, row) => cells.forEach((cell, column) => visit(row, column, cell)));
  else {
    const sheet = document.state.sheets[sheetIndex];
    for (const cell of sheet?.cells.values() || []) visit(cell.row, cell.column, cell);
    for (const [key] of document.drafts) {
      const [draftSheet, address] = key.split(':');
      if (Number(draftSheet) !== sheetIndex || sheet?.cells.has(address)) continue;
      const { row, column } = parseAddress(address);
      visit(row, column, { row, column, address, value: '' });
    }
  }
  matches.sort((left, right) => left.row - right.row || left.column - right.column);
  return { type: 'search', query, total, matches, sheetIndex };
}

function renderHtml(document, webview, extensionUri, language = 'en') {
  language = languageOf(language);
  const ui = JSON.stringify(messagesFor(language));
  const gridSource = browserGridSource();
  const pdfCanvasOutputScaleSource = pdfCanvasOutputScale.toString()
    .replaceAll('MAX_PDF_CANVAS_PIXELS', String(32 * 1024 * 1024))
    .replaceAll('MAX_PDF_CANVAS_DIMENSION', '16384');
  const nonce = crypto.randomBytes(18).toString('base64');
  const cspSource = webview?.cspSource || "'self'";
  const asset = name => webview && extensionUri ? webview.asWebviewUri(vscode.Uri.joinPath(extensionUri, 'media', name)).toString() : `media/${name}`;
  const pdfWorker = asset('pdf.worker.min.mjs');
  const pdfFonts = asset('pdf-standard-fonts/');
  const pptxWorker = asset('pptx.worker.js');
  const renderers = asset('renderers.js');
  const pdfLibrary = asset('pdf.min.js');
  const pptxCss = asset('pptx.css');
  const csp = `default-src 'none'; base-uri 'none'; form-action 'none'; connect-src ${cspSource}; img-src data: blob: ${cspSource}; font-src data: blob: ${cspSource}; style-src 'unsafe-inline' ${cspSource}; script-src 'nonce-${nonce}' ${cspSource}; worker-src blob: ${cspSource};`;
  return `<!doctype html><html lang="${language}"><head><meta charset="UTF-8"><meta http-equiv="Content-Security-Policy" content="${csp}"><meta name="viewport" content="width=device-width,initial-scale=1"><link rel="stylesheet" href="${pptxCss}"><style nonce="${nonce}">
  *{box-sizing:border-box}body{display:flex;flex-direction:column;height:100vh;margin:0;overflow:hidden;color:var(--vscode-foreground,#202124);background:var(--vscode-editor-background,#fff);font:13px var(--vscode-font-family)}#notice{padding:6px 10px;color:var(--vscode-descriptionForeground,#666);min-height:30px;background:var(--vscode-editor-background,#fff);border-bottom:1px solid var(--vscode-panel-border,#d6d6d6)}#toolbar{display:flex;flex-wrap:wrap;gap:6px;align-items:center;padding:6px 8px;border-bottom:1px solid var(--vscode-panel-border,#d6d6d6);background:var(--vscode-editorGroupHeader-tabsBackground,var(--vscode-sideBar-background,#f3f3f3));color:var(--vscode-foreground,#202124)}button{font:inherit;color:var(--vscode-button-secondaryForeground,var(--vscode-button-foreground,#202124));background:var(--vscode-button-secondaryBackground,var(--vscode-button-background,#ededed));border:1px solid var(--vscode-button-border,var(--vscode-panel-border,#d0d0d0));border-radius:2px;padding:4px 9px;min-height:26px;cursor:pointer}button:hover:not(:disabled){background:var(--vscode-button-secondaryHoverBackground,var(--vscode-button-hoverBackground,#ddd))}button:focus-visible,input:focus-visible{outline:1px solid var(--vscode-focusBorder,#007acc);outline-offset:1px}button:disabled{opacity:.55;cursor:default}button[aria-selected=true]{color:var(--vscode-tab-activeForeground,var(--vscode-foreground,#202124));background:var(--vscode-tab-activeBackground,var(--vscode-editor-background,#fff));border-color:var(--vscode-focusBorder,#007acc)}input{font:inherit;color:var(--vscode-input-foreground,#202124);background:var(--vscode-input-background,#fff);border:1px solid var(--vscode-input-border,var(--vscode-panel-border,#c8c8c8));border-radius:2px;padding:4px 7px;min-height:26px}input::placeholder{color:var(--vscode-input-placeholderForeground,#777)}#cellValueGroup,#searchGroup{display:flex;align-items:center;gap:6px}#cellValueGroup{flex:1;min-width:180px}#cellValueGroup label,#searchGroup label{color:var(--vscode-descriptionForeground,#666);white-space:nowrap}#formula{flex:1;min-width:120px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}#search{width:190px;height:26px}#address,#searchCount{color:var(--vscode-descriptionForeground,#666)}#address{min-width:42px}#searchCount{min-width:3em}#clearSearch{white-space:nowrap}#clearSearch[hidden]{display:none}#grid{position:relative;flex:1 1 auto;min-height:0;overflow:auto;outline:none;background:#fff;color:#202124}#grid,#readerCanvas{scrollbar-color:var(--vscode-scrollbarSlider-background,rgba(100,100,100,.4)) transparent;scrollbar-width:thin}#grid::-webkit-scrollbar,#readerCanvas::-webkit-scrollbar{width:12px;height:12px}#grid::-webkit-scrollbar-track,#readerCanvas::-webkit-scrollbar-track{background:transparent}#grid::-webkit-scrollbar-thumb,#readerCanvas::-webkit-scrollbar-thumb{background:var(--vscode-scrollbarSlider-background,rgba(100,100,100,.4));background-clip:padding-box;border:3px solid transparent;border-radius:6px}#grid::-webkit-scrollbar-thumb:hover,#readerCanvas::-webkit-scrollbar-thumb:hover{background:var(--vscode-scrollbarSlider-hoverBackground,rgba(100,100,100,.7));background-clip:padding-box}#grid::-webkit-scrollbar-thumb:active,#readerCanvas::-webkit-scrollbar-thumb:active{background:var(--vscode-scrollbarSlider-activeBackground,rgba(100,100,100,.8));background-clip:padding-box}#canvas{position:relative}.cell,.rowhead,.colhead{position:absolute;height:26px;padding:4px 6px;overflow:hidden;white-space:nowrap;text-overflow:ellipsis;border:1px solid var(--vscode-panel-border)}.cell{width:144px;background:#fff;color:#202124;border-color:#d6d6d6}.cell.numeric{text-align:right}.cell.in-range{box-shadow:inset 0 0 0 9999px rgba(0,100,220,.14)}.rowhead,.colhead{background:#f3f3f3;color:#444;text-align:center;border-color:#d6d6d6}.rowhead{width:52px;z-index:1}.colhead{top:0;z-index:1;user-select:none}.col-resize-handle{position:absolute;top:0;right:-3px;width:6px;height:100%;z-index:2;cursor:col-resize;touch-action:none}.col-resize-handle:focus{outline:1px solid var(--vscode-focusBorder)}.corner{position:absolute;width:52px;height:30px;top:0;left:0;background:#e8e8e8;border:1px solid #d6d6d6;z-index:2}.cell input.cell-editor{display:block;height:16px;min-height:16px;padding:0;border:0;border-radius:0;outline:1px solid var(--vscode-focusBorder,#007acc);background:#fff;color:#202124;font:inherit;line-height:16px}.cell.selected{outline:1px solid var(--vscode-focusBorder);outline-offset:-1px}.cell:focus{outline:1px solid var(--vscode-focusBorder);outline-offset:-1px}.cell.search-match{outline:2px solid var(--vscode-editor-findMatchBorder,#e8a317);outline-offset:-2px}.cell.search-current{outline:2px solid var(--vscode-editor-findMatchHighlightBorder,var(--vscode-focusBorder,#007acc));outline-offset:-2px}#sheets{display:flex;gap:2px;padding:4px 8px;border-bottom:1px solid var(--vscode-panel-border,#d6d6d6);background:var(--vscode-editorGroupHeader-tabsBackground,var(--vscode-sideBar-background,#f3f3f3));color:var(--vscode-foreground,#202124)}#sheets button{background:transparent;border-color:transparent;border-bottom:2px solid transparent}#sheets button[aria-selected=true]{background:var(--vscode-tab-activeBackground,var(--vscode-editor-background,#fff));border-color:var(--vscode-panel-border,#d6d6d6);border-bottom-color:var(--vscode-focusBorder,#007acc)}body.document-view{height:100vh;overflow:hidden}#reader{display:none;height:calc(100vh - 30px);overflow:hidden;background:var(--vscode-editor-background,#fff);color:var(--vscode-foreground,#202124);flex-direction:column}#readerToolbar{position:relative;z-index:3;display:flex;flex:0 0 auto;gap:6px;align-items:center;padding:6px 10px;background:var(--vscode-editorGroupHeader-tabsBackground,var(--vscode-sideBar-background,#f3f3f3));color:var(--vscode-foreground,#202124);border-bottom:1px solid var(--vscode-panel-border,#ddd)}#readerToolbar #pdfStatus{margin-left:8px;color:var(--vscode-descriptionForeground,#666);font-variant-numeric:tabular-nums}#readerCanvas{flex:1 1 auto;height:auto;min-height:0;padding:16px;overflow:auto;background:#f3f3f3;color:#202124}#readerCanvas .docx-wrapper{background:#f3f3f3!important;padding:16px 0 0!important}#readerCanvas section.docx{background:#fff!important;color:#202124!important;box-shadow:0 1px 4px #0002!important;outline:1px solid #dedede;margin-bottom:16px}#readerCanvas .flyfish-pptx-scale-box,#readerCanvas .flyfish-pptx-content,#readerCanvas .slide{background:#fff}#readerCanvas.pptx-view{overflow-x:hidden}#readerCanvas.pptx-view.pptx-zoomed,#readerCanvas.docx-view.docx-zoomed{overflow-x:auto}#readerCanvas.docx-view{overflow-x:hidden}#readerCanvas a{pointer-events:none}#pdfPage{display:block;max-width:none;height:auto;margin:0 auto;background:white;box-shadow:0 2px 8px #0005}#pdfStatus{min-width:9em}#readerStyles{display:none}
  </style><script nonce="${nonce}" src="${renderers}"></script><script nonce="${nonce}" src="${pdfLibrary}"></script><script nonce="${nonce}">window.pdfjsLib=globalThis.pdfjsLib;window.jinSheetAssets=${JSON.stringify({ pdfWorker, pdfFonts, pptxWorker })};</script></head><body><div id="notice" role="status" aria-live="polite"></div><div id="sheets" role="tablist"></div><div id="toolbar"><button id="insertRow" type="button" title="${message('insert-row-title', language)}">${message('insert-row', language)}</button><button id="deleteRow" type="button" title="${message('delete-row-title', language)}">${message('delete-row', language)}</button><button id="insertColumn" type="button" title="${message('insert-column-title', language)}">${message('insert-column', language)}</button><button id="deleteColumn" type="button" title="${message('delete-column-title', language)}">${message('delete-column', language)}</button><button id="autoFitColumn" type="button" title="${message('auto-fit-column-title', language)}">${message('auto-fit-column', language)}</button><span id="address"></span><div id="cellValueGroup"><label for="formula">${message('cell-value-label', language)}</label><input id="formula" type="text" aria-label="${message('cell-value', language)}" placeholder="${message('cell-value', language)}"></div><div id="searchGroup" role="search"><label for="search">${message('search-sheet-label', language)}</label><input id="search" type="search" aria-label="${message('search-sheet', language)}" placeholder="${message('find-hint', language)}"><button id="findButton" type="button" title="${message('find-button-title', language)}">${message('find-button', language)}</button><button id="clearSearch" type="button" title="${message('clear-search-title', language)}" aria-label="${message('clear-search-title', language)}" hidden>${message('clear-search', language)}</button><span id="searchCount" aria-live="polite"></span></div><span id="selectionSummary" role="status" aria-live="polite"></span></div><div id="grid" role="grid" aria-label="${message('spreadsheet', language)}" tabindex="0"><div id="canvas"></div></div><section id="reader" aria-label="${message('document-preview', language)}"><div id="readerToolbar"></div><div id="readerStyles"></div><div id="readerCanvas"></div></section><script nonce="${nonce}">
  const ui=${ui};${gridSource};const pdfCanvasOutputScale=${pdfCanvasOutputScaleSource};const t=key=>ui[key]||key;
  const vscode=acquireVsCodeApi();let model;let selected={row:0,column:0};let selectionAnchor={row:0,column:0};let selectedRange={startRow:0,endRow:0,startColumn:0,endColumn:0};let dragging=false;let dragged=false;let dragPointer=null;let dragMoved=false;let dragScrollTimer=null;let shiftSelecting=false;let navigationRequest=0;let rangeRequest=0;let renderedRange=null;let pendingRange=null;let matches=[];let matchedCells=new Set();let matchIndex=0;let searchedQuery='';let searchRequest=0;let scrollTimer;let renderGeneration=0;let renderedInstance=null;let pdfLoadingTask=null;let pdfDocument=null;let pdfPageNumber=1;let pdfZoom=1;let pdfPages=[];let pdfPageObserver=null;let pdfResizeObserver=null;let pdfBaseScale=1;let pdfViewportWidth=0;let pdfFitMode='width';const XLSX_MAX_ROWS=1048576;const XLSX_MAX_COLUMNS=16384;const RANGE_ROW_CHUNK=10;const RANGE_COLUMN_CHUNK=5;let pptxWorkerBlobUrl=null;let autoFitRequest=0;let docxUserZoom=100;let docxFitMode='width';let docxResizeObserver=null;let columnWidthsBySheet=new Map();let columnWidths=new Map();let columnOffsets=new Float64Array([0]);let columnResize=null;const label=n=>{let out='';while(n>0){n--;out=String.fromCharCode(65+n%26)+out;n=Math.floor(n/26)}return out};
  function releaseReader(){renderGeneration++;docxResizeObserver?.disconnect();docxResizeObserver=null;docxUserZoom=100;pdfPageObserver?.disconnect();pdfPageObserver=null;pdfResizeObserver?.disconnect();pdfResizeObserver=null;for(const page of pdfPages){page.token++;try{page.task?.cancel()}catch{}}pdfPages=[];try{renderedInstance?.destroy?.()}catch{}try{pdfLoadingTask?.destroy?.()}catch{}if(pptxWorkerBlobUrl){URL.revokeObjectURL(pptxWorkerBlobUrl);pptxWorkerBlobUrl=null}renderedInstance=null;pdfLoadingTask=null;pdfDocument=null;pdfViewportWidth=0;pdfFitMode='width';docxFitMode='width'}
  function readerError(error){const detail=error?.detail?' ('+error.detail+')':'';document.getElementById('notice').textContent=t('render-error')+String(error?.message||error)+detail}
  function documentBytes(m){const bytes=m.sourceBytes instanceof Uint8Array?m.sourceBytes:new Uint8Array(m.sourceBytes);return bytes.buffer.slice(bytes.byteOffset,bytes.byteOffset+bytes.byteLength)}
  function updatePdfStatus(pageNumber){const status=document.getElementById('pdfStatus');if(status)status.textContent=pageNumber+' / '+(pdfDocument?.numPages||0)}
  function cancelPdfPage(page){page.token++;page.rendering=false;try{page.task?.cancel()}catch{}page.task=null;page.canvas=null;if(page.shell?.isConnected)page.shell.replaceChildren()}
  async function renderPdfPage(info,generation){if(!pdfDocument||generation!==renderGeneration||info.task||info.rendering)return;let page;const token=++info.token;info.rendering=true;try{page=await pdfDocument.getPage(info.number);if(generation!==renderGeneration||token!==info.token)return;const base=page.getViewport({scale:1});const scale=Math.min(pdfBaseScale*pdfZoom,Math.sqrt(16000000/(base.width*base.height)));const viewport=page.getViewport({scale});const pageCanvas=document.createElement('canvas');pageCanvas.className='pdf-page-canvas';const outputScale=pdfCanvasOutputScale(viewport.width,viewport.height,window.devicePixelRatio||1);pageCanvas.width=Math.ceil(viewport.width*outputScale);pageCanvas.height=Math.ceil(viewport.height*outputScale);pageCanvas.style.width=viewport.width+'px';pageCanvas.style.height=viewport.height+'px';info.shell.style.width=viewport.width+'px';info.shell.style.height=viewport.height+'px';info.shell.replaceChildren(pageCanvas);info.canvas=pageCanvas;const task=page.render({canvasContext:pageCanvas.getContext('2d',{alpha:false}),viewport,transform:outputScale===1?null:[outputScale,0,0,outputScale,0,0]});info.task=task;try{await task.promise}catch(error){if(error?.name==='RenderingCancelledException')return;throw error}if(generation===renderGeneration&&token===info.token)updatePdfStatus(info.number)}catch(error){if(error?.name!=='RenderingCancelledException'&&generation===renderGeneration&&token===info.token)readerError(error)}finally{if(token===info.token){info.task=null;info.rendering=false;try{page?.cleanup?.()}catch{}}}}
  function renderVisiblePdfPages(generation){if(generation!==renderGeneration)return;const container=document.getElementById('readerCanvas'),root=container.getBoundingClientRect();let active=null,activeOverlap=0;for(const info of pdfPages){const bounds=info.shell.getBoundingClientRect();const overlap=Math.max(0,Math.min(bounds.bottom,root.bottom)-Math.max(bounds.top,root.top));if(overlap>0){renderPdfPage(info,generation);if(overlap>activeOverlap){active=info;activeOverlap=overlap}}}if(active){pdfPageNumber=active.number;updatePdfStatus(active.number)}}
  function layoutPdfPages(generation,anchorPage=pdfPageNumber){if(generation!==renderGeneration||!pdfPages.length)return;const container=document.getElementById('readerCanvas');pdfViewportWidth=container.clientWidth;const available=Math.max(120,pdfViewportWidth-48);const maxPageWidth=Math.max(...pdfPages.map(page=>page.width));pdfBaseScale=Math.min(1,available/maxPageWidth);if(pdfFitMode==='height'){const page=pdfPages[pdfPageNumber-1];if(page)pdfZoom=Math.max(.5,Math.min(3,Math.max(120,container.clientHeight-48)/(page.height*pdfBaseScale)))}for(const info of pdfPages){cancelPdfPage(info);const scale=Math.min(pdfBaseScale*pdfZoom,Math.sqrt(16000000/(info.width*info.height)));info.shell.style.width=Math.ceil(info.width*scale)+'px';info.shell.style.height=Math.ceil(info.height*scale)+'px'}pdfPageObserver?.disconnect();pdfPageObserver=new IntersectionObserver(entries=>{let active=null,activeOverlap=0;for(const entry of entries){const info=entry.target.__pdfPage;if(!info)continue;if(entry.isIntersecting)renderPdfPage(info,generation);else if(info.canvas||info.task)cancelPdfPage(info);const bounds=entry.boundingClientRect,root=document.getElementById('readerCanvas').getBoundingClientRect();const overlap=Math.max(0,Math.min(bounds.bottom,root.bottom)-Math.max(bounds.top,root.top));if(overlap>activeOverlap){active=info;activeOverlap=overlap}}if(active&&activeOverlap>0){pdfPageNumber=active.number;updatePdfStatus(active.number)}},{root:container,rootMargin:'900px 0px',threshold:0});for(const info of pdfPages)pdfPageObserver.observe(info.shell);const target=pdfPages[anchorPage-1];requestAnimationFrame(()=>{if(generation!==renderGeneration)return;if(target){const delta=target.shell.getBoundingClientRect().top-container.getBoundingClientRect().top-16;container.scrollTop=Math.max(0,container.scrollTop+delta)}requestAnimationFrame(()=>{if(generation!==renderGeneration)return;if(target)renderPdfPage(target,generation)})})}
  async function preparePdfPages(generation){const container=document.getElementById('readerCanvas');container.replaceChildren();pdfPages=[];for(let number=1;number<=pdfDocument.numPages;number++){if(generation!==renderGeneration)return;const page=await pdfDocument.getPage(number);const viewport=page.getViewport({scale:1});pdfPages.push({number,width:viewport.width,height:viewport.height,shell:null,canvas:null,task:null,rendering:false,token:0});try{page.cleanup()}catch{}const status=document.getElementById('pdfStatus');if(status)status.textContent=t('pdf-layout')+' '+number+' / '+pdfDocument.numPages}for(const info of pdfPages){const shell=document.createElement('div');shell.className='pdf-page-shell';shell.style.cssText='position:relative;margin:0 auto 16px;background:#fff;box-shadow:0 2px 8px #0005;';shell.setAttribute('role','group');shell.setAttribute('aria-label',t('page')+info.number);shell.__pdfPage=info;info.shell=shell;container.append(shell)}layoutPdfPages(generation,1);if(!pdfResizeObserver){pdfResizeObserver=new ResizeObserver(()=>{const width=container.clientWidth;if(Math.abs(width-pdfViewportWidth)>2)layoutPdfPages(generation,pdfPageNumber)});pdfResizeObserver.observe(container)}}
  function setPdfZoom(delta,generation){pdfFitMode=null;pdfZoom=Math.max(.5,Math.min(3,Math.round((pdfZoom+delta)*10)/10));layoutPdfPages(generation,pdfPageNumber)}
  async function openPdf(m,generation){if(!window.pdfjsLib)throw new Error(t('pdf-worker-missing'));const bytes=new Uint8Array(m.sourceBytes);window.pdfjsLib.GlobalWorkerOptions.workerSrc=window.jinSheetAssets.pdfWorker;const loadingStatus=document.getElementById('pdfStatus');const createTask=data=>{const task=window.pdfjsLib.getDocument({data,isEvalSupported:false,disableRange:true,disableStream:true,maxImageSize:16000000,enableXfa:false,standardFontDataUrl:window.jinSheetAssets.pdfFonts});if(loadingStatus)task.onProgress=({loaded,total})=>{if(loadingStatus.isConnected)loadingStatus.textContent=total?t('pdf-load')+Math.round(loaded/total*100)+'%':t('pdf-analyzing')};return task};const awaitTask=(task,timeoutMs,message,code)=>{let timerId;return Promise.race([task.promise,new Promise((_,reject)=>{timerId=setTimeout(()=>{const error=new Error(message);error.code=code;reject(error)},timeoutMs)})]).finally(()=>clearTimeout(timerId))};try{pdfLoadingTask=createTask(bytes);try{pdfDocument=await awaitTask(pdfLoadingTask,5000,t('pdf-worker-timeout'),'JINSHEET_PDF_WORKER_TIMEOUT')}catch(error){if(error.code!=='JINSHEET_PDF_WORKER_TIMEOUT')throw error;try{void pdfLoadingTask.destroy().catch(()=>{})}catch{}if(loadingStatus)loadingStatus.textContent=t('pdf-compat');const nativeWorker=window.Worker;window.Worker=undefined;try{pdfLoadingTask=createTask(new Uint8Array(m.sourceBytes));pdfDocument=await awaitTask(pdfLoadingTask,30000,t('pdf-fallback-timeout'),'JINSHEET_PDF_FALLBACK_TIMEOUT')}finally{window.Worker=nativeWorker}}}catch(error){try{await pdfLoadingTask?.destroy?.()}catch{}throw error}if(generation!==renderGeneration)return;if(pdfDocument.numPages<1||pdfDocument.numPages>1000)throw new Error(t('pdf-page-limit'));const bar=document.getElementById('readerToolbar');bar.replaceChildren();const minus=document.createElement('button');minus.textContent='−';minus.setAttribute('aria-label',t('zoom-out'));minus.onclick=()=>setPdfZoom(-.2,generation);const plus=document.createElement('button');plus.textContent='+';plus.setAttribute('aria-label',t('zoom-in'));plus.onclick=()=>setPdfZoom(.2,generation);const fit=document.createElement('button');fit.textContent=t('fit-width');fit.setAttribute('aria-label',t('fit-width'));fit.onclick=()=>{pdfFitMode='width';pdfZoom=1;layoutPdfPages(generation,pdfPageNumber)};const fitHeight=document.createElement('button');fitHeight.textContent=t('fit-height');fitHeight.setAttribute('aria-label',t('fit-height'));fitHeight.onclick=()=>{pdfFitMode='height';layoutPdfPages(generation,pdfPageNumber)};const status=document.createElement('span');status.id='pdfStatus';status.textContent=t('pdf-page-info');bar.append(minus,plus,fit,fitHeight,status);await preparePdfPages(generation);updatePdfStatus(1)}
  async function createPptxWorkerBlobUrl(){const response=await fetch(window.jinSheetAssets.pptxWorker);if(!response.ok)throw new Error(t('ppt-worker-read')+response.status);const source=await response.text();if(!source||source.length>2*1024*1024)throw new Error(t('ppt-worker-invalid'));const workerPrelude='self.setImmediate=self.setImmediate||function(callback,...args){return self.setTimeout(callback,0,...args)};self.clearImmediate=self.clearImmediate||function(handle){self.clearTimeout(handle)};';return URL.createObjectURL(new Blob([workerPrelude,source],{type:'text/javascript'}))}
  function setPptxZoom(percent,generation){if(generation!==renderGeneration||!renderedInstance)return;const zoom=Math.max(25,Math.min(300,percent));renderedInstance.setZoom(zoom);const canvas=document.getElementById('readerCanvas');canvas.classList.toggle('pptx-zoomed',zoom>100);if(zoom<=100)canvas.scrollLeft=0}
  function setDocxZoom(percent,generation){docxFitMode=docxFitMode||'custom';if(generation!==renderGeneration)return;docxUserZoom=Math.max(25,Math.min(300,percent));const canvas=document.getElementById('readerCanvas'),wrapper=canvas.querySelector('.docx-wrapper');if(!wrapper)return;wrapper.style.zoom='1';const pageWidths=[...wrapper.querySelectorAll('section.docx')].map(page=>page.getBoundingClientRect().width);const pageWidth=Math.max(1,...pageWidths);const fitScale=Math.min(1,Math.max(120,canvas.clientWidth-32)/pageWidth);wrapper.style.zoom=String(fitScale*docxUserZoom/100);canvas.classList.toggle('docx-zoomed',docxUserZoom>100);if(docxUserZoom<=100)canvas.scrollLeft=0}
  function fitDocxHeight(generation){const canvas=document.getElementById('readerCanvas'),wrapper=canvas.querySelector('.docx-wrapper');if(!wrapper)return;wrapper.style.zoom='1';const pages=[...wrapper.querySelectorAll('section.docx')];if(!pages.length)return;const pageHeight=Math.max(1,pages[0].getBoundingClientRect().height);const pageWidth=Math.max(1,...pages.map(page=>page.getBoundingClientRect().width));const fitScale=Math.min(1,Math.max(120,canvas.clientWidth-32)/pageWidth);docxFitMode='height';setDocxZoom(100*Math.max(120,canvas.clientHeight-32)/(pageHeight*fitScale),generation)}
  async function renderDocument(m){releaseReader();const generation=renderGeneration;const reader=document.getElementById('reader'),grid=document.getElementById('grid'),toolbar=document.getElementById('toolbar'),tabs=document.getElementById('sheets'),canvas=document.getElementById('readerCanvas'),styles=document.getElementById('readerStyles'),bar=document.getElementById('readerToolbar');reader.style.display='flex';document.body.classList.add('document-view');grid.style.display='none';toolbar.style.display='none';tabs.style.display='none';canvas.replaceChildren();styles.replaceChildren();bar.replaceChildren();canvas.style.padding='';canvas.style.maxWidth='';canvas.style.margin='';canvas.classList.remove('pptx-view','pptx-zoomed','docx-view','docx-zoomed');if(!m.supported||!m.sourceBytes){readerError(new Error(m.reason||t('not-previewable')));return}if(m.kind==='pdf'){const bar=document.getElementById('readerToolbar');bar.style.display='flex';const status=document.createElement('span');status.id='pdfStatus';status.textContent=t('pdf-loading');bar.append(status);try{await openPdf(m,generation)}catch(error){if(generation===renderGeneration){status.textContent=t('pdf-failed');readerError(error)}}return}if(m.format?.startsWith('ppt')){bar.style.display='flex';canvas.style.padding='0';canvas.classList.add('pptx-view');const pptxTarget=document.createElement('div');pptxTarget.style.cssText='width:100%;min-height:100%';canvas.replaceChildren(pptxTarget);try{const workerUrl=await createPptxWorkerBlobUrl();if(generation!==renderGeneration){URL.revokeObjectURL(workerUrl);return}pptxWorkerBlobUrl=workerUrl;renderedInstance=await window.jinSheetRenderers.PptxViewer.open(documentBytes(m),pptxTarget,{workerUrl,workerType:'classic',fitMode:'contain',lazySlides:true,zipLimits:{maxFileBytes:20*1024*1024},onError:readerError});if(generation!==renderGeneration){renderedInstance.destroy();return}const minus=document.createElement('button');minus.textContent='−';minus.setAttribute('aria-label',t('zoom-out'));minus.onclick=()=>setPptxZoom(renderedInstance.zoomPercent-10,generation);const plus=document.createElement('button');plus.textContent='+';plus.setAttribute('aria-label',t('zoom-in'));plus.onclick=()=>setPptxZoom(renderedInstance.zoomPercent+10,generation);const fit=document.createElement('button');fit.textContent=t('fit-width');fit.setAttribute('aria-label',t('fit-width'));fit.onclick=()=>{const width=renderedInstance.slideDimensions?.width;if(width>0)setPptxZoom(Math.min(100,canvas.clientWidth/width*100),generation)};const fitHeight=document.createElement('button');fitHeight.textContent=t('fit-height');fitHeight.setAttribute('aria-label',t('fit-height'));fitHeight.onclick=()=>{const height=renderedInstance.slideDimensions?.height;if(height>0)setPptxZoom(Math.min(100,canvas.clientHeight/height*100),generation)};bar.append(minus,plus,fit,fitHeight)}catch(error){if(generation===renderGeneration)readerError(error)}return}bar.style.display='flex';canvas.classList.add('docx-view');canvas.style.maxWidth='1100px';canvas.style.margin='0 auto';try{await window.jinSheetRenderers.renderDocxAsync(documentBytes(m),canvas,styles,{breakPages:true,renderHeaders:true,renderFooters:true,renderFootnotes:true,renderEndnotes:true,renderAltChunks:false,renderComments:false,renderChanges:false});for(const link of canvas.querySelectorAll('a')){link.removeAttribute('href');link.removeAttribute('target')}if(generation!==renderGeneration)return;const minus=document.createElement('button');minus.textContent='−';minus.setAttribute('aria-label',t('zoom-out'));minus.onclick=()=>{docxFitMode=null;setDocxZoom(docxUserZoom-10,generation)};const plus=document.createElement('button');plus.textContent='+';plus.setAttribute('aria-label',t('zoom-in'));plus.onclick=()=>{docxFitMode=null;setDocxZoom(docxUserZoom+10,generation)};const fit=document.createElement('button');fit.textContent=t('fit-width');fit.setAttribute('aria-label',t('fit-width'));fit.onclick=()=>{docxFitMode='width';setDocxZoom(100,generation)};const fitHeight=document.createElement('button');fitHeight.textContent=t('fit-height');fitHeight.setAttribute('aria-label',t('fit-height'));fitHeight.onclick=()=>fitDocxHeight(generation);bar.append(minus,plus,fit,fitHeight);docxFitMode='width';setDocxZoom(100,generation);docxResizeObserver=new ResizeObserver(()=>docxFitMode==='height'?fitDocxHeight(generation):setDocxZoom(docxUserZoom,generation));docxResizeObserver.observe(canvas)}catch(error){if(generation===renderGeneration)readerError(error)}}
  function draw(){if(!model)return;if(model.kind==='xlsx'){model.gridRowCount=XLSX_MAX_ROWS;model.gridColumnCount=XLSX_MAX_COLUMNS}else{model.gridRowCount=model.rowCount;model.gridColumnCount=model.columnCount}const activeSheet=model.sheetIndex||0;if(!columnWidthsBySheet.has(activeSheet)){const savedWidths=new Map(Object.entries(model.columnWidths||{}).map(([column,width])=>[Number(column),excelColumnWidthPixels(width)]));columnWidthsBySheet.set(activeSheet,savedWidths)}columnWidths=columnWidthsBySheet.get(activeSheet);columnOffsets=createColumnOffsets(model.gridColumnCount,columnWidths);const maxRow=Math.max(0,model.gridRowCount-1),maxColumn=Math.max(0,model.gridColumnCount-1);selected.row=Math.min(selected.row,maxRow);selected.column=Math.min(selected.column,maxColumn);selectionAnchor.row=Math.min(selectionAnchor.row,maxRow);selectionAnchor.column=Math.min(selectionAnchor.column,maxColumn);selectedRange.startRow=Math.min(selectedRange.startRow,maxRow);selectedRange.endRow=Math.min(selectedRange.endRow,maxRow);selectedRange.startColumn=Math.min(selectedRange.startColumn,maxColumn);selectedRange.endColumn=Math.min(selectedRange.endColumn,maxColumn);const notice=document.getElementById('notice');notice.textContent=(model.reason||'')+(model.dirty?' · '+t('modified'):'')+(model.readOnly?'':' · '+t('edit-shortcut'))+(model.rowCount>1000?' · '+t('scroll-rows'):'');const tabs=document.getElementById('sheets');tabs.replaceChildren();const grid=document.getElementById('grid'),toolbar=document.getElementById('toolbar'),canvas=document.getElementById('canvas');if(model.kind==='office'||model.kind==='pdf'){renderDocument(model);return}releaseReader();document.body.classList.remove('document-view');document.getElementById('reader').style.display='none';grid.style.display='';toolbar.style.display='';tabs.style.display='flex';for(const id of ['insertRow','deleteRow','insertColumn','deleteColumn'])document.getElementById(id).disabled=model.readOnly;canvas.style.position='relative';canvas.style.padding='0';canvas.style.maxWidth='';canvas.style.margin='';for(const sheet of model.sheets||[]){const button=document.createElement('button');button.textContent=sheet.name+(sheet.visible?'':' ('+t('hidden')+')');button.setAttribute('role','tab');button.setAttribute('aria-selected',String(sheet.index===model.sheetIndex));button.onclick=()=>vscode.postMessage({type:'refresh',sheetIndex:sheet.index});tabs.append(button)}rangeRequest++;renderedRange=null;pendingRange=null;canvas.style.width=(ROW_HEADER_WIDTH+columnOffsets[model.gridColumnCount])+'px';canvas.style.height=(30+model.gridRowCount*26)+'px';canvas.replaceChildren();if(!model.gridRowCount||!model.gridColumnCount){canvas.textContent=model.reason||t('no-data');return}const corner=document.createElement('div');corner.className='corner';canvas.append(corner);requestRange()}
  for(const [id,axis,action] of [['insertRow','row','insert'],['deleteRow','row','delete'],['insertColumn','column','insert'],['deleteColumn','column','delete']])document.getElementById(id).onclick=()=>{if(model?.readOnly)return;vscode.postMessage({type:'structure',axis,action,index:axis==='row'?selected.row:selected.column,sheetIndex:model?.sheetIndex||0})};
  document.getElementById('autoFitColumn').onclick=()=>{if(!model)return;const requestId=++autoFitRequest;vscode.postMessage({type:'auto-fit-column',requestId,column:selected.column,sheetIndex:model.sheetIndex||0})};
  function requestRange(){if(!model||!model.gridRowCount||!model.gridColumnCount)return;const grid=document.getElementById('grid');const visibleRow=Math.floor(grid.scrollTop/26);const rowStart=Math.floor(visibleRow/RANGE_ROW_CHUNK)*RANGE_ROW_CHUNK;const {columnStart,columnCount}=visibleColumnRange(columnOffsets,grid.scrollLeft,grid.clientWidth);const freezeRows=Math.max(0,Math.min(50,model.freezePane?.rows||0));const freezeColumns=Math.max(0,Math.min(10,model.freezePane?.columns||0));const rowCount=Math.min(60-freezeRows,Math.max(1,Math.ceil(grid.clientHeight/26)+RANGE_ROW_CHUNK),model.gridRowCount-rowStart);const {columnStart:visibleColumnStart,columnCount:visibleColumnCount}=visibleColumnRange(columnOffsets,grid.scrollLeft,grid.clientWidth,{maxColumns:Math.max(1,30-freezeColumns)});if(rowCount<1||visibleColumnCount<1)return;let range={rowStart,columnStart:visibleColumnStart,rowCount,columnCount:visibleColumnCount,sheetIndex:model.sheetIndex};if(freezeRows||freezeColumns){const viewRows=[...new Set([...Array.from({length:freezeRows},(_,index)=>index),...Array.from({length:rowCount},(_,index)=>rowStart+index)])].sort((a,b)=>a-b);const columnIndexes=[...new Set([...Array.from({length:freezeColumns},(_,index)=>index),...Array.from({length:visibleColumnCount},(_,index)=>visibleColumnStart+index)])].sort((a,b)=>a-b);if(viewRows.length>60||columnIndexes.length>30)return;range={...range,rowStart:viewRows[0]||0,columnStart:columnIndexes[0]||0,rowCount:viewRows.length,columnCount:columnIndexes.length,viewRows,columnIndexes}}else range={...range,columnStart,rowCount:Math.min(60,Math.max(1,Math.ceil(grid.clientHeight/26)+RANGE_ROW_CHUNK),model.gridRowCount-rowStart)};const same=(left,right)=>left&&right&&Object.keys(right).every(key=>key==='viewRows'||key==='columnIndexes'?(!right[key]||JSON.stringify(left[key])===JSON.stringify(right[key])):left[key]===right[key]);if(same(pendingRange,range))return;if(same(renderedRange,range)){if(pendingRange){rangeRequest++;pendingRange=null}return}const requestId=++rangeRequest;pendingRange={...range,requestId};vscode.postMessage({type:'range',requestId,...range})}
  function updateColumnGeometry(){if(!model)return;const canvas=document.getElementById('canvas');columnOffsets=createColumnOffsets(model.gridColumnCount,columnWidths);canvas.style.width=(ROW_HEADER_WIDTH+columnOffsets[model.gridColumnCount])+'px';for(const item of canvas.children){if(!item.classList.contains('colhead')&&!item.classList.contains('cell'))continue;const column=Number(item.dataset.column);const width=clampColumnWidth(columnWidths.get(column)??DEFAULT_COLUMN_WIDTH);item.style.left=(ROW_HEADER_WIDTH+columnOffsets[column])+'px';item.style.width=width+'px';const handle=[...item.children].find(child=>child.classList.contains('col-resize-handle'));handle?.setAttribute('aria-valuenow',String(width))}}function resizeColumn(column,width){const value=clampColumnWidth(width);columnWidths.set(column,value);updateColumnGeometry()}function commitColumnWidth(column){if(model?.kind==='xlsx'&&!model.readOnly)vscode.postMessage({type:'column-width',column,width:clampColumnWidth(columnWidths.get(column)??DEFAULT_COLUMN_WIDTH),sheetIndex:model.sheetIndex})}function renderRange(range){if(range.requestId!==rangeRequest||range.sheetIndex!==model.sheetIndex)return;renderedRange={rowStart:range.rowStart,columnStart:range.columnStart,rowCount:range.rowCount,columnCount:range.columnCount,sheetIndex:range.sheetIndex,viewRows:range.viewRows?[...range.viewRows]:undefined,columnIndexes:range.columnIndexes?[...range.columnIndexes]:undefined};pendingRange=null;const canvas=document.getElementById('canvas');for(const item of [...canvas.children])if(!item.classList.contains('corner'))item.remove();for(let c=0;c<range.columnCount;c++){const column=range.columnIndexes?.[c]??(range.columnStart+c);const head=document.createElement('div');head.className='colhead';head.dataset.column=String(column);const columnWidth=clampColumnWidth(columnWidths.get(column)??DEFAULT_COLUMN_WIDTH);head.style.width=columnWidth+'px';head.textContent=label(column+1);head.style.left=(ROW_HEADER_WIDTH+columnOffsets[column])+'px';head.setAttribute('role','columnheader');head.setAttribute('aria-label',t('aria-column')+label(column+1));const resizeHandle=document.createElement('div');resizeHandle.className='col-resize-handle';resizeHandle.setAttribute('role','separator');resizeHandle.setAttribute('aria-orientation','vertical');resizeHandle.setAttribute('aria-label',t('resize-column')+label(column+1));resizeHandle.setAttribute('aria-valuemin','48');resizeHandle.setAttribute('aria-valuemax','640');resizeHandle.setAttribute('aria-valuenow',String(columnWidth));resizeHandle.setAttribute('tabindex','0');resizeHandle.onmousedown=e=>{e.preventDefault();e.stopPropagation();columnResize={column,startX:e.clientX,startWidth:columnWidth}};resizeHandle.onclick=e=>e.stopPropagation();resizeHandle.onkeydown=e=>{if(e.key==='ArrowLeft'||e.key==='ArrowRight'){e.preventDefault();e.stopPropagation();resizeColumn(column,(columnWidths.get(column)??columnWidth)+(e.key==='ArrowRight'?10:-10));commitColumnWidth(column);requestRange();resizeHandle.setAttribute('aria-valuenow',String(clampColumnWidth(columnWidths.get(column))))}};head.append(resizeHandle);canvas.append(head)}for(let r=0;r<range.rowCount;r++){const row=range.viewRows?.[r]??(range.rowStart+r);const rowHead=document.createElement('div');rowHead.className='rowhead';rowHead.textContent=String(row+1);rowHead.style.top=(30+row*26)+'px';rowHead.dataset.viewRow=String(row);rowHead.style.left='0px';rowHead.setAttribute('aria-label',t('aria-row')+(row+1));canvas.append(rowHead);for(let c=0;c<range.columnCount;c++){const column=range.columnIndexes?.[c]??(range.columnStart+c);const data=range.rows[r][c];const inRange=row>=Math.min(selectedRange.startRow,selectedRange.endRow)&&row<=Math.max(selectedRange.startRow,selectedRange.endRow)&&column>=Math.min(selectedRange.startColumn,selectedRange.endColumn)&&column<=Math.max(selectedRange.startColumn,selectedRange.endColumn);const cell=document.createElement('div');cell.className='cell'+(selected.row===row&&selected.column===column?' selected':'')+(inRange?' in-range':'');cell.classList.toggle('numeric',data?.kind==='number');const style=data?.style;if(style){if(style.color)cell.style.color=style.color;if(style.fill)cell.style.backgroundColor=style.fill;if(style.fontSize)cell.style.fontSize=style.fontSize+'pt';if(style.bold)cell.style.fontWeight='700';if(style.italic)cell.style.fontStyle='italic';if(style.underline||style.strike)cell.style.textDecoration=[style.underline?'underline':'',style.strike?'line-through':''].filter(Boolean).join(' ');if(style.horizontal)cell.style.textAlign=style.horizontal;if(style.vertical){cell.style.paddingTop=style.vertical==='top'?'0px':style.vertical==='bottom'?'8px':'4px';cell.style.paddingBottom=style.vertical==='bottom'?'0px':style.vertical==='top'?'8px':'4px'}if(style.borders)for(const [side,border] of Object.entries(style.borders))cell.style['border'+side[0].toUpperCase()+side.slice(1)]='1px '+(border.line==='double'?'double':border.line==='dashed'?'dashed':border.line==='dotted'?'dotted':'solid')+' '+border.color}cell.id='cell-'+row+'-'+column;cell.tabIndex=selected.row===row&&selected.column===column?0:-1;cell.setAttribute('role','gridcell');cell.setAttribute('aria-selected',String(inRange));cell.setAttribute('aria-rowindex',String(row+1));cell.setAttribute('aria-colindex',String(column+1));cell.dataset.row=String(row);cell.dataset.column=String(column);cell.style.top=(30+row*26)+'px';cell.style.left=(ROW_HEADER_WIDTH+columnOffsets[column])+'px';cell.style.width=clampColumnWidth(columnWidths.get(column)??DEFAULT_COLUMN_WIDTH)+'px';const frozenRow=row<(model.freezePane?.rows||0),frozenColumn=column<(model.freezePane?.columns||0);if(frozenRow||frozenColumn)cell.style.zIndex=frozenRow&&frozenColumn?'3':'2';cell.textContent=data?.displayValue??data?.value??'';cell.title=cell.textContent;cell.__jinSheetData=data;cell.onmousedown=e=>{e.preventDefault();dragging=true;dragged=false;dragPointer={x:e.clientX,y:e.clientY};dragMoved=false;shiftSelecting=e.shiftKey;select(row,column,data,e.shiftKey);startCellDrag()};cell.onmouseenter=()=>{if(dragging){dragged=true;select(row,column,data,true)}};cell.onclick=()=>{if(dragged){dragged=false;shiftSelecting=false;return}if(shiftSelecting){select(row,column,data,true);shiftSelecting=false}else select(row,column,data)};cell.ondblclick=()=>edit(cell,row,column,data);cell.onkeydown=e=>{if(e.target&&e.target!==cell)return;if(e.key==='F2'||e.key==='Enter'){e.preventDefault();e.stopPropagation();edit(cell,row,column,data)}else if((e.isComposing||e.key==='Process'||e.key==='Unidentified')&&!e.ctrlKey&&!e.metaKey&&!e.altKey){edit(cell,row,column,data,'')}else if(e.key.length===1&&!e.ctrlKey&&!e.metaKey&&!e.altKey){e.preventDefault();e.stopPropagation();edit(cell,row,column,data,e.key)}};canvas.append(cell)}}const selectedRowIndex=(range.viewRows||Array.from({length:range.rowCount},(_,index)=>range.rowStart+index)).indexOf(selected.row),selectedColumnIndex=(range.columnIndexes||Array.from({length:range.columnCount},(_,index)=>range.columnStart+index)).indexOf(selected.column);if(selectedRowIndex>=0&&selectedColumnIndex>=0)select(selected.row,selected.column,range.rows[selectedRowIndex]?.[selectedColumnIndex],true,document.activeElement!==document.getElementById('search'));refreshSearchHighlights();syncHeadings()}
  function syncHeadings(){const grid=document.getElementById('grid'),canvas=document.getElementById('canvas'),freezeRows=model?.freezePane?.rows||0,freezeColumns=model?.freezePane?.columns||0;for(const item of canvas.children){if(item.classList.contains('colhead'))item.style.transform='translateY('+grid.scrollTop+'px)';else if(item.classList.contains('rowhead')){const frozen=Number(item.dataset.viewRow)<freezeRows;item.style.transform=frozen?'translate('+grid.scrollLeft+'px,'+grid.scrollTop+'px)':'translateX('+grid.scrollLeft+'px)';if(frozen)item.style.zIndex='3'}else if(item.classList.contains('cell')){const frozenRow=Number(item.dataset.row)<freezeRows,frozenColumn=Number(item.dataset.column)<freezeColumns;if(frozenRow||frozenColumn)item.style.transform='translate('+(frozenColumn?grid.scrollLeft:0)+'px,'+(frozenRow?grid.scrollTop:0)+'px)';else item.style.transform=''}else if(item.classList.contains('corner'))item.style.transform='translate('+grid.scrollLeft+'px,'+grid.scrollTop+'px)'}}
  let summaryRequest=0,summaryTimer;
  function refreshSelectionSummary(){clearTimeout(summaryTimer);const requestId=++summaryRequest;summaryTimer=setTimeout(()=>vscode.postMessage({type:'selection-summary',requestId,startRow:selectedRange.startRow,endRow:selectedRange.endRow,startColumn:selectedRange.startColumn,endColumn:selectedRange.endColumn,sheetIndex:model?.sheetIndex||0}),80)}
  function select(r,c,data,extend=false,focusCell=true){if(!extend)navigationRequest++;selected={row:r,column:c};if(!extend){selectionAnchor={row:r,column:c};selectedRange={startRow:r,endRow:r,startColumn:c,endColumn:c}}else selectedRange={startRow:selectionAnchor.row,endRow:r,startColumn:selectionAnchor.column,endColumn:c};document.querySelector('.cell.selected')?.classList.remove('selected');document.querySelector('[data-row="'+r+'"][data-column="'+c+'"]')?.classList.add('selected');for(const cell of document.querySelectorAll('.cell')){const row=Number(cell.dataset.row),column=Number(cell.dataset.column),inRange=row>=Math.min(selectedRange.startRow,selectedRange.endRow)&&row<=Math.max(selectedRange.startRow,selectedRange.endRow)&&column>=Math.min(selectedRange.startColumn,selectedRange.endColumn)&&column<=Math.max(selectedRange.startColumn,selectedRange.endColumn);cell.classList.toggle('in-range',inRange);cell.setAttribute('aria-selected',String(inRange));cell.tabIndex=row===r&&column===c?0:-1}if(focusCell)if(focusCell)document.querySelector('[data-row="'+r+'"][data-column="'+c+'"]')?.focus({preventScroll:true});document.getElementById('address').textContent=label(c+1)+(r+1);const formula=document.getElementById('formula');formula.value=data?.formula?('='+data.formula+'  · '+t('cache-value')+': '+(data.cachedValue??'')):(data?.value??'');formula.title=formula.value;formula.disabled=Boolean(model.readOnly||data?.formula||data?.kind==='error');refreshSelectionSummary()}
  function edit(cell,r,c,data,initial){if(!cell||model.readOnly||data?.formula||data?.kind==='error')return;const input=document.createElement('input');input.className='cell-editor';input.value=initial??data?.value??'';input.style.width=Math.max(24,clampColumnWidth(columnWidths.get(c)??DEFAULT_COLUMN_WIDTH)-12)+'px';cell.replaceChildren(input);input.focus();if(initial===undefined)input.select();else input.setSelectionRange(input.value.length,input.value.length);let done=false,committedChange=false;const commit=()=>{if(done)return committedChange;done=true;committedChange=input.value!==(data?.value??'');if(committedChange)vscode.postMessage({type:'edit',row:r,column:c,value:input.value,sheetIndex:model.sheetIndex});return committedChange};input.onkeydown=e=>{if(e.key==='Tab'){e.preventDefault();e.stopPropagation();commit();moveAfterCommit(r,c,e.shiftKey);input.blur()}else if(e.key==='Enter'){if(e.isComposing||e.keyCode===229)return;e.preventDefault();commit();input.blur();scrollToCell({row:Math.min((model.gridRowCount||model.rowCount)-1,r+1),column:c})}else if(e.key==='Escape'){e.preventDefault();done=true;renderedRange=null;pendingRange=null;input.blur();requestRange()}};input.onblur=()=>{if(!commit()){renderedRange=null;pendingRange=null;requestRange()}}}
  function scrollToCell(cell,extend=false,focusCell=true){select(cell.row,cell.column,undefined,extend,focusCell);const grid=document.getElementById('grid');const freezeRows=model?.freezePane?.rows||0,freezeColumns=model?.freezePane?.columns||0;if(cell.row>=freezeRows){const rowTop=30+cell.row*26,rowBottom=rowTop+26,visibleTop=30+freezeRows*26;if(rowTop-grid.scrollTop<visibleTop)grid.scrollTop=Math.max(0,rowTop-visibleTop);else if(rowBottom-grid.scrollTop>grid.clientHeight)grid.scrollTop=rowBottom-grid.clientHeight}if(cell.column>=freezeColumns){let frozenWidth=0;for(let column=0;column<freezeColumns;column++)frozenWidth+=clampColumnWidth(columnWidths.get(column)??DEFAULT_COLUMN_WIDTH);const cellLeft=52+columnOffsets[cell.column],cellRight=cellLeft+clampColumnWidth(columnWidths.get(cell.column)??DEFAULT_COLUMN_WIDTH),visibleLeft=52+frozenWidth;if(cellLeft-grid.scrollLeft<visibleLeft)grid.scrollLeft=Math.max(0,columnOffsets[cell.column]-frozenWidth);else if(cellRight-grid.scrollLeft>grid.clientWidth)grid.scrollLeft=cellRight-grid.clientWidth}requestRange()}function refreshSearchHighlights(){const current=matches[matchIndex];for(const cell of document.querySelectorAll('.cell')){const row=Number(cell.dataset.row),column=Number(cell.dataset.column),isMatch=matchedCells.has(row+':'+column);cell.classList.toggle('search-match',isMatch);cell.classList.toggle('search-current',isMatch&&current?.row===row&&current?.column===column)}}function runSearch(){const search=document.getElementById('search'),query=search.value;if(!query){matches=[];matchedCells.clear();matchIndex=0;searchedQuery='';searchRequest++;document.getElementById('searchCount').textContent='';document.getElementById('clearSearch').hidden=true;refreshSearchHighlights();return}if(query===searchedQuery&&matches.length){matchIndex=(matchIndex+1)%matches.length;scrollToCell(matches[matchIndex],false,false);refreshSearchHighlights();return}searchedQuery=query;matches=[];matchedCells.clear();matchIndex=0;refreshSearchHighlights();document.getElementById('searchCount').textContent='';const requestId=++searchRequest;vscode.postMessage({type:'search',query,requestId,sheetIndex:model?.sheetIndex||0})}function dragEdgeSpeed(position,start,end){const margin=24;if(position<start+margin)return -Math.min(24,Math.max(1,(start+margin-position)*.35));if(position>end-margin)return Math.min(24,Math.max(1,(position-(end-margin))*.35));return 0}function updateDragSelection(){if(!dragging||!dragPointer||!dragMoved||!model)return;const grid=document.getElementById('grid'),rect=grid.getBoundingClientRect();if(!rect.width||!rect.height)return;const deltaX=dragEdgeSpeed(dragPointer.x,rect.left,rect.right),deltaY=dragEdgeSpeed(dragPointer.y,rect.top,rect.bottom);if(deltaX)grid.scrollLeft=Math.max(0,Math.min(grid.scrollWidth-grid.clientWidth,grid.scrollLeft+deltaX));if(deltaY)grid.scrollTop=Math.max(0,Math.min(grid.scrollHeight-grid.clientHeight,grid.scrollTop+deltaY));const freezeRows=model?.freezePane?.rows||0,freezeColumns=model?.freezePane?.columns||0,x=Math.max(rect.left+ROW_HEADER_WIDTH+1,Math.min(rect.right-1,dragPointer.x)),y=Math.max(rect.top+30,Math.min(rect.bottom-1,dragPointer.y));const rowPixel=y<rect.top+30+freezeRows*26?y-rect.top-30:y-rect.top-30+grid.scrollTop,row=Math.max(0,Math.min(model.gridRowCount-1,Math.floor(rowPixel/26)));let frozenWidth=0;for(let column=0;column<freezeColumns;column++)frozenWidth+=clampColumnWidth(columnWidths.get(column)??DEFAULT_COLUMN_WIDTH);const columnPixel=x<rect.left+ROW_HEADER_WIDTH+frozenWidth?x-rect.left-ROW_HEADER_WIDTH:x-rect.left-ROW_HEADER_WIDTH+grid.scrollLeft,column=Math.max(0,Math.min(model.gridColumnCount-1,columnAtPixel(columnOffsets,columnPixel))),target=document.querySelector('[data-row="'+row+'"][data-column="'+column+'"]');if(row!==selected.row||column!==selected.column){dragged=true;select(row,column,target?.__jinSheetData,true)}}function startCellDrag(){clearInterval(dragScrollTimer);dragScrollTimer=setInterval(updateDragSelection,40)}function stopCellDrag(){dragging=false;dragPointer=null;dragMoved=false;clearInterval(dragScrollTimer);dragScrollTimer=null}function moveAfterCommit(row,column,backwards=false){const rowCount=model.gridRowCount||model.rowCount;const columnCount=model.gridColumnCount||model.columnCount;let nextRow=row,nextColumn=column+(backwards?-1:1);if(nextColumn>=columnCount){nextRow=Math.min(rowCount-1,nextRow+1);nextColumn=0}else if(nextColumn<0){nextRow=Math.max(0,nextRow-1);nextColumn=columnCount-1}scrollToCell({row:nextRow,column:nextColumn})}
  document.getElementById('grid').addEventListener('scroll',()=>{syncHeadings();clearTimeout(scrollTimer);scrollTimer=setTimeout(requestRange,60)});
  document.addEventListener('mousemove',e=>{if(dragging){dragPointer={x:e.clientX,y:e.clientY};dragMoved=true;updateDragSelection()}if(columnResize)resizeColumn(columnResize.column,columnResize.startWidth+e.clientX-columnResize.startX)});window.addEventListener('blur',stopCellDrag);document.addEventListener('mouseup',()=>{stopCellDrag();if(columnResize){const {column,startWidth}=columnResize;const changed=Math.abs((columnWidths.get(column)??startWidth)-startWidth)>=1;columnResize=null;requestRange();if(changed)commitColumnWidth(column)}});
  document.getElementById('formula').addEventListener('keydown',e=>{if(e.key==='Escape'){e.preventDefault();const cell=document.querySelector('[data-row="'+selected.row+'"][data-column="'+selected.column+'"]');select(selected.row,selected.column,{value:cell?.textContent||''});return}if((e.key==='Enter'||e.key==='Tab')&&!e.target.disabled){if(e.isComposing||e.keyCode===229)return;e.preventDefault();const row=selected.row,column=selected.column;vscode.postMessage({type:'edit',row,column,value:e.target.value,sheetIndex:model?.sheetIndex||0});if(e.key==='Tab')moveAfterCommit(row,column,e.shiftKey);else scrollToCell({row:Math.min((model.gridRowCount||model.rowCount)-1,row+1),column})}});
  const searchInput=document.getElementById('search'),clearSearchButton=document.getElementById('clearSearch');function resetSearchResults(){matches=[];matchedCells.clear();matchIndex=0;searchedQuery='';searchRequest++;document.getElementById('searchCount').textContent='';refreshSearchHighlights()}function clearSearch(){searchInput.value='';resetSearchResults();clearSearchButton.hidden=true;searchInput.focus({preventScroll:true})}searchInput.addEventListener('input',()=>{resetSearchResults();clearSearchButton.hidden=!searchInput.value});document.getElementById('findButton').addEventListener('click',()=>{searchInput.focus({preventScroll:true});runSearch()});clearSearchButton.addEventListener('click',clearSearch);
  searchInput.addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();runSearch()}if(e.key==='Escape'){e.preventDefault();clearSearch()}});
  document.addEventListener('keydown', e => {
    if (e.target.id === 'search') return;
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'f') {
      e.preventDefault(); const search = document.getElementById('search'); search.focus(); search.select(); return;
    }
    if (e.target.tagName === 'INPUT' || !model || !e.target.closest?.('#grid')) return;
    if (e.key === 'Delete' || e.key === 'Backspace') {
      if (model.readOnly) return;
      e.preventDefault(); vscode.postMessage({ type: 'edit', row: selected.row, column: selected.column, value: '', sheetIndex: model.sheetIndex }); return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c') {
      e.preventDefault(); vscode.postMessage({ type: 'copy', ...selected, range: selectedRange, sheetIndex: model.sheetIndex }); return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'v') {
      e.preventDefault(); vscode.postMessage({ type: 'paste', ...selected, sheetIndex: model.sheetIndex }); return;
    }
    if ((e.ctrlKey || e.metaKey) && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) {
      e.preventDefault();
      const direction = { ArrowUp: 'up', ArrowDown: 'down', ArrowLeft: 'left', ArrowRight: 'right' }[e.key];
      const requestId = ++navigationRequest;
      vscode.postMessage({ type: 'navigate', ...selected, direction, extend: e.shiftKey, requestId, sheetIndex: model.sheetIndex }); return;
    }
    let row = selected.row, column = selected.column;
    const usedRows = Math.max(1, model.rowCount || 0), usedColumns = Math.max(1, model.columnCount || 0);
    if (e.key === 'Home') { row = e.ctrlKey || e.metaKey ? 0 : row; column = 0; }
    else if (e.key === 'End') {
      if (e.ctrlKey || e.metaKey) row = usedRows - 1;
      column = usedColumns - 1;
    } else if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight', 'Tab', 'Enter'].includes(e.key)) {
      if (e.key === 'ArrowUp') row = Math.max(0, row - 1);
      if (e.key === 'ArrowDown' || e.key === 'Enter') row = Math.min((model.gridRowCount || model.rowCount) - 1, row + 1);
      if (e.key === 'ArrowLeft' || (e.key === 'Tab' && e.shiftKey)) column = Math.max(0, column - 1);
      if (e.key === 'ArrowRight' || (e.key === 'Tab' && !e.shiftKey)) column = Math.min((model.gridColumnCount || model.columnCount) - 1, column + 1);
    } else return;
    e.preventDefault(); scrollToCell({ row, column }, e.shiftKey && e.key !== 'Tab');
  });
  function acceptModel(next){if(model&&model.sheetIndex===next.sheetIndex){if(JSON.stringify(model.columnWidths||{})!==JSON.stringify(next.columnWidths||{}))columnWidthsBySheet.delete(next.sheetIndex);selectionAnchor={row:selected.row,column:selected.column};selectedRange={startRow:selected.row,endRow:selected.row,startColumn:selected.column,endColumn:selected.column}}model=next;draw()}
  window.addEventListener('message',e=>{if(e.data?.type==='error'){document.getElementById('notice').textContent=e.data.message;return}if(e.data?.type==='selection-summary'){if(e.data.requestId!==summaryRequest)return;const target=document.getElementById('selectionSummary');if(e.data.tooLarge){target.textContent=t('selection-too-large');return}const format=value=>new Intl.NumberFormat(undefined,{maximumFractionDigits:10}).format(value);target.textContent=t('selection-count')+': '+e.data.count+' · '+t('selection-numbers')+': '+e.data.numericCount+(e.data.sum===null?'':' · '+t('selection-sum')+': '+format(e.data.sum)+' · '+t('selection-average')+': '+format(e.data.average));return}if(e.data?.type==='model'){acceptModel(e.data);return}if(e.data?.type==='range'){renderRange(e.data);return}if(e.data?.type==='auto-fit-width'){if(e.data.requestId!==autoFitRequest||e.data.sheetIndex!==model?.sheetIndex||e.data.column!==selected.column)return;resizeColumn(e.data.column,e.data.width);commitColumnWidth(e.data.column);requestRange();return}if(e.data?.type==='navigation'){if(e.data.requestId!==navigationRequest)return;scrollToCell({row:e.data.row,column:e.data.column},e.data.extend);return}if(e.data?.type==='search'){if(e.data.requestId!==searchRequest||e.data.query!==document.getElementById('search').value)return;matches=e.data.matches;matchedCells=new Set(matches.map(match=>match.row+':'+match.column));matchIndex=0;document.getElementById('searchCount').textContent=e.data.total?e.data.total+ui['match-separator']+t('matches'):t('no-matches');refreshSearchHighlights();if(matches.length)scrollToCell(matches[0],false,false)}});vscode.postMessage({type:'ready'});
  </script></body></html>`;
}

function deactivate() {}
module.exports = { activate, deactivate, estimateDocumentColumnWidth, findMatches, getRange, isValidSheetRequest, matchesSelectedSheet, navigateCell, selectionSummary, parseClipboard, serializeClipboard, JinSheetDocument, JinSheetProvider, renderHtml, viewModel };
