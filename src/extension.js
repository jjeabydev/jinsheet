'use strict';

const vscode = require('vscode');
const fs = require('node:fs/promises');
const crypto = require('node:crypto');
const path = require('node:path');
const { applyEdits, encode, load } = require('./csv');
const { analyze, patchCells } = require('./xlsx');
const { analyzeOffice } = require('./office');

const CSV_LIMIT_BYTES = 10 * 1024 * 1024;
const CSV_LIMIT_ROWS = 10000;
const CSV_LIMIT_COLUMNS = 1000;
const MAX_CELL_CHARS = 100000;

function hash(bytes) { return crypto.createHash('sha256').update(bytes).digest('hex'); }

class JinSheetDocument {
  constructor(uri, kind, state, bytes, writable, reason) {
    this.uri = uri;
    this.kind = kind;
    this.state = state;
    this.originalBytes = Buffer.from(bytes);
    this.originalFingerprint = hash(bytes);
    this.writable = writable && uri.scheme === 'file';
    this.reason = this.writable ? reason : (reason || '원격·가상 파일 시스템에서는 안전한 저장을 보장할 수 없어 읽기 전용으로 엽니다.');
    this.drafts = new Map();
    this.emitter = new vscode.EventEmitter();
    this.onDidChange = this.emitter.event;
    this.renderEmitter = new vscode.EventEmitter();
    this.onDidRender = this.renderEmitter.event;
    this.saving = null;
  }

  get dirty() { return this.drafts.size > 0; }

  getCell(row, column, sheetIndex = 0) {
    if (this.kind === 'csv') return this.state.rows[row]?.[column] ? { ...this.state.rows[row][column], row, column, value: this.state.rows[row][column].value } : null;
    const sheet = this.state.sheets[sheetIndex];
    const address = addressFor(row + 1, column + 1);
    const cell = sheet?.cells.get(address);
    if (!cell) return { row, column, address, value: '', kind: 'blank', formula: null };
    return { ...cell, value: this.drafts.get(`${sheetIndex}:${address}`) ?? cell.value };
  }

  applyCellEdit(row, column, value, sheetIndex = 0) {
    if (!this.writable) throw new Error(this.reason || '읽기 전용 파일입니다.');
    if (!Number.isInteger(row) || !Number.isInteger(column) || row < 0 || column < 0 || typeof value !== 'string' || value.length > MAX_CELL_CHARS) throw new RangeError('셀 수정 요청이 유효하지 않습니다.');
    const cell = this.getCell(row, column, sheetIndex);
    if (!cell) throw new RangeError('셀 범위를 벗어났습니다.');
    if (this.kind === 'xlsx' && cell.kind === 'blank') throw new Error('현재는 통합문서에 이미 존재하는 셀만 수정할 수 있습니다.');
    if (this.kind === 'xlsx' && cell.formula) throw new Error('수식 셀은 수정할 수 없습니다.');
    const previous = cell.value;
    if (previous === value) return;
    this.setDraft(row, column, value, sheetIndex);
    this.emitter.fire({ label: '셀 값 편집', undo: () => { this.setDraft(row, column, previous, sheetIndex); this.renderEmitter.fire(); }, redo: () => { this.setDraft(row, column, value, sheetIndex); this.renderEmitter.fire(); } });
    this.renderEmitter.fire();
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
      const baseline = this.state.sheets[sheetIndex].cells.get(address).value;
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
    } else {
      const edits = [...this.drafts].map(([key, value]) => {
        const [sheet, address] = key.split(':');
        return { sheetIndex: Number(sheet), address, value };
      });
      output = patchCells(this.state, edits).bytes;
    }
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
    else this.state = analyze(output, path.extname(sourcePath));
    this.drafts.clear();
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
    if (this.kind === 'csv') this.originalCsvState = this.state; else this.state = analyze(output, path.extname(destination.fsPath));
    this.drafts.clear();
    this.renderEmitter.fire();
  }

  createOutput() {
    if (this.kind === 'csv') return encode(this.state.text, this.state.encoding, this.state.bom);
    return patchCells(this.state, [...this.drafts].map(([key, value]) => { const [sheet, address] = key.split(':'); return { sheetIndex: Number(sheet), address, value }; })).bytes;
  }

  async revert() {
    if (this.kind === 'csv') this.setInitialCsvState(load(this.originalBytes, path.extname(this.uri.fsPath)));
    else this.state = analyze(this.originalBytes, path.extname(this.uri.fsPath));
    this.drafts.clear();
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
    this.emitter = new vscode.EventEmitter();
    this.onDidChangeCustomDocument = this.emitter.event;
  }

  async openCustomDocument(uri) {
    const bytes = Buffer.from(await vscode.workspace.fs.readFile(uri));
    const extension = path.extname(uri.fsPath).toLowerCase();
    let document;
    if (extension === '.xlsx' || extension === '.xlsm') {
      const state = analyze(bytes, extension);
      document = new JinSheetDocument(uri, 'xlsx', state, bytes, state.writable, state.readOnlyReason);
    } else if (['.docx', '.docm', '.pptx', '.pptm'].includes(extension)) {
      let state, reason = 'Office 문서는 텍스트 미리보기 전용입니다. 편집·저장은 지원하지 않습니다.';
      try { state = analyzeOffice(bytes, extension); }
      catch (error) { reason = `문서 미리보기를 열 수 없습니다: ${error.message}`; state = { unsupported: true, format: 'office', title: 'Office 문서', sections: [] }; }
      document = new JinSheetDocument(uri, 'office', state, bytes, false, reason);
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
    const sub = document.onDidChange(change => this.emitter.fire({ document, label: change.label, undo: change.undo, redo: change.redo }));
    document.customDocumentSubscription = sub;
    return document;
  }

  async resolveCustomEditor(document, panel) {
    let selectedSheet = 0;
    panel.webview.options = { enableScripts: true, localResourceRoots: [] };
    panel.webview.html = renderHtml(document);
    panel.webview.onDidReceiveMessage(async message => {
      if (!message || typeof message !== 'object' || Array.isArray(message)) return;
      try {
        if (message.type === 'edit' && Number.isInteger(message.row) && Number.isInteger(message.column) && Number.isInteger(message.sheetIndex || 0)) {
          document.applyCellEdit(message.row, message.column, message.value, message.sheetIndex || 0);
        } else if (message.type === 'range') {
          selectedSheet = message.sheetIndex || 0;
          panel.webview.postMessage(getRange(document, message));
        } else if (message.type === 'copy' && Number.isInteger(message.row) && Number.isInteger(message.column)) {
          const cell = document.getCell(message.row, message.column, message.sheetIndex || 0);
          await vscode.env.clipboard.writeText(cell?.formula ? `=${cell.formula}` : (cell?.value ?? ''));
        } else if (message.type === 'paste' && Number.isInteger(message.row) && Number.isInteger(message.column)) {
          const value = await vscode.env.clipboard.readText();
          document.applyCellEdit(message.row, message.column, value, message.sheetIndex || 0);
        } else if (message.type === 'search' && typeof message.query === 'string' && message.query.length <= 1000) {
          panel.webview.postMessage(findMatches(document, message.query, message.sheetIndex || 0));
        } else if (message.type === 'ready' || message.type === 'refresh') {
          selectedSheet = message.sheetIndex || 0;
          await panel.webview.postMessage(viewModel(document, selectedSheet));
        }
      } catch (error) { panel.webview.postMessage({ type: 'error', message: String(error.message || error) }); }
    }, undefined, this.context.subscriptions);
    const sub = document.onDidRender(() => panel.webview.postMessage(viewModel(document, selectedSheet)));
    panel.onDidDispose(() => sub.dispose());
  }

  saveCustomDocument(document, token) { return document.save(token); }
  saveCustomDocumentAs(document, destination, token) { return document.saveAs(destination, token); }
  revertCustomDocument(document) { return document.revert(); }
  backupCustomDocument(document, context) { return document.backup(context); }
  dispose() { this.emitter.dispose(); }
}

function viewModel(document, sheetIndex = 0) {
  if (document.kind === 'office') return { type: 'model', kind: 'office', readOnly: true, reason: document.reason, dirty: false, title: document.state.title, format: document.state.format, sections: document.state.sections || [] };
  if (document.kind === 'csv') {
    if (!document.state.rows) return { type: 'model', kind: 'csv', readOnly: true, reason: document.reason, rows: [] };
    let columnCount = 0;
    for (const row of document.state.rows) columnCount = Math.max(columnCount, row.length);
    return { type: 'model', kind: 'csv', readOnly: !document.writable, reason: document.reason, dirty: document.dirty, rowCount: document.state.rows.length, columnCount, sheetIndex: 0 };
  }
  const sheet = document.state.sheets[sheetIndex];
  if (!sheet) return { type: 'model', kind: 'xlsx', readOnly: true, reason: document.reason || '표시할 시트가 없습니다.', sheets: [] };
  let usedRows = 0, usedColumns = 0;
  for (const cell of sheet.cells.values()) { usedRows = Math.max(usedRows, cell.row + 1); usedColumns = Math.max(usedColumns, cell.column + 1); }
  return {
    type: 'model', kind: 'xlsx', readOnly: !document.writable, reason: document.reason, dirty: document.dirty,
    rowCount: usedRows, columnCount: usedColumns,
    sheetIndex, sheets: document.state.sheets.map((item, index) => ({ name: item.name, index, visible: item.visible }))
  };
}

function getRange(document, message) {
  const { rowStart, columnStart, rowCount, columnCount, requestId } = message;
  if (![rowStart, columnStart, rowCount, columnCount, requestId].every(Number.isInteger) || rowStart < 0 || columnStart < 0 || rowCount < 1 || rowCount > 60 || columnCount < 1 || columnCount > 30) throw new RangeError('조회 범위가 유효하지 않습니다.');
  const sheetIndex = message.sheetIndex || 0;
  const rows = Array.from({ length: rowCount }, (_, r) => Array.from({ length: columnCount }, (_, c) => {
    const cell = document.getCell(rowStart + r, columnStart + c, sheetIndex);
    return cell ? { value: cell.value, formula: cell.formula || null, cachedValue: cell.cachedValue ?? null, address: cell.address || null } : null;
  }));
  return { type: 'range', requestId, rowStart, columnStart, rowCount, columnCount, sheetIndex, rows };
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
  else for (const cell of document.state.sheets[sheetIndex]?.cells.values() || []) visit(cell.row, cell.column, cell);
  return { type: 'search', query, total, matches, sheetIndex };
}

function renderHtml() {
  const nonce = crypto.randomBytes(18).toString('base64');
  const csp = `default-src 'none'; style-src 'nonce-${nonce}'; script-src 'nonce-${nonce}';`;
  return `<!doctype html><html lang="ko"><head><meta charset="UTF-8"><meta http-equiv="Content-Security-Policy" content="${csp}"><meta name="viewport" content="width=device-width,initial-scale=1"><style nonce="${nonce}">
  *{box-sizing:border-box}body{margin:0;color:var(--vscode-editor-foreground);background:var(--vscode-editor-background);font:13px var(--vscode-font-family)}#notice{padding:7px 10px;color:var(--vscode-editorWarning-foreground);min-height:30px}#toolbar{display:flex;gap:6px;align-items:center;padding:4px 8px;border-bottom:1px solid var(--vscode-panel-border)}button{color:var(--vscode-button-foreground);background:var(--vscode-button-background);border:0;padding:4px 8px}button[aria-selected=true]{background:var(--vscode-button-hoverBackground)}#formula{flex:1;padding:4px 8px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;border:1px solid var(--vscode-input-border)}#search{width:190px;height:26px;color:var(--vscode-input-foreground);background:var(--vscode-input-background);border:1px solid var(--vscode-input-border)}#grid{position:relative;overflow:auto;height:calc(100vh - 84px);outline:none}#canvas{position:relative}.cell,.rowhead,.colhead{position:absolute;height:26px;padding:4px 6px;overflow:hidden;white-space:nowrap;text-overflow:ellipsis;border:1px solid var(--vscode-panel-border)}.cell{width:144px;background:var(--vscode-editor-background)}.rowhead,.colhead{background:var(--vscode-editorWidget-background);text-align:center}.rowhead{width:52px}.colhead{width:144px;top:0}.corner{position:absolute;width:52px;height:30px;top:0;left:0;background:var(--vscode-editorWidget-background);z-index:2}.cell.selected{outline:1px solid var(--vscode-focusBorder);outline-offset:-1px}.cell:focus{outline:1px solid var(--vscode-focusBorder);outline-offset:-1px}#sheets{display:flex;gap:4px;padding:4px 8px;border-bottom:1px solid var(--vscode-panel-border)}
  </style></head><body><div id="notice" role="status" aria-live="polite"></div><div id="sheets" role="tablist"></div><div id="toolbar"><span id="address"></span><div id="formula" aria-label="셀 값 또는 수식"></div><input id="search" type="search" aria-label="시트에서 찾기" placeholder="찾기… (Ctrl/Cmd+F)"><span id="searchCount" aria-live="polite"></span></div><div id="grid" role="grid" aria-label="스프레드시트" tabindex="0"><div id="canvas"></div></div><script nonce="${nonce}">
  const vscode=acquireVsCodeApi();let model;let selected={row:0,column:0};let rangeRequest=0;let matches=[];let matchIndex=0;let scrollTimer;const label=n=>{let out='';while(n>0){n--;out=String.fromCharCode(65+n%26)+out;n=Math.floor(n/26)}return out};
  function draw(){if(!model)return;const notice=document.getElementById('notice');notice.textContent=(model.reason||'')+(model.dirty?' · 수정됨':'')+(model.rowCount>1000?' · 행은 스크롤해 불러옵니다.':'');const tabs=document.getElementById('sheets');tabs.replaceChildren();const grid=document.getElementById('grid'),toolbar=document.getElementById('toolbar'),canvas=document.getElementById('canvas');if(model.kind==='office'){grid.style.display='none';toolbar.style.display='none';tabs.style.display='none';canvas.replaceChildren();canvas.style.position='static';canvas.style.padding='16px';canvas.style.maxWidth='900px';canvas.style.margin='0 auto';const title=document.createElement('h1');title.textContent=model.title||'Office 문서';canvas.append(title);for(const section of model.sections||[]){const article=document.createElement('section');article.style.margin='0 0 24px';if(model.format==='pptx'){const heading=document.createElement('h2');heading.textContent=section.title;article.append(heading)}for(const paragraph of section.paragraphs||[]){const p=document.createElement('p');p.textContent=paragraph;p.style.whiteSpace='pre-wrap';article.append(p)}canvas.append(article)}if(!model.sections?.length){const empty=document.createElement('p');empty.textContent=model.reason||'미리 볼 텍스트가 없습니다.';canvas.append(empty)}return}grid.style.display='';toolbar.style.display='';tabs.style.display='flex';canvas.style.position='relative';canvas.style.padding='0';canvas.style.maxWidth='';canvas.style.margin='';for(const sheet of model.sheets||[]){const button=document.createElement('button');button.textContent=sheet.name+(sheet.visible?'':' (숨김)');button.setAttribute('role','tab');button.setAttribute('aria-selected',String(sheet.index===model.sheetIndex));button.onclick=()=>vscode.postMessage({type:'refresh',sheetIndex:sheet.index});tabs.append(button)}canvas.style.width=(52+model.columnCount*144)+'px';canvas.style.height=(30+model.rowCount*26)+'px';canvas.replaceChildren();if(!model.rowCount||!model.columnCount){canvas.textContent=model.reason||'표시할 데이터가 없습니다.';return}const corner=document.createElement('div');corner.className='corner';canvas.append(corner);requestRange()}
  function requestRange(){if(!model||!model.rowCount||!model.columnCount)return;const grid=document.getElementById('grid');const rowStart=Math.max(0,Math.floor(grid.scrollTop/26)-2);const columnStart=Math.max(0,Math.floor(Math.max(0,grid.scrollLeft-52)/144)-2);const rowCount=Math.min(60,Math.max(1,Math.ceil(grid.clientHeight/26)+5),model.rowCount-rowStart);const columnCount=Math.min(30,Math.max(1,Math.ceil(grid.clientWidth/144)+5),model.columnCount-columnStart);if(rowCount<1||columnCount<1)return;const requestId=++rangeRequest;const canvas=document.getElementById('canvas');for(const item of [...canvas.children])if(!item.classList.contains('corner'))item.remove();vscode.postMessage({type:'range',requestId,rowStart,columnStart,rowCount,columnCount,sheetIndex:model.sheetIndex})}
  function renderRange(range){if(range.requestId!==rangeRequest||range.sheetIndex!==model.sheetIndex)return;const canvas=document.getElementById('canvas');for(const item of [...canvas.children])if(!item.classList.contains('corner'))item.remove();for(let c=0;c<range.columnCount;c++){const column=range.columnStart+c;const head=document.createElement('div');head.className='colhead';head.textContent=label(column+1);head.style.left=(52+column*144)+'px';head.setAttribute('aria-label','열 '+label(column+1));canvas.append(head)}for(let r=0;r<range.rowCount;r++){const row=range.rowStart+r;const rowHead=document.createElement('div');rowHead.className='rowhead';rowHead.textContent=String(row+1);rowHead.style.top=(30+row*26)+'px';rowHead.setAttribute('aria-label','행 '+(row+1));canvas.append(rowHead);for(let c=0;c<range.columnCount;c++){const column=range.columnStart+c;const data=range.rows[r][c];const cell=document.createElement('div');cell.className='cell'+(selected.row===row&&selected.column===column?' selected':'');cell.tabIndex=0;cell.setAttribute('role','gridcell');cell.setAttribute('aria-rowindex',String(row+1));cell.setAttribute('aria-colindex',String(column+1));cell.dataset.row=String(row);cell.dataset.column=String(column);cell.style.top=(30+row*26)+'px';cell.style.left=(52+column*144)+'px';cell.textContent=data?.value??'';cell.title=cell.textContent;cell.onclick=()=>select(row,column,data);cell.ondblclick=()=>edit(cell,row,column,data);cell.onkeydown=e=>{if(e.key==='F2'||e.key==='Enter'){e.preventDefault();edit(cell,row,column,data)}};canvas.append(cell)}}select(selected.row,selected.column,range.rows[selected.row-range.rowStart]?.[selected.column-range.columnStart])}
  function select(r,c,data){selected={row:r,column:c};document.querySelector('.cell.selected')?.classList.remove('selected');document.querySelector('[data-row="'+r+'"][data-column="'+c+'"]')?.classList.add('selected');document.getElementById('address').textContent=label(c+1)+(r+1);const formula=document.getElementById('formula');formula.textContent=data?.formula?('='+data.formula+'  ·  캐시값: '+(data.cachedValue??'')):(data?.value??'');formula.title=formula.textContent}
  function edit(cell,r,c,data){if(!cell||model.readOnly)return;const input=document.createElement('input');input.value=data?.value??'';input.style.width='132px';cell.replaceChildren(input);input.focus();input.select();let done=false;const commit=()=>{if(done)return;done=true;vscode.postMessage({type:'edit',row:r,column:c,value:input.value,sheetIndex:model.sheetIndex})};input.onkeydown=e=>{if(e.key==='Enter'){commit();input.blur()}else if(e.key==='Escape'){done=true;requestRange()}};input.onblur=commit}
  function scrollToCell(cell){selected={row:cell.row,column:cell.column};const grid=document.getElementById('grid');grid.scrollTop=Math.max(0,cell.row*26);grid.scrollLeft=Math.max(0,cell.column*144);requestRange()}
  document.getElementById('grid').addEventListener('scroll',()=>{clearTimeout(scrollTimer);scrollTimer=setTimeout(requestRange,60)});
  document.getElementById('search').addEventListener('input',e=>vscode.postMessage({type:'search',query:e.target.value,sheetIndex:model?.sheetIndex||0}));
  document.getElementById('search').addEventListener('keydown',e=>{if(e.key==='Enter'&&matches.length){e.preventDefault();matchIndex=(matchIndex+1)%matches.length;scrollToCell(matches[matchIndex])}if(e.key==='Escape'){e.target.value='';matches=[];document.getElementById('searchCount').textContent=''}});
  document.addEventListener('keydown',e=>{if(e.target.id==='search')return;if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='f'){e.preventDefault();document.getElementById('search').focus();document.getElementById('search').select();return}if(e.target.tagName==='INPUT'||!model)return;if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='c'){e.preventDefault();vscode.postMessage({type:'copy',...selected,sheetIndex:model.sheetIndex});return}if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='v'){e.preventDefault();vscode.postMessage({type:'paste',...selected,sheetIndex:model.sheetIndex});return}if(['ArrowUp','ArrowDown','ArrowLeft','ArrowRight','Tab','Enter'].includes(e.key)){e.preventDefault();if(e.key==='ArrowUp')selected.row=Math.max(0,selected.row-1);if(e.key==='ArrowDown'||e.key==='Enter')selected.row=Math.min(model.rowCount-1,selected.row+1);if(e.key==='ArrowLeft')selected.column=Math.max(0,selected.column-1);if(e.key==='ArrowRight'||e.key==='Tab')selected.column=Math.min(model.columnCount-1,selected.column+1);scrollToCell(selected)}});
  window.addEventListener('message',e=>{if(e.data?.type==='error'){document.getElementById('notice').textContent=e.data.message;return}if(e.data?.type==='model'){model=e.data;draw();return}if(e.data?.type==='range'){renderRange(e.data);return}if(e.data?.type==='search'){if(e.data.query!==document.getElementById('search').value)return;matches=e.data.matches;matchIndex=0;document.getElementById('searchCount').textContent=e.data.total?e.data.total+'개 일치':'일치 항목 없음';if(matches.length)scrollToCell(matches[0])}});vscode.postMessage({type:'ready'});
  </script></body></html>`;
}

function deactivate() {}
module.exports = { activate, deactivate, findMatches, getRange, JinSheetDocument, JinSheetProvider, renderHtml, viewModel };
