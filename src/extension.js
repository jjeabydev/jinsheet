'use strict';

const vscode = require('vscode');
const fs = require('node:fs/promises');
const crypto = require('node:crypto');
const path = require('node:path');
const { unzipSync } = require('../vendor/node_modules/fflate');
const { applyEdits, encode, extendCells: extendCsvCells, load, parse: parseCsv, quote: quoteCsv, transformStructure: transformCsvStructure } = require('./csv');
const { LIMITS: XLSX_LIMITS, TABLE_STYLE_NAMES, analyze, calculateChartPreviews, conditionalStyleForCell, formatCellValue, parseRangeReference, patchCells, patchColumnWidths, patchRowHeights, patchRowHeightValues, patchPrintSettings, patchPrintArea, patchDefinedName, patchPrintTitles, validatePrintTitles, patchHeaderFooter, patchPageBreaks, tableFilterRange, patchAutoFilters, patchFilterVisibility, patchHiddenRows, patchHiddenColumns, patchOutlineGroups, patchFreezePanes, patchMergedCells, planTableHeaderEdits, patchTableColumnNames, planTableExpansion, prepareTableExpansion, patchCalculatedColumns, planCalculatedColumn, patchAddTable, patchTableStyles, patchTableRanges, patchTableTotals, totalsFunctionForCell, patchAddChart, patchChartMetadata, patchChartPlacement, patchChartSource, patchDeleteChart, patchMoveConditionalFormatRule, patchSheetVisibility, patchSheetZoom, patchSheetName, patchAddSheet, patchDeleteSheet, patchMoveSheet, translateFormulaReferences, transformStructure: transformXlsxStructure } = require('./xlsx');
const { createFormulaCalculator, evaluateFormula } = require('./formula');
const { analyzeOffice } = require('./office');
const { pdfCanvasOutputScale, validatePdfInput } = require('./pdf');
const { DEFAULT_COLUMN_WIDTH, ROW_HEADER_WIDTH, createColumnOffsets, visibleColumnRange, clampColumnWidth, excelColumnWidthPixels, estimateBestFitWidth, browserGridSource } = require('./grid');
const { languageOf, message, messagesFor, translateDiagnostic } = require('./i18n');
const { DEFAULT_PRINT_SETTINGS, renderPrintDocument, combinePrintDocuments } = require('./printing');
const { chartPrintGeometry } = require('./chart-print');
const { defaultHeaderFooter } = require('./header-footer');
const { createConditionalValueReader } = require('./conditional-style');
const { validateCustomFilter, validateTopFilter, validateDateGroups, parseDateGroups, DYNAMIC_FILTER_TYPES, ICON_SET_LENGTHS, validateDynamicFilter, dynamicFilterUsesClock, validateColorFilter, validateIconFilter, filterColumnCells, compileAutoFilterColumn } = require('./auto-filter');
const { splitDelimitedText } = require('./text-to-columns');
const { patchDataValidations, patchDataBarConditionalFormat, patchCellValueConditionalFormat, patchColorScaleConditionalFormat, patchDuplicateValuesConditionalFormat, patchIconSetConditionalFormat, patchClearConditionalFormats, patchDeleteConditionalFormatRule, patchEditConditionalFormatRule, patchHyperlinks } = require('./xlsx');

const CSV_LIMIT_BYTES = 10 * 1024 * 1024;
const CSV_LIMIT_ROWS = 10000;
const CSV_LIMIT_COLUMNS = 1000;
const MAX_CELL_CHARS = 100000;
const MAX_CLIPBOARD_CHARS = 1000000;

function xlsxDefaultColumnWidth(sheet, maximumDigitWidth = 7) {
  return Number.isFinite(sheet?.defaultColumnWidthCharacters)
    ? excelColumnWidthPixels(sheet.defaultColumnWidthCharacters, sheet.defaultColumnWidth ?? 64, maximumDigitWidth)
    : sheet?.defaultColumnWidth ?? 64;
}

function worksheetColumnLabel(column) {
  let value = column + 1, label = '';
  while (value) { const remainder = (value - 1) % 26; label = String.fromCharCode(65 + remainder) + label; value = Math.floor((value - 1) / 26); }
  return label;
}

function conditionalFormatRangeLabel(range) {
  const first = worksheetColumnLabel(range.left) + (range.top + 1), last = worksheetColumnLabel(range.right) + (range.bottom + 1);
  return first === last ? first : first + ':' + last;
}

function definedNameReferenceLabel(name) {
  const source = name.source, sheet = "'" + source.sheetName.replaceAll("'", "''") + "'!";
  return sheet + '$' + worksheetColumnLabel(source.left) + '$' + (source.top + 1) + ':$' + worksheetColumnLabel(source.right) + '$' + (source.bottom + 1);
}

function conditionalFormatRuleOption(rule, localized) {
  const type = localized['conditional-format-type-' + rule.type] || rule.type;
  const ranges = (rule.ranges || []).slice(0, 4).map(conditionalFormatRangeLabel);
  if ((rule.ranges || []).length > ranges.length) ranges.push('+' + ((rule.ranges || []).length - ranges.length));
  let condition = '';
  if (rule.type === 'cellIs') {
    const operator = localized['conditional-format-operator-' + rule.operator] || rule.operator;
    condition = operator + ' ' + (rule.values || []).join(' – ');
  } else if (rule.type === 'top10') {
    condition = localized[rule.top10?.bottom ? 'conditional-format-bottom' : 'conditional-format-top'] + ' ' + rule.top10?.rank + (rule.top10?.percent ? '%' : '');
  } else if (rule.type === 'expression') condition = '=' + rule.formula;
  else if (rule.type === 'dataBar') condition = rule.dataBar?.color || '';
  else if (rule.type === 'colorScale') condition = (rule.colorScale?.colors || []).join(' → ');
  else if (rule.type === 'iconSet') condition = rule.iconSet?.name || '';
  return {
    label: localized['conditional-format-priority'] + ' ' + rule.priority + ': ' + type,
    description: ranges.join(', ') + (condition ? ' · ' + condition : ''),
    detail: rule.stopIfTrue ? localized['conditional-format-stop-if-true'] : undefined,
    priority: rule.priority
  };
}

function dataValidationValues(document, sheetIndex, validation, calculator, row = validation?.ranges?.[0]?.top, column = validation?.ranges?.[0]?.left) {
  if (Array.isArray(validation?.values)) return validation.values;
  if (!calculator) return [];
  let source = validation?.source;
  if (validation?.sourceFormula) {
    const anchor = validation.ranges.find(range => row >= range.top && row <= range.bottom && column >= range.left && column <= range.right) || validation.ranges[0];
    const formula = translateFormulaReferences(validation.sourceFormula, row - anchor.top, column - anchor.left);
    const indirect = /^INDIRECT\((.+)\)$/i.exec(formula.trim());
    if (!indirect) return [];
    const argument = indirect[1].trim(), literal = /^"((?:[^"]|"")*)"$/.exec(argument);
    let reference = literal ? literal[1].replaceAll('""', '"') : evaluateFormula(argument, { locale: document.locale, date1904: document.state.date1904, getCell: (targetRow, targetColumn, sheetName) => {
      const targetSheet = sheetName ? document.state.sheets.findIndex(sheet => sheet.name.toLocaleLowerCase() === sheetName.toLocaleLowerCase()) : sheetIndex;
      return targetSheet < 0 ? '#REF!' : calculator.value(targetSheet, targetRow, targetColumn);
    } });
    if (typeof reference !== 'string' || reference.startsWith('#')) return [];
    source = parseRangeReference(reference);
    if (!source) {
      const name = /^([A-Z_\\][A-Z0-9_.]*)$/i.exec(reference)?.[1];
      const definition = name && (document.state.definedNames?.find(item => item.localSheetId === sheetIndex && item.name.toLocaleLowerCase() === name.toLocaleLowerCase()) || document.state.definedNames?.find(item => item.localSheetId === null && item.name.toLocaleLowerCase() === name.toLocaleLowerCase()));
      source = definition?.source || null;
    }
  }
  if (!source) return [];
  const { top, bottom, left, right } = source, values = [], seen = new Set();
  const sourceSheetIndex = source.sheetName ? document.state.sheets.findIndex(sheet => sheet.name.toLocaleLowerCase() === source.sheetName.toLocaleLowerCase()) : sheetIndex;
  if (sourceSheetIndex < 0) return [];
  for (let row = top; row <= bottom; row++) for (let column = left; column <= right; column++) {
    const value = calculator.value(sourceSheetIndex, row, column);
    if (value === null || value === undefined || value === '' || typeof value === 'string' && value.startsWith('#')) continue;
    const text = typeof value === 'boolean' ? (value ? 'TRUE' : 'FALSE') : String(value);
    if (!seen.has(text)) { seen.add(text); values.push(text); }
  }
  return values;
}

function dataValidationFormulaLimits(document, sheetIndex, validation, row, column, calculator) {
  if (!validation.formulas) return validation.values;
  const anchor = validation.ranges.find(range => row >= range.top && row <= range.bottom && column >= range.left && column <= range.right) || validation.ranges[0];
  return validation.formulas.map((formula, index) => {
    if (formula === null) return validation.values[index];
    const translated = translateFormulaReferences(formula, row - anchor.top, column - anchor.left);
    return evaluateFormula(translated, { locale: document.locale, date1904: document.state.date1904, getCell: (targetRow, targetColumn, sheetName) => {
      const targetSheet = sheetName ? document.state.sheets.findIndex(sheet => sheet.name.toLocaleLowerCase() === sheetName.toLocaleLowerCase()) : sheetIndex;
      return targetSheet < 0 ? '#REF!' : calculator.value(targetSheet, targetRow, targetColumn);
    } });
  });
}

function validationNumber(type, value, date1904 = false, locale = 'en-US') {
  if (type === 'textLength') return Number(value);
  const numeric = typeof value === 'number' || /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(String(value ?? ''));
  if (type === 'date') return numeric ? Number(value) : parseClipboardDate(String(value ?? ''), locale, date1904)?.serial;
  if (type === 'time') return numeric ? Number(value) : parseClipboardTime(String(value ?? ''))?.serial;
  return numeric ? Number(value) : NaN;
}

function dataValidationAllows(validation, value, allowedValues, limits = validation.values, date1904 = false, locale = 'en-US') {
  if (value === '') return validation.allowBlank;
  if (validation.type === 'list') return allowedValues.includes(value);
  const number = validation.type === 'textLength' ? value.length : validationNumber(validation.type, value, date1904, locale);
  if (!Number.isFinite(number)) return false;
  if (validation.type === 'whole' && !Number.isInteger(number) || validation.type === 'date' && (number < 0 || number > 2958465 - (date1904 ? 1462 : 0)) || validation.type === 'time' && (number < 0 || number > 1)) return false;
  const [rawFirst, rawSecond] = limits, first = validationNumber(validation.type, rawFirst, date1904, locale), second = validationNumber(validation.type, rawSecond, date1904, locale);
  if (!Number.isFinite(first) || ['between', 'notBetween'].includes(validation.operator) && !Number.isFinite(second)) return false;
  return validation.operator === 'equal' ? number === first
    : validation.operator === 'notEqual' ? number !== first
    : validation.operator === 'greaterThan' ? number > first
    : validation.operator === 'lessThan' ? number < first
    : validation.operator === 'greaterThanOrEqual' ? number >= first
    : validation.operator === 'lessThanOrEqual' ? number <= first
    : validation.operator === 'between' ? number >= first && number <= second
    : number < first || number > second;
}

function subtractValidationRange(source, cut) {
  const top = Math.max(source.top, cut.top), bottom = Math.min(source.bottom, cut.bottom), left = Math.max(source.left, cut.left), right = Math.min(source.right, cut.right);
  if (top > bottom || left > right) return [source];
  const remaining = [];
  if (source.top < top) remaining.push({ top: source.top, bottom: top - 1, left: source.left, right: source.right });
  if (bottom < source.bottom) remaining.push({ top: bottom + 1, bottom: source.bottom, left: source.left, right: source.right });
  if (source.left < left) remaining.push({ top, bottom, left: source.left, right: left - 1 });
  if (right < source.right) remaining.push({ top, bottom, left: right + 1, right: source.right });
  return remaining;
}

function parseValidationListInput(input, document, sheetIndex) {
  const raw = String(input).trim();
  if (!raw) throw new Error('목록 값이나 셀 범위를 입력하세요.');
  const reference = raw.startsWith('=') ? raw.slice(1).trim() : raw;
  const source = parseRangeReference(reference);
  if (source) return { source };
  const name = /^[A-Z_\\][A-Z0-9_.]*$/i.test(reference) ? reference : null;
  const definition = name && (document.state.definedNames?.find(item => item.localSheetId === sheetIndex && item.name.toLocaleLowerCase() === name.toLocaleLowerCase()) || document.state.definedNames?.find(item => item.localSheetId === null && item.name.toLocaleLowerCase() === name.toLocaleLowerCase()));
  if (definition) {
    if (!definition.source) throw new Error('셀 범위를 가리키는 정의된 이름만 목록 원본으로 사용할 수 있습니다.');
    return { source: definition.source };
  }
  if (raw.startsWith('=')) throw new Error('목록 원본에는 셀 범위, 정의된 이름 또는 쉼표로 구분한 값을 입력하세요.');
  const values = [], value = []; let quoted = false;
  for (let index = 0; index < raw.length; index++) {
    const character = raw[index];
    if (character === '"') {
      if (quoted && raw[index + 1] === '"') { value.push('"'); index++; }
      else quoted = !quoted;
    } else if (character === ',' && !quoted) { values.push(value.join('').trim()); value.length = 0; }
    else value.push(character);
  }
  if (quoted) throw new Error('목록 값의 따옴표 짝이 맞지 않습니다.');
  values.push(value.join('').trim());
  const formulaLength = values.join(',').replaceAll('"', '""').length + 2;
  if (values.length > 50 || values.some(item => item.length > 255) || formulaLength > 255) throw new Error('목록은 50개 이하이며 전체 길이는 255자 이하여야 합니다.');
  return { values };
}

function validationDateSerial(value, date1904) {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (!match) throw new Error('날짜는 YYYY-MM-DD 형식으로 입력하세요.');
  const year = Number(match[1]), month = Number(match[2]), day = Number(match[3]);
  if (year === 1900 && month === 2 && day === 29) {
    if (date1904) throw new Error('1904 날짜 체계에는 1900-02-29가 없습니다.');
    return 60;
  }
  const date = new Date(0); date.setUTCFullYear(year, month - 1, day); date.setUTCHours(0, 0, 0, 0);
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) throw new Error('날짜가 유효하지 않습니다.');
  const serial = Math.floor((date.getTime() - Date.UTC(1899, 11, 30)) / 86400000) - (date1904 ? 1462 : 0);
  if (serial < 0 || serial > 2958465 - (date1904 ? 1462 : 0)) throw new Error('날짜가 Excel에서 지원하는 범위를 벗어났습니다.');
  return serial;
}

function validationOperand(input, type, date1904, locale = 'en-US') {
  const raw = String(input).trim();
  if (!raw) throw new Error('유효성 기준값을 입력하세요.');
  if (raw.startsWith('=')) {
    const formula = raw.slice(1);
    if (!formula || formula.length > 255) throw new Error('유효성 기준 수식은 255자 이하여야 합니다.');
    const sample = evaluateFormula(formula, { locale, date1904, getCell: () => 1, getSparseRange: () => [] });
    if (typeof sample === 'string' && ['#NAME?', '#VALUE!', '#REF!'].includes(sample) && !(sample === '#REF!' && formula.includes('#REF!'))) throw new Error('유효성 기준 수식을 계산할 수 없습니다.');
    return { value: null, formula };
  }
  let number;
  if (type === 'date') number = validationDateSerial(raw, date1904);
  else if (type === 'time') {
    const match = /^(\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?$/.exec(raw);
    if (!match) throw new Error('시간은 HH:MM 또는 HH:MM:SS 형식으로 입력하세요.');
    const hour = Number(match[1]), minute = Number(match[2]), second = Number(match[3] || 0), fraction = Number('0.' + (match[4] || '0'));
    if (hour > 24 || minute > 59 || second > 59 || hour === 24 && (minute || second || fraction)) throw new Error('시간이 유효하지 않습니다.');
    number = (hour * 3600 + minute * 60 + second + fraction) / 86400;
  } else {
    if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(raw)) throw new Error('유효성 기준값은 숫자 또는 =로 시작하는 수식이어야 합니다.');
    number = Number(raw);
  }
  if (!Number.isFinite(number) || type === 'whole' && !Number.isInteger(number) || type === 'textLength' && (!Number.isInteger(number) || number < 0)) throw new Error('유효성 기준값의 숫자 형식이 올바르지 않습니다.');
  return { value: number, formula: null };
}

function buildDataValidationRule(document, sheetIndex, input) {
  const type = input?.type;
  if (!['list', 'whole', 'decimal', 'date', 'time', 'textLength'].includes(type)) throw new RangeError('데이터 유효성 종류가 유효하지 않습니다.');
  const rule = { type, operator: null, ranges: [], values: null, formulas: null, errorStyle: 'stop', allowBlank: true, showDropDown: false, showInputMessage: false, showErrorMessage: true, errorTitle: '', error: '', promptTitle: '', prompt: '' };
  if (type === 'list') return { ...rule, ...parseValidationListInput(input.source, document, sheetIndex) };
  const operators = ['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'];
  if (!operators.includes(input.operator)) throw new RangeError('데이터 유효성 비교 연산자가 유효하지 않습니다.');
  const first = validationOperand(input.first, type, document.state.date1904, document.locale), needsSecond = ['between', 'notBetween'].includes(input.operator);
  const second = needsSecond ? validationOperand(input.second, type, document.state.date1904, document.locale) : null;
  return { ...rule, operator: input.operator, values: [first.value, ...(second ? [second.value] : [])], formulas: first.formula || second?.formula ? [first.formula, ...(second ? [second.formula] : [])] : null };
}

function hash(bytes) { return crypto.createHash('sha256').update(bytes).digest('hex'); }

class JinSheetDocument {
  constructor(uri, kind, state, bytes, writable, reason) {
    this.uri = uri;
    this.locale = Intl.DateTimeFormat().resolvedOptions().locale || 'en-US';
    this.kind = kind;
    this.state = state;
    this.maximumDigitWidth = 7;
    this.originalBytes = Buffer.from(bytes);
    this.originalFingerprint = hash(bytes);
    this.savedRevision = 0;
    this.baselineXlsxFiles = kind === 'xlsx' ? state.files : null;
    this.writable = writable && uri.scheme === 'file';
    this.reason = this.writable ? reason : (reason || '원격·가상 파일 시스템에서는 안전한 저장을 보장할 수 없어 읽기 전용으로 엽니다.');
    this.drafts = new Map();
    this.styleDrafts = new Map();
    this.formulaDrafts = new Map();
    this.columnWidthDrafts = new Map();
    this.csvAutoFilter = null;
    this.csvOutline = { rowOutlineLevels: {}, columnOutlineLevels: {}, collapsedRows: new Set(), collapsedColumns: new Set() };
    this.csvHiddenRows = new Set();
    this.csvHiddenColumns = new Set();
    this.structureChanged = false;
    this.emitter = new vscode.EventEmitter();
    this.onDidChange = this.emitter.event;
    this.renderEmitter = new vscode.EventEmitter();
    this.onDidRender = this.renderEmitter.event;
    this.autoRowHeights = new Map();
    this.pendingAutoRowFits = new Set();
    this.renderRevision = 0; this.onDidRender(derived => {
      if (derived === 'row-layout') return;
      this.renderRevision++;
      this.autoRowHeights.clear();
      for (const pending of this.pendingAutoRowFits) pending.cancel();
    });
    this.saving = null;
  }

  get dirty() { return this.structureChanged || this.drafts.size > 0 || this.styleDrafts.size > 0 || this.formulaDrafts.size > 0 || this.columnWidthDrafts.size > 0; }

  beginAutoRowFit(sheetIndex, rows) {
    const pending = { sheetIndex, rows, revision: this.renderRevision };
    let resolve;
    pending.promise = new Promise(done => { resolve = done; });
    pending.cancel = () => {
      clearTimeout(pending.timer);
      this.pendingAutoRowFits.delete(pending);
      resolve();
    };
    // Hidden/disposed webviews must never block save indefinitely.
    pending.timer = setTimeout(pending.cancel, 2000);
    this.pendingAutoRowFits.add(pending);
    return pending;
  }

  completeAutoRowFit(pending, message) {
    if (!this.pendingAutoRowFits.has(pending)) return false;
    try {
      validateRowFitResult(pending, message, this.renderRevision, pending.sheetIndex);
      this.autoRowHeights.set(pending.sheetIndex, message.heights.map(item => ({ ...item })));
      this.renderEmitter.fire('row-layout');
      return true;
    } finally { pending.cancel(); }
  }

  async waitForAutoRowFits() {
    while (this.pendingAutoRowFits.size) await Promise.all([...this.pendingAutoRowFits].map(pending => pending.promise));
  }

  getCell(row, column, sheetIndex = 0) {
    if (this.kind === 'csv') return this.state.rows[row]?.[column] ? { ...this.state.rows[row][column], row, column, value: this.state.rows[row][column].value } : null;
    const sheet = this.state.sheets[sheetIndex];
    const address = addressFor(row + 1, column + 1);
    const cell = sheet?.cells.get(address);
    const draft = this.drafts.get(sheetIndex + ':' + address);
    const styleDraft = this.styleDrafts.get(sheetIndex + ':' + address);
    const formulaDraft = this.formulaDrafts.get(sheetIndex + ':' + address);
    const styleFormat = styleDraft && styleDraft.index !== null && /^(?:0|[1-9]\d*)$/.test(styleDraft.index)
      ? this.state.styleFormats?.[Number(styleDraft.index)] : null;
    const numberFormatId = styleDraft?.numberFormatId;
    const numberFormat = numberFormatId !== undefined && (styleDraft?.forceNumberFormat || (styleFormat?.numFmtId ?? 0) === 0);
    const styleOverlay = styleDraft ? {
      style: styleDraft.index,
      styleProperties: styleDraft.properties ?? styleFormat?.properties ?? null,
      numFmtId: numberFormat ? numberFormatId : styleFormat?.numFmtId ?? 0,
      formatCode: numberFormat ? (styleDraft.numberFormatCode || ({ 9: '0%', 10: '0.00%', 14: 'm/d/yy', 20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yy h:mm' })[numberFormatId]) : styleFormat?.formatCode ?? null,
      ...(styleDraft.kind !== undefined ? { kind: styleDraft.kind } : {})
    } : {};
    if (!cell) {
      const value = draft ?? '';
      const numericDraft = typeof value === 'string' && /^\s*[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?\s*$/.test(value) && Number.isFinite(Number(value.trim()));
      const kind = styleDraft?.kind ?? (value === '' ? 'blank' : numericDraft ? 'number' : /^(?:TRUE|FALSE)$/i.test(value) ? 'boolean' : 'string');
      return { row, column, address, value, kind, formula: formulaDraft ?? null, hasFormula: formulaDraft !== undefined && formulaDraft !== null, ...(this.state.date1904 ? { date1904: true } : {}), ...styleOverlay };
    }
    return { ...cell, value: draft ?? cell.value, ...(formulaDraft === undefined ? {} : { formula: formulaDraft, hasFormula: formulaDraft !== null }), ...styleOverlay, ...(draft === undefined ? {} : { kind: typedBlankKind(draft, styleOverlay.kind ?? cell.kind) }) };
  }

  applyCellEdit(row, column, value, sheetIndex = 0, acceptedWarnings = new Set(), options = {}) {
    if (!this.writable) throw new Error(this.reason || '읽기 전용 파일입니다.');
    const cell = this.getCell(row, column, sheetIndex);
    if (this.kind === 'xlsx' && typeof value === 'string') {
      if (value.startsWith('=')) {
        const formula = value.slice(1);
        if (!formula.length || formula.length > MAX_CELL_CHARS) throw new RangeError('셀 수식이 유효하지 않습니다.');
        this.applyCellEdits([{ row, column, value: cell?.cachedValue ?? cell?.value ?? '', formula }], sheetIndex, '셀 수식 편집', acceptedWarnings, true, options);
        return;
      }
      if (isFormulaCell(cell)) {
        this.applyCellEdits([{ row, column, value, formula: null }], sheetIndex, '셀 값 편집', acceptedWarnings, true, options);
        return;
      }
    }
    this.applyCellEdits([{ row, column, value }], sheetIndex, '셀 값 편집', acceptedWarnings, true, options);
  }

  applyRenamedTableEdits(plan, sheetIndex, label, acceptedWarnings, emit, options) {
    const before = this.captureSnapshot(), revision = this.renderRevision;
    const shadow = Object.create(this);
    shadow.drafts = new Map(this.drafts); shadow.styleDrafts = new Map(this.styleDrafts); shadow.formulaDrafts = new Map(this.formulaDrafts); shadow.columnWidthDrafts = new Map(this.columnWidthDrafts);
    const ensureCurrent = () => { if (this.renderRevision !== revision || this.savedRevision !== before.savedRevision || this.state !== before.state) throw new Error('문서가 변경되어 표 머리글 수정을 다시 실행해야 합니다.'); };
    const commit = () => {
      ensureCurrent();
      const drafts = shadow.xlsxDraftEdits(), current = drafts.length ? patchCells(shadow.state, drafts).model : shadow.state;
      shadow.state = patchTableColumnNames(current, plan.requests).model;
      this.state = shadow.state; this.drafts = new Map(); this.styleDrafts = new Map(); this.formulaDrafts = new Map(); this.structureChanged = !this.matchesSavedState();
      if (!emit) return true;
      const after = this.captureSnapshot();
      this.emitter.fire({ label: '표 머리글 변경', undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } }); this.renderEmitter.fire(); return true;
    };
    const wrap = error => {
      if (error.validationPrompt && typeof error.retry === 'function') { const retry = error.retry; error.retry = () => { ensureCurrent(); try { retry(); } catch (next) { throw wrap(next); } return commit(); }; }
      return error;
    };
    try { shadow.applyCellEdits(plan.edits, sheetIndex, label, acceptedWarnings, false, { ...options, skipHeaders: true }); } catch (error) { throw wrap(error); }
    return commit();
  }

  applyExpandedTableEdits(plan, sheetIndex, label, acceptedWarnings, emit, options) {
    const before = this.captureSnapshot(), revision = this.renderRevision, drafts = this.xlsxDraftEdits();
    let current = drafts.length ? patchCells(this.state, drafts).model : this.state;
    current = prepareTableExpansion(current, plan.requests);
    const skipCells = new Set(plan.edits.filter(edit => Number.isInteger(edit?.row) && Number.isInteger(edit?.column)).map(edit => sheetIndex + ':' + addressFor(edit.row + 1, edit.column + 1)));
    current = patchTableRanges(current, plan.requests, { skipCells }).model;
    const shadow = Object.create(this);
    shadow.state = current; shadow.drafts = new Map(); shadow.styleDrafts = new Map(); shadow.formulaDrafts = new Map(); shadow.columnWidthDrafts = new Map(this.columnWidthDrafts); shadow.structureChanged = true;
    const ensureCurrent = () => { if (this.renderRevision !== revision || this.savedRevision !== before.savedRevision || this.state !== before.state) throw new Error('문서가 변경되어 표 자동 확장을 다시 실행해야 합니다.'); };
    const commit = () => {
      ensureCurrent(); this.state = shadow.state; this.drafts = shadow.drafts; this.styleDrafts = shadow.styleDrafts; this.formulaDrafts = shadow.formulaDrafts; this.structureChanged = !this.matchesSavedState();
      if (!emit) return true;
      const after = this.captureSnapshot();
      this.emitter.fire({ label: '표 자동 확장', undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } }); this.renderEmitter.fire(); return true;
    };
    const wrap = error => {
      if (error.validationPrompt && typeof error.retry === 'function') {
        const retry = error.retry;
        error.retry = () => { ensureCurrent(); try { retry(); } catch (next) { throw wrap(next); } return commit(); };
      }
      return error;
    };
    try { shadow.applyCellEdits(plan.edits, sheetIndex, label, acceptedWarnings, false, { ...options, skipExpansion: true }); }
    catch (error) { throw wrap(error); }
    return commit();
  }

  applyCellEdits(edits, sheetIndex = 0, label = '셀 범위 붙여넣기', acceptedWarnings = new Set(), emit = true, options = {}) {
    if (!this.writable) throw new Error(this.reason || '읽기 전용 파일입니다.');
    const maximumEditCount = this.kind === 'xlsx' ? XLSX_LIMITS.cells * 2 : 20000;
    if (!Array.isArray(edits) || edits.length < 1 || edits.length > maximumEditCount || !Number.isInteger(sheetIndex) || sheetIndex < 0) throw new RangeError('셀 범위 수정 요청이 유효하지 않습니다.');
    if (this.kind === 'csv' && sheetIndex !== 0) throw new RangeError('CSV에는 첫 번째 시트만 있습니다.');
    if (this.kind === 'xlsx' && !this.state.sheets[sheetIndex]) throw new RangeError('수정할 시트를 찾을 수 없습니다.');
    if (this.kind === 'xlsx' && !options.skipExpansion) {
      const expansion = planTableExpansion(this.state, sheetIndex, edits, (row, column, index) => this.getCell(row, column, index));
      if (expansion.requests.length) return this.applyExpandedTableEdits(expansion, sheetIndex, label, acceptedWarnings, emit, options);
    }
    if (this.kind === 'xlsx' && !options.skipHeaders) {
      const headers = planTableHeaderEdits(this.state, sheetIndex, edits);
      if (headers.requests.length) return this.applyRenamedTableEdits(headers, sheetIndex, label, acceptedWarnings, emit, options);
      edits = headers.edits;
    }
    const requestedEdits = edits, calculatedMetadata = [...(options.calculatedMetadata || [])];
    if (this.kind === 'xlsx' && !options.skipCalculated && edits.length === 1 && typeof edits[0]?.formula === 'string' && this.getCell(edits[0].row, edits[0].column, sheetIndex)?.formula !== edits[0].formula) {
      const origin = edits[0], plan = planCalculatedColumn(this.state, { sheetIndex, row: origin.row, column: origin.column, formula: origin.formula }, (row, column, index) => this.getCell(row, column, index));
      if (plan) { edits = plan.edits.map(edit => edit.row === origin.row ? { ...edit, ...origin } : edit); calculatedMetadata.push(plan.metadata); label = '계산 열 수식 변경'; }
    }
    if (this.kind === 'xlsx' && (!options.skipCalculated || edits.length > 1)) for (const table of this.state.sheets[sheetIndex].tables || []) {
      const first = table.range.top + table.headerRowCount, last = table.range.bottom - table.totalsRowCount;
      for (const [index, formula] of (table.calculatedColumns || []).entries()) {
        if (!formula) continue;
        const column = table.range.left + index, cleared = new Set(edits.filter(edit => edit && edit.column === column && edit.row >= first && edit.row <= last && (Object.hasOwn(edit, 'formula') ? edit.formula === null : !['styleIndex','styleProperties','font','fillColor','alignment','border','numberFormatId'].some(key => Object.hasOwn(edit,key)))).map(edit => edit.row));
        if (cleared.size === last - first + 1) calculatedMetadata.push({ sheetIndex, tableName: table.name, column, formula: null });
      }
    }
    const seen = new Set(), changes = [];
    let validationCalculator = null, candidateCalculator = null, spillCalculator = null;
    for (const edit of edits) {
      const { row, column, value, styleIndex, styleProperties, kind } = edit || {};
      const hasFormula = this.kind === 'xlsx' && Object.hasOwn(edit || {}, 'formula'), formula = edit?.formula;
      const hasNumberFormat = this.kind === 'xlsx' && Object.hasOwn(edit || {}, 'numberFormatId');
      const hasFont = this.kind === 'xlsx' && Object.hasOwn(edit || {}, 'font');
      const hasFill = this.kind === 'xlsx' && Object.hasOwn(edit || {}, 'fillColor');
      const hasAlignment = this.kind === 'xlsx' && Object.hasOwn(edit || {}, 'alignment');
      const hasBorder = this.kind === 'xlsx' && Object.hasOwn(edit || {}, 'border');
      const forceNumberFormat = this.kind === 'xlsx' && edit?.forceNumberFormat === true;
      const numberFormatCode = this.kind === 'xlsx' && typeof edit?.numberFormatCode === 'string' ? edit.numberFormatCode : undefined;
      if (!Number.isInteger(row) || !Number.isInteger(column) || row < 0 || column < 0 || typeof value !== 'string' || value.length > MAX_CELL_CHARS) throw new RangeError('셀 수정 요청이 유효하지 않습니다.');
      if (hasFormula && formula !== null && (typeof formula !== 'string' || !formula.length || formula.length > MAX_CELL_CHARS || formula.startsWith('='))) throw new RangeError('셀 수식이 유효하지 않습니다.');
      const hasStyle = this.kind === 'xlsx' && Object.hasOwn(edit || {}, 'styleIndex');
      if (hasStyle && styleIndex !== null && (typeof styleIndex !== 'string' || !/^(?:0|[1-9]\d*)$/.test(styleIndex))) throw new RangeError('셀 서식 인덱스가 유효하지 않습니다.');
      const hasKind = this.kind === 'xlsx' && Object.hasOwn(edit || {}, 'kind');
      if (hasKind && !['blank', 'string', 'number', 'boolean', 'error'].includes(kind)) throw new RangeError('셀 값 종류가 유효하지 않습니다.');
      if (hasNumberFormat && (!([4, 9, 10, 14, 20, 21, 22].includes(edit.numberFormatId) || edit.numberFormatId >= 164 && edit.numberFormatId <= 65535 && numberFormatCode && numberFormatCode.length <= 255) || kind !== 'number' || (Object.hasOwn(edit, 'forceNumberFormat') && typeof edit.forceNumberFormat !== 'boolean'))) throw new RangeError('숫자 표시 서식 요청이 유효하지 않습니다.');
      if (hasFont && (!edit.font || typeof edit.font !== 'object' || Object.keys(edit.font).some(key => !['bold', 'italic', 'color', 'underline', 'fontSize', 'fontFamily', 'strike'].includes(key)) || typeof edit.font.bold !== 'boolean' || typeof edit.font.italic !== 'boolean' || (Object.hasOwn(edit.font, 'color') && edit.font.color !== null && (typeof edit.font.color !== 'string' || !/^#[0-9A-F]{6}$/i.test(edit.font.color))) || (Object.hasOwn(edit.font, 'underline') && typeof edit.font.underline !== 'boolean') || (Object.hasOwn(edit.font, 'strike') && typeof edit.font.strike !== 'boolean') || (Object.hasOwn(edit.font, 'fontSize') && edit.font.fontSize !== null && (typeof edit.font.fontSize !== 'number' || !Number.isFinite(edit.font.fontSize) || edit.font.fontSize < 1 || edit.font.fontSize > 409 || edit.font.fontSize * 2 % 1 !== 0)) || (Object.hasOwn(edit.font, 'fontFamily') && edit.font.fontFamily !== null && (typeof edit.font.fontFamily !== 'string' || !edit.font.fontFamily.trim() || edit.font.fontFamily.length > 64)))) throw new RangeError('글꼴 서식 요청이 유효하지 않습니다.');
      if (hasFill && edit.fillColor !== null && (typeof edit.fillColor !== 'string' || !/^#[0-9A-F]{6}$/i.test(edit.fillColor))) throw new RangeError('채우기 서식 요청이 유효하지 않습니다.');
      if (hasAlignment && (!edit.alignment || typeof edit.alignment !== 'object' || Object.keys(edit.alignment).length < 1 || Object.keys(edit.alignment).length > 3 || Object.keys(edit.alignment).some(axis => !['horizontal', 'vertical', 'wrapText'].includes(axis)) || (Object.hasOwn(edit.alignment, 'horizontal') && !['left', 'center', 'right'].includes(edit.alignment.horizontal)) || (Object.hasOwn(edit.alignment, 'vertical') && !['top', 'center', 'bottom'].includes(edit.alignment.vertical)) || (Object.hasOwn(edit.alignment, 'wrapText') && typeof edit.alignment.wrapText !== 'boolean'))) throw new RangeError('셀 맞춤 요청이 유효하지 않습니다.');
      if (hasBorder && (!edit.border || typeof edit.border !== 'object' || !Object.keys(edit.border).length || Object.keys(edit.border).some(side => !['left', 'right', 'top', 'bottom'].includes(side)) || Object.values(edit.border).some(edge => edge !== null && (!edge || typeof edge !== 'object' || !['hair', 'thin', 'medium', 'thick', 'dashed', 'dotted', 'double'].includes(edge.line) || (edge.color !== undefined && (typeof edge.color !== 'string' || !/^#[0-9A-F]{6}$/i.test(edge.color))))))) throw new RangeError('셀 테두리 요청이 유효하지 않습니다.');
      if (this.kind === 'csv' && (row >= CSV_LIMIT_ROWS || column >= CSV_LIMIT_COLUMNS)) throw new RangeError('CSV 행·열 제한을 벗어난 셀입니다.');
      if (this.kind === 'xlsx' && (row >= XLSX_LIMITS.rows || column >= XLSX_LIMITS.columns)) throw new RangeError('XLSX 행·열 제한을 벗어난 셀입니다.');
      const key = row + ':' + column;
      if (seen.has(key)) throw new RangeError('같은 셀을 중복 수정할 수 없습니다.');
      seen.add(key);
      const cell = this.getCell(row, column, sheetIndex);
      if (this.kind === 'xlsx') {
        const calculatedCell = (spillCalculator ||= createFormulaCalculator(this)).cell(sheetIndex, row, column);
        if (calculatedCell?.isSpill) throw new Error('동적 배열 결과의 일부 셀은 편집할 수 없습니다. 배열 수식이 있는 셀을 편집하세요.');
        const merged = this.state.sheets[sheetIndex].mergedCells?.find(range => row >= range.top && row <= range.bottom && column >= range.left && column <= range.right);
        if (merged && (row !== merged.top || column !== merged.left)) throw new Error('병합된 셀은 왼쪽 위 셀에서만 편집할 수 있습니다.');
        const validation = this.state.sheets[sheetIndex].dataValidations?.find(rule => rule.ranges.some(range => row >= range.top && row <= range.bottom && column >= range.left && column <= range.right));
        if (validation) {
          const allowedValues = dataValidationValues(this, sheetIndex, validation, validationCalculator ||= createFormulaCalculator(this), row, column);
          const calculated = hasFormula && typeof formula === 'string' ? (candidateCalculator ||= projectedFormulaCalculator(this, sheetIndex, edits)).value(sheetIndex, row, column) : value;
          const candidate = typeof calculated === 'boolean' ? (calculated ? 'TRUE' : 'FALSE') : calculated === null || calculated === undefined ? '' : String(calculated);
          const limits = dataValidationFormulaLimits(this, sheetIndex, validation, row, column, validationCalculator);
          if (validation.showErrorMessage !== false && !dataValidationAllows(validation, candidate, allowedValues, limits, this.state.date1904, this.locale)) {
            const customMessage = [validation.errorTitle, validation.error].filter(Boolean).join(': ');
            const error = new Error(customMessage || (validation.type === 'list' ? '이 셀은 허용 목록의 값만 입력할 수 있습니다.' : '이 셀의 데이터 유효성 조건을 만족하지 않습니다.'));
            error.userMessage = Boolean(customMessage);
            if (validation.errorStyle === 'stop') throw error;
            if (!acceptedWarnings.has(key)) {
              error.validationPrompt = { style: validation.errorStyle, message: customMessage || '이 값은 데이터 유효성 조건에 맞지 않습니다.', sheetIndex, row, column };
              error.retry = () => {
                const accepted = new Set(acceptedWarnings);
                accepted.add(key);
                return this.applyCellEdits(requestedEdits, sheetIndex, label, accepted, emit, options);
              };
              throw error;
            }
          }
        }
      }
      if (this.kind === 'xlsx' && ((isFormulaCell(cell) || cell?.kind === 'error') && !hasFormula)) throw new Error('수식·오류 셀이 포함된 범위는 수정할 수 없습니다.');
      if (this.kind === 'xlsx' && cell?.formulaType && cell.formulaType !== 'normal' && cell.formulaType !== 'shared' && hasFormula) throw new Error('배열 수식 셀은 수정할 수 없습니다.');
      const formulaChanged = hasFormula && cell?.formula !== formula;
      const kindChanged = hasKind && cell?.kind !== kind;
      const style = hasStyle || hasFont || hasFill || hasAlignment || hasBorder ? { index: hasStyle ? styleIndex : cell?.style ?? null, properties: styleProperties ?? cell?.styleProperties ?? null, ...(hasFont ? { font: edit.font } : {}), ...(hasFill ? { fillColor: edit.fillColor } : {}), ...(hasAlignment ? { alignment: edit.alignment } : {}), ...(hasBorder ? { border: edit.border } : {}) } : hasNumberFormat ? { index: cell?.style ?? null, properties: cell?.styleProperties ?? null, numberFormatId: edit.numberFormatId, ...(numberFormatCode ? { numberFormatCode } : {}), ...(forceNumberFormat ? { forceNumberFormat: true } : {}) } : undefined;
      const propertiesChanged = hasStyle && JSON.stringify(cell?.styleProperties ?? null) !== JSON.stringify(styleProperties ?? null);
      if ((cell?.value ?? '') !== value || (hasStyle && (cell?.style ?? null) !== styleIndex) || propertiesChanged || kindChanged || formulaChanged || hasNumberFormat || hasFont || hasFill || hasAlignment || hasBorder) changes.push({ row, column, value, kind: hasKind ? kind : undefined, style, formula: hasFormula ? formula : undefined, numberFormatId: hasNumberFormat ? edit.numberFormatId : undefined, numberFormatCode });
    }
    if (this.kind === 'xlsx' && calculatedMetadata.some(item => item.formula !== null)) {
      const pending = new Set();
      for (const edit of this.xlsxDraftEdits()) if (!this.state.sheets[edit.sheetIndex].cells.has(edit.address)) pending.add(edit.sheetIndex + ':' + edit.address);
      for (const edit of edits) { const address = addressFor(edit.row + 1, edit.column + 1); if (!this.state.sheets[sheetIndex].cells.has(address)) pending.add(sheetIndex + ':' + address); }
      if (this.state.sheets.reduce((count, sheet) => count + sheet.cells.size, 0) + pending.size > XLSX_LIMITS.cells) throw new RangeError('계산 열을 채우면 통합문서 100,000셀 제한을 초과합니다.');
    }
    const calculatedState = this.kind === 'xlsx' && calculatedMetadata.length ? patchCalculatedColumns(this.state, calculatedMetadata).model : this.state;
    if (!changes.length && calculatedState === this.state) return;
    const before = this.captureSnapshot();
    if (this.kind === 'csv') {
      const missingCells = changes.filter(({ row, column }) => !this.state.rows[row]?.[column]);
      if (missingCells.length) {
        this.state = extendCsvCells(this.state, path.extname(this.uri.fsPath), missingCells);
        this.originalCsvState = this.state;
        this.drafts.clear();
        this.structureChanged = true;
      }
    }
    for (const edit of changes) this.setDraft(edit.row, edit.column, edit.value, sheetIndex, edit.style, edit.formula, edit.kind, edit.numberFormatId, edit.numberFormatCode);
    if (this.kind === 'csv') this.state = this.makeCsvState();
    if (this.kind === 'xlsx' && calculatedState !== this.state) { this.state = calculatedState; this.structureChanged = !this.matchesSavedState(); }
    if (!emit) return true;
    const after = this.captureSnapshot();
    this.emitter.fire({
      label,
      undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); },
      redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); }
    });
    this.renderEmitter.fire();
  }

  applyCellEditsAcrossSheets(sheetEdits, label = '셀 범위 붙여넣기', acceptedWarnings = new Map()) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || '통합 문서 범위 수정은 XLSX에서만 지원합니다.');
    const seenSheets = new Set(); let editCount = 0;
    if (!Array.isArray(sheetEdits) || !sheetEdits.length || sheetEdits.some(item => {
      if (!item || !Number.isInteger(item.sheetIndex) || !this.state.sheets[item.sheetIndex] || seenSheets.has(item.sheetIndex) || !Array.isArray(item.edits) || !item.edits.length) return true;
      seenSheets.add(item.sheetIndex); editCount += item.edits.length;
      return editCount > 20000;
    })) throw new RangeError('통합 문서 범위 수정 요청이 유효하지 않습니다.');
    const before = this.captureSnapshot(), revision = this.renderRevision;
    let changed = false;
    try {
      for (const { sheetIndex, edits } of sheetEdits) {
        const accepted = acceptedWarnings.get(sheetIndex) || new Set();
        if (this.applyCellEdits(edits, sheetIndex, label, accepted, false) === true) changed = true;
      }
    } catch (error) {
      this.restoreSnapshot(before);
      if (error.validationPrompt && typeof error.retry === 'function') {
        const prompt = error.validationPrompt;
        error.retry = () => {
          if (this.renderRevision !== revision) throw new Error('문서가 변경되어 찾기 및 바꾸기를 다시 시도해야 합니다.');
          const nextAccepted = new Map([...acceptedWarnings].map(([sheetIndex, keys]) => [sheetIndex, new Set(keys)]));
          const keys = nextAccepted.get(prompt.sheetIndex) || new Set();
          keys.add(prompt.row + ':' + prompt.column);
          nextAccepted.set(prompt.sheetIndex, keys);
          return this.applyCellEditsAcrossSheets(sheetEdits, label, nextAccepted);
        };
      }
      throw error;
    }
    if (!changed) return false;
    const after = this.captureSnapshot();
    this.emitter.fire({
      label,
      undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); },
      redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); }
    });
    this.renderEmitter.fire();
    return true;
  }

  applyColumnWidth(column, pixelWidth, sheetIndex = 0) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || 'XLSX 열 너비를 저장할 수 없습니다.');
    if (!Number.isInteger(sheetIndex) || !this.state.sheets[sheetIndex] || !Number.isInteger(column) || column < 0 || column >= XLSX_LIMITS.columns || !Number.isFinite(pixelWidth) || pixelWidth < 1 || pixelWidth > 1785) throw new RangeError('열 너비 수정 요청이 유효하지 않습니다.');
    const key = sheetIndex + ':' + column;
    const sheet = this.state.sheets[sheetIndex], maximumDigitWidth = this.maximumDigitWidth || 7;
    const baseline = sheet.columnWidths[column];
    const pixels = clampColumnWidth(pixelWidth, true);
    const width = Math.round((pixels / maximumDigitWidth) * 100) / 100;
    const previous = this.columnWidthDrafts.get(key);
    const resetToDefault = baseline === undefined && Math.abs(pixels - xlsxDefaultColumnWidth(sheet, maximumDigitWidth)) < 1;
    const resetToBaseline = baseline !== undefined && Math.abs(excelColumnWidthPixels(baseline, sheet.defaultColumnWidth ?? 64, maximumDigitWidth) - pixels) < 1;
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

  applyAutoFilter(sheetIndex, filter) {
    if (!this.writable) throw new Error(this.reason || '읽기 전용 파일입니다.');
    if (this.kind !== 'xlsx' || !Number.isInteger(sheetIndex) || !this.state.sheets[sheetIndex]) throw new RangeError('자동 필터를 적용할 시트가 유효하지 않습니다.');
    if (JSON.stringify(this.state.sheets[sheetIndex].autoFilter ?? null) === JSON.stringify(filter ?? null)) return;
    const before = this.captureSnapshot(), patched = patchAutoFilters(this.state, [{ sheetIndex, filter }], { now: this.now });
    if (Buffer.compare(Buffer.from(patched.bytes), Buffer.from(this.state.bytes)) === 0) return;
    if (Object.keys(patched.model.files).length === Object.keys(this.state.files).length && Object.entries(this.state.files).every(([name, bytes]) => patched.model.files[name] && Buffer.compare(Buffer.from(bytes), Buffer.from(patched.model.files[name])) === 0)) return;
    this.state = patched.model;
    this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({
      label: '자동 필터 변경',
      undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); },
      redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); }
    });
    this.renderEmitter.fire();
  }

  applyDataValidation(sheetIndex, range, rule = null) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || '데이터 유효성 규칙을 수정할 수 없는 문서입니다.');
    const sheet = this.state.sheets[sheetIndex];
    if (!sheet || !range || ![range.top, range.bottom, range.left, range.right].every(Number.isInteger) || range.top < 0 || range.left < 0 || range.bottom < range.top || range.right < range.left || range.bottom >= XLSX_LIMITS.rows || range.right >= XLSX_LIMITS.columns) throw new RangeError('데이터 유효성 적용 범위가 유효하지 않습니다.');
    if (rule !== null && (typeof rule !== 'object' || Array.isArray(rule))) throw new TypeError('데이터 유효성 규칙이 유효하지 않습니다.');
    const before = this.captureSnapshot(), drafts = this.xlsxDraftEdits(), current = drafts.length ? patchCells(this.state, drafts).model : this.state;
    const validations = [];
    for (const existing of current.sheets[sheetIndex].dataValidations || []) {
      const ranges = existing.ranges.flatMap(source => subtractValidationRange(source, range));
      if (ranges.length > 64) throw new RangeError('데이터 유효성 적용 범위 개수가 제한을 벗어났습니다.');
      if (ranges.length) validations.push({ ...existing, ranges });
    }
    if (rule) validations.push({ ...rule, ranges: [range] });
    if (validations.length > 256) throw new RangeError('데이터 유효성 규칙 개수가 제한을 벗어났습니다.');
    const patched = patchDataValidations(current, sheetIndex, validations);
    if (patched.model === current) return false;
    this.state = patched.model;
    this.drafts.clear(); this.styleDrafts.clear(); this.formulaDrafts.clear();
    this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot(), label = rule ? '데이터 유효성 규칙 설정' : '데이터 유효성 규칙 지우기';
    this.emitter.fire({ label, undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } });
    this.renderEmitter.fire();
    return true;
  }

  applyDataBars(sheetIndex, range, color) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || '조건부 서식을 수정할 수 없는 문서입니다.');
    const before = this.captureSnapshot(), drafts = this.xlsxDraftEdits(), current = drafts.length ? patchCells(this.state, drafts).model : this.state;
    const patched = patchDataBarConditionalFormat(current, sheetIndex, range, color);
    if (patched.model === current) return false;
    this.state = patched.model;
    this.drafts.clear(); this.styleDrafts.clear(); this.formulaDrafts.clear();
    this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({ label: '데이터 막대 조건부 서식 적용', undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } });
    this.renderEmitter.fire();
    return true;
  }

  applyCellValueConditionalFormat(sheetIndex, range, operator, values) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || '조건부 서식을 수정할 수 없는 문서입니다.');
    const before = this.captureSnapshot(), drafts = this.xlsxDraftEdits(), current = drafts.length ? patchCells(this.state, drafts).model : this.state;
    const patched = patchCellValueConditionalFormat(current, sheetIndex, range, operator, values);
    if (patched.model === current) return false;
    this.state = patched.model;
    this.drafts.clear(); this.styleDrafts.clear(); this.formulaDrafts.clear();
    this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({ label: message('conditional-format-cell-rule-history', this.language), undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } });
    this.renderEmitter.fire();
    return true;
  }

  applyColorScale(sheetIndex, range, colors) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || '조건부 서식을 수정할 수 없는 문서입니다.');
    const before = this.captureSnapshot(), drafts = this.xlsxDraftEdits(), current = drafts.length ? patchCells(this.state, drafts).model : this.state;
    const patched = patchColorScaleConditionalFormat(current, sheetIndex, range, colors);
    if (patched.model === current) return false;
    this.state = patched.model;
    this.drafts.clear(); this.styleDrafts.clear(); this.formulaDrafts.clear();
    this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({ label: '색조 조건부 서식 적용', undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } });
    this.renderEmitter.fire();
    return true;
  }

  applyDuplicateValues(sheetIndex, range) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || '조건부 서식을 수정할 수 없는 문서입니다.');
    const before = this.captureSnapshot(), drafts = this.xlsxDraftEdits(), current = drafts.length ? patchCells(this.state, drafts).model : this.state;
    const patched = patchDuplicateValuesConditionalFormat(current, sheetIndex, range);
    if (patched.model === current) return false;
    this.state = patched.model;
    this.drafts.clear(); this.styleDrafts.clear(); this.formulaDrafts.clear();
    this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({ label: message('conditional-format-duplicates-history', this.language), undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } });
    this.renderEmitter.fire();
    return true;
  }

  applyIconSet(sheetIndex, range, name) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || '조건부 서식을 수정할 수 없는 문서입니다.');
    const before = this.captureSnapshot(), drafts = this.xlsxDraftEdits(), current = drafts.length ? patchCells(this.state, drafts).model : this.state;
    const patched = patchIconSetConditionalFormat(current, sheetIndex, range, name);
    if (patched.model === current) return false;
    this.state = patched.model;
    this.drafts.clear(); this.styleDrafts.clear(); this.formulaDrafts.clear();
    this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({ label: message('conditional-format-icon-set-history', this.language), undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } });
    this.renderEmitter.fire();
    return true;
  }

  clearConditionalFormats(sheetIndex, range) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || '조건부 서식을 수정할 수 없는 문서입니다.');
    const before = this.captureSnapshot(), drafts = this.xlsxDraftEdits(), current = drafts.length ? patchCells(this.state, drafts).model : this.state;
    const patched = patchClearConditionalFormats(current, sheetIndex, range);
    if (patched.model === current) return false;
    this.state = patched.model;
    this.drafts.clear(); this.styleDrafts.clear(); this.formulaDrafts.clear();
    this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({ label: '조건부 서식 규칙 지우기', undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } });
    this.renderEmitter.fire();
    return true;
  }

  deleteConditionalFormatRule(sheetIndex, priority) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || '조건부 서식을 수정할 수 없는 문서입니다.');
    const before = this.captureSnapshot(), drafts = this.xlsxDraftEdits(), current = drafts.length ? patchCells(this.state, drafts).model : this.state;
    const patched = patchDeleteConditionalFormatRule(current, sheetIndex, priority);
    if (patched.model === current) return false;
    this.state = patched.model;
    this.drafts.clear(); this.styleDrafts.clear(); this.formulaDrafts.clear();
    this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({ label: '조건부 서식 규칙 제거', undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } });
    this.renderEmitter.fire();
    return true;
  }

  moveConditionalFormatRule(sheetIndex, priority, direction) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || '조건부 서식을 수정할 수 없는 문서입니다.');
    const before = this.captureSnapshot(), drafts = this.xlsxDraftEdits(), current = drafts.length ? patchCells(this.state, drafts).model : this.state;
    const patched = patchMoveConditionalFormatRule(current, sheetIndex, priority, direction);
    if (patched.model === current) return false;
    this.state = patched.model;
    this.drafts.clear(); this.styleDrafts.clear(); this.formulaDrafts.clear();
    this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    const historyLabel = direction === 'up' ? 'conditional-format-move-up-history' : 'conditional-format-move-down-history';
    this.emitter.fire({ label: message(historyLabel, this.language), undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } });
    this.renderEmitter.fire();
    return true;
  }

  editConditionalFormatRule(sheetIndex, priority, edit) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || '조건부 서식을 수정할 수 없는 문서입니다.');
    const before = this.captureSnapshot(), drafts = this.xlsxDraftEdits(), current = drafts.length ? patchCells(this.state, drafts).model : this.state;
    const patched = patchEditConditionalFormatRule(current, sheetIndex, priority, edit);
    if (patched.model === current) return false;
    this.state = patched.model;
    this.drafts.clear(); this.styleDrafts.clear(); this.formulaDrafts.clear();
    this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({ label: message('conditional-format-edit-history', this.language), undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } });
    this.renderEmitter.fire();
    return true;
  }

  applyHyperlink(sheetIndex, row, column, link = null) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || '하이퍼링크를 수정할 수 없는 문서입니다.');
    const sheet = this.state.sheets[sheetIndex];
    if (!sheet || !Number.isInteger(row) || row < 0 || row >= XLSX_LIMITS.rows || !Number.isInteger(column) || column < 0 || column >= XLSX_LIMITS.columns) throw new RangeError('하이퍼링크를 설정할 셀 주소가 유효하지 않습니다.');
    const merged = (sheet.mergedCells || []).find(range => row >= range.top && row <= range.bottom && column >= range.left && column <= range.right);
    if (merged && (merged.top !== row || merged.left !== column)) throw new RangeError('병합 셀의 왼쪽 위 셀에만 하이퍼링크를 설정할 수 있습니다.');
    const before = this.captureSnapshot(), drafts = this.xlsxDraftEdits(), current = drafts.length ? patchCells(this.state, drafts).model : this.state;
    const patched = patchHyperlinks(current, sheetIndex, row, column, link);
    if (patched.model === current) return false;
    this.state = patched.model;
    this.drafts.clear(); this.styleDrafts.clear(); this.formulaDrafts.clear();
    this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot(), label = link ? '하이퍼링크 설정' : '하이퍼링크 제거';
    this.emitter.fire({ label, undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } });
    this.renderEmitter.fire();
    return true;
  }

  applyTextLayout(sheetIndex, edits, heights, label = '행 높이 자동 맞춤') {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || '읽기 전용 파일입니다.');
    const before = this.captureSnapshot();
    try {
      const styleChanged = edits.length ? this.applyCellEdits(edits, sheetIndex, label, new Set(), false) : false;
      const patched = patchRowHeightValues(this.state, sheetIndex, heights);
      if (!styleChanged && patched.model === this.state) return;
      this.state = patched.model; this.structureChanged = !this.matchesSavedState();
    } catch (error) { this.restoreSnapshot(before); throw error; }
    const after = this.captureSnapshot();
    this.emitter.fire({ label, undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } });
    this.renderEmitter.fire();
  }

  applyPageBreaks(sheetIndex, operation) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || 'XLSX 페이지 나누기를 저장할 수 없습니다.');
    const before = this.captureSnapshot(), patched = patchPageBreaks(this.state, sheetIndex, operation);
    if (patched.model === this.state) return;
    this.state = patched.model; this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({ label: '페이지 나누기 변경', undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } });
    this.renderEmitter.fire();
  }

  applyHeaderFooter(sheetIndex, settings, clear = false) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || 'XLSX 머리글·바닥글을 저장할 수 없습니다.');
    const before = this.captureSnapshot(), patched = patchHeaderFooter(this.state, sheetIndex, settings, clear);
    if (patched.model === this.state) return;
    this.state = patched.model; this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({ label: '머리글·바닥글 변경', undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } });
    this.renderEmitter.fire();
  }

  applyPrintTitles(sheetIndex, titles) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || 'XLSX 반복 제목을 저장할 수 없습니다.');
    const before = this.captureSnapshot(), patched = patchPrintTitles(this.state, sheetIndex, titles);
    if (patched.model === this.state) return;
    this.state = patched.model; this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({ label: '반복 제목 변경', undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } });
    this.renderEmitter.fire();
  }

  applyPrintArea(sheetIndex, range) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || 'XLSX 인쇄 영역을 저장할 수 없습니다.');
    const before = this.captureSnapshot(), patched = patchPrintArea(this.state, sheetIndex, range);
    if (patched.model === this.state) return;
    this.state = patched.model; this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({ label: '인쇄 영역 변경', undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } });
    this.renderEmitter.fire();
  }

  applyDefinedName(request) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || 'XLSX 이름 정의를 저장할 수 없습니다.');
    const before = this.captureSnapshot(), patched = patchDefinedName(this.state, request);
    if (patched.model === this.state) return;
    this.state = patched.model; this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({ label: message('defined-name-history', this.language), undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } });
    this.renderEmitter.fire();
  }

  applyPrintSettings(sheetIndex, settings) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || 'XLSX 페이지 설정을 저장할 수 없습니다.');
    const before = this.captureSnapshot(), patched = patchPrintSettings(this.state, sheetIndex, settings);
    if (patched.model === this.state) return;
    this.state = patched.model; this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({ label: '페이지 설정 변경', undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } });
    this.renderEmitter.fire();
  }

  applyRowHeights(sheetIndex, startRow, endRow, height) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || 'XLSX 행 높이를 저장할 수 없습니다.');
    const before = this.captureSnapshot(), patched = patchRowHeights(this.state, sheetIndex, startRow, endRow, height);
    if (patched.model === this.state) return;
    this.state = patched.model; this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({ label: '행 높이 변경', undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } });
    this.renderEmitter.fire();
  }

  applyHiddenRows(sheetIndex, rows, hidden) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || 'XLSX 행 숨김을 저장할 수 없습니다.');
    if (!Number.isInteger(sheetIndex) || !this.state.sheets[sheetIndex] || !Array.isArray(rows) || typeof hidden !== 'boolean') throw new RangeError('숨김 행 요청이 유효하지 않습니다.');
    const current = new Set(this.state.sheets[sheetIndex].hiddenRows || []);
    const changedRows = rows.filter(row => hidden ? !current.has(row) : current.has(row));
    if (!changedRows.length) return;
    const before = this.captureSnapshot(), patched = patchHiddenRows(this.state, [{ sheetIndex, rows: [...new Set(changedRows)].sort((a, b) => a - b), hidden }]);
    this.state = patched.model;
    this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({
      label: hidden ? '행 숨기기' : '행 숨김 해제',
      undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); },
      redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); }
    });
    this.renderEmitter.fire();
  }

  applyHiddenColumns(sheetIndex, columns, hidden) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || 'XLSX 열 숨김을 저장할 수 없습니다.');
    if (!Number.isInteger(sheetIndex) || !this.state.sheets[sheetIndex] || !Array.isArray(columns) || typeof hidden !== 'boolean') throw new RangeError('숨김 열 요청이 유효하지 않습니다.');
    const current = new Set(this.state.sheets[sheetIndex].hiddenColumns || []);
    const changed = columns.filter(column => hidden ? !current.has(column) : current.has(column));
    if (!changed.length) return;
    const before = this.captureSnapshot(), patched = patchHiddenColumns(this.state, [{ sheetIndex, columns: [...new Set(changed)].sort((a, b) => a - b), hidden }]);
    this.state = patched.model;
    this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({
      label: hidden ? '열 숨기기' : '열 숨김 해제',
      undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); },
      redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); }
    });
    this.renderEmitter.fire();
  }

  applyOutlineGroups(sheetIndex, axis, action, start, end, summary, level) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || 'XLSX 행·열 그룹을 저장할 수 없습니다.');
    const before = this.captureSnapshot(), patched = patchOutlineGroups(this.state, [{ sheetIndex, axis, action, start, end, summary, level }]);
    if (patched.model === this.state) return;
    this.state = patched.model;
    this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot(), subject = axis === 'row' ? '행' : '열', verb = { group: '그룹화', ungroup: '그룹 해제', collapse: '그룹 접기', expand: '그룹 펼치기', 'collapse-summary': '요약 그룹 모두 접기', 'expand-summary': '요약 그룹 모두 펼치기', 'show-level': '개요 표시 수준 변경' }[action];
    this.emitter.fire({
      label: subject + ' ' + verb,
      undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); },
      redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); }
    });
    this.renderEmitter.fire();
  }

  applyFreezePane(sheetIndex, freezePane) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || 'XLSX 고정 창을 저장할 수 없습니다.');
    if (!Number.isInteger(sheetIndex) || !this.state.sheets[sheetIndex] || !Number.isInteger(freezePane?.rows) || !Number.isInteger(freezePane?.columns)) throw new RangeError('고정 창 요청이 유효하지 않습니다.');
    const current = this.state.sheets[sheetIndex].freezePane || { rows: 0, columns: 0 };
    if (current.rows === freezePane.rows && current.columns === freezePane.columns) return;
    const before = this.captureSnapshot(), patched = patchFreezePanes(this.state, [{ sheetIndex, freezePane }]);
    this.state = patched.model;
    this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({
      label: freezePane.rows || freezePane.columns ? '고정 창 변경' : '고정 창 해제',
      undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); },
      redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); }
    });
    this.renderEmitter.fire();
  }

  applyMergedCells(sheetIndex, range, action, discardContents = false) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || 'XLSX 병합 셀을 저장할 수 없습니다.');
    const sheet = this.state.sheets[sheetIndex];
    if (!Number.isInteger(sheetIndex) || !sheet || !range || !['merge', 'unmerge'].includes(action) || !['top', 'bottom', 'left', 'right'].every(key => Number.isInteger(range[key])) || range.top < 0 || range.left < 0 || range.bottom < range.top || range.right < range.left || range.bottom >= XLSX_LIMITS.rows || range.right >= XLSX_LIMITS.columns || typeof discardContents !== 'boolean') throw new RangeError('병합 셀 요청이 유효하지 않습니다.');
    const currentRanges = [...(sheet.mergedCells || [])];
    let nextRanges = currentRanges, clearEdits = [];
    if (action === 'merge') {
      if (range.top === range.bottom && range.left === range.right) throw new RangeError('두 셀 이상을 선택해야 병합할 수 있습니다.');
      if (currentRanges.some(item => range.top <= item.bottom && range.bottom >= item.top && range.left <= item.right && range.right >= item.left)) throw new Error('선택 범위에 이미 병합된 셀이 포함되어 있습니다. 먼저 병합을 해제하세요.');
      const coordinates = new Map();
      for (const cell of sheet.cells.values()) if (cell.row >= range.top && cell.row <= range.bottom && cell.column >= range.left && cell.column <= range.right && (cell.row !== range.top || cell.column !== range.left)) coordinates.set(cell.row + ':' + cell.column, [cell.row, cell.column]);
      for (const key of [...this.drafts.keys(), ...this.formulaDrafts.keys()]) {
        const prefix = sheetIndex + ':', address = key.startsWith(prefix) ? key.slice(prefix.length) : '';
        if (!address) continue;
        const { row, column } = parseAddress(address);
        if (row >= range.top && row <= range.bottom && column >= range.left && column <= range.right && (row !== range.top || column !== range.left)) coordinates.set(row + ':' + column, [row, column]);
      }
      for (const [row, column] of coordinates.values()) {
        const cell = this.getCell(row, column, sheetIndex);
        if (cell && (cell.value !== '' && cell.value !== null && cell.value !== undefined || isFormulaCell(cell))) clearEdits.push({ sheetIndex, address: addressFor(row + 1, column + 1), value: '', formula: null, kind: 'blank' });
      }
      if (clearEdits.length && !discardContents) {
        const warning = new Error('병합 범위의 왼쪽 위 셀을 제외한 값과 수식이 삭제됩니다.');
        warning.needsConfirmation = true;
        throw warning;
      }
      nextRanges = [...currentRanges, { ...range }];
    } else {
      const removed = currentRanges.filter(item => range.top <= item.bottom && range.bottom >= item.top && range.left <= item.right && range.right >= item.left);
      if (!removed.length) return;
      const removeSet = new Set(removed);
      nextRanges = currentRanges.filter(item => !removeSet.has(item));
    }
    const before = this.captureSnapshot();
    let current = this.state;
    const drafts = this.xlsxDraftEdits();
    if (drafts.length) current = patchCells(current, drafts).model;
    if (clearEdits.length) current = patchCells(current, clearEdits).model;
    const patched = patchMergedCells(current, [{ sheetIndex, ranges: nextRanges }]);
    this.state = patched.model;
    this.drafts.clear(); this.styleDrafts.clear(); this.formulaDrafts.clear();
    this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({
      label: action === 'merge' ? '셀 병합' : '셀 병합 해제',
      undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); },
      redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); }
    });
    this.renderEmitter.fire();
  }

  applyCalculatedColumn(sheetIndex, row, column, mode = 'preserve', remove = false) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || '읽기 전용 파일입니다.');
    const table = this.state.sheets[sheetIndex]?.tables?.find(item => row >= item.range.top && row <= item.range.bottom && column >= item.range.left && column <= item.range.right);
    if (!table || !Number.isInteger(row) || !Number.isInteger(column) || !['preserve', 'all'].includes(mode) || typeof remove !== 'boolean') throw new RangeError('계산 열 요청이 유효하지 않습니다.');
    if (remove) {
      const before = this.captureSnapshot(), patched = patchCalculatedColumns(this.state, [{ sheetIndex, tableName: table.name, column, formula: null }]);
      if (patched.model === this.state) return { tableName: table.name, filled: 0, preserved: 0 };
      this.state = patched.model; this.structureChanged = !this.matchesSavedState(); const after = this.captureSnapshot();
      this.emitter.fire({ label: '계산 열 해제', undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } }); this.renderEmitter.fire();
      return { tableName: table.name, filled: 0, preserved: 0 };
    }
    const first = table.range.top + table.headerRowCount, last = table.range.bottom - table.totalsRowCount, anchor = Math.max(first, Math.min(row, last)), source = row >= first && row <= last ? this.getCell(row, column, sheetIndex) : null;
    const template = table.calculatedColumns?.[column - table.range.left], formula = source?.formula || (template ? translateFormulaReferences(template, anchor - first, 0) : null);
    if (!formula) throw new Error('계산 열에 적용할 수식 셀을 선택하세요.');
    const plan = planCalculatedColumn(this.state, { sheetIndex, row: anchor, column, formula, mode, forceOrigin: Boolean(source?.formula) }, (r, c, index) => this.getCell(r, c, index));
    if (!plan) throw new Error('계산 열에는 데이터 행이 필요합니다.');
    const result = { tableName: table.name, filled: plan.edits.length, preserved: plan.preserved };
    try { if (plan.edits.length) this.applyCellEdits(plan.edits, sheetIndex, '계산 열 수식 적용', new Set(), true, { skipCalculated: true, calculatedMetadata: [plan.metadata] }); }
    catch (error) { error.calculatedColumnResult = result; throw error; }
    return result;
  }

  applyTableTotals(sheetIndex, tableName, request) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || '읽기 전용 파일입니다.');
    const before = this.captureSnapshot(), drafts = this.xlsxDraftEdits(), current = drafts.length ? patchCells(this.state, drafts).model : this.state;
    const patched = patchTableTotals(current, { ...request, sheetIndex, tableName });
    if (Object.keys(patched.model.files).length === Object.keys(this.state.files).length && Object.entries(this.state.files).every(([name, bytes]) => patched.model.files[name] && Buffer.compare(Buffer.from(bytes), Buffer.from(patched.model.files[name])) === 0)) return;
    this.state = patched.model; this.drafts.clear(); this.styleDrafts.clear(); this.formulaDrafts.clear(); this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({ label: '표 합계 변경', undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } });
    this.renderEmitter.fire();
  }

  appendTableRow(sheetIndex, tableName, row, column, value, acceptedWarnings = new Set(), options = {}) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || '읽기 전용 파일입니다.');
    const sheet = this.state.sheets[sheetIndex], table = sheet?.tables.find(item => item.name === tableName);
    const last = table ? table.range.bottom - table.totalsRowCount : -1;
    const visibleColumns = table ? Array.from({ length: table.range.right - table.range.left + 1 }, (_, index) => table.range.left + index).filter(index => !sheet.hiddenColumns.includes(index)) : [];
    if (!table || row !== last || column !== visibleColumns.at(-1) || row < table.range.top + table.headerRowCount && !(table.headerRowCount === 1 && last === table.range.top) || value !== undefined && typeof value !== 'string') throw new RangeError('표 마지막 데이터 셀에서 Tab을 사용하세요.');
    if (table.range.bottom + 1 >= XLSX_LIMITS.rows) throw new RangeError('워크시트 마지막 행에서는 표 행을 추가할 수 없습니다.');
    const before = this.captureSnapshot(), revision = this.renderRevision, shadow = Object.create(this);
    shadow.drafts = new Map(this.drafts); shadow.styleDrafts = new Map(this.styleDrafts); shadow.formulaDrafts = new Map(this.formulaDrafts); shadow.columnWidthDrafts = new Map(this.columnWidthDrafts);
    const ensureCurrent = () => { if (this.renderRevision !== revision || this.savedRevision !== before.savedRevision || this.state !== before.state) throw new Error('문서가 변경되어 표 행 추가를 다시 실행해야 합니다.'); };
    const commit = (clearTableFilters = false) => {
      ensureCurrent(); const drafts = shadow.xlsxDraftEdits();
      let current = drafts.length ? patchCells(shadow.state, drafts).model : shadow.state;
      const target = table.range.bottom + 1, range = { ...table.range, bottom: target };
      const calculator = createFormulaCalculator({ state: current, getCell: (r, c, index = 0) => current.sheets[index]?.cells.get(addressFor(r + 1, c + 1)) || null });
      for (let c = range.left; c <= range.right; c++) {
        const cell = current.sheets[sheetIndex].cells.get(addressFor(target + 1, c + 1));
        if (cell?.formula || (cell?.value ?? '') !== '' || calculator.cell(sheetIndex, target, c)?.isSpill) throw new Error('표 아래에 데이터가 있어 새 행을 추가할 수 없습니다. 빈 행을 확보하세요.');
      }
      if (sheet.tables.some(other => other.name !== tableName && target >= other.range.top && target <= other.range.bottom && range.left <= other.range.right && range.right >= other.range.left) || sheet.mergedCells.some(merged => merged.top <= target && merged.bottom >= last + 1 && merged.left <= range.right && merged.right >= range.left)) throw new Error('새 표 행이 다른 표 또는 병합 셀과 겹칩니다.');
      const requests = [{ sheetIndex, tableName, range, headers: [] }];
      current = prepareTableExpansion(current, requests); current = patchTableRanges(current, requests).model;
      const probe = Object.create(shadow); probe.state = current; probe.drafts = new Map(); probe.styleDrafts = new Map(); probe.formulaDrafts = new Map();
      if ((hiddenRowsForSheet(probe, sheetIndex) || []).includes(last + 1)) {
        const expandedTable = current.sheets[sheetIndex].tables.find(item => item.name === tableName);
        if (!expandedTable?.autoFilter?.columns.length) throw new Error('새 표 행이 수동 숨김이나 다른 필터에 가려집니다. 해당 숨김을 해제한 뒤 추가하세요.');
        const cleared = patchAutoFilters(current, [{ sheetIndex, filter: { ...expandedTable.autoFilter, columns: [] } }], { now: this.now }).model;
        probe.state = cleared;
        if ((hiddenRowsForSheet(probe, sheetIndex) || []).includes(last + 1)) throw new Error('새 표 행이 수동 숨김이나 다른 필터에 가려집니다. 해당 숨김을 해제한 뒤 추가하세요.');
        if (!clearTableFilters) {
          const error = new Error('새 표 행이 현재 필터에 가려집니다. 필터를 해제한 뒤 추가하세요.');
          error.tableFilterPrompt = { tableName };
          error.retry = () => { ensureCurrent(); return commit(true); };
          throw error;
        }
        current = cleared;
      }
      this.state = current; this.drafts = new Map(); this.styleDrafts = new Map(); this.formulaDrafts = new Map(); this.structureChanged = !this.matchesSavedState();
      const after = this.captureSnapshot(); this.emitter.fire({ label: '표 데이터 행 추가', undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } }); this.renderEmitter.fire();
      return { row: last + 1, column: visibleColumns[0] };
    };
    const wrap = error => { if (error.validationPrompt && typeof error.retry === 'function') { const retry = error.retry; error.retry = () => { ensureCurrent(); try { retry(); } catch (next) { throw wrap(next); } return commit(); }; } return error; };
    try {
      if (value !== undefined) {
        const cell = shadow.getCell(row, column, sheetIndex), edit = value.startsWith('=') ? { row, column, value: cell?.cachedValue ?? cell?.value ?? '', formula: value.slice(1) } : { row, column, value, ...(isFormulaCell(cell) ? { formula: null } : {}) };
        shadow.applyCellEdits([edit], sheetIndex, '표 데이터 행 추가', acceptedWarnings, false, { ...options, skipExpansion: true });
      }
    }
    catch (error) { throw wrap(error); }
    return commit();
  }

  applyCreateTable(sheetIndex, range, name, language = 'en') {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || '읽기 전용 파일입니다.');
    const before=this.captureSnapshot(), drafts=this.xlsxDraftEdits(), current=drafts.length?patchCells(this.state,drafts).model:this.state;
    const patched=patchAddTable(current,{sheetIndex,range,name},{now:this.now,language});
    this.state=patched.model; this.drafts=new Map(); this.styleDrafts=new Map(); this.formulaDrafts=new Map(); this.structureChanged=!this.matchesSavedState();
    const after=this.captureSnapshot();
    this.emitter.fire({label:'Excel 표 생성',undo:()=>{this.restoreSnapshot(before);this.renderEmitter.fire()},redo:()=>{this.restoreSnapshot(after);this.renderEmitter.fire()}});this.renderEmitter.fire();
    return {name:this.state.sheets[sheetIndex].tables.at(-1).name};
  }

  applyTableStyle(sheetIndex, tableName, styleInfo) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || 'XLSX 표 스타일을 저장할 수 없습니다.');
    const table = this.state.sheets[sheetIndex]?.tables?.find(item => item.name === tableName);
    if (!table) throw new RangeError('선택한 Excel 표를 찾을 수 없습니다.');
    const before = this.captureSnapshot(), drafts = this.xlsxDraftEdits(), current = drafts.length ? patchCells(this.state, drafts).model : this.state;
    const patched = patchTableStyles(current, [{ sheetIndex, tableName, styleInfo }]);
    if (patched.model === this.state) return;
    this.state = patched.model; this.drafts.clear(); this.styleDrafts.clear(); this.formulaDrafts.clear(); this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({ label: 'Excel 표 스타일 변경', undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } });
    this.renderEmitter.fire();
  }

  applyTableResize(sheetIndex, tableName, range) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || 'XLSX 표 범위를 저장할 수 없습니다.');
    const table = this.state.sheets[sheetIndex]?.tables?.find(item => item.name === tableName);
    if (!table) throw new RangeError('선택한 Excel 표를 찾을 수 없습니다.');
    if (range?.top === table.range.top && range?.bottom === table.range.bottom && range?.left === table.range.left && range?.right === table.range.right) return;
    const before = this.captureSnapshot(), drafts = this.xlsxDraftEdits(), current = drafts.length ? patchCells(this.state, drafts).model : this.state;
    const patched = patchTableRanges(current, [{ sheetIndex, tableName, range }]);
    this.state = patched.model;
    this.drafts.clear(); this.styleDrafts.clear(); this.formulaDrafts.clear();
    this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({ label: 'Excel 표 크기 조정', undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } });
    this.renderEmitter.fire();
  }

  applyAddChart(sheetIndex, range, chartType = 'bar') {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || 'XLSX 차트를 추가할 수 없습니다.');
    const before = this.captureSnapshot(), drafts = this.xlsxDraftEdits(), current = drafts.length ? patchCells(this.state, drafts).model : this.state;
    const patched = patchAddChart(current, sheetIndex, range, chartType);
    this.state = patched.model;
    this.drafts.clear(); this.styleDrafts.clear(); this.formulaDrafts.clear();
    this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({ label: '차트 삽입', undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } });
    this.renderEmitter.fire();
  }

  applyChartPlacement(sheetIndex, chartIndex, placement) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || 'XLSX 차트 배치를 수정할 수 없습니다.');
    const before = this.captureSnapshot(), current = analyze(this.createOutput(), this.state.extension);
    const patched = patchChartPlacement(current, sheetIndex, chartIndex, placement);
    if (patched.model === current) return;
    this.state = patched.model;
    this.drafts.clear(); this.styleDrafts.clear(); this.formulaDrafts.clear(); this.columnWidthDrafts.clear();
    this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({ label: '차트 배치 변경', undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } });
    this.renderEmitter.fire();
  }

  applyChartMetadata(sheetIndex, chartIndex, changes) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || 'XLSX 차트 정보를 수정할 수 없습니다.');
    const before = this.captureSnapshot(), drafts = this.xlsxDraftEdits(), current = drafts.length ? patchCells(this.state, drafts).model : this.state;
    const patched = patchChartMetadata(current, sheetIndex, chartIndex, changes);
    if (patched.model === current) return;
    this.state = patched.model;
    this.drafts.clear(); this.styleDrafts.clear(); this.formulaDrafts.clear();
    this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({ label: '차트 제목·계열 이름 변경', undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } });
    this.renderEmitter.fire();
  }

  applyChartSource(sheetIndex, chartIndex, range) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || 'XLSX 차트 정보를 수정할 수 없습니다.');
    const before = this.captureSnapshot(), drafts = this.xlsxDraftEdits(), current = drafts.length ? patchCells(this.state, drafts).model : this.state;
    const patched = patchChartSource(current, sheetIndex, chartIndex, range);
    if (patched.model === current) return;
    this.state = patched.model;
    this.drafts.clear(); this.styleDrafts.clear(); this.formulaDrafts.clear();
    this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({ label: '차트 원본 범위 변경', undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } });
    this.renderEmitter.fire();
  }

  applyDeleteChart(sheetIndex, chartIndex) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || 'XLSX 차트 정보를 수정할 수 없습니다.');
    const before = this.captureSnapshot(), drafts = this.xlsxDraftEdits(), current = drafts.length ? patchCells(this.state, drafts).model : this.state;
    const patched = patchDeleteChart(current, sheetIndex, chartIndex);
    if (patched.model === current) return;
    this.state = patched.model;
    this.drafts.clear(); this.styleDrafts.clear(); this.formulaDrafts.clear();
    this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({ label: '차트 삭제', undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } });
    this.renderEmitter.fire();
  }

  applySheetVisibility(sheetIndex, visible) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || 'XLSX 워크시트 표시 상태를 저장할 수 없습니다.');
    const before = this.captureSnapshot(), patched = patchSheetVisibility(this.state, sheetIndex, visible);
    if (Buffer.compare(Buffer.from(patched.bytes), Buffer.from(this.state.bytes)) === 0) return;
    this.state = patched.model;
    this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({
      label: visible ? '시트 표시' : '시트 숨기기',
      undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); },
      redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); }
    });
    this.renderEmitter.fire();
  }

  applySheetZoom(sheetIndex, zoomScale) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || 'XLSX 워크시트 확대 비율을 저장할 수 없습니다.');
    const before = this.captureSnapshot(), patched = patchSheetZoom(this.state, sheetIndex, zoomScale);
    if (Buffer.compare(Buffer.from(patched.bytes), Buffer.from(this.state.bytes)) === 0) return;
    this.state = patched.model;
    this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({
      label: '시트 확대/축소 변경',
      undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); },
      redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); }
    });
    this.renderEmitter.fire();
  }

  applySheetName(sheetIndex, name) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || 'XLSX 시트 이름을 저장할 수 없습니다.');
    const before = this.captureSnapshot(), patched = patchSheetName(this.state, sheetIndex, name);
    if (Buffer.compare(Buffer.from(patched.bytes), Buffer.from(this.state.bytes)) === 0) return;
    this.state = patched.model;
    this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({
      label: '시트 이름 변경',
      undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); },
      redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); }
    });
    this.renderEmitter.fire();
  }

  applyAddSheet() {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || 'XLSX 워크시트를 추가할 수 없습니다.');
    const before = this.captureSnapshot(), patched = patchAddSheet(this.state);
    this.state = patched.model;
    this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({
      label: '워크시트 추가',
      undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); },
      redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); }
    });
    this.renderEmitter.fire();
    return patched.sheetIndex;
  }

  applyDeleteSheet(sheetIndex) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || 'XLSX 워크시트를 삭제할 수 없습니다.');
    const before = this.captureSnapshot(), patched = patchDeleteSheet(this.state, sheetIndex);
    this.state = patched.model;
    this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({
      label: '워크시트 삭제',
      undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); },
      redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); }
    });
    this.renderEmitter.fire();
    return patched.sheetIndex;
  }

  applyMoveSheet(sheetIndex, targetIndex) {
    if (!this.writable || this.kind !== 'xlsx') throw new Error(this.reason || 'XLSX 워크시트를 이동할 수 없습니다.');
    const before = this.captureSnapshot(), patched = patchMoveSheet(this.state, sheetIndex, targetIndex);
    if (patched.model === this.state) return targetIndex;
    this.state = patched.model;
    this.structureChanged = !this.matchesSavedState();
    const after = this.captureSnapshot();
    this.emitter.fire({ label: '워크시트 순서 변경', undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); }, redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); } });
    this.renderEmitter.fire();
    return patched.sheetIndex;
  }

  applyStructureChange(operation) {
    if (!this.writable) throw new Error(this.reason || '읽기 전용 파일입니다.');
    const { axis, action, index, count = 1, sheetIndex = 0, emit = true } = operation || {};
    if (!['row', 'column'].includes(axis) || !['insert', 'delete'].includes(action) || !Number.isInteger(index) || index < 0 || !Number.isInteger(count) || count < 1 || count > 10000 || !Number.isInteger(sheetIndex)) throw new RangeError('행·열 구조 변경 요청이 유효하지 않습니다.');
    const before = this.captureSnapshot();
    if (this.kind === 'csv') {
      if (sheetIndex !== 0) throw new RangeError('CSV에는 첫 번째 시트만 있습니다.');
      let state = this.state;
      for (let offset = 0; offset < count; offset++) state = transformCsvStructure(state, path.extname(this.uri.fsPath).toLowerCase(), { axis, action, index: action === 'insert' ? index + offset : index });
      this.state = state;
      this.originalCsvState = this.state;
    } else if (this.kind === 'xlsx') {
      const edits = this.xlsxDraftEdits();
      let current = edits.length ? patchCells(this.state, edits).model : this.state;
      for (let offset = 0; offset < count; offset++) current = transformXlsxStructure(current, { sheetIndex, axis, action, index: action === 'insert' ? index + offset : index }).model;
      this.state = current;
      if (axis === 'column') {
        const shiftedWidths = new Map();
        for (const [key, width] of this.columnWidthDrafts) {
          const [draftSheet, draftColumn] = key.split(':').map(Number);
          let nextColumn = draftColumn;
          if (draftSheet === sheetIndex && action === 'insert' && draftColumn >= index) nextColumn += count;
          else if (draftSheet === sheetIndex && action === 'delete' && draftColumn >= index && draftColumn < index + count) continue;
          else if (draftSheet === sheetIndex && action === 'delete' && draftColumn >= index + count) nextColumn -= count;
          shiftedWidths.set(draftSheet + ':' + nextColumn, width);
        }
        this.columnWidthDrafts = shiftedWidths;
      }
    } else throw new Error('이 문서 형식은 행·열 구조를 편집할 수 없습니다.');
    this.drafts.clear();
    this.styleDrafts.clear();
    this.formulaDrafts.clear();
    this.structureChanged = true;
    if (emit) {
      const after = this.captureSnapshot();
      this.emitter.fire({
        label: action === 'insert' ? '행·열 삽입' : '행·열 삭제',
        undo: () => { this.restoreSnapshot(before); this.renderEmitter.fire(); },
        redo: () => { this.restoreSnapshot(after); this.renderEmitter.fire(); }
      });
    }
    this.renderEmitter.fire();
  }

  captureSnapshot() {
    return { state: this.state, originalCsvState: this.originalCsvState, drafts: new Map(this.drafts), styleDrafts: new Map([...this.styleDrafts].map(([key, value]) => [key, { ...value }])), formulaDrafts: new Map(this.formulaDrafts), columnWidthDrafts: new Map(this.columnWidthDrafts), structureChanged: this.structureChanged, savedRevision: this.savedRevision };
  }

  restoreSnapshot(snapshot) {
    this.state = snapshot.state;
    this.drafts = new Map(snapshot.drafts);
    this.styleDrafts = new Map([...(snapshot.styleDrafts || [])].map(([key, value]) => [key, { ...value }]));
    this.formulaDrafts = new Map(snapshot.formulaDrafts || []);
    this.columnWidthDrafts = new Map(snapshot.columnWidthDrafts || []);
    if (snapshot.savedRevision !== this.savedRevision) {
      if (this.kind === 'csv') {
        this.originalCsvState = this.state;
        this.drafts.clear();
      } else {
        const edits = this.xlsxDraftEdits();
        if (edits.length) this.state = patchCells(this.state, edits).model;
        this.drafts.clear();
        this.styleDrafts.clear();
        this.formulaDrafts.clear();
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

  xlsxDraftEdits() {
    const edits = new Map();
    for (const [key, value] of this.drafts) {
      const [sheet, address] = key.split(':');
      const sheetIndex = Number(sheet), cell = this.state.sheets[sheetIndex]?.cells.get(address);
      const edit = { sheetIndex, address, value };
      if (this.formulaDrafts.has(key)) edit.formula = this.formulaDrafts.get(key);
      else if (cell?.formula) edit.formula = cell.formula;
      edits.set(key, edit);
    }
    for (const [key, style] of this.styleDrafts) {
      const [sheet, address] = key.split(':');
      const existing = edits.get(key), coordinate = parseAddress(address), cell = this.getCell(coordinate.row, coordinate.column, Number(sheet));
      edits.set(key, { ...(existing || { sheetIndex: Number(sheet), address, value: cell?.value ?? '', ...(cell?.formula ? { formula: cell.formula } : {}) }), style: style.index, ...(style.properties ? { styleProperties: style.properties } : {}), ...(style.font ? { font: style.font } : {}), ...(Object.hasOwn(style, 'fillColor') ? { fillColor: style.fillColor } : {}), ...(style.alignment ? { alignment: style.alignment } : {}), ...(style.border ? { border: style.border } : {}), ...(style.kind !== undefined ? { kind: style.kind } : {}), ...(style.numberFormatId !== undefined ? { numberFormatId: style.numberFormatId, numberFormatCode: style.numberFormatCode, ...(style.forceNumberFormat ? { forceNumberFormat: true } : {}) } : {}) });
    }
    for (const [key, formula] of this.formulaDrafts) {
      const [sheet, address] = key.split(':');
      const existing = edits.get(key), coordinate = parseAddress(address), cell = this.getCell(coordinate.row, coordinate.column, Number(sheet));
      edits.set(key, { ...(existing || { sheetIndex: Number(sheet), address, value: '' }), formula });
    }
    return [...edits.values()];
  }

  setDraft(row, column, value, sheetIndex, style, formula, kind, numberFormatId, numberFormatCode) {
    if (this.kind === 'csv') {
      const key = `0:${row}:${column}`;
      const baseline = this.baselineCsvValue(row, column);
      if (value === baseline) this.drafts.delete(key); else this.drafts.set(key, value);
    } else {
      const address = addressFor(row + 1, column + 1);
      const key = `${sheetIndex}:${address}`;
      const baseline = this.state.sheets[sheetIndex].cells.get(address)?.value ?? '';
      if (value === baseline) this.drafts.delete(key); else this.drafts.set(key, value);
      if (formula !== undefined) {
        const baselineFormula = this.state.sheets[sheetIndex].cells.get(address)?.formula ?? null;
        if (formula === baselineFormula) this.formulaDrafts.delete(key); else this.formulaDrafts.set(key, formula);
      }
      const previousStyle = this.styleDrafts.get(key);
      if (style || kind !== undefined || numberFormatId !== undefined) {
        const baselineCell = this.state.sheets[sheetIndex].cells.get(address), baselineStyle = baselineCell?.style ?? null;
        const index = style && Object.hasOwn(style, 'index') ? style.index : previousStyle?.index ?? baselineStyle;
        const properties = style && Object.hasOwn(style, 'properties') ? style.properties : previousStyle?.properties ?? baselineCell?.styleProperties ?? null;
        const font = style && Object.hasOwn(style, 'font') ? style.font : previousStyle?.font;
        const fillColor = style && Object.hasOwn(style, 'fillColor') ? style.fillColor : previousStyle?.fillColor;
        const alignment = style && Object.hasOwn(style, 'alignment') ? { ...(previousStyle?.alignment || {}), ...style.alignment } : previousStyle?.alignment;
        const border = style && Object.hasOwn(style, 'border') ? { ...(previousStyle?.border || {}), ...style.border } : previousStyle?.border;
        const nextNumberFormatId = numberFormatId ?? previousStyle?.numberFormatId;
        const nextNumberFormatCode = numberFormatCode ?? previousStyle?.numberFormatCode;
        const forceNumberFormat = style?.forceNumberFormat || previousStyle?.forceNumberFormat;
        const styleChanged = index !== baselineStyle || JSON.stringify(properties) !== JSON.stringify(baselineCell?.styleProperties ?? null), kindChanged = kind !== undefined && (baselineCell?.kind ?? 'blank') !== kind, numberFormatChanged = nextNumberFormatId !== undefined && baselineCell?.numFmtId !== nextNumberFormatId;
        if (!styleChanged && !kindChanged && !numberFormatChanged) this.styleDrafts.delete(key);
        else this.styleDrafts.set(key, { index, properties, ...(font ? { font } : {}), ...(fillColor !== undefined ? { fillColor } : {}), ...(alignment ? { alignment } : {}), ...(border ? { border } : {}), ...(kindChanged || nextNumberFormatId !== undefined ? { kind: kind ?? baselineCell?.kind ?? 'number' } : {}), ...(nextNumberFormatId !== undefined ? { numberFormatId: nextNumberFormatId } : {}), ...(nextNumberFormatCode ? { numberFormatCode: nextNumberFormatCode } : {}), ...(forceNumberFormat ? { forceNumberFormat: true } : {}) });
      }
    }
  }

  baselineCsvValue(row, column) { return this.originalCsvState.rows[row]?.[column]?.value ?? ''; }

  makeCsvState() {
    const edits = [...this.drafts].map(([key, value]) => {
      const [, row, column] = key.split(':').map(Number);
      return { row, column, value };
    }).filter(({ row, column, value }) => this.originalCsvState.rows[row]?.[column] && value !== this.baselineCsvValue(row, column));
    const text = applyEdits(this.originalCsvState.text, this.originalCsvState.rows, edits, this.originalCsvState.delimiter);
    const state = load(encode(text, this.originalCsvState.encoding, this.originalCsvState.bom), path.extname(this.uri.fsPath), this.originalCsvState.delimiter);
    return state;
  }

  setInitialCsvState(state) { this.originalCsvState = state; this.state = state; }

  async save(cancellationToken) {
    if (!this.dirty && !this.hasRelativeDateFilters()) return;
    if (cancellationToken?.isCancellationRequested) throw new Error('저장이 취소되었습니다.');
    if (this.saving) return this.saving;
    this.saving = this.saveCurrent(cancellationToken);
    try { await this.saving; } finally { this.saving = null; }
  }

  async saveCurrent(cancellationToken) {
    await this.waitForAutoRowFits();
    const sourcePath = this.uri.fsPath;
    const current = await fs.readFile(sourcePath);
    if (hash(current) !== this.originalFingerprint) throw new Error('파일이 외부에서 변경되어 저장을 중단했습니다. 다시 열거나 다른 이름으로 저장하세요.');
    let output;
    if (this.kind === 'csv') {
      output = encode(this.state.text, this.state.encoding, this.state.bom);
      const verified = load(output, path.extname(sourcePath), this.state.delimiter);
      if (!sameRows(verified.rows, this.state.rows) || verified.delimiter !== this.state.delimiter) throw new Error('저장 전 CSV 재파싱 검증에 실패했습니다. 원본을 유지했습니다.');
    } else output = this.createOutput();
    if (!this.dirty && this.kind === 'xlsx' && this.baselineXlsxFiles) {
      const outputFiles = unzipSync(output), baselineFiles = this.baselineXlsxFiles, names = Object.keys(outputFiles);
      if (names.length === Object.keys(baselineFiles).length && names.every(name => baselineFiles[name] && Buffer.compare(Buffer.from(outputFiles[name]), Buffer.from(baselineFiles[name])) === 0)) return;
    }
    if (!this.dirty && Buffer.compare(Buffer.from(output), current) === 0) return;
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
    this.styleDrafts.clear();
    this.formulaDrafts.clear();
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
    await this.waitForAutoRowFits();
    const output = this.dirty || this.hasRelativeDateFilters() ? this.createOutput() : this.originalBytes;
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
    this.styleDrafts.clear();
    this.formulaDrafts.clear();
    this.columnWidthDrafts.clear();
    this.structureChanged = false;
    this.savedRevision++;
    this.renderEmitter.fire();
  }

  hasRelativeDateFilters() {
    return this.writable && this.kind === 'xlsx' && this.state.sheets.some(sheet => [sheet.autoFilter, ...(sheet.tables || []).map(table => table.autoFilter)].some(filter => filter?.columns.some(column => dynamicFilterUsesClock(column.dynamic))));
  }

  createOutput() {
    if (this.kind === 'csv') return encode(this.state.text, this.state.encoding, this.state.bom);
    const patched = patchCells(this.state, this.xlsxDraftEdits());
    const widths = [...this.columnWidthDrafts].map(([key, width]) => { const [sheet, column] = key.split(':').map(Number); return { sheetIndex: sheet, column, width }; });
    let output = widths.length ? patchColumnWidths(patched.model, widths) : patched;
    for (const [sheetIndex, heights] of this.autoRowHeights) output = patchRowHeightValues(output.model, sheetIndex, heights);
    return patchFilterVisibility(output.model, { now: this.now }).bytes;
  }

  async revert() {
    if (this.kind === 'csv') { this.setInitialCsvState(load(this.originalBytes, path.extname(this.uri.fsPath))); this.csvAutoFilter = null; this.csvOutline = { rowOutlineLevels: {}, columnOutlineLevels: {}, collapsedRows: new Set(), collapsedColumns: new Set() }; this.csvHiddenRows = new Set(); this.csvHiddenColumns = new Set(); }
    else { this.state = analyze(this.originalBytes, path.extname(this.uri.fsPath)); this.baselineXlsxFiles = this.state.files; }
    this.drafts.clear();
    this.styleDrafts.clear();
    this.formulaDrafts.clear();
    this.columnWidthDrafts.clear();
    this.structureChanged = false;
    this.savedRevision++;
    this.renderEmitter.fire();
  }

  async backup(context) {
    await this.waitForAutoRowFits();
    const extension = path.extname(this.uri.fsPath) || '.data';
    const id = typeof context.destination === 'string'
      ? path.join(context.destination, `${crypto.randomBytes(12).toString('hex')}${extension}`)
      : context.destination?.fsPath;
    if (typeof id !== 'string' || !id) throw new TypeError('백업 대상 경로가 유효하지 않습니다.');
    const bytes = this.dirty || this.hasRelativeDateFilters() ? this.createOutput() : this.originalBytes;
    await fs.mkdir(path.dirname(id), { recursive: true });
    await fs.writeFile(id, bytes, { flag: 'wx', mode: 0o600 });
    return { id, delete: async () => fs.rm(id, { force: true }) };
  }

  dispose() {
    for (const pending of this.pendingAutoRowFits) pending.cancel();
    this.customDocumentSubscription?.dispose();
    this.onDispose?.();
    this.emitter.dispose();
    this.renderEmitter.dispose();
  }
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

function buildAutoSumFormula(document, row, column, sheetIndex = 0) {
  if (document?.kind !== 'xlsx' || !Number.isInteger(row) || row < 0 || !Number.isInteger(column) || column < 0 || !document.state?.sheets?.[sheetIndex]) return null;
  const containsNumber = (targetRow, targetColumn) => {
    const cell = document.getCell(targetRow, targetColumn, sheetIndex);
    if (!cell) return false;
    if (cell.formula) return true;
    if (cell.kind !== 'number') return false;
    return Number.isFinite(Number(cell.value));
  };
  const range = (startRow, endRow, startColumn, endColumn) => {
    const first = addressFor(startRow + 1, startColumn + 1), last = addressFor(endRow + 1, endColumn + 1);
    const filter = document.state.sheets[sheetIndex].autoFilter;
    const usesSubtotal = Boolean(filter && column >= filter.left && column <= filter.right && startRow > filter.top && endRow <= filter.bottom);
    return usesSubtotal ? `=SUBTOTAL(9,${first}:${last})` : `=SUM(${first}:${last})`;
  };
  let start = row;
  while (start > 0 && containsNumber(start - 1, column)) start--;
  if (start < row) return range(start, row - 1, column, column);
  start = column;
  while (start > 0 && containsNumber(row, start - 1)) start--;
  if (start < column) return range(row, row, start, column - 1);
  return null;
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
  context.subscriptions.push(vscode.commands.registerCommand('jinsheet.undo', () => provider.undoActive()));
  context.subscriptions.push(vscode.commands.registerCommand('jinsheet.redo', () => provider.redoActive()));
}

class JinSheetProvider {
  constructor(context) {
    this.context = context;
    this.language = languageOf(vscode.env.language);
    this.internalClipboard = null;
    this.clipboardWrite = Promise.resolve();
    this.emitter = new vscode.EventEmitter();
    this.onDidChangeCustomDocument = this.emitter.event;
    this.documents = new Map();
    this.editHistory = new WeakMap();
  }

  queueClipboardOperation(task) {
    const operation = this.clipboardWrite.then(task);
    // A failed copy is reported to its caller; later operations can still proceed.
    this.clipboardWrite = operation.then(() => undefined, () => undefined);
    return operation;
  }

  writeClipboardText(text, writer = value => vscode.env.clipboard.writeText(value)) { return this.queueClipboardOperation(() => writer(text)); }

  readClipboardText(reader = () => vscode.env.clipboard.readText()) { return this.queueClipboardOperation(reader); }

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
    this.documents.set(document.uri.toString(), document);
    document.onDispose = () => {
      const key = document.uri.toString();
      if (this.documents.get(key) === document) this.documents.delete(key);
    };
    this.editHistory.set(document, { undo: [], redo: [], applying: false });
    const sub = document.onDidChange(change => {
      const history = this.editHistory.get(document);
      if (!history?.applying) {
        history?.undo.push(change);
        if (history) history.redo.length = 0;
      }
      this.emitter.fire({ document, label: translateDiagnostic(change.label, this.language), undo: change.undo, redo: change.redo });
    });
    document.customDocumentSubscription = sub;
    return document;
  }

  async resolveCustomEditor(document, panel) {
    let selectedSheet = document.kind === 'xlsx' ? Math.max(0, document.state.sheets.findIndex(sheet => sheet.visible)) : 0, fitSequence = 0, pendingFit = null;
    let autoFitRevision = -1;
    const automaticFits = new Map();
    const postModel = reason => {
      if (reason === 'row-layout') {
        const sheet = document.state.sheets[selectedSheet];
        panel.webview.postMessage({ type: 'row-layout', sheetIndex: selectedSheet, rowHeights: { ...(sheet.rowHeights || {}), ...Object.fromEntries((document.autoRowHeights.get(selectedSheet) || []).map(item => [item.row, item.height * 4 / 3])) } });
        return;
      }
      let model;
      try { model = viewModel(document, selectedSheet, this.language); }
      catch (error) {
        panel.webview.postMessage({ type: 'error', message: translateDiagnostic(error.message || error, this.language) });
        model = viewModel(document, selectedSheet, this.language);
      }
      panel.webview.postMessage(model);
      if (document.kind === 'xlsx' && document.writable && autoFitRevision !== document.renderRevision) {
        autoFitRevision = document.renderRevision;
        automaticFits.clear();
        for (let sheetIndex = 0; sheetIndex < document.state.sheets.length; sheetIndex++) {
          let fit;
          // Keep ordinary editing available beyond the bounded measurement budget.
          try { fit = buildAutomaticRowFitRequest(document, sheetIndex, this.language); } catch { continue; }
          if (!fit.rows.length) continue;
          const requestId = ++fitSequence;
          const pending = document.beginAutoRowFit(sheetIndex, fit.rows.map(item => item.row));
          automaticFits.set(requestId, pending);
          pending.promise.then(() => automaticFits.delete(requestId));
          panel.webview.postMessage({ type: 'row-fit-request', background: true, requestId, sheetIndex, ...fit });
        }
      }
    };
    const mediaRoot = vscode.Uri.joinPath(this.context.extensionUri, 'media');
    panel.webview.options = { enableScripts: true, localResourceRoots: [mediaRoot] };
    panel.webview.html = renderHtml(document, panel.webview, this.context.extensionUri, this.language);
    panel.webview.onDidReceiveMessage(async message => {
      if (!message || typeof message !== 'object' || Array.isArray(message)) return;
      let replacementContext = null;
      try {
        const messageSheet = message.sheetIndex;
        if (message.tableAutoFill !== undefined && typeof message.tableAutoFill !== 'boolean') throw new RangeError('계산 열 요청이 유효하지 않습니다.');
        if (message.tableAutoExpand !== undefined && typeof message.tableAutoExpand !== 'boolean') throw new RangeError('표 자동 확장 요청이 유효하지 않습니다.');
        const cellOptions = { skipCalculated: message.tableAutoFill === false, skipExpansion: message.tableAutoExpand === false };
        const currentSheetMessage = matchesSelectedSheet(message, selectedSheet);
        const workbookMessage = message.scope === 'workbook' && document.kind === 'xlsx' && isValidSheetRequest(document, messageSheet);
        if (message.type === 'manage-defined-names' && currentSheetMessage && document.kind === 'xlsx') {
          const localized = messagesFor(this.language), sheet = document.state.sheets[messageSheet];
          const names = document.state.definedNames.filter(item => !item.printArea && !item.printTitles && !item.name.toLocaleLowerCase().startsWith('_xlnm.'));
          const addItem = { label: localized['defined-name-add'], action: 'add' };
          const selected = await vscode.window.showQuickPick([
            ...names.map(item => ({ label: item.name, description: item.localSheetId === null ? localized['defined-name-workbook-scope'] : document.state.sheets[item.localSheetId]?.name, detail: item.source ? definedNameReferenceLabel(item) : '', action: 'select', definition: item })),
            addItem
          ], { placeHolder: localized['defined-name-select'] });
          if (!selected) return;
          if (selected.action === 'add') {
            const name = await vscode.window.showInputBox({ prompt: localized['defined-name-name-prompt'], placeHolder: 'SalesTotal', maxLength: 255 });
            if (name === undefined) return;
            const reference = await vscode.window.showInputBox({ prompt: localized['defined-name-reference-prompt'], placeHolder: "'" + sheet.name.replaceAll("'", "''") + "'!$A$1:$A$10" });
            if (reference === undefined) return;
            const source = parseRangeReference(reference.trim(), sheet.name);
            if (!source) throw new Error(localized['defined-name-reference-invalid']);
            const scope = await vscode.window.showQuickPick([
              { label: localized['defined-name-workbook-scope'], localSheetId: null },
              { label: localized['defined-name-sheet-scope'].replace('{sheet}', sheet.name), localSheetId: messageSheet }
            ], { placeHolder: localized['defined-name-scope-prompt'] });
            if (!scope) return;
            document.applyDefinedName({ action: 'set', name: name.trim(), localSheetId: scope.localSheetId, source });
          } else {
            const definition = selected.definition;
            const action = await vscode.window.showQuickPick([
              { label: localized['defined-name-edit'], action: 'edit' },
              { label: localized['defined-name-delete'], action: 'delete' }
            ], { placeHolder: localized['defined-name-action'].replace('{name}', definition.name) });
            if (!action) return;
            if (action.action === 'delete') {
              const confirmation = await vscode.window.showWarningMessage(localized['defined-name-delete-confirm'].replace('{name}', definition.name), { modal: true }, localized['defined-name-delete']);
              if (confirmation !== localized['defined-name-delete']) return;
              document.applyDefinedName({ action: 'delete', name: definition.name, localSheetId: definition.localSheetId });
            } else {
              const reference = await vscode.window.showInputBox({ prompt: localized['defined-name-reference-prompt'], value: definedNameReferenceLabel(definition), placeHolder: "'" + sheet.name.replaceAll("'", "''") + "'!$A$1:$A$10" });
              if (reference === undefined) return;
              const source = parseRangeReference(reference.trim(), sheet.name);
              if (!source) throw new Error(localized['defined-name-reference-invalid']);
              document.applyDefinedName({ action: 'set', name: definition.name, localSheetId: definition.localSheetId, source });
            }
          }
          postModel();
        } else if (message.type === 'manage-data-validation' && currentSheetMessage && document.kind === 'xlsx') {
          const range = message.range;
          if (!range || ![range.top, range.bottom, range.left, range.right].every(Number.isInteger) || range.top < 0 || range.left < 0 || range.bottom < range.top || range.right < range.left || range.bottom >= XLSX_LIMITS.rows || range.right >= XLSX_LIMITS.columns) throw new RangeError('데이터 유효성 적용 범위가 유효하지 않습니다.');
          const korean = this.language === 'ko', typeItems = [
            { label: korean ? '목록' : 'List', type: 'list' },
            { label: korean ? '정수' : 'Whole number', type: 'whole' },
            { label: korean ? '소수' : 'Decimal', type: 'decimal' },
            { label: korean ? '날짜' : 'Date', type: 'date' },
            { label: korean ? '시간' : 'Time', type: 'time' },
            { label: korean ? '텍스트 길이' : 'Text length', type: 'textLength' },
            { label: korean ? '유효성 규칙 지우기' : 'Clear validation', type: 'clear' }
          ];
          const selectedType = await vscode.window.showQuickPick(typeItems, { placeHolder: korean ? '선택한 셀의 입력을 제한할 방법을 고르세요' : 'Choose how to restrict entries in the selected cells' });
          if (!selectedType) return;
          if (selectedType.type === 'clear') {
            document.applyDataValidation(messageSheet, range, null); postModel(); return;
          }
          let input = { type: selectedType.type };
          if (input.type === 'list') {
            input.source = await vscode.window.showInputBox({ prompt: korean ? '쉼표로 구분한 값, 셀 범위(예: =$A$1:$A$5) 또는 정의된 이름' : 'Comma-separated values, a cell range (for example =$A$1:$A$5), or a defined name', placeHolder: korean ? '승인,대기,반려' : 'Approved,Pending,Rejected' });
            if (input.source === undefined) return;
          } else {
            const operatorItems = [
              ['between', korean ? '사이' : 'between'], ['notBetween', korean ? '사이가 아님' : 'not between'], ['equal', korean ? '같음' : 'equal to'], ['notEqual', korean ? '같지 않음' : 'not equal to'],
              ['greaterThan', korean ? '보다 큼' : 'greater than'], ['lessThan', korean ? '보다 작음' : 'less than'], ['greaterThanOrEqual', korean ? '크거나 같음' : 'greater than or equal to'], ['lessThanOrEqual', korean ? '작거나 같음' : 'less than or equal to']
            ].map(([value, label]) => ({ value, label }));
            const selectedOperator = await vscode.window.showQuickPick(operatorItems, { placeHolder: korean ? '비교 조건을 고르세요' : 'Choose a comparison rule' });
            if (!selectedOperator) return;
            input.operator = selectedOperator.value;
            const examples = { whole: '10', decimal: '0.5', date: '2026-10-17', time: '17:30' };
            const prompt = korean ? '기준값을 입력하세요. 수식은 =로 시작합니다.' : 'Enter a limit. Start with = to use a formula.';
            input.first = await vscode.window.showInputBox({ prompt, placeHolder: examples[input.type] || '10' });
            if (input.first === undefined) return;
            if (['between', 'notBetween'].includes(input.operator)) {
              input.second = await vscode.window.showInputBox({ prompt: korean ? '두 번째 기준값을 입력하세요.' : 'Enter the second limit.', placeHolder: examples[input.type] || '20' });
              if (input.second === undefined) return;
            }
          }
          const rule = buildDataValidationRule(document, messageSheet, input);
          document.applyDataValidation(messageSheet, range, rule); postModel();
        } else if (message.type === 'manage-cell-value-rule' && currentSheetMessage && document.kind === 'xlsx') {
          const range = message.range, localized = messagesFor(this.language);
          if (!range || ![range.top, range.bottom, range.left, range.right].every(Number.isInteger) || range.top < 0 || range.left < 0 || range.bottom < range.top || range.right < range.left || range.bottom >= XLSX_LIMITS.rows || range.right >= XLSX_LIMITS.columns) throw new RangeError(localized['error-conditional-format-range']);
          const operators = ['equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual', 'between', 'notBetween'];
          const operator = await vscode.window.showQuickPick(operators.map(value => ({ label: localized['conditional-format-operator-' + value], value })), { placeHolder: localized['conditional-format-cell-rule-operator-prompt'] });
          if (!operator) return;
          const validateValue = input => {
            const text = String(input).trim();
            return /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(text) && Number.isFinite(Number(text)) ? undefined : localized['error-conditional-format-value'];
          };
          const first = await vscode.window.showInputBox({ prompt: localized['conditional-format-cell-rule-value-prompt'], validateInput: validateValue });
          if (first === undefined) return;
          const values = [Number(first)];
          if (operator.value === 'between' || operator.value === 'notBetween') {
            const second = await vscode.window.showInputBox({ prompt: localized['conditional-format-cell-rule-value-2-prompt'], validateInput: validateValue });
            if (second === undefined) return;
            values.push(Number(second));
          }
          document.applyCellValueConditionalFormat(messageSheet, range, operator.value, values); postModel();
        } else if (message.type === 'manage-data-bars' && currentSheetMessage && document.kind === 'xlsx') {
          const range = message.range;
          const localized = messagesFor(this.language);
          if (!range || ![range.top, range.bottom, range.left, range.right].every(Number.isInteger) || range.top < 0 || range.left < 0 || range.bottom < range.top || range.right < range.left || range.bottom >= XLSX_LIMITS.rows || range.right >= XLSX_LIMITS.columns) throw new RangeError(localized['error-conditional-format-range']);
          const options = [
            { label: localized['conditional-format-blue'], color: '#4472C4' },
            { label: localized['conditional-format-green'], color: '#70AD47' },
            { label: localized['conditional-format-orange'], color: '#ED7D31' }
          ];
          const selectedColor = await vscode.window.showQuickPick(options, { placeHolder: localized['conditional-format-color-prompt'] });
          if (!selectedColor) return;
          document.applyDataBars(messageSheet, range, selectedColor.color); postModel();
        } else if (message.type === 'manage-color-scale' && currentSheetMessage && document.kind === 'xlsx') {
          const range = message.range;
          const localized = messagesFor(this.language);
          if (!range || ![range.top, range.bottom, range.left, range.right].every(Number.isInteger) || range.top < 0 || range.left < 0 || range.bottom < range.top || range.right < range.left || range.bottom >= XLSX_LIMITS.rows || range.right >= XLSX_LIMITS.columns) throw new RangeError(localized['error-conditional-format-range']);
          const options = [
            { label: localized['conditional-format-scale-red-yellow-green'], colors: ['#F8696B', '#FFEB84', '#63BE7B'] },
            { label: localized['conditional-format-scale-green-yellow-red'], colors: ['#63BE7B', '#FFEB84', '#F8696B'] },
            { label: localized['conditional-format-scale-white-blue'], colors: ['#F2F2F2', '#5A8AC6'] }
          ];
          const selectedScale = await vscode.window.showQuickPick(options, { placeHolder: localized['conditional-format-scale-prompt'] });
          if (!selectedScale) return;
          document.applyColorScale(messageSheet, range, selectedScale.colors); postModel();
        } else if (message.type === 'apply-duplicate-values' && currentSheetMessage && document.kind === 'xlsx') {
          const range = message.range, localized = messagesFor(this.language);
          if (!range || ![range.top, range.bottom, range.left, range.right].every(Number.isInteger) || range.top < 0 || range.left < 0 || range.bottom < range.top || range.right < range.left || range.bottom >= XLSX_LIMITS.rows || range.right >= XLSX_LIMITS.columns) throw new RangeError(localized['error-conditional-format-range']);
          document.applyDuplicateValues(messageSheet, range); postModel();
        } else if (message.type === 'manage-icon-set' && currentSheetMessage && document.kind === 'xlsx') {
          const range = message.range;
          const localized = messagesFor(this.language);
          if (!range || ![range.top, range.bottom, range.left, range.right].every(Number.isInteger) || range.top < 0 || range.left < 0 || range.bottom < range.top || range.right < range.left || range.bottom >= XLSX_LIMITS.rows || range.right >= XLSX_LIMITS.columns) throw new RangeError(localized['error-conditional-format-range']);
          const choices = Object.entries(ICON_SET_LENGTHS).map(([name, count]) => ({
            label: localized['conditional-format-icon-style-' + name],
            description: localized['conditional-format-icon-count'].replace('{count}', String(count)),
            value: name
          }));
          const selectedSet = await vscode.window.showQuickPick(choices, { placeHolder: localized['conditional-format-icon-style-prompt'] });
          if (!selectedSet) return;
          document.applyIconSet(messageSheet, range, selectedSet.value); postModel();
        } else if (message.type === 'manage-conditional-format-rules' && currentSheetMessage && document.kind === 'xlsx') {
          const localized = messagesFor(this.language), rules = document.state.sheets[messageSheet]?.conditionalFormats || [];
          if (!rules.length) { await vscode.window.showInformationMessage(localized['conditional-format-no-rules']); return; }
          const selectedRule = await vscode.window.showQuickPick(rules.map(rule => conditionalFormatRuleOption(rule, localized)), { placeHolder: localized['conditional-format-manage-prompt'] });
          if (!selectedRule) return;
          const index = rules.findIndex(rule => rule.priority === selectedRule.priority), rule = rules[index], actions = [];
          if (rule && ['cellIs', 'expression', 'top10'].includes(rule.type)) actions.push({ label: localized['conditional-format-rule-edit-condition'], value: 'condition' });
          if (rule && ['cellIs', 'expression'].includes(rule.type)) {
            actions.push({ label: localized['conditional-format-rule-edit-font-color'], value: 'font-color' });
            actions.push({ label: localized['conditional-format-rule-edit-fill-color'], value: 'fill-color' });
            if (rule.style?.color) actions.push({ label: localized['conditional-format-rule-remove-font-color'], value: 'remove-font-color' });
            if (rule.style?.fill) actions.push({ label: localized['conditional-format-rule-remove-fill-color'], value: 'remove-fill-color' });
          }
          if (rule && ['dataBar', 'colorScale'].includes(rule.type)) actions.push({ label: localized['conditional-format-rule-edit-colors'], value: 'colors' });
          if (rule?.type === 'dataBar') actions.push({ label: localized['conditional-format-rule-edit-data-bar-display'], value: 'data-bar-display' });
          if (rule?.type === 'colorScale' && rule.colorScale?.stops?.length === 3) actions.push({ label: localized['conditional-format-rule-edit-threshold'], value: 'threshold' });
          if (rule?.type === 'iconSet') actions.push({ label: localized['conditional-format-rule-edit-icon-set'], value: 'icon-set' });
          if (index > 0) actions.push({ label: localized['conditional-format-rule-move-up'], value: 'up' });
          if (index >= 0 && index < rules.length - 1) actions.push({ label: localized['conditional-format-rule-move-down'], value: 'down' });
          actions.push({ label: localized['conditional-format-rule-remove'], value: 'remove' });
          const selectedAction = await vscode.window.showQuickPick(actions, { placeHolder: localized['conditional-format-rule-action-prompt'] });
          if (!selectedAction) return;
          if (selectedAction.value === 'remove') document.deleteConditionalFormatRule(messageSheet, selectedRule.priority);
          else if (selectedAction.value === 'up' || selectedAction.value === 'down') document.moveConditionalFormatRule(messageSheet, selectedRule.priority, selectedAction.value);
          else if (['font-color', 'fill-color'].includes(selectedAction.value)) {
            const property = selectedAction.value === 'font-color' ? 'fontColor' : 'fillColor';
            const color = await vscode.window.showInputBox({
              prompt: localized['conditional-format-' + property.replace(/[A-Z]/g, letter => '-' + letter.toLowerCase()) + '-prompt'],
              value: rule.style?.[property === 'fontColor' ? 'color' : 'fill'] || (property === 'fontColor' ? '#000000' : '#FFFF00'),
              validateInput: value => /^#[0-9a-f]{6}$/i.test(String(value).trim()) ? undefined : localized['error-conditional-format-color-value']
            });
            if (color === undefined) return;
            document.editConditionalFormatRule(messageSheet, selectedRule.priority, { kind: 'differential-color', property, color });
          } else if (['remove-font-color', 'remove-fill-color'].includes(selectedAction.value)) {
            document.editConditionalFormatRule(messageSheet, selectedRule.priority, {
              kind: 'differential-color',
              property: selectedAction.value === 'remove-font-color' ? 'fontColor' : 'fillColor',
              color: null
            });
          } else if (selectedAction.value === 'threshold' && rule.type === 'colorScale' && rule.colorScale?.stops?.length === 3) {
            const stop = rule.colorScale.stops[1];
            const selectedType = await vscode.window.showQuickPick(['num', 'percent', 'percentile'].map(value => ({ label: localized['conditional-format-threshold-type-' + value], value })), { placeHolder: localized['conditional-format-threshold-type-prompt'] });
            if (!selectedType) return;
            const rawValue = await vscode.window.showInputBox({
              prompt: localized['conditional-format-threshold-value-prompt'],
              value: String(stop.value ?? 50),
              validateInput: input => {
                const value = String(input).trim(), number = Number(value);
                if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(value) || !Number.isFinite(number) || selectedType.value !== 'num' && (number < 0 || number > 100)) return localized['error-conditional-format-threshold'];
                return undefined;
              }
            });
            if (rawValue === undefined) return;
            document.editConditionalFormatRule(messageSheet, selectedRule.priority, { kind: 'threshold', stop: { type: selectedType.value, value: rawValue } });
          } else if (selectedAction.value === 'data-bar-display' && rule.type === 'dataBar') {
            const current = rule.dataBar;
            const properties = [
              { label: localized['conditional-format-data-bar-min-length'], value: 'minLength' },
              { label: localized['conditional-format-data-bar-max-length'], value: 'maxLength' },
              { label: localized['conditional-format-data-bar-show-value'], value: 'showValue' },
              { label: localized['conditional-format-data-bar-fill'], value: 'gradient' }
            ];
            const selectedProperty = await vscode.window.showQuickPick(properties, { placeHolder: localized['conditional-format-data-bar-display-prompt'] });
            if (!selectedProperty) return;
            const display = { minLength: current.minLength, maxLength: current.maxLength, showValue: current.showValue, gradient: current.gradient };
            if (selectedProperty.value === 'minLength' || selectedProperty.value === 'maxLength') {
              const minimum = selectedProperty.value === 'minLength', property = selectedProperty.value;
              const value = await vscode.window.showInputBox({
                prompt: localized['conditional-format-data-bar-' + (minimum ? 'min-length' : 'max-length') + '-prompt'],
                value: String(current[property]),
                validateInput: input => {
                  const text = String(input).trim();
                  if (!/^(?:0|[1-9]\d?|100)$/.test(text)) return localized['error-conditional-format-data-bar-display'];
                  const number = Number(text), minLength = minimum ? number : current.minLength, maxLength = minimum ? current.maxLength : number;
                  return minLength <= maxLength ? undefined : localized['error-conditional-format-data-bar-display'];
                }
              });
              if (value === undefined) return;
              display[property] = Number(value);
            } else {
              const isValue = selectedProperty.value === 'showValue';
              const choices = isValue
                ? [{ label: localized['conditional-format-data-bar-show-values'], value: true }, { label: localized['conditional-format-data-bar-hide-values'], value: false }]
                : [{ label: localized['conditional-format-data-bar-gradient'], value: true }, { label: localized['conditional-format-data-bar-solid'], value: false }];
              const choice = await vscode.window.showQuickPick(choices, { placeHolder: localized['conditional-format-data-bar-' + (isValue ? 'show-value' : 'fill') + '-prompt'] });
              if (!choice) return;
              display[selectedProperty.value] = choice.value;
            }
            document.editConditionalFormatRule(messageSheet, selectedRule.priority, { kind: 'data-bar-display', dataBar: display });
          } else if (selectedAction.value === 'icon-set' && rule.type === 'iconSet') {
            const current = rule.iconSet;
            const thresholdDescription = stop => {
              const type = localized['conditional-format-icon-threshold-type-' + stop.type];
              const value = stop.type === 'formula' ? '=' + stop.formula : stop.type === 'min' || stop.type === 'max' ? '' : String(stop.value);
              return type + (value ? ': ' + value : '') + (stop.gte ? ' ≥' : ' >');
            };
            const properties = [
              { label: localized['conditional-format-icon-set-mapping'], value: 'mapping' },
              ...(!current.customIcons ? [{ label: localized['conditional-format-icon-set-style'], description: localized['conditional-format-icon-style-' + current.name], value: 'style' }] : []),
              { label: localized['conditional-format-icon-set-show-value'], description: current.showValue ? localized['conditional-format-icon-show-values'] : localized['conditional-format-icon-hide-values'], value: 'showValue' },
              { label: localized['conditional-format-icon-set-reverse'], description: current.reverse ? localized['conditional-format-icon-order-reversed'] : localized['conditional-format-icon-order-normal'], value: 'reverse' },
              ...current.thresholds.map((stop, index) => ({ label: localized['conditional-format-icon-threshold-option'].replace('{index}', String(index + 1)), description: thresholdDescription(stop), value: 'threshold:' + index }))
            ];
            const selectedProperty = await vscode.window.showQuickPick(properties, { placeHolder: localized['conditional-format-icon-set-prompt'] });
            if (!selectedProperty) return;
            if (selectedProperty.value === 'mapping') {
              const customIcons = [];
              for (let index = 0; index < current.thresholds.length; index++) {
                const existing = current.customIcons?.[index] || { iconSet: current.name, iconId: index };
                const setChoices = Object.entries(ICON_SET_LENGTHS).map(([name, count]) => ({ label: localized['conditional-format-icon-style-' + name], value: name, description: localized['conditional-format-icon-count'].replace('{count}', String(count)) }));
                const selectedSet = await vscode.window.showQuickPick(setChoices, { placeHolder: localized['conditional-format-icon-map-set'].replace('{index}', String(index + 1)) });
                if (!selectedSet) return;
                const count = ICON_SET_LENGTHS[selectedSet.value];
                const iconChoices = Array.from({ length: count }, (_, iconId) => ({ label: localized['conditional-format-icon-map-choice'].replace('{index}', String(iconId + 1)), description: iconId === existing.iconId && selectedSet.value === existing.iconSet ? localized['conditional-format-icon-map-current'] : '', value: iconId }));
                const selectedIcon = await vscode.window.showQuickPick(iconChoices, { placeHolder: localized['conditional-format-icon-map-id'].replace('{set}', localized['conditional-format-icon-style-' + selectedSet.value]).replace('{index}', String(index + 1)) });
                if (!selectedIcon) return;
                customIcons.push({ iconSet: selectedSet.value, iconId: selectedIcon.value });
              }
              document.editConditionalFormatRule(messageSheet, selectedRule.priority, { kind: 'icon-set-mapping', customIcons });
            } else if (selectedProperty.value === 'style') {
              const choices = Object.entries(ICON_SET_LENGTHS).map(([name, count]) => ({
                label: localized['conditional-format-icon-style-' + name],
                description: count === current.thresholds.length
                  ? localized['conditional-format-icon-style-preserve-thresholds']
                  : localized['conditional-format-icon-style-reset-thresholds'].replace('{count}', String(count)),
                value: name
              }));
              const selectedStyle = await vscode.window.showQuickPick(choices, { placeHolder: localized['conditional-format-icon-style-prompt'] });
              if (!selectedStyle) return;
              document.editConditionalFormatRule(messageSheet, selectedRule.priority, { kind: 'icon-set-style', name: selectedStyle.value });
            } else if (selectedProperty.value === 'showValue' || selectedProperty.value === 'reverse') {
              const showValue = selectedProperty.value === 'showValue';
              const choices = showValue
                ? [{ label: localized['conditional-format-icon-show-values'], value: true }, { label: localized['conditional-format-icon-hide-values'], value: false }]
                : [{ label: localized['conditional-format-icon-order-normal'], value: false }, { label: localized['conditional-format-icon-order-reversed'], value: true }];
              const choice = await vscode.window.showQuickPick(choices, { placeHolder: localized['conditional-format-icon-set-prompt'] });
              if (!choice) return;
              document.editConditionalFormatRule(messageSheet, selectedRule.priority, { kind: 'icon-set-display', iconSet: { showValue: showValue ? choice.value : current.showValue, reverse: showValue ? current.reverse : choice.value } });
            } else {
              const index = Number(selectedProperty.value.slice('threshold:'.length)), stop = current.thresholds[index];
              if (!Number.isInteger(index) || !stop) throw new Error(localized['error-conditional-format-icon-threshold']);
              const types = ['min', 'max', 'num', 'percent', 'percentile', 'formula'];
              const selectedType = await vscode.window.showQuickPick(types.map(value => ({ label: localized['conditional-format-icon-threshold-type-' + value], value })), { placeHolder: localized['conditional-format-icon-threshold-type-prompt'] });
              if (!selectedType) return;
              const updatedStop = { type: selectedType.value };
              if (selectedType.value === 'formula') {
                const formula = await vscode.window.showInputBox({
                  prompt: localized['conditional-format-icon-threshold-formula-prompt'],
                  value: stop.type === 'formula' ? '=' + stop.formula : '',
                  validateInput: value => {
                    const expression = String(value).trim().replace(/^=/, '');
                    if (!expression || expression.length > 256 || /[!\[\]]/.test(expression)) return localized['error-conditional-format-icon-threshold'];
                    const sample = evaluateFormula(expression, { getCell: () => 1 });
                    return sample === '#NAME?' || sample === '#VALUE!' ? localized['error-conditional-format-icon-threshold'] : undefined;
                  }
                });
                if (formula === undefined) return;
                updatedStop.formula = formula;
              } else if (selectedType.value !== 'min' && selectedType.value !== 'max') {
                const value = await vscode.window.showInputBox({
                  prompt: localized['conditional-format-icon-threshold-value-prompt'],
                  value: stop.type === selectedType.value ? String(stop.value) : '',
                  validateInput: input => {
                    const text = String(input).trim(), number = Number(text);
                    if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(text) || !Number.isFinite(number) || selectedType.value !== 'num' && (number < 0 || number > 100)) return localized['error-conditional-format-icon-threshold'];
                    return undefined;
                  }
                });
                if (value === undefined) return;
                updatedStop.value = value;
              }
              const comparison = await vscode.window.showQuickPick([
                { label: localized['conditional-format-icon-threshold-gte'], value: true },
                { label: localized['conditional-format-icon-threshold-gt'], value: false }
              ], { placeHolder: localized['conditional-format-icon-threshold-comparison-prompt'] });
              if (!comparison) return;
              updatedStop.gte = comparison.value;
              document.editConditionalFormatRule(messageSheet, selectedRule.priority, { kind: 'icon-set-threshold', index, stop: updatedStop });
            }
          } else if (selectedAction.value === 'condition' && rule.type === 'top10') {
            const direction = await vscode.window.showQuickPick([
              { label: localized['conditional-format-top'], value: false },
              { label: localized['conditional-format-bottom'], value: true }
            ], { placeHolder: localized['conditional-format-top-bottom-direction-prompt'] });
            if (!direction) return;
            const percent = await vscode.window.showQuickPick([
              { label: localized['conditional-format-count'], value: false },
              { label: localized['conditional-format-percent'], value: true }
            ], { placeHolder: localized['conditional-format-top-bottom-unit-prompt'] });
            if (!percent) return;
            const value = await vscode.window.showInputBox({
              prompt: localized['conditional-format-top-bottom-value-prompt'],
              value: String(Math.min(rule.top10.rank, percent.value ? 100 : 1000)),
              validateInput: input => {
                const text = String(input).trim(), maximum = percent.value ? 100 : 1000;
                return /^(?:[1-9]\d*)$/.test(text) && Number(text) <= maximum ? undefined : localized['error-conditional-format-top-bottom-value'].replace('{maximum}', String(maximum));
              }
            });
            if (value === undefined) return;
            document.editConditionalFormatRule(messageSheet, selectedRule.priority, { kind: 'condition', top10: { rank: Number(value), percent: percent.value, bottom: direction.value } });
          } else if (selectedAction.value === 'condition' && rule.type === 'cellIs') {
            const operators = ['equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual', 'between', 'notBetween'];
            const operator = await vscode.window.showQuickPick(operators.map(value => ({ label: localized['conditional-format-operator-' + value], value })), { placeHolder: localized['conditional-format-condition-operator-prompt'] });
            if (!operator) return;
            const validateValue = value => /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(String(value).trim()) && Number.isFinite(Number(value)) ? undefined : localized['error-conditional-format-value'];
            const first = await vscode.window.showInputBox({ prompt: localized['conditional-format-condition-value-1-prompt'], value: String(rule.values[0]), validateInput: validateValue });
            if (first === undefined) return;
            const values = [first];
            if (operator.value === 'between' || operator.value === 'notBetween') {
              const second = await vscode.window.showInputBox({ prompt: localized['conditional-format-condition-value-2-prompt'], value: String(rule.values[1] ?? ''), validateInput: validateValue });
              if (second === undefined) return;
              values.push(second);
            }
            document.editConditionalFormatRule(messageSheet, selectedRule.priority, { kind: 'condition', operator: operator.value, values });
          } else if (selectedAction.value === 'condition' && rule.type === 'expression') {
            const formula = await vscode.window.showInputBox({
              prompt: localized['conditional-format-condition-formula-prompt'],
              value: '=' + rule.formula,
              validateInput: value => {
                const expression = String(value).trim().replace(/^=/, '');
                if (!expression || expression.length > 256 || /[!:\[\]]/.test(expression)) return localized['error-conditional-format-formula'];
                const sample = evaluateFormula(expression, { getCell: () => 1 });
                return sample === '#NAME?' || sample === '#VALUE!' ? localized['error-conditional-format-formula'] : undefined;
              }
            });
            if (formula === undefined) return;
            document.editConditionalFormatRule(messageSheet, selectedRule.priority, { kind: 'condition', formula });
          } else if (selectedAction.value === 'colors' && ['dataBar', 'colorScale'].includes(rule.type)) {
            const existingColors = rule.type === 'dataBar' ? [rule.dataBar.color] : rule.colorScale.colors, colors = [];
            for (const [index, existingColor] of existingColors.entries()) {
              const color = await vscode.window.showInputBox({
                prompt: localized['conditional-format-color-stop-prompt'].replace('{index}', String(index + 1)),
                value: existingColor,
                validateInput: value => /^#[0-9a-f]{6}$/i.test(String(value).trim()) ? undefined : localized['error-conditional-format-color-value']
              });
              if (color === undefined) return;
              colors.push(color);
            }
            document.editConditionalFormatRule(messageSheet, selectedRule.priority, { kind: 'colors', colors });
          }
          postModel();
        } else if (message.type === 'clear-conditional-formats' && currentSheetMessage && document.kind === 'xlsx') {
          const range = message.range;
          const localized = messagesFor(this.language);
          if (!range || ![range.top, range.bottom, range.left, range.right].every(Number.isInteger) || range.top < 0 || range.left < 0 || range.bottom < range.top || range.right < range.left || range.bottom >= XLSX_LIMITS.rows || range.right >= XLSX_LIMITS.columns) throw new RangeError(localized['error-conditional-format-range']);
          document.clearConditionalFormats(messageSheet, range); postModel();
        } else if ((message.type === 'edit-hyperlink' || message.type === 'remove-hyperlink') && currentSheetMessage && document.kind === 'xlsx') {
          const row = message.row, column = message.column, localized = messagesFor(this.language);
          if (!Number.isInteger(row) || row < 0 || row >= XLSX_LIMITS.rows || !Number.isInteger(column) || column < 0 || column >= XLSX_LIMITS.columns) throw new RangeError(localized['hyperlink-invalid']);
          const sheet = document.state.sheets[messageSheet], existing = sheet.hyperlinks?.find(link => row >= link.top && row <= link.bottom && column >= link.left && column <= link.right) || null;
          if (message.type === 'remove-hyperlink') {
            document.applyHyperlink(messageSheet, row, column, null); postModel(); return;
          }
          const address = await vscode.window.showInputBox({
            prompt: localized['hyperlink-target-prompt'],
            placeHolder: localized['hyperlink-target-placeholder'],
            value: existing ? (existing.target || '#' + existing.location) : ''
          });
          if (address === undefined) return;
          const destination = address.trim();
          if (!destination) throw new RangeError(localized['hyperlink-invalid']);
          let link;
          if (destination.startsWith('#') || !/^[A-Za-z][A-Za-z0-9+.-]*:/.test(destination)) {
            const location = destination.replace(/^#/, '').trim();
            if (!resolveInternalHyperlink(document, messageSheet, location)) throw new RangeError(localized['hyperlink-location-invalid']);
            link = { ...(existing?.tooltip ? { tooltip: existing.tooltip } : {}), location };
          } else {
            if (!externalHyperlinkUri(destination)) throw new RangeError(localized['hyperlink-invalid']);
            link = { ...(existing?.tooltip ? { tooltip: existing.tooltip } : {}), target: destination };
          }
          const cell = document.getCell(row, column, messageSheet), currentText = existing?.display || (cell ? formatCellValue(cell, this.language) : ''), defaultText = currentText || destination;
          const display = await vscode.window.showInputBox({
            prompt: localized['hyperlink-display-prompt'],
            placeHolder: defaultText,
            value: defaultText
          });
          if (display === undefined) return;
          link.display = display || null;
          document.applyHyperlink(messageSheet, row, column, link); postModel();
        } else if (message.type === 'follow-hyperlink' && currentSheetMessage && document.kind === 'xlsx') {
          if (!Number.isInteger(message.row) || message.row < 0 || message.row >= XLSX_LIMITS.rows || !Number.isInteger(message.column) || message.column < 0 || message.column >= XLSX_LIMITS.columns) throw new RangeError(messagesFor(this.language)['hyperlink-invalid']);
          const link = document.state.sheets[messageSheet]?.hyperlinks?.find(item => message.row >= item.top && message.row <= item.bottom && message.column >= item.left && message.column <= item.right);
          if (!link) throw new RangeError(messagesFor(this.language)['hyperlink-invalid']);
          if (link.target) {
            const destination = externalHyperlinkUri(link.target);
            if (!destination) throw new RangeError(messagesFor(this.language)['hyperlink-invalid']);
            if (!await vscode.env.openExternal(destination)) throw new Error(messagesFor(this.language)['hyperlink-open-failed']);
          } else {
            const target = resolveInternalHyperlink(document, messageSheet, link.location);
            if (!target) throw new RangeError(messagesFor(this.language)['hyperlink-location-invalid']);
            selectedSheet = target.sheetIndex;
            postModel();
            panel.webview.postMessage({ type: 'hyperlink-navigation', ...target });
          }
        } else if (message.type === 'search-navigate') {
          if (document.kind !== 'xlsx' || !['sheet', 'workbook'].includes(message.scope) || !Number.isSafeInteger(message.requestId) || message.requestId < 1 || !isValidSheetRequest(document, message.sheetIndex) || !Number.isInteger(message.row) || message.row < 0 || message.row >= XLSX_LIMITS.rows || !Number.isInteger(message.column) || message.column < 0 || message.column >= XLSX_LIMITS.columns) throw new RangeError('검색 결과 이동 요청이 유효하지 않습니다.');
          if (selectedSheet !== message.sheetIndex) { selectedSheet = message.sheetIndex; postModel(); }
          panel.webview.postMessage({ type: 'search-navigation', requestId: message.requestId, sheetIndex: message.sheetIndex, row: message.row, column: message.column });
        } else if (message.type === 'defined-name-navigate') {
          if (document.kind !== 'xlsx' || !Number.isSafeInteger(message.requestId) || message.requestId < 1 || message.sheetIndex !== selectedSheet || !isValidSheetRequest(document, message.sheetIndex) || typeof message.name !== 'string' || message.name.length > 255) throw new RangeError('정의된 이름 이동 요청이 유효하지 않습니다.');
          const key = message.name.toLocaleLowerCase(), definition = document.state.definedNames.find(item => item.localSheetId === message.sheetIndex && item.name.toLocaleLowerCase() === key) || document.state.definedNames.find(item => item.localSheetId === null && item.name.toLocaleLowerCase() === key);
          if (!definition || !definition.source || definition.printArea || definition.printTitles || definition.name.toLocaleLowerCase().startsWith('_xlnm.')) throw new RangeError('이동할 정의된 셀 범위를 찾을 수 없습니다.');
          const targetSheet = document.state.sheets.findIndex(sheet => sheet.name.toLocaleLowerCase() === definition.source.sheetName.toLocaleLowerCase());
          if (targetSheet < 0 || !document.state.sheets[targetSheet].visible) throw new RangeError('정의된 이름의 대상 표시 시트를 찾을 수 없습니다.');
          selectedSheet = targetSheet;
          postModel();
          panel.webview.postMessage({ type: 'defined-name-navigation', requestId: message.requestId, sheetIndex: targetSheet, range: { startRow: definition.source.top, endRow: definition.source.bottom, startColumn: definition.source.left, endColumn: definition.source.right } });
        } else if (message.type === 'cell-reference-navigate') {
          const range = message.range;
          if (document.kind !== 'xlsx' || !Number.isSafeInteger(message.requestId) || message.requestId < 1 || message.sheetIndex !== selectedSheet || !isValidSheetRequest(document, message.sheetIndex) || !range || ![range.startRow, range.endRow, range.startColumn, range.endColumn].every(Number.isInteger) || Math.min(range.startRow, range.endRow) < 0 || Math.min(range.startColumn, range.endColumn) < 0 || Math.max(range.startRow, range.endRow) >= XLSX_LIMITS.rows || Math.max(range.startColumn, range.endColumn) >= XLSX_LIMITS.columns) throw new RangeError('시트 셀 주소 이동 요청이 유효하지 않습니다.');
          if (!isValidSheetRequest(document, message.targetSheetIndex) || !document.state.sheets[message.targetSheetIndex].visible) throw new RangeError('시트 셀 주소의 대상 표시 시트가 유효하지 않습니다.');
          selectedSheet = message.targetSheetIndex;
          postModel();
          panel.webview.postMessage({ type: 'cell-reference-navigation', requestId: message.requestId, sheetIndex: selectedSheet, range });
        } else if (message.type === 'refresh-dynamic-filters' && currentSheetMessage && document.kind === 'xlsx') { postModel(); return; }
        if (message.type === 'edit' && currentSheetMessage && Number.isInteger(message.row) && Number.isInteger(message.column)) {
          document.applyCellEdit(message.row, message.column, message.value, messageSheet, new Set(), cellOptions);
        } else if (message.type === 'insert-date-time' && currentSheetMessage && Number.isInteger(message.row) && Number.isInteger(message.column) && ['date', 'time'].includes(message.mode)) {
          const edit = buildDateTimeShortcutEdit(document, message.row, message.column, message.mode);
          document.applyCellEdits([edit], messageSheet, translateDiagnostic(message.mode === 'date' ? '현재 날짜 입력' : '현재 시간 입력', this.language));
        } else if (message.type === 'auto-sum' && currentSheetMessage && Number.isInteger(message.row) && message.row >= 0 && Number.isInteger(message.column) && message.column >= 0 && document.kind === 'xlsx') {
          const formula = buildAutoSumFormula(document, message.row, message.column, messageSheet);
          if (formula) document.applyCellEdit(message.row, message.column, formula, messageSheet, new Set(), { ...cellOptions, skipExpansion: true });
          else panel.webview.postMessage({ type: 'error', message: message('auto-sum-unavailable', this.language) });
        } else if (message.type === 'fill-range' && currentSheetMessage && Number.isInteger(message.row) && Number.isInteger(message.column) && typeof message.value === 'string') {
          const matrix = [[message.value]], formulas = buildClipboardFormulaMatrix(null, document.kind, message.row, message.column, matrix);
          const edits = excludeHiddenRowEdits(buildPasteEdits(matrix, message.row, message.column, message.range, undefined, formulas), hiddenRowsForSheet(document, messageSheet));
          if (document.kind === 'xlsx') for (const edit of edits) if (!Object.hasOwn(edit, 'formula')) edit.formula = null;
          document.applyCellEdits(edits, messageSheet, translateDiagnostic('선택 범위 값 채우기', this.language), new Set(), true, cellOptions);
        } else if (message.type === 'clear' && currentSheetMessage && Number.isInteger(message.row) && Number.isInteger(message.column)) {
          document.applyCellEdits(buildClearEdits(document, message.row, message.column, message.range, messageSheet, hiddenRowsForSheet(document, messageSheet)), messageSheet, translateDiagnostic('선택 범위 값 지우기', this.language));
        } else if (message.type === 'normal-font-width' && currentSheetMessage && document.kind === 'xlsx') {
          if (!Number.isInteger(message.maximumDigitWidth) || message.maximumDigitWidth < 1 || message.maximumDigitWidth > 600) throw new RangeError('Normal 글꼴의 숫자 폭 측정값이 유효하지 않습니다.');
          document.maximumDigitWidth = message.maximumDigitWidth;
        } else if (message.type === 'column-width' && currentSheetMessage && Number.isInteger(message.column) && message.column >= 0 && message.column < XLSX_LIMITS.columns && Number.isFinite(message.width)) {
          document.applyColumnWidth(message.column, message.width, messageSheet);
        } else if (message.type === 'auto-fit-column' && currentSheetMessage && Number.isInteger(message.requestId) && message.requestId > 0 && Number.isInteger(message.column) && message.column >= 0 && message.column < (document.kind === 'csv' ? CSV_LIMIT_COLUMNS : XLSX_LIMITS.columns)) {
          const width = estimateDocumentColumnWidth(document, message.column, messageSheet, this.language);
          panel.webview.postMessage({ type: 'auto-fit-width', requestId: message.requestId, column: message.column, width, sheetIndex: messageSheet });
        } else if (message.type === 'number-format-custom' && currentSheetMessage && document.kind === 'xlsx') {
          if (!isSupportedCustomNumberFormatCode(message.formatCode)) throw new Error('지원하는 사용자 지정 숫자 형식은 0, #, 쉼표, 소수점, 통화 기호와 % 조합입니다. 예: #,##0.00');
          const numberFormat = { numberFormatId: nextCustomNumberFormatId(document), numberFormatCode: message.formatCode };
          const edits = buildNumberFormatEdits(document, message.range, numberFormat, messageSheet);
          if (edits.length) document.applyCellEdits(edits, messageSheet, translateDiagnostic('사용자 지정 숫자 표시 형식 변경', this.language));
        } else if (message.type === 'number-format' && currentSheetMessage && document.kind === 'xlsx' && ['number', 'currency', 'percent', 'date', 'time'].includes(message.format)) {
          const numberFormat = message.format === 'number' ? 4 : message.format === 'percent' ? 9 : message.format === 'date' ? 14 : message.format === 'time' ? 20 : (() => {
            const currency = parseClipboardCurrency(this.language === 'ko' ? '₩1.00' : '$1.00', this.language, document.state);
            return currency && { numberFormatId: currency.numFmtId, numberFormatCode: currency.formatCode };
          })();
          if (!numberFormat) throw new Error('통화 표시 형식을 만들 수 없습니다.');
          const edits = buildNumberFormatEdits(document, message.range, numberFormat, messageSheet);
          if (edits.length) document.applyCellEdits(edits, messageSheet, translateDiagnostic('숫자 표시 형식 변경', this.language));
        } else if (message.type === 'font-format' && currentSheetMessage && document.kind === 'xlsx' && ['bold', 'italic', 'underline', 'strike'].includes(message.property)) {
          const edits = buildFontFormatEdits(document, message.range, message.property, messageSheet);
          if (edits.length) document.applyCellEdits(edits, messageSheet, translateDiagnostic('글꼴 서식 변경', this.language));
        } else if (message.type === 'font-family' && currentSheetMessage && document.kind === 'xlsx') {
          const edits = buildFontFormatEdits(document, message.range, 'fontFamily', messageSheet, message.family);
          if (edits.length) document.applyCellEdits(edits, messageSheet, translateDiagnostic('글꼴 계열 변경', this.language));
        } else if (message.type === 'font-size' && currentSheetMessage && document.kind === 'xlsx') {
          const edits = buildFontFormatEdits(document, message.range, 'fontSize', messageSheet, message.size);
          if (edits.length) document.applyCellEdits(edits, messageSheet, translateDiagnostic('글꼴 크기 변경', this.language));
        } else if (message.type === 'font-color' && currentSheetMessage && document.kind === 'xlsx') {
          const edits = buildFontColorEdits(document, message.range, message.color, messageSheet);
          if (edits.length) document.applyCellEdits(edits, messageSheet, translateDiagnostic('글꼴 색 변경', this.language));
        } else if (message.type === 'alignment-format' && currentSheetMessage && document.kind === 'xlsx') {
          const edits = buildAlignmentFormatEdits(document, message.range, message.axis, message.value, messageSheet);
          if (edits.length) document.applyCellEdits(edits, messageSheet, translateDiagnostic('셀 맞춤 변경', this.language));
        } else if (message.type === 'border-format' && currentSheetMessage && document.kind === 'xlsx') {
          const edits = buildBorderFormatEdits(document, message.range, message.preset, messageSheet, message.line, message.color);
          if (edits.length) document.applyCellEdits(edits, messageSheet, translateDiagnostic('셀 테두리 변경', this.language));
        } else if (message.type === 'fill-format' && currentSheetMessage && document.kind === 'xlsx') {
          const edits = buildFillFormatEdits(document, message.range, message.color, messageSheet);
          if (edits.length) document.applyCellEdits(edits, messageSheet, translateDiagnostic('셀 채우기 색 변경', this.language));
        } else if (message.type === 'calculated-column' && currentSheetMessage && document.kind === 'xlsx') {
          const table = document.state.sheets[messageSheet].tables.find(item => item.name === message.tableName);
          if (!table || !Number.isSafeInteger(message.requestId) || message.requestId < 1 || !Number.isInteger(message.row) || !Number.isInteger(message.column) || message.row < table.range.top || message.row > table.range.bottom || message.column < table.range.left || message.column > table.range.right) throw new RangeError('계산 열 요청이 유효하지 않습니다.');
          const result = document.applyCalculatedColumn(messageSheet, message.row, message.column, message.mode, message.remove);
          panel.webview.postMessage({ type: 'calculated-column-applied', requestId: message.requestId, sheetIndex: messageSheet, row: message.row, column: message.column, ...result });
        } else if (message.type === 'table-totals' && currentSheetMessage && document.kind === 'xlsx') {
          try {
            if (typeof message.tableName !== 'string' || !Number.isSafeInteger(message.requestId) || message.requestId < 1 || !Number.isInteger(message.row) || !Number.isInteger(message.column)) throw new RangeError('표 합계 요청이 유효하지 않습니다.');
            const currentTable = document.state.sheets[messageSheet].tables.find(item => item.name === message.tableName);
            if (!currentTable || message.row < currentTable.range.top || message.row > currentTable.range.bottom || message.column < currentTable.range.left || message.column > currentTable.range.right) throw new RangeError('표 합계 요청이 유효하지 않습니다.');
            document.applyTableTotals(messageSheet, message.tableName, { enabled: message.enabled, column: message.column, function: message.function, label: messagesFor(this.language)['table-totals-label'] });
            const table = document.state.sheets[messageSheet].tables.find(item => item.name === message.tableName);
            panel.webview.postMessage({ type: 'table-totals-applied', requestId: message.requestId, sheetIndex: messageSheet, tableName: table.name, row: table.totalsRowCount ? table.range.bottom : Math.min(message.row, table.range.bottom), column: message.column });
          } catch (error) { panel.webview.postMessage({ type: 'table-totals-error', requestId: message.requestId, sheetIndex: messageSheet, tableName: message.tableName, message: translateDiagnostic(error.message, this.language) }); }
        } else if (message.type === 'table-tab' && currentSheetMessage && document.kind === 'xlsx') {
          if (!Number.isSafeInteger(message.requestId) || message.requestId < 1 || typeof message.tableName !== 'string') throw new RangeError('표 행 추가 요청이 유효하지 않습니다.');
          const result = document.appendTableRow(messageSheet, message.tableName, message.row, message.column, message.value, new Set(), cellOptions);
          panel.webview.postMessage({ type: 'table-tab-applied', requestId: message.requestId, sheetIndex: messageSheet, tableName: message.tableName, ...result });
        } else if (message.type === 'create-table' && currentSheetMessage && document.kind === 'xlsx') {
          try {
            if (!Number.isSafeInteger(message.requestId) || message.requestId < 1) throw new RangeError('표 생성 요청이 유효하지 않습니다.');
            const result=document.applyCreateTable(messageSheet,message.range,message.name,this.language);
            panel.webview.postMessage({type:'table-created',requestId:message.requestId,sheetIndex:messageSheet,...result});
          } catch(error) { panel.webview.postMessage({type:'table-create-error',requestId:message.requestId,sheetIndex:messageSheet,message:translateDiagnostic(error.message,this.language)}); }
        } else if (message.type === 'table-style' && currentSheetMessage && document.kind === 'xlsx') {
          try {
            if (!Number.isSafeInteger(message.requestId) || message.requestId < 1 || typeof message.tableName !== 'string') throw new RangeError('표 스타일 요청이 유효하지 않습니다.');
            document.applyTableStyle(messageSheet, message.tableName, message.styleInfo);
            postModel();
            panel.webview.postMessage({ type: 'table-style-applied', requestId: message.requestId, sheetIndex: messageSheet, tableName: message.tableName });
          } catch (error) { panel.webview.postMessage({ type: 'table-style-error', requestId: message.requestId, sheetIndex: messageSheet, tableName: message.tableName, message: translateDiagnostic(error.message, this.language) }); }
        } else if (message.type === 'resize-table' && currentSheetMessage && document.kind === 'xlsx' && typeof message.tableName === 'string' && message.range && ['top', 'bottom', 'left', 'right'].every(key => Number.isInteger(message.range[key]))) {
          document.applyTableResize(messageSheet, message.tableName, message.range);
        } else if (message.type === 'insert-chart' && currentSheetMessage && document.kind === 'xlsx' && ['bar', 'line', 'area', 'scatter', 'pie', 'doughnut'].includes(message.chartType) && message.range && ['top', 'bottom', 'left', 'right'].every(key => Number.isInteger(message.range[key]))) {
          document.applyAddChart(messageSheet, message.range, message.chartType);
        } else if (message.type === 'chart-source' && currentSheetMessage && document.kind === 'xlsx' && Number.isInteger(message.chartIndex) && message.chartIndex >= 0 && message.range && ['top', 'bottom', 'left', 'right'].every(key => Number.isInteger(message.range[key]))) {
          document.applyChartSource(messageSheet, message.chartIndex, message.range);
        } else if (message.type === 'delete-chart' && currentSheetMessage && document.kind === 'xlsx' && Number.isInteger(message.chartIndex) && message.chartIndex >= 0) {
          document.applyDeleteChart(messageSheet, message.chartIndex);
        } else if (message.type === 'edit-chart' && currentSheetMessage && document.kind === 'xlsx' && Number.isInteger(message.chartIndex) && message.chartIndex >= 0 && typeof message.title === 'string' && message.title.length <= 512 && Array.isArray(message.seriesNames) && message.seriesNames.length <= 32 && message.seriesNames.every(name => typeof name === 'string' && name.length <= 256)) {
          document.applyChartMetadata(messageSheet, message.chartIndex, { title: message.title, seriesNames: message.seriesNames });
        } else if (message.type === 'chart-placement' && currentSheetMessage && document.kind === 'xlsx' && Number.isInteger(message.chartIndex) && ['twoCell', 'oneCell', 'absolute'].includes(message.placement)) {
          await document.waitForAutoRowFits();
          document.applyChartPlacement(messageSheet, message.chartIndex, message.placement);
        } else if (message.type === 'structure' && currentSheetMessage && ['row', 'column'].includes(message.axis) && ['insert', 'delete'].includes(message.action) && Number.isInteger(message.index)) {
          document.applyStructureChange({ axis: message.axis, action: message.action, index: message.index, count: message.count, sheetIndex: messageSheet });
        } else if (['wrap-text', 'row-auto-fit'].includes(message.type) && currentSheetMessage && document.kind === 'xlsx' && document.writable) {
          const edits = message.type === 'wrap-text' ? buildWrapTextEdits(document, message.range, messageSheet) : [];
          const bounds = message.type === 'wrap-text' ? selectionBounds(message.range, '줄 바꿈', Number.MAX_SAFE_INTEGER) : { top: message.startRow, bottom: message.endRow };
          const fit = buildRowFitRequest(document, bounds.top, bounds.bottom, messageSheet, edits, message.type === 'wrap-text', this.language);
          const label = message.type === 'wrap-text' ? '텍스트 줄 바꿈 변경' : '행 높이 자동 맞춤';
          if (!fit.rows.length) document.applyTextLayout(messageSheet, edits, [], label);
          else {
            const requestId = ++fitSequence;
            pendingFit = { requestId, revision: document.renderRevision, sheetIndex: messageSheet, rows: fit.rows.map(item => item.row), edits, label };
            panel.webview.postMessage({ type: 'row-fit-request', requestId, sheetIndex: messageSheet, ...fit });
          }
        } else if (message.type === 'row-fit-result') {
          const automatic = automaticFits.get(message.requestId);
          if (automatic) { document.completeAutoRowFit(automatic, message); return; }
          if (!pendingFit || message.requestId !== pendingFit.requestId) return;
          const pending = pendingFit; pendingFit = null;
          validateRowFitResult(pending, message, document.renderRevision, selectedSheet);
          document.applyTextLayout(messageSheet, pending.edits, message.heights, pending.label);
        } else if (message.type === 'page-settings' && currentSheetMessage && document.kind === 'xlsx') {
          document.applyPrintSettings(messageSheet, message.settings);
        } else if (message.type === 'page-break' && currentSheetMessage && document.kind === 'xlsx') {
          document.applyPageBreaks(messageSheet, { axis: message.axis, action: message.action, index: message.index });
        } else if (message.type === 'print-area' && currentSheetMessage && document.kind === 'xlsx' && ['set', 'clear'].includes(message.action)) {
          const range = message.action === 'set' ? selectionBounds(message.range, '인쇄 영역', Number.MAX_SAFE_INTEGER) : null;
          document.applyPrintArea(messageSheet, range);
        } else if (message.type === 'print-titles' && currentSheetMessage && document.kind === 'xlsx' && ['set', 'clear'].includes(message.action)) {
          const titles = message.action === 'clear' ? { rows: null, columns: null } : parsePrintTitleInputs(message.rows, message.columns);
          document.applyPrintTitles(messageSheet, titles);
        } else if (message.type === 'header-footer' && currentSheetMessage && document.kind === 'xlsx' && ['set', 'clear'].includes(message.action)) {
          document.applyHeaderFooter(messageSheet, message.settings, message.action === 'clear');
        } else if (message.type === 'export-print' && currentSheetMessage && ['csv', 'xlsx'].includes(document.kind)) {
          await document.waitForAutoRowFits();
          if (message.includeCharts !== undefined && typeof message.includeCharts !== 'boolean') throw new RangeError('차트 포함 요청이 유효하지 않습니다.');
          const html = buildPrintDocument(document, messageSheet, message.range ?? null, this.language, undefined, message.includeCharts !== false);
          const destination = await vscode.window.showSaveDialog({ title: messagesFor(this.language)['print-export-title'], defaultUri: vscode.Uri.file(document.uri.fsPath.replace(/\.[^.]+$/, '') + '-print.html'), filters: { HTML: ['html'] } });
          if (destination) {
            if (destination.scheme !== 'file' || path.extname(destination.fsPath).toLowerCase() !== '.html') throw new Error('인쇄 미리보기는 로컬 HTML 파일로 저장하세요.');
            const temporary = destination.fsPath + '.jinsheet-' + crypto.randomBytes(12).toString('hex') + '.tmp';
            try {
              await fs.writeFile(temporary, html, { encoding: 'utf8', flag: 'wx', mode: 0o600 });
              await fs.rename(temporary, destination.fsPath);
            } catch (error) {
              await fs.rm(temporary, { force: true }).catch(() => {}); throw error;
            }
            if (!await vscode.env.openExternal(destination)) await vscode.window.showInformationMessage(messagesFor(this.language)['print-open-manually']);
          }
        } else if (message.type === 'row-height' && currentSheetMessage && document.kind === 'xlsx' && Number.isInteger(message.startRow) && Number.isInteger(message.endRow) && (message.height === null || Number.isFinite(message.height))) {
          document.applyRowHeights(messageSheet, message.startRow, message.endRow, message.height);
        } else if (message.type === 'hidden-rows' && currentSheetMessage && ['xlsx', 'csv'].includes(document.kind) && ['hide', 'unhide-all'].includes(message.action)) {
          if (document.kind === 'csv') updateCsvHiddenAxis(document, messageSheet, 'row', message.action, message.startRow, message.endRow);
          else {
            const rows = message.action === 'hide' && Number.isInteger(message.startRow) && Number.isInteger(message.endRow) && message.startRow >= 0 && message.endRow >= 0
              ? Array.from({ length: Math.abs(message.endRow - message.startRow) + 1 }, (_, index) => Math.min(message.startRow, message.endRow) + index)
              : message.action === 'unhide-all' ? [...(document.state.sheets[messageSheet]?.hiddenRows || [])] : null;
            if (rows && rows.length) document.applyHiddenRows(messageSheet, rows, message.action === 'hide');
          }
        } else if (message.type === 'hidden-columns' && currentSheetMessage && ['xlsx', 'csv'].includes(document.kind) && ['hide', 'unhide-all'].includes(message.action)) {
          if (document.kind === 'csv') updateCsvHiddenAxis(document, messageSheet, 'column', message.action, message.startColumn, message.endColumn);
          else {
            const columns = message.action === 'hide' && Number.isInteger(message.startColumn) && Number.isInteger(message.endColumn) && message.startColumn >= 0 && message.endColumn >= 0
              ? Array.from({ length: Math.abs(message.endColumn - message.startColumn) + 1 }, (_, index) => Math.min(message.startColumn, message.endColumn) + index)
              : message.action === 'unhide-all' ? [...(document.state.sheets[messageSheet]?.hiddenColumns || [])] : null;
            if (columns && columns.length) document.applyHiddenColumns(messageSheet, columns, message.action === 'hide');
          }
        } else if (message.type === 'outline-groups' && currentSheetMessage && ['xlsx', 'csv'].includes(document.kind) && ['row', 'column'].includes(message.axis) && ['group', 'ungroup', 'collapse', 'expand', 'collapse-summary', 'expand-summary', 'show-level'].includes(message.action) && (['collapse-summary', 'expand-summary'].includes(message.action) ? Number.isInteger(message.summary) : message.action === 'show-level' ? Number.isInteger(message.level) : Number.isInteger(message.start) && Number.isInteger(message.end))) {
          if (document.kind === 'csv') updateCsvOutlineGroups(document, messageSheet, message.axis, message.action, message.start, message.end, message.summary, message.level);
          else document.applyOutlineGroups(messageSheet, message.axis, message.action, message.start, message.end, message.summary, message.level);
        } else if (message.type === 'freeze-pane' && currentSheetMessage && document.kind === 'xlsx' && Number.isInteger(message.rows) && Number.isInteger(message.columns)) {
          document.applyFreezePane(messageSheet, { rows: message.rows, columns: message.columns });
        } else if (message.type === 'merge-cells' && currentSheetMessage && document.kind === 'xlsx' && ['merge', 'unmerge'].includes(message.action) && message.range && ['startRow', 'endRow', 'startColumn', 'endColumn'].every(key => Number.isInteger(message.range[key]))) {
          const range = { top: Math.min(message.range.startRow, message.range.endRow), bottom: Math.max(message.range.startRow, message.range.endRow), left: Math.min(message.range.startColumn, message.range.endColumn), right: Math.max(message.range.startColumn, message.range.endColumn) };
          try { document.applyMergedCells(messageSheet, range, message.action); }
          catch (error) {
            if (!error.needsConfirmation) throw error;
            const confirmLabel = message('merge-cells-confirm', this.language);
            if (await vscode.window.showWarningMessage(message('merge-cells-data-warning', this.language), { modal: true }, confirmLabel) === confirmLabel) document.applyMergedCells(messageSheet, range, message.action, true);
          }
        } else if (message.type === 'navigate' && currentSheetMessage && ['up', 'down', 'left', 'right'].includes(message.direction)) {
          panel.webview.postMessage(navigateCell(document, message));
        } else if (message.type === 'range') {
          if (!currentSheetMessage) return;
          panel.webview.postMessage(getRange(document, message, this.language));
        } else if (message.type === 'text-to-columns-preview' && currentSheetMessage && ['csv', 'xlsx'].includes(document.kind) && Number.isSafeInteger(message.requestId) && message.requestId > 0 && typeof message.delimiter === 'string' && typeof message.mergeConsecutive === 'boolean') {
          const plan = planTextToColumns(document, message.range, message.delimiter, message.mergeConsecutive, messageSheet);
          panel.webview.postMessage({ type: 'text-to-columns-preview', requestId: message.requestId, sheetIndex: messageSheet, preview: plan.preview, conflictCount: plan.conflicts.length });
        } else if (message.type === 'remove-duplicates-preview' && currentSheetMessage && ['csv', 'xlsx'].includes(document.kind) && Number.isSafeInteger(message.requestId) && message.requestId > 0) {
          const plan = planRemoveDuplicates(document, message.range, message.columns, message.hasHeader, messageSheet);
          panel.webview.postMessage({ type: 'remove-duplicates-preview', requestId: message.requestId, sheetIndex: messageSheet, removedCount: plan.removedCount });
        } else if (message.type === 'selection-summary' && currentSheetMessage && Number.isInteger(message.requestId) && message.requestId > 0) {
          panel.webview.postMessage(selectionSummary(document, message));
        } else if (message.type === 'current-region' && currentSheetMessage && Number.isInteger(message.row) && message.row >= 0 && Number.isInteger(message.column) && message.column >= 0 && Number.isInteger(message.requestId) && message.requestId > 0) {
          panel.webview.postMessage({ type: 'current-region', requestId: message.requestId, row: message.row, column: message.column, sheetIndex: messageSheet, range: currentRegion(document, message.row, message.column, messageSheet, Boolean(message.includeSheet)) });
        } else if (message.type === 'copy' && currentSheetMessage && Number.isInteger(message.row) && Number.isInteger(message.column)) {
          const range = message.range || { startRow: message.row, endRow: message.row, startColumn: message.column, endColumn: message.column };
          const payload = buildCopyPayload(document, range, messageSheet, this.language, message.visibleRows, message.hiddenRows);
          payload.sourceUri = document.uri.toString();
          this.internalClipboard = payload;
          if (payload.text !== null) {
            try { await this.writeClipboardText(payload.text); }
            catch (error) { payload.failed = true; if (this.internalClipboard === payload) this.internalClipboard = null; throw error; }
          }
        } else if (message.type === 'cut' && currentSheetMessage && Number.isInteger(message.row) && Number.isInteger(message.column)) {
          this.internalClipboard = null;
          if (!document.writable) throw new Error(document.reason || '읽기 전용 파일입니다.');
          const range = message.range || { startRow: message.row, endRow: message.row, startColumn: message.column, endColumn: message.column };
          const fullColumn = document.kind === 'xlsx' && Math.abs(range.endRow - range.startRow) + 1 === XLSX_LIMITS.rows && range.startColumn === range.endColumn;
          const sparseCopy = fullColumn ? buildCopyPayload(document, range, messageSheet, this.language) : null;
          const hiddenRows = hiddenRowsForSheet(document, messageSheet);
          const bounds = selectionBounds(range, '복사', sparseCopy ? XLSX_LIMITS.rows : document.kind === 'xlsx' ? 20000 : 10000);
          const hiddenSet = new Set(hiddenRows || []);
          const visibleCutRows = hiddenRows && !sparseCopy ? Array.from({ length: bounds.bottom - bounds.top + 1 }, (_, offset) => bounds.top + offset).filter(row => !hiddenSet.has(row)) : undefined;
          const text = sparseCopy ? null : serializeClipboard(document, range, messageSheet, this.language, visibleCutRows);
          const cells = sparseCopy ? sparseCopy.cells.map(item => ({ row: item.rowOffset, column: 0 })) : null;
          const validateCell = (row, column) => {
            const cell = document.getCell(sparseCopy ? bounds.top + row : row, sparseCopy ? bounds.left + column : column, messageSheet);
            if ((cell?.kind === 'error' && !isFormulaCell(cell)) || (isFormulaCell(cell) && (!cell.formula || cell.formulaType))) throw new Error('공유·배열 수식 또는 일반 오류 셀이 포함된 범위는 이동할 수 없습니다.');
          };
          if (cells) for (const cell of cells) validateCell(cell.row, cell.column);
          else for (const row of visibleCutRows ?? Array.from({ length: bounds.bottom - bounds.top + 1 }, (_, offset) => bounds.top + offset)) for (let column = bounds.left; column <= bounds.right; column++) validateCell(row, column);
          if (text !== null) await this.writeClipboardText(text);
          if (selectedSheet !== messageSheet) return;
          panel.webview.postMessage({ type: 'cut-ready', range, sheetIndex: messageSheet });
        } else if (message.type === 'move-paste' && currentSheetMessage && Number.isInteger(message.row) && Number.isInteger(message.column)) {
          if (!document.writable) throw new Error(document.reason || '읽기 전용 파일입니다.');
          const fullColumn = document.kind === 'xlsx' && Math.abs(message.sourceRange?.endRow - message.sourceRange?.startRow) + 1 === XLSX_LIMITS.rows && message.sourceRange?.startColumn === message.sourceRange?.endColumn;
          const moveEdits = fullColumn ? buildSparseColumnMoveEdits(document, message.sourceRange, message.row, message.column, messageSheet) : buildMoveEdits(document, message.sourceRange, message.row, message.column, messageSheet, hiddenRowsForSheet(document, messageSheet));
          if (moveEdits.length) document.applyCellEdits(moveEdits, messageSheet, translateDiagnostic('셀 범위 이동', this.language), new Set(), true, cellOptions);
          panel.webview.postMessage({ type: 'cut-complete', sheetIndex: messageSheet });
        } else if (message.type === 'fill-selection' && currentSheetMessage && ['down', 'right'].includes(message.direction)) {
          const edits = buildFillEdits(document, message.range, message.direction, messageSheet, hiddenRowsForSheet(document, messageSheet));
          if (edits.length) document.applyCellEdits(edits, messageSheet, translateDiagnostic('선택 범위 채우기', this.language), new Set(), true, cellOptions);
        } else if (message.type === 'fill-handle' && currentSheetMessage) {
          const edits = buildFillRangeEdits(document, message.sourceRange, message.targetRange, messageSheet, hiddenRowsForSheet(document, messageSheet));
          if (edits.length) document.applyCellEdits(edits, messageSheet, translateDiagnostic('선택 범위 채우기', this.language), new Set(), true, cellOptions);
        } else if (message.type === 'sort-selection' && currentSheetMessage && (Array.isArray(message.sortKeys) ? message.sortKeys.length > 0 && message.sortKeys.length <= 64 && message.sortKeys.every(key => key && Number.isInteger(key.column) && ['ascending', 'descending'].includes(key.direction)) : ['ascending', 'descending'].includes(message.direction) && Number.isInteger(message.keyColumn))) {
          const edits = buildSortEdits(document, message.range, message.sortKeys || message.keyColumn, message.direction, Boolean(message.hasHeader), Boolean(message.expandSelection), messageSheet, hiddenRowsForSheet(document, messageSheet));
          if (edits.length) document.applyCellEdits(edits, messageSheet, translateDiagnostic('선택 범위 정렬', this.language), new Set(), true, cellOptions);
        } else if (message.type === 'remove-duplicates' && currentSheetMessage && ['csv', 'xlsx'].includes(document.kind) && Number.isSafeInteger(message.requestId) && message.requestId > 0) {
          const plan = planRemoveDuplicates(document, message.range, message.columns, message.hasHeader, messageSheet);
          if (plan.removedCount) {
            const labels = messagesFor(this.language), confirm = labels['remove-duplicates-confirm'];
            if (await vscode.window.showWarningMessage(labels['remove-duplicates-warning'].replace('{count}', String(plan.removedCount)), { modal: true }, confirm) !== confirm) {
              panel.webview.postMessage({ type: 'remove-duplicates-error', requestId: message.requestId, sheetIndex: messageSheet, message: labels['remove-duplicates-cancelled'] });
              return;
            }
            removeDuplicateRows(document, plan, messageSheet);
            postModel();
          }
          panel.webview.postMessage({ type: 'remove-duplicates-applied', requestId: message.requestId, sheetIndex: messageSheet, removedCount: plan.removedCount });
        } else if (message.type === 'text-to-columns' && currentSheetMessage && ['csv', 'xlsx'].includes(document.kind) && Number.isSafeInteger(message.requestId) && message.requestId > 0 && typeof message.delimiter === 'string' && typeof message.mergeConsecutive === 'boolean') {
          let edits;
          try { edits = buildTextToColumnsEdits(document, message.range, message.delimiter, message.mergeConsecutive, messageSheet); }
          catch (error) {
            if (!error.needsConfirmation) throw error;
            const labels = messagesFor(this.language), confirm = labels['text-to-columns-overwrite-confirm'];
            if (await vscode.window.showWarningMessage(labels['text-to-columns-overwrite-warning'], { modal: true }, confirm) !== confirm) {
              panel.webview.postMessage({ type: 'text-to-columns-error', requestId: message.requestId, sheetIndex: messageSheet, message: labels['text-to-columns-cancelled'] });
              return;
            }
            edits = buildTextToColumnsEdits(document, message.range, message.delimiter, message.mergeConsecutive, messageSheet, true);
          }
          document.applyCellEdits(edits, messageSheet, translateDiagnostic('텍스트 나누기', this.language), new Set(), true, { ...cellOptions, skipExpansion: true, skipHeaders: true });
          panel.webview.postMessage({ type: 'text-to-columns-applied', requestId: message.requestId, sheetIndex: messageSheet });
        } else if (message.type === 'filter-options' && currentSheetMessage && ['xlsx', 'csv'].includes(document.kind) && Number.isInteger(message.row) && Number.isInteger(message.column)) {
          const filter = autoFilterRange(document, messageSheet, message.row, message.column);
          panel.webview.postMessage({ type: 'filter-options', column: message.column, sheetIndex: messageSheet, filter, options: autoFilterOptions(document, messageSheet, message.column, filter) });
        } else if (message.type === 'auto-filter' && currentSheetMessage && ['xlsx', 'csv'].includes(document.kind) && Number.isInteger(message.row) && Number.isInteger(message.column)) {
          try {
            if (message.clear !== undefined && typeof message.clear !== 'boolean' || message.requestId !== undefined && (!Number.isSafeInteger(message.requestId) || message.requestId < 1)) throw new RangeError('자동 필터 요청이 유효하지 않습니다.');
            if (document.kind === 'csv' && ['iconFilter', 'colorFilter', 'top10', 'dateGroupText', 'dynamic'].some(key => Object.hasOwn(message, key))) throw new RangeError('CSV/TSV 자동 필터는 값 선택과 조건 비교만 지원합니다.');
            const baseFilter = autoFilterRange(document, messageSheet, message.row, message.column);
            if (Object.hasOwn(message, 'iconFilter') && ['colorFilter', 'custom', 'top10', 'dateGroupText', 'dynamic', 'values', 'blanks'].some(key => Object.hasOwn(message, key))) throw new RangeError('아이콘 필터는 같은 열의 다른 조건과 함께 지정할 수 없습니다.');
            if (Object.hasOwn(message, 'colorFilter') && ['custom', 'top10', 'dateGroupText', 'dynamic', 'values', 'blanks'].some(key => Object.hasOwn(message, key))) throw new RangeError('색상 필터는 같은 열의 다른 조건과 함께 지정할 수 없습니다.');
            if (Object.hasOwn(message, 'dynamic') && ['custom', 'top10', 'dateGroupText', 'values', 'blanks'].some(key => Object.hasOwn(message, key))) throw new RangeError('동적 필터는 같은 열의 다른 조건과 함께 지정할 수 없습니다.');
            if (Object.hasOwn(message, 'top10') && (message.custom !== undefined || message.values !== undefined || message.blanks !== undefined)) throw new RangeError('한 열의 순위 필터에 다른 종류의 조건을 함께 지정할 수 없습니다.');
            if (Object.hasOwn(message, 'dateGroupText') && (message.custom !== undefined || message.top10 !== undefined)) throw new RangeError('날짜 그룹은 같은 열의 조건·순위 필터와 함께 지정할 수 없습니다.');
            const filter = Object.hasOwn(message, 'iconFilter') && !message.clear ? updateAutoFilterIconColumn(document, messageSheet, message.column, message.iconFilter, baseFilter) : Object.hasOwn(message, 'colorFilter') && !message.clear ? updateAutoFilterColorColumn(document, messageSheet, message.column, message.colorFilter, baseFilter) : Object.hasOwn(message, 'dynamic') && !message.clear ? updateAutoFilterDynamicColumn(document, messageSheet, message.column, message.dynamic, baseFilter) : Object.hasOwn(message, 'dateGroupText') && !message.clear ? updateAutoFilterDateColumn(document, messageSheet, message.column, parseDateGroups(message.dateGroupText), message.values, message.blanks, baseFilter) : Object.hasOwn(message, 'top10') && !message.clear ? updateAutoFilterTopColumn(document, messageSheet, message.column, message.top10, baseFilter) : message.custom && !message.clear ? updateAutoFilterCustomColumn(document, messageSheet, message.column, message.custom, baseFilter) : updateAutoFilterColumn(document, messageSheet, message.column, message.values, message.blanks, Boolean(message.clear), baseFilter);
            if (document.kind === 'csv') {
              if (!document.writable) throw new Error(document.reason || '읽기 전용 파일입니다.');
              if (JSON.stringify(document.csvAutoFilter ?? null) !== JSON.stringify(filter ?? null)) {
                document.csvAutoFilter = filter;
                document.renderEmitter.fire();
              }
            } else document.applyAutoFilter(messageSheet, filter);
            panel.webview.postMessage({ type: 'filter-applied', sheetIndex: messageSheet, column: message.column, requestId: message.requestId });
          } catch (error) {
            panel.webview.postMessage({ type: 'filter-error', sheetIndex: messageSheet, column: message.column, requestId: message.requestId, message: translateDiagnostic(error.message, this.language) });
          }
        } else if (message.type === 'paste' && currentSheetMessage && Number.isInteger(message.row) && Number.isInteger(message.column)) {
          const targetSheet = messageSheet;
          const pendingCopy = this.internalClipboard;
          if (pendingCopy?.sparse && !isSameWorkbookClipboard(pendingCopy, document)) throw new Error('백만 행 전체 열 복사는 같은 XLSX 통합문서 안에서만 붙여넣을 수 있습니다.');
          const value = await this.readClipboardText(pendingCopy?.sparse ? () => null : undefined);
          if (selectedSheet !== targetSheet) return;
          const copied = pendingCopy && !pendingCopy.failed && (pendingCopy.text === value || pendingCopy.text === null) ? pendingCopy : null;
          if (copied?.sparse) {
            const edits = buildSparseColumnPasteEdits(document, copied, message.row, message.column, message.range, targetSheet);
            if (edits.length) document.applyCellEdits(edits, targetSheet, '셀 범위 붙여넣기', new Set(), true, cellOptions);
            return;
          }
          const matrix = copied ? copied.cells.map(line => line.map(cell => cell.value)) : parseClipboard(value);
          const sameWorkbook = isSameWorkbookClipboard(copied, document);
          const styleMatrix = sameWorkbook ? copied.cells.map(line => line.map(({ styleIndex, styleProperties, kind }) => ({ styleIndex, styleProperties, kind }))) : undefined;
          const numberFormatMatrix = !copied && document.kind === 'xlsx' && document.state.stylesPath
            ? matrix.map(line => line.map(value => parseClipboardDate(value, document.locale, document.state.date1904) ?? parseClipboardTime(value) ?? parseClipboardPercentage(value) ?? parseClipboardCurrency(value, this.language, document.state))) : undefined;
          if (numberFormatMatrix) for (let row = 0; row < matrix.length; row++) for (let column = 0; column < matrix[row].length; column++) {
            const numberFormat = numberFormatMatrix[row][column];
            if (numberFormat === null) continue;
            matrix[row][column] = String(numberFormat.serial);
          }
          const kindMatrix = buildClipboardKindMatrix(copied, document.kind, matrix);
          const formulaMatrix = buildClipboardFormulaMatrix(sameWorkbook ? copied : null, document.kind, message.row, message.column, matrix);
          const hiddenRows = hiddenRowsForSheet(document, targetSheet);
          const edits = excludeHiddenRowEdits(buildPasteEdits(matrix, message.row, message.column, message.range, styleMatrix, formulaMatrix, hiddenRows, kindMatrix, numberFormatMatrix), hiddenRows);
          if (document.kind === 'xlsx') for (const edit of edits) if (!Object.hasOwn(edit, 'formula')) edit.formula = null;
          if (edits.length) document.applyCellEdits(edits, targetSheet, '셀 범위 붙여넣기', new Set(), true, cellOptions);
        } else if (['replace-current', 'replace-all'].includes(message.type) && (currentSheetMessage || workbookMessage)) {
          const scope = message.scope === undefined ? 'sheet' : message.scope;
          replacementContext = { requestId: message.requestId, sheetIndex: messageSheet, query: message.query, mode: message.type, scope };
          if (!['sheet', 'workbook'].includes(scope) || scope === 'workbook' && document.kind !== 'xlsx' || !Number.isSafeInteger(message.requestId) || message.requestId < 1 || typeof message.query !== 'string' || message.query.length > 1000 || typeof message.replacement !== 'string' || message.replacement.length > MAX_CELL_CHARS || (message.row !== undefined && (!Number.isInteger(message.row) || message.row < 0)) || (message.column !== undefined && (!Number.isInteger(message.column) || message.column < 0)) || (message.row !== undefined) !== (message.column !== undefined) || (message.targetSheetIndex !== undefined && !isValidSheetRequest(document, message.targetSheetIndex))) throw new RangeError('찾기 및 바꾸기 요청이 유효하지 않습니다.');
          const targetSheet = message.type === 'replace-current' && Number.isInteger(message.targetSheetIndex) ? message.targetSheetIndex : messageSheet;
          const target = message.type === 'replace-current' && Number.isInteger(message.row) ? { row: message.row, column: message.column, sheetIndex: targetSheet } : null;
          const result = buildScopedReplaceEdits(document, message.query, message.replacement, targetSheet, target, message.type === 'replace-all', scope);
          replacementContext = { ...replacementContext, target: result.target ? { ...result.target, sheetIndex: result.target.sheetIndex ?? targetSheet } : null, replaced: result.replaced, skippedFormulas: result.skippedFormulas };
          Object.defineProperty(replacementContext, 'sendResult', { value: () => panel.webview.postMessage({ ...findMatches(document, message.query, messageSheet, scope), ...replacementContext, type: 'replace-result' }) });
          if (scope === 'workbook') {
            if (result.sheetEdits.length) document.applyCellEditsAcrossSheets(result.sheetEdits, translateDiagnostic('찾기 및 바꾸기', this.language));
          } else if (result.edits.length) document.applyCellEdits(result.edits, targetSheet, translateDiagnostic('찾기 및 바꾸기', this.language), new Set(), true, cellOptions);
          replacementContext.sendResult();
          replacementContext = null;
        } else if (message.type === 'search' && (currentSheetMessage || workbookMessage) && typeof message.query === 'string' && message.query.length <= 1000) {
          const scope = message.scope === undefined ? 'sheet' : message.scope;
          if (!['sheet', 'workbook'].includes(scope) || scope === 'workbook' && document.kind !== 'xlsx') throw new RangeError('검색 범위 요청이 유효하지 않습니다.');
          const after = message.after && Number.isInteger(message.after.row) && Number.isInteger(message.after.column) && (!Number.isInteger(message.after.sheetIndex) || message.after.sheetIndex >= 0) ? message.after : null;
          panel.webview.postMessage({ ...findMatches(document, message.query, messageSheet, scope, after), requestId: message.requestId });
        } else if (message.type === 'add-sheet' && currentSheetMessage && document.kind === 'xlsx') {
          selectedSheet = document.state.sheets.length;
          document.applyAddSheet();
        } else if (message.type === 'delete-sheet' && currentSheetMessage && document.kind === 'xlsx' && Number.isInteger(message.sheetIndex) && document.state.sheets.length > 1) {
          const localized = messagesFor(this.language), confirmText = localized['delete-sheet-confirm'];
          const choice = await vscode.window.showWarningMessage(localized['delete-sheet-warning'], { modal: true }, confirmText);
          if (choice !== confirmText) { postModel(); return; }
          const previousSheet = selectedSheet;
          try {
            document.applyDeleteSheet(message.sheetIndex);
            selectedSheet = Math.min(message.sheetIndex, document.state.sheets.length - 1);
            postModel();
          } catch (error) { selectedSheet = previousSheet; throw error; }
        } else if (message.type === 'sheet-visibility' && document.kind === 'xlsx' && Number.isInteger(message.sheetIndex) && typeof message.visible === 'boolean' && (message.visible || message.sheetIndex === selectedSheet)) {
          const previousSheet = selectedSheet;
          try {
            document.applySheetVisibility(message.sheetIndex, message.visible);
            if (message.visible) selectedSheet = message.sheetIndex;
            else if (selectedSheet === message.sheetIndex) selectedSheet = Math.max(0, document.state.sheets.findIndex(sheet => sheet.visible));
            postModel();
          } catch (error) { selectedSheet = previousSheet; throw error; }
        } else if (message.type === 'sheet-zoom' && document.kind === 'xlsx' && Number.isInteger(message.sheetIndex) && Number.isInteger(message.zoomScale) && message.sheetIndex === selectedSheet) { document.applySheetZoom(message.sheetIndex, message.zoomScale); postModel(); } else if (message.type === 'rename-sheet' && currentSheetMessage && document.kind === 'xlsx' && Number.isInteger(message.sheetIndex) && typeof message.name === 'string' && message.name.length <= 100) {
          document.applySheetName(message.sheetIndex, message.name);
        } else if (message.type === 'move-sheet' && currentSheetMessage && document.kind === 'xlsx' && Number.isInteger(message.sheetIndex) && Number.isInteger(message.targetIndex) && message.sheetIndex >= 0 && message.targetIndex >= 0 && message.sheetIndex < document.state.sheets.length && message.targetIndex < document.state.sheets.length) {
          const activeName = document.state.sheets[selectedSheet]?.name;
          document.applyMoveSheet(message.sheetIndex, message.targetIndex);
          selectedSheet = document.state.sheets.findIndex(sheet => sheet.name === activeName);
          postModel();
        } else if (message.type === 'ready' || message.type === 'refresh') {
          const requestedSheet = Number.isInteger(message.sheetIndex) ? message.sheetIndex : 0;
          if (!isValidSheetRequest(document, requestedSheet)) return;
          selectedSheet = requestedSheet;
          postModel();
        }
      } catch (error) {
        if (document.kind === 'xlsx' && !document.state.sheets[selectedSheet]) selectedSheet = Math.max(0, document.state.sheets.length - 1);
        let pending = error;
        const replaceContext = replacementContext || (['replace-current', 'replace-all'].includes(message.type) ? { requestId: message.requestId, sheetIndex: message.sheetIndex, query: message.query, mode: message.type } : null);
        const calculatedContext = message.type === 'calculated-column' ? { requestId: message.requestId, sheetIndex: message.sheetIndex, tableName: message.tableName, row: message.row, column: message.column, ...error.calculatedColumnResult } : null;
        const tabContext = message.type === 'table-tab' ? { requestId: message.requestId, sheetIndex: message.sheetIndex, tableName: message.tableName } : null;
        const textColumnsContext = ['text-to-columns', 'text-to-columns-preview'].includes(message.type) ? { requestId: message.requestId, sheetIndex: messageSheet, preview: message.type === 'text-to-columns-preview' } : null;
        const duplicateContext = ['remove-duplicates', 'remove-duplicates-preview'].includes(message.type) ? { requestId: message.requestId, sheetIndex: messageSheet } : null;
        while ((pending?.validationPrompt || pending?.tableFilterPrompt) && typeof pending.retry === 'function') {
          const korean = this.language === 'ko', { style } = pending.validationPrompt || {};
          const messageText = pending.validationPrompt?.message || (korean ? '이 값은 데이터 유효성 조건에 맞지 않습니다.' : 'This value does not satisfy the data validation rule.');
          let accepted = false;
          if (pending.tableFilterPrompt) {
            const localized = messagesFor(this.language), clearLabel = localized['table-tab-clear-filters'];
            accepted = await vscode.window.showWarningMessage(localized['table-tab-filter-prompt'].replace('{table}', pending.tableFilterPrompt.tableName), clearLabel, localized['table-tab-cancel']) === clearLabel;
          } else if (style === 'warning') {
            const continueLabel = korean ? '계속 입력' : 'Continue';
            const cancelLabel = korean ? '취소' : 'Cancel';
            accepted = await vscode.window.showWarningMessage(messageText, continueLabel, cancelLabel) === continueLabel;
          } else {
            const okLabel = korean ? '확인' : 'OK';
            accepted = await vscode.window.showInformationMessage(messageText, okLabel) === okLabel;
          }
          if (!accepted) { if (tabContext) panel.webview.postMessage({ type: 'table-tab-error', ...tabContext, message: messagesFor(this.language)['calculated-column-cancelled'] }); else if (calculatedContext) panel.webview.postMessage({ type: 'calculated-column-error', ...calculatedContext, message: messagesFor(this.language)['calculated-column-cancelled'] }); else if (replaceContext) panel.webview.postMessage({ type: 'replace-error', ...replaceContext, message: messagesFor(this.language)['replace-cancelled'] }); else if (textColumnsContext) panel.webview.postMessage({ type: 'text-to-columns-error', ...textColumnsContext, message: messagesFor(this.language)['text-to-columns-cancelled'] }); else postModel(); return; }
          try { const result = pending.retry(); postModel(); if (tabContext) panel.webview.postMessage({ type: 'table-tab-applied', ...tabContext, ...result }); else if (calculatedContext) panel.webview.postMessage({ type: 'calculated-column-applied', ...calculatedContext }); else if (textColumnsContext && !textColumnsContext.preview) panel.webview.postMessage({ type: 'text-to-columns-applied', requestId: textColumnsContext.requestId, sheetIndex: messageSheet }); else if (replaceContext) { replaceContext.sendResult?.(); return; } return; }
          catch (nextError) { pending = nextError; }
        }
        panel.webview.postMessage({ type: tabContext ? 'table-tab-error' : calculatedContext ? 'calculated-column-error' : replaceContext ? 'replace-error' : textColumnsContext ? 'text-to-columns-error' : duplicateContext ? 'remove-duplicates-error' : 'error', ...(tabContext || calculatedContext || replaceContext || textColumnsContext || duplicateContext || {}), message: pending.userMessage ? pending.message : translateDiagnostic(pending.message || pending, this.language) });
      }
    }, undefined, this.context.subscriptions);
    const sub = document.onDidRender(postModel);
    panel.onDidDispose(() => {
      sub.dispose();
      for (const pending of automaticFits.values()) pending.cancel();
      automaticFits.clear();
    });
  }

  saveCustomDocument(document, token) { return document.save(token).catch(error => { throw new Error(translateDiagnostic(error.message || error, this.language)); }); }
  saveCustomDocumentAs(document, destination, token) { return document.saveAs(destination, token).catch(error => { throw new Error(translateDiagnostic(error.message || error, this.language)); }); }
  revertCustomDocument(document) { return document.revert().catch(error => { throw new Error(translateDiagnostic(error.message || error, this.language)); }); }
  backupCustomDocument(document, context) { return document.backup(context); }

  activeDocumentHistory() {
    const uri = vscode.window.tabGroups.activeTabGroup.activeTab?.input?.uri;
    const document = uri && this.documents.get(uri.toString());
    const history = document && this.editHistory.get(document);
    return document && history ? { document, history } : null;
  }

  async undoActive() { return this.applyHistoryEdit('undo'); }
  async redoActive() { return this.applyHistoryEdit('redo'); }

  async applyHistoryEdit(direction) {
    const active = this.activeDocumentHistory();
    if (!active) return;
    const { document, history } = active;
    const source = direction === 'undo' ? history.undo : history.redo;
    const target = direction === 'undo' ? history.redo : history.undo;
    const edit = source.pop();
    if (!edit) return;
    history.applying = true;
    try {
      await edit[direction]();
      target.push(edit);
    } catch (error) {
      source.push(edit);
      throw error;
    } finally {
      history.applying = false;
    }
    // Undoing an edit after a save creates a new dirty state in VS Code's workbench.
    // Keep the extension's history authoritative because VS Code clears custom editor
    // undo entries at save checkpoints.
    if (document.dirty) {
      const inverse = direction === 'undo' ? edit.redo : edit.undo;
      const forward = direction === 'undo' ? edit.undo : edit.redo;
      this.emitter.fire({ document, label: translateDiagnostic(edit.label || '편집', this.language), undo: inverse, redo: forward });
    }
  }

  dispose() { this.emitter.dispose(); }
}

function matchesSelectedSheet(message, selectedSheet) {
  return Number.isInteger(message?.sheetIndex) && message.sheetIndex === selectedSheet;
}

function typedBlankKind(value, kind) {
  if (kind !== 'blank' || value === '') return kind;
  return typeof value === 'string' && /^\s*[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?\s*$/.test(value) && Number.isFinite(Number(value.trim())) ? 'number' : /^(?:TRUE|FALSE)$/i.test(value) ? 'boolean' : 'string';
}

function projectedFormulaCalculator(document, sheetIndex, edits) {
  const mutations = new Map(edits.map(edit => [edit.row + ':' + edit.column, edit])), drafts = new Map(document.drafts), formulaDrafts = new Map(document.formulaDrafts);
  for (const edit of edits) {
    const key = sheetIndex + ':' + addressFor(edit.row + 1, edit.column + 1);
    drafts.set(key, edit.value);
    if (Object.hasOwn(edit, 'formula')) formulaDrafts.set(key, edit.formula);
  }
  return createFormulaCalculator({ state: document.state, drafts, formulaDrafts, getCell(row, column, targetSheet = 0) {
    const cell = document.getCell(row, column, targetSheet), edit = targetSheet === sheetIndex && mutations.get(row + ':' + column);
    if (!edit) return cell;
    return { ...cell, row, column, address: addressFor(row + 1, column + 1), value: edit.value, kind: typedBlankKind(edit.value, edit.kind ?? cell?.kind ?? 'blank'), ...(Object.hasOwn(edit, 'formula') ? { formula: edit.formula, hasFormula: edit.formula !== null } : {}) };
  } });
}

function isValidSheetRequest(document, requestedSheet) {
  if (!Number.isInteger(requestedSheet) || requestedSheet < 0) return false;
  if (document.kind === 'csv') return requestedSheet === 0;
  if (document.kind === 'xlsx') return Boolean(document.state.sheets?.[requestedSheet]);
  return ['office', 'pdf'].includes(document.kind) && requestedSheet === 0;
}

function resolveInternalHyperlink(document, currentSheetIndex, location) {
  if (document?.kind !== 'xlsx' || !Number.isInteger(currentSheetIndex) || !document.state?.sheets?.[currentSheetIndex] || typeof location !== 'string' || location.length > 512) return null;
  const value = location.trim().replace(/^#/, '');
  let separator = -1, quoted = false;
  for (let index = 0; index < value.length; index++) {
    if (value[index] === "'") {
      if (quoted && value[index + 1] === "'") { index++; continue; }
      quoted = !quoted;
    } else if (value[index] === '!' && !quoted) {
      if (separator >= 0) return null;
      separator = index;
    }
  }
  if (quoted) return null;
  let sheetIndex = currentSheetIndex, reference = value;
  if (separator >= 0) {
    let name = value.slice(0, separator).trim();
    reference = value.slice(separator + 1).trim();
    if (name.startsWith("'") && name.endsWith("'")) name = name.slice(1, -1).replaceAll("''", "'");
    else if (name.includes("'")) return null;
    sheetIndex = document.state.sheets.findIndex(sheet => sheet.name.toLowerCase() === name.toLowerCase() && sheet.visible !== false);
    if (sheetIndex < 0) return null;
  }
  const addresses = reference.split(':');
  if (addresses.length > 2 || addresses.some(address => !address.trim())) return null;
  const parse = address => {
    const match = /^\$?([A-Za-z]{1,3})\$?([1-9]\d*)$/.exec(address.trim());
    if (!match) return null;
    let column = 0;
    for (const letter of match[1].toUpperCase()) column = column * 26 + letter.charCodeAt(0) - 64;
    const row = Number(match[2]) - 1;
    column--;
    return Number.isSafeInteger(row) && row >= 0 && row < XLSX_LIMITS.rows && column >= 0 && column < XLSX_LIMITS.columns ? { row, column } : null;
  };
  const start = parse(addresses[0]), end = addresses.length === 2 ? parse(addresses[1]) : start;
  if (!start || !end) return null;
  return { sheetIndex, row: start.row, column: start.column, range: { startRow: start.row, endRow: end.row, startColumn: start.column, endColumn: end.column } };
}

function externalHyperlinkUri(target) {
  if (typeof target !== 'string' || target.length > 2048 || /[\u0000-\u001f\u007f]/.test(target)) return null;
  try {
    const uri = vscode.Uri.parse(target, true), scheme = uri.scheme.toLowerCase();
    if (!['http', 'https', 'mailto', 'file'].includes(scheme)) return null;
    if ((scheme === 'http' || scheme === 'https') && (!uri.authority || uri.authority.includes('@'))) return null;
    if ((scheme === 'mailto' && !uri.path) || (scheme === 'file' && !uri.path)) return null;
    return uri;
  } catch { return null; }
}

function updateCsvHiddenAxis(document, sheetIndex, axis, action, start, end) {
  const maximum = axis === 'row' ? CSV_LIMIT_ROWS : axis === 'column' ? CSV_LIMIT_COLUMNS : 0, key = axis === 'row' ? 'csvHiddenRows' : 'csvHiddenColumns';
  if (!document?.writable || document.kind !== 'csv' || sheetIndex !== 0 || !maximum || !['hide', 'unhide-all'].includes(action)) throw new RangeError('CSV/TSV 행·열 숨김 요청이 유효하지 않습니다.');
  const before = new Set(document[key] || []), next = new Set(before);
  if (action === 'hide') {
    if (!Number.isInteger(start) || !Number.isInteger(end) || start < 0 || end < 0 || start >= maximum || end >= maximum || Math.abs(end - start) + 1 > maximum) throw new RangeError('CSV/TSV 행·열 숨김 범위가 유효하지 않습니다.');
    for (let index = Math.min(start, end); index <= Math.max(start, end); index++) next.add(index);
  } else next.clear();
  if (before.size === next.size && [...before].every(index => next.has(index))) return false;
  document[key] = next;
  const label = 'CSV/TSV ' + (axis === 'row' ? '행' : '열') + (action === 'hide' ? ' 숨김' : ' 숨김 해제'), after = new Set(next);
  document.emitter.fire({ label, undo: () => { document[key] = new Set(before); document.renderEmitter.fire(); }, redo: () => { document[key] = new Set(after); document.renderEmitter.fire(); } });
  document.renderEmitter.fire();
  return true;
}

function outlineGroupsForLevels(levels, collapsed, summaryAfter = true) {
  const groups = [];
  for (let level = 1; level <= 7; level++) {
    const coordinates = Object.keys(levels || {}).map(Number).filter(index => levels[index] >= level).sort((left, right) => left - right);
    let start = null, end = null;
    const push = () => {
      if (start === null) return;
      const summary = summaryAfter ? end + 1 : start - 1;
      groups.push({ start, end, level, summary, collapsed: collapsed?.has(summary) || false });
    };
    for (const coordinate of coordinates) {
      if (start === null) start = end = coordinate;
      else if (coordinate === end + 1) end = coordinate;
      else { push(); start = end = coordinate; }
    }
    push();
  }
  return groups.sort((left, right) => left.start - right.start || left.end - right.end || left.level - right.level);
}

function csvOutlineView(document) {
  const outline = document.csvOutline || {}, rowGroups = outlineGroupsForLevels(outline.rowOutlineLevels, outline.collapsedRows), columnGroups = outlineGroupsForLevels(outline.columnOutlineLevels, outline.collapsedColumns);
  const hiddenRows = [...new Set(rowGroups.filter(group => group.collapsed).flatMap(group => Array.from({ length: group.end - group.start + 1 }, (_, offset) => group.start + offset)))].sort((left, right) => left - right);
  const hiddenColumns = [...new Set(columnGroups.filter(group => group.collapsed).flatMap(group => Array.from({ length: group.end - group.start + 1 }, (_, offset) => group.start + offset)))].sort((left, right) => left - right);
  return { rowGroups, columnGroups, hiddenRows, hiddenColumns };
}

function cloneCsvOutline(outline) {
  return { rowOutlineLevels: { ...(outline?.rowOutlineLevels || {}) }, columnOutlineLevels: { ...(outline?.columnOutlineLevels || {}) }, collapsedRows: new Set(outline?.collapsedRows || []), collapsedColumns: new Set(outline?.collapsedColumns || []) };
}

function updateCsvOutlineGroups(document, sheetIndex, axis, action, start, end, summary, level) {
  const maximum = axis === 'row' ? CSV_LIMIT_ROWS : axis === 'column' ? CSV_LIMIT_COLUMNS : 0, summaryAction = action === 'collapse-summary' || action === 'expand-summary', levelAction = action === 'show-level';
  const validRange = summaryAction ? Number.isInteger(summary) && summary >= 0 && summary < maximum : levelAction ? Number.isInteger(level) && level >= 1 && level <= 8 : Number.isInteger(start) && Number.isInteger(end) && start >= 0 && end >= start && end < maximum && !(axis === 'row' && end - start + 1 > CSV_LIMIT_ROWS);
  if (!document?.writable || document.kind !== 'csv' || sheetIndex !== 0 || !['row', 'column'].includes(axis) || !['group', 'ungroup', 'collapse', 'expand', 'collapse-summary', 'expand-summary', 'show-level'].includes(action) || !validRange) throw new RangeError('CSV/TSV 행·열 그룹 요청이 유효하지 않습니다.');
  const before = cloneCsvOutline(document.csvOutline), next = cloneCsvOutline(document.csvOutline), levelsKey = axis === 'row' ? 'rowOutlineLevels' : 'columnOutlineLevels', collapsedKey = axis === 'row' ? 'collapsedRows' : 'collapsedColumns';
  const levels = next[levelsKey], collapsed = next[collapsedKey], groups = outlineGroupsForLevels(levels, collapsed), actionGroups = summaryAction ? groups.filter(group => group.summary === summary) : levelAction ? groups : [];
  if ((summaryAction || levelAction) && !actionGroups.length) throw new RangeError(summaryAction ? '선택한 요약 위치에 접거나 펼칠 그룹이 없습니다.' : '이 시트에 표시할 행·열 그룹이 없습니다.');
  const coordinates = summaryAction || levelAction ? (axis === 'column' ? [...new Set(actionGroups.flatMap(group => Array.from({ length: group.end - group.start + 1 }, (_, offset) => group.start + offset)))] : []) : Array.from({ length: end - start + 1 }, (_, offset) => start + offset);
  const levelUpdates = new Map(), collapsedUpdates = new Map();
  if (action === 'group') {
    const currentLevels = coordinates.map(index => levels[index] || 0), base = currentLevels[0];
    if (currentLevels.some(value => value !== base)) throw new RangeError('같은 개요 수준의 연속된 행·열만 새 그룹으로 묶을 수 있습니다.');
    if (base >= 7) throw new RangeError('행·열 그룹은 7단계까지만 만들 수 있습니다.');
    if (groups.some(group => group.collapsed && group.start <= end && group.end >= start)) throw new RangeError('접힌 그룹을 펼친 뒤 다시 그룹화하세요.');
    if (end + 1 >= maximum) throw new RangeError('요약 행·열을 둘 위치가 없어 그룹을 만들 수 없습니다.');
    for (const coordinate of coordinates) levelUpdates.set(coordinate, base + 1);
  } else if (action === 'ungroup') {
    const currentLevels = coordinates.map(index => levels[index] || 0);
    if (currentLevels.every(value => value === 0)) throw new RangeError('선택 범위에 해제할 그룹이 없습니다.');
    if (groups.some(group => group.collapsed && group.start <= end && group.end >= start)) throw new RangeError('그룹을 펼친 뒤 그룹 해제를 실행하세요.');
    for (const coordinate of coordinates) levelUpdates.set(coordinate, Math.max(0, (levels[coordinate] || 0) - 1));
  } else if (summaryAction) {
    const value = action === 'collapse-summary';
    if (actionGroups.every(group => group.collapsed === value)) throw new RangeError('선택한 요약 위치의 그룹은 이미 요청한 상태입니다.');
    collapsedUpdates.set(summary, value);
  } else if (levelAction) {
    const groupsBySummary = new Map();
    for (const group of groups) groupsBySummary.set(group.summary, [...(groupsBySummary.get(group.summary) || []), group]);
    for (const [coordinate, summaryGroups] of groupsBySummary) {
      const value = summaryGroups.some(group => group.level >= level);
      if (summaryGroups.some(group => group.collapsed !== value)) collapsedUpdates.set(coordinate, value);
    }
  } else {
    const matches = groups.filter(group => group.start === start && group.end === end && group.collapsed === (action === 'expand'));
    if (!matches.length) throw new RangeError('선택한 범위에 접거나 펼칠 그룹이 없습니다.');
    for (const group of matches) collapsedUpdates.set(group.summary, action === 'collapse');
  }
  for (const [coordinate, value] of levelUpdates) { if (value > 0) levels[coordinate] = value; else delete levels[coordinate]; }
  for (const [coordinate, value] of collapsedUpdates) { if (value) collapsed.add(coordinate); else collapsed.delete(coordinate); }
  const updatedGroups = outlineGroupsForLevels(levels, collapsed), validSummaries = new Set(updatedGroups.map(group => group.summary));
  for (const coordinate of collapsed) if (!validSummaries.has(coordinate)) collapsed.delete(coordinate);
  const changed = JSON.stringify(before.rowOutlineLevels) !== JSON.stringify(next.rowOutlineLevels) || JSON.stringify(before.columnOutlineLevels) !== JSON.stringify(next.columnOutlineLevels) || [...before.collapsedRows].sort((a, b) => a - b).join(',') !== [...next.collapsedRows].sort((a, b) => a - b).join(',') || [...before.collapsedColumns].sort((a, b) => a - b).join(',') !== [...next.collapsedColumns].sort((a, b) => a - b).join(',');
  if (!changed) return false;
  document.csvOutline = next;
  const subject = axis === 'row' ? '행' : '열', verb = { group: '그룹화', ungroup: '그룹 해제', collapse: '그룹 접기', expand: '그룹 펼치기', 'collapse-summary': '요약 그룹 모두 접기', 'expand-summary': '요약 그룹 모두 펼치기', 'show-level': '개요 표시 수준 변경' }[action];
  const after = cloneCsvOutline(next);
  document.emitter.fire({ label: 'CSV/TSV ' + subject + ' ' + verb, undo: () => { document.csvOutline = cloneCsvOutline(before); document.renderEmitter.fire(); }, redo: () => { document.csvOutline = cloneCsvOutline(after); document.renderEmitter.fire(); } });
  document.renderEmitter.fire();
  return true;
}

function hiddenRowsForSheet(document, sheetIndex) {
  if (document.kind === 'csv') {
    if (sheetIndex !== 0) return undefined;
    const filter = document.csvAutoFilter, hidden = new Set([...csvOutlineView(document).hiddenRows, ...(document.csvHiddenRows || [])]);
    if (filter?.columns.length) {
      for (const columnFilter of filter.columns) {
        const selected = new Set(columnFilter.values || []), matches = columnFilter.custom ? compileAutoFilterColumn(columnFilter, [], { coerceTextNumbers: true }) : null;
        for (let row = filter.top + 1; row <= filter.bottom; row++) {
          const cell = document.getCell(row, columnFilter.column, 0) || { value: '' }, value = String(cell.value ?? '');
          if (matches ? !matches({ ...cell, row, column: columnFilter.column }) : value === '' ? !columnFilter.blanks : !selected.has(value)) hidden.add(row);
        }
      }
    }
    return hidden.size ? [...hidden].sort((left, right) => left - right) : undefined;
  }
  if (document.kind !== 'xlsx') return undefined;
  const sheet = document.state.sheets[sheetIndex], filters = [sheet?.autoFilter, ...(sheet?.tables || []).map(table => table.autoFilter)].filter(filter => filter?.columns.length), hidden = new Set([...(sheet?.hiddenRows || []), ...(sheet?.outlineHiddenRows || [])]);
  if (filters.length) {
    const calculator = createFormulaCalculator(document);
    for (const filter of filters) {
      const predicates = filter.columns.map(column => ({ column: column.column, matches: compileAutoFilterColumn(column, () => filterColumnCells(filter, column.column, (row, col) => calculator.cell(sheetIndex, row, col)), { date1904: document.state.date1904, now: calculator.now, getStyle: cell => calculator.filterStyle(sheetIndex, cell.row, cell.column) }) }));
      for (let row = filter.top + 1; row <= filter.bottom; row++) {
        if (predicates.some(rule => !rule.matches({ ...(calculator.cell(sheetIndex, row, rule.column) ?? document.getCell(row, rule.column, sheetIndex)), row, column: rule.column }))) hidden.add(row);
      }
    }
  }
  return hidden.size ? [...hidden].sort((left, right) => left - right) : undefined;
}

function autoFilterOptions(document, sheetIndex, column, filterOverride = undefined) {
  if (document.kind === 'csv') {
    const filter = filterOverride ?? document.csvAutoFilter;
    const area = filter ? (filter.bottom - filter.top + 1) * (filter.right - filter.left + 1) : 0;
    if (sheetIndex !== 0 || !filter || ![filter.top, filter.bottom, filter.left, filter.right].every(Number.isInteger) || filter.top < 0 || filter.bottom <= filter.top || filter.bottom >= CSV_LIMIT_ROWS || filter.left < 0 || filter.right < filter.left || filter.right >= CSV_LIMIT_COLUMNS || !Number.isSafeInteger(area) || area > XLSX_LIMITS.cells || !Array.isArray(filter.columns) || !Number.isInteger(column) || column < filter.left || column > filter.right) throw new RangeError('선택한 열에 자동 필터가 없습니다.');
    const values = new Set();
    let blanks = false;
    for (let row = filter.top + 1; row <= filter.bottom; row++) {
      const value = String(document.getCell(row, column, 0)?.value ?? '');
      if (value === '') blanks = true; else values.add(value);
    }
    const sortedValues = [...values].sort(new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' }).compare);
    const current = filter.columns.find(item => item.column === column), selected = new Set(), matches = current?.custom ? compileAutoFilterColumn(current, [], { coerceTextNumbers: true }) : null;
    if (matches) for (let row = filter.top + 1; row <= filter.bottom; row++) {
      const cell = document.getCell(row, column, 0) || { value: '' };
      if (matches({ ...cell, row, column })) selected.add(String(cell.value ?? ''));
    }
    return {
      values: sortedValues,
      blanks,
      colors: { fills: [], fonts: [], fillCount: 0, fontCount: 0 },
      selectedValues: matches ? sortedValues.filter(value => selected.has(value)) : current?.values ? sortedValues.filter(value => current.values.includes(value)) : sortedValues,
      includeBlanks: matches ? matches({ value: '' }) : current?.blanks ?? blanks,
      ...(current?.custom ? { custom: { and: current.custom.and, conditions: current.custom.conditions.map(condition => ({ ...condition })) } } : {})
    };
  }
  const sheet = document.state.sheets[sheetIndex], filter = filterOverride ?? sheet?.autoFilter;
  const area = filter ? (filter.bottom - filter.top + 1) * (filter.right - filter.left + 1) : 0;
  if (!filter || ![filter.top, filter.bottom, filter.left, filter.right].every(Number.isInteger) || filter.top < 0 || filter.bottom < filter.top || filter.bottom === filter.top && !sheet?.tables?.some(table => { const range = tableFilterRange(table); return range && ['top','bottom','left','right'].every(key => range[key] === filter[key]); }) || filter.left < 0 || filter.right < filter.left || filter.bottom >= XLSX_LIMITS.rows || filter.right >= XLSX_LIMITS.columns || !Number.isSafeInteger(area) || area > XLSX_LIMITS.cells || !Array.isArray(filter.columns) || !Number.isInteger(column) || column < filter.left || column > filter.right) throw new RangeError('선택한 열에 자동 필터가 없습니다.');
  const calculator = createFormulaCalculator(document), values = new Set(), selected = new Set(), fills = new Map(), fonts = new Map(), icons = new Map(), current = filter.columns.find(item => item.column === column), matches = current?.custom || current?.top10 || current?.dateGroups || current?.dynamic || current?.colorFilter || current?.iconFilter ? compileAutoFilterColumn(current, () => filterColumnCells(filter, column, (row, col) => calculator.cell(sheetIndex, row, col)), { date1904: document.state.date1904, now: calculator.now, getStyle: cell => calculator.filterStyle(sheetIndex, cell.row, cell.column) }) : null;
  let blanks = false;
  for (let row = filter.top + 1; row <= filter.bottom; row++) {
    const cell = calculator.cell(sheetIndex, row, column), value = cell?.value ?? '';
    const style = calculator.filterStyle(sheetIndex, row, column);
    for (const [map, key] of [[fills, 'fill'], [fonts, 'color']]) { const color = style?.[key]?.toUpperCase() ?? null; map.set(color, (map.get(color) || 0) + 1); }
    if (style?.icon) { const key = style.icon.set + ':' + style.icon.id, entry = icons.get(key) || { iconSet: style.icon.set, iconId: style.icon.id, count: 0 }; entry.count++; icons.set(key, entry); }
    if (value === '') blanks = true; else values.add(String(value));
    if (matches?.({ ...cell, row, column })) selected.add(String(value));
  }
  const sortedValues = [...values].sort(new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' }).compare);
  const palette = map => [...map].sort(([a], [b]) => String(a || '').localeCompare(String(b || ''))).slice(0, 256).map(([color, count]) => ({ color, count }));
  return { values: sortedValues, blanks, colors: { fills: palette(fills), fonts: palette(fonts), fillCount: fills.size, fontCount: fonts.size }, ...(icons.size ? { iconFilters: [...icons.values()].sort((a, b) => a.iconSet.localeCompare(b.iconSet) || a.iconId - b.iconId) } : {}), selectedValues: matches ? sortedValues.filter(value => selected.has(value)) : current ? current.values.filter(value => values.has(value)) : sortedValues, includeBlanks: matches ? selected.has('') : current ? current.blanks : blanks, ...(current?.iconFilter ? { iconFilter: { ...current.iconFilter } } : {}), ...(current?.colorFilter ? { colorFilter: { ...current.colorFilter } } : {}), ...(current?.dynamic ? { dynamic: { ...current.dynamic } } : {}), ...(current?.dateGroups ? { dateGroups: current.dateGroups.map(group => ({ ...group })), dateValues: [...current.values] } : {}), ...(current?.top10 ? { top10: { ...current.top10 } } : {}), ...(current?.custom ? { custom: { and: current.custom.and, conditions: current.custom.conditions.map(condition => ({ ...condition })) } } : {}) };
}

function updateAutoFilterIconColumn(document, sheetIndex, column, iconFilter, filterOverride = undefined) {
  const current = filterOverride ?? document.state.sheets[sheetIndex]?.autoFilter;
  autoFilterOptions(document, sheetIndex, column, current); validateIconFilter(iconFilter);
  return { ...current, columns: [...current.columns.filter(item => item.column !== column), { column, iconFilter: { iconSet: iconFilter.iconSet, iconId: iconFilter.iconId } }] };
}

function updateAutoFilterColorColumn(document, sheetIndex, column, colorFilter, filterOverride = undefined) {
  const current = filterOverride ?? document.state.sheets[sheetIndex]?.autoFilter;
  autoFilterOptions(document, sheetIndex, column, current); validateColorFilter(colorFilter);
  return { ...current, columns: [...current.columns.filter(item => item.column !== column), { column, colorFilter: { cellColor: colorFilter.cellColor, color: colorFilter.color?.toUpperCase() ?? null } }] };
}

function updateAutoFilterDynamicColumn(document, sheetIndex, column, dynamic, filterOverride = undefined) {
  const current = filterOverride ?? document.state.sheets[sheetIndex]?.autoFilter;
  autoFilterOptions(document, sheetIndex, column, current); validateDynamicFilter(dynamic);
  return { ...current, columns: [...current.columns.filter(item => item.column !== column), { column, dynamic: { type: dynamic.type } }] };
}

function updateAutoFilterDateColumn(document, sheetIndex, column, dateGroups, values, blanks, filterOverride = undefined) {
  const current = filterOverride ?? document.state.sheets[sheetIndex]?.autoFilter;
  autoFilterOptions(document, sheetIndex, column, current); validateDateGroups(dateGroups);
  if (!dateGroups.length) throw new RangeError('날짜 그룹을 한 개 이상 입력하세요.');
  if (!Array.isArray(values) || values.length > XLSX_LIMITS.cells || values.some(value => typeof value !== 'string' || value.length > MAX_CELL_CHARS) || typeof blanks !== 'boolean') throw new RangeError('자동 필터 선택 값이 유효하지 않습니다.');
  return { ...current, columns: [...current.columns.filter(item => item.column !== column), { column, values: [...new Set(values)], blanks, dateGroups: dateGroups.map(group => ({ ...group })) }] };
}

function updateAutoFilterTopColumn(document, sheetIndex, column, top10, filterOverride = undefined) {
  const current = filterOverride ?? document.state.sheets[sheetIndex]?.autoFilter;
  autoFilterOptions(document, sheetIndex, column, current); validateTopFilter(top10);
  return { ...current, columns: [...current.columns.filter(item => item.column !== column), { column, top10: { top: top10.top, percent: top10.percent, value: top10.value } }] };
}

function updateAutoFilterCustomColumn(document, sheetIndex, column, custom, filterOverride = undefined) {
  if (document.kind === 'csv') {
    const current = filterOverride ?? document.csvAutoFilter;
    autoFilterOptions(document, sheetIndex, column, current);
    validateCustomFilter(custom); compileAutoFilterColumn({ column, custom }, [], { coerceTextNumbers: true });
    return { ...current, columns: [...current.columns.filter(item => item.column !== column), { column, custom: { and: custom.and, conditions: custom.conditions.map(condition => ({ ...condition })) } }] };
  }
  const current = filterOverride ?? document.state.sheets[sheetIndex]?.autoFilter;
  autoFilterOptions(document, sheetIndex, column, current);
  validateCustomFilter(custom); compileAutoFilterColumn({ column, custom });
  return { ...current, columns: [...current.columns.filter(item => item.column !== column), { column, custom: { and: custom.and, conditions: custom.conditions.map(condition => ({ ...condition })) } }] };
}

function updateAutoFilterColumn(document, sheetIndex, column, selectedValues, includeBlanks, clear = false, filterOverride = undefined) {
  if (document.kind === 'csv') {
    const current = filterOverride ?? document.csvAutoFilter;
    if (sheetIndex !== 0 || !current || column < current.left || column > current.right) throw new RangeError('선택한 열에 자동 필터가 없습니다.');
    const columns = current.columns.filter(item => item.column !== column);
    if (!clear) {
      const options = autoFilterOptions(document, sheetIndex, column, current), allowed = new Set(options.values);
      if (!Array.isArray(selectedValues) || selectedValues.length > XLSX_LIMITS.cells || selectedValues.some(value => typeof value !== 'string' || value.length > MAX_CELL_CHARS || !allowed.has(value)) || typeof includeBlanks !== 'boolean') throw new RangeError('자동 필터 선택 값이 유효하지 않습니다.');
      const uniqueValues = [...new Set(selectedValues)], selected = new Set(uniqueValues);
      const allValuesSelected = uniqueValues.length === options.values.length && options.values.every(value => selected.has(value));
      if (!allValuesSelected || includeBlanks !== options.blanks) columns.push({ column, values: uniqueValues, blanks: includeBlanks });
    }
    return { ...current, columns };
  }
  const sheet = document.state.sheets[sheetIndex], current = filterOverride ?? sheet?.autoFilter;
  const tableFilter = sheet?.tables?.some(table => { const range = tableFilterRange(table); return current && range && ['top', 'bottom', 'left', 'right'].every(key => range[key] === current[key]); });
  if (!sheet?.autoFilter && !tableFilter && clear) return null;
  if (!current || column < current.left || column > current.right) throw new RangeError('선택한 열에 자동 필터가 없습니다.');
  let columns = current.columns.filter(item => item.column !== column);
  if (!clear) {
    const options = autoFilterOptions(document, sheetIndex, column, current), allowed = new Set(options.values);
    if (!Array.isArray(selectedValues) || selectedValues.length > XLSX_LIMITS.cells || selectedValues.some(value => typeof value !== 'string' || !allowed.has(value)) || typeof includeBlanks !== 'boolean') throw new RangeError('자동 필터 선택 값이 유효하지 않습니다.');
    const uniqueValues = [...new Set(selectedValues)];
    const selected = new Set(uniqueValues);
    const allValuesSelected = uniqueValues.length === options.values.length && options.values.every(value => selected.has(value));
    if (!allValuesSelected || includeBlanks !== options.blanks) {
      columns.push({ column, values: uniqueValues, blanks: includeBlanks });
    }
  }
  return { ...current, columns };
}

function autoFilterRange(document, sheetIndex, row, column) {
  if (document.kind === 'csv') {
    const existing = document.csvAutoFilter;
    if (sheetIndex !== 0 || !Number.isInteger(row) || row < 0 || row >= CSV_LIMIT_ROWS || !Number.isInteger(column) || column < 0 || column >= CSV_LIMIT_COLUMNS) throw new RangeError('자동 필터 요청이 유효하지 않습니다.');
    if (existing && row >= existing.top && row <= existing.bottom && column >= existing.left && column <= existing.right) return existing;
    const region = currentRegion(document, row, column, sheetIndex);
    const area = (region.bottom - region.top + 1) * (region.right - region.left + 1);
    if (region.bottom <= region.top || region.right < region.left || !Number.isSafeInteger(area) || area > XLSX_LIMITS.cells) throw new RangeError('자동 필터에는 2행 이상의 100,000셀 이하 데이터 영역이 필요합니다.');
    return { ...region, columns: [] };
  }
  const sheet = document.state.sheets[sheetIndex], existing = sheet?.autoFilter;
  if (existing && row >= existing.top && row <= existing.bottom && column >= existing.left && column <= existing.right) return existing;
  const table = sheet?.tables?.find(item => row >= item.range.top && row <= item.range.bottom && column >= item.range.left && column <= item.range.right);
  if (table) {
    const range = tableFilterRange(table);
    if (!range || range.bottom < range.top) throw new RangeError('표 필터에는 표시된 머리글이 필요합니다.');
    return table.autoFilter || { ...range, columns: [] };
  }
  const region = currentRegion(document, row, column, sheetIndex);
  const area = (region.bottom - region.top + 1) * (region.right - region.left + 1);
  if (region.bottom <= region.top || region.right < region.left || !Number.isSafeInteger(area) || area > XLSX_LIMITS.cells) throw new RangeError('자동 필터에는 2행 이상의 100,000셀 이하 데이터 영역이 필요합니다.');
  return { ...region, columns: [] };
}

function excludeHiddenRowEdits(edits, hiddenRows) {
  if (!hiddenRows?.length) return edits;
  const hidden = new Set(hiddenRows);
  return edits.filter(edit => !hidden.has(edit.row));
}


function parsePrintTitleInputs(rows, columns) {
  if (typeof rows !== 'string' || typeof columns !== 'string' || rows.length > 32 || columns.length > 32) throw new RangeError('반복 제목은 50행·10열 이하의 유효한 범위를 입력하세요.');
  const titles = { rows: null, columns: null }, rowInput = rows.trim(), columnInput = columns.trim();
  if (rowInput) {
    const match = /^\$?([1-9]\d*)(?::\$?([1-9]\d*))?$/.exec(rowInput);
    if (!match) throw new RangeError('반복 제목은 50행·10열 이하의 유효한 범위를 입력하세요.');
    titles.rows = { start: Number(match[1]) - 1, end: Number(match[2] || match[1]) - 1 };
  }
  if (columnInput) {
    const match = /^\$?([A-Z]{1,3})(?::\$?([A-Z]{1,3}))?$/i.exec(columnInput);
    if (!match) throw new RangeError('반복 제목은 50행·10열 이하의 유효한 범위를 입력하세요.');
    titles.columns = { start: parseAddress(match[1].toUpperCase() + '1').column, end: parseAddress((match[2] || match[1]).toUpperCase() + '1').column };
  }
  return validatePrintTitles(titles);
}

function buildPrintDocument(document, sheetIndex = 0, range = null, language = 'en', budget = { cells: 0, characters: 0, output: { cells: 0, characters: 0 } }, includeCharts = true) {
  if (!['csv', 'xlsx'].includes(document.kind) || !Number.isInteger(sheetIndex) || sheetIndex < 0 || document.kind === 'xlsx' && !document.state.sheets[sheetIndex] || document.kind === 'csv' && sheetIndex !== 0) throw new RangeError('인쇄할 시트가 유효하지 않습니다.');
  const model = viewModel(document, sheetIndex, language), sheet = document.kind === 'xlsx' ? document.state.sheets[sheetIndex] : null;
  if (includeCharts && sheet?.chartPrintUnsupported) throw new RangeError('일부 차트를 해석할 수 없습니다. 차트 포함을 해제하고 인쇄하세요.');
  const geometry = includeCharts && sheet?.charts?.length ? chartPrintGeometry(model, sheet.chartAnchors || [], { ...sheet, defaultColumnWidth: xlsxDefaultColumnWidth(sheet, document.maximumDigitWidth), maximumDigitWidth: document.maximumDigitWidth, hiddenRows: sheet.nativeHiddenRows || sheet.hiddenRows }) : null;
  if (geometry && geometry.boxes.length !== model.charts.length) throw new RangeError('차트 위치를 해석할 수 없습니다. 차트 포함을 해제하고 인쇄하세요.');
  const printableBoxes = geometry?.boxes.filter(Boolean) || [];
  const usedRows = Math.max(model.rowCount, ...printableBoxes.map(box => box.bottom + 1)), usedColumns = Math.max(model.columnCount, ...printableBoxes.map(box => box.right + 1));
  if (sheet?.pageBreaksUnsupported && !sheet.printSettings?.fitToPage) throw new RangeError('지원하지 않는 페이지 나누기는 모두 지운 뒤 새로 설정하세요.');
  if (sheet?.headerFooterUnsupported) throw new RangeError('지원하지 않는 머리글·바닥글은 지운 뒤 새로 설정하세요.');
  if (!budget.headerContext) {
    const now = new Date();
    budget.headerContext = { fileName: path.basename(document.uri.fsPath), sheetName: sheet?.name || '', path: path.dirname(document.uri.fsPath) + path.sep, date: new Intl.DateTimeFormat(languageOf(language), { dateStyle: 'short' }).format(now), time: new Intl.DateTimeFormat(languageOf(language), { timeStyle: 'short' }).format(now) };
  }
  budget.headerOutput ||= { characters: 0 };
  if (sheet?.printTitlesUnsupported) throw new RangeError('저장된 반복 제목을 해석할 수 없어 인쇄를 중단했습니다.');
  const titles = validatePrintTitles(sheet?.printTitles || { rows: null, columns: null });
  const titleRows = titles.rows ? Array.from({ length: titles.rows.end - titles.rows.start + 1 }, (_, index) => titles.rows.start + index) : [];
  const titleColumns = titles.columns ? Array.from({ length: titles.columns.end - titles.columns.start + 1 }, (_, index) => titles.columns.start + index) : [];
  if (!range && sheet?.printAreaUnsupported) throw new RangeError('저장된 인쇄 영역을 해석할 수 없습니다. 인쇄할 범위를 직접 선택하세요.');
  if (!range && sheet?.printAreas?.length) {
    const documents = sheet.printAreas.map(area => {
      const bounds = { ...area };
      if (bounds.top === 0 && bounds.bottom === XLSX_LIMITS.rows - 1) bounds.bottom = Math.max(0, usedRows - 1);
      if (bounds.left === 0 && bounds.right === XLSX_LIMITS.columns - 1) bounds.right = Math.max(0, usedColumns - 1);
      try { return buildPrintDocument(document, sheetIndex, { startRow: bounds.top, endRow: bounds.bottom, startColumn: bounds.left, endColumn: bounds.right }, language, budget, includeCharts); }
      catch (error) { if (error.code === 'JINSHEET_PRINT_EMPTY') return null; throw error; }
    }).filter(Boolean);
    if (!documents.length) throw new RangeError('인쇄할 표시 셀이 없거나 인쇄 범위가 너무 큽니다.');
    return combinePrintDocuments(documents, language, { headerFooter: sheet.headerFooter, context: budget.headerContext, settings: sheet.printSettings || DEFAULT_PRINT_SETTINGS });
  }
  if (!range && (!usedRows || !usedColumns)) throw new RangeError('인쇄할 표시 셀이 없거나 인쇄 범위가 너무 큽니다.');
  const bounds = range ? selectionBounds(range, '인쇄', Number.MAX_SAFE_INTEGER) : { top: 0, bottom: Math.max(0, usedRows - 1), left: 0, right: Math.max(0, usedColumns - 1) };
  // Include a merged cell's anchor and whole extent when the selection intersects it.
  let expanded = true;
  while (expanded) {
    expanded = false;
    for (const merged of sheet?.mergedCells || []) if (merged.top <= bounds.bottom && merged.bottom >= bounds.top && merged.left <= bounds.right && merged.right >= bounds.left) {
      const next = { top: Math.min(bounds.top, merged.top), bottom: Math.max(bounds.bottom, merged.bottom), left: Math.min(bounds.left, merged.left), right: Math.max(bounds.right, merged.right) };
      if (Object.keys(next).some(key => next[key] !== bounds[key])) { Object.assign(bounds, next); expanded = true; }
    }
  }
  const maximumRows = sheet ? XLSX_LIMITS.rows : CSV_LIMIT_ROWS, maximumColumns = sheet ? XLSX_LIMITS.columns : CSV_LIMIT_COLUMNS;
  const count = (bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1);
  if (bounds.bottom >= maximumRows || bounds.right >= maximumColumns || !Number.isSafeInteger(count) || count > 20000 || bounds.bottom - bounds.top >= 2000 || bounds.right - bounds.left >= 100) throw new RangeError('인쇄는 2,000행·100열·20,000셀 이하 범위를 선택하세요.');
  const hiddenRows = new Set(model.hiddenRows || []), hiddenColumns = new Set(model.hiddenColumns || []);
  const bodyRows = Array.from({ length: bounds.bottom - bounds.top + 1 }, (_, index) => bounds.top + index), bodyColumns = Array.from({ length: bounds.right - bounds.left + 1 }, (_, index) => bounds.left + index);
  const rawRows = [...new Set([...bodyRows, ...titleRows])].sort((left, right) => left - right);
  const rawColumns = [...new Set([...bodyColumns, ...titleColumns])].sort((left, right) => left - right);
  budget.cells += rawRows.length * rawColumns.length;
  if (budget.cells > 20000) throw new RangeError('인쇄는 2,000행·100열·20,000셀 이하 범위를 선택하세요.');
  const rows = rawRows.filter(row => !hiddenRows.has(row)), columns = rawColumns.filter(column => !hiddenColumns.has(column));
  if (!bodyRows.some(row => !hiddenRows.has(row)) || !bodyColumns.some(column => !hiddenColumns.has(column))) {
    const error = new RangeError('인쇄할 표시 셀이 없거나 인쇄 범위가 너무 큽니다.'); error.code = 'JINSHEET_PRINT_EMPTY'; throw error;
  }
  for (const merged of sheet?.mergedCells || []) {
    if (!rows.some(row => row >= merged.top && row <= merged.bottom) || !columns.some(column => column >= merged.left && column <= merged.right)) continue;
    if (titles.rows && titles.rows.start <= merged.bottom && titles.rows.end >= merged.top && (titles.rows.start > merged.top || titles.rows.end < merged.bottom)
      || titles.columns && titles.columns.start <= merged.right && titles.columns.end >= merged.left && (titles.columns.start > merged.left || titles.columns.end < merged.right)) throw new RangeError('반복 제목 경계에 걸친 병합 셀 전체를 제목 범위에 포함하세요.');
  }
  const cells = new Map();
  for (let rowStart = 0; rowStart < rows.length; rowStart += 60) for (let columnStart = 0; columnStart < columns.length; columnStart += 30) {
    const viewRows = rows.slice(rowStart, rowStart + 60), columnIndexes = columns.slice(columnStart, columnStart + 30);
    const result = getRange(document, { sheetIndex, requestId: 1, rowStart: viewRows[0], columnStart: columnIndexes[0], rowCount: viewRows.length, columnCount: columnIndexes.length, viewRows, columnIndexes }, language);
    result.rows.forEach((items, rowIndex) => items.forEach((cell, columnIndex) => {
      budget.characters += String(cell?.displayValue ?? '').length;
      if (budget.characters > 2000000) throw new RangeError('인쇄할 텍스트가 너무 큽니다. 더 작은 범위를 선택하세요.');
      cells.set(viewRows[rowIndex] + ':' + columnIndexes[columnIndex], cell);
    }));
  }
  // Hidden merge anchors still supply the visible part of their merged cell.
  for (const merged of sheet?.mergedCells || []) if (!cells.has(merged.top + ':' + merged.left) && rows.some(row => row >= merged.top && row <= merged.bottom) && columns.some(column => column >= merged.left && column <= merged.right)) {
    const result = getRange(document, { sheetIndex, requestId: 1, rowStart: merged.top, columnStart: merged.left, rowCount: 1, columnCount: 1 }, language);
    budget.characters += String(result.rows[0][0]?.displayValue ?? '').length;
    if (budget.characters > 2000000) throw new RangeError('인쇄할 텍스트가 너무 큽니다. 더 작은 범위를 선택하세요.');
    cells.set(merged.top + ':' + merged.left, result.rows[0][0]);
  }
  const printChartHtml = chart => {
    const html = chartPreviewHtml([chart], language, false, true);
    budget.chartSourceCharacters = (budget.chartSourceCharacters || 0) + html.length;
    if (budget.chartSourceCharacters > 4000000) throw new RangeError('차트 인쇄 출력이 너무 큽니다. 범위를 줄이거나 차트 포함을 해제하세요.');
    return html;
  };
  const charts = geometry ? geometry.boxes.flatMap((box, index) => box && rows.some(row => geometry.rows.start(row) < box.y + box.height && geometry.rows.start(row) + geometry.rows.size(row) > box.y) && columns.some(column => geometry.columns.start(column) < box.x + box.width && geometry.columns.start(column) + geometry.columns.size(column) > box.x) ? [{ ...box, html: printChartHtml(model.charts[index]) }] : []) : [];
  const title = path.basename(document.uri.fsPath) + (sheet ? ' — ' + sheet.name : '');
  return renderPrintDocument({ title, rows, columns, cells, charts, rowPositions: geometry ? Object.fromEntries(rows.map(row => [row, geometry.rows.start(row)])) : {}, columnPositions: geometry ? Object.fromEntries(columns.map(column => [column, geometry.columns.start(column)])) : {}, mergedCells: sheet?.mergedCells || [], pageBreaks: sheet?.printSettings?.fitToPage ? { rows: [], columns: [] } : sheet?.pageBreaks, repeatRows: titleRows, repeatColumns: titleColumns, outputBudget: budget.output, headerFooter: sheet?.headerFooter, headerContext: budget.headerContext, headerBudget: budget.headerOutput, settings: sheet?.printSettings || DEFAULT_PRINT_SETTINGS, language,
    widths: Object.fromEntries(columns.map(column => [column, excelColumnWidthPixels(model.columnWidths?.[column], model.defaultColumnWidth ?? DEFAULT_COLUMN_WIDTH, model.maximumDigitWidth ?? 7)])),
    heights: Object.fromEntries(rows.map(row => [row, model.rowHeights?.[row] ?? model.defaultRowHeight ?? 26])),
    note: sheet?.charts?.length ? message(includeCharts ? 'print-charts-basic-note' : 'print-charts-note', language) : '' });
}

function viewModel(document, sheetIndex = 0, language = 'en') {
  const localizedReason = translateDiagnostic(document.reason, language);
  if (document.kind === 'office') return { type: 'model', kind: 'office', readOnly: true, reason: localizedReason, dirty: false, title: document.state.title, format: path.extname(document.uri.fsPath).slice(1), supported: !document.state.unsupported, sourceBytes: Uint8Array.from(document.originalBytes) };
  if (document.kind === 'pdf') return { type: 'model', kind: 'pdf', readOnly: true, reason: localizedReason, dirty: false, title: document.state.title, format: 'pdf', supported: !document.state.unsupported, sourceBytes: Uint8Array.from(document.originalBytes) };
  if (document.kind === 'csv') {
    if (!document.state.rows) return { type: 'model', kind: 'csv', readOnly: true, reason: localizedReason, rows: [] };
    let columnCount = 0;
    for (const row of document.state.rows) columnCount = Math.max(columnCount, row.length);
    const outline = csvOutlineView(document);
    return { type: 'model', kind: 'csv', readOnly: !document.writable, reason: localizedReason, dirty: document.dirty, rowCount: document.state.rows.length, columnCount, sheetIndex: 0, autoFilter: document.csvAutoFilter, hiddenRows: hiddenRowsForSheet(document, 0), hiddenColumns: [...new Set([...outline.hiddenColumns, ...(document.csvHiddenColumns || [])])].sort((left, right) => left - right), manuallyHiddenRows: [...(document.csvHiddenRows || [])].sort((left, right) => left - right), manuallyHiddenColumns: [...(document.csvHiddenColumns || [])].sort((left, right) => left - right), rowGroups: outline.rowGroups, columnGroups: outline.columnGroups };
  }
  const sheet = document.state.sheets[sheetIndex];
  if (!sheet) return { type: 'model', kind: 'xlsx', readOnly: true, reason: localizedReason || message('no-data', language), sheets: [] };
  let usedRows = 0, usedColumns = 0;
  for (const cell of sheet.cells.values()) { usedRows = Math.max(usedRows, cell.row + 1); usedColumns = Math.max(usedColumns, cell.column + 1); }
  for (const range of sheet.mergedCells || []) { usedRows = Math.max(usedRows, range.bottom + 1); usedColumns = Math.max(usedColumns, range.right + 1); }
  for (const [key] of document.drafts) {
    const [draftSheet, address] = key.split(':');
    if (Number(draftSheet) !== sheetIndex) continue;
    const coordinate = parseAddress(address);
    usedRows = Math.max(usedRows, coordinate.row + 1);
    usedColumns = Math.max(usedColumns, coordinate.column + 1);
  }
  for (const [key] of document.formulaDrafts || []) {
    const [draftSheet, address] = key.split(':');
    if (Number(draftSheet) !== sheetIndex) continue;
    const coordinate = parseAddress(address);
    usedRows = Math.max(usedRows, coordinate.row + 1);
    usedColumns = Math.max(usedColumns, coordinate.column + 1);
  }
  const chartCalculator = createFormulaCalculator(document);
  const spillExtent = chartCalculator.spillExtent(sheetIndex);
  if (spillExtent) { usedRows = Math.max(usedRows, spillExtent.bottom + 1); usedColumns = Math.max(usedColumns, spillExtent.right + 1); }
  const chartChanges = [...new Set([...(document.drafts?.keys() || []), ...(document.formulaDrafts?.keys() || [])])].map(key => {
    const separator = key.indexOf(':'), coordinate = parseAddress(key.slice(separator + 1));
    return { sheetIndex: Number(key.slice(0, separator)), row: coordinate.row, column: coordinate.column };
  });
  const charts = calculateChartPreviews(document.state, sheetIndex, chartCalculator, (row, column, targetSheet) => document.getCell(row, column, targetSheet), chartChanges);
  const columnWidths = { ...(sheet.columnWidths || {}) };
  for (const [key, width] of document.columnWidthDrafts || []) {
    const [draftSheet, column] = key.split(':').map(Number);
    if (draftSheet === sheetIndex) columnWidths[column] = width;
  }
  return {
    type: 'model', kind: 'xlsx', readOnly: !document.writable, reason: localizedReason, dirty: document.dirty,
    formulas: hasEffectiveFormulas(document), definedNames: document.state.definedNames || [], rowCount: usedRows, columnCount: usedColumns,
    sheetIndex, maximumDigitWidth: document.maximumDigitWidth, zoomScale: sheet.zoomScale ?? 100, printAreas: sheet.printAreas || [], printAreaUnsupported: Boolean(sheet.printAreaUnsupported), printTitles: sheet.printTitles || { rows: null, columns: null }, printTitlesUnsupported: Boolean(sheet.printTitlesUnsupported), pageBreaks: sheet.pageBreaks || { rows: [], columns: [], hasAny: false }, pageBreaksUnsupported: Boolean(sheet.pageBreaksUnsupported), headerFooter: sheet.headerFooter || defaultHeaderFooter(), headerFooterUnsupported: Boolean(sheet.headerFooterUnsupported), printSettings: sheet.printSettings, rowHeights: { ...(sheet.rowHeights || {}), ...Object.fromEntries((document.autoRowHeights?.get(sheetIndex) || []).map(item => [item.row, item.height * 4 / 3])) }, defaultRowHeight: sheet.defaultRowHeight || 20, defaultColumnWidth: xlsxDefaultColumnWidth(sheet, document.maximumDigitWidth), defaultColumnWidthCharacters: sheet.defaultColumnWidthCharacters ?? null, maximumDigitWidth: document.maximumDigitWidth, normalFont: document.state.normalFont, columnWidths, hiddenColumns: [...new Set([...(sheet.hiddenColumns || []), ...(sheet.outlineHiddenColumns || [])])], manuallyHiddenColumns: sheet.hiddenColumns || [], rowGroups: sheet.rowGroups || [], columnGroups: sheet.columnGroups || [], freezePane: sheet.freezePane || { rows: 0, columns: 0 }, autoFilter: sheet.autoFilter || sheet.tables?.find(table => table.autoFilter)?.autoFilter || null, hasDynamicDateFilters: [sheet.autoFilter, ...(sheet.tables || []).map(table => table.autoFilter)].some(filter => filter?.columns.some(item => dynamicFilterUsesClock(item.dynamic))), hasTables: Boolean(sheet.hasTables), tables: (sheet.tables || []).map(table => ({ name: table.name, range: table.range, headerRowCount: table.headerRowCount, totalsRowCount: table.totalsRowCount, styleInfo: table.styleInfo, columns: table.columns, calculatedColumns: table.calculatedColumns, totalColumns: table.totalColumns.map((entry, index) => ({ ...entry, function: table.totalsRowCount ? totalsFunctionForCell(table, table.range.left + index, document.getCell(table.range.bottom, table.range.left + index, sheetIndex)) : entry.function })) })), mergedCells: sheet.mergedCells || [], hyperlinks: (sheet.hyperlinks || []).map(({ top, bottom, left, right }) => ({ top, bottom, left, right })), charts, chartInsertionAvailable: charts.length < 32, chartPlacements: sheet.chartPlacements || [], chartPreviewHtml: chartPreviewHtml(charts, language, document.writable, false, sheet.chartPlacements), hiddenRows: hiddenRowsForSheet(document, sheetIndex), manuallyHiddenRows: sheet.hiddenRows || [], sheets: document.state.sheets.map((item, index) => ({ name: item.name, index, visible: item.visible }))
  };
}

function hasEffectiveFormulas(document) {
  if (!(document.formulaDrafts instanceof Map)) return Boolean(document.state.formulas);
  for (let sheetIndex = 0; sheetIndex < document.state.sheets.length; sheetIndex++) {
    for (const cell of document.state.sheets[sheetIndex].cells.values()) {
      if (isFormulaCell(cell) && isFormulaCell(document.getCell(cell.row, cell.column, sheetIndex))) return true;
    }
  }
  for (const formula of document.formulaDrafts.values()) if (formula !== null) return true;
  return false;
}

function chartPreviewHtml(charts, language = 'en', editable = false, forPrint = false, placements = []) {
  const escape = value => String(value).replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&#39;');
  const palette = ['#4f81bd', '#c0504d', '#9bbb59', '#8064a2', '#4bacc6', '#f79646', '#33558b', '#7f3f3f'];
  if (forPrint && charts?.length > 32) throw new RangeError('시트당 32개 차트까지 인쇄할 수 있습니다. 차트 포함을 해제하고 인쇄하세요.');
  const chartColor = index => forPrint && index >= palette.length ? `hsl(${Math.round(index * 137.508) % 360},60%,40%)` : palette[index % palette.length];
  return (charts || []).slice(0, 32).map((chart, chartIndex) => {
    if (forPrint && ((chart.series || []).length > 32 || (chart.series || []).some(item => item.values.length > 10000))) throw new RangeError('차트 인쇄는 32계열·계열당 10,000개 데이터까지 지원합니다. 차트 포함을 해제하고 인쇄하세요.');
    if (forPrint && ['pie', 'doughnut'].includes(chart.type) && chart.series.length !== 1) throw new RangeError('원형·도넛 인쇄는 1계열을 지원합니다. 차트 포함을 해제하고 인쇄하세요.');
    if (forPrint && (chart.series || []).reduce((sum, item) => sum + item.values.length, 0) > 50000) throw new RangeError('차트 인쇄 데이터는 차트당 합계 50,000개까지 지원합니다. 차트 포함을 해제하고 인쇄하세요.');
    let markupCharacters = 0;
    const boundedMarkup = html => { markupCharacters += html.length; if (forPrint && markupCharacters > 4000000) throw new RangeError('차트 인쇄 출력이 너무 큽니다. 차트 포함을 해제하고 인쇄하세요.'); return html; };
    const pushMark = (target, html) => target.push(boundedMarkup(html));
    const legend = (name, index, color) => { const x = 28 + index % 4 * 152, y = 174 + Math.floor(index / 4) * 16; return `<g data-print-legend="${index}"><rect x="${x}" y="${y}" width="9" height="9" fill="${color}"/><text x="${x + 12}" y="${y + 9}" textLength="${Math.min(132, String(name).length * 6)}" lengthAdjust="spacingAndGlyphs">${escape(name)}</text></g>`; };
    const placementControl = `<label>${language === 'ko' ? '셀 크기 변경 시 차트' : 'Chart behavior on cell resize'} <select data-chart-placement>${[['twoCell', language === 'ko' ? '함께 이동하고 크기 조정' : 'Move and size with cells'], ['oneCell', language === 'ko' ? '함께 이동하고 크기 고정' : 'Move with cells; fixed size'], ['absolute', language === 'ko' ? '위치와 크기 고정' : 'Fixed position and size']].map(([value, label]) => `<option value="${value}"${(placements[chartIndex] || 'twoCell') === value ? ' selected' : ''}>${label}</option>`).join('')}</select></label><button type="button" data-chart-placement-apply>${language === 'ko' ? '배치 적용' : 'Apply placement'}</button>`;
    const editor = editable ? `<div class="chart-metadata"><label>${language === 'ko' ? '제목' : 'Title'} <input data-chart-title maxlength="512" value="${escape(chart.title || '')}"></label>${(chart.series || []).map((item, index) => `<label>${language === 'ko' ? '계열' : 'Series'} ${index + 1} <input data-chart-series-name maxlength="256" value="${escape(item.name || '')}"></label>`).join('')}<button type="button" data-chart-edit>${language === 'ko' ? '차트 정보 적용' : 'Apply chart details'}</button><button type="button" data-chart-source title="${language === 'ko' ? '머리글 행과 범주 열을 포함한 새 데이터 범위를 먼저 선택하세요.' : 'First select new data including a header row and category column.'}">${language === 'ko' ? '선택 범위를 원본으로' : 'Use selected range'}</button><button type="button" data-chart-delete>${language === 'ko' ? '차트 삭제' : 'Delete chart'}</button>${placementControl}</div>` : '';
    const emptyChart = () => `<figure class="chart-card" data-chart-index="${chartIndex}"><figcaption>${escape(chart.title || (language === 'ko' ? '차트 미리보기' : 'Chart preview'))}</figcaption><div class="chart-empty">${language === 'ko' ? '표시할 차트 데이터가 없습니다.' : 'No chart data to display.'}</div>${editor}</figure>`;
    const series = (chart.series || []).slice(0, forPrint ? 32 : 8), count = Math.min(forPrint ? 10000 : 60, Math.max(0, ...series.map(item => item.values?.length || 0)));
    const svgHeight = forPrint ? Math.max(204, 188 + Math.ceil(series.length / 4) * 16) : 204;
    const pointAttributes = (item, seriesIndex, index, value) => forPrint ? ` data-chart-point="${index}" data-chart-series="${seriesIndex}" aria-label="${escape((item.name || `Series ${seriesIndex + 1}`) + ': ' + (item.categories?.[index] ?? index + 1) + ': ' + value)}"` : '';
    if (!count || !series.length) return emptyChart();
    const labels = Array.from({ length: count }, (_, index) => String(series[0].categories?.[index] ?? index + 1).slice(0, forPrint ? 256 : 32));
    const values = series.flatMap(item => item.values.slice(0, count).filter(value => typeof value === 'number' && Number.isFinite(value)));
    if (!values.length) return emptyChart();
    if (chart.type === 'scatter') {
      const xValues = series.flatMap(item => (item.xValues || []).slice(0, count).filter(value => typeof value === 'number' && Number.isFinite(value)));
      if (!xValues.length) return emptyChart();
      const minX = Math.min(...xValues), maxX = Math.max(...xValues), minY = Math.min(...values), maxY = Math.max(...values);
      const rangeX = maxX - minX || 1, rangeY = maxY - minY || 1, left = 48, top = 12, width = 548, height = 142;
      if (forPrint && (!Number.isFinite(rangeX) || !Number.isFinite(rangeY))) throw new RangeError('차트 수치 범위가 너무 큽니다. 차트 포함을 해제하고 인쇄하세요.');
      const marks = [], legends = [];
      for (const [seriesIndex, item] of series.entries()) {
        const color = chartColor(seriesIndex);
        pushMark(legends, forPrint ? legend(item.name || `Series ${seriesIndex + 1}`, seriesIndex, color) : `<g><circle cx="${left + seriesIndex * 122 + 5}" cy="180" r="4" fill="${color}"/><text x="${left + 14 + seriesIndex * 122}" y="183">${escape(item.name || `Series ${seriesIndex + 1}`).slice(0, 80)}</text></g>`);
        (item.xValues || []).slice(0, count).forEach((x, index) => {
          const y = item.values[index];
          if (!Number.isFinite(x) || !Number.isFinite(y)) return;
          const px = left + (x - minX) / rangeX * width, py = top + (maxY - y) / rangeY * height;
          pushMark(marks, `<circle cx="${px.toFixed(2)}" cy="${py.toFixed(2)}" r="4" fill="${color}"${pointAttributes(item, seriesIndex, index, x + ', ' + y)}/>`);
        });
      }
      const title = escape(chart.title || (language === 'ko' ? '차트 미리보기' : 'Chart preview'));
      const axes = `<line x1="${left}" y1="${top}" x2="${left}" y2="${top + height}" stroke="currentColor" opacity=".35"/><line x1="${left}" y1="${top + height}" x2="${left + width}" y2="${top + height}" stroke="currentColor" opacity=".35"/>`;
      return `<figure class="chart-card" data-chart-index="${chartIndex}"><figcaption>${title}</figcaption><svg viewBox="0 0 640 ${svgHeight}" role="img" aria-label="${title}">${axes}${marks.join('')}${legends.join('')}</svg>${editor}</figure>`;
    }
    if (chart.type === 'pie' || chart.type === 'doughnut') {
      const item = series[0], entries = item.values.slice(0, count).map((value, index) => ({ index, original: value, value: forPrint && Number.isFinite(value) ? Math.abs(value) : value, label: labels[index] })).filter(entry => Number.isFinite(entry.value) && entry.value > 0);
      const total = entries.reduce((sum, entry) => sum + entry.value, 0);
      if (forPrint && !Number.isFinite(total)) throw new RangeError('차트 수치 합계가 너무 큽니다. 차트 포함을 해제하고 인쇄하세요.');
      if (!total) return emptyChart();
      let angle = -Math.PI / 2;
      const cx = 320, cy = 91, radius = 68;
      const slices = entries.map((entry, index) => {
        const start = angle, end = angle + entry.value / total * Math.PI * 2;
        angle = end;
        const x1 = cx + radius * Math.cos(start), y1 = cy + radius * Math.sin(start), x2 = cx + radius * Math.cos(end), y2 = cy + radius * Math.sin(end);
        const large = end - start > Math.PI ? 1 : 0, color = chartColor(forPrint ? entry.index : index), precision = forPrint ? 6 : 2;
        const attributes = pointAttributes(item, 0, entry.index, entry.original);
        const path = forPrint && entries.length === 1 ? `<circle cx="${cx}" cy="${cy}" r="${radius}" fill="${color}"${attributes}/>` : `<path d="M ${cx} ${cy} L ${x1.toFixed(precision)} ${y1.toFixed(precision)} A ${radius} ${radius} 0 ${large} 1 ${x2.toFixed(precision)} ${y2.toFixed(precision)} Z" fill="${color}"${attributes}/>`;
        const legend = index < 12 ? forPrint ? `<g data-print-legend="${entry.index}"><rect x="${20 + index % 3 * 200}" y="${176 + Math.floor(index / 3) * 16}" width="9" height="9" fill="${color}"/><text x="${33 + index % 3 * 200}" y="${185 + Math.floor(index / 3) * 16}" textLength="${Math.min(175, entry.label.length * 6)}" lengthAdjust="spacingAndGlyphs">${escape(entry.label)}</text></g>` : `<g><rect x="${170 + index % 3 * 150}" y="${176 + Math.floor(index / 3) * 16}" width="9" height="9" fill="${color}"/><text x="${183 + index % 3 * 150}" y="${185 + Math.floor(index / 3) * 16}">${escape(entry.label).slice(0, 32)}</text></g>` : '';
        return { path: boundedMarkup(path), legend: boundedMarkup(legend) };
      });
      const hole = chart.type === 'doughnut' ? `<circle cx="${cx}" cy="${cy}" r="34" fill="var(--vscode-editor-background, white)"/>` : '';
      const summary = forPrint && entries.length > 12 ? `<text data-legend-summary="true" x="320" y="256" text-anchor="middle">${language === 'ko' ? `모든 ${entries.length}개 조각 표시 · 범례 ${entries.length - 12}개 생략` : `All ${entries.length} slices shown; ${entries.length - 12} legend labels omitted`}</text>` : '';
      const title = escape(chart.title || (language === 'ko' ? '차트 미리보기' : 'Chart preview'));
      return `<figure class="chart-card" data-chart-index="${chartIndex}"><figcaption>${title}</figcaption><svg viewBox="0 0 640 ${forPrint ? summary ? 270 : 248 : 220}" role="img" aria-label="${title}">${slices.map(slice => slice.path).join('')}${hole}${slices.map(slice => slice.legend).join('')}${summary}</svg>${editor}</figure>`;
    }
    const low = Math.min(0, ...values), high = Math.max(0, ...values), range = high - low || 1;
    if (forPrint && !Number.isFinite(range)) throw new RangeError('차트 수치 범위가 너무 큽니다. 차트 포함을 해제하고 인쇄하세요.');
    const horizontal = chart.type === 'bar' && chart.direction === 'horizontal';
    const left = horizontal ? 112 : 42, top = 12, width = horizontal ? 500 : 570, height = 142, baseY = top + high / range * height, baseX = left + (-low) / range * width;
    const escLabels = labels.map(label => escape(label.slice(0, 32)));
    const marks = [], legends = [];
    for (const [seriesIndex, item] of series.entries()) {
      const color = chartColor(seriesIndex);
      pushMark(legends, forPrint ? legend(item.name || `Series ${seriesIndex + 1}`, seriesIndex, color) : `<g><rect x="${left + seriesIndex * 122}" y="174" width="10" height="10" fill="${color}"/><text x="${left + 14 + seriesIndex * 122}" y="183">${escape(item.name || `Series ${seriesIndex + 1}`).slice(0, 80)}</text></g>`);
      if (chart.type === 'line' || chart.type === 'area') {
        if (forPrint) {
          const segments = []; let segment = [];
          for (const [index, value] of item.values.entries()) {
            if (!Number.isFinite(value)) { if (segment.length) segments.push(segment); segment = []; continue; }
            segment.push([left + (index + .5) * width / count, top + (high - value) / range * height]);
          }
          if (segment.length) segments.push(segment);
          for (const points of segments) {
            if (chart.type === 'area') pushMark(marks, `<polygon points="${points[0][0]},${baseY} ${points.map(point => point.join(',')).join(' ')} ${points.at(-1)[0]},${baseY}" fill="${color}" opacity=".22"/>`);
            if (points.length === 1) pushMark(marks, `<circle cx="${points[0][0]}" cy="${points[0][1]}" r="2" fill="${color}"/>`);
            else pushMark(marks, `<polyline points="${points.map(point => point.join(',')).join(' ')}" fill="none" stroke="${color}" stroke-width="2"/>`);
          }
          continue;
        }
        const points = item.values.slice(0, count).map((value, index) => typeof value === 'number' && Number.isFinite(value) ? `${left + (index + 0.5) * width / count},${top + (high - value) / range * height}` : null).filter(Boolean);
        if (chart.type === 'area' && points.length) {
          const firstX = left + (item.values.findIndex(value => typeof value === 'number' && Number.isFinite(value)) + 0.5) * width / count;
          const lastIndex = item.values.slice(0, count).map((value, index) => typeof value === 'number' && Number.isFinite(value) ? index : -1).filter(index => index >= 0).at(-1);
          const lastX = left + (lastIndex + 0.5) * width / count;
          pushMark(marks, `<polygon points="${firstX},${baseY} ${points.join(' ')} ${lastX},${baseY}" fill="${color}" opacity=".22"/>`);
        }
        if (points.length) pushMark(marks, `<polyline points="${points.join(' ')}" fill="none" stroke="${color}" stroke-width="2"/>`);
        continue;
      }
      item.values.slice(0, count).forEach((value, index) => {
        if (typeof value !== 'number' || !Number.isFinite(value)) return;
        if (forPrint && value === 0) return;
        const precision = forPrint ? 4 : 1;
        if (horizontal) {
          const band = height / count, barHeight = Math.max(forPrint ? 0 : 1, band * 0.76 / series.length), y = top + index * band + band * 0.12 + seriesIndex * barHeight;
          const valueX = left + (value - low) / range * width, x = Math.min(baseX, valueX), barWidth = Math.max(forPrint ? 0 : 1, Math.abs(baseX - valueX));
          pushMark(marks, `<rect x="${x.toFixed(precision)}" y="${y.toFixed(precision)}" width="${barWidth.toFixed(precision)}" height="${barHeight.toFixed(precision)}" fill="${color}"${pointAttributes(item, seriesIndex, index, value)}/>`);
        } else {
          const band = width / count, barWidth = Math.max(forPrint ? 0 : 1, band * 0.76 / series.length), x = left + index * band + band * 0.12 + seriesIndex * barWidth;
          const valueY = top + (high - value) / range * height, y = Math.min(baseY, valueY), barHeight = Math.max(forPrint ? 0 : 1, Math.abs(baseY - valueY));
          pushMark(marks, `<rect x="${x.toFixed(precision)}" y="${y.toFixed(precision)}" width="${barWidth.toFixed(precision)}" height="${barHeight.toFixed(precision)}" fill="${color}"${pointAttributes(item, seriesIndex, index, value)}/>`);
        }
      });
    }
    const ticks = escLabels.map((label, index) => forPrint && index % Math.ceil(count / 12) ? '' : horizontal
      ? `<text text-anchor="end" x="${left - 8}" y="${(top + (index + 0.5) * height / count + 4).toFixed(1)}">${label}</text>`
      : `<text text-anchor="middle" x="${(left + (index + 0.5) * width / count).toFixed(1)}" y="166">${label}</text>`).join('');
    const title = escape(chart.title || (language === 'ko' ? '차트 미리보기' : 'Chart preview'));
    const axes = horizontal
      ? `<line x1="${baseX.toFixed(1)}" y1="${top}" x2="${baseX.toFixed(1)}" y2="${top + height}" stroke="currentColor" opacity=".35"/>`
      : `<line x1="${left}" y1="${top}" x2="${left}" y2="${top + height}" stroke="currentColor" opacity=".35"/><line x1="${left}" y1="${baseY.toFixed(1)}" x2="${left + width}" y2="${baseY.toFixed(1)}" stroke="currentColor" opacity=".35"/>`;
    return `<figure class="chart-card" data-chart-index="${chartIndex}"><figcaption>${title}</figcaption><svg viewBox="0 0 640 ${svgHeight}" role="img" aria-label="${title}">${axes}${marks.join('')}${ticks}${legends.join('')}</svg>${editor}</figure>`;
  }).join('');
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
  const sheet = document.state.sheets[sheetIndex], seen = new Set(), calculator = createFormulaCalculator(document);
  for (const cell of sheet.cells.values()) {
    if (cell.column !== column) continue;
    const current = calculator.cell(sheetIndex, cell.row, column);
    values.push(formatCellValue(current, language));
    seen.add(cell.address);
  }
  for (const [key] of document.drafts) {
    const [draftSheet, address] = key.split(':');
    if (Number(draftSheet) !== sheetIndex || seen.has(address)) continue;
    const coordinate = parseAddress(address);
    if (coordinate.column !== column) continue;
    values.push(formatCellValue(calculator.cell(sheetIndex, coordinate.row, column), language));
  }
  for (const [key] of document.formulaDrafts || []) {
    const [draftSheet, address] = key.split(':');
    if (Number(draftSheet) !== sheetIndex || seen.has(address)) continue;
    const coordinate = parseAddress(address);
    if (coordinate.column !== column) continue;
    values.push(formatCellValue(calculator.cell(sheetIndex, coordinate.row, column), language));
  }
  return estimateBestFitWidth(values, { maximum: 1785 });
}

function tableStyleForCell(sheet, row, column) {
  const table = sheet?.tables?.find(item => row >= item.range.top && row <= item.range.bottom && column >= item.range.left && column <= item.range.right);
  const info = table?.styleInfo, match = /^TableStyle(Light|Medium|Dark)(\d+)$/.exec(info?.name || '');
  if (!table || !match || !TABLE_STYLE_NAMES.includes(info.name)) return null;
  const family = match[1], index = Number(match[2]), paletteIndex = Math.floor((index - 1) / 4) % 7;
  const accents = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#70AD47', '#5B9BD5', '#8064A2'];
  const accent = accents[paletteIndex];
  const lightRows = ['#F2F2F2', '#FCE4D6', '#E7E6E6', '#FFF2CC', '#E2F0D9', '#DDEBF7', '#E4DFEC'];
  const lightBands = ['#F7F7F7', '#FDF1EA', '#F3F2F2', '#FFF9E6', '#F0F7EB', '#EFF5FB', '#F4F1F8'];
  const mediumRows = ['#D9E2F3', '#FBE5D6', '#E7E6E6', '#FFF2CC', '#E2F0D9', '#DDEBF7', '#E4DFEC'];
  const darkRows = ['#D9E2F3', '#FBE5D6', '#E7E6E6', '#FFF2CC', '#E2F0D9', '#DDEBF7', '#E4DFEC'];
  const header = family === 'Light' ? (index === 1 ? '#F2F2F2' : lightRows[paletteIndex]) : family === 'Medium' ? accent : ['#404040', '#1F4E78', '#7030A0', '#7F6000', '#385723', '#1F4E78', '#4F3B60'][paletteIndex];
  const font = family === 'Dark' || family === 'Medium' ? '#FFFFFF' : paletteIndex === 0 && index === 1 ? '#222222' : accent;
  let fill = null, color = null;
  const headerRow = table.headerRowCount ? row === table.range.top : false;
  const totalsRow = table.totalsRowCount ? row === table.range.bottom : false;
  const firstDataRow = table.range.top + (table.headerRowCount || 0), lastDataRow = table.range.bottom - (table.totalsRowCount || 0);
  if (headerRow || totalsRow) { fill = header; color = font; }
  else if (row >= firstDataRow && row <= lastDataRow) {
    const dataIndex = row - firstDataRow, columnIndex = column - table.range.left;
    if (info.showRowStripes && dataIndex % 2 === 0) fill = family === 'Light' ? lightBands[paletteIndex] : family === 'Medium' ? mediumRows[paletteIndex] : darkRows[paletteIndex];
    if (info.showColumnStripes && columnIndex % 2 === 0) fill = family === 'Light' ? lightRows[paletteIndex] : mediumRows[paletteIndex];
    if (info.showFirstColumn && columnIndex === 0 || info.showLastColumn && columnIndex === table.range.right - table.range.left) fill ||= family === 'Light' ? lightRows[paletteIndex] : mediumRows[paletteIndex];
  }
  return fill || color ? { fill, color } : null;
}

function getRange(document, message, language = 'en') {
  const { rowStart, columnStart, rowCount, columnCount, requestId } = message;
  if (![rowStart, columnStart, rowCount, columnCount, requestId].every(Number.isInteger) || rowStart < 0 || columnStart < 0 || rowCount < 1 || rowCount > 60 || columnCount < 1 || columnCount > 30) throw new RangeError('조회 범위가 유효하지 않습니다.');
  const sheetIndex = Number.isInteger(message.sheetIndex) ? message.sheetIndex : 0;
  const viewRows = message.viewRows ?? Array.from({ length: rowCount }, (_, index) => rowStart + index);
  const viewIndexes = message.viewIndexes ?? viewRows;
  const columnIndexes = message.columnIndexes ?? Array.from({ length: columnCount }, (_, index) => columnStart + index);
  // CSV content has a bounded used range, while its editable grid extends to
  // the supported worksheet limits so empty cells remain directly selectable.
  const maximumRows = document.kind === 'csv' ? CSV_LIMIT_ROWS : XLSX_LIMITS.rows;
  const maximumColumns = document.kind === 'csv' ? CSV_LIMIT_COLUMNS : XLSX_LIMITS.columns;
  if (!Array.isArray(viewRows) || viewRows.length !== rowCount || viewRows.some((row, index) => !Number.isInteger(row) || row < 0 || row >= maximumRows || (index && row <= viewRows[index - 1]))) throw new RangeError('요청한 행 인덱스가 유효하지 않습니다.');
  if (!Array.isArray(viewIndexes) || viewIndexes.length !== rowCount || viewIndexes.some((row, index) => !Number.isInteger(row) || row < 0 || (index && row <= viewIndexes[index - 1]))) throw new RangeError('표시 행 인덱스가 유효하지 않습니다.');
  if (!Array.isArray(columnIndexes) || columnIndexes.length !== columnCount || columnIndexes.some((column, index) => !Number.isInteger(column) || column < 0 || column >= maximumColumns || (index && column <= columnIndexes[index - 1]))) throw new RangeError('요청한 열 인덱스가 유효하지 않습니다.');
  const rowIndexes = viewRows;
  const calculator = document.kind === 'xlsx' ? createFormulaCalculator(document) : null;
  const conditionalCellValue = calculator ? createConditionalValueReader(document,calculator,sheetIndex) : () => null;
  const worksheet = document.state?.sheets?.[sheetIndex];
  const rows = rowIndexes.map(row => columnIndexes.map(column => {
    const cell = calculator?.cell(sheetIndex, row, column) ?? document.getCell(row, column, sheetIndex);
    const validation = document.kind === 'xlsx' ? document.state?.sheets?.[sheetIndex]?.dataValidations?.find(rule => rule.ranges.some(range => row >= range.top && row <= range.bottom && column >= range.left && column <= range.right)) : null;
    const conditionalStyle = document.kind === 'xlsx' ? conditionalStyleForCell(document.state?.sheets?.[sheetIndex], row, column, cell, conditionalCellValue) : null;
    const tableStyle = document.kind === 'xlsx' ? tableStyleForCell(worksheet, row, column) : null;
    const hyperlink = document.kind === 'xlsx' ? worksheet?.hyperlinks?.find(link => row >= link.top && row <= link.bottom && column >= link.left && column <= link.right) : null;
    if (!cell && !conditionalStyle && !validation && !tableStyle && !hyperlink) return null;
    let style = document.kind === 'xlsx' ? cell?.styleProperties || null : null;
    if (conditionalStyle) {
      style = { ...(style || {}), ...conditionalStyle };
      if (style.borders && cell?.styleProperties?.borders && conditionalStyle.borders) style.borders = { ...cell.styleProperties.borders, ...conditionalStyle.borders };
    }
    const valueCell = cell || { value: '', kind: 'blank' };
    return { value: valueCell.value, displayValue: document.kind === 'xlsx' ? formatCellValue(valueCell, language) : valueCell.value, formula: valueCell.formula || null, hasFormula: isFormulaCell(valueCell), isSpill: Boolean(valueCell.isSpill), cachedValue: valueCell.cachedValue ?? null, address: valueCell.address || null, kind: valueCell.kind || null, style, tableStyle, hyperlink: hyperlink ? { display: hyperlink.display, tooltip: hyperlink.tooltip } : null, validation: validation ? { type: validation.type, values: validation.type === 'list' ? dataValidationValues(document, sheetIndex, validation, calculator, row, column) : [], operator: validation.operator, limits: validation.type === 'list' ? null : validation.values, allowBlank: validation.allowBlank, showDropDown: validation.showDropDown, showInputMessage: validation.showInputMessage, showErrorMessage: validation.showErrorMessage, errorTitle: validation.errorTitle, error: validation.error, promptTitle: validation.promptTitle, prompt: validation.prompt } : null };
  }));
  const mergedCells = document.kind === 'xlsx' ? document.state?.sheets?.[sheetIndex]?.mergedCells?.filter(range => viewRows.some(row => row >= range.top && row <= range.bottom) && columnIndexes.some(column => column >= range.left && column <= range.right)) || [] : [];
  return { type: 'range', requestId, rowStart, viewRows, viewIndexes, rowIndexes, columnStart, columnIndexes, rowCount, columnCount, sheetIndex, mergedCells, rows };
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
  const calculator = document.kind === 'xlsx' ? createFormulaCalculator(document) : null;
  let count = 0, numericCount = 0, sum = 0;
  for (let row = top; row <= bottom; row++) {
    for (let column = left; column <= right; column++) {
      const cell = document.getCell(row, column, sheetIndex);
      if (!cell) continue;
      const value = isFormulaCell(cell) ? calculator.value(sheetIndex, row, column) : cell.value;
      if (value === '' || value === null || value === undefined) continue;
      count++;
      const raw = String(value).trim();
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

function isFormulaCell(cell) { return Boolean(cell?.hasFormula || cell?.formula); }
function isFormulaCopyable(cell) { return Boolean(cell?.formula) && (!cell.formulaType || cell.formulaType === 'shared'); }

function currentRegion(document, row, column, sheetIndex = 0, includeSheet = false) {
  if (![row, column, sheetIndex].every(Number.isInteger) || row < 0 || column < 0 || sheetIndex < 0) throw new RangeError('선택 영역 요청이 유효하지 않습니다.');
  if (includeSheet) {
    if (document.kind === 'xlsx') {
      if (!document.state.sheets[sheetIndex]) throw new RangeError('선택할 시트를 찾을 수 없습니다.');
      return { top: 0, bottom: XLSX_LIMITS.rows - 1, left: 0, right: XLSX_LIMITS.columns - 1 };
    }
    if (sheetIndex !== 0) throw new RangeError('CSV에는 첫 번째 시트만 있습니다.');
    const rows = document.state.rows || [];
    let columnCount = 0;
    for (const cells of rows) columnCount = Math.max(columnCount, cells.length);
    return { top: 0, bottom: Math.max(0, rows.length - 1), left: 0, right: Math.max(0, columnCount - 1) };
  }
  const rowCells = new Map(), columnCells = new Map();
  const addCell = (cell, cellRow, cellColumn) => {
    if ((cell?.value ?? '') === '' && !isFormulaCell(cell)) return;
    if (!rowCells.has(cellRow)) rowCells.set(cellRow, new Set());
    if (!columnCells.has(cellColumn)) columnCells.set(cellColumn, new Set());
    rowCells.get(cellRow).add(cellColumn);
    columnCells.get(cellColumn).add(cellRow);
  };
  if (document.kind === 'csv') {
    if (sheetIndex !== 0) throw new RangeError('CSV에는 첫 번째 시트만 있습니다.');
    (document.state.rows || []).forEach((cells, cellRow) => cells.forEach((cell, cellColumn) => {
      const draft = document.drafts?.get(`0:${cellRow}:${cellColumn}`);
      addCell({ ...cell, value: draft ?? cell.value }, cellRow, cellColumn);
    }));
    for (const [key, value] of document.drafts || []) {
      const [draftSheet, draftRow, draftColumn] = key.split(':').map(Number);
      if (draftSheet === 0 && value !== '' && !document.state.rows[draftRow]?.[draftColumn]) addCell({ value }, draftRow, draftColumn);
    }
  } else {
    const sheet = document.state.sheets[sheetIndex];
    if (!sheet) throw new RangeError('선택할 시트를 찾을 수 없습니다.');
    for (const cell of sheet.cells.values()) {
      addCell(document.getCell(cell.row, cell.column, sheetIndex), cell.row, cell.column);
    }
    for (const [key, value] of document.drafts || []) {
      const [draftSheet, address] = key.split(':');
      if (Number(draftSheet) !== sheetIndex || value === '') continue;
      const coordinate = parseAddress(address);
      addCell({ value }, coordinate.row, coordinate.column);
    }
    for (const [key, formula] of document.formulaDrafts || []) {
      const [draftSheet, address] = key.split(':');
      if (Number(draftSheet) !== sheetIndex || formula === null) continue;
      const coordinate = parseAddress(address);
      addCell(document.getCell(coordinate.row, coordinate.column, sheetIndex), coordinate.row, coordinate.column);
    }
  }
  if (!rowCells.size) return { top: 0, bottom: 0, left: 0, right: 0 };
  const used = { top: Infinity, bottom: -1, left: Infinity, right: -1 };
  for (const cellRow of rowCells.keys()) { used.top = Math.min(used.top, cellRow); used.bottom = Math.max(used.bottom, cellRow); }
  for (const cellColumn of columnCells.keys()) { used.left = Math.min(used.left, cellColumn); used.right = Math.max(used.right, cellColumn); }
  let top = row, bottom = row, left = column, right = column, changed = true;
  const rowIntersects = (candidate, first, last) => [...(rowCells.get(candidate) || [])].some(value => value >= first && value <= last);
  const columnIntersects = (candidate, first, last) => [...(columnCells.get(candidate) || [])].some(value => value >= first && value <= last);
  while (changed) {
    changed = false;
    if (left > 0 && columnIntersects(left - 1, top, bottom)) { left--; changed = true; }
    if (right < XLSX_LIMITS.columns - 1 && columnIntersects(right + 1, top, bottom)) { right++; changed = true; }
    if (top > 0 && rowIntersects(top - 1, left, right)) { top--; changed = true; }
    if (bottom < XLSX_LIMITS.rows - 1 && rowIntersects(bottom + 1, left, right)) { bottom++; changed = true; }
  }
  if (top === row && bottom === row && left === column && right === column && !rowCells.get(row)?.has(column)) return { top: row, bottom: row, left: column, right: column };
  return { top, bottom, left, right };
}

function navigateCell(document, message) {
  const { row, column, direction, sheetIndex = 0, requestId } = message || {};
  if (![row, column, sheetIndex, requestId].every(Number.isInteger) || row < 0 || column < 0 || sheetIndex < 0 || !['up', 'down', 'left', 'right'].includes(direction)) throw new RangeError('키보드 이동 요청이 유효하지 않습니다.');
  if (document.kind === 'xlsx' && !document.state.sheets[sheetIndex]) throw new RangeError('이동할 시트를 찾을 수 없습니다.');
  if (document.kind === 'csv' && sheetIndex !== 0) throw new RangeError('CSV에는 첫 번째 시트만 있습니다.');
  if (document.kind === 'xlsx' && (row >= XLSX_LIMITS.rows || column >= XLSX_LIMITS.columns)) throw new RangeError('XLSX 행·열 제한을 벗어난 이동 요청입니다.');
  if (document.kind === 'csv' && (row >= CSV_LIMIT_ROWS || column >= CSV_LIMIT_COLUMNS)) throw new RangeError('CSV 행·열 제한을 벗어난 이동 요청입니다.');
  const vertical = direction === 'up' || direction === 'down';
  const step = direction === 'up' || direction === 'left' ? -1 : 1;
  const current = vertical ? row : column;
  const csvUsedRows = document.kind === 'csv' ? Math.max(0, document.state.rows.length - 1) : 0;
  const csvUsedColumns = document.kind === 'csv' ? Math.max(0, ...(document.state.rows[row]?.length ? [document.state.rows[row].length - 1] : [0])) : 0;
  const edge = vertical
    ? (document.kind === 'xlsx' ? XLSX_LIMITS.rows - 1 : row <= csvUsedRows ? csvUsedRows : CSV_LIMIT_ROWS - 1)
    : (document.kind === 'xlsx' ? XLSX_LIMITS.columns - 1 : column <= csvUsedColumns ? csvUsedColumns : CSV_LIMIT_COLUMNS - 1);
  const navigationResult = target => {
    let targetRow = vertical ? target : row;
    let targetColumn = vertical ? column : target;
    const mergedCell = document.kind === 'xlsx'
      ? document.state.sheets[sheetIndex].mergedCells?.find(range => targetRow >= range.top && targetRow <= range.bottom && targetColumn >= range.left && targetColumn <= range.right)
      : null;
    if (mergedCell) { targetRow = mergedCell.top; targetColumn = mergedCell.left; }
    return { type: 'navigation', requestId, row: targetRow, column: targetColumn, extend: Boolean(message.extend) };
  };
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
      if (isFormulaCell(currentCell) || currentCell.value !== '') occupied.add(vertical ? cell.row : cell.column);
    }
    for (const [key, value] of document.drafts) {
      const [draftSheet, address] = key.split(':');
      if (Number(draftSheet) !== sheetIndex) continue;
      const coordinate = parseAddress(address);
      if ((vertical ? coordinate.column : coordinate.row) !== (vertical ? column : row)) continue;
      const position = vertical ? coordinate.row : coordinate.column;
      if (value === '') occupied.delete(position); else occupied.add(position);
    }
    for (const [key, formula] of document.formulaDrafts || []) {
      const [draftSheet, address] = key.split(':');
      if (Number(draftSheet) !== sheetIndex) continue;
      const coordinate = parseAddress(address);
      if ((vertical ? coordinate.column : coordinate.row) !== (vertical ? column : row)) continue;
      const position = vertical ? coordinate.row : coordinate.column;
      if (formula === null) {
        const currentCell = document.getCell(coordinate.row, coordinate.column, sheetIndex);
        if (currentCell.value === '') occupied.delete(position);
      } else occupied.add(position);
    }
  }
  if (vertical && Array.isArray(message.hiddenRows)) {
    const rows = message.hiddenRows;
    if (rows.length > XLSX_LIMITS.rows || rows.some((value, index) => !Number.isInteger(value) || value < 0 || value > edge || (index && value <= rows[index - 1]))) throw new RangeError('숨김 행 인덱스가 유효하지 않습니다.');
    const hidden = new Set(rows);
    if (hidden.has(row)) throw new RangeError('현재 셀이 표시 행에 없습니다.');
    const nextVisible = (position, direction) => { let next = position + direction; while (next >= 0 && next <= edge && hidden.has(next)) next += direction; return next >= 0 && next <= edge ? next : null; };
    const visibleOccupied = new Set([...occupied].filter(position => !hidden.has(position)));
    const positions = [...visibleOccupied].filter(position => position * step > row * step).sort((a, b) => a * step - b * step);
    let target = row;
    if (visibleOccupied.has(row)) {
      let next = nextVisible(target, step);
      while (next !== null && visibleOccupied.has(next)) { target = next; next = nextVisible(target, step); }
      if (target === row) {
        const nextBlock = positions.find(position => position * step > target * step);
        if (nextBlock !== undefined) { target = nextBlock; let following = nextVisible(target, step); while (following !== null && visibleOccupied.has(following)) { target = following; following = nextVisible(target, step); } }
        else { let last = step > 0 ? edge : 0; while (hidden.has(last) && last >= 0 && last <= edge) last -= step; target = last; }
      }
    } else if (positions.length) { target = positions[0]; let following = nextVisible(target, step); while (following !== null && visibleOccupied.has(following)) { target = following; following = nextVisible(target, step); } }
    else { let last = step > 0 ? edge : 0; while (hidden.has(last) && last >= 0 && last <= edge) last -= step; target = last; }
    return navigationResult(target);
  }
  if (!vertical && Array.isArray(message.hiddenColumns)) {
    const columns = message.hiddenColumns;
    if (columns.length > XLSX_LIMITS.columns || columns.some((value, index) => !Number.isInteger(value) || value < 0 || value > edge || (index && value <= columns[index - 1]))) throw new RangeError('숨김 열 인덱스가 유효하지 않습니다.');
    const hidden = new Set(columns);
    if (hidden.has(column)) throw new RangeError('현재 셀이 표시 열에 없습니다.');
    const nextVisible = (position, direction) => { let next = position + direction; while (next >= 0 && next <= edge && hidden.has(next)) next += direction; return next >= 0 && next <= edge ? next : null; };
    const visibleOccupied = new Set([...occupied].filter(position => !hidden.has(position)));
    const positions = [...visibleOccupied].filter(position => position * step > column * step).sort((a, b) => a * step - b * step);
    let target = column;
    if (visibleOccupied.has(column)) {
      let next = nextVisible(target, step);
      while (next !== null && visibleOccupied.has(next)) { target = next; next = nextVisible(target, step); }
      if (target === column) {
        const nextBlock = positions.find(position => position * step > target * step);
        if (nextBlock !== undefined) { target = nextBlock; let following = nextVisible(target, step); while (following !== null && visibleOccupied.has(following)) { target = following; following = nextVisible(target, step); } }
        else { let last = step > 0 ? edge : 0; while (hidden.has(last) && last >= 0 && last <= edge) last -= step; target = last; }
      }
    } else if (positions.length) { target = positions[0]; let following = nextVisible(target, step); while (following !== null && visibleOccupied.has(following)) { target = following; following = nextVisible(target, step); } }
    else { let last = step > 0 ? edge : 0; while (hidden.has(last) && last >= 0 && last <= edge) last -= step; target = last; }
    return navigationResult(target);
  }
  if (vertical && Array.isArray(message.visibleRows)) {
    const rows = message.visibleRows;
    if (!rows.length || rows.length > XLSX_LIMITS.rows || rows.some((value, index) => !Number.isInteger(value) || value < 0 || value > edge || (index && value <= rows[index - 1]))) throw new RangeError('표시 행 인덱스가 유효하지 않습니다.');
    const currentIndex = rows.indexOf(row);
    if (currentIndex < 0) throw new RangeError('현재 셀이 표시 행에 없습니다.');
    const visibleOccupied = new Set();
    rows.forEach((value, index) => { if (occupied.has(value)) visibleOccupied.add(index); });
    const signIndex = position => position * step;
    const later = [...visibleOccupied].filter(position => signIndex(position) > signIndex(currentIndex)).sort((a, b) => signIndex(a) - signIndex(b));
    let targetIndex = currentIndex;
    if (visibleOccupied.has(currentIndex)) {
      while (targetIndex + step >= 0 && targetIndex + step < rows.length && visibleOccupied.has(targetIndex + step)) targetIndex += step;
      if (targetIndex === currentIndex) {
        const nextBlock = later.find(position => signIndex(position) > signIndex(targetIndex));
        if (nextBlock !== undefined) { targetIndex = nextBlock; while (targetIndex + step >= 0 && targetIndex + step < rows.length && visibleOccupied.has(targetIndex + step)) targetIndex += step; }
        else targetIndex = step > 0 ? rows.length - 1 : 0;
      }
    } else if (later.length) { targetIndex = later[0]; while (targetIndex + step >= 0 && targetIndex + step < rows.length && visibleOccupied.has(targetIndex + step)) targetIndex += step; }
    else targetIndex = step > 0 ? rows.length - 1 : 0;
    return navigationResult(rows[targetIndex]);
  }
  const sign = (position) => position * step;
  const positions = [...occupied].filter(position => sign(position) > sign(current)).sort((a, b) => sign(a) - sign(b));
  let target = current;
  if (occupied.has(current)) {
    while (target + step >= 0 && target + step <= edge && occupied.has(target + step)) target += step;
    if (target === current) {
      const nextBlock = positions.find(position => sign(position) > sign(target));
      if (nextBlock !== undefined) {
        target = nextBlock;
        while (target + step >= 0 && target + step <= edge && occupied.has(target + step)) target += step;
      } else target = step > 0 ? edge : 0;
    }
  } else if (positions.length) {
    target = positions[0];
    while (target + step >= 0 && target + step <= edge && occupied.has(target + step)) target += step;
  } else target = step > 0 ? edge : 0;
  return navigationResult(target);
}

function parseAddress(address) {
  const match = /^([A-Z]+)([1-9]\d*)$/.exec(address);
  if (!match) throw new RangeError('셀 주소를 키보드 이동에 사용할 수 없습니다.');
  let column = 0;
  for (const letter of match[1]) column = column * 26 + letter.charCodeAt(0) - 64;
  return { row: Number(match[2]) - 1, column: column - 1 };
}

function selectionBounds(range, operation = '복사', maxCells = 10000) {
  const { startRow, endRow, startColumn, endColumn } = range || {};
  if (![startRow, endRow, startColumn, endColumn].every(Number.isInteger) || Math.min(startRow, endRow, startColumn, endColumn) < 0) throw new RangeError(operation + ' 범위가 유효하지 않습니다.');
  const bounds = { top: Math.min(startRow, endRow), bottom: Math.max(startRow, endRow), left: Math.min(startColumn, endColumn), right: Math.max(startColumn, endColumn) };
  if ((bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1) > maxCells) {
    const action = operation === '복사' ? '복사할' : operation === '정렬' ? '정렬할' : '붙여넣을';
    throw new RangeError('한 번에 ' + action + ' 수 있는 셀은 ' + maxCells.toLocaleString() + '개까지입니다.');
  }
  return bounds;
}

function assertNoHyperlinkRange(document, sheetIndex, bounds, operation) {
  if (document.kind !== 'xlsx') return;
  const overlaps = document.state?.sheets?.[sheetIndex]?.hyperlinks?.some(link => bounds.top <= link.bottom && bounds.bottom >= link.top && bounds.left <= link.right && bounds.right >= link.left);
  if (overlaps) throw new Error('하이퍼링크가 포함된 범위는 ' + operation + '할 수 없습니다. 링크 참조를 보존하려면 해당 셀을 범위 작업에서 제외하세요.');
}

function serializeClipboard(document, range, sheetIndex = 0, language = 'en', rowIndexes = undefined) {
  const bounds = selectionBounds(range, '복사', Number.MAX_SAFE_INTEGER), rows = [];
  const copyRows = rowIndexes ?? Array.from({ length: bounds.bottom - bounds.top + 1 }, (_, index) => bounds.top + index);
  for (const row of copyRows) {
    const values = [];
    for (let column = bounds.left; column <= bounds.right; column++) {
      const cell = document.getCell(row, column, sheetIndex);
      const value = document.kind === 'xlsx'
        ? (cell?.formula ? '=' + cell.formula : formatCellValue(cell, language))
        : (cell?.value ?? '');
      values.push(quoteCsv(value, '\t', false));
    }
    rows.push(values.join('\t'));
  }
  return rows.join('\r\n');
}

function buildCopyPayload(document, range, sheetIndex = 0, language = 'en', visibleRows = undefined, hiddenRows = undefined) {
  const bounds = selectionBounds(range, '복사', Number.MAX_SAFE_INTEGER), maxRows = document.kind === 'xlsx' ? XLSX_LIMITS.rows : document.kind === 'csv' ? document.state?.rows?.length ?? Number.MAX_SAFE_INTEGER : 0;
  if (bounds.bottom >= maxRows) throw new RangeError('복사 범위가 시트 경계를 벗어났습니다.');
  if (visibleRows !== undefined && hiddenRows !== undefined) throw new RangeError('표시 행 정보가 중복되었습니다.');
  if (visibleRows !== undefined && (!Array.isArray(visibleRows) || visibleRows.length > maxRows || visibleRows.some((row, index) => !Number.isInteger(row) || row < 0 || row >= maxRows || (index && row <= visibleRows[index - 1])))) throw new RangeError('표시 행 인덱스가 유효하지 않습니다.');
  if (hiddenRows !== undefined && (!Array.isArray(hiddenRows) || hiddenRows.length > maxRows || hiddenRows.some((row, index) => !Number.isInteger(row) || row < 0 || row >= maxRows || (index && row <= hiddenRows[index - 1])))) throw new RangeError('숨김 행 인덱스가 유효하지 않습니다.');
  const hiddenSelection = hiddenRows?.filter(row => row >= bounds.top && row <= bounds.bottom) || [];
  const selectedRowCount = visibleRows === undefined ? bounds.bottom - bounds.top + 1 - hiddenSelection.length : visibleRows.filter(row => row >= bounds.top && row <= bounds.bottom).length;
  const sparseFullColumn = visibleRows === undefined && hiddenRows === undefined && document.kind === 'xlsx' && range?.startColumn === range?.endColumn && Math.abs(range.endRow - range.startRow) + 1 === XLSX_LIMITS.rows;
  const copyColumns = bounds.right - bounds.left + 1, maxCells = document.kind === 'xlsx' ? 20000 : 10000;
  if (selectedRowCount * copyColumns > maxCells && !sparseFullColumn) throw new RangeError('한 번에 복사할 수 있는 셀은 ' + maxCells.toLocaleString() + '개까지입니다.');
  const hiddenSet = new Set(hiddenSelection);
  const copyRows = [];
  if (!sparseFullColumn) {
    if (visibleRows !== undefined) copyRows.push(...visibleRows.filter(row => row >= bounds.top && row <= bounds.bottom));
    else for (let row = bounds.top; row <= bounds.bottom; row++) if (!hiddenSet.has(row)) copyRows.push(row);
  }
  if (sparseFullColumn) {
    const sparseCells = new Map(), sheet = document.state.sheets[sheetIndex];
    const addCell = (row, column) => {
      if (row < bounds.top || row > bounds.bottom || column !== bounds.left) return;
      const cell = document.getCell(row, column, sheetIndex);
      if (!cell || ((cell.value ?? '') === '' && !isFormulaCell(cell) && cell.kind !== 'error' && cell.style == null)) return;
      sparseCells.set(row - bounds.top, { rowOffset: row - bounds.top, columnOffset: 0, value: cell.value ?? '', kind: cell.kind ?? 'blank', formula: isFormulaCopyable(cell) ? cell.formula : null, styleIndex: cell.style ?? null, styleProperties: cell.styleProperties ?? null });
    };
    for (const cell of sheet?.cells.values() || []) if (cell.column === bounds.left) addCell(cell.row, bounds.left);
    for (const key of [...document.drafts.keys(), ...document.formulaDrafts.keys(), ...document.styleDrafts.keys()]) {
      const [draftSheet, address] = key.split(':');
      if (Number(draftSheet) === sheetIndex) { const coordinate = parseAddress(address); if (coordinate.column === bounds.left) addCell(coordinate.row, coordinate.column); }
    }
    const cells = [...sparseCells.values()].sort((a, b) => a.rowOffset - b.rowOffset);
    return { kind: 'xlsx', sparse: true, height: XLSX_LIMITS.rows, width: 1, originRow: bounds.top, originColumn: bounds.left, cells, text: null };
  }
  const cells = [];
  for (const row of copyRows) {
    const line = [];
    for (let column = bounds.left; column <= bounds.right; column++) {
      const cell = document.getCell(row, column, sheetIndex);
      line.push({ value: cell?.value ?? '', kind: cell?.kind ?? 'blank', formula: isFormulaCopyable(cell) ? cell.formula : null, row, column, styleIndex: cell?.style ?? null, styleProperties: cell?.styleProperties ?? null });
    }
    cells.push(line);
  }
  return { kind: document.kind, text: serializeClipboard(document, range, sheetIndex, language, copyRows), cells };
}

function buildSparseColumnPasteEdits(document, copied, row, column, range, sheetIndex = 0) {
  if (document.kind !== 'xlsx' || !Number.isInteger(row) || row < 0 || !Number.isInteger(column) || column < 0 || copied.width !== 1 || copied.height !== XLSX_LIMITS.rows) throw new RangeError('전체 열 붙여넣기 대상이 유효하지 않습니다.');
  if (row !== 0) throw new RangeError('전체 열은 첫 번째 행에 붙여넣어야 합니다.');
  const targetColumn = column;
  if (targetColumn >= XLSX_LIMITS.columns) throw new RangeError('붙여넣기 대상이 시트 경계를 벗어났습니다.');
  if (range) {
    const target = selectionBounds(range, '붙여넣기', XLSX_LIMITS.rows);
    if (target.top !== 0 || target.bottom !== XLSX_LIMITS.rows - 1 || target.left !== targetColumn || target.right !== targetColumn) throw new RangeError('전체 열 붙여넣기는 대상 열 전체를 선택해야 합니다.');
  }
  const candidates = new Set(copied.cells.map(cell => cell.rowOffset));
  const sheet = document.state.sheets[sheetIndex];
  for (const cell of sheet?.cells.values() || []) if (cell.column === targetColumn) candidates.add(cell.row);
  for (const key of [...document.drafts.keys(), ...document.formulaDrafts.keys(), ...document.styleDrafts.keys()]) {
    const [draftSheet, address] = key.split(':');
    if (Number(draftSheet) === sheetIndex) { const coordinate = parseAddress(address); if (coordinate.column === targetColumn) candidates.add(coordinate.row); }
  }
  if (candidates.size > XLSX_LIMITS.cells) throw new RangeError('한 번에 붙여넣을 수 있는 셀은 ' + XLSX_LIMITS.cells.toLocaleString() + '개까지입니다.');
  const sourceByRow = new Map(copied.cells.map(cell => [cell.rowOffset, cell]));
  return [...candidates].sort((a, b) => a - b).map(targetRow => {
    const source = sourceByRow.get(targetRow), hasFormula = Boolean(source?.formula), edit = { row: targetRow, column: targetColumn, value: hasFormula ? '' : source?.value ?? '', kind: source?.kind ?? 'blank', formula: hasFormula ? translateFormulaReferences(source.formula, targetRow - (copied.originRow + source.rowOffset), targetColumn - copied.originColumn) : null, styleIndex: source?.styleIndex ?? null, styleProperties: source?.styleProperties ?? null };
    return edit;
  });
}

function buildSparseColumnMoveEdits(document, sourceRange, targetRow, targetColumn, sheetIndex = 0) {
  const source = selectionBounds(sourceRange, '붙여넣기', XLSX_LIMITS.rows);
  if (document.kind !== 'xlsx' || source.top !== 0 || source.bottom !== XLSX_LIMITS.rows - 1 || source.left !== source.right || targetRow !== 0 || !Number.isInteger(targetColumn) || targetColumn < 0 || targetColumn >= XLSX_LIMITS.columns) throw new RangeError('전체 열 이동 대상이 유효하지 않습니다.');
  if (source.left === targetColumn) return [];
  if (document.state?.sheets?.[sheetIndex]?.hyperlinks?.some(link => link.left <= source.left && link.right >= source.left || link.left <= targetColumn && link.right >= targetColumn)) throw new Error('하이퍼링크가 있는 열은 이동할 수 없습니다.');
  const rows = new Set(), sheet = document.state.sheets[sheetIndex];
  for (const cell of sheet?.cells.values() || []) if (cell.column === source.left || cell.column === targetColumn) rows.add(cell.row);
  for (const key of [...document.drafts.keys(), ...document.formulaDrafts.keys(), ...document.styleDrafts.keys()]) {
    const [draftSheet, address] = key.split(':');
    if (Number(draftSheet) === sheetIndex) { const coordinate = parseAddress(address); if (coordinate.column === source.left || coordinate.column === targetColumn) rows.add(coordinate.row); }
  }
  if (rows.size > XLSX_LIMITS.cells) throw new RangeError('한 번에 이동할 수 있는 셀은 ' + XLSX_LIMITS.cells.toLocaleString() + '개까지입니다.');
  const edits = [];
  for (const row of [...rows].sort((a, b) => a - b)) {
    const sourceCell = document.getCell(row, source.left, sheetIndex), targetCell = document.getCell(row, targetColumn, sheetIndex);
    if ((sourceCell?.kind === 'error' && !isFormulaCell(sourceCell)) || (isFormulaCell(sourceCell) && (!sourceCell.formula || sourceCell.formulaType)) || targetCell?.formulaType) throw new Error('공유·배열 수식 또는 일반 오류 셀이 포함된 범위는 이동할 수 없습니다.');
    const hasFormula = isFormulaCell(sourceCell), formula = hasFormula ? translateFormulaReferences(sourceCell.formula, 0, targetColumn - source.left) : null;
    edits.push({ row, column: targetColumn, value: hasFormula ? '' : sourceCell?.value ?? '', kind: sourceCell?.kind ?? 'blank', formula, styleIndex: sourceCell?.style ?? null, styleProperties: sourceCell?.styleProperties ?? null });
    edits.push({ row, column: source.left, value: '', formula: null, styleIndex: null, styleProperties: null, kind: 'blank' });
  }
  if (edits.length > XLSX_LIMITS.cells * 2) throw new RangeError('한 번에 이동할 수 있는 셀 수정은 ' + (XLSX_LIMITS.cells * 2).toLocaleString() + '개까지입니다.');
  return edits;
}

function parseClipboard(text) {
  if (typeof text !== 'string' || text.length > MAX_CLIPBOARD_CHARS) throw new RangeError('붙여넣기 내용은 1,000,000자까지 허용합니다.');
  if (!text) return [];
  const values = /[\t\r\n]/.test(text)
    ? parseCsv(text, '\t', 20000).map(row => row.map(cell => cell.value))
    : [[text]];
  if (values.some(row => row.some(value => value.length > MAX_CELL_CHARS))) throw new RangeError('붙여넣기 셀은 100,000자까지 허용합니다.');
  return values;
}

function inferClipboardCellKind(value) {
  if (typeof value !== 'string') return 'blank';
  if (value === '') return 'blank';
  if (/^\s*[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?\s*$/.test(value) && Number.isFinite(Number(value.trim()))) return 'number';
  if (/^(?:TRUE|FALSE)$/i.test(value)) return 'boolean';
  return 'string';
}

function parseClipboardDate(value, locale = 'en-US', date1904 = false) {
  if (typeof value !== 'string') return null;
  let year, month, day, time = null;
  let match = /^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?)?$/.exec(value);
  if (match) {
    [, year, month, day] = match;
    if (match[4] !== undefined) time = [Number(match[4]), Number(match[5]), Number(match[6] || 0), Number((match[7] || '').padEnd(3, '0') || 0)];
  } else {
    match = /^(\d{4})[./]\s*(\d{1,2})[./]\s*(\d{1,2})\.?(?:[ T](\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?)?$/.exec(value);
    if (match) {
      [, year, month, day] = match;
      if (match[4] !== undefined) time = [Number(match[4]), Number(match[5]), Number(match[6] || 0), Number((match[7] || '').padEnd(3, '0') || 0)];
    } else if ((match = /^(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(value))) {
      [, year, month, day] = match;
      if (match[4] !== undefined) time = [Number(match[4]), Number(match[5]), Number(match[6] || 0), 0];
    } else {
      match = /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(value);
      if (!match) return null;
      const [, first, second, rawYear] = match;
      if (clipboardLocaleUsesDayFirstDates(locale)) { day = first; month = second; }
      else { month = first; day = second; }
      year = rawYear;
      if (match[4] !== undefined) time = [Number(match[4]), Number(match[5]), Number(match[6] || 0), 0];
    }
  }
  year = Number(year); month = Number(month); day = Number(day);
  if (year < 1900 || year > 9999) return null;
  const date = new Date(0);
  date.setUTCFullYear(year, month - 1, day);
  date.setUTCHours(0, 0, 0, 0);
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return null;
  const [hour = 0, minute = 0, second = 0, millisecond = 0] = time || [];
  if (hour > 23 || minute > 59 || second > 59) return null;
  date.setUTCHours(hour, minute, second, millisecond);
  const utcDays = (date.getTime() - Date.UTC(1899, 11, 31)) / 86400000;
  const serial = utcDays + (date.getTime() >= Date.UTC(1900, 2, 1) ? 1 : 0) - (date1904 ? 1462 : 0);
  if (serial < 0) return null;
  return { serial, numFmtId: time ? 22 : 14 };
}

function clipboardLocaleUsesDayFirstDates(locale) {
  try {
    const parts = new Intl.DateTimeFormat(locale, { day: 'numeric', month: 'numeric', year: 'numeric', timeZone: 'UTC' }).formatToParts(new Date(Date.UTC(2006, 10, 22)));
    return parts.findIndex(part => part.type === 'day') < parts.findIndex(part => part.type === 'month');
  } catch { return false; }
}

function buildDateTimeShortcutEdit(document, row, column, mode, now = new Date()) {
  if (!document || !['date', 'time'].includes(mode) || !Number.isInteger(row) || row < 0 || !Number.isInteger(column) || column < 0 || !(now instanceof Date) || !Number.isFinite(now.getTime())) throw new RangeError('날짜·시간 입력 요청이 유효하지 않습니다.');
  const two = value => String(value).padStart(2, '0');
  const isoDate = [String(now.getFullYear()).padStart(4, '0'), two(now.getMonth() + 1), two(now.getDate())].join('-');
  const value = mode === 'date' ? isoDate : two(now.getHours()) + ':' + two(now.getMinutes());
  if (document.kind !== 'xlsx') return { row, column, value };
  if (!document.state.stylesPath) throw new Error('XLSX 통합문서에 날짜·시간 표시 서식을 저장할 스타일 정의가 없습니다.');
  if (mode === 'date') {
    const parsed = parseClipboardDate(isoDate, 'en', document.state.date1904);
    if (!parsed) throw new Error('현재 날짜를 Excel 날짜 값으로 변환할 수 없습니다.');
    return { row, column, value: String(parsed.serial), kind: 'number', numberFormatId: 14 };
  }
  const serial = (now.getHours() * 60 + now.getMinutes()) / 1440;
  return { row, column, value: String(serial), kind: 'number', numberFormatId: 20 };
}

function parseClipboardTime(value) {
  if (typeof value !== 'string') return null;
  const match = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM|오전|오후)?$/i.exec(value.trim());
  if (!match) return null;
  let hour = Number(match[1]);
  const minute = Number(match[2]), second = Number(match[3] || 0), period = match[4]?.toUpperCase();
  if (minute > 59 || second > 59) return null;
  if (period) {
    if (hour < 1 || hour > 12) return null;
    if (period === 'AM' || period === '오전') hour %= 12;
    else hour = hour % 12 + 12;
  } else if (hour > 23) return null;
  return { serial: (hour * 3600 + minute * 60 + second) / 86400, numFmtId: match[3] ? 21 : 20 };
}

function parseClipboardPercentage(value) {
  if (typeof value !== 'string') return null;
  const match = /^([+-]?(?:(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+))%$/.exec(value.trim());
  if (!match) return null;
  const percent = Number(match[1].replace(/,/g, ''));
  if (!Number.isFinite(percent)) return null;
  const decimals = (match[1].split('.')[1] || '').length;
  return { serial: percent / 100, numFmtId: decimals ? 10 : 9 };
}

function parseClipboardCurrency(value, language, model) {
  if (typeof value !== 'string' || !model?.stylesPath || !model.files?.[model.stylesPath]) return null;
  let text = value.trim(), currency = null;
  const currencies = { '$': 'USD', USD: 'USD', '€': 'EUR', EUR: 'EUR', '£': 'GBP', GBP: 'GBP', '¥': 'JPY', JPY: 'JPY', '₩': 'KRW', KRW: 'KRW', 원: 'KRW' };
  const prefix = /^([+-]?)\s*(USD|EUR|GBP|JPY|KRW|[$€£¥₩])\s*/i.exec(text);
  if (prefix) { currency = currencies[prefix[2].toUpperCase()] || currencies[prefix[2]]; text = prefix[1] + text.slice(prefix[0].length); }
  else {
    const suffix = /\s*(원|USD|EUR|GBP|JPY|KRW)$/i.exec(text);
    if (suffix) { currency = currencies[suffix[1].toUpperCase()] || currencies[suffix[1]]; text = text.slice(0, suffix.index); }
  }
  if (!currency || !/^[+-]?(?:\d[\d.,]*|[.,]\d+)$/.test(text)) return null;
  let sign = '';
  if (/^[+-]/.test(text)) { sign = text[0]; text = text.slice(1); }
  const dot = text.lastIndexOf('.'), comma = text.lastIndexOf(',');
  let decimalSeparator = null;
  if (dot >= 0 && comma >= 0) decimalSeparator = dot > comma ? '.' : ',';
  else if (dot >= 0) decimalSeparator = '.';
  else if (comma >= 0 && (currency === 'EUR' || (language === 'ko' && currency !== 'KRW')) && text.length - comma - 1 <= 2) decimalSeparator = ',';
  const groupingSeparator = decimalSeparator === '.' ? ',' : decimalSeparator === ',' ? '.' : (text.includes(',') ? ',' : '.');
  let integerPart = text, fractionalPart = '';
  if (decimalSeparator) {
    const split = text.lastIndexOf(decimalSeparator);
    integerPart = text.slice(0, split); fractionalPart = text.slice(split + 1);
    if (!fractionalPart || !/^\d{1,4}$/.test(fractionalPart)) return null;
  }
  if (integerPart.includes(decimalSeparator) || (integerPart.includes(groupingSeparator)
    ? integerPart.split(groupingSeparator).some((part, index) => index === 0 ? !/^\d{1,3}$/.test(part) : !/^\d{3}$/.test(part))
    : !/^\d+$/.test(integerPart))) return null;
  const normalized = sign + integerPart.split(groupingSeparator).join('') + (decimalSeparator ? '.' + fractionalPart : '');
  const amount = Number(normalized);
  if (!Number.isFinite(amount)) return null;
  const symbol = { USD: '$', EUR: '€', GBP: '£', JPY: '¥', KRW: '₩' }[currency];
  const decimals = fractionalPart.length;
  const formatCode = '"' + symbol + '"#,##0' + (decimals ? '.' + '0'.repeat(decimals) : '');
  const stylesXml = new TextDecoder().decode(model.files[model.stylesPath]);
  const existing = [...stylesXml.matchAll(/<numFmt\b[^>]*\bnumFmtId="(\d+)"[^>]*\bformatCode="([^"]*)"[^>]*\/>/g)]
    .find(match => match[2].replace(/&quot;/g, '"').replace(/&amp;/g, '&') === formatCode);
  let numFmtId = existing ? Number(existing[1]) : 164;
  if (!existing) {
    const ids = new Set([...stylesXml.matchAll(/\bnumFmtId="(\d+)"/g)].map(match => Number(match[1])));
    while (ids.has(numFmtId) && numFmtId < 65536) numFmtId++;
    if (numFmtId >= 65536) return null;
  }
  return { serial: amount, numFmtId, formatCode };
}

function buildClipboardKindMatrix(copied, targetKind, matrix) {
  if (targetKind !== 'xlsx') return undefined;
  if (copied?.kind === 'xlsx') return copied.cells.map(line => line.map(cell => cell.kind ?? inferClipboardCellKind(cell.value)));
  return matrix.map(line => line.map(inferClipboardCellKind));
}

function buildClipboardFormulaMatrix(copied, targetKind, targetRow, targetColumn, matrix) {
  if (targetKind !== 'xlsx') return undefined;
  if (copied?.kind === 'xlsx') return copied.cells.map(line => line.map(cell => cell.formula ? { formula: cell.formula, row: cell.row, column: cell.column } : null));
  return matrix.map((line, rowOffset) => line.map((cell, columnOffset) => typeof cell === 'string' && cell.startsWith('=') && cell.length > 1
    ? { formula: cell.slice(1), row: targetRow + rowOffset, column: targetColumn + columnOffset }
    : null));
}

function isSameWorkbookClipboard(copied, document) {
  return copied?.kind === 'xlsx' && typeof copied.sourceUri === 'string' && copied.sourceUri === document?.uri?.toString();
}

function buildMoveEdits(document, sourceRange, targetRow, targetColumn, sheetIndex = 0, hiddenRows = []) {
  if (!Number.isInteger(targetRow) || targetRow < 0 || !Number.isInteger(targetColumn) || targetColumn < 0) throw new RangeError('붙여넣기 대상이 유효하지 않습니다.');
  const source = selectionBounds(sourceRange, '붙여넣기', document.kind === 'xlsx' ? 20000 : 10000), width = source.right - source.left + 1;
  assertNoHyperlinkRange(document, sheetIndex, source, '이동');
  const hidden = new Set(hiddenRows), sourceRows = [];
  for (let row = source.top; row <= source.bottom; row++) if (!hidden.has(row)) sourceRows.push(row);
  const maximumRows = document.kind === 'csv' ? CSV_LIMIT_ROWS : XLSX_LIMITS.rows, targetRows = [];
  for (let row = targetRow; targetRows.length < sourceRows.length; row++) {
    if (row >= maximumRows) throw new RangeError('붙여넣기 대상이 시트 경계를 벗어났습니다.');
    if (!hidden.has(row)) targetRows.push(row);
  }
  if (sourceRows.length * width > (document.kind === 'xlsx' ? 20000 : 10000)) throw new RangeError('붙여넣기 범위는 허용된 셀 수를 넘습니다.');
  if (!sourceRows.length) return [];
  selectionBounds({ startRow: targetRows[0], endRow: targetRows.at(-1), startColumn: targetColumn, endColumn: targetColumn + width - 1 }, '붙여넣기', document.kind === 'xlsx' ? 20000 : 10000);
  assertNoHyperlinkRange(document, sheetIndex, { top: targetRows[0], bottom: targetRows.at(-1), left: targetColumn, right: targetColumn + width - 1 }, '붙여넣기');
  const destination = [], editsByCell = new Map();
  for (let rowOffset = 0; rowOffset < sourceRows.length; rowOffset++) for (let columnOffset = 0; columnOffset < width; columnOffset++) {
    const fromRow = sourceRows[rowOffset], fromColumn = source.left + columnOffset, toRow = targetRows[rowOffset], toColumn = targetColumn + columnOffset;
    const sourceCell = document.getCell(fromRow, fromColumn, sheetIndex), targetCell = document.getCell(toRow, toColumn, sheetIndex);
    if ((sourceCell?.kind === 'error' && !isFormulaCell(sourceCell)) || (isFormulaCell(sourceCell) && (!sourceCell.formula || sourceCell.formulaType)) || (isFormulaCell(targetCell) && targetCell.formulaType)) throw new Error('공유·배열 수식 또는 일반 오류 셀이 포함된 범위는 이동할 수 없습니다.');
    const edit = { row: toRow, column: toColumn, value: sourceCell?.value ?? '' };
    if (document.kind === 'xlsx') edit.formula = isFormulaCell(sourceCell) ? sourceCell.formula : null;
    if (document.kind === 'xlsx') { edit.styleIndex = sourceCell?.style ?? null; edit.styleProperties = sourceCell?.styleProperties ?? null; edit.kind = sourceCell?.kind ?? 'blank'; }
    destination.push(edit);
  }
  for (const edit of destination) editsByCell.set(edit.row + ':' + edit.column, edit);
  const destinationCells = new Set(destination.map(edit => edit.row + ':' + edit.column));
  for (const row of sourceRows) for (let column = source.left; column <= source.right; column++) {
    if (destinationCells.has(row + ':' + column)) continue;
    editsByCell.set(row + ':' + column, { row, column, value: '', ...(document.kind === 'xlsx' ? { formula: null, styleIndex: null, styleProperties: null, kind: 'blank' } : {}) });
  }
  const maximumEdits = document.kind === 'xlsx' ? 40000 : 20000;
  if (editsByCell.size > maximumEdits) throw new RangeError('한 번에 이동할 수 있는 수정 셀은 ' + maximumEdits.toLocaleString() + '개까지입니다.');
  return [...editsByCell.values()];
}

function buildFillEdits(document, range, direction, sheetIndex = 0, hiddenRows = []) {
  if (!['down', 'right'].includes(direction)) throw new RangeError('채우기 방향이 유효하지 않습니다.');
  const bounds = selectionBounds(range, '붙여넣기', document.kind === 'xlsx' ? 20000 : 10000), edits = [];
  assertNoHyperlinkRange(document, sheetIndex, bounds, '채우기');
  const hidden = new Set(hiddenRows);
  const fillCell = (source, target, row, column, sourceRow, sourceColumn) => {
    const sourceHasFormula = isFormulaCell(source), targetHasFormula = isFormulaCell(target);
    if ((source?.kind === 'error' && !sourceHasFormula) || (sourceHasFormula && (!source.formula || source.formulaType)) || (target?.kind === 'error' && document.kind !== 'xlsx') || (targetHasFormula && (target.formulaType || document.kind !== 'xlsx'))) throw new Error('공유·배열 수식 또는 일반 오류 셀이 포함된 범위는 수정할 수 없습니다.');
    const edit = sourceHasFormula
      ? { row, column, value: '', formula: typeof source.formula === 'string' ? translateFormulaReferences(source.formula, row - sourceRow, column - sourceColumn) : null }
      : { row, column, value: source?.value ?? '', ...(document.kind === 'xlsx' ? { formula: null } : {}) };
    if (sourceHasFormula && !edit.formula) throw new Error('수식 참조를 확인할 수 없는 셀은 채울 수 없습니다.');
    if (document.kind === 'xlsx') { edit.styleIndex = source?.style ?? null; edit.styleProperties = source?.styleProperties ?? null; edit.kind = source?.kind ?? 'blank'; }
    edits.push(edit);
  };
  if (direction === 'down') {
    let sourceRow = bounds.top;
    while (sourceRow <= bounds.bottom && hidden.has(sourceRow)) sourceRow++;
    if (sourceRow > bounds.bottom || sourceRow === bounds.bottom) return edits;
    for (let column = bounds.left; column <= bounds.right; column++) {
      const source = document.getCell(sourceRow, column, sheetIndex);
      for (let row = sourceRow + 1; row <= bounds.bottom; row++) {
        if (hidden.has(row)) continue;
        const target = document.getCell(row, column, sheetIndex);
        fillCell(source, target, row, column, sourceRow, column);
      }
    }
  } else {
    if (bounds.left === bounds.right) return edits;
    for (let row = bounds.top; row <= bounds.bottom; row++) {
      if (hidden.has(row)) continue;
      const source = document.getCell(row, bounds.left, sheetIndex);
      for (let column = bounds.left + 1; column <= bounds.right; column++) {
        const target = document.getCell(row, column, sheetIndex);
        fillCell(source, target, row, column, row, bounds.left);
      }
    }
  }
  return edits;
}

function buildFillRangeEdits(document, sourceRange, targetRange, sheetIndex = 0, hiddenRows = []) {
  const source = selectionBounds(sourceRange, '채우기'), target = selectionBounds(targetRange, '채우기');
  assertNoHyperlinkRange(document, sheetIndex, source, '채우기');
  assertNoHyperlinkRange(document, sheetIndex, target, '채우기');
  if (target.top > source.top || target.left > source.left || target.bottom < source.bottom || target.right < source.right) throw new RangeError('채우기 대상은 원본 선택 범위를 포함해야 합니다.');
  const sourceHeight = source.bottom - source.top + 1, sourceWidth = source.right - source.left + 1;
  const readSeries = values => {
    if (!values.length) return null;
    const texts = values.map(value => String(value?.value ?? ''));
    const seriesLists = [
        ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'],
        ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'],
        ['월요일', '화요일', '수요일', '목요일', '금요일', '토요일', '일요일'],
        ['월', '화', '수', '목', '금', '토', '일'],
        Array.from({ length: 12 }, (_, index) => (index + 1) + '월'),
        ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
        ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
      ];
    const listSeries = text => {
      for (const list of seriesLists) {
        const index = list.findIndex(value => value.toLocaleLowerCase() === text.toLocaleLowerCase());
        if (index >= 0) return { values: list, index, input: text };
      }
      return null;
    };
    const parseDateText = text => {
      const match = /^(\d{4})([-/.])(\d{1,2})\2(\d{1,2})$/.exec(text);
      if (!match) return null;
      const year = Number(match[1]), month = Number(match[3]), day = Number(match[4]);
      const date = new Date(0);
      date.setUTCFullYear(year, month - 1, day);
      date.setUTCHours(0, 0, 0, 0);
      if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return null;
      const monthLength = [4, 6, 9, 11].includes(month) ? 30 : month === 2 ? (year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0) ? 29 : 28) : 31;
      return { year, month, day, separator: match[2], monthWidth: match[3].length, dayWidth: match[4].length, serial: Math.trunc(date.getTime() / 86400000), monthIndex: year * 12 + month - 1, endOfMonth: day === monthLength };
    };
    if (values.length === 1) {
      const date = parseDateText(texts[0]);
      if (date) return { mode: 'date', first: date.serial, last: date.serial, delta: 1, separator: date.separator, monthWidth: date.monthWidth, dayWidth: date.dayWidth };
      const value = values[0];
      const formatCode = String(value?.formatCode || '');
      if (value?.kind === 'number' && Number.isFinite(Number(texts[0])) &&
          ((Number(value.numFmtId) >= 14 && Number(value.numFmtId) <= 22) || /[yd]/i.test(formatCode))) {
        const serial = Number(texts[0]);
        return { mode: 'excel-date', first: serial, last: serial, delta: 1, decimalPlaces: 10 };
      }
      const item = listSeries(texts[0]);
      return item ? { mode: 'list', values: item.values, firstIndex: item.index, lastIndex: item.index, delta: 1, input: item.input } : null;
    }
    const numbers = values.map((value, index) => {
      const text = texts[index];
      if (value?.kind !== 'number' && !/^[+-]?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?$/.test(text)) return null;
      const number = Number(text);
      return Number.isFinite(number) ? { text, number } : null;
    });
    if (numbers.every(Boolean)) {
      const decimalPlaces = Math.min(12, Math.max(...numbers.map(({ text }) => {
        const [mantissa, exponent = '0'] = text.toLowerCase().split('e');
        return Math.max(0, (mantissa.split('.')[1] || '').length - Number(exponent));
      })));
      return { mode: 'number', first: numbers[0].number, last: numbers.at(-1).number, delta: numbers.at(-1).number - numbers.at(-2).number, decimalPlaces };
    }
    const dateParts = texts.map(parseDateText);
    if (dateParts.every(Boolean)) {
      if (!dateParts.every(part => part.separator === dateParts[0].separator && part.monthWidth === dateParts[0].monthWidth && part.dayWidth === dateParts[0].dayWidth)) return null;
      const monthDeltas = dateParts.slice(1).map((part, index) => part.monthIndex - dateParts[index].monthIndex);
      const sameMonthStep = monthDeltas[0] !== 0 && monthDeltas.every(delta => delta === monthDeltas[0]);
      const sameDay = dateParts.every(part => part.day === dateParts[0].day), endOfMonth = dateParts.every(part => part.endOfMonth);
      if (sameMonthStep && (sameDay || endOfMonth)) return { mode: 'month', first: dateParts[0].serial, last: dateParts.at(-1).serial, firstParts: dateParts[0], lastParts: dateParts.at(-1), delta: monthDeltas[0], endOfMonth };
      const serials = dateParts.map(part => part.serial);
      return { mode: 'date', first: serials[0], last: serials.at(-1), delta: serials.at(-1) - serials.at(-2), separator: dateParts[0].separator, monthWidth: dateParts[0].monthWidth, dayWidth: dateParts[0].dayWidth };
    }
    const listParts = texts.map(listSeries);
    if (listParts.every(Boolean) && listParts.every(part => part.values === listParts[0].values)) {
      const listLength = listParts[0].values.length;
      const forwardDelta = (listParts[1].index - listParts[0].index + listLength) % listLength;
      const delta = forwardDelta === 1 ? 1 : forwardDelta === listLength - 1 ? -1 : 0;
      if (delta && listParts.every((part, index) => part.index === ((listParts[0].index + delta * index) % listLength + listLength) % listLength)) {
        return { mode: 'list', values: listParts[0].values, firstIndex: listParts[0].index, lastIndex: listParts.at(-1).index, delta, input: listParts[0].input };
      }
    }
    const suffixes = texts.map(text => /^(.*\D)(\d+)$/.exec(text));
    if (suffixes.every(match => match && match[1] === suffixes[0][1])) {
      const values = suffixes.map(match => Number(match[2]));
      if (values.every(Number.isSafeInteger)) return { mode: 'suffix', prefix: suffixes[0][1], width: Math.max(...suffixes.map(match => match[2].length)), first: values[0], last: values.at(-1), delta: values.at(-1) - values.at(-2) };
    }
    return null;
  };
  const makeSeriesValue = (series, before, steps) => {
    if (series.mode === 'list') {
      const offset = (before ? series.firstIndex : series.lastIndex) + series.delta * (before ? -steps : steps);
      let value = series.values[((offset % series.values.length) + series.values.length) % series.values.length];
      if (series.input === series.input.toLocaleLowerCase()) value = value.toLocaleLowerCase();
      else if (series.input === series.input.toLocaleUpperCase()) value = value.toLocaleUpperCase();
      return value;
    }
    const value = before ? series.first - series.delta * steps : series.last + series.delta * steps;
    if (!Number.isFinite(value)) return undefined;
    if (series.mode === 'number') return String(Number(value.toFixed(series.decimalPlaces)));
    if (series.mode === 'excel-date') return String(Number(value.toFixed(series.decimalPlaces)));
    if (series.mode === 'month') {
      const base = before ? series.firstParts : series.lastParts;
      const monthIndex = base.monthIndex + series.delta * (before ? -steps : steps);
      const year = Math.floor(monthIndex / 12), month = monthIndex - year * 12 + 1;
      if (year < 1 || year > 9999) return undefined;
      const monthLength = [4, 6, 9, 11].includes(month) ? 30 : month === 2 ? (year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0) ? 29 : 28) : 31;
      const day = series.endOfMonth ? monthLength : Math.min(base.day, monthLength);
      return String(year).padStart(4, '0') + base.separator + String(month).padStart(base.monthWidth, '0') + base.separator + String(day).padStart(base.dayWidth, '0');
    }
    if (series.mode === 'date') {
      const date = new Date(value * 86400000);
      const iso = Number.isNaN(date.getTime()) ? '' : date.toISOString().slice(0, 10);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(iso)) return undefined;
      const dateParts = iso.split('-');
      return dateParts[0] + (series.separator || '-') + String(Number(dateParts[1])).padStart(series.monthWidth || 2, '0') + (series.separator || '-') + String(Number(dateParts[2])).padStart(series.dayWidth || 2, '0');
    }
    const sign = value < 0 ? '-' : '';
    return series.prefix + sign + String(Math.abs(value)).padStart(series.width, '0');
  };
  let horizontalSeries = null, verticalSeries = null;
  if (sourceHeight === 1 && (sourceWidth > 1 || target.left < source.left || target.right > source.right)) horizontalSeries = readSeries(Array.from({ length: sourceWidth }, (_, offset) => document.getCell(source.top, source.left + offset, sheetIndex)));
  if (sourceWidth === 1 && (sourceHeight > 1 || target.top < source.top || target.bottom > source.bottom)) verticalSeries = readSeries(Array.from({ length: sourceHeight }, (_, offset) => document.getCell(source.top + offset, source.left, sheetIndex)));
  const edits = [];
  const hidden = new Set(hiddenRows);
  for (let row = target.top; row <= target.bottom; row++) for (let column = target.left; column <= target.right; column++) {
    if (hidden.has(row)) continue;
    if (row >= source.top && row <= source.bottom && column >= source.left && column <= source.right) continue;
    let sourceRow = source.top + (((row - source.top) % sourceHeight) + sourceHeight) % sourceHeight;
    let sourceColumn = source.left + (((column - source.left) % sourceWidth) + sourceWidth) % sourceWidth;
    let seriesValue;
    if (horizontalSeries && row === source.top && column < source.left) {
      sourceColumn = source.left;
      seriesValue = makeSeriesValue(horizontalSeries, true, source.left - column);
    } else if (horizontalSeries && row === source.top && column > source.right) {
      sourceColumn = source.right;
      seriesValue = makeSeriesValue(horizontalSeries, false, column - source.right);
    } else if (verticalSeries && column === source.left && row < source.top) {
      sourceRow = source.top;
      seriesValue = makeSeriesValue(verticalSeries, true, source.top - row);
    } else if (verticalSeries && column === source.left && row > source.bottom) {
      sourceRow = source.bottom;
      seriesValue = makeSeriesValue(verticalSeries, false, row - source.bottom);
    }
    const sourceCell = document.getCell(sourceRow, sourceColumn, sheetIndex), targetCell = document.getCell(row, column, sheetIndex);
    const sourceHasFormula = isFormulaCell(sourceCell);
    if ((sourceCell?.kind === 'error' && !sourceHasFormula) || (sourceHasFormula && (!sourceCell.formula || sourceCell.formulaType)) || (targetCell?.kind === 'error' && document.kind !== 'xlsx') || (isFormulaCell(targetCell) && (targetCell.formulaType || document.kind !== 'xlsx'))) throw new Error('공유·배열 수식 또는 일반 오류 셀이 포함된 범위는 수정할 수 없습니다.');
    const edit = sourceHasFormula
      ? { row, column, value: '', formula: typeof sourceCell.formula === 'string' ? translateFormulaReferences(sourceCell.formula, row - sourceRow, column - sourceColumn) : null }
      : { row, column, value: seriesValue ?? sourceCell?.value ?? '', ...(document.kind === 'xlsx' ? { formula: null } : {}) };
    if (sourceHasFormula && !edit.formula) throw new Error('수식 참조를 확인할 수 없는 셀은 채울 수 없습니다.');
    if (document.kind === 'xlsx') { edit.styleIndex = sourceCell?.style ?? null; edit.styleProperties = sourceCell?.styleProperties ?? null; edit.kind = sourceCell?.kind ?? 'blank'; }
    edits.push(edit);
  }
  return edits;
}

function buildNumberFormatEdits(document, range, numberFormatId, sheetIndex = 0) {
  const format = typeof numberFormatId === 'number' ? { numberFormatId } : numberFormatId;
  if (document.kind !== 'xlsx' || !format || !Number.isInteger(format.numberFormatId) || !([4, 9, 14, 20].includes(format.numberFormatId) || (format.numberFormatId >= 164 && format.numberFormatId <= 65535 && typeof format.numberFormatCode === 'string' && format.numberFormatCode.length <= 255)) || !Number.isInteger(sheetIndex) || !document.state.sheets[sheetIndex]) throw new RangeError('숫자 표시 형식 요청이 유효하지 않습니다.');
  const bounds = selectionBounds(range, '숫자 표시 형식', Number.MAX_SAFE_INTEGER);
  if (bounds.bottom >= XLSX_LIMITS.rows || bounds.right >= XLSX_LIMITS.columns) throw new RangeError('숫자 표시 형식 범위가 시트 한도를 벗어났습니다.');
  const candidates = new Map(), sheet = document.state.sheets[sheetIndex];
  for (const cell of sheet.cells.values()) if (cell.row >= bounds.top && cell.row <= bounds.bottom && cell.column >= bounds.left && cell.column <= bounds.right) candidates.set(cell.row + ':' + cell.column, [cell.row, cell.column]);
  for (const key of [...document.drafts.keys(), ...document.formulaDrafts.keys()]) {
    const [sheetText, address] = key.split(':');
    if (Number(sheetText) !== sheetIndex) continue;
    const coordinate = parseAddress(address);
    if (coordinate.row >= bounds.top && coordinate.row <= bounds.bottom && coordinate.column >= bounds.left && coordinate.column <= bounds.right) candidates.set(coordinate.row + ':' + coordinate.column, [coordinate.row, coordinate.column]);
  }
  const edits = [];
  for (const [row, column] of candidates.values()) {
    const cell = document.getCell(row, column, sheetIndex);
    if (cell?.kind !== 'number' || cell.formulaType || (cell.hasFormula && !cell.formula)) continue;
    edits.push({ row, column, value: cell.value, kind: 'number', ...format, forceNumberFormat: true, ...(cell.formula ? { formula: cell.formula } : {}) });
  }
  return edits;
}

function nextCustomNumberFormatId(document) {
  const stylesPath = document.state.stylesPath;
  const stylesBytes = stylesPath && document.state.files?.[stylesPath];
  if (!stylesBytes) throw new Error('사용자 지정 숫자 형식을 추가할 스타일 정의가 없습니다.');
  const stylesXml = new TextDecoder().decode(stylesBytes);
  const used = new Set([...stylesXml.matchAll(/\bnumFmtId\s*=\s*(["'])(\d+)\1/g)].map(match => Number(match[2])));
  for (let id = 164; id <= 65535; id++) if (!used.has(id)) return id;
  throw new Error('사용할 수 있는 사용자 지정 숫자 형식 ID가 없습니다.');
}

function isSupportedCustomNumberFormatCode(formatCode) {
  return typeof formatCode === 'string' && formatCode.length <= 255 && /^(?:[$€£¥₩])?(?:#,##0|0)(?:\.[0#]{1,10})?%?$/.test(formatCode);
}

function buildFontFormatEdits(document, range, property, sheetIndex = 0, value = undefined) {
  if (document.kind !== 'xlsx' || !['bold', 'italic', 'underline', 'strike', 'fontSize', 'fontFamily'].includes(property) || !Number.isInteger(sheetIndex) || !document.state.sheets[sheetIndex] || (property === 'fontSize' ? value !== null && (typeof value !== 'number' || !Number.isFinite(value) || value < 1 || value > 409 || value * 2 % 1 !== 0) : property === 'fontFamily' ? value !== null && (typeof value !== 'string' || !value.trim() || value.length > 64) : value !== undefined)) throw new RangeError('글꼴 서식 요청이 유효하지 않습니다.');
  const bounds = selectionBounds(range, '글꼴 서식', Number.MAX_SAFE_INTEGER);
  if (bounds.bottom >= XLSX_LIMITS.rows || bounds.right >= XLSX_LIMITS.columns) throw new RangeError('글꼴 서식 범위가 시트 한도를 벗어났습니다.');
  const candidates = new Map(), sheet = document.state.sheets[sheetIndex];
  for (const cell of sheet.cells.values()) if (cell.row >= bounds.top && cell.row <= bounds.bottom && cell.column >= bounds.left && cell.column <= bounds.right) candidates.set(cell.row + ':' + cell.column, [cell.row, cell.column]);
  for (const key of [...document.drafts.keys(), ...document.formulaDrafts.keys(), ...document.styleDrafts.keys()]) {
    const [sheetText, address] = key.split(':');
    if (Number(sheetText) !== sheetIndex) continue;
    const coordinate = parseAddress(address);
    if (coordinate.row >= bounds.top && coordinate.row <= bounds.bottom && coordinate.column >= bounds.left && coordinate.column <= bounds.right) candidates.set(coordinate.row + ':' + coordinate.column, [coordinate.row, coordinate.column]);
  }
  const selectionArea = (bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1);
  if (selectionArea <= 32768) for (let row = bounds.top; row <= bounds.bottom; row++) for (let column = bounds.left; column <= bounds.right; column++) candidates.set(row + ':' + column, [row, column]);
  const cells = [...candidates.values()].map(([row, column]) => ({ row, column, cell: document.getCell(row, column, sheetIndex) }))
    .filter(({ cell }) => cell && !cell.formulaType && cell.kind !== 'error');
  if (!cells.length) return [];
  const enabled = property === 'fontSize' ? value !== null : property === 'fontFamily' ? value !== null : !cells.every(({ cell }) => cell.styleProperties?.[property] === true);
  return cells.map(({ row, column, cell }) => {
    const properties = { ...(cell.styleProperties || {}) };
    if (property === 'fontSize') { if (value === null) delete properties.fontSize; else properties.fontSize = value; }
    else if (property === 'fontFamily') { if (value === null) delete properties.fontFamily; else properties.fontFamily = value.trim(); }
    else if (enabled) properties[property] = true; else delete properties[property];
    const nextProperties = Object.keys(properties).length ? properties : null;
    return {
      row, column, value: cell.value ?? '', kind: cell.kind ?? 'blank',
      styleIndex: cell.style ?? null, styleProperties: nextProperties,
      font: { bold: properties.bold === true, italic: properties.italic === true, underline: properties.underline === true, fontSize: properties.fontSize ?? null, ...(property === 'fontFamily' || properties.fontFamily ? { fontFamily: property === 'fontFamily' ? value === null ? null : value.trim() : properties.fontFamily } : {}), ...(property === 'strike' || properties.strike === true ? { strike: property === 'strike' ? enabled : true } : {}) },
      ...(cell.formula ? { formula: cell.formula } : {})
    };
  });
}

function buildFontColorEdits(document, range, color, sheetIndex = 0) {
  if (document.kind !== 'xlsx' || (color !== null && (typeof color !== 'string' || !/^#[0-9A-F]{6}$/i.test(color))) || !Number.isInteger(sheetIndex) || !document.state.sheets[sheetIndex]) throw new RangeError('글꼴 색 요청이 유효하지 않습니다.');
  const bounds = selectionBounds(range, '글꼴 색', Number.MAX_SAFE_INTEGER);
  if (bounds.bottom >= XLSX_LIMITS.rows || bounds.right >= XLSX_LIMITS.columns) throw new RangeError('글꼴 색 범위가 시트 한도를 벗어났습니다.');
  const candidates = new Map(), sheet = document.state.sheets[sheetIndex];
  for (const cell of sheet.cells.values()) if (cell.row >= bounds.top && cell.row <= bounds.bottom && cell.column >= bounds.left && cell.column <= bounds.right) candidates.set(cell.row + ':' + cell.column, [cell.row, cell.column]);
  for (const key of [...document.drafts.keys(), ...document.formulaDrafts.keys(), ...document.styleDrafts.keys()]) {
    const [sheetText, address] = key.split(':');
    if (Number(sheetText) !== sheetIndex) continue;
    const coordinate = parseAddress(address);
    if (coordinate.row >= bounds.top && coordinate.row <= bounds.bottom && coordinate.column >= bounds.left && coordinate.column <= bounds.right) candidates.set(coordinate.row + ':' + coordinate.column, [coordinate.row, coordinate.column]);
  }
  const selectionArea = (bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1);
  if (selectionArea <= 32768) for (let row = bounds.top; row <= bounds.bottom; row++) for (let column = bounds.left; column <= bounds.right; column++) candidates.set(row + ':' + column, [row, column]);
  return [...candidates.values()].map(([row, column]) => ({ row, column, cell: document.getCell(row, column, sheetIndex) }))
    .filter(({ cell }) => cell && !cell.formulaType && cell.kind !== 'error')
    .map(({ row, column, cell }) => {
      const properties = { ...(cell.styleProperties || {}) };
      if (color === null) delete properties.color; else properties.color = color.toUpperCase();
      return {
        row, column, value: cell.value ?? '', kind: cell.kind ?? 'blank',
        styleIndex: cell.style ?? null, styleProperties: Object.keys(properties).length ? properties : null,
        font: { bold: properties.bold === true, italic: properties.italic === true, color: color === null ? null : color.toUpperCase(), underline: properties.underline === true, fontSize: properties.fontSize ?? null, ...(properties.strike === true ? { strike: true } : {}) },
        ...(cell.formula ? { formula: cell.formula } : {})
      };
    });
}

function buildAlignmentFormatEdits(document, range, axis, value, sheetIndex = 0) {
  const values = axis === 'horizontal' ? ['left', 'center', 'right'] : axis === 'vertical' ? ['top', 'center', 'bottom'] : axis === 'wrapText' ? [true, false] : [];
  if (document.kind !== 'xlsx' || !values.includes(value) || !Number.isInteger(sheetIndex) || !document.state.sheets[sheetIndex]) throw new RangeError('셀 맞춤 요청이 유효하지 않습니다.');
  const bounds = selectionBounds(range, '셀 맞춤', Number.MAX_SAFE_INTEGER);
  if (bounds.bottom >= XLSX_LIMITS.rows || bounds.right >= XLSX_LIMITS.columns) throw new RangeError('셀 맞춤 범위가 시트 한도를 벗어났습니다.');
  const candidates = new Map(), sheet = document.state.sheets[sheetIndex];
  for (const cell of sheet.cells.values()) if (cell.row >= bounds.top && cell.row <= bounds.bottom && cell.column >= bounds.left && cell.column <= bounds.right) candidates.set(cell.row + ':' + cell.column, [cell.row, cell.column]);
  for (const key of [...document.drafts.keys(), ...document.formulaDrafts.keys(), ...document.styleDrafts.keys()]) {
    const [sheetText, address] = key.split(':');
    if (Number(sheetText) !== sheetIndex) continue;
    const coordinate = parseAddress(address);
    if (coordinate.row >= bounds.top && coordinate.row <= bounds.bottom && coordinate.column >= bounds.left && coordinate.column <= bounds.right) candidates.set(coordinate.row + ':' + coordinate.column, [coordinate.row, coordinate.column]);
  }
  const selectionArea = (bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1);
  if (selectionArea <= 32768) for (let row = bounds.top; row <= bounds.bottom; row++) for (let column = bounds.left; column <= bounds.right; column++) candidates.set(row + ':' + column, [row, column]);
  return [...candidates.values()].map(([row, column]) => ({ row, column, cell: document.getCell(row, column, sheetIndex) }))
    .filter(({ cell }) => cell && !cell.formulaType && cell.kind !== 'error')
    .map(({ row, column, cell }) => {
      const properties = { ...(cell.styleProperties || {}), [axis]: value };
      return {
        row, column, value: cell.value ?? '', kind: cell.kind ?? 'blank',
        styleIndex: cell.style ?? null, styleProperties: properties,
        alignment: { [axis]: value }, ...(cell.formula ? { formula: cell.formula } : {})
      };
    });
}

function validateRowFitResult(pending, message, revision, selectedSheet) {
  if (pending.sheetIndex !== message.sheetIndex || pending.sheetIndex !== selectedSheet || pending.revision !== revision) throw new Error('시트가 변경되어 행 자동 맞춤을 다시 실행해야 합니다.');
  if (!Array.isArray(message.heights) || message.heights.length !== pending.rows.length || message.heights.some((item, index) => !item || item.row !== pending.rows[index] || !Number.isFinite(item.height) || item.height < .75 || item.height > 409.5)) throw new Error('행 자동 맞춤 결과가 유효하지 않습니다.');
}

function buildWrapTextEdits(document, range, sheetIndex = 0) {
  const edits = buildAlignmentFormatEdits(document, range, 'wrapText', true, sheetIndex).filter(edit => !document.state.sheets[sheetIndex].mergedCells?.some(merged => edit.row >= merged.top && edit.row <= merged.bottom && edit.column >= merged.left && edit.column <= merged.right && (edit.row !== merged.top || edit.column !== merged.left)));
  const enabled = !edits.length || !edits.every(edit => document.getCell(edit.row, edit.column, sheetIndex)?.styleProperties?.wrapText);
  return edits.map(edit => ({ ...edit, styleProperties: { ...edit.styleProperties, wrapText: enabled }, alignment: { wrapText: enabled } }));
}

function buildAutomaticRowFitRequest(document, sheetIndex = 0, language = 'en') {
  const sheet = document.state.sheets[sheetIndex];
  const manual = new Set(sheet.customRowHeights || []);
  const rows = new Set(Object.keys(sheet.rowHeights || {}).map(Number).filter(row => !manual.has(row)));
  const coordinates = new Map([...sheet.cells.values()].map(cell => [cell.address, cell]));
  for (const key of [...document.drafts.keys(), ...document.formulaDrafts.keys(), ...document.styleDrafts.keys()]) {
    const [index, address] = key.split(':');
    if (Number(index) === sheetIndex) coordinates.set(address, parseAddress(address));
  }
  for (const { row, column } of coordinates.values()) {
    if (!manual.has(row) && document.getCell(row, column, sheetIndex)?.styleProperties?.wrapText) rows.add(row);
  }
  return buildRowFitRequest(document, 0, XLSX_LIMITS.rows - 1, sheetIndex, [...rows].map(row => ({ row })), true, language);
}

function buildRowFitRequest(document, startRow, endRow, sheetIndex = 0, edits = [], automatic = false, language = 'en') {
  const sheet = document.state.sheets[sheetIndex];
  if (document.kind !== 'xlsx' || !sheet || !Number.isInteger(startRow) || !Number.isInteger(endRow) || startRow < 0 || endRow < startRow || endRow >= XLSX_LIMITS.rows) throw new RangeError('행 자동 맞춤 범위가 유효하지 않습니다.');
  const targets = new Set();
  if (automatic) for (const edit of edits) targets.add(edit.row);
  else if (endRow - startRow < 10000) for (let row = startRow; row <= endRow; row++) targets.add(row);
  else {
    if (startRow !== 0 || endRow !== XLSX_LIMITS.rows - 1) throw new RangeError('행 자동 맞춤은 10,000행 이하 또는 전체 행을 선택하세요.');
    for (const cell of sheet.cells.values()) targets.add(cell.row);
    for (const row of Object.keys(sheet.rowHeights || {})) targets.add(Number(row));
  }
  const coordinates = new Map([...sheet.cells.values()].map(cell => [cell.row + ':' + cell.column, [cell.row, cell.column]]));
  for (const key of [...document.drafts.keys(), ...document.formulaDrafts.keys(), ...document.styleDrafts.keys()]) {
    const [index, address] = key.split(':'); if (Number(index) !== sheetIndex) continue;
    const coordinate = parseAddress(address); coordinates.set(coordinate.row + ':' + coordinate.column, [coordinate.row, coordinate.column]);
    if (!automatic && endRow - startRow >= 10000) targets.add(coordinate.row);
  }
  if (automatic) for (const row of sheet.customRowHeights || []) targets.delete(row);
  if (targets.size > 10000) throw new RangeError('행 자동 맞춤은 한 번에 10,000행 이하를 지원합니다.');
  const rows = new Map([...targets].sort((left, right) => left - right).map(row => [row, { row, cells: [] }]));
  const overlays = new Map(edits.map(edit => [edit.row + ':' + edit.column, edit.styleProperties]));
  const calculator = createFormulaCalculator(document), mergedOnly = new Set();
  let characters = 0, count = 0;
  for (const [row, column] of coordinates.values()) {
    const target = rows.get(row); if (!target || sheet.hiddenColumns?.includes(column)) continue;
    if (sheet.mergedCells?.some(merged => row >= merged.top && row <= merged.bottom && column >= merged.left && column <= merged.right)) { mergedOnly.add(row); continue; }
    const cell = calculator.cell(sheetIndex, row, column); if (!cell) continue;
    const text = formatCellValue(cell, language), style = overlays.get(row + ':' + column) || cell.styleProperties || {};
    characters += text.length; if (++count > 32768 || characters > 2000000) throw new RangeError('행 자동 맞춤 데이터가 너무 큽니다. 더 작은 범위를 선택하세요.');
    const width = excelColumnWidthPixels(document.columnWidthDrafts.get(sheetIndex + ':' + column) ?? sheet.columnWidths[column], xlsxDefaultColumnWidth(sheet, document.maximumDigitWidth), document.maximumDigitWidth);
    target.cells.push({ text, width, fontSize: style.fontSize || null, fontFamily: style.fontFamily || null, bold: Boolean(style.bold), italic: Boolean(style.italic), wrapText: Boolean(style.wrapText) });
  }
  return { rows: [...rows.values()].filter(row => row.cells.length || !mergedOnly.has(row.row)), defaultHeight: sheet.defaultRowHeight || 26 };
}

function buildBorderFormatEdits(document, range, preset, sheetIndex = 0, line = 'thin', color = '#000000') {
  const presets = ['none', 'outside', 'all', 'top', 'bottom', 'left', 'right', 'inside-horizontal', 'inside-vertical'];
  const lines = ['hair', 'thin', 'medium', 'thick', 'dashed', 'dotted', 'double'];
  if (document.kind !== 'xlsx' || !presets.includes(preset) || !Number.isInteger(sheetIndex) || !document.state.sheets[sheetIndex] || !lines.includes(line) || typeof color !== 'string' || !/^#[0-9A-F]{6}$/i.test(color)) throw new RangeError('셀 테두리 요청이 유효하지 않습니다.');
  const bounds = selectionBounds(range, '셀 테두리', Number.MAX_SAFE_INTEGER);
  if (bounds.bottom >= XLSX_LIMITS.rows || bounds.right >= XLSX_LIMITS.columns) throw new RangeError('셀 테두리 범위가 시트 한도를 벗어났습니다.');
  const candidates = new Map(), sheet = document.state.sheets[sheetIndex];
  for (const cell of sheet.cells.values()) if (cell.row >= bounds.top && cell.row <= bounds.bottom && cell.column >= bounds.left && cell.column <= bounds.right) candidates.set(cell.row + ':' + cell.column, [cell.row, cell.column]);
  for (const key of [...document.drafts.keys(), ...document.formulaDrafts.keys(), ...document.styleDrafts.keys()]) {
    const [sheetText, address] = key.split(':');
    if (Number(sheetText) !== sheetIndex) continue;
    const coordinate = parseAddress(address);
    if (coordinate.row >= bounds.top && coordinate.row <= bounds.bottom && coordinate.column >= bounds.left && coordinate.column <= bounds.right) candidates.set(coordinate.row + ':' + coordinate.column, [coordinate.row, coordinate.column]);
  }
  const selectionArea = (bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1);
  if (selectionArea <= 32768) for (let row = bounds.top; row <= bounds.bottom; row++) for (let column = bounds.left; column <= bounds.right; column++) candidates.set(row + ':' + column, [row, column]);
  const result = [];
  for (const [row, column] of candidates.values()) {
    const cell = document.getCell(row, column, sheetIndex);
    if (!cell || cell.formulaType || cell.kind === 'error') continue;
    const sides = [];
    if (preset === 'all') sides.push('left', 'right', 'top', 'bottom');
    else if (preset === 'none') sides.push('left', 'right', 'top', 'bottom');
    else if (preset === 'outside') {
      if (column === bounds.left) sides.push('left');
      if (column === bounds.right) sides.push('right');
      if (row === bounds.top) sides.push('top');
      if (row === bounds.bottom) sides.push('bottom');
    } else if (preset === 'inside-horizontal') {
      if (row < bounds.bottom) sides.push('bottom');
    } else if (preset === 'inside-vertical') {
      if (column < bounds.right) sides.push('right');
    } else sides.push(preset);
    if (!sides.length) continue;
    const border = {}, properties = { ...(cell.styleProperties || {}) }, borders = { ...(properties.borders || {}) };
    for (const side of sides) {
      const edge = preset === 'none' ? null : { line, color: color.toUpperCase() };
      border[side] = edge;
      if (edge) borders[side] = edge; else delete borders[side];
    }
    if (Object.keys(borders).length) properties.borders = borders; else delete properties.borders;
    result.push({ row, column, value: cell.value ?? '', kind: cell.kind ?? 'blank', styleIndex: cell.style ?? null, styleProperties: Object.keys(properties).length ? properties : null, border, ...(cell.formula ? { formula: cell.formula } : {}) });
  }
  return result;
}

function buildFillFormatEdits(document, range, color, sheetIndex = 0) {
  if (document.kind !== 'xlsx' || (color !== null && (typeof color !== 'string' || !/^#[0-9A-F]{6}$/i.test(color))) || !Number.isInteger(sheetIndex) || !document.state.sheets[sheetIndex]) throw new RangeError('채우기 서식 요청이 유효하지 않습니다.');
  const bounds = selectionBounds(range, '채우기 서식', Number.MAX_SAFE_INTEGER);
  if (bounds.bottom >= XLSX_LIMITS.rows || bounds.right >= XLSX_LIMITS.columns) throw new RangeError('채우기 서식 범위가 시트 한도를 벗어났습니다.');
  const candidates = new Map(), sheet = document.state.sheets[sheetIndex];
  for (const cell of sheet.cells.values()) if (cell.row >= bounds.top && cell.row <= bounds.bottom && cell.column >= bounds.left && cell.column <= bounds.right) candidates.set(cell.row + ':' + cell.column, [cell.row, cell.column]);
  for (const key of [...document.drafts.keys(), ...document.formulaDrafts.keys(), ...document.styleDrafts.keys()]) {
    const [sheetText, address] = key.split(':');
    if (Number(sheetText) !== sheetIndex) continue;
    const coordinate = parseAddress(address);
    if (coordinate.row >= bounds.top && coordinate.row <= bounds.bottom && coordinate.column >= bounds.left && coordinate.column <= bounds.right) candidates.set(coordinate.row + ':' + coordinate.column, [coordinate.row, coordinate.column]);
  }
  const selectionArea = (bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1);
  if (selectionArea <= 32768) for (let row = bounds.top; row <= bounds.bottom; row++) for (let column = bounds.left; column <= bounds.right; column++) candidates.set(row + ':' + column, [row, column]);
  return [...candidates.values()].map(([row, column]) => ({ row, column, cell: document.getCell(row, column, sheetIndex) }))
    .filter(({ cell }) => cell && !cell.formulaType && cell.kind !== 'error')
    .map(({ row, column, cell }) => {
      const properties = { ...(cell.styleProperties || {}) };
      if (color === null) delete properties.fill; else properties.fill = color.toUpperCase();
      return {
        row, column, value: cell.value ?? '', kind: cell.kind ?? 'blank',
        styleIndex: cell.style ?? null, styleProperties: Object.keys(properties).length ? properties : null,
        fillColor: color === null ? null : color.toUpperCase(),
        ...(cell.formula ? { formula: cell.formula } : {})
      };
    });
}

function buildSortEdits(document, range, keyColumn, direction, hasHeader = false, expandSelection = true, sheetIndex = 0, hiddenRows = []) {
  const sortKeys = Array.isArray(keyColumn) ? keyColumn : [{ column: keyColumn, direction }];
  if (!sortKeys.length || sortKeys.length > 64 || sortKeys.some(key => !key || !Number.isInteger(key.column) || key.column < 0 || !['ascending', 'descending'].includes(key.direction)) || new Set(sortKeys.map(key => key.column)).size !== sortKeys.length) throw new RangeError('정렬 기준이 유효하지 않습니다.');
  const maximumSortCells = document.kind === 'xlsx' ? XLSX_LIMITS.cells : 20000;
  let bounds = selectionBounds(range, '정렬', maximumSortCells);
  if (sortKeys.some(key => key.column < bounds.left || key.column > bounds.right)) throw new RangeError('정렬 기준 열이 선택 범위에 없습니다.');
  if (expandSelection) {
    let top = bounds.top, bottom = bounds.bottom, left = bounds.left, right = bounds.right;
    const ensureArea = () => { if ((bottom - top + 1) * (right - left + 1) > 10000) throw new RangeError('정렬 선택 영역은 10,000개 셀까지 확장할 수 있습니다.'); };
    const columnHasData = column => {
      for (let row = top; row <= bottom; row++) if ((document.getCell(row, column, sheetIndex)?.value ?? '') !== '') return true;
      return false;
    };
    const rowHasData = row => {
      for (let column = left; column <= right; column++) if ((document.getCell(row, column, sheetIndex)?.value ?? '') !== '') return true;
      return false;
    };
    const maximumColumns = document.kind === 'csv' ? CSV_LIMIT_COLUMNS : XLSX_LIMITS.columns;
    const sheet = document.kind === 'xlsx' ? document.state.sheets[sheetIndex] : null;
    let maximumRows = document.kind === 'csv' ? document.state.rows.length : bottom + 1;
    if (document.kind === 'xlsx') for (const cell of sheet?.cells.values() || []) maximumRows = Math.max(maximumRows, cell.row + 1);
    while (left > 0 && columnHasData(left - 1)) { left--; ensureArea(); }
    while (right + 1 < maximumColumns && columnHasData(right + 1)) { right++; ensureArea(); }
    if (top === bottom) {
      while (top > 0 && rowHasData(top - 1)) { top--; ensureArea(); }
      while (bottom + 1 < maximumRows && rowHasData(bottom + 1)) { bottom++; ensureArea(); }
      while (left > 0 && columnHasData(left - 1)) { left--; ensureArea(); }
      while (right + 1 < maximumColumns && columnHasData(right + 1)) { right++; ensureArea(); }
    }
    if (top !== bounds.top || bottom !== bounds.bottom || left !== bounds.left || right !== bounds.right) bounds = selectionBounds({ startRow: top, endRow: bottom, startColumn: left, endColumn: right }, '붙여넣기');
  }
  assertNoHyperlinkRange(document, sheetIndex, bounds, '정렬');
  const firstRow = bounds.top + (hasHeader ? 1 : 0);
  const sortRows = [];
  const hidden = new Set(hiddenRows);
  for (let row = firstRow; row <= bounds.bottom; row++) if (!hidden.has(row)) sortRows.push(row);
  if (sortRows.length < 2) return [];
  const rows = [], formulaCalculator = document.kind === 'xlsx' ? createFormulaCalculator(document) : null;
  for (const row of sortRows) {
    const cells = [];
    for (let column = bounds.left; column <= bounds.right; column++) {
      const cell = document.getCell(row, column, sheetIndex);
      const hasFormula = isFormulaCell(cell);
      if ((cell?.kind === 'error' && !hasFormula) || (hasFormula && (!cell.formula || cell.formulaType))) throw new Error('공유·배열 수식 또는 일반 오류 셀이 포함된 범위는 정렬할 수 없습니다.');
      const value = hasFormula ? formulaCalculator.value(sheetIndex, row, column) : cell?.value ?? '';
      if (hasFormula && typeof value === 'string' && /^#(?:REF|VALUE|NAME|DIV\/0|N\/A|NUM|NULL|SPILL|CALC)!/.test(value)) throw new Error('오류를 반환하는 수식 셀이 포함된 범위는 정렬할 수 없습니다.');
      cells.push({ value: String(value ?? ''), rawValue: cell?.value ?? '', formula: hasFormula ? cell.formula : null, sourceRow: row, style: cell?.style ?? null, styleProperties: cell?.styleProperties ?? null, kind: typeof value === 'number' ? 'number' : typeof value === 'boolean' ? 'boolean' : cell?.kind ?? 'blank' });
    }
    rows.push({ sourceRow: row, cells, keys: sortKeys.map(key => ({ ...key, value: cells[key.column - bounds.left] })) });
  }
  const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' });
  const compare = (left, right) => {
    for (let index = 0; index < sortKeys.length; index++) {
      const key = sortKeys[index], leftKey = left.keys[index].value, rightKey = right.keys[index].value;
      const leftValue = leftKey.value, rightValue = rightKey.value;
      const leftBlank = leftValue === '', rightBlank = rightValue === '';
      if (leftBlank !== rightBlank) return leftBlank ? 1 : -1;
      if (leftBlank) continue;
      const leftNumber = leftKey.kind === 'number' || /^\s*[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?\s*$/.test(leftValue);
      const rightNumber = rightKey.kind === 'number' || /^\s*[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?\s*$/.test(rightValue);
      const leftRank = leftNumber ? 0 : leftKey.kind === 'boolean' ? 1 : 2, rightRank = rightNumber ? 0 : rightKey.kind === 'boolean' ? 1 : 2;
      let result = leftRank - rightRank;
      if (!result) result = leftNumber && rightNumber && Number.isFinite(Number(leftValue)) && Number.isFinite(Number(rightValue)) ? Number(leftValue) - Number(rightValue) : collator.compare(leftValue, rightValue);
      if (result) return (key.direction === 'ascending' ? result : -result);
    }
    return left.sourceRow - right.sourceRow;
  };
  rows.sort(compare);
  const edits = [];
  for (let offset = 0; offset < rows.length; offset++) {
    const targetRow = sortRows[offset], sourceRow = rows[offset].sourceRow;
    for (let columnOffset = 0; columnOffset < rows[offset].cells.length; columnOffset++) {
      const column = bounds.left + columnOffset, source = rows[offset].cells[columnOffset], target = document.getCell(targetRow, column, sheetIndex);
      const formula = source.formula ? translateFormulaReferences(source.formula, targetRow - source.sourceRow, 0) : null;
      if ((target?.value ?? '') === source.rawValue && (document.kind !== 'xlsx' || ((target?.style ?? null) === source.style && (target?.formula ?? null) === formula))) continue;
      const edit = { row: targetRow, column, value: source.rawValue };
      if (document.kind === 'xlsx') { edit.styleIndex = source.style; edit.styleProperties = source.styleProperties; edit.formula = formula; edit.kind = source.kind; }
      edits.push(edit);
    }
  }
  return edits;
}

function planTextToColumns(document, range, delimiter, mergeConsecutive = false, sheetIndex = 0) {
  if (!document?.writable || !['csv', 'xlsx'].includes(document.kind)) throw new Error('현재 문서에서는 텍스트 나누기를 사용할 수 없습니다.');
  const bounds = selectionBounds(range, '붙여넣기', 10000);
  if (bounds.left !== bounds.right) throw new RangeError('텍스트 나누기는 한 번에 한 열만 선택하세요.');
  const maxRows = document.kind === 'csv' ? CSV_LIMIT_ROWS : XLSX_LIMITS.rows;
  const maxColumns = document.kind === 'csv' ? CSV_LIMIT_COLUMNS : XLSX_LIMITS.columns;
  if (bounds.bottom >= maxRows || bounds.left >= maxColumns) throw new RangeError('텍스트 나누기 범위가 시트 경계를 벗어났습니다.');
  if (document.kind === 'xlsx' && !document.state.sheets[sheetIndex]) throw new RangeError('나눌 시트를 찾을 수 없습니다.');
  const rows = [];
  let width = 1;
  for (let row = bounds.top; row <= bounds.bottom; row++) {
    const cell = document.getCell(row, bounds.left, sheetIndex);
    if (isFormulaCell(cell)) throw new Error('수식 셀이 포함된 선택 범위는 텍스트로 나눌 수 없습니다. 수식을 값으로 바꾸세요.');
    const value = cell?.value === null || cell?.value === undefined ? '' : String(cell.value);
    const fields = splitDelimitedText(value, delimiter, mergeConsecutive);
    rows.push({ row, value, fields });
    width = Math.max(width, fields.length);
  }
  if (width < 2) throw new Error('선택 범위에서 지정한 구분자를 찾지 못했습니다.');
  if (bounds.left + width > maxColumns) throw new RangeError('나눈 결과가 워크시트의 마지막 열을 벗어납니다.');
  const area = rows.length * width;
  if (area > 20000) throw new RangeError('텍스트 나누기는 한 번에 20,000개 이하의 결과 셀을 만들 수 있습니다.');
  const outputBounds = { ...bounds, right: bounds.left + width - 1 };
  assertNoHyperlinkRange(document, sheetIndex, outputBounds, '텍스트 나누기');
  if (document.kind === 'xlsx') {
    const tables = document.state.sheets[sheetIndex].tables || [];
    if (tables.some(table => outputBounds.top <= table.range.bottom && outputBounds.bottom >= table.range.top && outputBounds.left <= table.range.right && outputBounds.right >= table.range.left)) throw new Error('표 안에서는 텍스트 나누기를 사용할 수 없습니다. 표를 일반 범위로 바꾼 뒤 다시 시도하세요.');
  }
  const edits = [], conflicts = [];
  for (const item of rows) {
    if (!item.value) continue;
    for (let offset = 0; offset < width; offset++) {
      const value = item.fields[offset] ?? '', column = bounds.left + offset, existing = document.getCell(item.row, column, sheetIndex);
      if (offset > 0 && (existing?.value ?? '') !== '' && (existing?.value ?? '') !== value || offset > 0 && isFormulaCell(existing)) conflicts.push({ row: item.row, column });
      const edit = { row: item.row, column, value };
      if (document.kind === 'xlsx') { edit.kind = 'string'; edit.formula = null; }
      edits.push(edit);
    }
  }
  if (!edits.length) throw new Error('나눌 텍스트가 없습니다.');
  return { edits, conflicts, preview: rows.slice(0, 6).map(item => item.fields), width };
}

function planRemoveDuplicates(document, range, columns, hasHeader = true, sheetIndex = 0) {
  if (!document?.writable || !['csv', 'xlsx'].includes(document.kind)) throw new Error('현재 문서에서는 중복 행 제거를 사용할 수 없습니다.');
  const bounds = selectionBounds(range, '중복 제거', 100000);
  if (document.kind === 'csv' && sheetIndex !== 0 || document.kind === 'xlsx' && !document.state.sheets[sheetIndex]) throw new RangeError('중복 제거 시트를 찾을 수 없습니다.');
  if (typeof hasHeader !== 'boolean' || !Array.isArray(columns) || !columns.length || columns.some(column => !Number.isInteger(column) || column < bounds.left || column > bounds.right) || new Set(columns).size !== columns.length) throw new RangeError('중복 비교 열이 유효하지 않습니다.');
  const firstRow = bounds.top + (hasHeader ? 1 : 0);
  if (firstRow > bounds.bottom) throw new RangeError('중복 제거 범위에 데이터 행이 없습니다.');
  const tables = document.kind === 'xlsx' ? document.state.sheets[sheetIndex].tables || [] : [];
  const calculator = document.kind === 'xlsx' ? createFormulaCalculator(document) : null;
  const seen = new Set(), duplicateRows = [];
  for (let row = firstRow; row <= bounds.bottom; row++) {
    const values = columns.map(column => {
      const cell = document.getCell(row, column, sheetIndex);
      const value = document.kind === 'xlsx' && isFormulaCell(cell) ? calculator.value(sheetIndex, row, column) : cell?.value;
      return String(value ?? '').toLocaleLowerCase();
    });
    const key = JSON.stringify(values);
    if (seen.has(key)) duplicateRows.push(row);
    else seen.add(key);
  }
  if (document.kind === 'xlsx' && duplicateRows.length) {
    const sheet = document.state.sheets[sheetIndex], deleted = new Set(duplicateRows);
    if (sheet.mergedCells?.some(range => [...deleted].some(row => row >= range.top && row <= range.bottom))) throw new Error('삭제 대상 행이 병합 셀과 겹쳐 중복 행을 안전하게 제거할 수 없습니다.');
    for (const table of tables) {
      const headerEnd = table.range.top + (table.headerRowCount || 0) - 1, totalsStart = table.range.bottom - (table.totalsRowCount || 0) + 1;
      if (duplicateRows.some(row => row >= table.range.top && row <= table.range.bottom && (row <= headerEnd || row >= totalsStart))) throw new Error('Excel 표의 머리글·합계 행은 중복 제거로 삭제할 수 없습니다.');
    }
  }
  return { range: bounds, duplicateRows, removedCount: duplicateRows.length, hasHeader, columns: [...columns] };
}

function removeDuplicateRows(document, plan, sheetIndex = 0) {
  if (!plan || !Array.isArray(plan.duplicateRows) || plan.duplicateRows.some((row, index) => !Number.isInteger(row) || row < 0 || index && row <= plan.duplicateRows[index - 1])) throw new RangeError('중복 행 제거 계획이 유효하지 않습니다.');
  if (!plan.duplicateRows.length) return;
  const before = document.captureSnapshot();
  try {
    for (const row of [...plan.duplicateRows].reverse()) document.applyStructureChange({ axis: 'row', action: 'delete', index: row, sheetIndex, emit: false });
  } catch (error) {
    document.restoreSnapshot(before);
    document.renderEmitter.fire();
    throw error;
  }
  const after = document.captureSnapshot();
  document.emitter.fire({ label: '중복 행 제거', undo: () => { document.restoreSnapshot(before); document.renderEmitter.fire(); }, redo: () => { document.restoreSnapshot(after); document.renderEmitter.fire(); } });
}

function buildTextToColumnsEdits(document, range, delimiter, mergeConsecutive = false, sheetIndex = 0, overwriteConfirmed = false) {
  const plan = planTextToColumns(document, range, delimiter, mergeConsecutive, sheetIndex);
  if (plan.conflicts.length && !overwriteConfirmed) {
    const error = new Error('텍스트 나누기 결과가 오른쪽 셀의 기존 내용을 바꿉니다. 계속하려면 덮어쓰기를 확인하세요.');
    error.needsConfirmation = true;
    error.conflicts = plan.conflicts.length;
    throw error;
  }
  return plan.edits;
}

function buildPasteEdits(matrix, row, column, range, styleMatrix, formulaMatrix, hiddenRows = [], kindMatrix, numberFormatMatrix) {
  if (!Array.isArray(matrix) || !Number.isInteger(row) || row < 0 || !Number.isInteger(column) || column < 0) throw new RangeError('붙여넣기 대상이 유효하지 않습니다.');
  const cellCount = matrix.reduce((count, values) => count + (Array.isArray(values) ? values.length : 0), 0);
  if (matrix.some(values => !Array.isArray(values)) || cellCount > 20000) throw new RangeError('한 번에 붙여넣을 수 있는 셀은 20,000개까지입니다.');
  if (styleMatrix && (!Array.isArray(styleMatrix) || styleMatrix.length !== matrix.length || styleMatrix.some((line, index) => !Array.isArray(line) || line.length !== matrix[index].length))) throw new RangeError('붙여넣기 셀 서식 범위가 유효하지 않습니다.');
  if (formulaMatrix && (!Array.isArray(formulaMatrix) || formulaMatrix.length !== matrix.length || formulaMatrix.some((line, index) => !Array.isArray(line) || line.length !== matrix[index].length))) throw new RangeError('붙여넣기 셀 수식 범위가 유효하지 않습니다.');
  if (kindMatrix && (!Array.isArray(kindMatrix) || kindMatrix.length !== matrix.length || kindMatrix.some((line, index) => !Array.isArray(line) || line.length !== matrix[index].length || line.some(kind => !['blank', 'string', 'number', 'boolean', 'error'].includes(kind))))) throw new RangeError('붙여넣기 셀 값 종류 범위가 유효하지 않습니다.');
  if (numberFormatMatrix && (!Array.isArray(numberFormatMatrix) || numberFormatMatrix.length !== matrix.length || numberFormatMatrix.some((line, index) => !Array.isArray(line) || line.length !== matrix[index].length || line.some(value => value !== null && !(Number.isInteger(value) && [9, 10, 14, 20, 21, 22].includes(value)) && (!Number.isInteger(value.numFmtId) || ![9, 10, 14, 20, 21, 22].includes(value.numFmtId) && (value.numFmtId < 164 || typeof value.formatCode !== 'string' || value.formatCode.length > 255)))))) throw new RangeError('붙여넣기 숫자 표시 범위가 유효하지 않습니다.');
  if (!cellCount) return [];
  const createEdit = (targetRow, targetColumn, sourceRow, sourceColumn) => {
    const edit = { row: targetRow, column: targetColumn, value: matrix[sourceRow][sourceColumn] };
    if (styleMatrix) { edit.styleIndex = styleMatrix[sourceRow][sourceColumn]?.styleIndex ?? null; edit.styleProperties = styleMatrix[sourceRow][sourceColumn]?.styleProperties ?? null; if (styleMatrix[sourceRow][sourceColumn]?.kind) edit.kind = styleMatrix[sourceRow][sourceColumn].kind; }
    if (kindMatrix) edit.kind = kindMatrix[sourceRow][sourceColumn];
    if (numberFormatMatrix?.[sourceRow]?.[sourceColumn] !== null && numberFormatMatrix?.[sourceRow]?.[sourceColumn] !== undefined) { const numberFormat = numberFormatMatrix[sourceRow][sourceColumn]; edit.numberFormatId = Number.isInteger(numberFormat) ? numberFormat : numberFormat.numFmtId; if (numberFormat.formatCode) edit.numberFormatCode = numberFormat.formatCode; edit.kind = 'number'; }
    const formula = formulaMatrix?.[sourceRow]?.[sourceColumn];
    if (formula?.formula) {
      edit.formula = translateFormulaReferences(formula.formula, targetRow - formula.row, targetColumn - formula.column);
      edit.value = '';
    }
    return edit;
  };
  if (cellCount === 1 && matrix.length === 1 && matrix[0].length === 1 && range) {
    const bounds = selectionBounds(range, '붙여넣기', 20000), area = (bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1);
    if (area > 1) {
      const edits = [];
      const hidden = new Set(hiddenRows);
      for (let targetRow = bounds.top; targetRow <= bounds.bottom; targetRow++) {
        if (hidden.has(targetRow)) continue;
        for (let targetColumn = bounds.left; targetColumn <= bounds.right; targetColumn++) edits.push(createEdit(targetRow, targetColumn, 0, 0));
      }
      return edits;
    }
  }
  if (range && matrix.length > 0 && matrix.every(values => values.length === matrix[0].length) && matrix[0].length > 0) {
    const selectionArea = (Math.abs(range.endRow - range.startRow) + 1) * (Math.abs(range.endColumn - range.startColumn) + 1);
    if (selectionArea <= 20000) {
      const bounds = selectionBounds(range, '붙여넣기', 20000), targetColumns = bounds.right - bounds.left + 1;
      const hidden = new Set(hiddenRows), visibleTargetRows = [];
      for (let targetRow = bounds.top; targetRow <= bounds.bottom; targetRow++) if (!hidden.has(targetRow)) visibleTargetRows.push(targetRow);
      const targetRows = visibleTargetRows.length;
      const sourceRows = matrix.length, sourceColumns = matrix[0].length, area = targetRows * targetColumns;
      if (area > 1 && cellCount > 1 && row === bounds.top && column === bounds.left && (targetRows % sourceRows !== 0 || targetColumns % sourceColumns !== 0)) throw new RangeError('클립보드 범위와 선택 영역의 크기가 맞지 않습니다.');
      if (row === bounds.top && column === bounds.left && area >= cellCount && targetRows % sourceRows === 0 && targetColumns % sourceColumns === 0) {
        const edits = [];
        for (let rowOffset = 0; rowOffset < targetRows; rowOffset++) for (let columnOffset = 0; columnOffset < targetColumns; columnOffset++) {
          edits.push(createEdit(visibleTargetRows[rowOffset], bounds.left + columnOffset, rowOffset % sourceRows, columnOffset % sourceColumns));
        }
        return edits;
      }
      if (row === bounds.top && column === bounds.left && matrix.length <= targetRows) {
        const edits = [];
        for (let rowOffset = 0; rowOffset < matrix.length; rowOffset++) for (let columnOffset = 0; columnOffset < matrix[rowOffset].length; columnOffset++) {
          edits.push(createEdit(visibleTargetRows[rowOffset], bounds.left + columnOffset, rowOffset, columnOffset));
        }
        return edits;
      }
    }
  }
  const edits = [];
  for (let rowOffset = 0; rowOffset < matrix.length; rowOffset++) {
    for (let columnOffset = 0; columnOffset < matrix[rowOffset].length; columnOffset++) edits.push(createEdit(row + rowOffset, column + columnOffset, rowOffset, columnOffset));
  }
  return edits;
}

function buildClearEdits(document, row, column, range, sheetIndex = 0, hiddenRows = []) {
  const source = range || { startRow: row, endRow: row, startColumn: column, endColumn: column };
  const { startRow, endRow, startColumn, endColumn } = source;
  if (![startRow, endRow, startColumn, endColumn].every(Number.isInteger) || Math.min(startRow, endRow, startColumn, endColumn) < 0) throw new RangeError('지우기 범위가 유효하지 않습니다.');
  const bounds = { top: Math.min(startRow, endRow), bottom: Math.max(startRow, endRow), left: Math.min(startColumn, endColumn), right: Math.max(startColumn, endColumn) };
  const maximumRows = document.kind === 'xlsx' ? XLSX_LIMITS.rows : document.kind === 'csv' ? CSV_LIMIT_ROWS : 0;
  const maximumColumns = document.kind === 'xlsx' ? XLSX_LIMITS.columns : document.kind === 'csv' ? CSV_LIMIT_COLUMNS : 0;
  if (bounds.bottom >= maximumRows || bounds.right >= maximumColumns) throw new RangeError('지우기 범위가 시트 경계를 벗어났습니다.');
  const area = (bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1);
  if (!Number.isSafeInteger(area)) throw new RangeError('지우기 범위가 유효하지 않습니다.');
  if (area > 10000) {
    const candidates = new Set();
    if (document.kind === 'csv') {
      for (let currentRow = bounds.top; currentRow <= bounds.bottom; currentRow++) {
        const cells = document.state.rows[currentRow] || [];
        for (let currentColumn = bounds.left; currentColumn <= Math.min(bounds.right, cells.length - 1); currentColumn++) candidates.add(currentRow + ':' + currentColumn);
      }
    } else if (document.kind === 'xlsx') {
      for (const cell of document.state.sheets[sheetIndex]?.cells.values() || []) if (cell.row >= bounds.top && cell.row <= bounds.bottom && cell.column >= bounds.left && cell.column <= bounds.right) candidates.add(cell.row + ':' + cell.column);
      for (const [key] of document.drafts) {
        const [draftSheet, address] = key.split(':');
        if (Number(draftSheet) !== sheetIndex) continue;
        const coordinate = parseAddress(address);
        if (coordinate.row >= bounds.top && coordinate.row <= bounds.bottom && coordinate.column >= bounds.left && coordinate.column <= bounds.right) candidates.add(coordinate.row + ':' + coordinate.column);
      }
      for (const [key] of document.formulaDrafts || []) {
        const [draftSheet, address] = key.split(':');
        if (Number(draftSheet) !== sheetIndex) continue;
        const coordinate = parseAddress(address);
        if (coordinate.row >= bounds.top && coordinate.row <= bounds.bottom && coordinate.column >= bounds.left && coordinate.column <= bounds.right) candidates.add(coordinate.row + ':' + coordinate.column);
      }
    }
    const edits = [];
    for (const key of candidates) {
      const [currentRow, currentColumn] = key.split(':').map(Number), cell = document.getCell(currentRow, currentColumn, sheetIndex);
      if (!cell || ((cell.value ?? '') === '' && !isFormulaCell(cell) && cell.kind !== 'error')) continue;
      const edit = { row: currentRow, column: currentColumn, value: '' };
      if (document.kind === 'xlsx' && (isFormulaCell(cell) || cell.kind === 'error')) edit.formula = null;
      edits.push(edit);
      if (edits.length > 20000) throw new RangeError('한 번에 지울 수 있는 값은 20,000개까지입니다.');
    }
    return excludeHiddenRowEdits(edits, hiddenRows);
  }
  selectionBounds(source, '지우기');
  const edits = buildPasteEdits([['']], bounds.top, bounds.left, {
    startRow: bounds.top, endRow: bounds.bottom, startColumn: bounds.left, endColumn: bounds.right
  });
  if (document.kind === 'xlsx') for (const edit of edits) {
    const cell = document.getCell(edit.row, edit.column, sheetIndex);
    if (isFormulaCell(cell) || cell?.kind === 'error') edit.formula = null;
  }
  return excludeHiddenRowEdits(edits, hiddenRows);
}

const searchResultCaches = new WeakMap();

function findMatches(document, query, sheetIndex, scope = 'sheet', after = null) {
  scope = document.kind === 'xlsx' && scope === 'workbook' ? 'workbook' : 'sheet';
  const needle = query.toLocaleLowerCase();
  if (!needle) return { type: 'search', query, total: 0, matches: [] };
  const pageSize = 100;
  const compare = (left, right) => (left.sheetIndex ?? sheetIndex) - (right.sheetIndex ?? sheetIndex) || left.row - right.row || left.column - right.column;
  const afterKey = after && Number.isInteger(after.row) && Number.isInteger(after.column)
    ? { row: after.row, column: after.column, sheetIndex: scope === 'workbook' && Number.isInteger(after.sheetIndex) ? after.sheetIndex : sheetIndex }
    : null;
  const revision = Number.isInteger(document.renderRevision) ? document.renderRevision : null;
  const cacheKey = JSON.stringify([needle, scope, sheetIndex]);
  const cached = revision === null ? null : searchResultCaches.get(document);
  let result = cached?.revision === revision && cached.key === cacheKey ? cached.result : null;
  if (!result) {
    const matches = [];
    let cacheable = revision !== null;
    const calculator = document.kind === 'xlsx' ? createFormulaCalculator(document) : null;
    const sheetIndexes = scope === 'workbook' ? document.state.sheets.map((_, index) => index) : [sheetIndex];
    const visit = (currentSheetIndex, row, column, base) => {
      const cell = document.kind === 'xlsx' ? document.getCell(row, column, currentSheetIndex) : null;
      if (cell?.formula && /\b(?:NOW|TODAY|RAND|RANDBETWEEN|RANDARRAY|INDIRECT|OFFSET|CELL|INFO)\s*\(/i.test(cell.formula)) cacheable = false;
      const value = document.kind === 'csv' ? (document.drafts.get(`0:${row}:${column}`) ?? base.value)
        : isFormulaCell(cell) ? calculator.value(currentSheetIndex, row, column)
          : (document.drafts.get(`${currentSheetIndex}:${base.address}`) ?? base.value);
      if (String(value).toLocaleLowerCase().includes(needle)) matches.push({ row, column, address: base.address || addressFor(row + 1, column + 1), ...(scope === 'workbook' ? { sheetIndex: currentSheetIndex, sheetName: document.state.sheets[currentSheetIndex].name } : {}) });
    };
    for (const currentSheetIndex of sheetIndexes) {
      if (document.kind === 'csv') document.state.rows?.forEach((cells, row) => cells.forEach((cell, column) => visit(0, row, column, cell)));
      else {
        const sheet = document.state.sheets[currentSheetIndex];
        for (const cell of sheet?.cells.values() || []) visit(currentSheetIndex, cell.row, cell.column, cell);
        for (const [key] of document.drafts) {
          const [draftSheet, address] = key.split(':');
          if (Number(draftSheet) !== currentSheetIndex || sheet?.cells.has(address)) continue;
          const { row, column } = parseAddress(address);
          visit(currentSheetIndex, row, column, { row, column, address, value: '' });
        }
        for (const [key] of document.formulaDrafts || []) {
          const [draftSheet, address] = key.split(':');
          if (Number(draftSheet) !== currentSheetIndex || sheet?.cells.has(address) || document.drafts.has(`${currentSheetIndex}:${address}`)) continue;
          const { row, column } = parseAddress(address);
          visit(currentSheetIndex, row, column, { row, column, address, value: '' });
        }
      }
    }
    matches.sort(compare);
    result = { matches, total: matches.length };
    if (cacheable) searchResultCaches.set(document, { key: cacheKey, revision, result });
    else searchResultCaches.delete(document);
  }
  let start = 0;
  if (afterKey) {
    let low = 0, high = result.matches.length;
    while (low < high) { const middle = (low + high) >> 1; if (compare(result.matches[middle], afterKey) <= 0) low = middle + 1; else high = middle; }
    start = low;
  }
  const wrapped = Boolean(afterKey && result.total && start >= result.total);
  if (wrapped) start = 0;
  return { type: 'search', query, total: result.total, matches: result.matches.slice(start, start + pageSize), sheetIndex, ...(afterKey ? { wrapped } : {}), ...(scope === 'workbook' ? { scope } : {}) };
}

function buildReplaceEdits(document, query, replacement, sheetIndex = 0, target = null, allCells = true) {
  const text = query;
  if (!text) return { edits: [], replaced: 0, skippedFormulas: 0, target: null };
  if (typeof replacement !== 'string' || replacement.length > MAX_CELL_CHARS) throw new RangeError('바꿀 값은 100,000자 이하여야 합니다.');
  const needle = text.toLocaleLowerCase(), escaped = text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const expression = new RegExp(escaped, 'giu');
  const candidates = new Map(), add = (row, column, cell) => candidates.set(`${row}:${column}`, { row, column, cell });
  if (document.kind === 'csv') {
    for (const [row, cells] of (document.state.rows || []).entries()) for (const [column, cell] of cells.entries()) add(row, column, cell);
  } else {
    const sheet = document.state.sheets[sheetIndex];
    if (!sheet) throw new RangeError('수정할 시트를 찾을 수 없습니다.');
    for (const cell of sheet.cells.values()) add(cell.row, cell.column, cell);
    for (const [key] of document.drafts || []) {
      const [draftSheet, address] = key.split(':');
      if (Number(draftSheet) !== sheetIndex || sheet.cells.has(address)) continue;
      const { row, column } = parseAddress(address);
      add(row, column, { row, column, address, value: '' });
    }
    for (const [key] of document.formulaDrafts || []) {
      const [draftSheet, address] = key.split(':');
      if (Number(draftSheet) !== sheetIndex || sheet.cells.has(address) || document.drafts?.has(`${sheetIndex}:${address}`)) continue;
      const { row, column } = parseAddress(address);
      add(row, column, { row, column, address, value: '' });
    }
  }
  const calculator = document.kind === 'xlsx' ? createFormulaCalculator(document) : null;
  const ordered = [...candidates.values()].sort((left, right) => left.row - right.row || left.column - right.column);
  const edits = [];
  let replaced = 0, skippedFormulas = 0, matchedTarget = null;
  for (const candidate of ordered) {
    const { row, column, cell: base } = candidate;
    if (target && (target.row !== row || target.column !== column)) continue;
    const address = base.address || addressFor(row + 1, column + 1);
    const draftKey = document.kind === 'csv' ? `0:${row}:${column}` : `${sheetIndex}:${address}`;
    const cell = document.kind === 'xlsx' ? document.getCell(row, column, sheetIndex) : null;
    const formulaDraft = document.formulaDrafts?.has(draftKey) || false;
    const formulaCell = document.kind === 'xlsx' && (isFormulaCell(cell) || formulaDraft);
    const value = document.kind === 'xlsx' && formulaCell
      ? calculator.value(sheetIndex, row, column)
      : (document.drafts?.get(draftKey) ?? base.value);
    if (!String(value ?? '').toLocaleLowerCase().includes(needle)) {
      if (target) break;
      continue;
    }
    matchedTarget = { row, column, address };
    if (formulaCell) skippedFormulas++;
    else {
      let occurrences = 0;
      const nextValue = String(value ?? '').replace(expression, () => { occurrences++; return replacement; });
      if (nextValue.length > MAX_CELL_CHARS) throw new RangeError('바꾼 값은 100,000자 이하여야 합니다.');
      replaced += occurrences;
      if (nextValue !== String(value ?? '')) edits.push({ row, column, value: nextValue, ...(document.kind === 'xlsx' ? { formula: null, kind: 'string' } : {}) });
      if (edits.length > 20000) throw new RangeError('한 번에 바꿀 수 있는 셀은 20,000개까지입니다.');
    }
    if (target || !allCells) break;
  }
  return { edits, replaced, skippedFormulas, target: matchedTarget };
}

function buildScopedReplaceEdits(document, query, replacement, sheetIndex = 0, target = null, allCells = true, scope = 'sheet') {
  if (scope !== 'workbook' || document.kind !== 'xlsx') {
    const result = buildReplaceEdits(document, query, replacement, sheetIndex, target, allCells);
    return { ...result, sheetEdits: result.edits.length ? [{ sheetIndex, edits: result.edits }] : [] };
  }
  const allSheetIndexes = document.state.sheets.map((_, index) => index);
  const sheetIndexes = target ? [target.sheetIndex ?? sheetIndex] : allCells ? allSheetIndexes : [sheetIndex, ...allSheetIndexes.filter(index => index !== sheetIndex)];
  const sheetEdits = []; let replaced = 0, skippedFormulas = 0, matchedTarget = null, editedCells = 0;
  for (const currentSheetIndex of sheetIndexes) {
    if (!Number.isInteger(currentSheetIndex) || !document.state.sheets[currentSheetIndex]) throw new RangeError('수정할 시트를 찾을 수 없습니다.');
    const localTarget = target ? { row: target.row, column: target.column } : null;
    const result = buildReplaceEdits(document, query, replacement, currentSheetIndex, localTarget, allCells);
    replaced += result.replaced;
    skippedFormulas += result.skippedFormulas;
    if (result.target) matchedTarget = { ...result.target, sheetIndex: currentSheetIndex };
    if (result.edits.length) {
      editedCells += result.edits.length;
      if (editedCells > 20000) throw new RangeError('통합 문서에서 한 번에 바꿀 수 있는 셀은 20,000개까지입니다.');
      sheetEdits.push({ sheetIndex: currentSheetIndex, edits: result.edits });
    }
    if (!allCells && result.target) break;
  }
  return { sheetEdits, edits: sheetEdits.flatMap(item => item.edits), replaced, skippedFormulas, target: matchedTarget };
}

function renderPageBreakControls(language) {
  return '<fieldset class="print-controls"><legend>' + message('page-breaks', language) + '</legend>' + [['breakAddRow','break-add-row'],['breakRemoveRow','break-remove-row'],['breakAddColumn','break-add-column'],['breakRemoveColumn','break-remove-column'],['breakReset','break-reset']].map(([id,key]) => '<button id="' + id + '" type="button">' + message(key, language) + '</button>').join('') + '<span id="pageBreakStatus" role="status" aria-live="polite"></span></fieldset>';
}

function renderHeaderFooterControls(language) {
  const checkbox = (id, key) => '<label><input id="' + id + '" type="checkbox">' + message(key, language) + '</label>';
  const fields = ['header', 'footer'].flatMap(part => ['left', 'center', 'right'].map(side => '<label for="hf-' + part + '-' + side + '">' + message('hf-' + part + '-' + side, language) + '<textarea id="hf-' + part + '-' + side + '" rows="2" maxlength="255"></textarea></label>')).join('');
  const tokens = [['&P','hf-page'],['&N','hf-pages'],['&F','hf-file'],['&A','hf-sheet'],['&D','hf-date'],['&T','hf-time'],['&Z','hf-path'],['&&','hf-amp']];
  return '<fieldset id="hfPanel"><legend>' + message('header-footer', language) + '</legend><div class="print-controls"><label for="hfVariant">' + message('hf-variant', language) + '</label><select id="hfVariant">' + ['odd','even','first'].map(variant => '<option value="' + variant + '">' + message('hf-' + variant, language) + '</option>').join('') + '</select>' + checkbox('hfDifferentFirst','hf-different-first') + checkbox('hfDifferentEven','hf-different-even') + checkbox('hfScale','hf-scale') + checkbox('hfAlign','hf-align') + '</div><div class="hf-fields">' + fields + '</div><div class="print-controls"><label for="hfToken">' + message('hf-token', language) + '</label><select id="hfToken">' + tokens.map(([value,key]) => '<option value="' + value.replaceAll('&','&amp;') + '">' + message(key,language) + '</option>').join('') + '</select><button id="hfInsert" type="button">' + message('hf-insert',language) + '</button><button id="hfApply" type="button">' + message('hf-apply',language) + '</button><button id="hfClear" type="button">' + message('hf-clear',language) + '</button><span id="hfStatus" role="status" aria-live="polite"></span></div><p>' + message('hf-hint',language).replaceAll('&','&amp;') + '</p></fieldset>';
}

function renderFilterConditionControls(language) {
  const operators = ['equal', 'notEqual', 'lessThan', 'lessThanOrEqual', 'greaterThan', 'greaterThanOrEqual', 'contains', 'notContains', 'beginsWith', 'endsWith'];
  const condition = index => '<div class="filter-condition-row"><select id="filterOperator' + index + '" aria-label="' + message('filter-operator-' + index, language) + '">' + operators.map(operator => '<option value="' + operator + '">' + message('filter-op-' + operator, language) + '</option>').join('') + '</select><textarea rows="1" id="filterConditionValue' + index + '" maxlength="32767" aria-label="' + message('filter-value-' + index, language) + '" placeholder="' + message('filter-condition-placeholder', language) + '"></textarea></div>';
  return '<label>' + message('filter-mode', language) + '<select id="filterMode"><option value="values">' + message('filter-mode-values', language) + '</option><option value="conditions">' + message('filter-mode-conditions', language) + '</option><option value="top10">' + message('filter-mode-top', language) + '</option><option value="dates">' + message('filter-mode-dates', language) + '</option><option value="dynamic">' + message('filter-mode-dynamic', language) + '</option><option value="color">' + message('filter-mode-color', language) + '</option><option value="icons">' + message('filter-mode-icons', language) + '</option></select></label><fieldset id="filterConditionControls" hidden><legend>' + message('filter-mode-conditions', language) + '</legend>' + condition(1) + '<label><input id="filterUseSecond" type="checkbox">' + message('filter-second-condition', language) + '</label><select id="filterLogic" disabled><option value="and">' + message('filter-logic-and', language) + '</option><option value="or">' + message('filter-logic-or', language) + '</option></select><fieldset id="filterSecondCondition" disabled>' + condition(2) + '</fieldset><p>' + message('filter-condition-help', language) + '</p></fieldset><fieldset id="filterTopControls" hidden><legend>' + message('filter-mode-top', language) + '</legend><label>' + message('filter-top-direction', language) + '<select id="filterTopDirection"><option value="top">' + message('filter-top', language) + '</option><option value="bottom">' + message('filter-bottom', language) + '</option></select></label><label>' + message('filter-top-value', language) + '<input id="filterTopValue" type="number" value="10" min="1" max="500" step="1"></label><label>' + message('filter-top-unit', language) + '<select id="filterTopUnit"><option value="items">' + message('filter-top-items', language) + '</option><option value="percent">' + message('filter-top-percent', language) + '</option></select></label><p>' + message('filter-top-help', language) + '</p></fieldset><fieldset id="filterDateControls" hidden><legend>' + message('filter-mode-dates', language) + '</legend><label for="filterDateGroups">' + message('filter-date-groups', language) + '</label><textarea id="filterDateGroups" rows="4" maxlength="32767" placeholder="2026-10"></textarea><label><input id="filterDateBlanks" type="checkbox">' + message('filter-date-blanks', language) + '</label><label id="filterDateKeepLabel" hidden><input id="filterDateKeepValues" type="checkbox" checked><span id="filterDateValuesCount"></span></label><p>' + message('filter-date-help', language) + '</p></fieldset><fieldset id="filterDynamicControls" hidden><legend>' + message('filter-mode-dynamic', language) + '</legend><label for="filterDynamicType">' + message('filter-dynamic-type', language) + '</label><select id="filterDynamicType">' + DYNAMIC_FILTER_TYPES.map(type => '<option value="' + type + '">' + message('filter-dynamic-' + type, language) + '</option>').join('') + '</select><p>' + message('filter-dynamic-help', language) + '</p></fieldset><fieldset id="filterColorControls" hidden><legend>' + message('filter-mode-color', language) + '</legend><label>' + message('filter-color-target', language) + '<select id="filterColorTarget"><option value="fill">' + message('filter-color-fill', language) + '</option><option value="font">' + message('filter-color-font', language) + '</option></select></label><label>' + message('filter-color-value', language) + '<select id="filterColorValue"></select></label><label id="filterColorCustomLabel" hidden>' + message('filter-color-custom', language) + '<input id="filterColorCustom" type="color" value="#FF0000"></label><span id="filterColorPreview" style="display:inline-block;padding:2px 8px;border:1px solid currentColor">Aa</span><p id="filterColorLimit" hidden>' + message('filter-color-limit', language) + '</p><p>' + message('filter-color-help', language) + '</p></fieldset><fieldset id="filterIconControls" hidden><legend>' + message('filter-mode-icons', language) + '</legend><label>' + message('filter-icon-choice', language) + '<select id="filterIconValue"></select></label><p>' + message('filter-icon-help', language) + '</p></fieldset>';
}

function renderPrintControls(language) {
  const select = (id, label, options) => '<label for="' + id + '">' + message(label, language) + '</label><select id="' + id + '">' + options.map(([value, key]) => '<option value="' + value + '">' + message(key, language) + '</option>').join('') + '</select>';
  const number = (id, label, min, max, step = 1) => '<label for="' + id + '">' + message(label, language) + '</label><input id="' + id + '" type="number" min="' + min + '" max="' + max + '" step="' + step + '" style="width:78px">';
  return '<details id="printPanel"><summary>' + message('page-setup', language) + '</summary><div class="print-controls">'
    + '<label><input id="printIncludeCharts" type="checkbox" checked>' + message('print-include-charts', language) + '</label>'
    + select('printPaper', 'paper-size', [[9, 'paper-a4'], [8, 'paper-a3'], [1, 'paper-letter']])
    + select('printOrientation', 'print-orientation', [['portrait', 'print-portrait'], ['landscape', 'print-landscape']])
    + select('printPageOrder', 'page-order', [['downThenOver', 'page-order-down'], ['overThenDown', 'page-order-over']])
    + number('printScale', 'print-scale', 10, 400)
    + '<label><input id="printUseFirstPage" type="checkbox">' + message('custom-first-page', language) + '</label>' + number('printFirstPage', 'first-page-number', 0, 32767)
    + '<label><input id="printFit" type="checkbox">' + message('print-fit', language) + '</label>'
    + number('printFitWidth', 'print-fit-width', 0, 10) + number('printFitHeight', 'print-fit-height', 0, 10)
    + '<label><input id="printGridlines" type="checkbox">' + message('print-gridlines', language) + '</label><label><input id="printHeadings" type="checkbox">' + message('print-headings', language) + '</label></div><fieldset class="print-controls"><legend>' + message('print-margins', language) + '</legend>'
    + ['left', 'right', 'top', 'bottom', 'header', 'footer'].map(side => number('printMargin-' + side, 'margin-' + side, 0, 5, .01)).join('')
    + '</fieldset><button id="applyPageSettings" type="button">' + message('apply-page-settings', language) + '</button><div class="print-controls"><button id="setPrintArea" type="button">' + message('set-print-area', language) + '</button><button id="clearPrintArea" type="button">' + message('clear-print-area', language) + '</button><span id="printAreaStatus" role="status" aria-live="polite"></span></div><fieldset class="print-controls"><legend>' + message('apply-print-titles', language) + '</legend><label for="printTitleRows">' + message('repeat-rows', language) + '</label><input id="printTitleRows" type="text" maxlength="32" placeholder="1:3" style="width:100px"><label for="printTitleColumns">' + message('repeat-columns', language) + '</label><input id="printTitleColumns" type="text" maxlength="32" placeholder="A:B" style="width:100px"><span>' + message('repeat-titles-hint', language) + '</span><button id="applyPrintTitles" type="button">' + message('apply-print-titles', language) + '</button><button id="clearPrintTitles" type="button">' + message('clear-print-titles', language) + '</button><span id="printTitleStatus" role="status" aria-live="polite"></span></fieldset>' + renderHeaderFooterControls(language) + renderPageBreakControls(language) + '</details>';
}

function renderTableCreationControls(language) {
  return `<label for="newTableName">${message('new-table-name',language)}</label><input id="newTableName" type="text" maxlength="255" placeholder="${message('new-table-auto-name',language)}" title="${message('new-table-title',language)}"><button id="createTable" type="button" title="${message('new-table-title',language)}">${message('new-table',language)}</button>` + renderTableStyleControls(language);
}

function renderTableStyleControls(language) {
  const groups = [['Light', 21], ['Medium', 28], ['Dark', 11]];
  const options = groups.map(([family, count]) => '<optgroup label="' + message('table-style-' + family.toLowerCase(), language) + '">' + Array.from({ length: count }, (_, index) => '<option value="TableStyle' + family + (index + 1) + '">' + family + ' ' + (index + 1) + '</option>').join('') + '</optgroup>').join('');
  return '<label for="tableStyleName">' + message('table-style-name', language) + '</label><select id="tableStyleName" style="max-width:180px" title="' + message('table-style-title', language) + '"><option value="">' + message('table-style-none', language) + '</option>' + options + '</select>' +
    '<label><input id="tableStyleFirstColumn" type="checkbox">' + message('table-style-first-column', language) + '</label>' +
    '<label><input id="tableStyleLastColumn" type="checkbox">' + message('table-style-last-column', language) + '</label>' +
    '<label><input id="tableStyleRows" type="checkbox">' + message('table-style-row-stripes', language) + '</label>' +
    '<label><input id="tableStyleColumns" type="checkbox">' + message('table-style-column-stripes', language) + '</label>' +
    '<button id="applyTableStyle" type="button" title="' + message('table-style-title', language) + '">' + message('table-style-apply', language) + '</button><span id="tableStyleNote" role="status" aria-live="polite"></span>';
}

function buildSortDialogBootstrap(language) {
  const labels = JSON.stringify({
    custom: message('sort-custom', language), title: message('sort-dialog-title', language), level: message('sort-level', language),
    by: message('sort-by', language), thenBy: message('sort-then-by', language), order: message('sort-order', language),
    ascending: message('sort-ascending', language), descending: message('sort-descending', language), add: message('sort-add-level', language),
    remove: message('sort-remove-level', language), apply: message('sort-apply', language), cancel: message('sort-cancel', language),
    header: message('sort-has-header', language), expand: message('sort-expand-selection', language), expandLimitHint: message('sort-expand-limit-hint', language), duplicate: message('sort-invalid-duplicate', language),
    limit: message('sort-level-limit', language), column: message('aria-column', language)
  });
  return `(()=>{
    const labels=${labels};
    const button=document.createElement('button');button.id='customSort';button.type='button';button.textContent=labels.custom;button.title=labels.title;
    const toolbar=document.getElementById('toolbar'),anchor=document.getElementById('filterValuesButton');
    if(anchor?.parentElement)anchor.parentElement.insertBefore(button,anchor);else toolbar?.append(button);
    const panel=document.createElement('div');panel.id='sortPanel';panel.setAttribute('role','dialog');panel.setAttribute('aria-label',labels.title);panel.hidden=true;
    Object.assign(panel.style,{position:'fixed',zIndex:'30',top:'70px',right:'12px',maxHeight:'65vh',minWidth:'280px',maxWidth:'min(620px, calc(100vw - 24px))',overflow:'auto',padding:'10px',background:'var(--vscode-editorWidget-background, #fff)',color:'var(--vscode-editorWidget-foreground, #202124)',border:'1px solid var(--vscode-widget-border, #aaa)',boxShadow:'0 3px 12px #0003'});
    const title=document.createElement('strong');title.textContent=labels.title;panel.append(title);
    const levels=document.createElement('div');levels.id='sortLevels';Object.assign(levels.style,{display:'grid',gap:'6px',marginTop:'8px'});panel.append(levels);
    const add=document.createElement('button');add.id='sortAddLevel';add.type='button';add.textContent=labels.add;panel.append(add);
    const options=document.createElement('div');options.className='sort-options';Object.assign(options.style,{display:'flex',flexWrap:'wrap',gap:'10px',marginTop:'8px'});panel.append(options);
    const headerLabel=document.createElement('label'),header=document.createElement('input'),headerText=document.createElement('span');header.type='checkbox';header.id='sortDialogHasHeader';headerText.textContent=labels.header;headerLabel.append(header);headerLabel.append(headerText);options.append(headerLabel);
    const expandLabel=document.createElement('label'),expand=document.createElement('input'),expandText=document.createElement('span');expand.type='checkbox';expand.id='sortDialogExpandSelection';expandText.textContent=labels.expand;expandLabel.append(expand);expandLabel.append(expandText);options.append(expandLabel);
    const expandHint=document.createElement('small');expandHint.id='sortExpandLimitHint';expandHint.textContent=labels.expandLimitHint;expandHint.style.flexBasis='100%';expandHint.style.opacity='.8';options.append(expandHint);
    const error=document.createElement('div');error.id='sortError';error.setAttribute('role','status');panel.append(error);
    const actions=document.createElement('div');actions.className='sort-actions';Object.assign(actions.style,{display:'flex',gap:'6px',marginTop:'8px'});panel.append(actions);
    const apply=document.createElement('button');apply.id='sortApply';apply.type='button';apply.textContent=labels.apply;actions.append(apply);
    const cancel=document.createElement('button');cancel.id='sortCancel';cancel.type='button';cancel.textContent=labels.cancel;actions.append(cancel);
    document.body.append(panel);
    let criteria=[],sortRange=null;
    const columnLabel=number=>{let value=number+1,name='';while(value){const remainder=(value-1)%26;name=String.fromCharCode(65+remainder)+name;value=Math.floor((value-1)/26)}return name};
    const addLevel=(initialColumn)=>{
      if(criteria.length>=64){error.textContent=labels.limit;return}
      const index=criteria.length,row=document.createElement('div');row.className='sort-level';Object.assign(row.style,{display:'grid',gridTemplateColumns:'minmax(80px, 1fr) minmax(100px, 1.5fr) minmax(90px, 1fr) auto',gap:'6px',alignItems:'center'});
      const label=document.createElement('span');label.textContent=index===0?labels.by:labels.thenBy;row.append(label);
      const column=document.createElement('select');column.setAttribute('aria-label',labels.column+' '+(index===0?labels.by:labels.thenBy));
      const top=Math.min(sortRange.startRow,sortRange.endRow),left=Math.min(sortRange.startColumn,sortRange.endColumn),right=Math.max(sortRange.startColumn,sortRange.endColumn);
      for(let columnIndex=left;columnIndex<=right;columnIndex++){
        const option=document.createElement('option');option.value=String(columnIndex);
        const cell=document.querySelector('[data-row="'+top+'"][data-column="'+columnIndex+'"]');
        const heading=cell?.__jinSheetData?.value;
        option.textContent=columnLabel(columnIndex)+(heading!==undefined&&heading!==''?' — '+String(heading):'');column.append(option);
      }
      const used=new Set(criteria.map(item=>Number(item.column.value)));
      const wanted=Number.isInteger(initialColumn)&&initialColumn>=left&&initialColumn<=right&&!used.has(initialColumn)?initialColumn:Array.from({length:right-left+1},(_,offset)=>left+offset).find(item=>!used.has(item));
      if(wanted!==undefined)column.value=String(wanted);
      row.append(column);
      const direction=document.createElement('select');direction.setAttribute('aria-label',labels.order);
      for(const [value,text] of [['ascending',labels.ascending],['descending',labels.descending]]){const option=document.createElement('option');option.value=value;option.textContent=text;direction.append(option)}
      direction.value='ascending';
      row.append(direction);
      const remove=document.createElement('button');remove.type='button';remove.textContent=labels.remove;remove.setAttribute('aria-label',labels.remove+' '+(index+1));remove.disabled=index===0;
      remove.onclick=()=>{if(criteria.length<=1)return;criteria=criteria.filter(item=>item.row!==row);row.remove();criteria.forEach((item,level)=>{item.label.textContent=level===0?labels.by:labels.thenBy;item.column.setAttribute('aria-label',labels.column+' '+(level===0?labels.by:labels.thenBy));item.remove.disabled=level===0;item.remove.setAttribute('aria-label',labels.remove+' '+(level+1))});error.textContent=''};
      row.append(remove);levels.append(row);criteria.push({row,label,column,direction,remove});add.disabled=criteria.length>=Math.min(64,right-left+1);
    };
    const close=()=>{panel.hidden=true;button.focus()};
    button.onclick=()=>{
      if(!model||model.readOnly)return;
      sortRange={...selectedRange};criteria=[];levels.replaceChildren();error.textContent='';
      header.checked=document.getElementById('sortHasHeader')?.checked||false;expand.checked=document.getElementById('sortExpandSelection')?.checked!==false;
      add.disabled=false;addLevel(selected.column);panel.hidden=false;criteria[0]?.column.focus();
    };
    add.onclick=()=>{if(sortRange)addLevel(undefined)};
    apply.onclick=()=>{
      if(!criteria.length||!sortRange)return;
      const sortKeys=criteria.map(item=>({column:Number(item.column.value),direction:item.direction.value}));
      if(new Set(sortKeys.map(item=>item.column)).size!==sortKeys.length){error.textContent=labels.duplicate;return}
      if(!model||model.readOnly)return;
      document.getElementById('sortHasHeader').checked=header.checked;document.getElementById('sortExpandSelection').checked=expand.checked;
      vscode.postMessage({type:'sort-selection',sortKeys,range:sortRange,hasHeader:header.checked,expandSelection:expand.checked,sheetIndex:model.sheetIndex||0});close();
    };
    cancel.onclick=close;
    panel.onkeydown=event=>{if(event.key==='Escape'){event.preventDefault();event.stopPropagation();close()}else if(event.key==='Enter'&&event.target?.tagName==='SELECT'){event.preventDefault();apply.click()}};
  })();`;
}

function buildTextToColumnsBootstrap(language) {
  const labels = JSON.stringify({
    button: message('text-to-columns', language), title: message('text-to-columns-title', language), delimiter: message('text-to-columns-delimiter', language),
    comma: message('text-to-columns-comma', language), tab: message('text-to-columns-tab', language), semicolon: message('text-to-columns-semicolon', language),
    space: message('text-to-columns-space', language), other: message('text-to-columns-other', language), custom: message('text-to-columns-custom', language),
    merge: message('text-to-columns-merge', language), preview: message('text-to-columns-preview', language), empty: message('text-to-columns-preview-empty', language),
    collision: message('text-to-columns-collision', language), apply: message('text-to-columns-apply', language), cancel: message('text-to-columns-cancel', language)
  });
  return `(()=>{
    const labels=${labels},toolbar=document.getElementById('toolbar');
    const button=document.createElement('button');button.id='textToColumns';button.type='button';button.textContent=labels.button;button.title=labels.title;
    const anchor=document.getElementById('customSort');if(anchor?.parentElement)anchor.parentElement.insertBefore(button,anchor.nextSibling);else toolbar?.append(button);
    const panel=document.createElement('section');panel.id='textToColumnsPanel';panel.setAttribute('role','dialog');panel.setAttribute('aria-label',labels.title);panel.hidden=true;
    Object.assign(panel.style,{position:'fixed',zIndex:'30',top:'70px',right:'12px',maxHeight:'70vh',minWidth:'280px',maxWidth:'min(620px, calc(100vw - 24px))',overflow:'auto',padding:'10px',background:'var(--vscode-editorWidget-background,#fff)',color:'var(--vscode-editorWidget-foreground,#202124)',border:'1px solid var(--vscode-widget-border,#aaa)',boxShadow:'0 3px 12px #0003'});
    const heading=document.createElement('strong');heading.textContent=labels.title;panel.append(heading);
    const controls=document.createElement('div');Object.assign(controls.style,{display:'flex',flexWrap:'wrap',gap:'6px',alignItems:'center',marginTop:'8px'});panel.append(controls);
    const delimiterLabel=document.createElement('label');delimiterLabel.textContent=labels.delimiter;const mode=document.createElement('select');mode.id='textToColumnsDelimiterMode';for(const [value,text] of [['comma',labels.comma],['tab',labels.tab],['semicolon',labels.semicolon],['space',labels.space],['other',labels.other]]){const option=document.createElement('option');option.value=value;option.textContent=text;mode.append(option)}delimiterLabel.append(mode);controls.append(delimiterLabel);
    const custom=document.createElement('input');custom.id='textToColumnsCustomDelimiter';custom.maxLength=8;custom.placeholder=labels.custom;custom.hidden=true;controls.append(custom);
    const mergeLabel=document.createElement('label'),merge=document.createElement('input'),mergeText=document.createElement('span');merge.type='checkbox';merge.id='textToColumnsMerge';mergeText.textContent=labels.merge;mergeLabel.append(merge,mergeText);controls.append(mergeLabel);
    const previewTitle=document.createElement('div');previewTitle.textContent=labels.preview;previewTitle.style.marginTop='8px';panel.append(previewTitle);
    const status=document.createElement('div');status.id='textToColumnsStatus';status.setAttribute('role','status');status.style.margin='4px 0';panel.append(status);
    const preview=document.createElement('div');preview.id='textToColumnsPreview';preview.style.overflow='auto';panel.append(preview);
    const actions=document.createElement('div');Object.assign(actions.style,{display:'flex',gap:'6px',marginTop:'8px'});panel.append(actions);
    const apply=document.createElement('button');apply.id='textToColumnsApply';apply.type='button';apply.textContent=labels.apply;actions.append(apply);
    const cancel=document.createElement('button');cancel.id='textToColumnsCancel';cancel.type='button';cancel.textContent=labels.cancel;actions.append(cancel);document.body.append(panel);
    let targetRange=null,previewRequest=0,applyRequest=0,previewTimer=null;
    const currentDelimiter=()=>mode.value==='tab'?'\\t':mode.value==='semicolon'?';':mode.value==='space'?' ':mode.value==='other'?custom.value:',';
    const close=()=>{panel.hidden=true;button.focus();targetRange=null;apply.disabled=false};
    const requestPreview=()=>{clearTimeout(previewTimer);previewTimer=setTimeout(()=>{if(!targetRange)return;const requestId=++previewRequest;status.textContent='';preview.replaceChildren();vscode.postMessage({type:'text-to-columns-preview',requestId,range:targetRange,delimiter:currentDelimiter(),mergeConsecutive:merge.checked,sheetIndex:model?.sheetIndex||0})},80)};
    button.onclick=()=>{if(!model||model.readOnly||!['csv','xlsx'].includes(model.kind))return;commitFormulaEdit();targetRange={...selectedRange};mode.value='comma';custom.value='';custom.hidden=true;merge.checked=false;status.textContent='';preview.replaceChildren();panel.hidden=false;requestPreview();mode.focus()};
    mode.onchange=()=>{custom.hidden=mode.value!=='other';requestPreview()};custom.oninput=requestPreview;merge.onchange=requestPreview;
    apply.onclick=()=>{if(!targetRange||apply.disabled)return;const delimiter=currentDelimiter();if(!delimiter){status.textContent=labels.custom;custom.focus();return}clearTimeout(previewTimer);apply.disabled=true;applyRequest=++previewRequest;previewRequest=applyRequest;vscode.postMessage({type:'text-to-columns',requestId:applyRequest,range:targetRange,delimiter,mergeConsecutive:merge.checked,sheetIndex:model?.sheetIndex||0})};
    cancel.onclick=close;
    panel.onkeydown=event=>{if(event.key==='Escape'){event.preventDefault();event.stopPropagation();close()}else if(event.key==='Enter'&&event.target!==apply&&event.target!==cancel){event.preventDefault();apply.click()}};
    window.addEventListener('message',event=>{const data=event.data;if(data?.type==='text-to-columns-preview'){if(data.requestId!==previewRequest||data.sheetIndex!==model?.sheetIndex)return;preview.replaceChildren();if(!data.preview?.length){status.textContent=labels.empty;return}if(data.conflictCount)status.textContent=labels.collision;const table=document.createElement('table');table.style.borderCollapse='collapse';for(const row of data.preview){const tr=document.createElement('tr');for(const value of row){const td=document.createElement('td');td.textContent=value;td.style.cssText='border:1px solid var(--vscode-panel-border,#ccc);padding:3px 6px;white-space:pre-wrap';tr.append(td)}table.append(tr)}preview.append(table);return}if(data?.type==='text-to-columns-applied'){if(data.requestId!==applyRequest)return;close();return}if(data?.type==='text-to-columns-error'){if(data.requestId&&data.requestId!==applyRequest&&data.requestId!==previewRequest)return;apply.disabled=false;status.textContent=data.message||labels.empty;return}if(data?.type==='error'&&!panel.hidden){apply.disabled=false;status.textContent=data.message}});
    window.addEventListener('message',event=>{if(event.data?.type==='model'){button.disabled=Boolean(event.data.readOnly||!['csv','xlsx'].includes(event.data.kind))}});
  })();`;
}

function buildRemoveDuplicatesBootstrap(language) {
  const labels = JSON.stringify({
    button: message('remove-duplicates', language), title: message('remove-duplicates-title', language), header: message('remove-duplicates-header', language),
    columns: message('remove-duplicates-columns', language), preview: message('remove-duplicates-preview', language), empty: message('remove-duplicates-empty', language),
    apply: message('remove-duplicates-apply', language), cancel: message('remove-duplicates-cancel', language), applied: message('remove-duplicates-applied', language),
    rows: message('remove-duplicates-preview-count', language)
  });
  return `(()=>{
    const labels=${labels},toolbar=document.getElementById('toolbar');
    const button=document.createElement('button');button.id='removeDuplicates';button.type='button';button.textContent=labels.button;button.title=labels.title;
    const anchor=document.getElementById('textToColumns');if(anchor?.parentElement)anchor.parentElement.insertBefore(button,anchor.nextSibling);else toolbar?.append(button);
    const panel=document.createElement('section');panel.id='removeDuplicatesPanel';panel.setAttribute('role','dialog');panel.setAttribute('aria-label',labels.title);panel.hidden=true;
    Object.assign(panel.style,{position:'fixed',zIndex:'30',top:'70px',right:'12px',maxHeight:'70vh',minWidth:'280px',maxWidth:'min(620px, calc(100vw - 24px))',overflow:'auto',padding:'10px',background:'var(--vscode-editorWidget-background,#fff)',color:'var(--vscode-editorWidget-foreground,#202124)',border:'1px solid var(--vscode-widget-border,#aaa)',boxShadow:'0 3px 12px #0003'});
    const heading=document.createElement('strong');heading.textContent=labels.title;panel.append(heading);
    const headerLabel=document.createElement('label'),header=document.createElement('input'),headerText=document.createElement('span');header.type='checkbox';header.id='removeDuplicatesHeader';header.checked=true;headerText.textContent=' '+labels.header;headerLabel.append(header,headerText);panel.append(headerLabel);
    const columnTitle=document.createElement('div');columnTitle.textContent=labels.columns;columnTitle.style.marginTop='8px';panel.append(columnTitle);
    const columns=document.createElement('div');columns.id='removeDuplicatesColumns';Object.assign(columns.style,{display:'flex',flexWrap:'wrap',gap:'6px'});panel.append(columns);
    const previewTitle=document.createElement('div');previewTitle.textContent=labels.preview;previewTitle.style.marginTop='8px';panel.append(previewTitle);
    const status=document.createElement('div');status.id='removeDuplicatesStatus';status.setAttribute('role','status');status.style.margin='4px 0';panel.append(status);
    const actions=document.createElement('div');Object.assign(actions.style,{display:'flex',gap:'6px',marginTop:'8px'});panel.append(actions);
    const apply=document.createElement('button');apply.id='removeDuplicatesApply';apply.type='button';apply.textContent=labels.apply;actions.append(apply);
    const cancel=document.createElement('button');cancel.id='removeDuplicatesCancel';cancel.type='button';cancel.textContent=labels.cancel;actions.append(cancel);document.body.append(panel);
    let targetRange=null,requestId=0,activeRequest=0;
    const close=()=>{panel.hidden=true;button.focus();targetRange=null;apply.disabled=false};
    const requestPreview=()=>{if(!targetRange)return;const id=++requestId;activeRequest=id;status.textContent=labels.empty;apply.disabled=true;vscode.postMessage({type:'remove-duplicates-preview',requestId:id,range:targetRange,columns:Array.from(columns.children).flatMap(label=>Array.from(label.children).filter(input=>input.tagName==='INPUT'&&input.checked)).map(input=>Number(input.value)),hasHeader:header.checked,sheetIndex:model?.sheetIndex||0})};
    const columnName=index=>{let value=index+1,result='';while(value){const remainder=(value-1)%26;result=String.fromCharCode(65+remainder)+result;value=Math.floor((value-1)/26)}return result};
    button.onclick=()=>{if(!model||model.readOnly||!['csv','xlsx'].includes(model.kind))return;const selection={...selectedRange};commitFormulaEdit();targetRange={startRow:Math.min(selection.startRow,selection.endRow),endRow:Math.max(selection.startRow,selection.endRow),startColumn:Math.min(selection.startColumn,selection.endColumn),endColumn:Math.max(selection.startColumn,selection.endColumn)};columns.replaceChildren();for(let column=targetRange.startColumn;column<=targetRange.endColumn;column++){const label=document.createElement('label'),input=document.createElement('input');input.type='checkbox';input.value=String(column);input.checked=true;input.onchange=requestPreview;const text=document.createElement('span');text.textContent=' '+columnName(column);label.append(input,text);columns.append(label)}header.checked=true;status.textContent='';panel.hidden=false;requestPreview();header.focus()};
    header.onchange=requestPreview;
    apply.onclick=()=>{const selectedColumns=Array.from(columns.children).flatMap(label=>Array.from(label.children).filter(input=>input.tagName==='INPUT'&&input.checked)).map(input=>Number(input.value));if(!targetRange||!selectedColumns.length||apply.disabled)return;const id=++requestId;activeRequest=id;apply.disabled=true;vscode.postMessage({type:'remove-duplicates',requestId:id,range:targetRange,columns:selectedColumns,hasHeader:header.checked,sheetIndex:model?.sheetIndex||0})};
    cancel.onclick=close;panel.onkeydown=event=>{if(event.key==='Escape'){event.preventDefault();event.stopPropagation();close()}else if(event.key==='Enter'&&event.target!==apply&&event.target!==cancel){event.preventDefault();apply.click()}};
    window.addEventListener('message',event=>{const data=event.data;if(data?.type==='remove-duplicates-preview'){if(data.requestId!==activeRequest||data.sheetIndex!==model?.sheetIndex)return;status.textContent=labels.rows.replace('{count}',String(data.removedCount));apply.disabled=!data.removedCount;return}if(data?.type==='remove-duplicates-applied'){if(data.requestId!==activeRequest||data.sheetIndex!==model?.sheetIndex)return;document.getElementById('notice').textContent=labels.applied.replace('{count}',String(data.removedCount));close();return}if(data?.type==='remove-duplicates-error'){if(data.requestId!==activeRequest)return;apply.disabled=false;status.textContent=data.message||labels.empty;return}if(data?.type==='model'){button.disabled=Boolean(data.readOnly||!['csv','xlsx'].includes(data.kind));if(!panel.hidden&&targetRange&&!activeRequest)requestPreview()}});
  })();`;
}

function renderHtmlRaw(document, webview, extensionUri, language = 'en') {
  language = languageOf(language);
  const ui = JSON.stringify(messagesFor(language));
  const sortDialogBootstrap = buildSortDialogBootstrap(language);
  const textToColumnsBootstrap = buildTextToColumnsBootstrap(language);
  const removeDuplicatesBootstrap = buildRemoveDuplicatesBootstrap(language);
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
  .rowhead.manual-row-break::before{content:"";position:absolute;left:0;right:0;top:0;border-top:2px dashed var(--vscode-focusBorder,#007acc);pointer-events:none}.colhead.manual-column-break::before{content:"";position:absolute;left:0;top:0;bottom:0;border-left:2px dashed var(--vscode-focusBorder,#007acc);pointer-events:none}.hf-fields{display:grid;grid-template-columns:repeat(3,minmax(100px,1fr));gap:6px}.hf-fields label{display:flex;flex-direction:column}.hf-fields textarea{width:100%;min-height:46px;font:inherit;background:var(--vscode-input-background,#fff);color:var(--vscode-input-foreground,#202124);border:1px solid var(--vscode-input-border,#ccc)}#printPanel{padding:4px 8px;max-height:40vh;overflow:auto;border-bottom:1px solid var(--vscode-panel-border)}.print-controls{display:flex;flex-wrap:wrap;gap:6px;align-items:center}.print-controls input[type=checkbox]{width:auto}.print-controls fieldset{margin:4px 0}
  *{box-sizing:border-box}body{display:flex;flex-direction:column;height:100vh;margin:0;overflow:hidden;color:var(--vscode-foreground,#202124);background:var(--vscode-editor-background,#fff);font:13px var(--vscode-font-family)}#notice{padding:6px 10px;color:var(--vscode-descriptionForeground,#666);min-height:30px;background:var(--vscode-editor-background,#fff);border-bottom:1px solid var(--vscode-panel-border,#d6d6d6)}#toolbar{display:flex;flex:0 0 auto;max-height:35vh;overflow:auto;flex-wrap:wrap;gap:6px;align-items:center;padding:6px 8px;border-bottom:1px solid var(--vscode-panel-border,#d6d6d6);background:var(--vscode-editorGroupHeader-tabsBackground,var(--vscode-sideBar-background,#f3f3f3));color:var(--vscode-foreground,#202124)}button{font:inherit;color:var(--vscode-button-secondaryForeground,var(--vscode-button-foreground,#202124));background:var(--vscode-button-secondaryBackground,var(--vscode-button-background,#ededed));border:1px solid var(--vscode-button-border,var(--vscode-panel-border,#d0d0d0));border-radius:2px;padding:4px 9px;min-height:26px;cursor:pointer}button:hover:not(:disabled){background:var(--vscode-button-secondaryHoverBackground,var(--vscode-button-hoverBackground,#ddd))}button:focus-visible,input:focus-visible{outline:1px solid var(--vscode-focusBorder,#007acc);outline-offset:1px}button:disabled{opacity:.55;cursor:default}button[aria-selected=true]{color:var(--vscode-tab-activeForeground,var(--vscode-foreground,#202124));background:var(--vscode-tab-activeBackground,var(--vscode-editor-background,#fff));border-color:var(--vscode-focusBorder,#007acc)}input,select{font:inherit;color:var(--vscode-input-foreground,#202124);background:var(--vscode-input-background,#fff);border:1px solid var(--vscode-input-border,var(--vscode-panel-border,#c8c8c8));border-radius:2px;padding:4px 7px;min-height:26px}input::placeholder{color:var(--vscode-input-placeholderForeground,#777)}#cellValueGroup,#searchGroup,#replaceGroup{display:flex;align-items:center;gap:6px}#cellValueGroup{flex:1;min-width:180px}#cellValueGroup label,#searchGroup label,#replaceGroup label{color:var(--vscode-descriptionForeground,#666);white-space:nowrap}#formula{flex:1;min-width:120px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}#search{width:190px;height:26px}#searchScope{max-width:140px}#replaceValue{width:150px;height:26px}#filterPanel{position:fixed;z-index:20;top:70px;right:12px;max-height:65vh;min-width:220px;max-width:340px;overflow:auto;padding:10px;background:var(--vscode-editorWidget-background,#fff);color:var(--vscode-editorWidget-foreground,#202124);border:1px solid var(--vscode-widget-border,#aaa);box-shadow:0 3px 12px #0003}#filterPanel[hidden]{display:none}#filterConditionControls[hidden],#filterValueControls[hidden]{display:none}#filterEditor{border:0;padding:0;margin:0;min-width:0}#filterConditionControls fieldset{min-width:0}.filter-condition-row{display:flex;gap:6px;margin:6px 0}.filter-condition-row select{width:48%;min-width:0}.filter-condition-row textarea{flex:1;min-width:0;width:0;min-height:28px;max-height:96px;resize:vertical;font:inherit}#filterConditionControls p{font-size:12px;line-height:1.4}#filterPanel>label,#filterValuesList label{display:flex;gap:6px;align-items:center;padding:3px 0}#filterSearch{width:100%;margin:8px 0}#filterValuesList{max-height:45vh;overflow:auto;display:grid;gap:2px}#filterShowMore[hidden]{display:none}#filterPanel>div:last-child{display:flex;gap:5px;margin-top:8px}#searchCount{color:var(--vscode-descriptionForeground,#666)}#address{width:74px;min-width:74px;text-align:center}#goToStatus{color:var(--vscode-errorForeground,#c00)}#searchCount{min-width:3em}#clearSearch{white-space:nowrap}#clearSearch[hidden]{display:none}#grid{position:relative;flex:1 1 auto;min-height:0;overflow:auto;outline:none;background:#fff;color:#202124}#grid,#readerCanvas{scrollbar-color:var(--vscode-scrollbarSlider-background,rgba(100,100,100,.4)) transparent;scrollbar-width:thin}#grid::-webkit-scrollbar,#readerCanvas::-webkit-scrollbar{width:12px;height:12px}#grid::-webkit-scrollbar-track,#readerCanvas::-webkit-scrollbar-track{background:transparent}#grid::-webkit-scrollbar-thumb,#readerCanvas::-webkit-scrollbar-thumb{background:var(--vscode-scrollbarSlider-background,rgba(100,100,100,.4));background-clip:padding-box;border:3px solid transparent;border-radius:6px}#grid::-webkit-scrollbar-thumb:hover,#readerCanvas::-webkit-scrollbar-thumb:hover{background:var(--vscode-scrollbarSlider-hoverBackground,rgba(100,100,100,.7));background-clip:padding-box}#grid::-webkit-scrollbar-thumb:active,#readerCanvas::-webkit-scrollbar-thumb:active{background:var(--vscode-scrollbarSlider-activeBackground,rgba(100,100,100,.8));background-clip:padding-box}#canvas{position:relative}#chartPreview{flex:0 0 auto;max-height:min(230px,25vh);overflow:auto;padding:8px;display:flex;gap:12px;border-top:1px solid var(--vscode-panel-border,#d6d6d6)}#chartPreview[hidden]{display:none}.chart-card{margin:0;min-width:320px;max-width:640px;flex:1 1 420px}.chart-card figcaption{font-weight:600;margin-bottom:2px}.chart-card svg{display:block;width:100%;height:190px;color:var(--vscode-foreground,#202124);font:11px var(--vscode-font-family,Arial)}.chart-metadata{display:flex;flex-wrap:wrap;gap:6px;align-items:center;margin-top:4px}.chart-metadata label{display:flex;gap:4px;align-items:center;color:var(--vscode-descriptionForeground,#666)}.chart-metadata input{width:140px}.chart-metadata button{min-height:24px}.cell,.rowhead,.colhead{position:absolute;height:26px;padding:4px 6px;overflow:hidden;white-space:nowrap;text-overflow:ellipsis;border:1px solid var(--vscode-panel-border)}.cell{width:144px;background:#fff;color:#202124;border-color:#d6d6d6}.cell.numeric{text-align:right}.cell.hyperlink{color:var(--vscode-textLink-foreground,#006ab1);text-decoration:underline;cursor:pointer}.cell .fill-handle{position:absolute;right:0;bottom:0;width:6px;height:6px;min-height:0;padding:0;border:1px solid #fff;border-radius:0;background:var(--vscode-focusBorder,#007acc);cursor:crosshair;z-index:5}.cell.in-range{box-shadow:inset 0 0 0 9999px rgba(0,100,220,.14)}.rowhead,.colhead{background:#f3f3f3;color:#444;text-align:center;border-color:#d6d6d6}.rowhead{width:52px;z-index:1}.rowhead.has-outline-toggle{padding-right:16px}.outline-bracket{position:absolute;pointer-events:none;z-index:2;color:var(--vscode-descriptionForeground,#777)}.rowhead .outline-row-bracket{top:0;bottom:0;width:5px;border-left:1px solid currentColor}.rowhead .outline-bracket-start{border-top:1px solid currentColor}.rowhead .outline-bracket-end{border-bottom:1px solid currentColor}.colhead.has-outline-brackets{padding-bottom:8px}.colhead .outline-column-bracket{left:0;right:0;bottom:1px;border-top:1px solid currentColor}.colhead .outline-bracket-start:before,.colhead .outline-bracket-end:after{content:"";position:absolute;top:0;bottom:0;border-left:1px solid currentColor}.colhead .outline-bracket-start:before{left:0}.colhead .outline-bracket-end:after{right:0}.outline-toggle{position:absolute;display:inline-flex;align-items:center;justify-content:center;width:14px;height:14px;min-width:14px;min-height:14px;padding:0;border:1px solid var(--vscode-panel-border,#888);border-radius:2px;background:var(--vscode-editor-background,#fff);color:var(--vscode-foreground,#222);font:11px/12px var(--vscode-font-family);z-index:3;cursor:pointer}.rowhead .outline-toggle{right:2px;top:50%;transform:translateY(-50%)}.colhead .outline-toggle{left:2px;top:50%;transform:translateY(-50%)}.outline-levels{display:inline-flex;gap:2px;vertical-align:middle}.outline-levels[hidden]{display:none}.outline-level{min-width:20px;height:20px;padding:0 3px;line-height:1}.outline-toggle:focus{outline:1px solid var(--vscode-focusBorder,#007acc)}.colhead{top:0;z-index:1;user-select:none}.col-resize-handle{position:absolute;top:0;right:-3px;width:6px;height:100%;z-index:2;cursor:col-resize;touch-action:none}.col-resize-handle:focus{outline:1px solid var(--vscode-focusBorder)}.row-resize-handle{position:absolute;left:0;bottom:0;width:100%;height:6px;cursor:ns-resize;z-index:4;touch-action:none}.row-resize-handle:focus{outline:1px solid var(--vscode-focusBorder)}.corner{position:absolute;width:52px;height:30px;top:0;left:0;background:#e8e8e8;border:1px solid #d6d6d6;z-index:2}.cell textarea.cell-editor{display:block;height:16px;min-height:16px;padding:0;border:0;border-radius:0;outline:1px solid var(--vscode-focusBorder,#007acc);background:#fff;color:#202124;font:inherit;line-height:16px;white-space:pre-wrap;overflow:auto;resize:none;user-select:text;-webkit-user-select:text}.cell.selected{outline:1px solid var(--vscode-focusBorder);outline-offset:-1px}.cell:focus{outline:1px solid var(--vscode-focusBorder);outline-offset:-1px}.cell.search-match{outline:2px solid var(--vscode-editor-findMatchBorder,#e8a317);outline-offset:-2px}.cell.search-current{outline:2px solid var(--vscode-editor-findMatchHighlightBorder,var(--vscode-focusBorder,#007acc));outline-offset:-2px}.cell.cut-source{outline:1px dashed var(--vscode-editorWarning-foreground,#c90);outline-offset:-2px}#sheets{display:flex;gap:2px;padding:4px 8px;border-bottom:1px solid var(--vscode-panel-border,#d6d6d6);background:var(--vscode-editorGroupHeader-tabsBackground,var(--vscode-sideBar-background,#f3f3f3));color:var(--vscode-foreground,#202124)}#sheets button{background:transparent;border-color:transparent;border-bottom:2px solid transparent}#sheets button[aria-selected=true]{background:var(--vscode-tab-activeBackground,var(--vscode-editor-background,#fff));border-color:var(--vscode-panel-border,#d6d6d6);border-bottom-color:var(--vscode-focusBorder,#007acc)}body.document-view{height:100vh;overflow:hidden}#reader{display:none;height:calc(100vh - 30px);overflow:hidden;background:var(--vscode-editor-background,#fff);color:var(--vscode-foreground,#202124);flex-direction:column}#readerToolbar{position:relative;z-index:3;display:flex;flex:0 0 auto;gap:6px;align-items:center;padding:6px 10px;background:var(--vscode-editorGroupHeader-tabsBackground,var(--vscode-sideBar-background,#f3f3f3));color:var(--vscode-foreground,#202124);border-bottom:1px solid var(--vscode-panel-border,#ddd)}#readerToolbar #pdfStatus{margin-left:8px;color:var(--vscode-descriptionForeground,#666);font-variant-numeric:tabular-nums}#readerCanvas{flex:1 1 auto;height:auto;min-height:0;padding:16px;overflow:auto;background:#f3f3f3;color:#202124}#readerCanvas .docx-wrapper{background:#f3f3f3!important;padding:16px 0 0!important}#readerCanvas section.docx{background:#fff!important;color:#202124!important;box-shadow:0 1px 4px #0002!important;outline:1px solid #dedede;margin-bottom:16px}#readerCanvas .flyfish-pptx-scale-box,#readerCanvas .flyfish-pptx-content,#readerCanvas .slide{background:#fff}#readerCanvas.pptx-view{overflow-x:hidden}#readerCanvas.pptx-view.pptx-zoomed,#readerCanvas.docx-view.docx-zoomed{overflow-x:auto}#readerCanvas.docx-view{overflow-x:hidden}#readerCanvas a{pointer-events:none}#pdfPage{display:block;max-width:none;height:auto;margin:0 auto;background:white;box-shadow:0 2px 8px #0005}#pdfStatus{min-width:9em}#readerStyles{display:none}
  </style><script nonce="${nonce}" src="${renderers}"></script><script nonce="${nonce}" src="${pdfLibrary}"></script><script nonce="${nonce}">window.pdfjsLib=globalThis.pdfjsLib;window.jinSheetAssets=${JSON.stringify({ pdfWorker, pdfFonts, pptxWorker })};</script></head><body><div id="notice" role="status" aria-live="polite"></div><div id="sheets" role="tablist"></div><div id="toolbar"><div role="group" aria-label="${message('zoom-level', language)}" style="display:inline-flex;align-items:center;gap:3px"><button id="zoomOut" type="button" title="${message('zoom-out', language)}" aria-label="${message('zoom-out', language)}">−</button><input id="sheetZoom" type="number" min="10" max="400" step="5" value="100" style="width:56px;text-align:right" aria-label="${message('zoom-level', language)}"><span>%</span><button id="zoomIn" type="button" title="${message('zoom-in', language)}" aria-label="${message('zoom-in', language)}">+</button></div><button id="insertRow" type="button" title="${message('insert-row-title', language)}">${message('insert-row', language)}</button><button id="deleteRow" type="button" title="${message('delete-row-title', language)}">${message('delete-row', language)}</button><button id="wrapText" type="button" aria-pressed="false">${message('wrap-text', language)}</button><button id="autoFitRow" type="button">${message('auto-fit-row', language)}</button><label for="rowHeight">${message('row-height', language)}</label><input id="rowHeight" type="number" min="0.75" max="409.5" step="0.25" value="19.5" style="width:76px"><button id="applyRowHeight" type="button">${message('row-height-apply', language)}</button><button id="resetRowHeight" type="button">${message('row-height-default', language)}</button><button id="hideRows" type="button" title="${message('hide-rows-title', language)}">${message('hide-rows', language)}</button><button id="unhideRows" type="button" title="${message('unhide-rows-title', language)}">${message('unhide-rows', language)}</button><button id="hideColumns" type="button" title="${message('hide-columns-title', language)}">${message('hide-columns', language)}</button><button id="unhideColumns" type="button" title="${message('unhide-columns-title', language)}">${message('unhide-columns', language)}</button><button id="groupRows" type="button">${message('group-rows', language)}</button><button id="ungroupRows" type="button">${message('ungroup-rows', language)}</button><button id="collapseRows" type="button">${message('collapse-rows', language)}</button><button id="expandRows" type="button">${message('expand-rows', language)}</button><span id="rowOutlineLevels" class="outline-levels" role="group" aria-label="${message('row-outline-levels', language)}" hidden></span><button id="groupColumns" type="button">${message('group-columns', language)}</button><button id="ungroupColumns" type="button">${message('ungroup-columns', language)}</button><button id="collapseColumns" type="button">${message('collapse-columns', language)}</button><button id="expandColumns" type="button">${message('expand-columns', language)}</button><span id="columnOutlineLevels" class="outline-levels" role="group" aria-label="${message('column-outline-levels', language)}" hidden></span><small id="csvOutlineSessionNote" hidden>${message('csv-outline-session', language)}</small><button id="freezePanes" type="button">${message('freeze-panes', language)}</button><button id="freezeTopRow" type="button">${message('freeze-top-row', language)}</button><button id="freezeFirstColumn" type="button">${message('freeze-first-column', language)}</button><button id="unfreezePanes" type="button">${message('unfreeze-panes', language)}</button><button id="mergeCells" type="button">${message('merge-cells', language)}</button><button id="unmergeCells" type="button">${message('unmerge-cells', language)}</button><button id="insertColumn" type="button" title="${message('insert-column-title', language)}">${message('insert-column', language)}</button><button id="deleteColumn" type="button" title="${message('delete-column-title', language)}">${message('delete-column', language)}</button>${renderTableCreationControls(language)}<button id="resizeTable" type="button" title="${message('resize-table-title', language)}">${message('resize-table', language)}</button><label><input id="autoExpandTables" type="checkbox" checked title="${message('auto-table-title', language)}">${message('auto-table', language)}</label><label title="${message('calculated-column-auto-title', language)}"><input id="autoCalculatedColumns" type="checkbox" checked>${message('calculated-column-auto', language)}</label><span id="calculatedColumnLabel" style="width:180px;max-width:180px;overflow:hidden;white-space:nowrap;text-overflow:ellipsis"></span><select id="calculatedColumnScope" aria-label="${message('calculated-column-scope', language)}"><option value="preserve">${message('calculated-column-preserve', language)}</option><option value="all">${message('calculated-column-all', language)}</option></select><button id="applyCalculatedColumn" type="button">${message('calculated-column-apply', language)}</button><button id="clearCalculatedColumn" type="button" title="${message('calculated-column-clear-title', language)}">${message('calculated-column-clear', language)}</button><button id="tableTotals" type="button" aria-pressed="false">${message('table-totals', language)}</button><span id="tableTotalsColumn" style="width:180px;max-width:180px;overflow:hidden;white-space:nowrap;text-overflow:ellipsis"></span><select id="tableTotalsFunction" aria-label="${message('table-totals-function', language)}"><option value="none">${message('table-totals-none', language)}</option><option value="sum">${message('table-totals-sum', language)}</option><option value="average">${message('table-totals-average', language)}</option><option value="count">${message('table-totals-count', language)}</option><option value="countNums">${message('table-totals-countNums', language)}</option><option value="min">${message('table-totals-min', language)}</option><option value="max">${message('table-totals-max', language)}</option><option value="stdDev">${message('table-totals-stdDev', language)}</option><option value="var">${message('table-totals-var', language)}</option><option value="custom" disabled>${message('table-totals-custom', language)}</option></select><button id="applyTableTotalFunction" type="button">${message('table-totals-apply', language)}</button><button id="insertChart" type="button" title="${message('insert-chart-title', language)}">${message('insert-chart', language)}</button><select id="chartType" aria-label="${message('chart-type', language)}" title="${message('chart-type', language)}"><option value="bar">${message('chart-column', language)}</option><option value="line">${message('chart-line', language)}</option><option value="area">${message('chart-area', language)}</option><option value="scatter">${message('chart-scatter', language)}</option><option value="pie">${message('chart-pie', language)}</option><option value="doughnut">${message('chart-doughnut', language)}</option></select><button id="autoSum" type="button" title="${message('auto-sum-title', language)}">${message('auto-sum', language)}</button><button id="autoFitColumn" type="button" title="${message('auto-fit-column-title', language)}">${message('auto-fit-column', language)}</button><button id="formatNumber" type="button" title="${message('format-number-title', language)}">${message('format-number', language)}</button><button id="formatCurrency" type="button" title="${message('format-currency-title', language)}">${message('format-currency', language)}</button><button id="formatPercent" type="button" title="${message('format-percent-title', language)}">${message('format-percent', language)}</button><button id="formatDate" type="button" title="${message('format-date-title', language)}">${message('format-date', language)}</button><button id="formatTime" type="button" title="${message('format-time-title', language)}">${message('format-time', language)}</button><input id="customNumberFormat" type="text" maxlength="255" aria-label="${message('custom-number-format', language)}" placeholder="${message('custom-number-format-placeholder', language)}"><button id="applyCustomNumberFormat" type="button" title="${message('custom-number-format-title', language)}">${message('custom-number-format', language)}</button><span>${message('alignment-label', language)}</span><button id="alignLeft" type="button" title="${message('align-left-title', language)}">L</button><button id="alignCenter" type="button" title="${message('align-center-title', language)}">C</button><button id="alignRight" type="button" title="${message('align-right-title', language)}">R</button><button id="alignTop" type="button" title="${message('align-top-title', language)}">↑</button><button id="alignMiddle" type="button" title="${message('align-middle-title', language)}">↕</button><button id="alignBottom" type="button" title="${message('align-bottom-title', language)}">↓</button><button id="formatBold" type="button" aria-label="${message('format-bold-title', language)}" title="${message('format-bold-title', language)}"><strong>B</strong></button><button id="formatItalic" type="button" aria-label="${message('format-italic-title', language)}" title="${message('format-italic-title', language)}"><em>I</em></button><button id="formatUnderline" type="button" aria-label="${message('format-underline-title', language)}" title="${message('format-underline-title', language)}"><u>U</u></button><button id="formatStrike" type="button" aria-label="${message('format-strike-title', language)}" title="${message('format-strike-title', language)}"><s>S</s></button><label for="fontFamily">${message('font-family', language)}</label><input id="fontFamily" type="text" maxlength="64" aria-label="${message('font-family-title', language)}" title="${message('font-family-title', language)}" placeholder="${message('font-family-placeholder', language)}"><button id="applyFontFamily" type="button" title="${message('font-family-title', language)}">${message('font-family-apply', language)}</button><button id="clearFontFamily" type="button" title="${message('font-family-default', language)}">${message('font-family-default', language)}</button><label for="fontSize">${message('font-size', language)}</label><input id="fontSize" type="number" min="1" max="409" step="0.5" value="11" aria-label="${message('font-size-title', language)}" title="${message('font-size-title', language)}"><button id="applyFontSize" type="button" title="${message('font-size-apply-title', language)}">${message('font-size-apply', language)}</button><button id="clearFontSize" type="button" title="${message('font-size-clear-title', language)}">${message('font-size-clear', language)}</button><label for="fontColor" title="${message('font-color-title', language)}">${message('font-color', language)}</label><input id="fontColor" type="color" value="#000000" aria-label="${message('font-color-title', language)}" title="${message('font-color-title', language)}"><button id="clearFontColor" type="button" title="${message('clear-font-color-title', language)}">${message('clear-font-color', language)}</button><label for="formatFillColor" title="${message('format-fill-title', language)}">${message('format-fill', language)}</label><input id="formatFillColor" type="color" value="#fff2cc" aria-label="${message('format-fill-title', language)}" title="${message('format-fill-title', language)}"><button id="clearFillColor" type="button" title="${message('clear-fill-title', language)}">${message('clear-fill', language)}</button><label for="borderPreset">${message('border-preset', language)}</label><select id="borderPreset" aria-label="${message('border-preset-title', language)}" title="${message('border-preset-title', language)}"><option value="none">${message('border-none', language)}</option><option value="outside">${message('border-outside', language)}</option><option value="all">${message('border-all', language)}</option><option value="top">${message('border-top', language)}</option><option value="bottom">${message('border-bottom', language)}</option><option value="left">${message('border-left', language)}</option><option value="right">${message('border-right', language)}</option><option value="inside-horizontal">${message('border-inside-horizontal', language)}</option><option value="inside-vertical">${message('border-inside-vertical', language)}</option></select><label for="borderLine">${message('border-line-label', language)}</label><select id="borderLine" aria-label="${message('border-line-title', language)}" title="${message('border-line-title', language)}"><option value="thin">${message('border-thin', language)}</option><option value="medium">${message('border-medium', language)}</option><option value="thick">${message('border-thick', language)}</option><option value="dashed">${message('border-dashed', language)}</option><option value="dotted">${message('border-dotted', language)}</option><option value="double">${message('border-double', language)}</option></select><label for="borderColor" title="${message('border-color-title', language)}">${message('border-color', language)}</label><input id="borderColor" type="color" value="#000000" aria-label="${message('border-color-title', language)}" title="${message('border-color-title', language)}"><button id="applyBorderPreset" type="button" title="${message('apply-border-title', language)}">${message('apply-border', language)}</button><button id="sortAscending" type="button" title="${message('sort-ascending-title', language)}">${message('sort-ascending', language)}</button><button id="sortDescending" type="button" title="${message('sort-descending-title', language)}">${message('sort-descending', language)}</button><button id="filterValuesButton" type="button" title="${message('filter-values-title', language)}">${message('filter-values', language)}</button><label for="sortHasHeader"><input id="sortHasHeader" type="checkbox">${message('sort-has-header', language)}</label><label for="sortExpandSelection"><input id="sortExpandSelection" type="checkbox" checked>${message('sort-expand-selection', language)}</label><input id="address" type="text" aria-label="${message('cell-reference', language)}" title="${message('cell-reference', language)} (Ctrl/Cmd+G)" placeholder="${message('cell-reference-hint', language)}" autocomplete="off" spellcheck="false"><span id="goToStatus" role="status" aria-live="polite"></span><div id="cellValueGroup"><label for="formula">${message('cell-value-label', language)}</label><input id="formula" type="text" list="cellValidationOptions" aria-label="${message('cell-value', language)}" placeholder="${message('cell-value', language)}"><datalist id="cellValidationOptions"></datalist></div><div id="searchGroup" role="search"><label for="search">${message('search-sheet-label', language)}</label><input id="search" type="search" aria-label="${message('search-sheet', language)}" placeholder="${message('find-hint', language)}"><button id="findButton" type="button" title="${message('find-button-title', language)}">${message('find-button', language)}</button><button id="clearSearch" type="button" title="${message('clear-search-title', language)}" aria-label="${message('clear-search-title', language)}" hidden>${message('clear-search', language)}</button><select id="searchScope" aria-label="${message('search-scope-label', language)}"><option value="sheet">${message('search-scope-sheet', language)}</option><option value="workbook">${message('search-scope-workbook', language)}</option></select><span id="searchCount" aria-live="polite"></span></div><div id="replaceGroup" role="group" aria-label="${message('replace-group-label', language)}"><label for="replaceValue">${message('replace-value', language)}</label><input id="replaceValue" type="text" maxlength="100000" aria-label="${message('replace-value', language)}"><button id="replaceCurrentButton" type="button">${message('replace-current', language)}</button><button id="replaceAllButton" type="button">${message('replace-all', language)}</button><span id="replaceStatus" role="status" aria-live="polite"></span></div><button id="printSheet" type="button">${message('print-sheet', language)}</button><button id="printSelection" type="button">${message('print-selection', language)}</button><span id="selectionSummary" role="status" aria-live="polite"></span></div>${renderPrintControls(language)}<div id="filterPanel" role="dialog" aria-label="${message('filter-dialog-title', language)}" hidden><strong id="filterTitle"></strong><small id="filterSessionNote" hidden>${message('filter-csv-session', language)}</small><fieldset id="filterEditor">${renderFilterConditionControls(language)}<div id="filterValueControls"><label><input id="filterSelectAll" type="checkbox">${message('filter-select-all', language)}</label><input id="filterSearch" type="search" aria-label="${message('filter-search', language)}" placeholder="${message('filter-search', language)}"><div id="filterValuesList"></div><button id="filterShowMore" type="button" hidden>${message('filter-show-more', language)}</button></div></fieldset><div><button id="filterApply" type="button">${message('filter-apply', language)}</button><button id="filterClear" type="button">${message('filter-clear', language)}</button><button id="filterCancel" type="button">${message('filter-cancel', language)}</button></div></div><div id="grid" role="grid" aria-label="${message('spreadsheet', language)}" tabindex="0"><div id="canvas"></div></div><div id="chartPreview" hidden></div><section id="reader" aria-label="${message('document-preview', language)}"><div id="readerToolbar"></div><div id="readerStyles"></div><div id="readerCanvas"></div></section><script nonce="${nonce}">
  const ui=${ui};const tableStyleNames=${JSON.stringify(TABLE_STYLE_NAMES)};${gridSource};const pdfCanvasOutputScale=${pdfCanvasOutputScaleSource};const t=key=>ui[key]||key;
  const bridge=acquireVsCodeApi();const vscode={postMessage(payload){if(['edit','paste','auto-sum','fill-range','fill-selection','fill-handle','move-paste','sort-selection'].includes(payload.type)){if(document.getElementById('autoCalculatedColumns')?.checked===false)payload={...payload,tableAutoFill:false};if(document.getElementById('autoExpandTables')?.checked===false)payload={...payload,tableAutoExpand:false}}bridge.postMessage(payload)}};let model;let gridZoom=1;let rowView;let rowMetrics;let selected={row:0,column:0};let selectionAnchor={row:0,column:0};let selectedRange={startRow:0,endRow:0,startColumn:0,endColumn:0};let selectionsBySheet=new Map();let pendingCutRange=null;let dragging=false;let dragAxis=null;let dragged=false;let dragPointer=null;let dragMoved=false;let dragScrollTimer=null;let shiftSelecting=false;let navigationRequest=0;let currentRegionRequest=0;let currentRegionStage=0;let rangeRequest=0;let renderedRange=null;let pendingRange=null;let matches=[];let matchedCells=new Set();let matchIndex=0;let searchedQuery='';let searchRequest=0;let searchPagePending=false;let replaceRequest=0;let pendingReplaceRequest=0;let scrollTimer;let renderGeneration=0;let renderedInstance=null;let pdfLoadingTask=null;let pdfDocument=null;let pdfPageNumber=1;let pdfZoom=1;let pdfPages=[];let pdfPageObserver=null;let pdfResizeObserver=null;let pdfBaseScale=1;let pdfViewportWidth=0;let pdfFitMode='width';const XLSX_MAX_ROWS=1048576;const XLSX_MAX_COLUMNS=16384;const RANGE_ROW_CHUNK=10;const RANGE_COLUMN_CHUNK=5;let pptxWorkerBlobUrl=null;let autoFitRequest=0;let docxUserZoom=100;let docxFitMode='width';let docxResizeObserver=null;let columnWidthsBySheet=new Map();let columnWidths=new Map();let columnOffsets=new Float64Array([0]);let columnResize=null;let rowResize=null;let rowBoundaryFocus=null;let fillDrag=null;const label=n=>{let out='';while(n>0){n--;out=String.fromCharCode(65+n%26)+out;n=Math.floor(n/26)}return out};
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
  function keepGridFocus(){const grid=document.getElementById('grid');if(grid.contains(document.activeElement)&&document.activeElement!==grid)grid.focus({preventScroll:true})}
  function fitBoxAxis(item,size,vertical){const cache=vertical?'__rowBox':'__columnBox',first=vertical?'Top':'Left',last=vertical?'Bottom':'Right',threshold=vertical?26:32;if(size>=threshold&&!item[cache])return;const keys=['padding'+first,'padding'+last,'border'+first+'Width','border'+last+'Width'];if(!item[cache])item[cache]=Object.fromEntries(keys.map(key=>[key,item.style[key]||'']));for(const key of keys)item.style[key]=item[cache][key];if(size>=threshold){delete item[cache];return}item.style[keys[0]]='0';item.style[keys[1]]='0';const start=Number.parseFloat(item.style[keys[2]])||1,end=Number.parseFloat(item.style[keys[3]])||1;item.style[keys[2]]=(size<2?0:Math.min(start,size/2))+'px';item.style[keys[3]]=Math.min(end,size<2?size:size/2)+'px'}
  function fitColumnBox(item,width){fitBoxAxis(item,width,false)}
  function fitRowBox(item,height){fitBoxAxis(item,height,true)}
  function gridColumnWidth(width){return clampColumnWidth(width,model?.kind==='xlsx')}
  function draw(){if(!model)return;syncPrintControls();if(model.kind==='xlsx'){model.gridRowCount=XLSX_MAX_ROWS;model.gridColumnCount=XLSX_MAX_COLUMNS}else if(model.kind==='csv'){model.gridRowCount=10000;model.gridColumnCount=1000}else{model.gridRowCount=model.rowCount;model.gridColumnCount=model.columnCount}rowView=createRowView(model.gridRowCount,model.visibleRows,model.hiddenRows);model.gridDisplayRowCount=rowView.displayCount;rowMetrics=createRowMetrics(rowView,model.rowHeights||{},model.defaultRowHeight||26,gridZoom);const activeSheet=model.sheetIndex||0;if(!columnWidthsBySheet.has(activeSheet)){const savedWidths=new Map(Object.entries(model.columnWidths||{}).map(([column,width])=>[Number(column),excelColumnWidthPixels(width,DEFAULT_COLUMN_WIDTH,model.maximumDigitWidth??7)]));columnWidthsBySheet.set(activeSheet,savedWidths)}columnWidths=columnWidthsBySheet.get(activeSheet);columnOffsets=createColumnOffsets(model.gridColumnCount,columnWidths,model.hiddenColumns||[],model.defaultColumnWidth??DEFAULT_COLUMN_WIDTH,model.kind==='xlsx',gridZoom);const maxRow=Math.max(0,model.gridRowCount-1),maxColumn=Math.max(0,model.gridColumnCount-1);selected.row=Math.min(selected.row,maxRow);selected.column=Math.min(selected.column,maxColumn);if(model.hiddenColumns?.includes(selected.column)){let next=selected.column+1;while(next<=maxColumn&&model.hiddenColumns.includes(next))next++;if(next>maxColumn){next=selected.column-1;while(next>=0&&model.hiddenColumns.includes(next))next--}if(next>=0&&next<=maxColumn){selected.column=next;selectionAnchor.column=next;selectedRange={startRow:selected.row,endRow:selected.row,startColumn:next,endColumn:next}}}selectionAnchor.row=Math.min(selectionAnchor.row,maxRow);selectionAnchor.column=Math.min(selectionAnchor.column,maxColumn);selectedRange.startRow=Math.min(selectedRange.startRow,maxRow);selectedRange.endRow=Math.min(selectedRange.endRow,maxRow);selectedRange.startColumn=Math.min(selectedRange.startColumn,maxColumn);selectedRange.endColumn=Math.min(selectedRange.endColumn,maxColumn);const notice=document.getElementById('notice');notice.textContent=(model.reason||'')+(model.formulas?' · '+t('formula-cache-notice'):'')+(model.dirty?' · '+t('modified'):'')+(model.readOnly?'':' · '+t('edit-shortcut'))+(model.rowCount>1000?' · '+t('scroll-rows'):'');const tabs=document.getElementById('sheets');tabs.replaceChildren();const grid=document.getElementById('grid'),toolbar=document.getElementById('toolbar'),canvas=document.getElementById('canvas');canvas.style.fontSize=gridZoom+'em';if(model.kind==='office'||model.kind==='pdf'){renderDocument(model);return}releaseReader();document.body.classList.remove('document-view');document.getElementById('reader').style.display='none';grid.style.display='';toolbar.style.display='';tabs.style.display='flex';for(const id of ['insertRow','deleteRow','insertColumn','deleteColumn'])document.getElementById(id).disabled=Boolean(model.readOnly);syncTableResizeButton();syncRowHeightControls();document.getElementById('autoSum').disabled=Boolean(model.readOnly||model.kind!=='xlsx');for(const id of ['hideRows','hideColumns','unhideRows','unhideColumns'])document.getElementById(id).disabled=Boolean(model.readOnly||!['xlsx','csv'].includes(model.kind));for(const id of ['groupRows','ungroupRows','collapseRows','expandRows','groupColumns','ungroupColumns','collapseColumns','expandColumns'])document.getElementById(id).disabled=Boolean(model.readOnly||!['xlsx','csv'].includes(model.kind));for(const id of ['freezePanes','freezeTopRow','freezeFirstColumn','unfreezePanes'])document.getElementById(id).disabled=Boolean(model.readOnly||model.kind!=='xlsx');document.getElementById('unhideRows').disabled=Boolean(model.readOnly||!['xlsx','csv'].includes(model.kind)||!(model.manuallyHiddenRows?.length));document.getElementById('unhideColumns').disabled=Boolean(model.readOnly||!['xlsx','csv'].includes(model.kind)||!(model.manuallyHiddenColumns?.length));document.getElementById('unfreezePanes').disabled=Boolean(model.readOnly||model.kind!=='xlsx'||(!(model.freezePane?.rows)&&!(model.freezePane?.columns)));const groupRowStart=Math.min(selectedRange.startRow,selectedRange.endRow),groupRowEnd=Math.max(selectedRange.startRow,selectedRange.endRow),groupColumnStart=Math.min(selectedRange.startColumn,selectedRange.endColumn),groupColumnEnd=Math.max(selectedRange.startColumn,selectedRange.endColumn),rowGroupMatches=(model.rowGroups||[]).filter(group=>group.start===groupRowStart&&group.end===groupRowEnd),columnGroupMatches=(model.columnGroups||[]).filter(group=>group.start===groupColumnStart&&group.end===groupColumnEnd),outlineDisabled=Boolean(model.readOnly||!['xlsx','csv'].includes(model.kind));syncOutlineLevelControls();document.getElementById('csvOutlineSessionNote').hidden=model.kind!=='csv';for(const id of ['groupRows','ungroupRows','groupColumns','ungroupColumns'])document.getElementById(id).disabled=outlineDisabled;document.getElementById('collapseRows').disabled=outlineDisabled||!rowGroupMatches.some(group=>!group.collapsed);document.getElementById('expandRows').disabled=outlineDisabled||!rowGroupMatches.some(group=>group.collapsed);document.getElementById('collapseColumns').disabled=outlineDisabled||!columnGroupMatches.some(group=>!group.collapsed);document.getElementById('expandColumns').disabled=outlineDisabled||!columnGroupMatches.some(group=>group.collapsed); for(const id of ['sortAscending','sortDescending'])document.getElementById(id).disabled=model.readOnly;for(const id of ['formatNumber','formatCurrency','formatPercent','formatDate','formatTime','applyCustomNumberFormat','customNumberFormat','formatBold','formatItalic','formatUnderline','formatStrike','fontFamily','applyFontFamily','clearFontFamily','fontSize','applyFontSize','clearFontSize','alignLeft','alignCenter','alignRight','alignTop','alignMiddle','alignBottom','fontColor','clearFontColor','formatFillColor','clearFillColor','borderPreset','borderLine','borderColor','applyBorderPreset'])document.getElementById(id).disabled=Boolean(model.readOnly||model.kind!== 'xlsx');document.getElementById('filterValuesButton').disabled=Boolean(model.readOnly||model.kind!=='xlsx'||!(model.autoFilter||(model.tables?.some(table=>table.headerRowCount===1)||(model.rowCount>1&&model.columnCount>0))));canvas.style.position='relative';canvas.style.padding='0';canvas.style.maxWidth='';canvas.style.margin='';for(const sheet of model.sheets||[]){if(model.kind==='xlsx'&&!sheet.visible)continue;const button=document.createElement('button');button.textContent=sheet.name+(sheet.visible?'':' ('+t('hidden')+')');button.setAttribute('role','tab');button.setAttribute('aria-selected',String(sheet.index===model.sheetIndex));button.onclick=()=>vscode.postMessage({type:'refresh',sheetIndex:sheet.index});button.ondblclick=event=>{if(model.readOnly||model.kind!=='xlsx')return;event.preventDefault();const input=document.createElement('input');input.value=sheet.name;input.setAttribute('aria-label',t('rename-sheet'));button.replaceChildren(input);input.focus();input.select();let done=false;const finish=commit=>{if(done)return;done=true;const name=input.value.trim();button.textContent=sheet.name+(sheet.visible?'':' ('+t('hidden')+')');if(commit&&name&&name!==sheet.name)vscode.postMessage({type:'rename-sheet',sheetIndex:sheet.index,name})};input.addEventListener('keydown',event=>{if(event.key==='Enter'){event.preventDefault();finish(true)}else if(event.key==='Escape'){event.preventDefault();finish(false)}});input.addEventListener('blur',()=>finish(false))};tabs.append(button)}if(model.kind==='xlsx'){const hideSheet=document.createElement('button');hideSheet.textContent=t('hide-sheet');hideSheet.setAttribute('type','button');hideSheet.setAttribute('aria-label',t('hide-sheet'));hideSheet.title=t('hide-sheet');hideSheet.disabled=Boolean(model.readOnly||(model.sheets||[]).filter(sheet=>sheet.visible).length<=1);hideSheet.onclick=()=>vscode.postMessage({type:'sheet-visibility',sheetIndex:model.sheetIndex,visible:false});tabs.append(hideSheet);const hiddenSheets=(model.sheets||[]).filter(sheet=>!sheet.visible);if(hiddenSheets.length){const unhideSheet=document.createElement('select');unhideSheet.setAttribute('aria-label',t('unhide-sheet'));unhideSheet.title=t('unhide-sheet');unhideSheet.disabled=Boolean(model.readOnly);const placeholder=document.createElement('option');placeholder.value='';placeholder.textContent=t('unhide-sheet');unhideSheet.append(placeholder);for(const sheet of hiddenSheets){const option=document.createElement('option');option.value=String(sheet.index);option.textContent=sheet.name;unhideSheet.append(option)}unhideSheet.onchange=()=>{if(unhideSheet.value==='')return;const sheetIndex=Number(unhideSheet.value);unhideSheet.value='';vscode.postMessage({type:'sheet-visibility',sheetIndex,visible:true})};tabs.append(unhideSheet)}const addSheet=document.createElement('button');addSheet.textContent='+';addSheet.setAttribute('type','button');addSheet.setAttribute('aria-label',t('add-sheet'));addSheet.title=t('add-sheet');addSheet.disabled=Boolean(model.readOnly);addSheet.onclick=()=>vscode.postMessage({type:'add-sheet',sheetIndex:model.sheetIndex});tabs.append(addSheet);const deleteSheet=document.createElement('button');deleteSheet.textContent='−';deleteSheet.setAttribute('type','button');deleteSheet.setAttribute('aria-label',t('delete-sheet'));deleteSheet.title=t('delete-sheet');deleteSheet.disabled=Boolean(model.readOnly||(model.sheets||[]).length<=1);deleteSheet.onclick=()=>vscode.postMessage({type:'delete-sheet',sheetIndex:model.sheetIndex});tabs.append(deleteSheet)}rangeRequest++;renderedRange=null;pendingRange=null;canvas.style.width=(ROW_HEADER_WIDTH+columnOffsets[model.gridColumnCount])+'px';canvas.style.height=(30+rowMetrics.totalHeight)+'px';keepGridFocus();canvas.replaceChildren();if(!model.gridRowCount||!model.gridColumnCount){canvas.textContent=model.reason||t('no-data');return}const corner=document.createElement('div');corner.className='corner';canvas.append(corner);requestRange()}
  for(const[id,height]of [['applyRowHeight',undefined],['resetRowHeight',null]])document.getElementById(id).onclick=()=>{if(model?.readOnly||model?.kind!=='xlsx')return;const value=height===null?null:Number(document.getElementById('rowHeight').value);if(value!==null&&(!Number.isFinite(value)||value<.75||value>409.5))return;vscode.postMessage({type:'row-height',startRow:Math.min(selectedRange.startRow,selectedRange.endRow),endRow:Math.max(selectedRange.startRow,selectedRange.endRow),height:value,sheetIndex:model.sheetIndex||0})};
  document.getElementById('wrapText').onclick=()=>{if(model?.readOnly||model?.kind!=='xlsx')return;vscode.postMessage({type:'wrap-text',range:{...selectedRange},sheetIndex:model.sheetIndex||0})};document.getElementById('autoFitRow').onclick=()=>{if(model?.readOnly||model?.kind!=='xlsx')return;vscode.postMessage({type:'row-auto-fit',startRow:Math.min(selectedRange.startRow,selectedRange.endRow),endRow:Math.max(selectedRange.startRow,selectedRange.endRow),sheetIndex:model.sheetIndex||0})};
  document.getElementById('hideRows').onclick=()=>{if(model?.readOnly||!['xlsx','csv'].includes(model?.kind))return;vscode.postMessage({type:'hidden-rows',action:'hide',startRow:selectedRange.startRow,endRow:selectedRange.endRow,sheetIndex:model.sheetIndex||0})};document.getElementById('unhideRows').onclick=()=>{if(model?.readOnly||!['xlsx','csv'].includes(model?.kind))return;vscode.postMessage({type:'hidden-rows',action:'unhide-all',sheetIndex:model.sheetIndex||0})};document.getElementById('hideColumns').onclick=()=>{if(model?.readOnly||!['xlsx','csv'].includes(model?.kind))return;vscode.postMessage({type:'hidden-columns',action:'hide',startColumn:selectedRange.startColumn,endColumn:selectedRange.endColumn,sheetIndex:model.sheetIndex||0})};document.getElementById('unhideColumns').onclick=()=>{if(model?.readOnly||!['xlsx','csv'].includes(model?.kind))return;vscode.postMessage({type:'hidden-columns',action:'unhide-all',sheetIndex:model.sheetIndex||0})};const sendOutlineGroup=(axis,action)=>{if(model?.readOnly||!['xlsx','csv'].includes(model?.kind))return;const start=axis==='row'?Math.min(selectedRange.startRow,selectedRange.endRow):Math.min(selectedRange.startColumn,selectedRange.endColumn),end=axis==='row'?Math.max(selectedRange.startRow,selectedRange.endRow):Math.max(selectedRange.startColumn,selectedRange.endColumn);vscode.postMessage({type:'outline-groups',axis,action,start,end,sheetIndex:model.sheetIndex||0})};for(const[id,axis,action]of [['groupRows','row','group'],['ungroupRows','row','ungroup'],['collapseRows','row','collapse'],['expandRows','row','expand'],['groupColumns','column','group'],['ungroupColumns','column','ungroup'],['collapseColumns','column','collapse'],['expandColumns','column','expand']])document.getElementById(id).onclick=()=>sendOutlineGroup(axis,action);for(const[id,rows,columns]of [['freezePanes',selected.row,selected.column],['freezeTopRow',1,0],['freezeFirstColumn',0,1]])document.getElementById(id).onclick=()=>{if(model?.readOnly||model?.kind!=='xlsx')return;vscode.postMessage({type:'freeze-pane',rows,columns,sheetIndex:model.sheetIndex||0})};document.getElementById('unfreezePanes').onclick=()=>{if(model?.readOnly||model?.kind!=='xlsx')return;vscode.postMessage({type:'freeze-pane',rows:0,columns:0,sheetIndex:model.sheetIndex||0})};document.getElementById('mergeCells').onclick=()=>{if(model?.readOnly||model?.kind!=='xlsx')return;vscode.postMessage({type:'merge-cells',action:'merge',range:selectedRange,sheetIndex:model.sheetIndex||0})};document.getElementById('unmergeCells').onclick=()=>{if(model?.readOnly||model?.kind!=='xlsx')return;vscode.postMessage({type:'merge-cells',action:'unmerge',range:selectedRange,sheetIndex:model.sheetIndex||0})};
  for(const [id,axis,action] of [['insertRow','row','insert'],['deleteRow','row','delete'],['insertColumn','column','insert'],['deleteColumn','column','delete']])document.getElementById(id).onclick=()=>{if(model?.readOnly)return;const start=axis==='row'?Math.min(selectedRange.startRow,selectedRange.endRow):Math.min(selectedRange.startColumn,selectedRange.endColumn),end=axis==='row'?Math.max(selectedRange.startRow,selectedRange.endRow):Math.max(selectedRange.startColumn,selectedRange.endColumn),count=end-start+1;vscode.postMessage({type:'structure',axis,action,index:start,...(count>1?{count}:{}),sheetIndex:model?.sheetIndex||0})};
  document.getElementById('autoFitColumn').onclick=()=>{if(!model)return;const requestId=++autoFitRequest;vscode.postMessage({type:'auto-fit-column',requestId,column:selected.column,sheetIndex:model.sheetIndex||0})};document.getElementById('autoSum').onclick=()=>{if(!model||model.readOnly||model.kind!=='xlsx')return;vscode.postMessage({type:'auto-sum',row:selected.row,column:selected.column,sheetIndex:model.sheetIndex||0})};
  for(const [id,format] of [['formatNumber','number'],['formatCurrency','currency'],['formatPercent','percent'],['formatDate','date'],['formatTime','time']])document.getElementById(id).onclick=()=>{if(!model||model.readOnly||model.kind!=='xlsx')return;vscode.postMessage({type:'number-format',range:{...selectedRange},format,sheetIndex:model.sheetIndex||0})};document.getElementById('applyCustomNumberFormat').onclick=()=>{if(!model||model.readOnly||model.kind!=='xlsx')return;vscode.postMessage({type:'number-format-custom',range:{...selectedRange},formatCode:document.getElementById('customNumberFormat').value.trim(),sheetIndex:model.sheetIndex||0})};for(const [id,property] of [['formatBold','bold'],['formatItalic','italic'],['formatUnderline','underline'],['formatStrike','strike']])document.getElementById(id).onclick=()=>{if(!model||model.readOnly||model.kind!=='xlsx')return;vscode.postMessage({type:'font-format',range:{...selectedRange},property,sheetIndex:model.sheetIndex||0})};for(const [id,axis,value] of [['alignLeft','horizontal','left'],['alignCenter','horizontal','center'],['alignRight','horizontal','right'],['alignTop','vertical','top'],['alignMiddle','vertical','center'],['alignBottom','vertical','bottom']])document.getElementById(id).onclick=()=>{if(!model||model.readOnly||model.kind!=='xlsx')return;vscode.postMessage({type:'alignment-format',range:{...selectedRange},axis,value,sheetIndex:model.sheetIndex||0})};document.getElementById('applyFontFamily').onclick=()=>{if(!model||model.readOnly||model.kind!=='xlsx')return;const family=document.getElementById('fontFamily').value.trim();if(!family||family.length>64)return;vscode.postMessage({type:'font-family',range:{...selectedRange},family,sheetIndex:model.sheetIndex||0})};document.getElementById('clearFontFamily').onclick=()=>{if(!model||model.readOnly||model.kind!=='xlsx')return;vscode.postMessage({type:'font-family',range:{...selectedRange},family:null,sheetIndex:model.sheetIndex||0})};document.getElementById('applyFontSize').onclick=()=>{if(!model||model.readOnly||model.kind!=='xlsx')return;const size=Number(document.getElementById('fontSize').value);if(!Number.isFinite(size))return;vscode.postMessage({type:'font-size',range:{...selectedRange},size,sheetIndex:model.sheetIndex||0})};document.getElementById('clearFontSize').onclick=()=>{if(!model||model.readOnly||model.kind!=='xlsx')return;vscode.postMessage({type:'font-size',range:{...selectedRange},size:null,sheetIndex:model.sheetIndex||0})};document.getElementById('fontColor').onchange=e=>{if(!model||model.readOnly||model.kind!=='xlsx')return;vscode.postMessage({type:'font-color',range:{...selectedRange},color:e.target.value.toUpperCase(),sheetIndex:model.sheetIndex||0})};document.getElementById('clearFontColor').onclick=()=>{if(!model||model.readOnly||model.kind!=='xlsx')return;vscode.postMessage({type:'font-color',range:{...selectedRange},color:null,sheetIndex:model.sheetIndex||0})};document.getElementById('formatFillColor').onchange=e=>{if(!model||model.readOnly||model.kind!=='xlsx')return;vscode.postMessage({type:'fill-format',range:{...selectedRange},color:e.target.value.toUpperCase(),sheetIndex:model.sheetIndex||0})};document.getElementById('clearFillColor').onclick=()=>{if(!model||model.readOnly||model.kind!=='xlsx')return;vscode.postMessage({type:'fill-format',range:{...selectedRange},color:null,sheetIndex:model.sheetIndex||0})};document.getElementById('applyBorderPreset').onclick=()=>{if(!model||model.readOnly||model.kind!=='xlsx')return;vscode.postMessage({type:'border-format',range:{...selectedRange},preset:document.getElementById('borderPreset').value,line:document.getElementById('borderLine').value||'thin',color:(document.getElementById('borderColor').value||'#000000').toUpperCase(),sheetIndex:model.sheetIndex||0})};
  for(const [id,direction] of [['sortAscending','ascending'],['sortDescending','descending']])document.getElementById(id).onclick=()=>{if(!model||model.readOnly)return;vscode.postMessage({type:'sort-selection',direction,range:selectedRange,keyColumn:selected.column,hasHeader:document.getElementById('sortHasHeader').checked,expandSelection:document.getElementById('sortExpandSelection').checked,sheetIndex:model.sheetIndex||0})};
  function rowDisplayIndex(row){let low=0,high=rowView?.displayCount||0;while(low<high){const middle=Math.floor((low+high)/2);if(rowView.rowAt(middle)<row)low=middle+1;else high=middle}return low}function frozenRowsHeight(){return rowMetrics?.top(rowDisplayIndex(model?.freezePane?.rows||0))||0}
  function requestRange(){if(!model||!model.gridRowCount||!model.gridColumnCount||!model.gridDisplayRowCount)return;const grid=document.getElementById('grid');const visibleRow=rowMetrics.indexAtPixel(grid.scrollTop);const rowStart=Math.floor(visibleRow/RANGE_ROW_CHUNK)*RANGE_ROW_CHUNK;const freezeRows=Math.max(0,Math.min(50,model.freezePane?.rows||0));const freezeColumns=Math.max(0,Math.min(10,model.freezePane?.columns||0));const rowCount=Math.min(60-freezeRows,Math.max(1,(rowMetrics.uniformHeight?Math.ceil(grid.clientHeight/rowMetrics.uniformHeight):rowMetrics.indexAtPixel(grid.scrollTop+grid.clientHeight-.001)-rowStart+1)+RANGE_ROW_CHUNK),model.gridDisplayRowCount-rowStart);const {columnStart,columnCount}=visibleColumnRange(columnOffsets,grid.scrollLeft,grid.clientWidth);const {columnStart:visibleColumnStart,columnCount:visibleColumnCount}=visibleColumnRange(columnOffsets,grid.scrollLeft,grid.clientWidth,{maxColumns:Math.max(1,30-freezeColumns)});if(rowCount<1||visibleColumnCount<1)return;let range={rowStart,columnStart:visibleColumnStart,rowCount,columnCount:visibleColumnCount,sheetIndex:model.sheetIndex};if(model.visibleRows||model.hiddenRows){const positions=new Map();for(let index=0;index<rowCount;index++){const displayIndex=rowStart+index;const actualRow=rowView.rowAt(displayIndex);if(actualRow>=0)positions.set(actualRow,displayIndex)}for(let row=0;row<freezeRows;row++){const displayIndex=rowView.indexOf(row);if(displayIndex>=0)positions.set(row,displayIndex)}const entries=[...positions].sort((a,b)=>a[0]-b[0]);const viewRows=entries.map(entry=>entry[0]),viewIndexes=entries.map(entry=>entry[1]);const columnIndexes=[...new Set([...Array.from({length:freezeColumns},(_,index)=>index),...Array.from({length:visibleColumnCount},(_,index)=>visibleColumnStart+index)])].sort((a,b)=>a-b);if(viewRows.length>60||columnIndexes.length>30)return;range={...range,rowStart:viewIndexes[0]||0,columnStart:columnIndexes[0]||0,rowCount:viewRows.length,columnCount:columnIndexes.length,viewRows,viewIndexes,columnIndexes}}else if(freezeRows||freezeColumns){const viewRows=[...new Set([...Array.from({length:freezeRows},(_,index)=>index),...Array.from({length:rowCount},(_,index)=>rowStart+index)])].sort((a,b)=>a-b);const columnIndexes=[...new Set([...Array.from({length:freezeColumns},(_,index)=>index),...Array.from({length:visibleColumnCount},(_,index)=>visibleColumnStart+index)])].sort((a,b)=>a-b);if(viewRows.length>60||columnIndexes.length>30)return;range={...range,rowStart:viewRows[0]||0,columnStart:columnIndexes[0]||0,rowCount:viewRows.length,columnCount:columnIndexes.length,viewRows,columnIndexes}}else range={...range,columnStart:visibleColumnStart,rowCount:Math.min(60,Math.max(1,(rowMetrics.uniformHeight?Math.ceil(grid.clientHeight/rowMetrics.uniformHeight):rowMetrics.indexAtPixel(grid.scrollTop+grid.clientHeight-.001)-rowStart+1)+RANGE_ROW_CHUNK),model.gridDisplayRowCount-rowStart)};const same=(left,right)=>left&&right&&Object.keys(right).every(key=>['viewRows','viewIndexes','columnIndexes'].includes(key)?(!right[key]||JSON.stringify(left[key])===JSON.stringify(right[key])):left[key]===right[key]);if(same(pendingRange,range))return;if(same(renderedRange,range)){if(pendingRange){rangeRequest++;pendingRange=null}return}const requestId=++rangeRequest;pendingRange={...range,requestId};vscode.postMessage({type:'range',requestId,...range})}
  function rowResizeBounds(row){const top=Math.min(selectedRange.startRow,selectedRange.endRow),bottom=Math.max(selectedRange.startRow,selectedRange.endRow);const wholeRows=selectionAnchor.axis==='row'||Math.min(selectedRange.startColumn,selectedRange.endColumn)===0&&Math.max(selectedRange.startColumn,selectedRange.endColumn)===(model?.gridColumnCount||1)-1;return wholeRows&&row>=top&&row<=bottom?{startRow:top,endRow:bottom}:{startRow:row,endRow:row}}
  function updateRowGeometry(){rowMetrics=createRowMetrics(rowView,model.rowHeights||{},model.defaultRowHeight||26,gridZoom);const canvas=document.getElementById('canvas');canvas.style.height=(30+rowMetrics.totalHeight)+'px';for(const item of canvas.children){if(item.classList.contains('rowhead')){const index=Number(item.dataset.viewRow);item.style.top=(30+rowMetrics.top(index))+'px';item.style.height=rowMetrics.height(index)+'px';fitRowBox(item,rowMetrics.height(index));const handle=[...item.children].find(child=>child.classList.contains('row-resize-handle'));handle?.setAttribute('aria-valuenow',String(Math.round(rowMetrics.height(index)*.75*100)/100))}else if(item.classList.contains('cell')){const row=Number(item.dataset.row),index=rowView.indexOf(row),merged=model.mergedCells?.find(range=>row===range.top&&Number(item.dataset.column)===range.left);item.style.top=(30+rowMetrics.top(index))+'px';item.style.height=(merged?rowMetrics.top(rowDisplayIndex(merged.bottom+1))-rowMetrics.top(index):rowMetrics.height(index))+'px';fitRowBox(item,Number.parseFloat(item.style.height))}}syncHeadings();syncRowHeightControls()}
  function previewRowHeight(bounds,pixels){const height=clampRowHeightPixels(pixels);if(bounds.startRow===0&&bounds.endRow===model.gridRowCount-1){model.defaultRowHeight=height;model.rowHeights={}}else{model.rowHeights={...(model.rowHeights||{})};for(let row=bounds.startRow;row<=bounds.endRow;row++)model.rowHeights[row]=height}updateRowGeometry();return height}
  function commitRowHeight(bounds,height){vscode.postMessage({type:'row-height',...bounds,height:Math.round(height*.75*100)/100,sheetIndex:model.sheetIndex||0})}
  function cancelRowResize(){if(!rowResize)return;const previous=rowResize;rowResize=null;document.body.style.cursor=previous.cursor;if(!previous.moved)return;if(model?.sheetIndex===previous.sheetIndex){model.rowHeights=previous.heights;model.defaultRowHeight=previous.defaultHeight;updateRowGeometry();requestRange()}}
  function updateColumnGeometry(){if(!model)return;const canvas=document.getElementById('canvas');columnOffsets=createColumnOffsets(model.gridColumnCount,columnWidths,model.hiddenColumns||[],model.defaultColumnWidth??DEFAULT_COLUMN_WIDTH,model.kind==='xlsx',gridZoom);canvas.style.width=(ROW_HEADER_WIDTH+columnOffsets[model.gridColumnCount])+'px';for(const item of canvas.children){if(!item.classList.contains('colhead')&&!item.classList.contains('cell'))continue;const column=Number(item.dataset.column);const width=item.classList.contains('cell')?columnOffsets[Number(item.dataset.mergeRight)+1]-columnOffsets[column]:columnOffsets[column+1]-columnOffsets[column];item.style.left=(ROW_HEADER_WIDTH+columnOffsets[column])+'px';item.style.width=width+'px';fitColumnBox(item,width);const handle=[...item.children].find(child=>child.classList.contains('col-resize-handle'));handle?.setAttribute('aria-valuenow',String(width))}}function resizeColumn(column,width){const value=gridColumnWidth(width);columnWidths.set(column,value);updateColumnGeometry()}function commitColumnWidth(column){if(model?.kind==='xlsx'&&!model.readOnly)vscode.postMessage({type:'column-width',column,width:gridColumnWidth(columnWidths.get(column)??model.defaultColumnWidth??DEFAULT_COLUMN_WIDTH),sheetIndex:model.sheetIndex})}function createConditionalIcon(icon){if(!icon||typeof icon.set!=='string'||!Number.isInteger(icon.id))return null;const svg=document.createElementNS('http:'+'//www.w3.org/2000/svg','svg'),set=icon.set,id=icon.id,count=Number(set[0])||3,traffic=set.includes('TrafficLights'),gray=set.includes('Gray'),colors=traffic?['#63BE7B','#FFEB84','#F8696B']:gray?['#666666','#999999','#CCCCCC']:['#63BE7B','#FFEB84','#F8696B'],color=colors[Math.min(colors.length-1,Math.floor(id/(Math.max(1,count-1))*2))]||colors[0];svg.setAttribute('viewBox','0 0 16 16');svg.setAttribute('width','14');svg.setAttribute('height','14');svg.setAttribute('aria-hidden','true');svg.style.cssText='position:absolute;left:2px;top:50%;transform:translateY(-50%);pointer-events:none;z-index:1';const shape=(name,attrs)=>{const item=document.createElementNS('http:'+'//www.w3.org/2000/svg',name);for(const[key,value]of Object.entries(attrs))item.setAttribute(key,value);svg.append(item);return item};if(set.includes('Arrows')){const points=['8,1 14,7 10,7 10,15 6,15 6,7 2,7','15,8 9,2 9,6 1,6 1,10 9,10 9,14','8,15 2,9 6,9 6,1 10,1 10,9 14,9','1,8 7,14 7,10 15,10 15,6 7,6 7,2'][id%4];shape('polygon',{points,fill:color,stroke:'#666','stroke-width':'.5'})}else if(traffic){shape('circle',{cx:'8',cy:'8',r:'6.2',fill:'#F2F2F2',stroke:'#666','stroke-width':'1'});shape('circle',{cx:'8',cy:'8',r:'4.2',fill:color})}else if(set.includes('Flags')){shape('path',{d:'M3 1v14M4 2h9l-2 3 2 3H4z',fill:color,stroke:'#555','stroke-width':'1','stroke-linejoin':'round'})}else if(set.includes('Symbols')){if(id===0)shape('path',{d:'M2 8l4 4 8-8',fill:'none',stroke:color,'stroke-width':'3','stroke-linecap':'round','stroke-linejoin':'round'});else if(id===1){shape('rect',{x:'6.5',y:'2',width:'3',height:'8',rx:'1.5',fill:color});shape('circle',{cx:'8',cy:'13',r:'1.5',fill:color})}else shape('path',{d:'M3 3l10 10M13 3L3 13',fill:'none',stroke:color,'stroke-width':'3','stroke-linecap':'round'})}else if(set.includes('Rating')||set.includes('Quarters')){if(set.includes('Rating')){for(let index=0;index<Math.max(1,count-id);index++){const x=2+index*(12/Math.max(1,count-id));shape('circle',{cx:String(x+1),cy:'8',r:'1.7',fill:color})}}else{shape('circle',{cx:'8',cy:'8',r:'6.5',fill:'#EEE',stroke:'#777','stroke-width':'1'});shape('path',{d:id===0?'M8 8V1.5A6.5 6.5 0 1 1 1.5 8z':id===1?'M8 8V1.5A6.5 6.5 0 0 1 14.5 8z':'M8 8V1.5A6.5 6.5 0 0 1 11.25 2.37z',fill:color})}}else{const polygon=count===3?'8,1 15,14 1,14':count===4?'8,1 15,8 8,15 1,8':'8,1 14,5 12,14 4,14 2,5';shape('polygon',{points:polygon,fill:color,stroke:'#666','stroke-width':'1'})}return svg}function syncOutlineLevelControls(){for(const axis of ['row','column']){const container=document.getElementById(axis==='row'?'rowOutlineLevels':'columnOutlineLevels'),groups=model?.[axis==='row'?'rowGroups':'columnGroups']||[],hidden=Boolean(model?.readOnly||!['xlsx','csv'].includes(model?.kind)||!groups.length);container.hidden=hidden;if(hidden){if(container.children.length)container.replaceChildren();continue}const maximum=Math.min(8,Math.max(...groups.map(group=>group.level))+1);if(container.children.length===maximum)continue;container.replaceChildren();for(let level=1;level<=maximum;level++){const button=document.createElement('button');button.type='button';button.className='outline-level';button.textContent=String(level);button.dataset.outlineLevel=String(level);button.title=t('show-outline-level')+' '+level;button.setAttribute('aria-label',button.title);button.onclick=()=>vscode.postMessage({type:'outline-groups',axis,action:'show-level',level,sheetIndex:model.sheetIndex||0});container.append(button)}}}function addOutlineBracket(parent,axis,groups,coordinate){for(const group of groups||[]){if(group.collapsed||coordinate<group.start||coordinate>group.end)continue;const bracket=document.createElement('span'),level=Math.min(7,Math.max(1,Number(group.level)||1));bracket.className='outline-bracket outline-'+axis+'-bracket';bracket.dataset.level=String(level);bracket.setAttribute('aria-hidden','true');if(axis==='row'){bracket.style.left=(2+(level-1)*2)+'px';if(coordinate===group.start)bracket.classList.add('outline-bracket-start');if(coordinate===group.end)bracket.classList.add('outline-bracket-end')}else{parent.classList.add('has-outline-brackets');bracket.style.top=(17+level)+'px';if(coordinate===group.start)bracket.classList.add('outline-bracket-start');if(coordinate===group.end)bracket.classList.add('outline-bracket-end')}parent.append(bracket)}}function addOutlineToggle(parent,axis,groups){if(model?.readOnly||!['xlsx','csv'].includes(model?.kind))return;const collapsed=groups.every(group=>group.collapsed),action=collapsed?'expand-summary':'collapse-summary',control=document.createElement('button'),ranges=groups.map(group=>{const first=axis==='row'?String(group.start+1):label(group.start+1),last=axis==='row'?String(group.end+1):label(group.end+1);return first===last?first:first+'–'+last}).join(', ');control.type='button';control.className='outline-toggle';control.textContent=collapsed?'+':'−';control.title=(collapsed?t(axis==='row'?'expand-rows':'expand-columns'):t(axis==='row'?'collapse-rows':'collapse-columns'))+' '+ranges;control.setAttribute('aria-label',control.title);control.setAttribute('aria-expanded',String(!collapsed));control.tabIndex=0;control.onmousedown=event=>{event.preventDefault();event.stopPropagation()};control.onclick=event=>{event.preventDefault();event.stopPropagation();vscode.postMessage({type:'outline-groups',axis,action,summary:groups[0].summary,sheetIndex:model.sheetIndex||0})};if(axis==='row')parent.classList.add('has-outline-toggle');parent.append(control)}function renderRange(range){if(range.requestId!==rangeRequest||range.sheetIndex!==model.sheetIndex)return;renderedRange={rowStart:range.rowStart,columnStart:range.columnStart,rowCount:range.rowCount,columnCount:range.columnCount,sheetIndex:range.sheetIndex,viewRows:range.viewRows?[...range.viewRows]:undefined,viewIndexes:range.viewIndexes?[...range.viewIndexes]:undefined,columnIndexes:range.columnIndexes?[...range.columnIndexes]:undefined};pendingRange=null;const canvas=document.getElementById('canvas');canvas.style.fontSize=(13*gridZoom)+'px';keepGridFocus();for(const item of [...canvas.children])if(!item.classList.contains('corner'))item.remove();for(let c=0;c<range.columnCount;c++){const column=range.columnIndexes?.[c]??(range.columnStart+c);const head=document.createElement('div');head.className='colhead';if(model.pageBreaks?.columns?.includes(column)){head.classList.add('manual-column-break');head.title=t('break-add-column')}head.dataset.column=String(column);const columnWidth=gridColumnWidth(columnWidths.get(column)??model.defaultColumnWidth??DEFAULT_COLUMN_WIDTH);head.style.width=(columnOffsets[column+1]-columnOffsets[column])+'px';fitColumnBox(head,columnOffsets[column+1]-columnOffsets[column]);head.textContent=label(column+1);head.style.left=(ROW_HEADER_WIDTH+columnOffsets[column])+'px';head.setAttribute('role','columnheader');head.setAttribute('aria-label',t('aria-column')+label(column+1));addOutlineBracket(head,'column',model.columnGroups||[],column);const outlineMatches=(model.columnGroups||[]).filter(group=>group.summary===column);if(outlineMatches.length)addOutlineToggle(head,'column',outlineMatches);const resizeHandle=document.createElement('div');resizeHandle.className='col-resize-handle';resizeHandle.setAttribute('role','separator');resizeHandle.setAttribute('aria-orientation','vertical');resizeHandle.setAttribute('aria-label',t('resize-column')+label(column+1));resizeHandle.setAttribute('aria-valuemin',model.kind==='xlsx'?'1':'48');resizeHandle.setAttribute('aria-valuemax',model.kind==='xlsx'?'1785':'640');resizeHandle.setAttribute('aria-valuenow',String(columnWidth));resizeHandle.setAttribute('tabindex','0');resizeHandle.onmousedown=e=>{e.preventDefault();e.stopPropagation();columnResize={column,startX:e.clientX,startWidth:columnWidth}};resizeHandle.onclick=e=>e.stopPropagation();resizeHandle.onkeydown=e=>{if(e.key==='ArrowLeft'||e.key==='ArrowRight'){e.preventDefault();e.stopPropagation();resizeColumn(column,(columnWidths.get(column)??columnWidth)+(e.key==='ArrowRight'?10:-10));commitColumnWidth(column);requestRange();resizeHandle.setAttribute('aria-valuenow',String(gridColumnWidth(columnWidths.get(column))))}};head.onmousedown=e=>{if(e.target===resizeHandle)return;e.preventDefault();dragging=true;dragAxis='column';dragged=false;dragPointer={x:e.clientX,y:e.clientY};dragMoved=false;shiftSelecting=e.shiftKey;selectAxis('column',column,e.shiftKey);startCellDrag()};head.onmouseenter=e=>{if(e.target!==resizeHandle&&dragging&&dragAxis==='column'){dragged=true;selectAxis('column',column,true)}};head.onclick=e=>{if(e.target===resizeHandle)return;if(dragged){dragged=false;shiftSelecting=false;return}selectAxis('column',column,e.shiftKey)};head.append(resizeHandle);canvas.append(head)}for(let r=0;r<range.rowCount;r++){const row=range.viewRows?.[r]??rowView.rowAt(range.rowStart+r);const displayIndex=range.viewIndexes?.[r]??(range.rowStart+r);const rowHead=document.createElement('div');rowHead.className='rowhead';if(model.pageBreaks?.rows?.includes(row)){rowHead.classList.add('manual-row-break');rowHead.title=t('break-add-row')}rowHead.textContent=String(row+1);rowHead.style.height=rowMetrics.height(displayIndex)+'px';rowHead.style.top=(30+rowMetrics.top(displayIndex))+'px';rowHead.dataset.viewRow=String(displayIndex);rowHead.style.left='0px';rowHead.setAttribute('aria-label',t('aria-row')+(row+1));rowHead.dataset.row=String(row);addOutlineBracket(rowHead,'row',model.rowGroups||[],row);const outlineMatches=(model.rowGroups||[]).filter(group=>group.summary===row);if(outlineMatches.length)addOutlineToggle(rowHead,'row',outlineMatches);let rowHandle=null;if(model.kind==='xlsx'&&!model.readOnly){rowHandle=document.createElement('div');rowHandle.className='row-resize-handle';rowHandle.dataset.row=String(row);rowHandle.setAttribute('role','separator');rowHandle.setAttribute('aria-orientation','horizontal');rowHandle.setAttribute('aria-label',t('resize-row')+(row+1));rowHandle.setAttribute('aria-valuemin','.75');rowHandle.setAttribute('aria-valuemax','409.5');rowHandle.setAttribute('aria-valuenow',String(Math.round(rowMetrics.height(displayIndex)*.75*100)/100));rowHandle.tabIndex=0;rowHandle.onmousedown=e=>{e.preventDefault();e.stopPropagation();const bounds=rowResizeBounds(row);if(bounds.endRow-bounds.startRow>=10000&&!(bounds.startRow===0&&bounds.endRow===model.gridRowCount-1)){document.getElementById('notice').textContent=t('error-row-height-range');return}rowResize={...bounds,sheetIndex:model.sheetIndex,row,startY:e.clientY,startHeight:rowMetrics.height(displayIndex)/gridZoom,height:rowMetrics.height(displayIndex)/gridZoom,heights:model.rowHeights,defaultHeight:model.defaultRowHeight,cursor:document.body.style.cursor||'',moved:false};document.body.style.cursor='ns-resize'};rowHandle.onclick=e=>e.stopPropagation();rowHandle.ondblclick=e=>{e.preventDefault();e.stopPropagation();cancelRowResize();vscode.postMessage({type:'row-auto-fit',...rowResizeBounds(row),sheetIndex:model.sheetIndex||0})};rowHandle.onkeydown=e=>{if(e.key==='Escape'){e.preventDefault();e.stopPropagation();cancelRowResize();return}if(e.key==='ArrowUp'||e.key==='ArrowDown'){e.preventDefault();e.stopPropagation();const bounds=rowResizeBounds(row);if(bounds.endRow-bounds.startRow>=10000&&!(bounds.startRow===0&&bounds.endRow===model.gridRowCount-1))return;rowBoundaryFocus={row,sheetIndex:model.sheetIndex,source:rowHandle};const height=previewRowHeight(bounds,rowMetrics.height(rowView.indexOf(row))+(e.key==='ArrowDown'?4/3:-4/3));rowHandle.focus({preventScroll:true});commitRowHeight(bounds,height)}};rowHead.append(rowHandle)}rowHead.onmousedown=e=>{if(e.target===rowHandle)return;e.preventDefault();dragging=true;dragAxis='row';dragged=false;dragPointer={x:e.clientX,y:e.clientY};dragMoved=false;shiftSelecting=e.shiftKey;selectAxis('row',row,e.shiftKey);startCellDrag()};rowHead.onmouseenter=()=>{if(dragging&&dragAxis==='row'){dragged=true;selectAxis('row',row,true)}};rowHead.onclick=e=>{if(e.target===rowHandle)return;if(dragged){dragged=false;shiftSelecting=false;return}selectAxis('row',row,e.shiftKey)};fitRowBox(rowHead,Number.parseFloat(rowHead.style.height));canvas.append(rowHead);for(let c=0;c<range.columnCount;c++){const column=range.columnIndexes?.[c]??(range.columnStart+c);const data=range.rows[r][c];const merged=model.mergedCells?.find(range=>row>=range.top&&row<=range.bottom&&column>=range.left&&column<=range.right);if(merged&&(row!==merged.top||column!==merged.left))continue;const inRange=row>=Math.min(selectedRange.startRow,selectedRange.endRow)&&row<=Math.max(selectedRange.startRow,selectedRange.endRow)&&column>=Math.min(selectedRange.startColumn,selectedRange.endColumn)&&column<=Math.max(selectedRange.startColumn,selectedRange.endColumn);const cell=document.createElement('div');cell.className='cell'+(selected.row===row&&selected.column===column?' selected':'')+(inRange?' in-range':'');cell.classList.toggle('numeric',data?.kind==='number');cell.classList.toggle('hyperlink',Boolean(data?.hyperlink));const tableStyle=data?.tableStyle;if(tableStyle){if(tableStyle.fill)cell.style.backgroundColor=tableStyle.fill;if(tableStyle.color)cell.style.color=tableStyle.color}const style=data?.style;if(style){if(style.color)cell.style.color=style.color;if(style.fill)cell.style.backgroundColor=style.fill;if(style.fontSize)cell.style.fontSize=(style.fontSize*gridZoom)+'pt';if(style.fontFamily)cell.style.fontFamily='\"'+style.fontFamily.replaceAll('\"','')+'\"';if(style.bold)cell.style.fontWeight='700';if(style.italic)cell.style.fontStyle='italic';if(style.underline||style.strike||data?.hyperlink)cell.style.textDecoration=[style.underline||data?.hyperlink?'underline':'',style.strike?'line-through':''].filter(Boolean).join(' ');if(style.horizontal)cell.style.textAlign=style.horizontal;if(style.wrapText){cell.style.whiteSpace='pre-wrap';cell.style.overflowWrap='anywhere'}if(style.vertical){cell.style.paddingTop=style.vertical==='top'?'0px':style.vertical==='bottom'?'8px':'4px';cell.style.paddingBottom=style.vertical==='bottom'?'0px':style.vertical==='top'?'8px':'4px'}if(style.dataBar){const bar=style.dataBar,start=bar.start??0,end=bar.end??bar.width;if(bar.axis!==undefined){cell.style.backgroundImage=bar.width===0?'none':bar.gradient?'linear-gradient(90deg,transparent 0%,transparent '+start+'%,'+bar.color+' '+start+'%,'+bar.color+' '+end+'%,transparent '+end+'%,transparent 100%)':'linear-gradient(90deg,transparent '+start+'%,'+bar.color+' '+start+'%,'+bar.color+' '+end+'%,transparent '+end+'%)'}else cell.style.backgroundImage=bar.gradient?'linear-gradient(90deg,'+bar.color+' 0%,transparent '+bar.width+'%)':'linear-gradient(90deg,'+bar.color+' 0%,'+bar.color+' '+bar.width+'%,transparent '+bar.width+'%)';if(!bar.showValue)cell.style.color='transparent'}if(style.borders)for(const [side,border] of Object.entries(style.borders))cell.style['border'+side[0].toUpperCase()+side.slice(1)]='1px '+(border.line==='double'?'double':border.line==='dashed'?'dashed':border.line==='dotted'?'dotted':'solid')+' '+border.color}cell.id='cell-'+row+'-'+column;cell.tabIndex=selected.row===row&&selected.column===column?0:-1;cell.setAttribute('role','gridcell');cell.setAttribute('aria-selected',String(inRange));cell.setAttribute('aria-rowindex',String(row+1));cell.setAttribute('aria-colindex',String(column+1));cell.dataset.row=String(row);cell.dataset.column=String(column);cell.dataset.mergeRight=String(merged?.right??column);cell.style.height=rowMetrics.height(displayIndex)+'px';cell.style.top=(30+rowMetrics.top(displayIndex))+'px';cell.style.left=(ROW_HEADER_WIDTH+columnOffsets[column])+'px';cell.style.width=(columnOffsets[(merged?.right??column)+1]-columnOffsets[column])+'px';if(merged)cell.style.height=(rowMetrics.top(rowDisplayIndex(merged.bottom+1))-rowMetrics.top(displayIndex))+'px';const frozenRow=row<(model.freezePane?.rows||0),frozenColumn=column<(model.freezePane?.columns||0);if(merged)cell.style.zIndex='4';else if(frozenRow||frozenColumn)cell.style.zIndex=frozenRow&&frozenColumn?'3':'2';cell.textContent=data?.hyperlink?.display??data?.displayValue??data?.value??'';if(data?.style?.icon){const iconNode=createConditionalIcon(data.style.icon);if(iconNode)cell.prepend(iconNode);if(data.style.icon.showValue===false)cell.style.color='transparent';else cell.style.paddingLeft='18px'}const prompt=data?.validation?.showInputMessage&&[data.validation.promptTitle,data.validation.prompt].filter(Boolean).join('\\n');cell.title=[cell.textContent,prompt,data?.hyperlink?(data.hyperlink.tooltip||t('hyperlink-follow')):null].filter(Boolean).join('\\n');if(prompt)cell.setAttribute('aria-description',prompt);const borders=data?.styleProperties?.borders||{};for(const side of ['top','right','bottom','left']){const edge=borders[side];if(edge){const color=edge.color||'#000000',style=edge.line==='double'?'double':edge.line==='dashed'?'dashed':edge.line==='dotted'?'dotted':edge.line==='medium'||edge.line==='thick'?'solid':'solid',width=edge.line==='thick'?'3px':edge.line==='medium'?'2px':'1px';cell.style['border'+side[0].toUpperCase()+side.slice(1)]=width+' '+style+' '+color}}cell.__jinSheetData=data;cell.onmousedown=e=>{if(e.target&&e.target!==cell)return;dragAxis=null;if(formulaDirty)commitFormulaEdit();e.preventDefault();dragging=true;dragged=false;dragPointer={x:e.clientX,y:e.clientY};dragMoved=false;shiftSelecting=e.shiftKey;select(row,column,data,e.shiftKey);startCellDrag()};cell.onmouseenter=()=>{if(dragging){dragged=true;select(row,column,data,true)}};cell.onclick=e=>{if(e?.target&&e.target!==cell)return;if(data?.hyperlink&&(e.ctrlKey||e.metaKey)&&!e.shiftKey){e.preventDefault();e.stopPropagation();select(row,column,data);vscode.postMessage({type:'follow-hyperlink',row,column,sheetIndex:model.sheetIndex||0});return}if(dragged){dragged=false;shiftSelecting=false;return}if(shiftSelecting){select(row,column,data,true);shiftSelecting=false}else select(row,column,data)};cell.ondblclick=e=>{if(e.target&&e.target!==cell)return;const position=document.caretPositionFromPoint?.(e?.clientX,e?.clientY)||document.caretRangeFromPoint?.(e?.clientX,e?.clientY);const caretOffset=position&&cell.contains(position.offsetNode)?position.offset:undefined;edit(cell,row,column,data,undefined,{selectAll:false,caretOffset})};cell.onkeydown=e=>{if(e.target&&e.target!==cell)return;if(data?.hyperlink&&(e.ctrlKey||e.metaKey)&&e.key==='Enter'){e.preventDefault();e.stopPropagation();vscode.postMessage({type:'follow-hyperlink',row,column,sheetIndex:model.sheetIndex||0});return}if(e.key==='F2'){e.preventDefault();e.stopPropagation();edit(cell,row,column,data,undefined,{selectAll:false})}else if((e.isComposing||e.key==='Process'||e.key==='Unidentified')&&!e.ctrlKey&&!e.metaKey&&!e.altKey){edit(cell,row,column,data,'')}else if(e.key.length===1&&!e.ctrlKey&&!e.metaKey&&!e.altKey){e.preventDefault();e.stopPropagation();edit(cell,row,column,data,e.key)}};if(!model.readOnly&&row===Math.max(selectedRange.startRow,selectedRange.endRow)&&column===Math.max(selectedRange.startColumn,selectedRange.endColumn)){const handle=document.createElement("button");handle.type="button";handle.className="fill-handle";handle.title=t("fill-selection");handle.setAttribute("aria-label",t("fill-selection"));handle.tabIndex=-1;handle.onmousedown=e=>{e.preventDefault();e.stopPropagation();fillDrag={sourceRange:{...selectedRange}};dragging=true;dragAxis="fill";dragPointer={x:e.clientX,y:e.clientY};dragMoved=false;startCellDrag()};cell.append(handle)}if(model.kind==='xlsx'){cell.style.display='flex';cell.style.flexDirection='column';cell.style.justifyContent=data?.style?.vertical==='top'?'flex-start':data?.style?.vertical==='center'?'center':'flex-end';cell.style.paddingTop='0px';cell.style.paddingBottom='0px'}fitColumnBox(cell,Number.parseFloat(cell.style.width));fitRowBox(cell,Number.parseFloat(cell.style.height));canvas.append(cell)}}const selectedRowIndex=(range.viewRows||Array.from({length:range.rowCount},(_,index)=>rowView.rowAt(range.rowStart+index))).indexOf(selected.row),selectedColumnIndex=(range.columnIndexes||Array.from({length:range.columnCount},(_,index)=>range.columnStart+index)).indexOf(selected.column);if(selectedRowIndex>=0&&selectedColumnIndex>=0){const preservedAnchor={...selectionAnchor},preservedRange={...selectedRange};select(selected.row,selected.column,range.rows[selectedRowIndex]?.[selectedColumnIndex],true,(document.activeElement===document.body||document.getElementById('grid').contains(document.activeElement))&&!rowBoundaryFocus);selectionAnchor=preservedAnchor;selectedRange=preservedRange;const top=Math.min(selectedRange.startRow,selectedRange.endRow),bottom=Math.max(selectedRange.startRow,selectedRange.endRow),left=Math.min(selectedRange.startColumn,selectedRange.endColumn),right=Math.max(selectedRange.startColumn,selectedRange.endColumn);for(const item of document.querySelectorAll('.cell')){const itemRow=Number(item.dataset.row),itemColumn=Number(item.dataset.column),inRange=itemRow>=top&&itemRow<=bottom&&itemColumn>=left&&itemColumn<=right;item.classList.toggle('in-range',inRange);item.setAttribute('aria-selected',String(inRange))}}refreshSearchHighlights();refreshCutHighlights();syncHeadings();if(rowBoundaryFocus?.restore){const head=[...canvas.children].find(item=>item.classList.contains('rowhead')&&rowView.rowAt(Number(item.dataset.viewRow))===rowBoundaryFocus.row),handle=head&&[...head.children].find(item=>item.classList.contains('row-resize-handle'));handle?.focus({preventScroll:true});rowBoundaryFocus=null}}
  function syncHeadings(){const grid=document.getElementById('grid'),canvas=document.getElementById('canvas'),freezeRows=model?.freezePane?.rows||0,freezeColumns=model?.freezePane?.columns||0;for(const item of canvas.children){if(item.classList.contains('colhead'))item.style.transform='translateY('+grid.scrollTop+'px)';else if(item.classList.contains('rowhead')){const frozen=rowView.rowAt(Number(item.dataset.viewRow))<freezeRows;item.style.transform=frozen?'translate('+grid.scrollLeft+'px,'+grid.scrollTop+'px)':'translateX('+grid.scrollLeft+'px)';if(frozen)item.style.zIndex='3'}else if(item.classList.contains('cell')){const frozenRow=Number(item.dataset.row)<freezeRows,frozenColumn=Number(item.dataset.column)<freezeColumns;if(frozenRow||frozenColumn)item.style.transform='translate('+(frozenColumn?grid.scrollLeft:0)+'px,'+(frozenRow?grid.scrollTop:0)+'px)';else item.style.transform=''}else if(item.classList.contains('corner'))item.style.transform='translate('+grid.scrollLeft+'px,'+grid.scrollTop+'px)'}}
  let summaryRequest=0,summaryTimer;
  function refreshSelectionSummary(){clearTimeout(summaryTimer);const requestId=++summaryRequest;summaryTimer=setTimeout(()=>vscode.postMessage({type:'selection-summary',requestId,startRow:selectedRange.startRow,endRow:selectedRange.endRow,startColumn:selectedRange.startColumn,endColumn:selectedRange.endColumn,sheetIndex:model?.sheetIndex||0}),80)}
  let calculatedSequence=0,pendingCalculatedColumn=null;const autoCalculatedColumns=document.getElementById('autoCalculatedColumns'),autoExpandTables=document.getElementById('autoExpandTables');autoCalculatedColumns.checked=true;autoExpandTables.checked=true;
  function syncCalculatedControls(){const table=selectedTotalsTable(),apply=document.getElementById('applyCalculatedColumn'),clear=document.getElementById('clearCalculatedColumn'),scope=document.getElementById('calculatedColumnScope'),labelElement=document.getElementById('calculatedColumnLabel');if(!apply||!clear||!scope)return;if(pendingCalculatedColumn&&!model?.tables?.some(item=>item.name===pendingCalculatedColumn.tableName))pendingCalculatedColumn=null;const disabled=Boolean(model?.kind!=='xlsx'||model?.readOnly||!table||pendingCalculatedColumn||pendingTableTotals),index=table?selected.column-table.range.left:0,base=table?.calculatedColumns?.[index],data=document.querySelector('#cell-'+selected.row+'-'+selected.column)?.__jinSheetData,body=table&&selected.row>=table.range.top+(table.headerRowCount||0)&&selected.row<=table.range.bottom-(table.totalsRowCount||0);apply.disabled=Boolean(disabled||!(base||body&&data?.formula));clear.disabled=Boolean(disabled||!base);scope.disabled=apply.disabled;autoCalculatedColumns.disabled=Boolean(model?.kind!=='xlsx'||model?.readOnly||!(model?.tables?.length));autoExpandTables.disabled=Boolean(model?.kind!=='xlsx'||model?.readOnly||!(model?.tables?.length));autoExpandTables.title=t('auto-table-title');labelElement.textContent=table?t('calculated-column-label')+(table.columns?.[index]??label(selected.column+1)):'';labelElement.title=labelElement.textContent}
  function submitCalculatedColumn(remove=false){const table=selectedTotalsTable();if(!table||pendingCalculatedColumn||pendingTableTotals||model?.readOnly)return;const requestId=++calculatedSequence;pendingCalculatedColumn={requestId,sheetIndex:model.sheetIndex,tableName:table.name,column:selected.column};vscode.postMessage({type:'calculated-column',requestId,sheetIndex:model.sheetIndex,tableName:table.name,row:selected.row,column:selected.column,mode:document.getElementById('calculatedColumnScope').value||'preserve',remove});syncCalculatedControls();syncTableTotalsControls()}
  document.getElementById('applyCalculatedColumn').onclick=()=>submitCalculatedColumn(false);document.getElementById('clearCalculatedColumn').onclick=()=>submitCalculatedColumn(true);
  let tableTotalsSequence=0,pendingTableTotals=null;
  function selectedTotalsTable(){return(model?.tables||[]).find(table=>selected.row>=table.range.top&&selected.row<=table.range.bottom&&selected.column>=table.range.left&&selected.column<=table.range.right)}
  function syncTableTotalsControls(){const button=document.getElementById('tableTotals'),input=document.getElementById('tableTotalsFunction'),apply=document.getElementById('applyTableTotalFunction'),columnLabel=document.getElementById('tableTotalsColumn');if(!button||!input||!apply)return;if(pendingTableTotals&&!model?.tables?.some(table=>table.name===pendingTableTotals.tableName))pendingTableTotals=null;const table=selectedTotalsTable(),disabled=Boolean(model?.readOnly||model?.kind!=='xlsx'||!table||pendingTableTotals||pendingCalculatedColumn);button.disabled=disabled;button.setAttribute('aria-pressed',String(Boolean(table?.totalsRowCount)));input.disabled=Boolean(disabled||!table?.totalsRowCount);const index=table?selected.column-table.range.left:0,current=table?.totalColumns?.[index]?.function||'none';if(!pendingTableTotals&&document.activeElement!==input)input.value=['none','sum','average','count','countNums','min','max','stdDev','var'].includes(current)?current:'custom';apply.disabled=Boolean(input.disabled||input.value==='custom');columnLabel.textContent=table?t('table-totals-column')+(table.columns?.[index]??label(selected.column+1)):'';columnLabel.title=columnLabel.textContent}
  function submitTableTotals(kind){const table=selectedTotalsTable();if(!table||pendingTableTotals||model?.readOnly)return;const requestId=++tableTotalsSequence;pendingTableTotals={requestId,sheetIndex:model.sheetIndex,tableName:table.name};const payload={type:'table-totals',requestId,sheetIndex:model.sheetIndex,tableName:table.name,row:selected.row,column:selected.column};if(kind==='toggle')payload.enabled=!table.totalsRowCount;else{payload.function=document.getElementById('tableTotalsFunction').value;if(payload.function==='custom'){pendingTableTotals=null;return}pendingTableTotals.function=payload.function}vscode.postMessage(payload);syncTableTotalsControls()}
  document.getElementById('tableTotals').onclick=()=>submitTableTotals('toggle');document.getElementById('applyTableTotalFunction').onclick=()=>submitTableTotals('function');document.getElementById('tableTotalsFunction').onchange=()=>{document.getElementById('applyTableTotalFunction').disabled=Boolean(document.getElementById('tableTotalsFunction').disabled||document.getElementById('tableTotalsFunction').value==='custom')};
  function syncMergeButtons(){const merge=document.getElementById('mergeCells'),unmerge=document.getElementById('unmergeCells');if(!merge||!unmerge)return;const top=Math.min(selectedRange.startRow,selectedRange.endRow),bottom=Math.max(selectedRange.startRow,selectedRange.endRow),left=Math.min(selectedRange.startColumn,selectedRange.endColumn),right=Math.max(selectedRange.startColumn,selectedRange.endColumn),hasMerged=(model?.mergedCells||[]).some(range=>top<=range.bottom&&bottom>=range.top&&left<=range.right&&right>=range.left);merge.disabled=Boolean(model?.readOnly||model?.kind!=='xlsx'||top===bottom&&left===right||hasMerged);unmerge.disabled=Boolean(model?.readOnly||model?.kind!=='xlsx'||!hasMerged)}let createTableSequence=0,pendingCreateTable=null,tableStyleSequence=0,pendingTableStyle=null,selectedStyleTableKey=null,currentCustomTableStyle=null,tableStyleSelectionTouched=false,tableStyleOptionsTouched=false;
  function selectedTableStyleTarget(){return(model?.tables||[]).find(table=>selected.row>=table.range.top&&selected.row<=table.range.bottom&&selected.column>=table.range.left&&selected.column<=table.range.right)}
  function syncTableStyleControls(){const select=document.getElementById('tableStyleName'),apply=document.getElementById('applyTableStyle'),note=document.getElementById('tableStyleNote');if(!select||!apply)return;if(pendingTableStyle&&pendingTableStyle.sheetIndex!==model?.sheetIndex)pendingTableStyle=null;const table=selectedTableStyleTarget(),key=table?model.sheetIndex+':'+table.name:null;if(key!==selectedStyleTableKey){selectedStyleTableKey=key;currentCustomTableStyle=null;tableStyleSelectionTouched=false;tableStyleOptionsTouched=false}const info=table?.styleInfo||null;if(!table){select.value='';for(const id of ['tableStyleFirstColumn','tableStyleLastColumn','tableStyleRows','tableStyleColumns'])document.getElementById(id).checked=false;if(note)note.textContent=''}else{if(!tableStyleSelectionTouched){const name=info?.name||'';currentCustomTableStyle=name&&!tableStyleNames.includes(name)?name:null;select.value=tableStyleNames.includes(name)?name:'';if(note)note.textContent=currentCustomTableStyle?t('table-style-custom')+' '+currentCustomTableStyle:''}if(!tableStyleOptionsTouched){document.getElementById('tableStyleFirstColumn').checked=Boolean(info?.showFirstColumn);document.getElementById('tableStyleLastColumn').checked=Boolean(info?.showLastColumn);document.getElementById('tableStyleRows').checked=Boolean(info?.showRowStripes);document.getElementById('tableStyleColumns').checked=Boolean(info?.showColumnStripes)}}const disabled=Boolean(model?.kind!=='xlsx'||model?.readOnly||!table||pendingTableStyle);select.disabled=disabled;for(const id of ['tableStyleFirstColumn','tableStyleLastColumn','tableStyleRows','tableStyleColumns'])document.getElementById(id).disabled=disabled;apply.disabled=disabled}
  function submitTableStyle(){const table=selectedTableStyleTarget();if(!table||pendingTableStyle||model?.readOnly)return;commitFormulaEdit();const selectedName=document.getElementById('tableStyleName').value,name=selectedName||(!tableStyleSelectionTouched?currentCustomTableStyle:null),styleInfo={name,showFirstColumn:document.getElementById('tableStyleFirstColumn').checked,showLastColumn:document.getElementById('tableStyleLastColumn').checked,showRowStripes:document.getElementById('tableStyleRows').checked,showColumnStripes:document.getElementById('tableStyleColumns').checked};if(!name)for(const key of ['showFirstColumn','showLastColumn','showRowStripes','showColumnStripes'])styleInfo[key]=false;pendingTableStyle={requestId:++tableStyleSequence,sheetIndex:model.sheetIndex,tableName:table.name,styleInfo};vscode.postMessage({type:'table-style',...pendingTableStyle,styleInfo});syncTableStyleControls()}
  function syncCreateTableControls(){const button=document.getElementById('createTable'),input=document.getElementById('newTableName');if(!button||!input)return;if(pendingCreateTable&&pendingCreateTable.sheetIndex!==model?.sheetIndex)pendingCreateTable=null;const range={top:Math.min(selectedRange.startRow,selectedRange.endRow),bottom:Math.max(selectedRange.startRow,selectedRange.endRow),left:Math.min(selectedRange.startColumn,selectedRange.endColumn),right:Math.max(selectedRange.startColumn,selectedRange.endColumn)},overlap=(model?.tables||[]).some(table=>range.top<=table.range.bottom&&range.bottom>=table.range.top&&range.left<=table.range.right&&range.right>=table.range.left);button.disabled=Boolean(model?.readOnly||model?.kind!=='xlsx'||pendingCreateTable||overlap||(range.bottom-range.top+1)*(range.right-range.left+1)>100000);input.disabled=Boolean(model?.readOnly||model?.kind!=='xlsx'||pendingCreateTable)}
  function submitCreateTable(){if(document.getElementById('createTable').disabled||pendingCreateTable)return;commitFormulaEdit();const range={top:Math.min(selectedRange.startRow,selectedRange.endRow),bottom:Math.max(selectedRange.startRow,selectedRange.endRow),left:Math.min(selectedRange.startColumn,selectedRange.endColumn),right:Math.max(selectedRange.startColumn,selectedRange.endColumn)},name=document.getElementById('newTableName').value.trim();pendingCreateTable={requestId:++createTableSequence,sheetIndex:model.sheetIndex,range,name};vscode.postMessage({type:'create-table',...pendingCreateTable});syncCreateTableControls()}
  function syncTableResizeButton(){syncCreateTableControls();syncTableStyleControls();syncCalculatedControls();syncTableTotalsControls();const button=document.getElementById('resizeTable');if(!button)return;const range={top:Math.min(selectedRange.startRow,selectedRange.endRow),bottom:Math.max(selectedRange.startRow,selectedRange.endRow),left:Math.min(selectedRange.startColumn,selectedRange.endColumn),right:Math.max(selectedRange.startColumn,selectedRange.endColumn)},table=(model?.tables||[]).find(item=>item.range.top===range.top&&range.left<=item.range.right&&range.right>=item.range.left&&range.bottom>=range.top+Math.max(1,(item.headerRowCount||0)+(item.totalsRowCount||0))-1&&range.right>=range.left);button.disabled=Boolean(model?.readOnly||model?.kind!=='xlsx'||!table);button.title=t('resize-table-title')}function syncChartButton(){const button=document.getElementById('insertChart'),type=document.getElementById('chartType');if(!button)return;const height=Math.abs(selectedRange.endRow-selectedRange.startRow)+1,width=Math.abs(selectedRange.endColumn-selectedRange.startColumn)+1;button.disabled=Boolean(model?.readOnly||model?.kind!=='xlsx'||!model.chartInsertionAvailable||height<2||width<2);if(type)type.disabled=button.disabled;for(const source of document.querySelectorAll('[data-chart-source]'))source.disabled=Boolean(model?.readOnly||model?.kind!=='xlsx'||height<2||width<2)}function syncRowHeightControls(){syncPageBreakControls();const input=document.getElementById('rowHeight'),disabled=Boolean(model?.readOnly||model?.kind!=='xlsx');for(const id of ['rowHeight','applyRowHeight','resetRowHeight','wrapText','autoFitRow'])document.getElementById(id).disabled=disabled;const selectedData=document.querySelector('[data-row="'+selected.row+'"][data-column="'+selected.column+'"]')?.__jinSheetData;document.getElementById('wrapText').setAttribute('aria-pressed',String(Boolean(selectedData?.style?.wrapText)));if(input&&document.activeElement!==input){const pixels=model?.rowHeights?.[selected.row]??model?.defaultRowHeight??26;input.value=String(Math.round(pixels*.75*100)/100)}}function select(r,c,data,extend=false,focusCell=true){const merged=model?.mergedCells?.find(range=>r>=range.top&&r<=range.bottom&&c>=range.left&&c<=range.right);if(merged){r=merged.top;c=merged.left;data=document.querySelector("#cell-"+r+"-"+c)?.__jinSheetData??data}if((selected.row!==r||selected.column!==c)&&formulaDirty)commitFormulaEdit();if(!extend)navigationRequest++;if(!extend||selected.row!==r||selected.column!==c){currentRegionRequest++;currentRegionStage=0}selected={row:r,column:c};if(!extend){selectionAnchor={row:r,column:c};selectedRange={startRow:r,endRow:r,startColumn:c,endColumn:c}}else selectedRange={startRow:selectionAnchor.row,endRow:r,startColumn:selectionAnchor.column,endColumn:c};document.querySelector('.cell.selected')?.classList.remove('selected');document.querySelector('[data-row="'+r+'"][data-column="'+c+'"]')?.classList.add('selected');for(const cell of document.querySelectorAll('.cell')){const row=Number(cell.dataset.row),column=Number(cell.dataset.column),inRange=row>=Math.min(selectedRange.startRow,selectedRange.endRow)&&row<=Math.max(selectedRange.startRow,selectedRange.endRow)&&column>=Math.min(selectedRange.startColumn,selectedRange.endColumn)&&column<=Math.max(selectedRange.startColumn,selectedRange.endColumn);cell.classList.toggle('in-range',inRange);cell.setAttribute('aria-selected',String(inRange));cell.tabIndex=row===r&&column===c?0:-1}if(focusCell)document.querySelector('[data-row="'+r+'"][data-column="'+c+'"]')?.focus({preventScroll:true});document.getElementById('address').value=label(c+1)+(r+1);document.getElementById('address').setAttribute('aria-invalid','false');document.getElementById('goToStatus').textContent='';const formula=document.getElementById('formula');if(!formulaDirty){formula.value=data?.formula?('='+data.formula):(data?.value??'');formula.title=data?.formula?formula.value+'  · '+t('cache-value')+': '+(data.cachedValue??''):formula.value;formula.dataset.originalValue=formula.value;const list=document.getElementById('cellValidationOptions');if(list)list.replaceChildren(...(data?.validation&&!data.validation.showDropDown?data.validation.values:[]).map(value=>{const option=document.createElement('option');option.value=value;return option}))}formula.disabled=Boolean(model.readOnly||(data?.kind==='error'&&!data?.hasFormula));if(!dragging)syncFillHandle();syncMergeButtons();syncTableResizeButton();syncChartButton();syncRowHeightControls();if(typeof syncHyperlinkControls==='function')syncHyperlinkControls();refreshSelectionSummary()}
function syncFillHandle(){for(const handle of document.querySelectorAll('.fill-handle'))handle.remove();if(!model||model.readOnly)return;const row=Math.max(selectedRange.startRow,selectedRange.endRow),column=Math.max(selectedRange.startColumn,selectedRange.endColumn),cell=document.querySelector('[data-row="'+row+'"][data-column="'+column+'"]');if(!cell)return;const handle=document.createElement('button');handle.type='button';handle.className='fill-handle';handle.title=t('fill-selection');handle.setAttribute('aria-label',t('fill-selection'));handle.tabIndex=-1;handle.onmousedown=e=>{e.preventDefault();e.stopPropagation();fillDrag={sourceRange:{...selectedRange}};dragging=true;dragAxis='fill';dragPointer={x:e.clientX,y:e.clientY};dragMoved=false;startCellDrag()};cell.append(handle)}
  function selectAxis(axis,index,extend=false){const limit=axis==='row'?model?.gridRowCount:model?.gridColumnCount;if(!Number.isInteger(index)||index<0||index>=limit)return;const priorAnchor=selectionAnchor,anchorIndex=extend&&priorAnchor.axis===axis?(axis==='row'?priorAnchor.row:priorAnchor.column):index;const row=axis==='row'?index:selected.row,column=axis==='column'?index:selected.column,cell=document.querySelector('[data-row="'+row+'"][data-column="'+column+'"]');select(row,column,cell?.__jinSheetData,false);selectionAnchor=axis==='row'?{row:anchorIndex,column:0,axis}:{row:0,column:anchorIndex,axis};selectedRange=axis==='row'?{startRow:anchorIndex,endRow:index,startColumn:0,endColumn:Math.max(0,(model.gridColumnCount||model.columnCount||1)-1)}:{startRow:0,endRow:Math.max(0,(model.gridRowCount||model.rowCount||1)-1),startColumn:anchorIndex,endColumn:index};const top=Math.min(selectedRange.startRow,selectedRange.endRow),bottom=Math.max(selectedRange.startRow,selectedRange.endRow),left=Math.min(selectedRange.startColumn,selectedRange.endColumn),right=Math.max(selectedRange.startColumn,selectedRange.endColumn);for(const item of document.querySelectorAll('.cell')){const itemRow=Number(item.dataset.row),itemColumn=Number(item.dataset.column),inRange=itemRow>=top&&itemRow<=bottom&&itemColumn>=left&&itemColumn<=right;item.classList.toggle('in-range',inRange);item.setAttribute('aria-selected',String(inRange))}refreshSelectionSummary()}
  function edit(cell,r,c,data,initial,options={}){if(!cell||model.readOnly||(data?.kind==='error'&&!data?.hasFormula))return;pendingCutRange=null;refreshCutHighlights();const input=document.createElement('textarea');input.className='cell-editor';input.rows=1;input.wrap='off';input.value=initial??((data?.hasFormula||data?.formula)?'='+data.formula:(data?.value??''));input.style.width=Math.max(24,gridColumnWidth(columnWidths.get(c)??model.defaultColumnWidth??DEFAULT_COLUMN_WIDTH)-12)+'px';const editBox={overflow:cell.style.overflow,zIndex:cell.style.zIndex,justifyContent:cell.style.justifyContent};cell.style.overflow='visible';cell.style.zIndex='8';cell.style.justifyContent='flex-start';cell.replaceChildren(input);input.focus();if(initial!==undefined)input.setSelectionRange(input.value.length,input.value.length);else if(data?.hasFormula||data?.formula)input.setSelectionRange(input.value.length,input.value.length);else if(Number.isInteger(options.caretOffset))input.setSelectionRange(Math.max(0,Math.min(input.value.length,options.caretOffset)),Math.max(0,Math.min(input.value.length,options.caretOffset)));else if(options.selectAll===false)input.setSelectionRange(input.value.length,input.value.length);else input.select();let done=false,committedChange=false;const commit=()=>{if(done)return committedChange;done=true;const originalValue=(data?.hasFormula||data?.formula)?'='+data.formula:(data?.value??'');committedChange=input.value!==originalValue;if(committedChange)vscode.postMessage({type:'edit',row:r,column:c,value:input.value,sheetIndex:model.sheetIndex});return committedChange};input.addEventListener('input',()=>{input.style.height='auto';input.style.height=Math.min(120,Math.max(16,input.scrollHeight||16))+'px'});input.onkeydown=e=>{if(e.key==='Enter'&&e.altKey&&!e.ctrlKey&&!e.metaKey){if(e.isComposing||e.keyCode===229)return;e.preventDefault();const start=input.selectionStart,end=input.selectionEnd;input.setRangeText('\\n',start,end,'end');input.style.height='auto';input.style.height=Math.min(120,Math.max(16,input.scrollHeight||16))+'px'}else if(e.key==='Tab'){if(e.isComposing||e.keyCode===229)return;e.preventDefault();e.stopPropagation();if(tryTableTab(r,c,e.shiftKey,input.value)){done=true;committedChange=true;input.blur();return}commit();moveAfterCommit(r,c,e.shiftKey,'Tab');input.blur()}else if(e.key==='Enter'&&(e.ctrlKey||e.metaKey)){if(e.isComposing||e.keyCode===229)return;e.preventDefault();done=true;committedChange=false;vscode.postMessage({type:'fill-range',row:r,column:c,value:input.value,range:selectedRange,sheetIndex:model.sheetIndex});input.blur()}else if(e.key==='Enter'){if(e.isComposing||e.keyCode===229)return;e.preventDefault();commit();input.blur();moveAfterCommit(r,c,e.shiftKey,'Enter')}else if(e.key==='Escape'){e.preventDefault();done=true;renderedRange=null;pendingRange=null;input.blur();requestRange()}};input.onblur=()=>{cell.style.overflow=editBox.overflow||'';cell.style.zIndex=editBox.zIndex||'';cell.style.justifyContent=editBox.justifyContent||'';if(!commit()){renderedRange=null;pendingRange=null;requestRange()}}}
  function scrollToCell(cell,extend=false,focusCell=true){const renderedCell=document.querySelector('[data-row="'+cell.row+'"][data-column="'+cell.column+'"]');select(cell.row,cell.column,renderedCell?.__jinSheetData,extend,focusCell);const grid=document.getElementById('grid');const freezeRows=model?.freezePane?.rows||0,freezeColumns=model?.freezePane?.columns||0,displayRow=rowView?.indexOf(cell.row)??cell.row;if(cell.row>=freezeRows){const rowTop=30+rowMetrics.top(displayRow),rowBottom=rowTop+rowMetrics.height(displayRow),visibleTop=30+frozenRowsHeight();if(rowTop-grid.scrollTop<visibleTop)grid.scrollTop=Math.max(0,rowTop-visibleTop);else if(rowBottom-grid.scrollTop>grid.clientHeight)grid.scrollTop=rowBottom-grid.clientHeight}if(cell.column>=freezeColumns){let frozenWidth=0;for(let column=0;column<freezeColumns;column++)frozenWidth+=gridColumnWidth(columnWidths.get(column)??model.defaultColumnWidth??DEFAULT_COLUMN_WIDTH);const cellLeft=52+columnOffsets[cell.column],cellRight=cellLeft+gridColumnWidth(columnWidths.get(cell.column)??model.defaultColumnWidth??DEFAULT_COLUMN_WIDTH),visibleLeft=52+frozenWidth;if(cellLeft-grid.scrollLeft<visibleLeft)grid.scrollLeft=Math.max(0,columnOffsets[cell.column]-frozenWidth);else if(cellRight-grid.scrollLeft>grid.clientWidth)grid.scrollLeft=cellRight-grid.clientWidth}requestRange()}function refreshCutHighlights(){const bounds=pendingCutRange?{top:Math.min(pendingCutRange.startRow,pendingCutRange.endRow),bottom:Math.max(pendingCutRange.startRow,pendingCutRange.endRow),left:Math.min(pendingCutRange.startColumn,pendingCutRange.endColumn),right:Math.max(pendingCutRange.startColumn,pendingCutRange.endColumn)}:null;for(const cell of document.querySelectorAll(".cell")){const row=Number(cell.dataset.row),column=Number(cell.dataset.column);cell.classList.toggle("cut-source",Boolean(bounds&&row>=bounds.top&&row<=bounds.bottom&&column>=bounds.left&&column<=bounds.right))}}function refreshSearchHighlights(){const current=matches[matchIndex];for(const cell of document.querySelectorAll('.cell')){const row=Number(cell.dataset.row),column=Number(cell.dataset.column),isMatch=matchedCells.has(row+':'+column),isCurrent=isMatch&&current?.row===row&&current?.column===column&&(current.sheetIndex===undefined||current.sheetIndex===model?.sheetIndex);cell.classList.toggle('search-match',isMatch);cell.classList.toggle('search-current',isCurrent)}}function runSearch(){const search=document.getElementById('search'),query=search.value;if(!query){matches=[];matchedCells.clear();matchIndex=0;searchedQuery='';searchPagePending=false;searchRequest++;document.getElementById('searchCount').textContent='';document.getElementById('clearSearch').hidden=true;refreshSearchHighlights();return}if(query===searchedQuery&&matches.length){if(searchPagePending)return;if(matchIndex+1<matches.length){matchIndex++;focusSearchMatch();refreshSearchHighlights();return}const last=matches[matches.length-1];searchPagePending=true;const requestId=++searchRequest;vscode.postMessage({type:'search',query,requestId,sheetIndex:model?.sheetIndex||0,scope:searchScope.value,after:{sheetIndex:last.sheetIndex,row:last.row,column:last.column}});return}searchedQuery=query;matches=[];matchedCells.clear();matchIndex=0;searchPagePending=false;refreshSearchHighlights();document.getElementById('searchCount').textContent='';const requestId=++searchRequest;vscode.postMessage({type:'search',query,requestId,sheetIndex:model?.sheetIndex||0,scope:searchScope.value})}function dragEdgeSpeed(position,start,end){const margin=24;if(position<start+margin)return -Math.min(24,Math.max(1,(start+margin-position)*.35));if(position>end-margin)return Math.min(24,Math.max(1,(position-(end-margin))*.35));return 0}function rowAtDisplayPixel(pixel){if(!rowView?.displayCount)return 0;const index=rowMetrics.indexAtPixel(pixel);return rowView.rowAt(index)}function updateDragSelection(){if(!dragging||!dragPointer||!dragMoved||!model)return;const grid=document.getElementById('grid'),rect=grid.getBoundingClientRect();if(!rect.width||!rect.height)return;if(dragAxis==='row'){const delta=dragEdgeSpeed(dragPointer.y,rect.top,rect.bottom);if(delta)grid.scrollTop=Math.max(0,Math.min(grid.scrollHeight-grid.clientHeight,grid.scrollTop+delta));const freezeRows=model?.freezePane?.rows||0,y=Math.max(rect.top+30,Math.min(rect.bottom-1,dragPointer.y)),pixel=y<rect.top+30+frozenRowsHeight()?y-rect.top-30:y-rect.top-30+grid.scrollTop,row=rowAtDisplayPixel(pixel);selectAxis('row',row,true);return}if(dragAxis==='column'){const delta=dragEdgeSpeed(dragPointer.x,rect.left,rect.right);if(delta)grid.scrollLeft=Math.max(0,Math.min(grid.scrollWidth-grid.clientWidth,grid.scrollLeft+delta));const x=Math.max(rect.left+ROW_HEADER_WIDTH+1,Math.min(rect.right-1,dragPointer.x)),columnPixel=x-rect.left-ROW_HEADER_WIDTH+grid.scrollLeft,column=Math.max(0,Math.min(model.gridColumnCount-1,columnAtPixel(columnOffsets,columnPixel)));selectAxis('column',column,true);return}const deltaX=dragEdgeSpeed(dragPointer.x,rect.left,rect.right),deltaY=dragEdgeSpeed(dragPointer.y,rect.top,rect.bottom);if(deltaX)grid.scrollLeft=Math.max(0,Math.min(grid.scrollWidth-grid.clientWidth,grid.scrollLeft+deltaX));if(deltaY)grid.scrollTop=Math.max(0,Math.min(grid.scrollHeight-grid.clientHeight,grid.scrollTop+deltaY));const freezeRows=model?.freezePane?.rows||0,freezeColumns=model?.freezePane?.columns||0,x=Math.max(rect.left+ROW_HEADER_WIDTH+1,Math.min(rect.right-1,dragPointer.x)),y=Math.max(rect.top+30,Math.min(rect.bottom-1,dragPointer.y));const rowPixel=y<rect.top+30+frozenRowsHeight()?y-rect.top-30:y-rect.top-30+grid.scrollTop,row=rowAtDisplayPixel(rowPixel);let frozenWidth=0;for(let column=0;column<freezeColumns;column++)frozenWidth+=gridColumnWidth(columnWidths.get(column)??model.defaultColumnWidth??DEFAULT_COLUMN_WIDTH);const columnPixel=x<rect.left+ROW_HEADER_WIDTH+frozenWidth?x-rect.left-ROW_HEADER_WIDTH:x-rect.left-ROW_HEADER_WIDTH+grid.scrollLeft,column=Math.max(0,Math.min(model.gridColumnCount-1,columnAtPixel(columnOffsets,columnPixel))),target=document.querySelector('[data-row="'+row+'"][data-column="'+column+'"]');if(dragAxis==="fill"&&fillDrag){extendFillSelection(row,column)}else if(row!==selected.row||column!==selected.column){dragged=true;select(row,column,target?.__jinSheetData,true)}}function extendFillSelection(row,column){if(!fillDrag)return;const source=fillDrag.sourceRange,sourceTop=Math.min(source.startRow,source.endRow),sourceBottom=Math.max(source.startRow,source.endRow),sourceLeft=Math.min(source.startColumn,source.endColumn),sourceRight=Math.max(source.startColumn,source.endColumn);if(row>=sourceTop&&row<=sourceBottom&&column>=sourceLeft&&column<=sourceRight){selectionAnchor={row:source.startRow,column:source.startColumn};const original=document.querySelector('[data-row='+source.endRow+'][data-column='+source.endColumn+']');select(source.endRow,source.endColumn,original?.__jinSheetData,true);return;}const anchor={row:row<sourceTop?sourceBottom:sourceTop,column:column<sourceLeft?sourceRight:sourceLeft},active={row:row<sourceTop?row:Math.max(row,sourceBottom),column:column<sourceLeft?column:Math.max(column,sourceRight)};selectionAnchor=anchor;const target=document.querySelector('[data-row='+active.row+'][data-column='+active.column+']');select(active.row,active.column,target?.__jinSheetData,true)}function startCellDrag(){clearInterval(dragScrollTimer);dragScrollTimer=setInterval(updateDragSelection,40)}function stopCellDrag(){dragging=false;dragAxis=null;dragPointer=null;dragMoved=false;clearInterval(dragScrollTimer);dragScrollTimer=null}function moveVisibleRow(row,step,top=0,bottom=(model?.gridRowCount||model?.rowCount||1)-1,wrap=false){if(!rowView?.displayCount)return row;let low=0,high=rowView.displayCount;while(low<high){const middle=(low+high)>>1;if(rowView.rowAt(middle)<top)low=middle+1;else high=middle}const first=low;low=0;high=rowView.displayCount;while(low<high){const middle=(low+high)>>1;if(rowView.rowAt(middle)<=bottom)low=middle+1;else high=middle}const last=low-1;if(first>last)return row;let index=rowView.indexOf(row);if(index<first||index>last)return row;index+=step;if(index>last)index=wrap?first:last;if(index<first)index=wrap?last:first;return rowView.rowAt(index)}let tableTabSequence=0,pendingTableTab=null;
  function tryTableTab(row,column,backwards,value){if(backwards||model?.readOnly||model?.kind!=='xlsx'||!autoExpandTables.checked||autoExpandTables.disabled||selectedRange.startRow!==selectedRange.endRow||selectedRange.startColumn!==selectedRange.endColumn)return false;if(pendingTableTab)return row===pendingTableTab.row&&column===pendingTableTab.column&&(value===undefined||value===pendingTableTab.value);const table=(model.tables||[]).find(item=>row===item.range.bottom-(item.totalsRowCount||0)&&(row>=item.range.top+(item.headerRowCount||0)||(item.headerRowCount===1&&row===item.range.top))&&column===Array.from({length:item.range.right-item.range.left+1},(_,index)=>item.range.left+index).filter(index=>!(model.hiddenColumns||[]).includes(index)).at(-1));if(!table)return false;pendingTableTab={requestId:++tableTabSequence,sheetIndex:model.sheetIndex,tableName:table.name,row,column,...(value!==undefined?{value}:{})};vscode.postMessage({type:'table-tab',...pendingTableTab});return true}
  function moveAfterCommit(row,column,backwards=false,key='Tab'){const rowCount=model.gridRowCount||model.rowCount;const columnCount=model.gridColumnCount||model.columnCount;const top=Math.min(selectedRange.startRow,selectedRange.endRow),bottom=Math.max(selectedRange.startRow,selectedRange.endRow),left=Math.min(selectedRange.startColumn,selectedRange.endColumn),right=Math.max(selectedRange.startColumn,selectedRange.endColumn),withinSelection=(bottom>top||right>left)&&row>=top&&row<=bottom&&column>=left&&column<=right;let nextRow=row,nextColumn=column;if(withinSelection&&key==='Tab'){nextColumn+=backwards?-1:1;if(nextColumn>right){nextColumn=left;nextRow=moveVisibleRow(nextRow,1,top,bottom,true)}else if(nextColumn<left){nextColumn=right;nextRow=moveVisibleRow(nextRow,-1,top,bottom,true)}}else if(withinSelection&&key==='Enter'){nextRow=moveVisibleRow(nextRow,backwards?-1:1,top,bottom,true)}else if(key==='Tab'){nextColumn+=backwards?-1:1;if(nextColumn>=columnCount){nextRow=moveVisibleRow(nextRow,1);nextColumn=0}else if(nextColumn<0){nextRow=moveVisibleRow(nextRow,-1);nextColumn=columnCount-1}for(let count=0;count<columnCount&&(model.hiddenColumns||[]).includes(nextColumn);count++){nextColumn+=backwards?-1:1;if(nextColumn>=columnCount||nextColumn<0){nextRow=moveVisibleRow(nextRow,backwards?-1:1);nextColumn=backwards?columnCount-1:0}}}else nextRow=moveVisibleRow(row,backwards?-1:1);if(!withinSelection&&key==='Tab'){const table=(model.tables||[]).find(item=>row>=item.range.top+(item.headerRowCount||0)&&row<=item.range.bottom-(item.totalsRowCount||0)&&column>=item.range.left&&column<=item.range.right);if(table){const columns=Array.from({length:table.range.right-table.range.left+1},(_,index)=>table.range.left+index).filter(index=>!(model.hiddenColumns||[]).includes(index)),first=table.range.top+(table.headerRowCount||0),last=table.range.bottom-(table.totalsRowCount||0);if(!backwards&&column===columns.at(-1)&&row<last){const next=moveVisibleRow(row,1,first,last);if(next!==row){nextRow=next;nextColumn=columns[0]}}else if(backwards&&column===columns[0]&&row>first){const next=moveVisibleRow(row,-1,first,last);if(next!==row){nextRow=next;nextColumn=columns.at(-1)}}}}const currentMerge=model.mergedCells?.find(range=>row>=range.top&&row<=range.bottom&&column>=range.left&&column<=range.right);if(currentMerge){if(key==='Tab'&&!backwards&&nextColumn>column&&nextColumn<=currentMerge.right)nextColumn=currentMerge.right+1;if(key==='Enter'&&!backwards&&nextRow>row&&nextRow<=currentMerge.bottom)nextRow=moveVisibleRow(currentMerge.bottom+1,1)}const anchor={...selectionAnchor},range={...selectedRange};scrollToCell({row:nextRow,column:nextColumn},withinSelection);if(withinSelection){selectionAnchor=anchor;selectedRange=range;for(const item of document.querySelectorAll('.cell')){const itemRow=Number(item.dataset.row),itemColumn=Number(item.dataset.column),inRange=itemRow>=top&&itemRow<=bottom&&itemColumn>=left&&itemColumn<=right;item.classList.toggle('in-range',inRange);item.setAttribute('aria-selected',String(inRange))}refreshSelectionSummary()}}
  document.getElementById('grid').addEventListener('scroll',()=>{syncHeadings();clearTimeout(scrollTimer);scrollTimer=setTimeout(requestRange,60)});
  document.addEventListener('mousemove',e=>{if(dragging){dragPointer={x:e.clientX,y:e.clientY};dragMoved=true;updateDragSelection()}if(columnResize)resizeColumn(columnResize.column,columnResize.startWidth+(e.clientX-columnResize.startX)/gridZoom);if(rowResize){const delta=e.clientY-rowResize.startY;if(Math.abs(delta)>.2||rowResize.moved){if(Math.abs(delta)>.2)rowResize.moved=true;rowResize.height=previewRowHeight({startRow:rowResize.startRow,endRow:rowResize.endRow},rowResize.startHeight+delta/gridZoom)}}});window.addEventListener('blur',()=>{fillDrag=null;stopCellDrag();cancelRowResize()});document.addEventListener('mouseup',()=>{if(fillDrag){const sourceRange=fillDrag.sourceRange,targetRange={...selectedRange};const changed=targetRange.startRow!==sourceRange.startRow||targetRange.endRow!==sourceRange.endRow||targetRange.startColumn!==sourceRange.startColumn||targetRange.endColumn!==sourceRange.endColumn;fillDrag=null;if(changed)vscode.postMessage({type:"fill-handle",sourceRange,targetRange,sheetIndex:model?.sheetIndex||0})}stopCellDrag();syncFillHandle();if(rowResize){if(rowResize.moved){const current=rowResize;rowResize=null;document.body.style.cursor=current.cursor;commitRowHeight({startRow:current.startRow,endRow:current.endRow},current.height);requestRange()}else cancelRowResize()}if(columnResize){const {column,startWidth}=columnResize;const changed=Math.abs((columnWidths.get(column)??startWidth)-startWidth)>=1;columnResize=null;requestRange();if(changed)commitColumnWidth(column)}});
  const formulaInput=document.getElementById('formula');let formulaDirty=false;function commitFormulaEdit(){if(!formulaDirty||formulaInput.disabled)return;const value=formulaInput.value;formulaDirty=false;formulaInput.dataset.originalValue=value;vscode.postMessage({type:'edit',row:selected.row,column:selected.column,value,sheetIndex:model?.sheetIndex||0})}formulaInput.addEventListener('input',()=>{formulaDirty=formulaInput.value!==formulaInput.dataset.originalValue});formulaInput.addEventListener('blur',()=>commitFormulaEdit());formulaInput.addEventListener('keydown',e=>{if(e.key==='Escape'){e.preventDefault();formulaDirty=false;formulaInput.value=formulaInput.dataset.originalValue||'';formulaInput.title=formulaInput.value;return}if((e.key==='Enter'||e.key==='Tab')&&!e.target.disabled){if(e.isComposing||e.keyCode===229)return;e.preventDefault();const row=selected.row,column=selected.column;pendingCutRange=null;refreshCutHighlights();if(e.key==='Enter'&&(e.ctrlKey||e.metaKey)){formulaDirty=false;formulaInput.dataset.originalValue=e.target.value;vscode.postMessage({type:'fill-range',row,column,value:e.target.value,range:selectedRange,sheetIndex:model?.sheetIndex||0})}else{if(e.key==='Tab'&&tryTableTab(row,column,e.shiftKey,formulaDirty?formulaInput.value:undefined)){formulaDirty=false;return}commitFormulaEdit();if(e.key==='Tab')moveAfterCommit(row,column,e.shiftKey,'Tab');else moveAfterCommit(row,column,e.shiftKey,'Enter')}}});
  const addressInput=document.getElementById('address'),goToStatus=document.getElementById('goToStatus');function parseCellReference(reference){const match=/^\\$?([A-Za-z]{1,3})\\$?([1-9]\\d*)$/.exec(reference.trim());if(!match)return null;let column=0;for(const char of match[1].toUpperCase())column=column*26+char.charCodeAt(0)-64;const row=Number(match[2])-1;column--;if(!Number.isSafeInteger(row)||!Number.isSafeInteger(column)||row<0||column<0||row>=(model?.gridRowCount||model?.rowCount||0)||column>=(model?.gridColumnCount||model?.columnCount||0))return null;return{row,column}}function applyGoToAddress(){if(!model)return false;const parts=addressInput.value.trim().split(':');if(parts.length>2||parts.some(part=>!part.trim())){goToStatus.textContent=t('invalid-cell-reference');addressInput.setAttribute('aria-invalid','true');addressInput.focus();addressInput.select();return false}const start=parseCellReference(parts[0]),end=parts.length===2?parseCellReference(parts[1]):start;if(!start||!end){goToStatus.textContent=t('invalid-cell-reference');addressInput.setAttribute('aria-invalid','true');addressInput.focus();addressInput.select();return false}goToStatus.textContent='';addressInput.setAttribute('aria-invalid','false');commitFormulaEdit();scrollToCell(start,false,true);selectionAnchor={...selected};selectedRange={startRow:selected.row,endRow:end.row,startColumn:selected.column,endColumn:end.column};for(const cell of document.querySelectorAll('.cell')){const row=Number(cell.dataset.row),column=Number(cell.dataset.column),inRange=row>=Math.min(selectedRange.startRow,selectedRange.endRow)&&row<=Math.max(selectedRange.startRow,selectedRange.endRow)&&column>=Math.min(selectedRange.startColumn,selectedRange.endColumn)&&column<=Math.max(selectedRange.startColumn,selectedRange.endColumn);cell.classList.toggle('in-range',inRange);cell.setAttribute('aria-selected',String(inRange))}addressInput.value=label(selected.column+1)+(selected.row+1);document.getElementById('grid').focus({preventScroll:true});syncFillHandle();syncChartButton();syncTableResizeButton();refreshSelectionSummary();refreshSearchHighlights();requestRange();return true}addressInput.addEventListener('focus',()=>addressInput.select());addressInput.addEventListener('input',()=>{goToStatus.textContent='';addressInput.setAttribute('aria-invalid','false')});addressInput.addEventListener('keydown',event=>{if(event.key==='Enter'&&!event.isComposing){event.preventDefault();applyGoToAddress()}else if(event.key==='Escape'){event.preventDefault();goToStatus.textContent='';addressInput.setAttribute('aria-invalid','false');addressInput.value=label(selected.column+1)+(selected.row+1);addressInput.blur()}});const searchInput=document.getElementById('search'),searchScope=document.getElementById('searchScope'),clearSearchButton=document.getElementById('clearSearch'),replaceInput=document.getElementById('replaceValue'),replaceStatus=document.getElementById('replaceStatus'),replaceCurrentButton=document.getElementById('replaceCurrentButton'),replaceAllButton=document.getElementById('replaceAllButton');function syncSearchMatchesForSheet(){matchedCells=new Set(matches.filter(match=>match.sheetIndex===undefined||match.sheetIndex===model?.sheetIndex).map(match=>match.row+':'+match.column))}function focusSearchMatch(){const match=matches[matchIndex];if(!match)return;if(Number.isInteger(match.sheetIndex)&&match.sheetIndex!==model?.sheetIndex){vscode.postMessage({type:'search-navigate',scope:searchScope.value,requestId:searchRequest,sheetIndex:match.sheetIndex,row:match.row,column:match.column});return}scrollToCell(match,false,false)}function resetSearchResults(){matches=[];matchedCells.clear();matchIndex=0;searchedQuery='';searchPagePending=false;searchRequest++;document.getElementById('searchCount').textContent='';refreshSearchHighlights()}function clearSearch(){searchInput.value='';resetSearchResults();replaceStatus.textContent='';clearSearchButton.hidden=true;searchInput.focus({preventScroll:true})}searchInput.addEventListener('input',()=>{resetSearchResults();replaceStatus.textContent='';clearSearchButton.hidden=!searchInput.value});searchScope.addEventListener('change',()=>{resetSearchResults();replaceStatus.textContent=''});document.getElementById('findButton').addEventListener('click',()=>{searchInput.focus({preventScroll:true});runSearch()});clearSearchButton.addEventListener('click',clearSearch);
  function submitReplace(mode){const query=searchInput.value;if(!query||model?.readOnly||pendingReplaceRequest)return;const requestId=++replaceRequest,current=searchedQuery===query&&matches.length?matches[matchIndex]:null;pendingReplaceRequest=requestId;replaceCurrentButton.disabled=true;replaceAllButton.disabled=true;replaceStatus.textContent='';vscode.postMessage({type:mode,query,replacement:replaceInput.value,requestId,sheetIndex:model?.sheetIndex||0,scope:searchScope.value,...(mode==='replace-current'&&current?{row:current.row,column:current.column,targetSheetIndex:current.sheetIndex??model?.sheetIndex??0}:{})})}replaceCurrentButton.addEventListener('click',()=>submitReplace('replace-current'));replaceAllButton.addEventListener('click',()=>submitReplace('replace-all'));replaceInput.addEventListener('input',()=>{replaceStatus.textContent=''});replaceInput.addEventListener('keydown',e=>{if(e.key==='Enter'&&!e.isComposing){e.preventDefault();submitReplace('replace-current')}});searchInput.addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();runSearch()}if(e.key==='Escape'){e.preventDefault();clearSearch()}});
  document.addEventListener('keydown', e => {
    if (rowResize && e.key === 'Escape') { e.preventDefault(); cancelRowResize(); return; }
    if (model && ['xlsx', 'csv'].includes(model.kind) && (e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'g') {
      e.preventDefault(); addressInput.focus(); addressInput.select(); return;
    }
    if (e.target.id === 'search') return;
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'f') {
      e.preventDefault(); const search = document.getElementById('search'); search.focus(); search.select(); return;
    }
    if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA' || !model || !e.target.closest?.('#grid')) return;
    if (e.altKey && !e.ctrlKey && !e.metaKey && e.key === '=') {
      e.preventDefault();
      if (model.kind === 'xlsx' && !model.readOnly) vscode.postMessage({ type: 'auto-sum', row: selected.row, column: selected.column, sheetIndex: model.sheetIndex });
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.code === 'Semicolon' && !e.altKey && !model.readOnly) {
      e.preventDefault();
      vscode.postMessage({ type: 'insert-date-time', row: selected.row, column: selected.column, mode: e.shiftKey ? 'time' : 'date', sheetIndex: model.sheetIndex });
      return;
    }
    if (e.key === ' ' && e.shiftKey && (e.ctrlKey || e.metaKey)) {
      e.preventDefault();
      const requestId = ++currentRegionRequest;
      vscode.postMessage({ type: 'current-region', ...selected, includeSheet: true, requestId, sheetIndex: model.sheetIndex });
      return;
    }
    if (e.key === ' ' && e.shiftKey) {
      e.preventDefault(); selectAxis('row', selected.row); return;
    }
    if (e.key === ' ' && (e.ctrlKey || e.metaKey)) {
      e.preventDefault(); selectAxis('column', selected.column); return;
    }
    if (e.key === 'Escape' && pendingCutRange) { e.preventDefault(); pendingCutRange = null; refreshCutHighlights(); return; }
    if (e.key === 'Delete' || e.key === 'Backspace') {
      if (model.readOnly) return;
      if (e.key === 'Backspace' && selectedRange.startRow === selectedRange.endRow && selectedRange.startColumn === selectedRange.endColumn) {
        e.preventDefault(); e.stopPropagation(); const cell = document.querySelector('.cell.selected'); edit(cell, selected.row, selected.column, cell?.__jinSheetData, '', { selectAll: false }); return;
      }
      e.preventDefault(); pendingCutRange = null; refreshCutHighlights(); vscode.postMessage({ type: 'clear', ...selectionAnchor, range: selectedRange, sheetIndex: model.sheetIndex }); return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c') {
      e.preventDefault(); pendingCutRange = null; refreshCutHighlights(); vscode.postMessage({ type: 'copy', ...selected, range: selectedRange, sheetIndex: model.sheetIndex, ...(model.visibleRows ? { visibleRows: model.visibleRows } : model.hiddenRows ? { hiddenRows: model.hiddenRows } : {}), ...(model.hiddenColumns?.length ? { hiddenColumns: model.hiddenColumns } : {}) }); return;
    }
    if ((e.ctrlKey || e.metaKey) && ['d', 'r'].includes(e.key.toLowerCase())) {
      if (model.readOnly) return;
      e.preventDefault(); pendingCutRange = null; refreshCutHighlights(); vscode.postMessage({ type: 'fill-selection', direction: e.key.toLowerCase() === 'd' ? 'down' : 'right', range: selectedRange, sheetIndex: model.sheetIndex }); return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'x') {
      if (model.readOnly) return;
      e.preventDefault(); pendingCutRange = { ...selectedRange }; refreshCutHighlights(); vscode.postMessage({ type: 'cut', ...selected, range: selectedRange, sheetIndex: model.sheetIndex }); return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'v') {
      e.preventDefault(); if (pendingCutRange) { vscode.postMessage({ type: 'move-paste', ...selectionAnchor, sourceRange: pendingCutRange, sheetIndex: model.sheetIndex }); pendingCutRange = null; refreshCutHighlights(); } else vscode.postMessage({ type: 'paste', ...selectionAnchor, range: selectedRange, sheetIndex: model.sheetIndex }); return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'a') {
      e.preventDefault();
      const requestId = ++currentRegionRequest;
      vscode.postMessage({ type: 'current-region', ...selected, includeSheet: currentRegionStage > 0, requestId, sheetIndex: model.sheetIndex });
      return;
    }
    if ((e.ctrlKey || e.metaKey) && ['PageUp', 'PageDown'].includes(e.key)) {
      e.preventDefault();
      const step = e.key === 'PageDown' ? 1 : -1;
      const candidates = (model.sheets || []).filter(sheet => sheet.visible && (step > 0 ? sheet.index > model.sheetIndex : sheet.index < model.sheetIndex));
      candidates.sort((left, right) => step * (left.index - right.index));
      if (candidates.length) vscode.postMessage({ type: 'refresh', sheetIndex: candidates[0].index });
      return;
    }
    if (e.key === 'PageUp' || e.key === 'PageDown') {
      e.preventDefault();
      const pagePixels = Math.max(1, document.getElementById('grid').clientHeight - 30 - frozenRowsHeight());
      const direction = e.key === 'PageDown' ? 1 : -1;
      const displayIndex = rowView.indexOf(selected.row);
      const pixel = displayIndex < 0 ? 0 : Math.max(0, rowMetrics.top(displayIndex) + direction * pagePixels);
      let nextIndex = rowMetrics.indexAtPixel(pixel);
      if (direction < 0 && rowMetrics.top(nextIndex) < pixel) nextIndex = Math.min(rowView.displayCount - 1, nextIndex + 1);
      const row = displayIndex < 0 ? selected.row : rowView.rowAt(nextIndex);
      scrollToCell({ row, column: selected.column }, e.shiftKey);
      return;
    }
    if (e.key === 'Tab') {
      e.preventDefault();
      if(!tryTableTab(selected.row,selected.column,e.shiftKey))moveAfterCommit(selected.row, selected.column, e.shiftKey, 'Tab');
      return;
    }
    if (e.key === 'Enter') {
      e.preventDefault();
      moveAfterCommit(selected.row, selected.column, e.shiftKey, 'Enter');
      return;
    }
    if ((e.ctrlKey || e.metaKey) && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) {
      e.preventDefault();
      const direction = { ArrowUp: 'up', ArrowDown: 'down', ArrowLeft: 'left', ArrowRight: 'right' }[e.key];
      const requestId = ++navigationRequest;
      vscode.postMessage({ type: 'navigate', ...selected, direction, extend: e.shiftKey, requestId, sheetIndex: model.sheetIndex, ...(model.visibleRows ? { visibleRows: model.visibleRows } : model.hiddenRows ? { hiddenRows: model.hiddenRows } : {}) }); return;
    }
    let row = selected.row, column = selected.column;
    const usedRows = Math.max(1, model.rowCount || 0), usedColumns = Math.max(1, model.columnCount || 0);
    if (e.key === 'Home') { row = e.ctrlKey || e.metaKey ? 0 : row; column = 0; }
    else if (e.key === 'End') {
      if (e.ctrlKey || e.metaKey) row = usedRows - 1;
      column = usedColumns - 1;
    } else if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) {
      if (e.key === 'ArrowUp') row = rowView.rowAt(Math.max(0,rowView.indexOf(row)-1));
      if (e.key === 'ArrowDown') row = rowView.rowAt(Math.min(rowView.displayCount-1,rowView.indexOf(row)+1));
      if (e.key === 'ArrowLeft') { let candidate=selected.column-1; while(candidate>=0&&(model.hiddenColumns||[]).includes(candidate))candidate--; if(candidate>=0)column=candidate; }
      if (e.key === 'ArrowRight') { let candidate=selected.column+1; while(candidate<(model.gridColumnCount||model.columnCount)&&(model.hiddenColumns||[]).includes(candidate))candidate++; if(candidate<model.gridColumnCount)column=candidate; }
    } else return;
    const currentMerge=model.mergedCells?.find(range=>selected.row>=range.top&&selected.row<=range.bottom&&selected.column>=range.left&&selected.column<=range.right);if(currentMerge){if(e.key==='ArrowRight'&&column>selected.column&&column<=currentMerge.right)column=currentMerge.right+1<(model.gridColumnCount||model.columnCount)?currentMerge.right+1:selected.column;if(e.key==='ArrowDown'&&row>selected.row&&row<=currentMerge.bottom){const nextIndex=rowView.indexOf(currentMerge.bottom)+1;row=nextIndex<rowView.displayCount?rowView.rowAt(nextIndex):selected.row}}else{const targetMerge=model.mergedCells?.find(range=>row>=range.top&&row<=range.bottom&&column>=range.left&&column<=range.right);if(targetMerge){row=targetMerge.top;column=targetMerge.left}}
    e.preventDefault(); scrollToCell({ row, column }, e.shiftKey && e.key !== 'Tab');
  });
  const sheetZoomInput=document.getElementById('sheetZoom');const sendSheetZoom=scale=>{const value=Number(scale);if(!model||model.kind!=='xlsx'||model.readOnly)return;if(!Number.isInteger(value)||value<10||value>400){sheetZoomInput.value=String(model.zoomScale||100);return}sheetZoomInput.value=String(value);vscode.postMessage({type:'sheet-zoom',sheetIndex:model.sheetIndex,zoomScale:value})};document.getElementById('zoomOut').onclick=()=>sendSheetZoom(Number(sheetZoomInput.value)-10);document.getElementById('zoomIn').onclick=()=>sendSheetZoom(Number(sheetZoomInput.value)+10);sheetZoomInput.addEventListener('change',()=>sendSheetZoom(sheetZoomInput.value));const sheetTabs=document.getElementById('sheets');let draggedSheetIndex=null;sheetTabs.addEventListener('pointerdown',event=>{const tab=event.target.closest('[role="tab"]');if(tab)tab.draggable=Boolean(model?.kind==='xlsx'&&!model.readOnly)});sheetTabs.addEventListener('dragstart',event=>{const tab=event.target.closest('[role="tab"]');if(!tab||!model||model.readOnly)return;draggedSheetIndex=[...sheetTabs.querySelectorAll('[role="tab"]')].indexOf(tab);event.dataTransfer.effectAllowed='move';event.dataTransfer.setData('text/plain',String(draggedSheetIndex))});sheetTabs.addEventListener('dragover',event=>{if(event.target.closest('[role="tab"]')){event.preventDefault();event.dataTransfer.dropEffect='move'}});sheetTabs.addEventListener('drop',event=>{const tab=event.target.closest('[role="tab"]'),tabs=[...sheetTabs.querySelectorAll('[role="tab"]')],targetIndex=tabs.indexOf(tab);if(draggedSheetIndex===null||targetIndex<0||targetIndex===draggedSheetIndex)return;event.preventDefault();vscode.postMessage({type:'move-sheet',sheetIndex:draggedSheetIndex,targetIndex,activeSheetIndex:model?.sheetIndex});draggedSheetIndex=null});sheetTabs.addEventListener('dragend',()=>{draggedSheetIndex=null});
  let printSettingsSignature='';
  const printDefaults=${JSON.stringify(DEFAULT_PRINT_SETTINGS)};
  const printInputs=['printPaper','printOrientation','printPageOrder','printScale','printFirstPage','printUseFirstPage','printFit','printFitWidth','printFitHeight','printGridlines','printHeadings',...['left','right','top','bottom','header','footer'].map(side=>'printMargin-'+side)];
  function syncPrintControls(){document.getElementById('printIncludeCharts').disabled=model?.kind!=='xlsx';syncPageBreakControls();syncPrintAreaControls();syncPrintTitleControls();syncHeaderFooterControls();const spreadsheet=['csv','xlsx'].includes(model?.kind);document.getElementById('printPanel').hidden=!spreadsheet;for(const id of ['printSheet','printSelection'])document.getElementById(id).disabled=!spreadsheet;for(const id of [...printInputs,'applyPageSettings'])document.getElementById(id).disabled=model?.kind!=='xlsx'||model.readOnly;const settings=model?.printSettings||printDefaults,signature=JSON.stringify([model?.kind,model?.sheetIndex,settings]);if(signature===printSettingsSignature){syncPrintFitControls();return}printSettingsSignature=signature;document.getElementById('printPaper').value=String(settings.paperSize);document.getElementById('printOrientation').value=settings.orientation;document.getElementById('printPageOrder').value=settings.pageOrder||'downThenOver';document.getElementById('printScale').value=String(settings.scale);document.getElementById('printFirstPage').value=String(settings.firstPageNumber??1);document.getElementById('printUseFirstPage').checked=Boolean(settings.useFirstPageNumber);document.getElementById('printFit').checked=settings.fitToPage;document.getElementById('printFitWidth').value=String(settings.fitToWidth);document.getElementById('printFitHeight').value=String(settings.fitToHeight);document.getElementById('printGridlines').checked=settings.gridLines;document.getElementById('printHeadings').checked=settings.headings;for(const side of ['left','right','top','bottom','header','footer'])document.getElementById('printMargin-'+side).value=String(settings.margins[side]);syncPrintFitControls()}
  function syncPrintFitControls(){const disabled=model?.kind!=='xlsx'||model.readOnly,fit=document.getElementById('printFit').checked;document.getElementById('printScale').disabled=disabled||fit;document.getElementById('printFirstPage').disabled=disabled||!document.getElementById('printUseFirstPage').checked;document.getElementById('printFitWidth').disabled=disabled||!fit;document.getElementById('printFitHeight').disabled=disabled||!fit}
  function syncPrintAreaControls(){const areas=model?.printAreas||[],description=areas.length?areas.map(area=>area.top===0&&area.bottom===XLSX_MAX_ROWS-1?label(area.left+1)+':'+label(area.right+1):area.left===0&&area.right===XLSX_MAX_COLUMNS-1?(area.top+1)+':'+(area.bottom+1):label(area.left+1)+(area.top+1)+':'+label(area.right+1)+(area.bottom+1)).join(', '):t('print-area-none'),disabled=model?.kind!=='xlsx'||model.readOnly;document.getElementById('printAreaStatus').textContent=model?.printAreaUnsupported?t('error-print-area-unsupported'):t('print-area')+': '+description;document.getElementById('printSheet').title=model?.printAreaUnsupported?t('error-print-area-unsupported'):t('print-area')+': '+description;document.getElementById('setPrintArea').disabled=disabled;document.getElementById('clearPrintArea').disabled=disabled||!areas.length}
  for(const [id,action] of [['setPrintArea','set'],['clearPrintArea','clear']])document.getElementById(id).onclick=()=>{if(model?.kind!=='xlsx'||model.readOnly)return;commitFormulaEdit();vscode.postMessage({type:'print-area',sheetIndex:model.sheetIndex||0,action,range:action==='set'?{...selectedRange}:null})};
  let printTitleSignature='';
  function syncPrintTitleControls(){const titles=model?.printTitles||{rows:null,columns:null},signature=JSON.stringify([model?.kind,model?.sheetIndex,titles]),disabled=model?.kind!=='xlsx'||model.readOnly;for(const id of ['printTitleRows','printTitleColumns','applyPrintTitles','clearPrintTitles'])document.getElementById(id).disabled=disabled;document.getElementById('printTitleStatus').textContent=model?.printTitlesUnsupported?t('error-print-titles-unsupported'):'';if(signature===printTitleSignature)return;printTitleSignature=signature;document.getElementById('printTitleRows').value=titles.rows?(titles.rows.start+1)+':'+(titles.rows.end+1):'';document.getElementById('printTitleColumns').value=titles.columns?label(titles.columns.start+1)+':'+label(titles.columns.end+1):''}
  for(const [id,action] of [['applyPrintTitles','set'],['clearPrintTitles','clear']])document.getElementById(id).onclick=()=>{if(model?.kind!=='xlsx'||model.readOnly)return;commitFormulaEdit();const rows=document.getElementById('printTitleRows'),columns=document.getElementById('printTitleColumns');if(action==='clear'){rows.value='';columns.value=''}vscode.postMessage({type:'print-titles',sheetIndex:model.sheetIndex||0,action,rows:rows.value,columns:columns.value})};
  const hfDefaults=${JSON.stringify(defaultHeaderFooter())};
  let hfDraft=JSON.parse(JSON.stringify(hfDefaults)),hfSignature='',hfVariant='odd',hfFocused='hf-header-center';
  const hfFields=['header','footer'].flatMap(part=>['left','center','right'].map(side=>({id:'hf-'+part+'-'+side,part,side})));
  function showHfVariant(){document.getElementById('hfVariant').value=hfVariant;for(const field of hfFields)document.getElementById(field.id).value=hfDraft[hfVariant][field.part][field.side];document.getElementById('hfDifferentFirst').checked=hfDraft.differentFirst;document.getElementById('hfDifferentEven').checked=hfDraft.differentOddEven;document.getElementById('hfScale').checked=hfDraft.scaleWithDoc;document.getElementById('hfAlign').checked=hfDraft.alignWithMargins}
  function readHfVariant(){for(const field of hfFields)hfDraft[hfVariant][field.part][field.side]=document.getElementById(field.id).value;hfDraft.differentFirst=document.getElementById('hfDifferentFirst').checked;hfDraft.differentOddEven=document.getElementById('hfDifferentEven').checked;hfDraft.scaleWithDoc=document.getElementById('hfScale').checked;hfDraft.alignWithMargins=document.getElementById('hfAlign').checked}
  function syncHeaderFooterControls(){const settings=model?.headerFooter||hfDefaults,unsupported=Boolean(model?.headerFooterUnsupported),disabled=model?.kind!=='xlsx'||model.readOnly;for(const id of [...hfFields.map(field=>field.id),'hfVariant','hfDifferentFirst','hfDifferentEven','hfScale','hfAlign','hfToken','hfInsert','hfApply'])document.getElementById(id).disabled=disabled||unsupported;document.getElementById('hfClear').disabled=disabled;document.getElementById('hfStatus').textContent=unsupported?t('error-hf-unsupported'):'';const signature=JSON.stringify([model?.kind,model?.sheetIndex,settings,unsupported]);if(signature===hfSignature)return;hfSignature=signature;hfDraft=JSON.parse(JSON.stringify(settings));hfVariant='odd';showHfVariant();document.getElementById('hfToken').value='&P'}
  for(const field of hfFields){const input=document.getElementById(field.id);input.onfocus=()=>{hfFocused=field.id};input.oninput=readHfVariant}
  for(const id of ['hfDifferentFirst','hfDifferentEven','hfScale','hfAlign'])document.getElementById(id).onchange=readHfVariant;
  document.getElementById('hfVariant').onchange=()=>{readHfVariant();hfVariant=document.getElementById('hfVariant').value;showHfVariant()};
  document.getElementById('hfInsert').onclick=()=>{const input=document.getElementById(hfFocused);if(input.disabled)return;const value=document.getElementById('hfToken').value,start=input.selectionStart??input.value.length,end=input.selectionEnd??start,next=input.value.slice(0,start)+value+input.value.slice(end);if(next.length>255){document.getElementById('hfStatus').textContent=t('error-hf-length');return}input.value=next;input.focus();input.setSelectionRange?.(start+value.length,start+value.length);readHfVariant()};
  for(const [id,action] of [['hfApply','set'],['hfClear','clear']])document.getElementById(id).onclick=()=>{if(model?.kind!=='xlsx'||model.readOnly||action==='set'&&model.headerFooterUnsupported)return;commitFormulaEdit();readHfVariant();if(action==='clear'){hfDraft=JSON.parse(JSON.stringify(hfDefaults));hfVariant='odd';showHfVariant()}vscode.postMessage({type:'header-footer',sheetIndex:model.sheetIndex||0,action,settings:JSON.parse(JSON.stringify(hfDraft))})};
  function syncPageBreakControls(){const breaks=model?.pageBreaks||{rows:[],columns:[]},unsupported=Boolean(model?.pageBreaksUnsupported),disabled=model?.kind!=='xlsx'||model.readOnly;document.getElementById('breakAddRow').disabled=disabled||unsupported||selected.row===0||breaks.rows.includes(selected.row);document.getElementById('breakRemoveRow').disabled=disabled||unsupported||!breaks.rows.includes(selected.row);document.getElementById('breakAddColumn').disabled=disabled||unsupported||selected.column===0||breaks.columns.includes(selected.column);document.getElementById('breakRemoveColumn').disabled=disabled||unsupported||!breaks.columns.includes(selected.column);document.getElementById('breakReset').disabled=disabled||!(breaks.hasAny||breaks.rows.length||breaks.columns.length||unsupported);const list=(items,format)=>items.length?items.slice(0,10).map(format).join(', ')+(items.length>10?' … (+'+(items.length-10)+')':''):t('break-none');let status=unsupported?t('error-page-break-unsupported'):t('break-rows')+': '+list(breaks.rows,value=>value+1)+' · '+t('break-columns')+': '+list(breaks.columns,value=>label(value+1));if(!unsupported&&breaks.hasAny&&!breaks.rows.length&&!breaks.columns.length)status+=' · '+t('break-other-info');if(model?.printSettings?.fitToPage&&(breaks.rows.length||breaks.columns.length))status+=' · '+t('break-fit-note');const target=document.getElementById('pageBreakStatus');if(target.textContent!==status)target.textContent=status}
  for(const [id,axis,action] of [['breakAddRow','row','add'],['breakRemoveRow','row','remove'],['breakAddColumn','column','add'],['breakRemoveColumn','column','remove']])document.getElementById(id).onclick=()=>{if(model?.kind!=='xlsx'||model.readOnly)return;commitFormulaEdit();vscode.postMessage({type:'page-break',sheetIndex:model.sheetIndex||0,axis,action,index:axis==='row'?selected.row:selected.column})};
  document.getElementById('breakReset').onclick=()=>{if(model?.kind!=='xlsx'||model.readOnly)return;commitFormulaEdit();vscode.postMessage({type:'page-break',sheetIndex:model.sheetIndex||0,action:'reset'})};
  document.getElementById('printFit').onchange=syncPrintFitControls;
  document.getElementById('printUseFirstPage').onchange=syncPrintFitControls;
  document.getElementById('applyPageSettings').onclick=()=>{if(model?.kind!=='xlsx'||model.readOnly)return;const settings={orientation:document.getElementById('printOrientation').value,pageOrder:document.getElementById('printPageOrder').value,paperSize:Number(document.getElementById('printPaper').value),scale:Number(document.getElementById('printScale').value),firstPageNumber:Number(document.getElementById('printFirstPage').value),useFirstPageNumber:document.getElementById('printUseFirstPage').checked,fitToPage:document.getElementById('printFit').checked,fitToWidth:Number(document.getElementById('printFitWidth').value),fitToHeight:Number(document.getElementById('printFitHeight').value),gridLines:document.getElementById('printGridlines').checked,headings:document.getElementById('printHeadings').checked,margins:Object.fromEntries(['left','right','top','bottom','header','footer'].map(side=>[side,Number(document.getElementById('printMargin-'+side).value)]))};vscode.postMessage({type:'page-settings',sheetIndex:model.sheetIndex||0,settings})};
  for(const [id,selection] of [['printSheet',false],['printSelection',true]])document.getElementById(id).onclick=()=>{if(!['csv','xlsx'].includes(model?.kind))return;commitFormulaEdit();const range=selection?{...selectedRange}:null;vscode.postMessage({type:'export-print',sheetIndex:model.sheetIndex||0,range,includeCharts:document.getElementById('printIncludeCharts').checked})};
  let dynamicFilterTimer=null,dynamicFilterDay=null;
  function scheduleDynamicFilterRefresh(reset=false){clearTimeout(dynamicFilterTimer);dynamicFilterTimer=null;if(!model?.hasDynamicDateFilters){dynamicFilterDay=null;return}const now=new Date(),day=now.getFullYear()+'-'+now.getMonth()+'-'+now.getDate();if(reset||dynamicFilterDay===null)dynamicFilterDay=day;if(dynamicFilterDay!==day){const active=document.activeElement,editingControl=active&&['INPUT','TEXTAREA','SELECT'].includes(active.tagName)&&active.id!=='search';if(document.hidden||editingControl||document.querySelector('.cell-editor')||formulaDirty||!filterPanel.hidden||rowResize||columnResize||dragging||pendingTableTotals||pendingCalculatedColumn||pendingTableTab){dynamicFilterTimer=setTimeout(()=>scheduleDynamicFilterRefresh(),60000);return}dynamicFilterDay=day;vscode.postMessage({type:'refresh-dynamic-filters',sheetIndex:model.sheetIndex})}const next=new Date(now.getFullYear(),now.getMonth(),now.getDate()+1);dynamicFilterTimer=setTimeout(()=>scheduleDynamicFilterRefresh(),Math.max(100,next.getTime()-now.getTime()+100))}
  window.addEventListener('focus',()=>scheduleDynamicFilterRefresh());document.addEventListener('visibilitychange',()=>{if(!document.hidden)scheduleDynamicFilterRefresh()});
  function updateNormalFontMetrics(next){if(next.kind!=='xlsx')return;const font=next.normalFont||{},family=typeof font.family==='string'&&font.family.trim()&&font.family.length<=100?font.family:'Calibri',fontSize=Number.isFinite(font.size)&&font.size>=1&&font.size<=409?font.size:11;let context;try{context=document.createElement('canvas').getContext?.('2d')}catch{return}if(!context)return;context.font=(font.italic?'italic ':'')+(font.bold?'bold ':'')+fontSize+'pt '+JSON.stringify(family)+', sans-serif';const measured=Math.max(...Array.from({length:10},(_,digit)=>context.measureText(String(digit)).width));if(!Number.isFinite(measured))return;const isExcelDefaultFont=family.toLocaleLowerCase()==='calibri'&&fontSize===11&&!font.bold&&!font.italic;const maximumDigitWidth=isExcelDefaultFont?7:Math.max(1,Math.min(600,Math.round(measured)));if(next.maximumDigitWidth===maximumDigitWidth)return;next.maximumDigitWidth=maximumDigitWidth;if(Number.isFinite(next.defaultColumnWidthCharacters))next.defaultColumnWidth=excelColumnWidthPixels(next.defaultColumnWidthCharacters,64,maximumDigitWidth);columnWidthsBySheet.delete(next.sheetIndex);vscode.postMessage({type:'normal-font-width',maximumDigitWidth,sheetIndex:next.sheetIndex})}
  function acceptModel(next){const activeZoom=Number(next.zoomScale)||100;gridZoom=Math.max(.1,Math.min(4,activeZoom/100));const zoomControl=document.getElementById('sheetZoom');if(zoomControl){zoomControl.value=String(activeZoom);zoomControl.disabled=next.kind!=='xlsx'||Boolean(next.readOnly)}for(const id of ['zoomOut','zoomIn']){const button=document.getElementById(id);if(button)button.disabled=next.kind!=='xlsx'||Boolean(next.readOnly)}const sheetChanged=Boolean(model&&model.sheetIndex!==next.sheetIndex);if(model&&(model.sheetIndex!==next.sheetIndex||model.kind!==next.kind)){pendingReplaceRequest=0;pendingTableTotals=null;pendingCalculatedColumn=null;pendingTableTab=null;closeFilterPanel();}if(rowBoundaryFocus){const active=document.activeElement;const keep=rowBoundaryFocus.sheetIndex===next.sheetIndex&&(active===rowBoundaryFocus.source||active===document.body&&!rowBoundaryFocus.source.isConnected||active?.classList?.contains('row-resize-handle')&&Number(active.dataset.row)===rowBoundaryFocus.row);if(keep)rowBoundaryFocus.restore=true;else rowBoundaryFocus=null}if(rowResize)document.body.style.cursor=rowResize.cursor;rowResize=null;if(model&&model.sheetIndex===next.sheetIndex){if(JSON.stringify(model.columnWidths||{})!==JSON.stringify(next.columnWidths||{}))columnWidthsBySheet.delete(next.sheetIndex)}else if(model){currentRegionRequest++;currentRegionStage=0;pendingCutRange=null;selectionsBySheet.set(model.sheetIndex,{selected:{...selected},selectionAnchor:{...selectionAnchor},selectedRange:{...selectedRange}});const saved=selectionsBySheet.get(next.sheetIndex);if(saved){selected={...saved.selected};selectionAnchor={...saved.selectionAnchor};selectedRange={...saved.selectedRange}}else{selected={row:0,column:0};selectionAnchor={row:0,column:0};selectedRange={startRow:0,endRow:0,startColumn:0,endColumn:0}}}model=next;updateNormalFontMetrics(model);draw();if(!searchScope.value)searchScope.value='sheet';searchScope.disabled=next.kind!=='xlsx';if(next.kind!=='xlsx'&&searchScope.value==='workbook'){searchScope.value='sheet';resetSearchResults()}if(sheetChanged&&searchScope.value==='sheet')resetSearchResults();syncSearchMatchesForSheet();refreshSearchHighlights();document.getElementById("replaceCurrentButton").disabled=Boolean(next.readOnly||pendingReplaceRequest);document.getElementById("replaceAllButton").disabled=Boolean(next.readOnly||pendingReplaceRequest);scheduleDynamicFilterRefresh(true);const chartPreview=document.getElementById("chartPreview");if(chartPreview){chartPreview.innerHTML=next.chartPreviewHtml?next.chartPreviewHtml:"";chartPreview.hidden=!next.chartPreviewHtml;syncChartButton()}}
  const filterPanel=document.getElementById('filterPanel'),filterList=document.getElementById('filterValuesList'),filterSelectAll=document.getElementById('filterSelectAll'),filterSearch=document.getElementById('filterSearch'),filterShowMore=document.getElementById('filterShowMore');
  const FILTER_OPTION_PAGE_SIZE=200;let dateMixedValues=[],activeFilterColumn=null,filterOptions=[],filterPage=0,filterRequestCounter=0,pendingFilterRequest=null;
  const filterMode=document.getElementById('filterMode'),filterUseSecond=document.getElementById('filterUseSecond'),filterLogic=document.getElementById('filterLogic');
  let filterColorPalette={fills:[],fonts:[]};
  function syncColorPreview(){const value=document.getElementById('filterColorValue').value,custom=document.getElementById('filterColorCustom');document.getElementById('filterColorCustomLabel').hidden=value!=='custom';const color=value==='none'?null:value==='custom'?custom.value:value,preview=document.getElementById('filterColorPreview'),fill=document.getElementById('filterColorTarget').value==='fill';preview.style.backgroundColor=fill?(color||'transparent'):'transparent';preview.style.color=fill?'inherit':(color||'inherit')}
  function renderColorChoices(preferred){const fill=document.getElementById('filterColorTarget').value==='fill',list=fill?filterColorPalette.fills:filterColorPalette.fonts,input=document.getElementById('filterColorValue');input.replaceChildren();for(const item of list||[]){const option=document.createElement('option');option.value=item.color||'none';option.textContent=(item.color||t(fill?'filter-color-no-fill':'filter-color-auto'))+' · '+item.count;input.append(option)}if(!input.children.length){const option=document.createElement('option');option.value='none';option.textContent=t(fill?'filter-color-no-fill':'filter-color-auto');input.append(option)}const custom=document.createElement('option');custom.value='custom';custom.textContent=t('filter-color-custom');input.append(custom);if(preferred!==undefined&&!Array.from(input.children).some(option=>option.value===(preferred||'none'))){document.getElementById('filterColorCustom').value=preferred||'#FF0000';input.value='custom'}else input.value=preferred===undefined?input.children[0].value:(preferred||'none');document.getElementById('filterColorLimit').hidden=(fill?filterColorPalette.fillCount:filterColorPalette.fontCount)<=256||!(fill?filterColorPalette.fillCount:filterColorPalette.fontCount);syncColorPreview()}
  document.getElementById('filterColorTarget').onchange=()=>renderColorChoices();document.getElementById('filterColorValue').onchange=syncColorPreview;document.getElementById('filterColorCustom').oninput=syncColorPreview;
  function syncFilterMode(){const conditions=filterMode.value==='conditions';document.getElementById('filterValueControls').hidden=filterMode.value!=='values';document.getElementById('filterConditionControls').hidden=!conditions;document.getElementById('filterTopControls').hidden=filterMode.value!=='top10';document.getElementById('filterDateControls').hidden=filterMode.value!=='dates';document.getElementById('filterDynamicControls').hidden=filterMode.value!=='dynamic';document.getElementById('filterColorControls').hidden=filterMode.value!=='color';const iconSelect=document.getElementById('filterIconValue'),iconOption=filterMode.options?.[filterMode.options.length-1];if(iconOption)iconOption.disabled=!iconSelect.options.length;document.getElementById('filterIconControls').hidden=filterMode.value!=='icons';const percent=document.getElementById('filterTopUnit').value==='percent',input=document.getElementById('filterTopValue');input.max=percent?'100':'500';input.min=percent?'0':'1';input.step=percent?'any':'1';document.getElementById('filterSecondCondition').disabled=!filterUseSecond.checked;filterLogic.disabled=!filterUseSecond.checked}
  filterMode.onchange=syncFilterMode;filterUseSecond.onchange=syncFilterMode;document.getElementById('filterTopUnit').onchange=syncFilterMode;
  function filterUiCondition(condition){let operator=condition.operator,value=condition.value,escapes=0;for(let index=value.length-2;index>=0&&value[index]==='~';index--)escapes++;const trailing=value.endsWith('*')&&escapes%2===0;if(value.length>1&&trailing&&value.startsWith('*')&&['equal','notEqual'].includes(operator)){operator=operator==='equal'?'contains':'notContains';value=value.slice(1,-1)}else if(operator==='equal'&&value.length>1&&trailing){operator='beginsWith';value=value.slice(0,-1)}else if(operator==='equal'&&value.length>1&&value.startsWith('*')){operator='endsWith';value=value.slice(1)}return{operator,value}}
  function encodedFilterCondition(index){let operator=document.getElementById('filterOperator'+index).value,value=document.getElementById('filterConditionValue'+index).value;if(['contains','notContains','beginsWith','endsWith'].includes(operator)){value=(['contains','notContains','endsWith'].includes(operator)?'*':'')+value+(['contains','notContains','beginsWith'].includes(operator)?'*':'');operator=operator==='notContains'?'notEqual':'equal'}return{operator,value}}
  function submitFilter(clear=false){if(activeFilterColumn===null||pendingFilterRequest!==null)return;const requestId=++filterRequestCounter,payload={type:'auto-filter',row:selected.row,column:activeFilterColumn,sheetIndex:model?.sheetIndex||0,requestId};if(clear)payload.clear=true;else if(filterMode.value==='color'){const value=document.getElementById('filterColorValue').value;payload.colorFilter={cellColor:document.getElementById('filterColorTarget').value==='fill',color:value==='none'?null:value==='custom'?document.getElementById('filterColorCustom').value:value}}else if(filterMode.value==='icons'){const [iconSet,iconId]=document.getElementById('filterIconValue').value.split(':');payload.iconFilter={iconSet,iconId:Number(iconId)}}else if(filterMode.value==='dynamic')payload.dynamic={type:document.getElementById('filterDynamicType').value};else if(filterMode.value==='dates'){payload.dateGroupText=document.getElementById('filterDateGroups').value;payload.values=document.getElementById('filterDateKeepValues').checked?[...dateMixedValues]:[];payload.blanks=document.getElementById('filterDateBlanks').checked}else if(filterMode.value==='top10')payload.top10={top:document.getElementById('filterTopDirection').value==='top',percent:document.getElementById('filterTopUnit').value==='percent',value:Number(document.getElementById('filterTopValue').value)};else if(filterMode.value==='conditions')payload.custom={and:filterLogic.value==='and',conditions:[encodedFilterCondition(1),...(filterUseSecond.checked?[encodedFilterCondition(2)]:[])]};else{const item=filterOptions.find(option=>option.blank);payload.values=filterOptions.filter(option=>!option.blank&&option.checked).map(option=>option.value);payload.blanks=Boolean(item?.checked)}pendingFilterRequest=requestId;document.getElementById('filterEditor').disabled=true;document.getElementById('filterApply').disabled=true;document.getElementById('filterClear').disabled=true;vscode.postMessage(payload)}
  function closeFilterPanel(){filterPanel.hidden=true;pendingFilterRequest=null;document.getElementById('filterEditor').disabled=false;document.getElementById('filterApply').disabled=false;document.getElementById('filterClear').disabled=false;activeFilterColumn=null;filterOptions=[];filterSearch.value='';filterPage=0;document.getElementById('grid').focus()}
  function matchingFilterOptions(){const query=filterSearch.value.trim().toLocaleLowerCase();return filterOptions.filter(item=>!query||(item.blank?t('filter-blank'):item.value).toLocaleLowerCase().includes(query))}
  function updateFilterSelectAll(){const matching=matchingFilterOptions();filterSelectAll.checked=matching.length>0&&matching.every(item=>item.checked);filterSelectAll.indeterminate=matching.some(item=>item.checked)&&!filterSelectAll.checked}
  function renderFilterOptionPage(){const matching=matchingFilterOptions(),shown=matching.slice(0,(filterPage+1)*FILTER_OPTION_PAGE_SIZE);filterList.replaceChildren();for(const item of shown){const optionLabel=document.createElement('label'),input=document.createElement('input'),text=document.createElement('span');input.type='checkbox';input.checked=item.checked;input.onchange=()=>{item.checked=input.checked;updateFilterSelectAll()};text.textContent=item.blank?t('filter-blank'):item.value;optionLabel.append(input);optionLabel.append(text);filterList.append(optionLabel)}filterShowMore.hidden=shown.length>=matching.length;updateFilterSelectAll()}
  function renderFilterOptions(data){if(data.sheetIndex!==model?.sheetIndex||data.column!==activeFilterColumn)return;const selectedValues=new Set(data.options.selectedValues);filterOptions=data.options.values.map(value=>({value,checked:selectedValues.has(value)}));if(data.options.blanks)filterOptions.push({value:'',checked:data.options.includeBlanks,blank:true});filterSearch.value='';filterPage=0;const custom=data.options.custom,top10=data.options.top10,dateGroups=data.options.dateGroups;dateMixedValues=data.options.dateValues||[];document.getElementById('filterDateGroups').value=(dateGroups||[]).map(group=>{const fields=['year','month','day','hour','minute','second'],last=fields.indexOf(group.dateTimeGrouping);return fields.slice(0,last+1).map((key,index)=>(index===0?'':index<=2?'-':index===3?'T':':')+String(group[key]).padStart(index===0?4:2,'0')).join('')}).join(String.fromCharCode(10));document.getElementById('filterDateBlanks').checked=Boolean(dateGroups&&data.options.includeBlanks);document.getElementById('filterDateKeepValues').checked=true;document.getElementById('filterDateKeepLabel').hidden=!dateMixedValues.length;document.getElementById('filterDateValuesCount').textContent=t('filter-date-keep-values')+' ('+dateMixedValues.length+')';document.getElementById('filterDynamicType').value=data.options.dynamic?.type||'today';filterColorPalette=data.options.colors||{fills:[],fonts:[]};document.getElementById('filterColorTarget').value=data.options.colorFilter?.cellColor===false?'font':'fill';renderColorChoices(data.options.colorFilter?.color);const iconSelect=document.getElementById('filterIconValue'),iconFilters=[...(data.options.iconFilters||[])];if(data.options.iconFilter&&!iconFilters.some(item=>item.iconSet===data.options.iconFilter.iconSet&&item.iconId===data.options.iconFilter.iconId))iconFilters.push({...data.options.iconFilter,count:0});iconSelect.replaceChildren(...iconFilters.map(item=>{const option=document.createElement('option');option.value=item.iconSet+':'+item.iconId;option.textContent=item.iconSet+' · '+t('filter-icon-number')+' '+(item.iconId+1)+' ('+item.count+')';return option}));iconSelect.value=data.options.iconFilter?data.options.iconFilter.iconSet+':'+data.options.iconFilter.iconId:iconFilters[0]?.iconSet+':'+iconFilters[0]?.iconId||'';filterMode.value=data.options.iconFilter?'icons':data.options.colorFilter?'color':data.options.dynamic?'dynamic':dateGroups?'dates':top10?'top10':custom?'conditions':'values';document.getElementById('filterTopDirection').value=top10?.top===false?'bottom':'top';document.getElementById('filterTopValue').value=String(top10?.value??10);document.getElementById('filterTopUnit').value=top10?.percent?'percent':'items';filterUseSecond.checked=Boolean(custom&&custom.conditions.length===2);filterLogic.value=custom?.and?'and':custom?'or':'and';for(const index of [1,2]){const condition=filterUiCondition(custom?.conditions[index-1]||{operator:'equal',value:''});document.getElementById('filterOperator'+index).value=condition.operator;document.getElementById('filterConditionValue'+index).value=condition.value}for(const option of Array.from(filterMode.options||[]))option.disabled=model?.kind==='csv'&&!['values','conditions'].includes(option.value);syncFilterMode();renderFilterOptionPage();document.getElementById('filterTitle').textContent=t('filter-dialog-title')+' · '+label(data.column+1);document.getElementById('filterSessionNote').hidden=model?.kind!=='csv';filterPanel.hidden=false}
  document.getElementById('filterValuesButton').onclick=()=>{if(pendingFilterRequest!==null)return;if(!['xlsx','csv'].includes(model?.kind)||model.readOnly||(!model.autoFilter&&!(model.rowCount>1&&model.columnCount>0)))return;activeFilterColumn=selected.column;vscode.postMessage({type:'filter-options',row:selected.row,column:activeFilterColumn,sheetIndex:model.sheetIndex})};
  filterSearch.oninput=()=>{filterPage=0;renderFilterOptionPage()};filterShowMore.onclick=()=>{filterPage++;renderFilterOptionPage()};
  filterSelectAll.onchange=()=>{for(const item of matchingFilterOptions())item.checked=filterSelectAll.checked;renderFilterOptionPage()};
  document.getElementById('filterApply').onclick=()=>submitFilter();
  document.getElementById('filterClear').onclick=()=>submitFilter(true);document.getElementById('filterCancel').onclick=closeFilterPanel;
  filterPanel.onkeydown=e=>{if(e.isComposing||e.keyCode===229)return;if(e.key==='Escape'){e.preventDefault();e.stopPropagation();closeFilterPanel()}else if(e.key==='Enter'&&e.altKey&&e.target?.tagName==='TEXTAREA'){e.preventDefault();e.stopPropagation();e.target.setRangeText(String.fromCharCode(10),e.target.selectionStart,e.target.selectionEnd,'end')}else if(e.key==='Enter'&&(['INPUT','TEXTAREA'].includes(e.target?.tagName)||['filterDynamicType','filterColorValue','filterColorCustom'].includes(e.target?.id))&&e.target.type!=='checkbox'){e.preventDefault();e.stopPropagation();submitFilter()}};
  window.addEventListener('message',e=>{if(['table-style-applied','table-style-error'].includes(e.data?.type)){const data=e.data,attempt=pendingTableStyle;if(!attempt||data.requestId!==attempt.requestId||data.sheetIndex!==attempt.sheetIndex||data.tableName!==attempt.tableName)return;pendingTableStyle=null;if(data.type==='table-style-error'){syncTableStyleControls();document.getElementById('notice').textContent=data.message}else{tableStyleSelectionTouched=false;tableStyleOptionsTouched=false;selectedStyleTableKey=null;syncTableStyleControls();document.getElementById('notice').textContent=t('table-style-applied')+' '+data.tableName}return}if(['table-created','table-create-error'].includes(e.data?.type)){const data=e.data;if(!pendingCreateTable||data.requestId!==pendingCreateTable.requestId||data.sheetIndex!==pendingCreateTable.sheetIndex||data.sheetIndex!==model?.sheetIndex)return;const attempted=pendingCreateTable;pendingCreateTable=null;syncCreateTableControls();if(data.type==='table-create-error'){document.getElementById('notice').textContent=data.message;document.getElementById('newTableName').focus()}else{const input=document.getElementById('newTableName');if(input.value.trim()===attempted.name)input.value='';document.getElementById('notice').textContent=t('table-created')+' '+data.name}return}if(['table-tab-applied','table-tab-error'].includes(e.data?.type)){const data=e.data,attempt=pendingTableTab;if(!attempt||data.requestId!==attempt.requestId||data.sheetIndex!==attempt.sheetIndex||data.tableName!==attempt.tableName)return;pendingTableTab=null;if(model?.sheetIndex!==attempt.sheetIndex||selected.row!==attempt.row||selected.column!==attempt.column)return;const active=document.activeElement;if(active&&active!==document.body&&active!==formulaInput&&!document.getElementById('grid').contains(active))return;if(data.type==='table-tab-applied')scrollToCell({row:data.row,column:data.column});else{document.getElementById('notice').textContent=data.message;if(attempt.value!==undefined){formulaInput.value=attempt.value;formulaDirty=true;formulaInput.focus()}}return}if(['calculated-column-applied','calculated-column-error'].includes(e.data?.type)){const data=e.data;if(!pendingCalculatedColumn||data.requestId!==pendingCalculatedColumn.requestId||data.sheetIndex!==pendingCalculatedColumn.sheetIndex||data.tableName!==pendingCalculatedColumn.tableName||data.column!==pendingCalculatedColumn.column||data.sheetIndex!==model?.sheetIndex)return;pendingCalculatedColumn=null;syncCalculatedControls();syncTableTotalsControls();if(data.type==='calculated-column-error'){document.getElementById('notice').textContent=data.message;document.getElementById('calculatedColumnScope').focus()}else{document.getElementById('notice').textContent=t('calculated-column-applied')+' '+data.filled+' · '+t('calculated-column-kept')+' '+data.preserved;scrollToCell({row:data.row,column:data.column})}return}if(['table-totals-applied','table-totals-error'].includes(e.data?.type)){const data=e.data;if(!pendingTableTotals||data.requestId!==pendingTableTotals.requestId||data.sheetIndex!==pendingTableTotals.sheetIndex||data.tableName!==pendingTableTotals.tableName||data.sheetIndex!==model?.sheetIndex)return;const attempted=pendingTableTotals;pendingTableTotals=null;syncTableTotalsControls();if(data.type==='table-totals-error'){document.getElementById('notice').textContent=data.message;if(attempted.function){document.getElementById('tableTotalsFunction').value=attempted.function;document.getElementById('applyTableTotalFunction').disabled=false;document.getElementById('tableTotalsFunction').focus()}else document.getElementById('tableTotals').focus()}else scrollToCell({row:data.row,column:data.column});return}if(['filter-applied','filter-error'].includes(e.data?.type)){if(e.data.sheetIndex!==model?.sheetIndex||e.data.column!==activeFilterColumn||e.data.requestId!==pendingFilterRequest)return;if(e.data.type==='filter-applied')closeFilterPanel();else{pendingFilterRequest=null;document.getElementById('filterEditor').disabled=false;document.getElementById('filterApply').disabled=false;document.getElementById('filterClear').disabled=false;document.getElementById('notice').textContent=e.data.message;(filterMode.value==='icons'?document.getElementById('filterIconValue'):filterMode.value==='color'?document.getElementById('filterColorValue'):filterMode.value==='dynamic'?document.getElementById('filterDynamicType'):filterMode.value==='dates'?document.getElementById('filterDateGroups'):filterMode.value==='top10'?document.getElementById('filterTopValue'):filterMode.value==='conditions'?document.getElementById('filterConditionValue1'):filterSearch).focus()}return}if(e.data?.type==='row-layout'){if(e.data.sheetIndex!==model?.sheetIndex||rowResize)return;model.rowHeights=e.data.rowHeights;updateRowGeometry();syncFillHandle();if(!document.querySelector('.cell-editor')&&!formulaDirty&&!columnResize)requestRange();return}if(e.data?.type==='row-fit-request'){if(!e.data.background&&e.data.sheetIndex!==model?.sheetIndex)return;const measure=document.createElement('div'),heights=[];measure.className='cell';measure.style.position='absolute';measure.style.left='-10000px';measure.style.top='0';measure.style.visibility='hidden';measure.style.height='auto';measure.style.minHeight='0';measure.style.lineHeight='normal';document.body.append(measure);try{for(const row of e.data.rows){let pixels=e.data.defaultHeight||26;for(const cell of row.cells){measure.style.width=cell.width+'px';measure.style.fontSize=cell.fontSize?cell.fontSize+'pt':'13px';measure.style.fontFamily=cell.fontFamily?'"'+cell.fontFamily.replaceAll('"','')+'"':'inherit';measure.style.fontWeight=cell.bold?'700':'normal';measure.style.fontStyle=cell.italic?'italic':'normal';measure.style.whiteSpace=cell.wrapText?'pre-wrap':'nowrap';measure.style.overflowWrap=cell.wrapText?'anywhere':'normal';measure.textContent=cell.text||' ';pixels=Math.max(pixels,measure.getBoundingClientRect().height)}heights.push({row:row.row,height:Math.min(409.5,Math.max(.75,Math.ceil(pixels*3)/4))})}}finally{measure.remove()}vscode.postMessage({type:'row-fit-result',requestId:e.data.requestId,sheetIndex:e.data.sheetIndex,heights});return}if(e.data?.type==='error'){pendingCutRange=null;refreshCutHighlights();document.getElementById('notice').textContent=e.data.message;return}if(e.data?.type==='selection-summary'){if(e.data.requestId!==summaryRequest)return;const target=document.getElementById('selectionSummary');if(e.data.tooLarge){target.textContent=t('selection-too-large');return}const format=value=>new Intl.NumberFormat(undefined,{maximumFractionDigits:10}).format(value);target.textContent=t('selection-count')+': '+e.data.count+' · '+t('selection-numbers')+': '+e.data.numericCount+(e.data.sum===null?'':' · '+t('selection-sum')+': '+format(e.data.sum)+' · '+t('selection-average')+': '+format(e.data.average));return}if(e.data?.type==='cut-ready'&&e.data.sheetIndex===model?.sheetIndex){pendingCutRange=e.data.range;refreshCutHighlights();return}if(e.data?.type==='cut-complete'){pendingCutRange=null;refreshCutHighlights();return}if(e.data?.type==='model'){acceptModel(e.data);return}if(e.data?.type==='range'){renderRange(e.data);return}if(e.data?.type==='auto-fit-width'){if(e.data.requestId!==autoFitRequest||e.data.sheetIndex!==model?.sheetIndex||e.data.column!==selected.column)return;resizeColumn(e.data.column,e.data.width);commitColumnWidth(e.data.column);requestRange();return}if(e.data?.type==='hyperlink-navigation'){if(e.data.sheetIndex!==model?.sheetIndex)return;scrollToCell({row:e.data.row,column:e.data.column},false,true);if(e.data.range){selectionAnchor={row:e.data.range.startRow,column:e.data.range.startColumn};selectedRange={...e.data.range};for(const cell of document.querySelectorAll('.cell')){const row=Number(cell.dataset.row),column=Number(cell.dataset.column),inRange=row>=Math.min(selectedRange.startRow,selectedRange.endRow)&&row<=Math.max(selectedRange.startRow,selectedRange.endRow)&&column>=Math.min(selectedRange.startColumn,selectedRange.endColumn)&&column<=Math.max(selectedRange.startColumn,selectedRange.endColumn);cell.classList.toggle('in-range',inRange);cell.setAttribute('aria-selected',String(inRange))}refreshSelectionSummary();requestRange()}return}if(e.data?.type==='navigation'){if(e.data.requestId!==navigationRequest)return;scrollToCell({row:e.data.row,column:e.data.column},e.data.extend);return}if(e.data?.type==='current-region'){if(e.data.requestId!==currentRegionRequest||e.data.sheetIndex!==model?.sheetIndex||e.data.row!==selected.row||e.data.column!==selected.column)return;const range=e.data.range;selectionAnchor={row:range.top,column:range.left};selectedRange={startRow:range.top,endRow:range.bottom,startColumn:range.left,endColumn:range.right};currentRegionStage=1;for(const cell of document.querySelectorAll('.cell')){const row=Number(cell.dataset.row),column=Number(cell.dataset.column),inRange=row>=range.top&&row<=range.bottom&&column>=range.left&&column<=range.right;cell.classList.toggle('in-range',inRange);cell.setAttribute('aria-selected',String(inRange))}refreshSelectionSummary();requestRange();return}if(e.data?.type==='filter-options'){renderFilterOptions(e.data);return}if(e.data?.type==='search-navigation'){if(e.data.requestId!==searchRequest||e.data.sheetIndex!==model?.sheetIndex)return;scrollToCell({row:e.data.row,column:e.data.column},false,false);refreshSearchHighlights();return}if(e.data?.type==='replace-error'){if(e.data.requestId!==pendingReplaceRequest)return;pendingReplaceRequest=0;replaceCurrentButton.disabled=Boolean(model?.readOnly);replaceAllButton.disabled=Boolean(model?.readOnly);replaceStatus.textContent=e.data.message;return}if(e.data?.type==='replace-result'){if(e.data.requestId!==pendingReplaceRequest)return;pendingReplaceRequest=0;replaceCurrentButton.disabled=Boolean(model?.readOnly);replaceAllButton.disabled=Boolean(model?.readOnly);replaceStatus.textContent=t('replace-result').replace('{count}',String(e.data.replaced||0)).replace('{formulas}',String(e.data.skippedFormulas||0));if(e.data.query===searchInput.value){matches=(e.data.matches||[]).map(match=>({...match,sheetIndex:Number.isInteger(match.sheetIndex)?match.sheetIndex:e.data.sheetIndex??model?.sheetIndex??0}));syncSearchMatchesForSheet();matchIndex=0;if(e.data.mode==='replace-current'&&e.data.target&&matches.length){const target=e.data.target,targetSheet=target.sheetIndex??e.data.sheetIndex??model?.sheetIndex??0,next=matches.findIndex(match=>match.sheetIndex>targetSheet||match.sheetIndex===targetSheet&&(match.row>target.row||match.row===target.row&&match.column>target.column));matchIndex=next>=0?next:0}searchedQuery=e.data.query;document.getElementById('searchCount').textContent=e.data.total?e.data.total+ui['match-separator']+t('matches'):t('no-matches');refreshSearchHighlights();if(matches.length)focusSearchMatch()}return}if(e.data?.type==='search'){if(e.data.requestId!==searchRequest||e.data.query!==document.getElementById('search').value)return;searchPagePending=false;matches=(e.data.matches||[]).map(match=>({...match,sheetIndex:Number.isInteger(match.sheetIndex)?match.sheetIndex:e.data.sheetIndex??model?.sheetIndex??0}));syncSearchMatchesForSheet();matchIndex=0;document.getElementById('searchCount').textContent=e.data.total?e.data.total+ui['match-separator']+t('matches'):t('no-matches');refreshSearchHighlights();if(matches.length)focusSearchMatch()}});document.addEventListener('click',event=>{if(event.target?.id!=='resizeTable'||model?.readOnly||model?.kind!=='xlsx')return;const range={top:Math.min(selectedRange.startRow,selectedRange.endRow),bottom:Math.max(selectedRange.startRow,selectedRange.endRow),left:Math.min(selectedRange.startColumn,selectedRange.endColumn),right:Math.max(selectedRange.startColumn,selectedRange.endColumn)},table=(model.tables||[]).find(item=>item.range.top===range.top&&range.left<=item.range.right&&range.right>=item.range.left&&range.bottom>=range.top+Math.max(1,(item.headerRowCount||0)+(item.totalsRowCount||0))-1&&range.right>=range.left);if(table)vscode.postMessage({type:'resize-table',tableName:table.name,range,sheetIndex:model.sheetIndex||0})});document.getElementById('chartPreview').addEventListener('click',event=>{const remove=event.target?.closest?.('[data-chart-delete]'),source=event.target?.closest?.('[data-chart-source]'),placement=event.target?.closest?.('[data-chart-placement-apply]'),button=remove||source||placement||event.target?.closest?.('[data-chart-edit]');if(!button||model?.readOnly||model?.kind!=='xlsx')return;const card=button.closest('[data-chart-index]');if(!card)return;if(placement){const value=card.querySelector('[data-chart-placement]')?.value;if(!['twoCell','oneCell','absolute'].includes(value))return;commitFormulaEdit();vscode.postMessage({type:'chart-placement',chartIndex:Number(card.dataset.chartIndex),placement:value,sheetIndex:model.sheetIndex||0});return}if(remove){vscode.postMessage({type:'delete-chart',chartIndex:Number(card.dataset.chartIndex),sheetIndex:model.sheetIndex||0});return}if(source){const range={top:Math.min(selectedRange.startRow,selectedRange.endRow),bottom:Math.max(selectedRange.startRow,selectedRange.endRow),left:Math.min(selectedRange.startColumn,selectedRange.endColumn),right:Math.max(selectedRange.startColumn,selectedRange.endColumn)};if(range.bottom<=range.top||range.right<=range.left)return;vscode.postMessage({type:'chart-source',chartIndex:Number(card.dataset.chartIndex),range,sheetIndex:model.sheetIndex||0});return}const title=card.querySelector('[data-chart-title]'),names=[...card.querySelectorAll('[data-chart-series-name]')];if(!title)return;vscode.postMessage({type:'edit-chart',chartIndex:Number(card.dataset.chartIndex),title:title.value,seriesNames:names.map(input=>input.value),sheetIndex:model.sheetIndex||0})});document.getElementById('insertChart').onclick=()=>{if(model?.readOnly||model?.kind!=='xlsx')return;const range={top:Math.min(selectedRange.startRow,selectedRange.endRow),bottom:Math.max(selectedRange.startRow,selectedRange.endRow),left:Math.min(selectedRange.startColumn,selectedRange.endColumn),right:Math.max(selectedRange.startColumn,selectedRange.endColumn)};vscode.postMessage({type:'insert-chart',range,chartType:document.getElementById('chartType').value,sheetIndex:model.sheetIndex||0})};document.getElementById('createTable').onclick=submitCreateTable;document.getElementById('applyTableStyle').onclick=submitTableStyle;document.getElementById('tableStyleName').onchange=()=>{tableStyleSelectionTouched=true;if(document.getElementById('tableStyleNote'))document.getElementById('tableStyleNote').textContent='';syncTableStyleControls()};for(const id of ['tableStyleFirstColumn','tableStyleLastColumn','tableStyleRows','tableStyleColumns'])document.getElementById(id).onchange=()=>{tableStyleOptionsTouched=true};document.getElementById('newTableName').addEventListener('keydown',event=>{if(event.key==='Enter'&&!event.isComposing){event.preventDefault();submitCreateTable()}});${sortDialogBootstrap}${textToColumnsBootstrap}${removeDuplicatesBootstrap}vscode.postMessage({type:'ready'});
  </script></body></html>`;
}

function renderHtml(document, webview, extensionUri, language = 'en') {
  const localizedLanguage = languageOf(language), xlsx = document?.kind === 'xlsx';
  const conditionalFormatButton = xlsx ? '<button id="dataBars" type="button" title="' + message('conditional-format-title', localizedLanguage) + '">' + message('conditional-format-button', localizedLanguage) + '</button>' : '';
  const cellValueRuleButton = xlsx ? '<button id="cellValueRule" type="button" title="' + message('conditional-format-cell-rule-title', localizedLanguage) + '">' + message('conditional-format-cell-rule-button', localizedLanguage) + '</button>' : '';
  const colorScaleButton = xlsx ? '<button id="colorScale" type="button" title="' + message('conditional-format-scale-title', localizedLanguage) + '">' + message('conditional-format-scale-button', localizedLanguage) + '</button>' : '';
  const duplicateValuesButton = xlsx ? '<button id="duplicateValues" type="button" title="' + message('conditional-format-duplicates-title', localizedLanguage) + '">' + message('conditional-format-duplicates-button', localizedLanguage) + '</button>' : '';
  const iconSetButton = xlsx ? '<button id="iconSet" type="button" title="' + message('conditional-format-icon-set-title', localizedLanguage) + '">' + message('conditional-format-icon-set-button', localizedLanguage) + '</button>' : '';
  const clearConditionalFormatsButton = xlsx ? '<button id="clearConditionalFormats" type="button" title="' + message('conditional-format-clear-title', localizedLanguage) + '">' + message('conditional-format-clear-button', localizedLanguage) + '</button>' : '';
  const manageConditionalFormatsButton = xlsx ? '<button id="manageConditionalFormats" type="button" title="' + message('conditional-format-manage-title', localizedLanguage) + '">' + message('conditional-format-manage-button', localizedLanguage) + '</button>' : '';
  const validationButton = xlsx ? '<button id="dataValidation" type="button" title="' + message('data-validation-title', localizedLanguage) + '">' + message('data-validation-button', localizedLanguage) + '</button>' : '';
  const hyperlinkButtons = xlsx ? '<button id="editHyperlink" type="button" title="' + message('hyperlink-edit-title', localizedLanguage) + '">' + message('hyperlink-edit', localizedLanguage) + '</button><button id="removeHyperlink" type="button" title="' + message('hyperlink-remove-title', localizedLanguage) + '">' + message('hyperlink-remove', localizedLanguage) + '</button>' : '';
  const definedNamesButton = xlsx ? '<button id="manageDefinedNames" type="button" title="' + message('defined-name-title', localizedLanguage) + '">' + message('defined-name-button', localizedLanguage) + '</button>' : '';
  const html = renderHtmlRaw(document, webview, extensionUri, localizedLanguage).replace('<button id="formatNumber"', cellValueRuleButton + conditionalFormatButton + colorScaleButton + duplicateValuesButton + iconSetButton + manageConditionalFormatsButton + clearConditionalFormatsButton + validationButton + hyperlinkButtons + definedNamesButton + '<button id="formatNumber"');
  const setup = xlsx ? "document.getElementById('dataBars').onclick=()=>{if(model?.readOnly||model?.kind!=='xlsx')return;commitFormulaEdit();const range={top:Math.min(selectedRange.startRow,selectedRange.endRow),bottom:Math.max(selectedRange.startRow,selectedRange.endRow),left:Math.min(selectedRange.startColumn,selectedRange.endColumn),right:Math.max(selectedRange.startColumn,selectedRange.endColumn)};vscode.postMessage({type:'manage-data-bars',range,sheetIndex:model.sheetIndex||0})};document.getElementById('colorScale').onclick=()=>{if(model?.readOnly||model?.kind!=='xlsx')return;commitFormulaEdit();const range={top:Math.min(selectedRange.startRow,selectedRange.endRow),bottom:Math.max(selectedRange.startRow,selectedRange.endRow),left:Math.min(selectedRange.startColumn,selectedRange.endColumn),right:Math.max(selectedRange.startColumn,selectedRange.endColumn)};vscode.postMessage({type:'manage-color-scale',range,sheetIndex:model.sheetIndex||0})};document.getElementById('dataValidation').onclick=()=>{if(model?.readOnly||model?.kind!=='xlsx')return;commitFormulaEdit();const range={top:Math.min(selectedRange.startRow,selectedRange.endRow),bottom:Math.max(selectedRange.startRow,selectedRange.endRow),left:Math.min(selectedRange.startColumn,selectedRange.endColumn),right:Math.max(selectedRange.startColumn,selectedRange.endColumn)};vscode.postMessage({type:'manage-data-validation',range,sheetIndex:model.sheetIndex||0})};function syncHyperlinkControls(){const edit=document.getElementById('editHyperlink'),remove=document.getElementById('removeHyperlink');if(!edit||!remove)return;const disabled=Boolean(model?.readOnly||model?.kind!=='xlsx'),link=(model?.hyperlinks||[]).some(item=>selected.row>=item.top&&selected.row<=item.bottom&&selected.column>=item.left&&selected.column<=item.right);edit.disabled=disabled;remove.disabled=disabled||!link}document.getElementById('editHyperlink').onclick=()=>{if(model?.readOnly||model?.kind!=='xlsx')return;commitFormulaEdit();vscode.postMessage({type:'edit-hyperlink',row:selected.row,column:selected.column,sheetIndex:model.sheetIndex||0})};document.getElementById('removeHyperlink').onclick=()=>{if(model?.readOnly||model?.kind!=='xlsx')return;commitFormulaEdit();vscode.postMessage({type:'remove-hyperlink',row:selected.row,column:selected.column,sheetIndex:model.sheetIndex||0})};window.addEventListener('message',event=>{if(event.data?.type==='model'){const dataBars=document.getElementById('dataBars'),colorScale=document.getElementById('colorScale'),validation=document.getElementById('dataValidation');if(dataBars)dataBars.disabled=Boolean(event.data.readOnly||event.data.kind!=='xlsx');if(colorScale)colorScale.disabled=Boolean(event.data.readOnly||event.data.kind!=='xlsx');if(validation)validation.disabled=Boolean(event.data.readOnly||event.data.kind!=='xlsx');syncHyperlinkControls()}});" : '';
  const clearConditionalFormatsSetup = xlsx ? "document.getElementById('clearConditionalFormats').onclick=()=>{if(model?.readOnly||model?.kind!=='xlsx')return;commitFormulaEdit();const range={top:Math.min(selectedRange.startRow,selectedRange.endRow),bottom:Math.max(selectedRange.startRow,selectedRange.endRow),left:Math.min(selectedRange.startColumn,selectedRange.endColumn),right:Math.max(selectedRange.startColumn,selectedRange.endColumn)};vscode.postMessage({type:'clear-conditional-formats',range,sheetIndex:model.sheetIndex||0})};window.addEventListener('message',event=>{if(event.data?.type==='model'){const button=document.getElementById('clearConditionalFormats');if(button)button.disabled=Boolean(event.data.readOnly||event.data.kind!=='xlsx')}});" : '';
  const manageConditionalFormatsSetup = xlsx ? "document.getElementById('manageConditionalFormats').onclick=()=>{if(model?.readOnly||model?.kind!=='xlsx')return;commitFormulaEdit();vscode.postMessage({type:'manage-conditional-format-rules',sheetIndex:model.sheetIndex||0})};window.addEventListener('message',event=>{if(event.data?.type==='model'){const button=document.getElementById('manageConditionalFormats');if(button)button.disabled=Boolean(event.data.readOnly||event.data.kind!=='xlsx')}});" : '';
  const iconSetSetup = xlsx ? "document.getElementById('iconSet').onclick=()=>{if(model?.readOnly||model?.kind!=='xlsx')return;commitFormulaEdit();const range={top:Math.min(selectedRange.startRow,selectedRange.endRow),bottom:Math.max(selectedRange.startRow,selectedRange.endRow),left:Math.min(selectedRange.startColumn,selectedRange.endColumn),right:Math.max(selectedRange.startColumn,selectedRange.endColumn)};vscode.postMessage({type:'manage-icon-set',range,sheetIndex:model.sheetIndex||0})};window.addEventListener('message',event=>{if(event.data?.type==='model'){const button=document.getElementById('iconSet');if(button)button.disabled=Boolean(event.data.readOnly||event.data.kind!=='xlsx')}});" : '';
  const definedNamesSetup = xlsx ? "document.getElementById('manageDefinedNames').onclick=()=>{if(model?.readOnly||model?.kind!=='xlsx')return;commitFormulaEdit();vscode.postMessage({type:'manage-defined-names',sheetIndex:model.sheetIndex||0})};window.addEventListener('message',event=>{if(event.data?.type==='model'){const button=document.getElementById('manageDefinedNames');if(button)button.disabled=Boolean(event.data.readOnly||event.data.kind!=='xlsx')}});" : '';
  const duplicateValuesSetup = xlsx ? "document.getElementById('duplicateValues').onclick=()=>{if(model?.readOnly||model?.kind!=='xlsx')return;commitFormulaEdit();const range={top:Math.min(selectedRange.startRow,selectedRange.endRow),bottom:Math.max(selectedRange.startRow,selectedRange.endRow),left:Math.min(selectedRange.startColumn,selectedRange.endColumn),right:Math.max(selectedRange.startColumn,selectedRange.endColumn)};vscode.postMessage({type:'apply-duplicate-values',range,sheetIndex:model.sheetIndex||0})};window.addEventListener('message',event=>{if(event.data?.type==='model'){const button=document.getElementById('duplicateValues');if(button)button.disabled=Boolean(event.data.readOnly||event.data.kind!=='xlsx')}});" : '';
  const cellValueRuleSetup = xlsx ? "document.getElementById('cellValueRule').onclick=()=>{if(model?.readOnly||model?.kind!=='xlsx')return;commitFormulaEdit();const range={top:Math.min(selectedRange.startRow,selectedRange.endRow),bottom:Math.max(selectedRange.startRow,selectedRange.endRow),left:Math.min(selectedRange.startColumn,selectedRange.endColumn),right:Math.max(selectedRange.startColumn,selectedRange.endColumn)};vscode.postMessage({type:'manage-cell-value-rule',range,sheetIndex:model.sheetIndex||0})};window.addEventListener('message',event=>{if(event.data?.type==='model'){const button=document.getElementById('cellValueRule');if(button)button.disabled=Boolean(event.data.readOnly||event.data.kind!=='xlsx')}});" : '';
  const qualifiedReferenceSetup = xlsx ? "let qualifiedNavigationRequest=0;const qualifiedAddress=document.getElementById('address');qualifiedAddress.addEventListener('keydown',event=>{if(event.key!=='Enter'||event.isComposing||model?.kind!=='xlsx')return;const text=qualifiedAddress.value.trim(),separator=text.lastIndexOf('!');if(separator<1)return;const sheetText=text.slice(0,separator).trim(),reference=text.slice(separator+1).trim();let sheetName;if(sheetText.startsWith(\"'\")){if(!/^'(?:[^']|'')*'$/.test(sheetText)){event.preventDefault();event.stopImmediatePropagation();document.getElementById('goToStatus').textContent=t('invalid-cell-reference');qualifiedAddress.setAttribute('aria-invalid','true');qualifiedAddress.focus();qualifiedAddress.select();return}sheetName=sheetText.slice(1,-1).replaceAll(\"''\",\"'\")}else{if(sheetText.includes(\"'\")){event.preventDefault();event.stopImmediatePropagation();document.getElementById('goToStatus').textContent=t('invalid-cell-reference');qualifiedAddress.setAttribute('aria-invalid','true');return}sheetName=sheetText}const target=(model.sheets||[]).find(sheet=>sheet.name.toLocaleLowerCase()===sheetName.toLocaleLowerCase()),parts=reference.split(':');const parseAddressPart=value=>{const match=/^\\$?([A-Za-z]{1,3})\\$?([1-9]\\d*)$/.exec(value);if(!match)return null;let column=0;for(const character of match[1].toUpperCase())column=column*26+character.charCodeAt(0)-64;const row=Number(match[2])-1;column--;if(!Number.isSafeInteger(row)||!Number.isSafeInteger(column)||row<0||column<0||row>=1048576||column>=16384)return null;return{row,column}};const start=parts.length<=2?parseAddressPart(parts[0]):null,end=parts.length===2?parseAddressPart(parts[1]):start;if(parts.some(part=>!part.trim())||!start||!end||!target){event.preventDefault();event.stopImmediatePropagation();document.getElementById('goToStatus').textContent=t('invalid-cell-reference');qualifiedAddress.setAttribute('aria-invalid','true');qualifiedAddress.focus();qualifiedAddress.select();return}event.preventDefault();event.stopImmediatePropagation();if(target.index===model.sheetIndex){qualifiedAddress.value=reference;applyGoToAddress();return}const requestId=++qualifiedNavigationRequest;commitFormulaEdit();qualifiedAddress.setAttribute('aria-invalid','false');document.getElementById('goToStatus').textContent='';vscode.postMessage({type:'cell-reference-navigate',requestId,sheetIndex:model.sheetIndex,targetSheetIndex:target.index,range:{startRow:start.row,endRow:end.row,startColumn:start.column,endColumn:end.column}})},true);window.addEventListener('message',event=>{const data=event.data;if(data?.type!=='cell-reference-navigation'||data.requestId!==qualifiedNavigationRequest||data.sheetIndex!==model?.sheetIndex)return;qualifiedNavigationRequest=0;const range=data.range;if(!range)return;scrollToCell({row:range.startRow,column:range.startColumn},false,true);selectionAnchor={row:range.startRow,column:range.startColumn};selectedRange={...range};const top=Math.min(range.startRow,range.endRow),bottom=Math.max(range.startRow,range.endRow),left=Math.min(range.startColumn,range.endColumn),right=Math.max(range.startColumn,range.endColumn);for(const cell of document.querySelectorAll('.cell')){const row=Number(cell.dataset.row),column=Number(cell.dataset.column),inRange=row>=top&&row<=bottom&&column>=left&&column<=right;cell.classList.toggle('in-range',inRange);cell.setAttribute('aria-selected',String(inRange))}qualifiedAddress.value=label(range.startColumn+1)+(range.startRow+1);qualifiedAddress.setAttribute('aria-invalid','false');document.getElementById('goToStatus').textContent='';document.getElementById('grid').focus({preventScroll:true});syncFillHandle();refreshSelectionSummary();refreshSearchHighlights();requestRange()});" : '';
  const definedNameNavigationSetup = xlsx ? "let definedNameNavigationRequest=0;const nameAddress=document.getElementById('address');nameAddress.addEventListener('keydown',event=>{if(event.key!=='Enter'||event.isComposing||model?.kind!=='xlsx')return;const name=nameAddress.value.trim(),key=name.toLocaleLowerCase(),definition=(model.definedNames||[]).find(item=>item.localSheetId===model.sheetIndex&&item.name.toLocaleLowerCase()===key)||(model.definedNames||[]).find(item=>item.localSheetId===null&&item.name.toLocaleLowerCase()===key);if(!definition)return;event.preventDefault();event.stopImmediatePropagation();const source=definition.source;if(!source||definition.printArea||definition.printTitles||definition.name.toLocaleLowerCase().startsWith('_xlnm.')){document.getElementById('goToStatus').textContent=t('invalid-cell-reference');nameAddress.setAttribute('aria-invalid','true');nameAddress.focus();nameAddress.select();return}const targetSheet=(model.sheets||[]).find(sheet=>sheet.name.toLocaleLowerCase()===source.sheetName.toLocaleLowerCase())?.index;if(!Number.isInteger(targetSheet)){document.getElementById('goToStatus').textContent=t('invalid-cell-reference');nameAddress.setAttribute('aria-invalid','true');return}let columnName=column=>{let result='';for(;column>0;column=Math.floor((column-1)/26))result=String.fromCharCode(65+(column-1)%26)+result;return result};const start=columnName(source.left+1)+(source.top+1),end=columnName(source.right+1)+(source.bottom+1);if(targetSheet===model.sheetIndex){nameAddress.value=start===end?start:start+':'+end;applyGoToAddress();return}const requestId=++definedNameNavigationRequest;commitFormulaEdit();nameAddress.setAttribute('aria-invalid','false');document.getElementById('goToStatus').textContent='';vscode.postMessage({type:'defined-name-navigate',requestId,name:definition.name,sheetIndex:model.sheetIndex})},true);window.addEventListener('message',event=>{const data=event.data;if(data?.type!=='defined-name-navigation'||data.requestId!==definedNameNavigationRequest||data.sheetIndex!==model?.sheetIndex)return;definedNameNavigationRequest=0;const range=data.range;if(!range||![range.startRow,range.endRow,range.startColumn,range.endColumn].every(Number.isInteger))return;scrollToCell({row:range.startRow,column:range.startColumn},false,true);selectionAnchor={row:range.startRow,column:range.startColumn};selectedRange={...range};const top=Math.min(range.startRow,range.endRow),bottom=Math.max(range.startRow,range.endRow),left=Math.min(range.startColumn,range.endColumn),right=Math.max(range.startColumn,range.endColumn);for(const cell of document.querySelectorAll('.cell')){const row=Number(cell.dataset.row),column=Number(cell.dataset.column),inRange=row>=top&&row<=bottom&&column>=left&&column<=right;cell.classList.toggle('in-range',inRange);cell.setAttribute('aria-selected',String(inRange))}nameAddress.value=label(range.startColumn+1)+(range.startRow+1);nameAddress.setAttribute('aria-invalid','false');document.getElementById('goToStatus').textContent='';document.getElementById('grid').focus({preventScroll:true});syncFillHandle();refreshSelectionSummary();refreshSearchHighlights();requestRange()});" : '';
  return html.replace("document.getElementById('createTable').onclick=submitCreateTable;", setup + cellValueRuleSetup + iconSetSetup + clearConditionalFormatsSetup + manageConditionalFormatsSetup + definedNamesSetup + duplicateValuesSetup + qualifiedReferenceSetup + definedNameNavigationSetup + "document.getElementById('createTable').onclick=submitCreateTable;");
}

function deactivate() {}
module.exports = { activate, deactivate, updateCsvOutlineGroups, autoFilterRange, autoFilterOptions, updateAutoFilterColumn, updateAutoFilterCustomColumn, updateAutoFilterTopColumn, updateAutoFilterDateColumn, updateAutoFilterDynamicColumn, updateAutoFilterColorColumn, currentRegion, chartPreviewHtml, estimateDocumentColumnWidth, findMatches, buildReplaceEdits, buildScopedReplaceEdits, getRange, isValidSheetRequest, resolveInternalHyperlink, matchesSelectedSheet, navigateCell, selectionSummary, parseClipboard, inferClipboardCellKind, parseClipboardDate, parseClipboardTime, parseClipboardPercentage, parseClipboardCurrency, buildDateTimeShortcutEdit, buildAutoSumFormula, buildClipboardKindMatrix, buildClipboardFormulaMatrix, isSameWorkbookClipboard, buildClearEdits, buildMoveEdits, buildSparseColumnMoveEdits, buildFillEdits, buildFillRangeEdits, buildSortEdits, planTextToColumns, buildTextToColumnsEdits, planRemoveDuplicates, removeDuplicateRows, buildNumberFormatEdits, buildFontFormatEdits, buildFontColorEdits, buildAlignmentFormatEdits, buildWrapTextEdits, buildAutomaticRowFitRequest, buildRowFitRequest, validateRowFitResult, buildBorderFormatEdits, buildFillFormatEdits, buildPasteEdits, buildSparseColumnPasteEdits, serializeClipboard, buildCopyPayload, JinSheetDocument, JinSheetProvider, buildPrintDocument, parsePrintTitleInputs, renderHtml, viewModel };
module.exports.buildDataValidationRule = buildDataValidationRule;
