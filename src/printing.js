'use strict';

const crypto = require('node:crypto');
const { validatePrintSettings } = require('./xlsx');
const { defaultHeaderFooter, validateHeaderFooter, renderHeaderFooterText } = require('./header-footer');
const { printAxisRuns, chartPrintLayers } = require('./chart-print');

const DEFAULT_PRINT_SETTINGS = Object.freeze({ orientation: 'portrait', paperSize: 9, pageOrder: 'downThenOver', firstPageNumber: 1, useFirstPageNumber: false, scale: 100, fitToPage: false, fitToWidth: 1, fitToHeight: 1, gridLines: false, headings: false, margins: Object.freeze({ left: .7, right: .7, top: .75, bottom: .75, header: .3, footer: .3 }) });
const escape = value => String(value ?? '').replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&#39;');
const color = value => /^#[0-9a-f]{6}$/i.test(value || '') ? value : null;

function printCellBox(cell, width, height) {
  const sizes = { hair: .5, thin: 1, medium: 2, thick: 3, dashed: 1, dotted: 1, double: 3 }, borders = cell?.style?.borders || {};
  const box = { paddingX: width < 32 ? 0 : 6, paddingY: height < 26 ? 0 : 3 };
  for (const side of ['left', 'right', 'top', 'bottom']) box[side] = sizes[borders[side]?.line] ?? 1;
  if (box.left + box.right > width) { box.left = width < 2 ? 0 : Math.min(box.left, width / 2); box.right = Math.min(box.right, width < 2 ? width : width / 2); }
  if (box.top + box.bottom > height) { box.top = height < 2 ? 0 : Math.min(box.top, height / 2); box.bottom = Math.min(box.bottom, height < 2 ? height : height / 2); }
  box.innerHeight = Math.max(0, height - box.paddingY * 2 - box.top - box.bottom);
  return box;
}

function printCellStyle(cell, gridLines, box) {
  const style = cell?.style || {}, rules = ['overflow:hidden', 'padding:' + box.paddingY + 'px ' + box.paddingX + 'px', 'box-sizing:border-box', 'white-space:' + (style.wrapText ? 'pre-wrap' : 'pre'), 'overflow-wrap:anywhere'];
  rules.push('border:1px solid ' + (gridLines ? '#999999' : 'transparent'));
  if (color(style.color)) rules.push('color:' + style.color);
  if (color(style.fill)) rules.push('background-color:' + style.fill);
  if (Number.isFinite(style.fontSize) && style.fontSize >= 1 && style.fontSize <= 409) rules.push('font-size:' + style.fontSize + 'pt');
  if (typeof style.fontFamily === 'string' && /^[\p{L}\p{N} _.,-]{1,100}$/u.test(style.fontFamily)) rules.push('font-family:"' + style.fontFamily + '"');
  if (style.bold) rules.push('font-weight:700');
  if (style.italic) rules.push('font-style:italic');
  const decoration = [style.underline ? 'underline' : '', style.strike ? 'line-through' : ''].filter(Boolean).join(' ');
  if (decoration) rules.push('text-decoration:' + decoration);
  rules.push('text-align:' + (['left', 'center', 'right'].includes(style.horizontal) ? style.horizontal : cell?.kind === 'number' ? 'right' : 'left'));
  rules.push('vertical-align:' + ({ top: 'top', center: 'middle', bottom: 'bottom' }[style.vertical] || 'bottom'));
  const lines = { hair: '.5px solid', thin: '1px solid', medium: '2px solid', thick: '3px solid', dashed: '1px dashed', dotted: '1px dotted', double: '3px double' };
  for (const side of ['left', 'right', 'top', 'bottom']) {
    const border = style.borders?.[side];
    if (border && lines[border.line]) rules.push('border-' + side + ':' + lines[border.line] + ' ' + (color(border.color) || '#000000'));
    rules.push('border-' + side + '-width:' + box[side] + 'px');
  }
  return escape(rules.join(';'));
}

function columnLabel(column) {
  let label = ''; for (let value = column + 1; value; value = Math.floor((value - 1) / 26)) label = String.fromCharCode(65 + (value - 1) % 26) + label;
  return label;
}

// A merged cell must stay on one page. Group every row/column connected by a merge.
function axisGroups(indexes, sizes, merges, start, end) {
  const groups = [];
  for (let position = 0; position < indexes.length;) {
    let last = position;
    for (let scan = position; scan <= last; scan++) for (const merged of merges) {
      if (indexes[scan] < merged[start] || indexes[scan] > merged[end]) continue;
      while (last + 1 < indexes.length && indexes[last + 1] <= merged[end]) last++;
    }
    const items = indexes.slice(position, last + 1);
    groups.push({ items, size: items.reduce((sum, index) => sum + sizes[index], 0) }); position = last + 1;
  }
  return groups;
}

function splitBands(groups, capacity, breaks = []) {
  const bands = []; let band = [], size = 0, previous = null;
  for (const group of groups) {
    const forced = previous !== null && breaks.some(position => position > previous && position <= group.items[0]);
    if (band.length && (forced || size + group.size > capacity + .001)) { bands.push(band); band = []; size = 0; }
    band.push(...group.items); size += group.size;
    previous = group.items.at(-1);
  }
  if (band.length) bands.push(band);
  return bands;
}

function decoratePrintPage(page, index, total, { headerFooter = defaultHeaderFooter(), context = {}, settings = DEFAULT_PRINT_SETTINGS, headerBudget = { characters: 0 } } = {}) {
  validateHeaderFooter(headerFooter);
  const pageNumber = (settings.useFirstPageNumber ? settings.firstPageNumber ?? 1 : 1) + index;
  const variant = headerFooter.differentFirst && index === 0 ? 'first' : headerFooter.differentOddEven && pageNumber % 2 === 0 ? 'even' : 'odd';
  const scale = Number(/data-print-scale="([\d.e+-]+)"/.exec(page)?.[1]) || 1, margin = settings.margins;
  const values = { ...context, page: pageNumber, pages: total };
  let decoration = '';
  for (const part of ['header', 'footer']) {
    const sections = headerFooter[variant][part];
    if (!['left', 'center', 'right'].some(side => sections[side])) continue;
    const rendered = Object.fromEntries(['left', 'center', 'right'].map(side => [side, renderHeaderFooterText(sections[side], values, headerBudget)]));
    if (!Object.values(rendered).some(Boolean)) continue;
    const extent = (part === 'header' ? margin.top : margin.bottom) - margin[part];
    if (extent < .08) throw new RangeError('머리글·바닥글이 본문과 겹치지 않도록 페이지 여백을 늘리세요.');
    const left = headerFooter.alignWithMargins ? margin.left : 0, right = headerFooter.alignWithMargins ? margin.right : 0;
    const position = part === 'header' ? 'top' : 'bottom', baseSize = 11 * (headerFooter.scaleWithDoc ? scale : 1);
    decoration += '<div class="print-' + part + '" data-hf-variant="' + variant + '" style="' + position + ':' + margin[part] + 'in;left:' + left + 'in;right:' + right + 'in;height:' + extent + 'in">';
    for (const side of ['left', 'center', 'right']) decoration += '<div class="hf-cell" data-hf-part="' + part + '" data-hf-side="' + side + '" style="text-align:' + side + ';font-size:' + baseSize + 'pt" data-hf-font-size="' + baseSize + '">' + '<div class="hf-content">' + rendered[side] + '</div></div>';
    decoration += '</div>';
  }
  page = page.replace(/<!--print-hf-start-->[\s\S]*?<!--print-hf-end-->/g, '');
  return page.replace(/^<section\b[^>]*>/, opening => opening + '<!--print-hf-start-->' + decoration + '<!--print-hf-end-->');
}

function renderPrintDocument({ title, rows, columns, widths, heights, cells, mergedCells = [], charts = [], rowPositions = {}, columnPositions = {}, pageBreaks = { rows: [], columns: [] }, repeatRows = [], repeatColumns = [], outputBudget = { cells: 0, characters: 0 }, headerFooter = defaultHeaderFooter(), headerContext = {}, headerBudget = { characters: 0 }, settings = DEFAULT_PRINT_SETTINGS, language = 'en', note = '' }) {
  validatePrintSettings(settings);
  if (!Array.isArray(repeatRows) || !Array.isArray(repeatColumns) || repeatRows.length > 50 || repeatColumns.length > 10) throw new RangeError('반복 제목은 50행·10열 이하의 유효한 범위를 입력하세요.');
  if (!pageBreaks || !['rows', 'columns'].every(axis => Array.isArray(pageBreaks[axis]) && pageBreaks[axis].length <= 4096 && pageBreaks[axis].every(position => Number.isInteger(position) && position > 0))) throw new RangeError('페이지 나누기 요청이 유효하지 않습니다.');
  if (!rows.length || !columns.length || rows.length * columns.length > 20000) throw new RangeError('인쇄할 표시 셀이 없거나 인쇄 범위가 너무 큽니다.');
  const paper = { 1: [8.5, 11], 8: [11.6929, 16.5354], 9: [8.2677, 11.6929] }[settings.paperSize];
  const [paperWidth, paperHeight] = settings.orientation === 'landscape' ? [...paper].reverse() : paper;
  const margin = settings.margins, width = (paperWidth - margin.left - margin.right) * 96, height = (paperHeight - margin.top - margin.bottom) * 96;
  const headingWidth = settings.headings ? 40 : 0, headingHeight = settings.headings ? 24 : 0;
  const titleRows = [...new Set(repeatRows)].filter(row => rows.includes(row)).sort((left, right) => left - right), titleColumns = [...new Set(repeatColumns)].filter(column => columns.includes(column)).sort((left, right) => left - right);
  const titleRowSet = new Set(titleRows), titleColumnSet = new Set(titleColumns);
  const declaredRows = new Set(repeatRows), declaredColumns = new Set(repeatColumns);
  const rowBreaks = settings.fitToPage ? [] : pageBreaks.rows.filter(position => !declaredRows.has(position)), columnBreaks = settings.fitToPage ? [] : pageBreaks.columns.filter(position => !declaredColumns.has(position));
  const visibleMerges = mergedCells.filter(merged => rows.some(row => row >= merged.top && row <= merged.bottom) && columns.some(column => column >= merged.left && column <= merged.right));
  for (const merged of visibleMerges) {
    const visibleRows = rows.filter(row => row >= merged.top && row <= merged.bottom), visibleColumns = columns.filter(column => column >= merged.left && column <= merged.right);
    if (rowBreaks.some(position => position > visibleRows[0] && position <= visibleRows.at(-1)) || columnBreaks.some(position => position > visibleColumns[0] && position <= visibleColumns.at(-1))) throw new RangeError('페이지 나누기가 병합 셀을 가르지 않도록 위치를 조정하세요.');
  }
  const titleHeight = titleRows.reduce((sum, row) => sum + heights[row], 0), titleWidth = titleColumns.reduce((sum, column) => sum + widths[column], 0);
  const rowGroups = axisGroups(rows.filter(row => !titleRowSet.has(row)), heights, visibleMerges, 'top', 'bottom'), columnGroups = axisGroups(columns.filter(column => !titleColumnSet.has(column)), widths, visibleMerges, 'left', 'right');
  const totalWidth = columns.reduce((sum, column) => sum + widths[column], headingWidth), totalHeight = rows.reduce((sum, row) => sum + heights[row], headingHeight);
  let scale = settings.fitToPage ? Math.min(1, settings.fitToWidth ? width * settings.fitToWidth / totalWidth : 1, settings.fitToHeight ? height * settings.fitToHeight / totalHeight : 1) : settings.scale / 100;
  // Keep indivisible merged blocks visible even if a fixed scale exceeds a page.
  scale = Math.min(scale, width / (Math.max(0, ...columnGroups.map(group => group.size)) + headingWidth + titleWidth), height / (Math.max(0, ...rowGroups.map(group => group.size)) + headingHeight + titleHeight));
  let rowBands, columnBands;
  for (let attempt = 0; attempt < 100; attempt++) {
    rowBands = rowGroups.length ? splitBands(rowGroups, height / scale - headingHeight - titleHeight, rowBreaks) : [[]];
    columnBands = columnGroups.length ? splitBands(columnGroups, width / scale - headingWidth - titleWidth, columnBreaks) : [[]];
    if (!settings.fitToPage || (!settings.fitToWidth || columnBands.length <= settings.fitToWidth) && (!settings.fitToHeight || rowBands.length <= settings.fitToHeight)) break;
    scale *= .95;
  }
  if (rowBands.length * columnBands.length > 500) throw new RangeError('인쇄 페이지가 500장을 초과합니다. 더 작은 범위를 선택하세요.');
  const covered = new Map(), anchors = new Map();
  for (const merged of visibleMerges) {
    const visibleRows = rows.filter(row => row >= merged.top && row <= merged.bottom), visibleColumns = columns.filter(column => column >= merged.left && column <= merged.right);
    if (!visibleRows.length || !visibleColumns.length) continue;
    if (visibleRows.some(row => titleRowSet.has(row)) && visibleRows.some(row => !titleRowSet.has(row)) || visibleColumns.some(column => titleColumnSet.has(column)) && visibleColumns.some(column => !titleColumnSet.has(column))) throw new RangeError('반복 제목 경계에 걸친 병합 셀 전체를 제목 범위에 포함하세요.');
    const anchor = visibleRows[0] + ':' + visibleColumns[0];
    anchors.set(anchor, { rowSpan: visibleRows.length, columnSpan: visibleColumns.length, height: visibleRows.reduce((sum, row) => sum + heights[row], 0), width: visibleColumns.reduce((sum, column) => sum + widths[column], 0), cell: cells.get(merged.top + ':' + merged.left) });
    for (const row of visibleRows) for (const column of visibleColumns) if (row + ':' + column !== anchor) covered.set(row + ':' + column, true);
  }
  const pages = [];
  const pageBands = settings.pageOrder === 'overThenDown' ? rowBands.flatMap(bodyRows => columnBands.map(bodyColumns => ({ bodyRows, bodyColumns }))) : columnBands.flatMap(bodyColumns => rowBands.map(bodyRows => ({ bodyRows, bodyColumns })));
  for (const { bodyRows, bodyColumns } of pageBands) {
    const rowBand = [...titleRows, ...bodyRows], columnBand = [...titleColumns, ...bodyColumns];
    let table = '<table><colgroup>' + (settings.headings ? '<col style="width:40px">' : '') + columnBand.map(column => '<col style="width:' + widths[column] + 'px">').join('') + '</colgroup>';
    if (settings.headings) table += '<thead><tr style="height:24px"><th></th>' + columnBand.map(column => '<th>' + columnLabel(column) + '</th>').join('') + '</tr></thead>';
    table += '<tbody>';
    for (const row of rowBand) {
      table += '<tr data-print-row="' + row + '"' + (titleRowSet.has(row) ? ' data-repeat-row="true"' : '') + ' style="height:' + heights[row] + 'px">' + (settings.headings ? '<th>' + (row + 1) + '</th>' : '');
      for (const column of columnBand) {
        const key = row + ':' + column; if (covered.has(key)) continue;
        const merged = anchors.get(key), cell = merged ? merged.cell : cells.get(key), span = merged ? ' rowspan="' + merged.rowSpan + '" colspan="' + merged.columnSpan + '"' : '';
        const box = printCellBox(cell, merged ? merged.width : widths[column], merged ? merged.height : heights[row]);
        const text = String(cell?.displayValue ?? cell?.value ?? '');
        outputBudget.cells++; outputBudget.characters += text.length;
        if (outputBudget.cells > 50000 || outputBudget.characters > 2000000) throw new RangeError('반복 제목을 포함한 인쇄 출력이 너무 큽니다. 더 작은 범위나 제목을 선택하세요.');
        table += '<td' + span + (titleColumnSet.has(column) ? ' data-repeat-column="true"' : '') + ' style="' + printCellStyle(cell, settings.gridLines, box) + '"><div style="max-height:' + box.innerHeight + 'px;overflow:hidden">' + escape(text) + '</div></td>';
      }
      table += '</tr>';
    }
    table += '</tbody></table>';
    const naturalWidth = columnBand.reduce((sum, column) => sum + widths[column], headingWidth);
    const layers = charts.length ? chartPrintLayers(charts, printAxisRuns(columnBand, widths, columnPositions, headingWidth), printAxisRuns(rowBand, heights, rowPositions, headingHeight), outputBudget) : '';
    const naturalHeight = rowBand.reduce((sum, row) => sum + heights[row], headingHeight);
    pages.push('<section class="page" aria-label="' + (pages.length + 1) + '" data-print-scale="' + scale + '"><div class="sheet" style="width:' + naturalWidth + 'px;'+ (charts.length ? 'height:' + naturalHeight + 'px;' : '') + 'transform:scale(' + scale + ')">' + table + layers + '</div></section>');
  }
  const nonce = crypto.randomBytes(16).toString('base64'), korean = language === 'ko';
  const printLabel = korean ? '인쇄 / PDF 저장' : 'Print / Save PDF';
  const decorated = pages.map((page, index) => decoratePrintPage(page, index, pages.length, { headerFooter, context: headerContext, settings, headerBudget }));
  return '<!doctype html><html lang="' + (korean ? 'ko' : 'en') + '"><head><meta charset="UTF-8"><meta http-equiv="Content-Security-Policy" content="default-src \'none\'; style-src \'unsafe-inline\'; script-src \'nonce-' + nonce + '\'; base-uri \'none\'; form-action \'none\'"><title>' + escape(title) + '</title><style>'
    + '@page{size:' + paperWidth + 'in ' + paperHeight + 'in;margin:0}*{box-sizing:border-box}body{margin:0;font:13px Arial,sans-serif;background:#ddd;color:#202124}nav{padding:12px;background:white;position:sticky;top:0;z-index:1}nav span{margin-left:12px}.page{position:relative;width:' + paperWidth + 'in;height:' + paperHeight + 'in;padding:' + margin.top + 'in ' + margin.right + 'in ' + margin.bottom + 'in ' + margin.left + 'in;background:white;margin:12px auto;overflow:hidden;break-after:page;page-break-after:always}.page:last-child{break-after:auto;page-break-after:auto}.sheet{transform-origin:top left;position:relative;overflow:hidden}.print-chart-clip{position:absolute;overflow:hidden;pointer-events:none}.print-chart{position:absolute;background:white;color:#202124}.print-chart .chart-card{margin:0;width:100%;height:100%;display:flex;flex-direction:column;padding:4px;overflow:hidden;border:1px solid #ccc;background:white}.print-chart figcaption{font-size:12px;max-height:30%;overflow:hidden;text-align:center}.print-chart svg{width:100%;flex:1;min-height:0;font:10px Arial,sans-serif}.print-chart .chart-empty{font-size:12px}.print-header,.print-footer{position:absolute;display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:4px}.hf-content{position:absolute;left:0;width:100%}.print-header .hf-content{top:0}.print-footer .hf-content{bottom:0}.hf-cell{position:relative;white-space:pre-wrap;overflow-wrap:anywhere;line-height:1.2;min-width:0;overflow:hidden;height:100%;font-family:Arial,sans-serif;color:#202124}table{border-collapse:collapse;table-layout:fixed;width:100%}th{padding:0;overflow:hidden;white-space:nowrap;background:#eee;font-weight:normal;border:1px solid #999}td{height:inherit}th,td{print-color-adjust:exact;-webkit-print-color-adjust:exact}@media print{body{background:white}nav{display:none}.page{margin:0}}'
    + '</style></head><body><nav><button id="print" type="button">' + printLabel + '</button><span>' + escape(title) + '</span><span id="printPageCount">' + pages.length + (korean ? '쪽' : ' pages') + '</span><span>' + escape(note) + '</span><span id="printWarning" role="alert"></span></nav>' + decorated.join('') + '<script nonce="' + nonce + '">function fitHeaders(){let valid=true;for(const cell of document.querySelectorAll(".hf-cell")){let size=Number(cell.dataset.hfFontSize);const content=cell.querySelector(".hf-content"),fits=()=>content.scrollHeight<=cell.clientHeight+1&&content.scrollWidth<=cell.clientWidth+1;cell.style.fontSize=size+"pt";if(!fits()&&size>4){let low=4,high=size;cell.style.fontSize=low+"pt";if(fits()){for(let i=0;i<9;i++){const middle=(low+high)/2;cell.style.fontSize=middle+"pt";if(fits())low=middle;else high=middle}cell.style.fontSize=low+"pt"}}if(!fits())valid=false}document.getElementById("print").disabled=!valid;document.getElementById("printWarning").textContent=valid?"":' + JSON.stringify(korean ? '머리글·바닥글이 여백에 들어가지 않습니다. 문구를 줄이거나 페이지 여백을 늘리세요.' : 'The header/footer does not fit. Shorten its text or increase the page margins.') + ';return valid}document.fonts.ready.then(fitHeaders);window.addEventListener("beforeprint",fitHeaders);document.getElementById("print").addEventListener("click",()=>{if(fitHeaders())window.print()});</script></body></html>';
}

function combinePrintDocuments(documents, language = 'en', decoration = null) {
  if (!Array.isArray(documents) || !documents.length || documents.length > 16) throw new RangeError('인쇄 영역 개수가 유효하지 않습니다.');
  if (documents.length === 1) return documents[0];
  const pages = documents.flatMap(html => html.match(/<section class="page"[\s\S]*?<\/section>/g) || []);
  if (pages.length > 500) throw new RangeError('인쇄 페이지가 500장을 초과합니다. 더 작은 범위를 선택하세요.');
  const first = documents[0], start = first.indexOf('<section class="page"'), end = first.lastIndexOf('</section>') + '</section>'.length;
  if (start < 0 || end <= start || pages.length < documents.length) throw new RangeError('인쇄 영역 출력이 유효하지 않습니다.');
  const headerBudget = { characters: 0 };
  const output = first.slice(0, start) + pages.map((page, index) => {
    const renumbered = page.replace(/aria-label="\d+"/, 'aria-label="' + (index + 1) + '"');
    return decoration ? decoratePrintPage(renumbered, index, pages.length, { ...decoration, headerBudget }) : renumbered;
  }).join('') + first.slice(end);
  return output.replace(/<span id="printPageCount">[^<]*<\/span>/, '<span id="printPageCount">' + pages.length + (language === 'ko' ? '쪽' : ' pages') + '</span>');
}

module.exports = { DEFAULT_PRINT_SETTINGS, renderPrintDocument, combinePrintDocuments };
