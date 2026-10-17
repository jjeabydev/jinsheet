'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { DEFAULT_PRINT_SETTINGS, renderPrintDocument, combinePrintDocuments } = require('../src/printing');
const { chartPrintGeometry, printAxisRuns, drawingAxes, drawingPoint, drawingMarker } = require('../src/chart-print');

function projection(rows = 2, columns = 2) {
  const rowIndexes = Array.from({ length: rows }, (_, index) => index), columnIndexes = Array.from({ length: columns }, (_, index) => index);
  return { title: 'Print <sample>', rows: rowIndexes, columns: columnIndexes, widths: Object.fromEntries(columnIndexes.map(index => [index, 144])), heights: Object.fromEntries(rowIndexes.map(index => [index, 26])), cells: new Map(rowIndexes.flatMap(row => columnIndexes.map(column => [row + ':' + column, { displayValue: 'value-' + row + '-' + column, style: {} }]))), settings: { ...DEFAULT_PRINT_SETTINGS, margins: { ...DEFAULT_PRINT_SETTINGS.margins } } };
}
const pageCount = html => (html.match(/<section class="page"/g) || []).length;

test('차트 좌표는 숨김·행 높이·열 너비와 세 종류 앵커를 반영한다', () => {
  const geometry = chartPrintGeometry({ columnWidths: { 0: 10 }, rowHeights: { 0: 40 }, hiddenRows: [1], hiddenColumns: [1], defaultRowHeight: 26 }, [
    { printable: true, type: 'twoCellAnchor', from: { row: 0, column: 0, x: 3, y: 4 }, to: { row: 3, column: 3, x: 5, y: 6 } },
    { printable: true, type: 'oneCellAnchor', from: { row: 2, column: 2, x: 0, y: 0 }, width: 100, height: 60 },
    { printable: true, type: 'absoluteAnchor', x: 10, y: 20, width: 50, height: 30 }, { printable: false }
  ]);
  assert.equal(geometry.columns.start(2), 70); assert.equal(geometry.rows.start(2), 40);
  assert.equal(geometry.boxes[0].x, 3); assert.equal(geometry.boxes[0].height, 68);
  assert.equal(geometry.boxes[1].x, 70); assert.equal(geometry.boxes[1].y, 40);
  assert.equal(geometry.boxes[2].width, 50); assert.equal(geometry.boxes[3], null);
  assert.throws(() => chartPrintGeometry({}, [null]), /차트 위치/);
  assert.throws(() => chartPrintGeometry({}, [{ printable: true, type: 'absoluteAnchor', x: -1, y: 0, width: -1, height: 10 }]), /차트 위치/);
  const far = chartPrintGeometry({}, [{ printable: true, type: 'oneCellAnchor', from: { row: 1000000, column: 16000, x: 0, y: 0 }, width: 10, height: 10 }]);
  assert.equal(far.boxes[0].bottom, 1000000); assert.equal(far.boxes[0].right, 16000);
  const axes = drawingAxes({}), end = { x: axes.columns.start(16384), y: axes.rows.start(1048576) }, marker = drawingMarker(end, axes);
  assert.equal(marker.column, 16383); assert.equal(marker.row, 1048575); assert.deepEqual(drawingPoint(marker, axes), end);
  assert.throws(() => drawingMarker({ ...end, x: end.x + 1 }, axes), /시트 범위/); assert.throws(() => drawingMarker({ x: NaN, y: 0 }, axes), /시트 범위/);
});

test('차트는 페이지 경계·반복 제목·선택 영역에 맞춰 원본 위치에서 잘린다', () => {
  const input = projection(4, 4); input.pageBreaks = { rows: [2], columns: [2] };
  input.rowPositions = { 0: 0, 1: 26, 2: 52, 3: 78 }; input.columnPositions = { 0: 0, 1: 144, 2: 288, 3: 432 };
  input.charts = [{ x: 200, y: 30, width: 200, height: 50, html: '<figure><svg></svg></figure>' }];
  let html = renderPrintDocument(input); assert.equal(pageCount(html), 4); assert.equal((html.match(/data-print-chart=/g) || []).length, 4);
  assert.match(html, /left:-88px;top:-22px;width:200px;height:50px/);
  input.repeatRows = [0]; input.repeatColumns = [0]; html = renderPrintDocument(input);
  assert.match(html, /left:144px;top:26px;width:112px;height:28px/);
  input.rows = [2, 3]; input.columns = [2, 3]; input.repeatRows = []; input.repeatColumns = [];
  html = renderPrintDocument(input); assert.equal((html.match(/data-print-chart=/g) || []).length, 1);
  assert.match(html, /left:0px;top:0px;width:112px;height:28px/);
  assert.deepEqual(printAxisRuns([0, 3], { 0: 26, 3: 26 }, { 0: 0, 3: 78 }, 24), [{ source: 0, target: 24, length: 26 }, { source: 78, target: 50, length: 26 }]);
});

test('인쇄 HTML은 용지·여백·서식·표시 값을 적용하고 외부 콘텐츠를 실행하지 않는다', () => {
  const input = projection();
  input.cells.set('0:0', { displayValue: '<img src=x onerror=alert(1)>', style: { color: '#123456', fill: '#FFFFAA', fontFamily: 'Malgun Gothic', fontSize: 16, bold: true, italic: true, underline: true, wrapText: true, borders: { bottom: { line: 'double', color: '#ABCDEF' } } } });
  const html = renderPrintDocument(input);
  assert.equal(pageCount(html), 1); assert.match(html, /size:8\.2677in 11\.6929in/); assert.match(html, /padding:0\.75in 0\.7in 0\.75in 0\.7in/);
  assert.match(html, /Print &lt;sample&gt;/); assert.match(html, /&lt;img src=x onerror=alert\(1\)&gt;/); assert.doesNotMatch(html, /<img/);
  assert.match(html, /color:#123456/); assert.match(html, /background-color:#FFFFAA/); assert.match(html, /font-weight:700/); assert.match(html, /border-bottom:3px double #ABCDEF/);
  assert.match(html, /window\.print\(\)/); assert.match(html, /default-src 'none'/);
});

test('인쇄 페이지 분할은 모든 셀을 한 번씩 보존하고 행·열 머리글을 반복한다', () => {
  const input = projection(90, 8); input.settings.headings = true;
  const html = renderPrintDocument(input);
  assert.ok(pageCount(html) > 1); assert.equal((html.match(/<thead>/g) || []).length, pageCount(html));
  for (const cell of input.cells.values()) assert.equal(html.split('>' + cell.displayValue + '<').length - 1, 1, cell.displayValue);
});

test('페이지 수에 맞춤과 가로 방향·A3·Letter 설정을 인쇄 출력에 반영한다', () => {
  const input = projection(90, 8); input.settings.fitToPage = true; input.settings.fitToWidth = 1; input.settings.fitToHeight = 1;
  assert.equal(pageCount(renderPrintDocument(input)), 1);
  input.settings.orientation = 'landscape'; input.settings.paperSize = 8;
  assert.match(renderPrintDocument(input), /size:16\.5354in 11\.6929in/);
  input.settings.paperSize = 1; assert.match(renderPrintDocument(input), /size:11in 8\.5in/);
});

test('병합 셀은 숨김 행이 있어도 한 번 출력하고 페이지 경계에서 분리하지 않는다', () => {
  const input = projection(60, 2);
  input.mergedCells = [{ top: 0, bottom: 49, left: 0, right: 1 }];
  input.rows = input.rows.filter(row => row !== 1);
  const html = renderPrintDocument(input);
  assert.match(html, /rowspan="49" colspan="2"/); assert.equal(html.split('>value-0-0<').length - 1, 1); assert.doesNotMatch(html, />value-3-1</);
  assert.match(html, /max-height:1266px/);
});

test('인쇄는 빈 영역·페이지 폭주와 잘못된 설정을 거부한다', () => {
  const empty = projection(); empty.rows = []; assert.throws(() => renderPrintDocument(empty));
  const large = projection(1000, 20); large.settings.scale = 400; for (const row of large.rows) large.heights[row] = 500; for (const column of large.columns) large.widths[column] = 600;
  assert.throws(() => renderPrintDocument(large), /500/);
  const invalid = projection(); invalid.settings.scale = NaN; assert.throws(() => renderPrintDocument(invalid));
});

test('여러 인쇄 영역은 별도 페이지와 하나의 인쇄 버튼으로 결합된다', () => {
  const first = projection(), second = projection(); second.cells.set('0:0', { displayValue: 'Second print area' });
  const html = combinePrintDocuments([renderPrintDocument(first), renderPrintDocument(second)], 'ko');
  assert.equal(pageCount(html), 2); assert.equal((html.match(/id="print"/g) || []).length, 1);
  assert.match(html, /id="printPageCount">2쪽</); assert.match(html, /aria-label="2"/); assert.match(html, /Second print area/);
});

test('인쇄 제목 행·열은 페이지마다 반복하고 원래 본문에서 중복하지 않는다', () => {
  const input = projection(90, 8); input.repeatRows = [0]; input.repeatColumns = [0];
  const html = renderPrintDocument(input), pages = html.match(/<section class="page"[\s\S]*?<\/section>/g);
  assert.ok(pages.length > 1);
  for (const page of pages) {
    assert.equal((page.match(/data-print-row="0"/g) || []).length, 1);
    assert.match(page, /data-repeat-column="true"/); assert.equal(page.split('>value-0-0<').length - 1, 1);
  }
  for (const row of input.rows.filter(row => row !== 0)) for (const column of input.columns.filter(column => column !== 0)) assert.equal(html.split('>value-' + row + '-' + column + '<').length - 1, 1, 'each non-title body value remains present once');
});

test('반복 제목만 있는 범위와 페이지 수 맞춤도 정상 출력한다', () => {
  const input = projection(); input.repeatRows = [0, 1]; input.repeatColumns = [0, 1];
  assert.equal(pageCount(renderPrintDocument(input)), 1);
  const large = projection(90, 8); large.repeatRows = [0]; large.repeatColumns = [0]; large.settings.fitToPage = true; large.settings.fitToWidth = 1; large.settings.fitToHeight = 1;
  assert.equal(pageCount(renderPrintDocument(large)), 1);
});

test('반복 제목의 병합 셀 경계와 출력량 제한을 검사한다', () => {
  const merged = projection(3, 3); merged.mergedCells = [{ top: 0, bottom: 1, left: 0, right: 0 }]; merged.repeatRows = [0];
  assert.throws(() => renderPrintDocument(merged), /병합/);
  merged.repeatRows = [0, 1]; assert.match(renderPrintDocument(merged), /rowspan="2"/);
  const input = projection(); input.outputBudget = { cells: 49999, characters: 0 };
  assert.throws(() => renderPrintDocument(input), /출력/);
  input.outputBudget = { cells: 0, characters: 2000000 }; assert.throws(() => renderPrintDocument(input), /출력/);
});

test('머리글·바닥글은 시작 번호와 첫 쪽·홀짝 쪽 및 전체 쪽 수를 반영한다', () => {
  const { defaultHeaderFooter } = require('../src/header-footer'), input = projection(90, 8), config = defaultHeaderFooter();
  config.differentFirst = true; config.differentOddEven = true;
  config.first.header.center = 'First &P / &N'; config.odd.header.center = 'Odd &P / &N &F &A'; config.even.header.center = 'Even &P / &N'; config.odd.footer.right = 'Footer';
  input.headerFooter = config; input.headerContext = { fileName: 'report.xlsx', sheetName: 'Data' }; input.settings.firstPageNumber = 10; input.settings.useFirstPageNumber = true;
  const html = renderPrintDocument(input), pages = html.match(/<section class="page"[\s\S]*?<\/section>/g), text = page => page.match(/<!--print-hf-start-->([\s\S]*?)<!--print-hf-end-->/)[1].replace(/<[^>]*>/g, '');
  assert.ok(pages.length >= 3); assert.equal(text(pages[0]), 'First 10 / ' + pages.length);
  assert.match(text(pages[1]), /^Odd 11 \/ \d+ report.xlsx DataFooter$/); assert.equal(text(pages[2]), 'Even 12 / ' + pages.length);
});

test('여러 인쇄 영역의 쪽 수와 첫 쪽·홀짝 머리글은 결합 후 다시 계산한다', () => {
  const { defaultHeaderFooter } = require('../src/header-footer'), config = defaultHeaderFooter(), input = projection();
  config.differentFirst = true; config.differentOddEven = true; config.first.header.center = 'First &P/&N'; config.odd.header.center = 'Odd &P/&N'; config.even.header.center = 'Even &P/&N'; input.headerFooter = config;
  const first = renderPrintDocument(input), second = renderPrintDocument(input), html = combinePrintDocuments([first, second], 'en', { headerFooter: config, settings: input.settings });
  const pages = html.match(/<section class="page"[\s\S]*?<\/section>/g), text = page => page.match(/<!--print-hf-start-->([\s\S]*?)<!--print-hf-end-->/)[1].replace(/<[^>]*>/g, '');
  assert.equal(text(pages[0]), 'First 1/2'); assert.equal(text(pages[1]), 'Even 2/2'); assert.doesNotMatch(pages[1], />First /);
});

test('머리글 배율·여백 정렬과 본문 겹침 및 빈 서식만 있는 구역을 처리한다', () => {
  const { defaultHeaderFooter } = require('../src/header-footer'), input = projection(); input.headerFooter = defaultHeaderFooter(); input.headerFooter.odd.header.left = 'Header'; input.settings.scale = 50;
  assert.match(renderPrintDocument(input), /font-size:5\.5pt/);
  input.headerFooter.scaleWithDoc = false; input.headerFooter.alignWithMargins = false;
  assert.match(renderPrintDocument(input), /font-size:11pt/); assert.match(renderPrintDocument(input), /left:0in;right:0in/);
  input.settings.margins.header = input.settings.margins.top; assert.throws(() => renderPrintDocument(input), /여백/);
  input.headerFooter.odd.header.left = '&B'; assert.doesNotMatch(renderPrintDocument(input), /class="print-header"/);
});

test('수동 행·열 나누기는 경계에 맞게 네 쪽을 만들고 인쇄 순서를 따른다', () => {
  const input = projection(); input.pageBreaks = { rows: [1], columns: [1] };
  const values = html => (html.match(/<section class="page"[\s\S]*?<\/section>/g) || []).map(page => />(value-\d+-\d+)</.exec(page)[1]);
  const down = renderPrintDocument(input); assert.equal(pageCount(down), 4); assert.deepEqual(values(down), ['value-0-0', 'value-1-0', 'value-0-1', 'value-1-1']);
  input.settings.pageOrder = 'overThenDown'; assert.deepEqual(values(renderPrintDocument(input)), ['value-0-0', 'value-0-1', 'value-1-0', 'value-1-1']);
  input.settings.fitToPage = true; input.settings.fitToWidth = 1; input.settings.fitToHeight = 1;
  assert.equal(pageCount(renderPrintDocument(input)), 1, 'fit-to-pages ignores stored manual breaks');
});

test('숨김 행 사이의 나누기는 빈 쪽을 만들지 않고 반복 제목을 유지한다', () => {
  const input = projection(6, 2); input.rows = [0, 1, 4, 5]; input.repeatRows = [0]; input.pageBreaks = { rows: [2, 3, 4, 99], columns: [] };
  const html = renderPrintDocument(input), pages = html.match(/<section class="page"[\s\S]*?<\/section>/g);
  assert.equal(pages.length, 2); for (const page of pages) assert.equal(page.split('>value-0-0<').length - 1, 1);
  assert.match(pages[0], />value-1-0</); assert.doesNotMatch(pages[0], />value-4-0</); assert.match(pages[1], />value-4-0</); assert.match(pages[1], />value-5-0</);
});

test('수동 나누기는 출력되는 병합 셀만 보호하고 맞춤에서는 무시한다', () => {
  const input = projection(4, 2); input.mergedCells = [{ top: 0, bottom: 2, left: 0, right: 0 }]; input.pageBreaks = { rows: [1], columns: [] };
  assert.throws(() => renderPrintDocument(input), /병합/);
  input.settings.fitToPage = true; assert.ok(renderPrintDocument(input));
  input.settings.fitToPage = false; input.mergedCells = [{ top: 0, bottom: 2, left: 10, right: 11 }];
  assert.equal(pageCount(renderPrintDocument(input)), 2, 'an unprinted merge does not group or block printed rows');
});
