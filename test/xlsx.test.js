'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { unzipSync, zipSync } = require('../vendor/node_modules/fflate');
const { SaxesParser } = require('../vendor/node_modules/saxes');
const { analyze, formatCellValue, inspectZip, patchCell, patchCells, patchColumnWidths, patchRowHeights, patchRowHeightValues, patchPrintSettings, patchPrintArea, patchDefinedName, patchPrintTitles, patchHeaderFooter, patchPageBreaks, patchAutoFilters, patchFilterVisibility, patchHiddenRows, patchHiddenColumns, patchOutlineGroups, patchFreezePanes, patchMergedCells, patchAddTable, patchTableStyles, patchTableRanges, patchTableTotals, patchAddChart, patchChartMetadata, patchChartPlacement, patchChartSource, patchDeleteChart, patchClearConditionalFormats, patchDeleteConditionalFormatRule, patchMoveConditionalFormatRule, patchEditConditionalFormatRule, patchIconSetConditionalFormat, patchCellValueConditionalFormat, patchDuplicateValuesConditionalFormat, patchSheetName, patchAddSheet, patchDeleteSheet, patchMoveSheet, translateFormulaReferences, transformStructure } = require('../src/xlsx');
const { createFormulaCalculator, evaluateFormula } = require('../src/formula');
const { chartPrintGeometry } = require('../src/chart-print');
const { ICON_SET_LENGTHS } = require('../src/auto-filter');

test('숫자 셀 비교 조건부 서식은 범위·차등 채우기를 생성하고 저장 후 다시 계산한다', () => {
  const original = analyze(fixture());
  assert.equal(original.writable, true, original.readOnlyReason);
  const changed = patchCellValueConditionalFormat(original, 0, { top: 0, bottom: 0, left: 0, right: 2 }, 'greaterThanOrEqual', [40]);
  assert.equal(changed.model.writable, true, changed.model.readOnlyReason);
  const rule = changed.model.sheets[0].conditionalFormats[0];
  assert.equal(rule.type, 'cellIs');
  assert.equal(rule.operator, 'greaterThanOrEqual');
  assert.deepEqual(rule.values, [40]);
  assert.deepEqual(rule.ranges, [{ top: 0, bottom: 0, left: 0, right: 2 }]);
  assert.deepEqual(rule.style, { fill: '#F4CCCC' });
  const reopened = analyze(changed.bytes);
  assert.equal(reopened.writable, true, reopened.readOnlyReason);
  const styleAt = address => {
    const cell = reopened.sheets[0].cells.get(address) || null;
    const match = /^([A-Z]+)(\d+)$/.exec(address);
    let column = 0;
    for (const letter of match[1]) column = column * 26 + letter.charCodeAt(0) - 64;
    return require('../src/xlsx').conditionalStyleForCell(reopened.sheets[0], Number(match[2]) - 1, column - 1, cell, (row, col) => reopened.sheets[0].cells.get(String.fromCharCode(65 + col) + (row + 1))?.value ?? null, evaluateFormula);
  };
  assert.deepEqual(styleAt('B1'), { fill: '#F4CCCC' });
  assert.equal(styleAt('A1'), null, 'text is excluded from numeric rules');
  assert.equal(styleAt('C1'), null, 'booleans are excluded from numeric rules');
  const zeroRule = patchCellValueConditionalFormat(original, 0, { top: 0, bottom: 0, left: 3, right: 3 }, 'equal', [0]);
  assert.equal(require('../src/xlsx').conditionalStyleForCell(zeroRule.model.sheets[0], 0, 3, null, () => null, evaluateFormula), null, 'blank cells are excluded even when zero matches the condition');
  const less = patchCellValueConditionalFormat(original, 0, { top: 0, bottom: 0, left: 1, right: 1 }, 'between', [50, 30]);
  const between = less.model.sheets[0].conditionalFormats[0];
  assert.deepEqual(between.values, [50, 30]);
  assert.deepEqual(require('../src/xlsx').conditionalStyleForCell(less.model.sheets[0], 0, 1, less.model.sheets[0].cells.get('B1'), () => 42, evaluateFormula), { fill: '#F4CCCC' }, 'between uses an inclusive normalized interval');
  for (const [operator, values, color] of [['unknown', [1]], ['between', [1]], ['equal', [Infinity]], ['equal', [1], 'red']]) assert.throws(() => patchCellValueConditionalFormat(original, 0, { top: 0, bottom: 0, left: 0, right: 0 }, operator, values, color), /숫자 조건부 서식 설정이 유효하지 않습니다/);
});

test('수식 안의 Excel 오류 리터럴은 값으로 전파되고 선택되지 않은 분기에는 전파되지 않는다', () => {
  for (const error of ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#N/A!', '#CALC!', '#CIRC!', '#SPILL!', '#GETTING_DATA']) {
    assert.equal(evaluateFormula(error), error);
    assert.equal(evaluateFormula(error + '+2'), error);
  }
  assert.equal(evaluateFormula('IF(FALSE,#REF!,1)'), 1);
  assert.equal(evaluateFormula('IF(TRUE,#REF!,1)'), '#REF!');
  assert.equal(evaluateFormula('TRUE()'), true, 'TRUE accepts Excel function-call syntax');
  assert.equal(evaluateFormula('FALSE()'), false, 'FALSE accepts Excel function-call syntax');
  assert.equal(evaluateFormula('TRUE(1)'), '#VALUE!', 'TRUE does not accept arguments');
});

test('기본 수식 계산은 연산자 우선순위·셀 참조·범위 집계와 오류를 처리한다', () => {
  const cells = new Map([["0:0", 2], ["0:1", 3], ["1:0", 4], ["1:1", 5]]);
  const context = { getCell: (row, column) => cells.get(row + ':' + column) ?? null };
  assert.equal(evaluateFormula('1+2*3'), 7);
  assert.equal(evaluateFormula('(1+2)^3'), 27);
  assert.equal(evaluateFormula('0^(-1)'), '#DIV/0!', 'a negative exponent on zero reports division by zero');
  assert.equal(evaluateFormula('POWER(0,-1)'), '#DIV/0!', 'POWER reports division by zero for a negative exponent on zero');
  assert.equal(evaluateFormula('(-1)^0.5'), '#NUM!', 'a fractional power of a negative base reports a numeric error');
  assert.equal(evaluateFormula('-2^2'), -4);
  assert.equal(evaluateFormula('50%+10%'), 0.6);
  assert.equal(evaluateFormula('A1+$B$1', context), 5);
  assert.equal(evaluateFormula('Z9', context), 0);
  assert.equal(evaluateFormula('SUM(A1:B2)', context), 14);
  assert.equal(evaluateFormula('TEXT(1234.5,"#,##0.00")'), '1,234.50');
  assert.equal(evaluateFormula('TEXT(0.25,"0.0%")'), '25.0%');
  assert.equal(evaluateFormula('TEXT(1234,"$#,##0")'), '$1,234');
  assert.equal(evaluateFormula('TEXT(45292,"yyyy-mm-dd")'), '2024-01-01');
  assert.equal(evaluateFormula('TEXT(45292,"m/d/yy")'), '1/1/24');
  assert.equal(evaluateFormula('TEXT(45292,"dd mmm yyyy")'), '01 Jan 2024');
  assert.equal(evaluateFormula('TEXT(0.5,"h:mm AM/PM")'), '12:00 PM');
  assert.equal(evaluateFormula('TEXT(45292.5,"yyyy-mm-dd h:mm")'), '2024-01-01 12:00');
  assert.equal(evaluateFormula('TEXT(0.999999,"hh:mm:ss")'), '00:00:00');
  assert.equal(evaluateFormula('TEXT(2,"unsupported")'), '#VALUE!');
  assert.equal(evaluateFormula('VALUE("$1,234.50")'), 1234.5);
  assert.equal(evaluateFormula('VALUE("25%")'), 0.25);
  assert.equal(evaluateFormula('VALUE("1,23")'), '#VALUE!');
  assert.equal(evaluateFormula('VALUE(42)'), 42);
  assert.equal(evaluateFormula('NUMBERVALUE("1,234.5")'), 1234.5);
  assert.equal(evaluateFormula('NUMBERVALUE("2.500,27",",",".")'), 2500.27);
  assert.equal(evaluateFormula('NUMBERVALUE("1 234,5 %",","," ")'), 12.345);
  assert.equal(evaluateFormula('NUMBERVALUE("12,5%%",",",".")'), 0.00125);
  assert.equal(evaluateFormula('NUMBERVALUE("-1.25e2")'), -125);
  assert.equal(evaluateFormula('NUMBERVALUE("1,2",",",",")'), '#VALUE!');
  assert.equal(evaluateFormula('NUMBERVALUE("12.3.4")'), '#VALUE!');
  assert.equal(evaluateFormula('NUMBERVALUE("")'), '#VALUE!');
  assert.equal(evaluateFormula('SUMSQ(A1:B2)', context), 54, 'SUMSQ sums the squares of numeric range entries');
  assert.equal(evaluateFormula('SUMXMY2(A1:B2,A1:B2)', context), 0, 'SUMXMY2 sums squared pairwise differences');
  assert.equal(evaluateFormula('SUMX2MY2(A1:B2,A1:B2)', context), 0, 'SUMX2MY2 subtracts paired sums of squares');
  assert.equal(evaluateFormula('SUMX2PY2(A1:B2,A1:B2)', context), 108, 'SUMX2PY2 adds paired sums of squares');
  const pairwise = { getCell: (row, column) => ({ '0:0': 1, '1:0': 2, '2:0': 3, '0:1': 2, '1:1': 4, '2:1': 5 })[row + ':' + column] ?? null };
  assert.equal(evaluateFormula('SUMXMY2(A1:A3,B1:B3)', pairwise), 9);
  assert.equal(evaluateFormula('SUMX2MY2(A1:A3,B1:B3)', pairwise), -31);
  assert.equal(evaluateFormula('SUMX2PY2(A1:A3,B1:B3)', pairwise), 59);
  assert.equal(evaluateFormula('SUMXMY2(A1:A2,B1:B3)', pairwise), '#N/A', 'paired array formulas require matching dimensions');
  assert.equal(evaluateFormula('DEVSQ(A1:B2)', context), 5, 'DEVSQ sums squared deviations from the mean');
  assert.equal(evaluateFormula('AVEDEV(A1:B2)', context), 1, 'AVEDEV averages absolute deviations');
  assert.equal(evaluateFormula('DEVSQ(1E308,1E308)'), 0, 'equal large values have zero squared deviation without overflowing the mean');
  assert.equal(evaluateFormula('AVEDEV(1E308,1E308)'), 0, 'equal large values have zero absolute deviation without overflowing the mean');
  assert.equal(evaluateFormula('AVEDEV(-1E308,1E308)'), 1E308, 'absolute deviation remains finite for large opposite-signed values');
  assert.equal(evaluateFormula('DEVSQ(-1E308,1E308)'), '#NUM!', 'squared deviation reports a result beyond numeric range');
  assert.equal(evaluateFormula('GEOMEAN(A1:B2)', context), 120 ** 0.25, 'GEOMEAN supports numeric ranges');
  assert.equal(evaluateFormula('HARMEAN(A1:B2)', context), 4 / (1 / 2 + 1 / 3 + 1 / 4 + 1 / 5), 'HARMEAN supports numeric ranges');
  const regression = { getCell: (row, column) => ({ '0:0': 1, '1:0': 2, '2:0': 3, '0:1': 2, '1:1': 4, '2:1': 5 })[row + ':' + column] ?? null };
  assert.ok(Math.abs(evaluateFormula('CORREL(A1:A3,B1:B3)', regression) - 3 / Math.sqrt(2 * (14 / 3))) < 1e-12, 'CORREL aligns numeric cells by position');
const largeLinearRegression = { getCell: (row, column) => column === 0 ? 1E150 * (row + 1) : 1E150 * (2 * row + 1) };
assert.equal(evaluateFormula('CORREL(A1:A3,B1:B3)', largeLinearRegression), 1, 'correlation avoids overflowing the product of finite squared deviations');
assert.equal(evaluateFormula('RSQ(B1:B3,A1:A3)', largeLinearRegression), 1, 'RSQ preserves a perfect fit for large finite data');
const overflowVarianceRegression = { getCell: (row, column) => column === 0 ? 1E155 * (row + 1) : 1E155 * (2 * row + 1) };
assert.equal(evaluateFormula('CORREL(A1:A3,B1:B3)', overflowVarianceRegression), 1, 'correlation rescales data when squared deviations overflow');
assert.ok(Math.abs(evaluateFormula('SLOPE(B1:B3,A1:A3)', overflowVarianceRegression) - 2) < 1e-12);
assert.ok(Math.abs(evaluateFormula('INTERCEPT(B1:B3,A1:A3) / 1E155', overflowVarianceRegression) + 1) < 1e-12);
assert.equal(evaluateFormula('RSQ(B1:B3,A1:A3)', overflowVarianceRegression), 1);
assert.ok(evaluateFormula('STEYX(B1:B3,A1:A3)', overflowVarianceRegression) < 1E141);
assert.equal(evaluateFormula('COVARIANCE.P(A1:A3,B1:B3)', overflowVarianceRegression), '#NUM!', 'covariance reports a result beyond numeric range');
assert.equal(evaluateFormula('FORECAST.LINEAR(4E155,B1:B3,A1:A3)', overflowVarianceRegression), 7E155);
  assert.equal(evaluateFormula('PEARSON(A1:A3,B1:B3)', regression), evaluateFormula('CORREL(A1:A3,B1:B3)', regression));
  assert.equal(evaluateFormula('COVARIANCE.P(A1:A3,B1:B3)', regression), 1);
  assert.equal(evaluateFormula('COVARIANCE.S(A1:A3,B1:B3)', regression), 1.5);
  assert.equal(evaluateFormula('SLOPE(B1:B3,A1:A3)', regression), 1.5);
  assert.ok(Math.abs(evaluateFormula('INTERCEPT(B1:B3,A1:A3)', regression) - 2 / 3) < 1e-12);
  assert.ok(Math.abs(evaluateFormula('RSQ(B1:B3,A1:A3)', regression) - 9 / (2 * (14 / 3))) < 1e-12);
  assert.ok(Math.abs(evaluateFormula('STEYX(B1:B3,A1:A3)', regression) - Math.sqrt(1 / 6)) < 1e-12);
  assert.ok(Math.abs(evaluateFormula('FORECAST.LINEAR(4,B1:B3,A1:A3)', regression) - 20 / 3) < 1e-12);
const largeConstantPairs = { getCell: (row, column) => column === 0 ? 1E308 : 1E308 };
assert.equal(evaluateFormula('CORREL(A1:A2,B1:B2)', largeConstantPairs), '#DIV/0!');
assert.equal(evaluateFormula('COVARIANCE.P(A1:A2,B1:B2)', largeConstantPairs), 0);
const largeConstantY = { getCell: (row, column) => column === 0 ? row + 1 : 1E308 };
assert.equal(evaluateFormula('SLOPE(B1:B3,A1:A3)', largeConstantY), 0);
assert.equal(evaluateFormula('INTERCEPT(B1:B3,A1:A3)', largeConstantY), 1E308);
assert.equal(evaluateFormula('STEYX(B1:B3,A1:A3)', largeConstantY), 0);
assert.equal(evaluateFormula('FORECAST.LINEAR(4,B1:B3,A1:A3)', largeConstantY), 1E308);
  assert.equal(evaluateFormula('SLOPE(A1:A2,A1:A3)', regression), '#N/A', 'paired regression ranges must have matching dimensions');
  assert.equal(evaluateFormula('CORREL(A1:A2,A1:A2)', { getCell: () => 2 }), '#DIV/0!', 'correlation rejects a constant input series');
  assert.equal(evaluateFormula('GEOMEAN(-1,2)'), '#NUM!');
  assert.equal(evaluateFormula('HARMEAN(0,2)'), '#NUM!');
  assert.equal(evaluateFormula('DEVSQ()'), '#NUM!');
  assert.equal(evaluateFormula('SUMSQ()'), 0);
  const payment = evaluateFormula('PMT(0.01,12,1000)');
  assert.ok(payment < 0, 'PMT follows Excel cash-flow sign conventions');
  assert.ok(Math.abs(evaluateFormula(`RATE(12,${payment},1000)`) - 0.01) < 1e-9, 'RATE recovers the interest rate from a payment schedule');
  assert.ok(Math.abs(evaluateFormula(`PV(0.01,12,${payment})`) - 1000) < 1e-9, 'PV reverses a PMT calculation');
  assert.ok(Math.abs(evaluateFormula(`FV(0.01,12,${payment},1000)`)) < 1e-9, 'FV reaches zero when funded by the matching present value');
  assert.ok(Math.abs(evaluateFormula(`NPER(0.01,${payment},1000)`) - 12) < 1e-9, 'NPER reverses the payment schedule');
  const duePayment = evaluateFormula('PMT(0.01,12,1000,0,1)');
  assert.ok(Math.abs(evaluateFormula(`PV(0.01,12,${duePayment},0,1)`) - 1000) < 1e-9, 'type 1 models payments at the beginning of each period');
assert.ok(Math.abs(evaluateFormula(`RATE(12,${duePayment},1000,0,1)`) - 0.01) < 1e-9, 'RATE supports beginning-of-period payments');
const cashFlowContext = { getCell: row => [-100, 39, 59, 55, 20][row] ?? null };
assert.ok(Math.abs(evaluateFormula('NPV(0.1,A1:A5)', cashFlowContext) - 35.634053809041816) < 1e-12, 'NPV discounts the first supplied value at period one');
assert.equal(evaluateFormula('NPV(0,-100,20)'), -80);
assert.ok(Math.abs(evaluateFormula('NPV(0.1,A1:A3)', { getCell: row => [-100, null, 200][row] ?? null }) - (-100 / 1.1 + 200 / 1.1 ** 3)) < 1e-12, 'blank cash-flow cells keep their original period positions');
assert.equal(evaluateFormula('NPV(-1,A1)', cashFlowContext), '#NUM!');
assert.ok(Math.abs(evaluateFormula('IRR(A1:A5)', cashFlowContext) - 0.2809484211599611) < 1e-10);
assert.ok(Math.abs(evaluateFormula('IRR(A1:A5,0.2)', cashFlowContext) - 0.2809484211599611) < 1e-10, 'IRR accepts an initial guess');
assert.equal(evaluateFormula('IRR(A1:A2)', { getCell: row => [-1, -2][row] }), '#NUM!');
assert.ok(Math.abs(evaluateFormula('IRR(A1:A3)', { getCell: row => [-100, null, 121][row] ?? null }) - 0.1) < 1e-10, 'blank cash-flow cells retain their periods');
assert.ok(Math.abs(evaluateFormula('MIRR(A1:A4,0.1,0.12)', { getCell: row => [-100, 30, 40, 50][row] }) - 0.09815669244631553) < 1e-12);
assert.ok(Math.abs(evaluateFormula('MIRR(A1:A4,0,0)', { getCell: row => [-100, 30, 40, 50][row] }) - (1.2 ** (1 / 3) - 1)) < 1e-12);
assert.equal(evaluateFormula('MIRR(A1:A2,0.1,0.1)', { getCell: row => [1, 2][row] }), '#DIV/0!');
assert.equal(evaluateFormula('MIRR(A1:A2,-1,0.1)', { getCell: row => [-1, 2][row] }), '#NUM!');
const datedCashFlows = { getCell: (row, column) => ({ '0:0': -100, '1:0': 110, '0:1': 1, '1:1': 366 })[row + ':' + column] ?? null };
assert.ok(Math.abs(evaluateFormula('XNPV(0.1,A1:A2,B1:B2)', datedCashFlows)) < 1e-12);
assert.ok(Math.abs(evaluateFormula('XIRR(A1:A2,B1:B2)', datedCashFlows) - 0.1) < 1e-10);
assert.ok(Math.abs(evaluateFormula('XIRR(A1:A2,B1:B2,0.2)', datedCashFlows) - 0.1) < 1e-10, 'XIRR accepts an initial guess');
assert.equal(evaluateFormula('XNPV(-1,A1:A2,B1:B2)', datedCashFlows), '#NUM!');
assert.equal(evaluateFormula('XIRR(A1:A2,B1:B2)', { getCell: (row, column) => column === 0 ? [-100, -20][row] : [1, 366][row] }), '#NUM!');
assert.equal(evaluateFormula('XIRR(A1:A2,B1:B1)', datedCashFlows), '#NUM!');
assert.equal(evaluateFormula('XNPV(0.1,A1:A2,B1:B2)', { getCell: (row, column) => column === 0 ? [-100, 110][row] : ['date', 366][row] }), '#NUM!');
assert.equal(evaluateFormula('IRR(A1:A2,-1)', { getCell: row => [-1, 2][row] }), '#NUM!');
assert.equal(evaluateFormula('IRR(A1:A2)', { getCell: row => ['#N/A', 2][row] }), '#N/A');
assert.equal(evaluateFormula('PMT(0,10,1000)'), -100);
  assert.equal(evaluateFormula('PV(0,10,-100)'), 1000);
  assert.equal(evaluateFormula('FV(0,10,-100,1000)'), 0);
  assert.equal(evaluateFormula('NPER(0,-100,1000)'), 10);
  assert.equal(evaluateFormula('PMT(0.1,0,1000)'), '#NUM!');
  assert.equal(evaluateFormula('FV(0.1,10,-100,0,2)'), '#NUM!');
  assert.equal(evaluateFormula('RATE(0,-100,1000)'), '#NUM!');
  assert.equal(evaluateFormula('RATE(12,-100,1000,0,2)'), '#NUM!');
assert.equal(evaluateFormula('RATE(12,0,1000)'), '#NUM!');
assert.equal(evaluateFormula('MAXIFS(B1:B2,A1:A2,">2")', context), 5);
assert.ok(Math.abs(evaluateFormula('EFFECT(0.0525,4)') - 0.0535427) < 1e-7);
assert.ok(Math.abs(evaluateFormula('NOMINAL(EFFECT(0.0525,4),4)') - 0.0525) < 1e-12);
assert.equal(evaluateFormula('EFFECT(0,12)'), 0);
assert.equal(evaluateFormula('EFFECT(0.05,0)'), '#NUM!');
assert.equal(evaluateFormula('NOMINAL(0.05,1.5)'), '#NUM!');
const monthlyPayment = evaluateFormula('PMT(0.1/12,12,10000)');
assert.ok(Math.abs(evaluateFormula('CUMIPMT(0.1/12,12,10000,1,1,0)') + 10000 * 0.1 / 12) < 1e-10);
assert.ok(Math.abs(evaluateFormula('CUMPRINC(0.1/12,12,10000,1,1,0)') - (monthlyPayment + 10000 * 0.1 / 12)) < 1e-10);
const dueMonthlyPayment = evaluateFormula('PMT(0.1/12,12,10000,0,1)');
assert.ok(Math.abs(evaluateFormula('CUMIPMT(0.1/12,12,10000,1,1,1)') + (10000 + dueMonthlyPayment) * 0.1 / 12) < 1e-10);
assert.equal(evaluateFormula('CUMPRINC(0.1/12,12,10000,1,1,1)'), dueMonthlyPayment);
assert.ok(Math.abs(evaluateFormula('CUMIPMT(0.1/12,12,10000,1,12,0)') + evaluateFormula('CUMPRINC(0.1/12,12,10000,1,12,0)') - monthlyPayment * 12) < 1e-8);
assert.equal(evaluateFormula('CUMIPMT(0.1/12,12,10000,0,1,0)'), '#NUM!');
assert.equal(evaluateFormula('CUMPRINC(0.1/12,12,10000,1,13,0)'), '#NUM!');
assert.equal(evaluateFormula('SLN(30000,7500,10)'), 2250);
assert.equal(evaluateFormula('SLN(30000,7500,0)'), '#NUM!');
assert.ok(Math.abs(evaluateFormula('SYD(30000,7500,10,1)') - 4090.909090909091) < 1e-10);
assert.ok(Math.abs(evaluateFormula('SYD(30000,7500,10,10)') - 409.09090909090907) < 1e-10);
assert.equal(evaluateFormula('SYD(30000,7500,10,11)'), '#NUM!');
assert.equal(evaluateFormula('DDB(2400,300,10,1)'), 480);
assert.equal(evaluateFormula('DDB(2400,300,10,2)'), 384);
assert.ok(Math.abs(evaluateFormula('DDB(2400,300,10,10)') - (2400 * 0.8 ** 9 - 300)) < 1e-10, 'DDB stops at salvage value');
assert.equal(evaluateFormula('DDB(2400,300,10,1,0)'), '#NUM!');
assert.equal(evaluateFormula('DDB(2400,300,10,11)'), '#NUM!');
assert.equal(evaluateFormula('VDB(2400,300,10,0,1)'), 480);
assert.equal(evaluateFormula('VDB(2400,300,10,0.5,1,2,TRUE)'), 240, 'VDB prorates a partial period');
assert.ok(Math.abs(evaluateFormula('VDB(2400,300,10,1.5,2.5,2,TRUE)') - 345.6) < 1e-10);
assert.ok(Math.abs(evaluateFormula('VDB(2400,300,10,0,10)') - 2100) < 1e-10, 'VDB switches to straight-line depreciation and reaches salvage value');
assert.ok(Math.abs(evaluateFormula('VDB(2400,300,10,0,10,2,TRUE)') - 2100) < 1e-10, 'VDB can disable the straight-line switch');
assert.equal(evaluateFormula('VDB(2400,300,10,2,2)'), '#NUM!');
assert.ok(Math.abs(evaluateFormula('DB(1000000,100000,6,1,7)') - 186083.33333333334) < 1e-10);
assert.ok(Math.abs(evaluateFormula('DB(1000000,100000,6,2,7)') - (1000000 - 186083.33333333334) * 0.319) < 1e-10);
assert.equal(evaluateFormula('DB(1000000,100000,6,8,7)'), '#NUM!');
assert.equal(evaluateFormula('DB(1000000,100000,6,1,0)'), '#NUM!');
const scheduleContext = { getCell: row => [0.02, 0.03, 0.04][row] ?? null };
assert.ok(Math.abs(evaluateFormula('FVSCHEDULE(1000,A1:A3)', scheduleContext) - 1092.624) < 1e-10);
assert.equal(evaluateFormula('FVSCHEDULE(1000,A1:A1)', { getCell: () => -1 }), '#NUM!');
assert.equal(evaluateFormula('FVSCHEDULE(1000)'), '#VALUE!');
assert.ok(Math.abs(evaluateFormula('RRI(10,10000,15000)') - 0.0413797439924106) < 1e-12);
assert.ok(Math.abs(evaluateFormula('PDURATION(0.1,1000,2000)') - 7.272540897341713) < 1e-12);
assert.equal(evaluateFormula('ISPMT(0.1,1,10,10000)'), -900);
assert.equal(evaluateFormula('ISPMT(0.1,10,10,10000)'), 0);
assert.equal(evaluateFormula('RRI(0,100,200)'), '#NUM!');
assert.equal(evaluateFormula('DOLLARDE(1.02,16)'), 1.125);
assert.equal(evaluateFormula('DOLLARFR(1.125,16)'), 1.02);
assert.equal(evaluateFormula('DOLLARDE(-1.02,16)'), -1.125);
assert.equal(evaluateFormula('DOLLARFR(-1.125,16)'), -1.02);
assert.equal(evaluateFormula('DOLLARFR(1.1,16)'), 1.016);
assert.equal(evaluateFormula('DOLLARDE(1.02,16.9)'), 1.125, 'DOLLARDE truncates a fractional denominator');
assert.equal(evaluateFormula('DOLLARFR(1.125,16.9)'), 1.02, 'DOLLARFR truncates a fractional denominator');
assert.equal(evaluateFormula('DOLLARDE(1.02,0)'), '#NUM!');
assert.equal(evaluateFormula('DOLLARFR(1.02,0)'), '#NUM!');
assert.equal(evaluateFormula('BASE(7,2)'), '111');
assert.equal(evaluateFormula('BASE(100,16)'), '64');
assert.equal(evaluateFormula('BASE(15,2,10)'), '0000001111');
assert.equal(evaluateFormula('BASE(7.9,2)'), '111');
assert.equal(evaluateFormula('DECIMAL("111",2)'), 7);
assert.equal(evaluateFormula('DECIMAL("FF",16)'), 255);
assert.equal(evaluateFormula('DECIMAL("zap",36)'), 45745);
assert.equal(evaluateFormula('DECIMAL("ff",16)'), 255);
assert.equal(evaluateFormula('BASE(-1,2)'), '#NUM!');
assert.equal(evaluateFormula('BASE(-0.5,2)'), '#NUM!');
assert.equal(evaluateFormula('BASE(7,1)'), '#NUM!');
assert.equal(evaluateFormula('BASE(7,2,256)'), '#NUM!');
assert.equal(evaluateFormula('DECIMAL("102",2)'), '#NUM!');
assert.equal(evaluateFormula('DECIMAL("",16)'), '#NUM!');
assert.equal(evaluateFormula('ACCRINTM(DATE(2024,1,1),DATE(2025,1,1),0.1,1000,0)'), 100);
assert.ok(Math.abs(evaluateFormula('ACCRINTM(DATE(2024,1,1),DATE(2025,1,1),0.1,1000,3)') - 100 * 366 / 365) < 1e-12);
assert.equal(evaluateFormula('ACCRINTM(DATE(2024,1,1),DATE(2025,1,1),0.1,1000,5)'), '#NUM!');
assert.equal(evaluateFormula('ACCRINTM(DATE(2024,1,1),DATE(2024,1,1),0.1,1000)'), '#NUM!');
assert.equal(evaluateFormula('PDURATION(0,100,200)'), '#NUM!');
assert.equal(evaluateFormula('ISPMT(0.1,11,10,10000)'), '#NUM!');
assert.equal(evaluateFormula('RRI(10,0,200)'), '#NUM!');
const settlement = evaluateFormula('DATE(2024,1,1)'), maturity = evaluateFormula('DATE(2024,7,1)');
const investmentYearFraction = evaluateFormula(`YEARFRAC(${settlement},${maturity},3)`);
assert.ok(Math.abs(evaluateFormula(`DISC(${settlement},${maturity},97,100,3)`) - 0.03 / investmentYearFraction) < 1e-12);
assert.ok(Math.abs(evaluateFormula(`INTRATE(${settlement},${maturity},97,100,3)`) - (3 / 97) / investmentYearFraction) < 1e-12);
assert.ok(Math.abs(evaluateFormula(`RECEIVED(${settlement},${maturity},100,0.05,3)`) - 100 / (1 - 0.05 * investmentYearFraction)) < 1e-12);
assert.ok(Math.abs(evaluateFormula(`PRICEDISC(${settlement},${maturity},0.05,100,3)`) - 100 * (1 - 0.05 * investmentYearFraction)) < 1e-12);
const discountPrice = evaluateFormula(`PRICEDISC(${settlement},${maturity},0.05,100,3)`);
assert.ok(Math.abs(evaluateFormula(`YIELDDISC(${settlement},${maturity},${discountPrice},100,3)`) - 0.05 / (1 - 0.05 * investmentYearFraction)) < 1e-12);
const quotedBondPrice = evaluateFormula('PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0)');
assert.ok(Math.abs(quotedBondPrice - 94.6343616213) < 1e-9);
assert.ok(Math.abs(evaluateFormula(`YIELD(DATE(2008,2,15),DATE(2017,11,15),0.0575,${quotedBondPrice},100,2,0)`) - 0.065) < 1e-10);
assert.ok(Math.abs(evaluateFormula('PRICE(DATE(2024,2,29),DATE(2024,8,31),0.06,0.05,100,2,1)') - 100.48780487804878) < 1e-12);
assert.equal(evaluateFormula(`DISC(${settlement},${settlement},97,100)`), '#NUM!');
assert.equal(evaluateFormula(`INTRATE(${settlement},${maturity},97,100,5)`), '#NUM!');
assert.equal(evaluateFormula(`RECEIVED(${settlement},${maturity},100,-0.01)`), '#NUM!');
assert.equal(evaluateFormula(`PRICEDISC(${settlement},${maturity},-0.01,100)`), '#NUM!');
assert.equal(evaluateFormula(`YIELDDISC(${settlement},${maturity},0,100)`), '#NUM!');
assert.equal(evaluateFormula('PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,-2,100,2)'), '#NUM!');
assert.equal(evaluateFormula('PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,3)'), '#NUM!');
assert.equal(evaluateFormula('YIELD(DATE(2008,2,15),DATE(2017,11,15),0.0575,0,100,2)'), '#NUM!');
assert.equal(evaluateFormula('MINIFS(B1:B2,A1:A2,">2")', context), 5);
  assert.equal(evaluateFormula('MAXIFS(B1:B2,A1:A2,">1",B1:B2,"<5")', context), 3, 'MAXIFS applies every criteria pair to the aligned positions');
  assert.equal(evaluateFormula('MINIFS(B1:B2,A1:A2,">100")', context), 0, 'MAXIFS and MINIFS return zero when no numeric value matches');
  assert.equal(evaluateFormula('MAXIFS(B1:B2,A1:A2)', context), '#VALUE!', 'MAXIFS requires a criteria range and criterion pair');
  assert.equal(evaluateFormula('FIND("E","Excel")'), 1, 'FIND is case-sensitive');
  assert.equal(evaluateFormula('FIND("e","Excel",3)'), 4, 'FIND honors the one-based start position');
  assert.equal(evaluateFormula('FIND("e","Excel",5)'), '#VALUE!', 'FIND reports text that is not found from the requested position');
  assert.equal(evaluateFormula('SEARCH("e","Excel")'), 1, 'SEARCH ignores case');
  assert.equal(evaluateFormula('SEARCH("a?c","abc")'), 1, 'SEARCH supports wildcard characters');
  assert.equal(evaluateFormula('SEARCH("a~*c","a*c")'), 1, 'SEARCH supports tilde-escaped literal wildcards');
  assert.equal(evaluateFormula('SEARCH("~ap*","apple")'), '#VALUE!', 'SEARCH does not discard a tilde before ordinary text');
  assert.equal(evaluateFormula('SEARCH("~ap*","~apple")'), 1, 'SEARCH treats a tilde before ordinary text literally');
  assert.equal(evaluateFormula('SEARCH("a","😀a",2)'), 2, 'SEARCH start positions count Unicode characters');
  assert.equal(evaluateFormula('SUBSTITUTE("abab","a","x",2)'), 'abxb', 'SUBSTITUTE can replace only the requested occurrence');
  assert.equal(evaluateFormula('SUBSTITUTE("abab","a","x")'), 'xbxb', 'SUBSTITUTE replaces every occurrence by default');
  assert.equal(evaluateFormula('REPLACE("abcdef",2,3,"X")'), 'aXef');
  assert.equal(evaluateFormula('REPLACE("A😀B",2,1,"x")'), 'AxB', 'REPLACE counts Unicode characters');
  assert.equal(evaluateFormula('EXACT("Excel","Excel")'), true);
  assert.equal(evaluateFormula('EXACT("Excel","EXCEL")'), false, 'EXACT distinguishes letter case');
  assert.equal(evaluateFormula('PROPER("jin sheet-EXCEL 테스트")'), 'Jin Sheet-Excel 테스트');
  assert.equal(evaluateFormula('REPT("ab",2.8)'), 'abab', 'REPT truncates the repeat count');
  assert.equal(evaluateFormula('REPT("x",-1)'), '#VALUE!');
  assert.equal(evaluateFormula('CLEAN(A1)', { getCell: () => 'a\nb' }), 'ab');
  assert.equal(evaluateFormula('T("text")'), 'text');
  assert.equal(evaluateFormula('T(42)'), '');
  assert.equal(evaluateFormula('N(TRUE)'), 1);
  assert.equal(evaluateFormula('N("42")'), 0);
  assert.equal(evaluateFormula('ISERROR(1/0)'), true);
  assert.equal(evaluateFormula('ISERR(1/0)'), true);
  assert.equal(evaluateFormula('ISNA(A1)', { getCell: () => '#N/A' }), true);
  assert.equal(evaluateFormula('ISNONTEXT(A1)', { getCell: () => 42 }), true);
  assert.equal(evaluateFormula('ISNONTEXT("text")'), false);
  assert.equal(evaluateFormula('ISNONTEXT("")'), false, 'an empty string is still text');
  assert.equal(evaluateFormula('ISODD(-3.9)'), true, 'parity uses the integer part of the absolute value');
  assert.equal(evaluateFormula('ISEVEN(-3.9)'), false);
  assert.equal(evaluateFormula('ISEVEN("4.8")'), true, 'numeric text is coerced');
  assert.equal(evaluateFormula('ISODD("text")'), '#VALUE!');
  assert.equal(evaluateFormula('NA()'), '#N/A');
  assert.equal(evaluateFormula('NA(1)'), '#VALUE!');
  assert.equal(evaluateFormula('ISERR(A1)', { getCell: () => '#N/A' }), false);
  assert.equal(evaluateFormula('ERROR.TYPE(1/0)'), 2);
  assert.equal(evaluateFormula('ERROR.TYPE(#N/A!)'), 7, 'the #N/A! alias has the same error type as #N/A');
  assert.equal(evaluateFormula('ERROR.TYPE(#SPILL!)'), 9, 'Excel assigns error type 9 to spill errors');
  assert.equal(evaluateFormula('ERROR.TYPE(#CALC!)'), 14, 'Excel assigns error type 14 to calculation errors');
  assert.equal(evaluateFormula('ERROR.TYPE("plain text")'), '#N/A');
  assert.equal(evaluateFormula('SIGN(-2.7)'), -1);
  assert.equal(evaluateFormula('SIGN(0)'), 0);
  assert.equal(evaluateFormula('QUOTIENT(-7,2)'), -3, 'QUOTIENT truncates toward zero');
  assert.equal(evaluateFormula('QUOTIENT(7,0)'), '#DIV/0!');
  assert.equal(evaluateFormula('EVEN(-1.1)'), -2);
  assert.equal(evaluateFormula('EVEN(0)'), 0);
  assert.equal(evaluateFormula('ODD(2)'), 3);
  assert.equal(evaluateFormula('ODD(0)'), 1);
  assert.equal(evaluateFormula('WEEKDAY(DATE(2024,1,1))'), 2, 'the default week starts on Sunday');
  assert.equal(evaluateFormula('WEEKDAY(DATE(2024,1,1),2)'), 1, 'return type 2 starts the week on Monday');
  assert.equal(evaluateFormula('WEEKDAY(DATE(2024,1,1),3)'), 0);
  assert.equal(evaluateFormula('WEEKDAY(60)'), 5, 'Excel serial 60 is its fictitious Thursday');
  assert.equal(evaluateFormula('WEEKDAY(DATE(2024,1,1),4)'), '#NUM!');
  assert.equal(evaluateFormula('WEEKNUM(DATE(2021,1,1))'), 1);
  assert.equal(evaluateFormula('WEEKNUM(DATE(2021,1,1),2)'), 1);
  assert.equal(evaluateFormula('WEEKNUM(DATE(2021,1,4),2)'), 2);
  assert.equal(evaluateFormula('WEEKNUM(DATE(2021,1,1),21)'), 53);
  assert.equal(evaluateFormula('ISOWEEKNUM(DATE(2021,1,1))'), 53);
  assert.equal(evaluateFormula('ISOWEEKNUM(DATE(2021,1,4))'), 1);
  assert.equal(evaluateFormula('WEEKNUM(DATE(2024,1,1),3)'), '#NUM!');
  assert.equal(evaluateFormula('EDATE(DATE(2024,1,31),1)'), evaluateFormula('DATE(2024,2,29)'), 'EDATE clamps to the target month end');
  assert.equal(evaluateFormula('EOMONTH(DATE(2024,1,15),1)'), evaluateFormula('DATE(2024,2,29)'));
  assert.equal(evaluateFormula('EOMONTH(DATE(2024,1,15),0)'), evaluateFormula('DATE(2024,1,31)'));
  assert.equal(evaluateFormula('EDATE(DATE(2024,1,1),-1)'), evaluateFormula('DATE(2023,12,1)'));
  assert.equal(evaluateFormula('DAYS(DATE(2024,3,1),DATE(2024,2,28))'), 2);
  assert.equal(evaluateFormula('NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,7))'), 5, 'NETWORKDAYS counts weekdays inclusively');
  assert.equal(evaluateFormula('NETWORKDAYS(DATE(2024,1,7),DATE(2024,1,1))'), -5, 'reversed date arguments return a negative count');
  assert.equal(evaluateFormula('NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,7),DATE(2024,1,3))'), 4, 'holidays exclude weekdays');
  assert.equal(evaluateFormula('NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,7),DATE(2024,1,6))'), 5, 'holidays on weekends do not reduce the count');
  assert.equal(evaluateFormula('NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),11)'), 6, 'custom weekend codes can exclude a single weekend day');
  assert.equal(evaluateFormula('NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),"0000011")'), 5, 'seven-character weekend masks map Monday through Sunday');
  assert.equal(evaluateFormula('NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),11,DATE(2024,1,2))'), 5, 'custom weekend rules still apply weekday holidays');
  assert.equal(evaluateFormula('WORKDAY(DATE(2024,1,5),1)'), evaluateFormula('DATE(2024,1,8)'), 'WORKDAY skips the weekend');
  assert.equal(evaluateFormula('WORKDAY(DATE(2024,1,5),1,DATE(2024,1,8))'), evaluateFormula('DATE(2024,1,9)'), 'WORKDAY skips a weekday holiday');
  assert.equal(evaluateFormula('WORKDAY.INTL(DATE(2024,1,5),1,11)'), evaluateFormula('DATE(2024,1,6)'), 'custom weekends allow Saturday workdays');
  assert.equal(evaluateFormula('WORKDAY.INTL(DATE(2024,1,5),1,"0000011")'), evaluateFormula('DATE(2024,1,8)'), 'custom weekend masks work with offsets');
  assert.equal(evaluateFormula('WORKDAY.INTL(DATE(2024,1,5),1,11,DATE(2024,1,6))'), evaluateFormula('DATE(2024,1,8)'), 'custom holiday arguments are honored');
  assert.equal(evaluateFormula('WORKDAY(DATE(2024,1,8),-1)'), evaluateFormula('DATE(2024,1,5)'), 'negative offsets move backwards over weekdays');
  assert.equal(evaluateFormula('WORKDAY(DATE(2024,1,8),0)'), evaluateFormula('DATE(2024,1,8)'), 'zero offset preserves the starting date');
  assert.equal(evaluateFormula('NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,7),A1:A2)', { getCell: row => row === 0 ? evaluateFormula('DATE(2024,1,3)') : evaluateFormula('DATE(2024,1,3)') }), 4, 'duplicate holidays from a range are counted once');
  assert.equal(evaluateFormula('WORKDAY(2958465,1)'), '#NUM!', 'WORKDAY rejects dates outside Excel limits');
  assert.equal(evaluateFormula('NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),"1111111")'), '#NUM!', 'weekend masks must leave at least one working day');
assert.equal(evaluateFormula('DAYS360(DATE(2024,1,1),DATE(2024,12,31))'), 360);
assert.equal(evaluateFormula('DATEDIF(DATE(2020,1,1),DATE(2023,1,1),"Y")'), 3);
assert.equal(evaluateFormula('DATEDIF(DATE(2020,1,1),DATE(2023,1,1),"M")'), 36);
assert.equal(evaluateFormula('DATEDIF(DATE(2020,1,1),DATE(2023,1,1),"D")'), 1096);
assert.equal(evaluateFormula('DATEDIF(DATE(2020,1,15),DATE(2022,6,14),"YM")'), 4);
assert.equal(evaluateFormula('DATEDIF(DATE(2020,1,1),DATE(2020,12,31),"YD")'), 365);
assert.equal(evaluateFormula('DATEDIF(DATE(2020,1,31),DATE(2020,2,15),"MD")'), 15);
assert.equal(evaluateFormula('DATEDIF(DATE(2020,1,31),DATE(2020,3,1),"MD")'), -1);
assert.equal(evaluateFormula('DATEDIF(DATE(2020,1,15),DATE(2022,6,14),"YD")'), 150);
assert.equal(evaluateFormula('DATEDIF(DATE(2020,2,29),DATE(2021,3,1),"YD")'), 0);
assert.equal(evaluateFormula('DATEDIF(DATE(2020,1,1),DATE(2021,1,1),"y")'), 1, 'DATEDIF units are case-insensitive');
assert.equal(evaluateFormula('DATEDIF(DATE(2021,1,1),DATE(2020,1,1),"D")'), '#NUM!');
assert.equal(evaluateFormula('DATEDIF(DATE(2020,1,1),DATE(2021,1,1),"Q")'), '#NUM!');
assert.equal(evaluateFormula('DATEDIF(DATE(2020,1,1),DATE(2021,1,1),1)'), '#VALUE!');
assert.equal(evaluateFormula('DAYS360(DATE(2024,1,1),DATE(2024,12,31),TRUE)'), 359);
assert.equal(evaluateFormula('YEARFRAC(DATE(2024,1,1),DATE(2025,1,1))'), 1);
assert.equal(evaluateFormula('DATEVALUE("2024-02-29")'), evaluateFormula('DATE(2024,2,29)'));
assert.equal(evaluateFormula('DATEVALUE("2024년 2월 29일")'), evaluateFormula('DATE(2024,2,29)'));
assert.equal(evaluateFormula('DATEVALUE("2/29/2024")'), evaluateFormula('DATE(2024,2,29)'));
assert.equal(evaluateFormula('DATEVALUE("Feb 29, 2024")'), evaluateFormula('DATE(2024,2,29)'));
assert.equal(evaluateFormula('DATEVALUE("2024-02-29T12:30:00")'), evaluateFormula('DATE(2024,2,29)'));
assert.equal(evaluateFormula('DATEVALUE("2024-02-30")'), '#VALUE!');
assert.equal(evaluateFormula('DATEVALUE(45352)'), '#VALUE!');
assert.equal(evaluateFormula('YEARFRAC(DATE(2024,1,1),DATE(2025,1,1),1)'), 1);
  assert.equal(evaluateFormula('YEARFRAC(DATE(2024,1,1),DATE(2025,1,1),2)'), 366 / 360);
  assert.equal(evaluateFormula('YEARFRAC(DATE(2024,1,1),DATE(2025,1,1),3)'), 366 / 365);
  assert.equal(evaluateFormula('YEARFRAC(DATE(2024,1,1),DATE(2024,12,31),4)'), 359 / 360);
assert.equal(evaluateFormula('YEARFRAC(DATE(2025,1,1),DATE(2024,1,1))'), -1);
assert.equal(evaluateFormula('TIMEVALUE("12:00 PM")'), 0.5);
assert.equal(evaluateFormula('TIMEVALUE("00:15:30")'), 930 / 86400);
assert.equal(evaluateFormula('TIMEVALUE("오후 1:30")'), 13.5 / 24);
assert.equal(evaluateFormula('TIMEVALUE("2024-01-01T12:00:00")'), 0.5);
assert.equal(evaluateFormula('TIMEVALUE("24:00")'), '#VALUE!');
assert.equal(evaluateFormula('TIMEVALUE(0.5)'), '#VALUE!');
assert.equal(evaluateFormula('COUPDAYBS(DATE(2008,1,25),DATE(2008,11,15),2,0)'), 70);
assert.equal(evaluateFormula('COUPDAYS(DATE(2008,1,25),DATE(2008,11,15),2,0)'), 180);
assert.equal(evaluateFormula('COUPDAYSNC(DATE(2008,1,25),DATE(2008,11,15),2,0)'), 110);
assert.equal(evaluateFormula('COUPNCD(DATE(2008,1,25),DATE(2008,11,15),2,0)'), evaluateFormula('DATE(2008,5,15)'));
assert.equal(evaluateFormula('COUPPCD(DATE(2008,1,25),DATE(2008,11,15),2,0)'), evaluateFormula('DATE(2007,11,15)'));
assert.equal(evaluateFormula('COUPNUM(DATE(2008,1,25),DATE(2008,11,15),2,0)'), 2);
assert.equal(evaluateFormula('COUPDAYBS(DATE(2008,1,25),DATE(2008,11,15),2,1)'), 71);
assert.equal(evaluateFormula('COUPDAYS(DATE(2008,1,25),DATE(2008,11,15),2,1)'), 182);
assert.equal(evaluateFormula('COUPDAYSNC(DATE(2008,1,25),DATE(2008,11,15),2,1)'), 111);
assert.equal(evaluateFormula('COUPNCD(DATE(2024,1,15),DATE(2024,8,31),2)'), evaluateFormula('DATE(2024,2,29)'), 'coupon schedule preserves maturity month-end dates');
assert.equal(evaluateFormula('COUPPCD(DATE(2024,2,29),DATE(2024,8,31),2)'), evaluateFormula('DATE(2024,2,29)'), 'a settlement on a coupon date is the previous coupon date');
assert.equal(evaluateFormula('COUPNCD(DATE(2024,2,29),DATE(2024,8,31),2)'), evaluateFormula('DATE(2024,8,31)'));
assert.equal(evaluateFormula('COUPNUM(DATE(2024,2,29),DATE(2024,8,31),2)'), 1);
assert.equal(evaluateFormula('COUPNUM(DATE(2024,1,15),DATE(2024,8,31),3)'), '#NUM!');
  const lookupContext = { getCell: (row, column) => ({ '0:0': 'Alpha', '0:1': 10, '1:0': 'Beta', '1:1': 20, '2:0': 'A*Star', '2:1': 30 })[row + ':' + column] ?? null };
assert.equal(evaluateFormula('MATCH("a*",A1:A3,0)', lookupContext), 1, 'MATCH exact mode supports case-insensitive Excel wildcards');
assert.equal(evaluateFormula('MATCH("A~*S?ar",A1:A3,0)', lookupContext), 3, 'MATCH lets tilde escape a literal wildcard');
assert.equal(evaluateFormula('XMATCH("Beta",A1:A3)', lookupContext), 2);
assert.equal(evaluateFormula('XMATCH("A~*S?ar",A1:A3,2)', lookupContext), 3, 'XMATCH wildcard mode honors escaped wildcard characters');
assert.equal(evaluateFormula('XMATCH("Beta",A1:A3,0,-1)', lookupContext), 2, 'XMATCH can search backwards');
const sortedLookupContext = { getCell: row => [10, 20, 30, 40][row] };
assert.equal(evaluateFormula('XMATCH(25,A1:A4,1)', sortedLookupContext), 3, 'XMATCH returns the next larger item');
assert.equal(evaluateFormula('XMATCH(25,A1:A4,-1)', sortedLookupContext), 2, 'XMATCH returns the next smaller item');
assert.equal(evaluateFormula('XMATCH(25,A1:A4,-1,2)', sortedLookupContext), 2, 'XMATCH binary search supports ascending ranges');
assert.equal(evaluateFormula('XMATCH(25,A1:A4,1,-2)', { getCell: row => [40, 30, 20, 10][row] }), 2, 'XMATCH binary search supports descending ranges');
assert.equal(evaluateFormula('XMATCH(25,A1:A4,-1,-2)', { getCell: row => [40, 30, 20, 10][row] }), 3);
assert.equal(evaluateFormula('XMATCH(99,A1:A4)', sortedLookupContext), '#N/A');
assert.equal(evaluateFormula('XMATCH(1,A1:A4,3)', sortedLookupContext), '#VALUE!');
assert.equal(evaluateFormula('VLOOKUP("B?ta",A1:B3,2,FALSE)', lookupContext), 20, 'VLOOKUP exact mode supports wildcards');
  assert.equal(evaluateFormula('VLOOKUP("B~?ta",A1:B3,2,FALSE)', lookupContext), '#N/A', 'tilde-escaped wildcard characters are matched literally');
  assert.equal(evaluateFormula('COUNTBLANK(A1:A2)', { getCell: (row, column) => [null, '', 0][row] }), 2);
  assert.equal(evaluateFormula('COUNTBLANK(A1)', { getCell: () => null }), 1);
  assert.equal(evaluateFormula('A1', { getCell: () => null }), 0, 'a blank cell reference evaluates to zero when returned directly');
  assert.equal(evaluateFormula('COUNTBLANK(A1)', { getCell: () => '#N/A' }), 0);
  assert.equal(evaluateFormula('COUNTA(A1:A3)', { getCell: (row, column) => [null, '', 0][row] }), 2);
  assert.equal(evaluateFormula('COUNTA("")'), 1, 'COUNTA includes a zero-length string value');
  assert.equal(evaluateFormula('COUNTA(A1)', { getCell: () => null }), 0, 'COUNTA excludes a truly blank cell reference');
  const averageAContext = { getCell: row => [2, true, false, 'text', null][row] ?? null };
  assert.equal(evaluateFormula('AVERAGEA(A1:A5)', averageAContext), 0.75, 'AVERAGEA includes logical and text cells but skips blanks');
  assert.equal(evaluateFormula('MAXA(A1:A5)', averageAContext), 2);
  assert.equal(evaluateFormula('MINA(A1:A5)', averageAContext), 0);
  assert.equal(evaluateFormula('AVERAGEA(2,TRUE,"text")'), 1);
  assert.equal(evaluateFormula('MAXA("text",FALSE)'), 0);
  assert.equal(evaluateFormula('AVERAGEA(A1:A2)', { getCell: () => null }), '#DIV/0!');
  assert.equal(evaluateFormula('MAXA(A1:A2)', { getCell: () => null }), 0);
  assert.equal(evaluateFormula('MAXA(SEQUENCE(100,100))'), 10000, 'large arrays are aggregated without argument spreading');
  assert.equal(evaluateFormula('AVERAGEA(1E308,1E308)'), '#NUM!', 'overflow is reported as a spreadsheet error');
  assert.equal(evaluateFormula('SUM(1E308,1E308)'), '#NUM!', 'SUM reports numeric overflow instead of returning Infinity');
  assert.equal(evaluateFormula('AVERAGE(1E308,1E308)'), '#NUM!', 'AVERAGE reports an overflowing intermediate sum');
  assert.equal(evaluateFormula('SUM(1E308,1E308,-1E308)'), '#NUM!', 'SUM matches the LibreOffice reference when an intermediate sum overflows');
  assert.equal(evaluateFormula('AVERAGE(1E308,1E308,-1E308)'), '#NUM!', 'AVERAGE matches the LibreOffice reference when its intermediate sum overflows');
  assert.equal(evaluateFormula('ISBLANK(A1)', { getCell: () => null }), true);
  assert.equal(evaluateFormula('ISBLANK("")'), false);
  const formulaInspection = { getSheetIndex: name => ({ data: 1, current: 0 })[name.toLowerCase()] ?? -1, getCellFormula: (row, column, sheetName) => sheetName === 'Data' && row === 1 && column === 2 ? 'SUM(A1:A3)' : row === 0 && column === 0 ? '1+2' : null };
  assert.equal(evaluateFormula('ISFORMULA(A1)', formulaInspection), true);
  assert.equal(evaluateFormula('ISFORMULA(B2)', formulaInspection), false);
  assert.equal(evaluateFormula('FORMULATEXT(A1)', formulaInspection), '=1+2');
  assert.equal(evaluateFormula('FORMULATEXT(Data!C2)', formulaInspection), '=SUM(A1:A3)');
  assert.equal(evaluateFormula('FORMULATEXT(B2)', formulaInspection), '#N/A');
  assert.equal(evaluateFormula('ISFORMULA(Missing!A1)', formulaInspection), '#REF!');
  assert.equal(evaluateFormula('ISFORMULA(A1,A2)', formulaInspection), '#VALUE!');
  assert.equal(evaluateFormula('ISREF(A1)', formulaInspection), true);
  assert.equal(evaluateFormula('ISREF(Data!A1:B2)', formulaInspection), true);
  assert.equal(evaluateFormula('ISREF(Missing!A1)', formulaInspection), false);
  assert.equal(evaluateFormula('ISREF(1+2)', formulaInspection), false);
  const sheetInfo = { getSheetIndex: name => ({ data: 1, current: 0 })[String(name || 'current').toLowerCase()] ?? -1, getSheetCount: () => 3 };
  assert.equal(evaluateFormula('SHEET()', sheetInfo), 1);
  assert.equal(evaluateFormula('SHEET(Data!A1)', sheetInfo), 2);
  assert.equal(evaluateFormula('SHEETS()', sheetInfo), 3);
  assert.equal(evaluateFormula('SHEETS(Data!A1:B2)', sheetInfo), 1);
  assert.equal(evaluateFormula('SHEET(Missing!A1)', sheetInfo), '#REF!');
  assert.equal(evaluateFormula('SHEETS(Data!A1:Missing!B2)', sheetInfo), '#REF!');
  assert.equal(evaluateFormula('SHEET(A1,A2)', sheetInfo), '#VALUE!');
  const formulaOrigin = { ...sheetInfo, formulaRow: 4, formulaColumn: 2 };
  assert.equal(evaluateFormula('ROW()', formulaOrigin), 5);
  assert.equal(evaluateFormula('COLUMN()', formulaOrigin), 3);
  assert.equal(evaluateFormula('ROW(Data!B3)', formulaOrigin), 3);
  assert.equal(evaluateFormula('COLUMN(Data!B3)', formulaOrigin), 2);
  assert.equal(evaluateFormula('ROW(A1:B4)'), 1);
  assert.equal(evaluateFormula('COLUMN(A1:C4)'), 1);
  assert.equal(evaluateFormula('ROW(Missing!A1)', formulaOrigin), '#REF!');
  assert.equal(evaluateFormula('ROWS(A1:C4)'), 4);
  assert.equal(evaluateFormula('COLUMNS(A1:C4)'), 3);
  assert.equal(evaluateFormula('ROWS(A1)'), 1);
  assert.equal(evaluateFormula('COLUMNS(A1)'), 1);
  assert.equal(evaluateFormula('ADDRESS(1,1)'), '$A$1');
  assert.equal(evaluateFormula('ADDRESS(3,28,2)'), 'AB$3');
  assert.equal(evaluateFormula('ADDRESS(1,1,3)'), '$A1');
  assert.equal(evaluateFormula('ADDRESS(1,1,4)'), 'A1');
  assert.equal(evaluateFormula('ADDRESS(1,1,1,FALSE)'), 'R1C1');
  assert.equal(evaluateFormula('ADDRESS(2,3,4,FALSE)'), 'R[2]C[3]');
  assert.equal(evaluateFormula("ADDRESS(1,1,1,TRUE,\"Sales O'Brien\")"), "'Sales O''Brien'!$A$1");
  assert.equal(evaluateFormula('ADDRESS(1048577,1)'), '#VALUE!');
  assert.equal(evaluateFormula('ADDRESS(1/0,1)'), '#DIV/0!');
  assert.equal(evaluateFormula('ISNUMBER(1)'), true);
  assert.equal(evaluateFormula('ISNUMBER("1")'), false);
  assert.equal(evaluateFormula('ISTEXT("")'), true);
  assert.equal(evaluateFormula('ISTEXT(TRUE)'), false);
  assert.equal(evaluateFormula('ISLOGICAL(TRUE)'), true);
  assert.equal(evaluateFormula('ISLOGICAL(1)'), false);
  assert.equal(evaluateFormula('ISBLANK(A1,A2)'), '#VALUE!');
  assert.equal(evaluateFormula("'Sales O''Brien'!A1*2", { getCell: (row, column, sheet) => sheet === "Sales O'Brien" ? 21 : null }), 42);
  assert.equal(evaluateFormula('AVERAGE(A1:B2)', context), 3.5);
  const directTextAggregateContext = { getCell: (row, column) => [[5, '2', true]][row]?.[column] ?? null };
  assert.equal(evaluateFormula('SUM(1,"2")'), 3, 'SUM coerces numeric text written directly as an argument');
  assert.equal(evaluateFormula('SUM(1,"text")'), '#VALUE!', 'SUM rejects nonnumeric text written directly as an argument');
  assert.equal(evaluateFormula('SUM(B1)', directTextAggregateContext), 0, 'SUM ignores text returned by a single-cell reference');
  assert.equal(evaluateFormula('SUM(C1)', directTextAggregateContext), 0, 'SUM ignores a logical value returned by a single-cell reference');
  assert.equal(evaluateFormula('SUM(A1:C1)', directTextAggregateContext), 5, 'SUM ignores text and logical values in a referenced range');
  assert.equal(evaluateFormula('SUM(TRUE(),2)'), 3, 'SUM includes a logical value passed directly as an argument');
  assert.equal(evaluateFormula('AVERAGE(1,"2")'), 1.5, 'AVERAGE counts numeric text written directly as an argument');
  assert.equal(evaluateFormula('AVERAGE(B1)', directTextAggregateContext), '#DIV/0!', 'AVERAGE ignores text returned by a single-cell reference');
  assert.equal(evaluateFormula('AVERAGE(A1:C1)', directTextAggregateContext), 5, 'AVERAGE ignores text and logical values in a referenced range');
  assert.equal(evaluateFormula('COUNT(1,"2")'), 2, 'COUNT counts numeric text written directly as an argument');
  assert.equal(evaluateFormula('COUNT(TRUE(),2)'), 2, 'COUNT counts logical values passed directly as arguments');
  assert.equal(evaluateFormula('COUNT(B1)', directTextAggregateContext), 0, 'COUNT ignores text returned by a single-cell reference');
  assert.equal(evaluateFormula('COUNT(C1)', directTextAggregateContext), 0, 'COUNT ignores a logical value returned by a single-cell reference');
  assert.equal(evaluateFormula('COUNT(A1:C1)', directTextAggregateContext), 1, 'COUNT ignores text and logical values in a referenced range');
  assert.equal(evaluateFormula('MEDIAN(1,5,3,9)'), 4);
  assert.equal(evaluateFormula('MEDIAN(1E308,1E308)'), 1E308, 'the midpoint avoids overflowing the sum of large values');
  assert.equal(evaluateFormula('MEDIAN(-1E308,1E308)'), 0, 'the midpoint handles large opposite-signed values');
  assert.equal(evaluateFormula('MEDIAN(5E-324,5E-324)'), 5E-324, 'the midpoint preserves the smallest positive subnormal values');
  assert.equal(evaluateFormula('MEDIAN(1,5,3)'), 3);
  assert.equal(evaluateFormula('MEDIAN(A1:B2)', context), 3.5);
  assert.equal(evaluateFormula('MEDIAN()'), '#NUM!');
  assert.equal(evaluateFormula('PRODUCT(2,-3,4)'), -24);
  assert.equal(evaluateFormula('PRODUCT()'), '#VALUE!');
  assert.equal(evaluateFormula('PRODUCT(A1:B2)', context), 120);
  assert.equal(evaluateFormula('LARGE(A1:B2,2)', context), 4);
  assert.equal(evaluateFormula('SMALL(A1:B2,2)', context), 3);
  assert.equal(evaluateFormula('LARGE(A1:B2,0)', context), '#NUM!');
  assert.equal(evaluateFormula('SMALL(A1:B2,9)', context), '#NUM!');
  assert.equal(evaluateFormula('LARGE(A1:A2,1)', { getCell: row => row === 0 ? '#N/A' : 2 }), '#N/A');
  assert.equal(evaluateFormula('RANK.EQ(4,A1:B2)', context), 2);
  assert.equal(evaluateFormula('RANK.EQ(3,A1:B2,1)', context), 2, 'ascending rank reverses the order');
  assert.equal(evaluateFormula('RANK.EQ(3,A1:B2,1,2)', context), '#VALUE!');
  assert.equal(evaluateFormula('RANK.AVG(3,A1:B2)', { getCell: (row, column) => [[5, 3], [3, 1]][row][column] }), 2.5);
  assert.equal(evaluateFormula('RANK.EQ(9,A1:B2)', context), '#N/A');
  assert.equal(evaluateFormula('PERCENTRANK.INC(A1:A4,2)', { getCell: row => [1, 2, 3, 4][row] ?? null }), 0.333);
  assert.equal(evaluateFormula('PERCENTRANK.INC(A1:A4,2.5)', { getCell: row => [1, 2, 3, 4][row] ?? null }), 0.5);
  assert.equal(evaluateFormula('PERCENTRANK.INC(A1:A4,1)', { getCell: row => [1, 2, 3, 4][row] ?? null }), 0);
  assert.equal(evaluateFormula('PERCENTRANK.INC(A1:A4,4)', { getCell: row => [1, 2, 3, 4][row] ?? null }), 1);
  assert.equal(evaluateFormula('PERCENTRANK.INC(A1:A4,2.345,2)', { getCell: row => [1, 2, 3, 4][row] ?? null }), 0.45);
  assert.equal(evaluateFormula('PERCENTRANK.INC(A1:A4,2.345,2.9)', { getCell: row => [1, 2, 3, 4][row] ?? null }), 0.45);
  assert.equal(evaluateFormula('PERCENTRANK.INC(A1:A4,0)', { getCell: row => [1, 2, 3, 4][row] ?? null }), '#N/A');
  assert.equal(evaluateFormula('PERCENTRANK.EXC(A1:A4,2)', { getCell: row => [1, 2, 3, 4][row] ?? null }), 0.4);
  assert.equal(evaluateFormula('PERCENTRANK.EXC(A1:A4,2.5)', { getCell: row => [1, 2, 3, 4][row] ?? null }), 0.5);
  assert.equal(evaluateFormula('PERCENTRANK.EXC(A1:A4,0)', { getCell: row => [1, 2, 3, 4][row] ?? null }), '#N/A');
  assert.equal(evaluateFormula('PERCENTRANK.EXC(A1:A4,4)', { getCell: row => [1, 2, 3, 4][row] ?? null }), 0.8);
  assert.equal(evaluateFormula('PERCENTRANK.INC(A1:A4,2,0)', { getCell: row => [1, 2, 3, 4][row] ?? null }), '#NUM!');
  assert.equal(evaluateFormula('PERCENTRANK.INC(A1:A4,2,16)', { getCell: row => [1, 2, 3, 4][row] ?? null }), '#NUM!');
  assert.equal(evaluateFormula('PERCENTILE.INC(A1:B2,0.5)', context), 3.5);
  assert.equal(evaluateFormula('PERCENTILE(A1:B2,0.5)', context), 3.5);
const extremePercentileContext = { getCell: row => row === 0 ? -1E308 : 1E308 };
assert.equal(evaluateFormula('PERCENTILE.INC(A1:A2,0.25)', extremePercentileContext), -5E307);
assert.equal(evaluateFormula('PERCENTRANK.INC(A1:A2,0)', extremePercentileContext), 0.5);
  assert.equal(evaluateFormula('PERCENTILE.INC(A1:B2,-0.1)', context), '#NUM!');
  assert.equal(evaluateFormula('QUARTILE.INC(A1:B2,2)', context), 3.5);
  assert.equal(evaluateFormula('QUARTILE(A1:B2,2)', context), 3.5);
  assert.equal(evaluateFormula('QUARTILE.INC(A1:B2,5)', context), '#NUM!');
  const exclusivePercentileContext = { getCell: row => [1, 2, 3, 4][row] ?? null };
  assert.equal(evaluateFormula('PERCENTILE.EXC(A1:A4,0.2)', exclusivePercentileContext), 1);
  assert.equal(evaluateFormula('PERCENTILE.EXC(A1:A4,0.4)', exclusivePercentileContext), 2);
  assert.equal(evaluateFormula('PERCENTILE.EXC(A1:A4,0.5)', exclusivePercentileContext), 2.5);
  assert.equal(evaluateFormula('PERCENTILE.EXC(A1:A4,0)', exclusivePercentileContext), '#NUM!');
  assert.equal(evaluateFormula('PERCENTILE.EXC(A1:A4,1)', exclusivePercentileContext), '#NUM!');
  assert.equal(evaluateFormula('QUARTILE.EXC(A1:A4,1)', exclusivePercentileContext), 1.25);
  assert.equal(evaluateFormula('QUARTILE.EXC(A1:A4,2)', exclusivePercentileContext), 2.5);
  assert.equal(evaluateFormula('QUARTILE.EXC(A1:A4,3)', exclusivePercentileContext), 3.75);
  assert.equal(evaluateFormula('QUARTILE.EXC(A1:A4,0)', exclusivePercentileContext), '#NUM!');
  assert.equal(evaluateFormula('VAR.S(A1:B2)', context), 5 / 3);
  assert.equal(evaluateFormula('VAR.P(A1:B2)', context), 1.25);
  assert.equal(evaluateFormula('STDEV.S(A1:B2)', context), Math.sqrt(5 / 3));
  assert.equal(evaluateFormula('STDEV.P(A1:B2)', context), Math.sqrt(1.25));
  assert.equal(evaluateFormula('STDEV.P(1E308,1E308)'), 0, 'a constant large population has zero deviation without overflowing its mean');
  assert.equal(evaluateFormula('VAR.P(1E308,1E308)'), 0, 'a constant large population has zero variance without overflowing its mean');
  assert.equal(evaluateFormula('STDEV.P(-1E308,1E308)'), 1E308, 'standard deviation can remain finite when variance exceeds numeric range');
  assert.equal(evaluateFormula('VAR.P(-1E308,1E308)'), '#NUM!', 'variance reports a result beyond numeric range');
  const statisticsAContext = { getCell: row => [2, true, false, 'text', null][row] ?? null };
  assert.equal(evaluateFormula('VARA(A1:A5)', statisticsAContext), 11 / 12);
  assert.equal(evaluateFormula('VARPA(A1:A5)', statisticsAContext), 11 / 16);
  assert.equal(evaluateFormula('STDEVA(A1:A5)', statisticsAContext), Math.sqrt(11 / 12));
  assert.equal(evaluateFormula('STDEVPA(A1:A5)', statisticsAContext), Math.sqrt(11 / 16));
  assert.equal(evaluateFormula('VARA(2,TRUE,FALSE,"text")'), 11 / 12);
assert.equal(evaluateFormula('VARA(1E308,1E308)'), 0);
assert.equal(evaluateFormula('STDEVA(1E308,1E308)'), 0);
assert.equal(evaluateFormula('STDEVA(-1E308,1E308)'), Math.SQRT2 * 1E308);
assert.equal(evaluateFormula('VARPA(-1E308,1E308)'), '#NUM!');
  assert.equal(evaluateFormula('VARA(A1:A1)', statisticsAContext), '#DIV/0!');
  assert.equal(evaluateFormula('VARPA(A1:A1)', statisticsAContext), 0);
  assert.equal(evaluateFormula('NORM.S.DIST(0,TRUE)'), 0.5);
  assert.ok(Math.abs(evaluateFormula('NORM.S.DIST(1.96,TRUE)') - 0.9750021049) < 1e-7);
  assert.ok(Math.abs(evaluateFormula('NORM.S.DIST(0,FALSE)') - 1 / Math.sqrt(2 * Math.PI)) < 1e-14);
  assert.equal(evaluateFormula('NORM.DIST(0,0,1,TRUE)'), 0.5);
  assert.ok(Math.abs(evaluateFormula('NORM.DIST(0,0,2,FALSE)') - 1 / (2 * Math.sqrt(2 * Math.PI))) < 1e-14);
  const wideNormalDensity = evaluateFormula('NORM.DIST(0,0,1E308,FALSE)');
  assert.ok(Math.abs(wideNormalDensity / 3.989422804014326e-309 - 1) < 1e-12, 'normal density stays nonzero when the direct denominator product overflows');
  assert.ok(Math.abs(evaluateFormula('NORM.DIST(1E308,-1E308,1E308,TRUE)') - 0.977249868) < 1e-6, 'standardized difference avoids overflowing opposite-signed values');
  assert.equal(evaluateFormula('NORM.S.INV(0.5)'), 0);
  assert.ok(Math.abs(evaluateFormula('NORM.S.INV(0.975)') - 1.9599639845) < 1e-7);
  assert.ok(Math.abs(evaluateFormula('NORM.S.DIST(NORM.S.INV(0.975),TRUE)') - 0.975) < 1e-7);
  assert.equal(evaluateFormula('NORM.INV(0.5,10,2)'), 10);
  assert.equal(evaluateFormula('NORM.INV(1E-300,1E308,1E308)'), '#NUM!', 'inverse normal reports a non-finite scaled result as a spreadsheet error');
  assert.equal(evaluateFormula('NORM.S.INV(0)'), '#NUM!');
  assert.equal(evaluateFormula('NORM.INV(0.5,0,0)'), '#NUM!');
  assert.equal(evaluateFormula('NORM.S.DIST(1,1)'), '#VALUE!');
  assert.equal(evaluateFormula('NORM.DIST(1,0,-1,TRUE)'), '#NUM!');
  assert.ok(Math.abs(evaluateFormula('BINOM.DIST(2,4,0.5,FALSE)') - 0.375) < 1e-14);
  assert.ok(Math.abs(evaluateFormula('BINOM.DIST(2,4,0.5,TRUE)') - 0.6875) < 1e-14);
  assert.equal(evaluateFormula('BINOM.DIST(0,10,0,FALSE)'), 1);
  assert.equal(evaluateFormula('BINOM.DIST(10,10,1,TRUE)'), 1);
  const largeBinomialCdf = evaluateFormula('BINOM.DIST(100000,100000,0.5,TRUE)');
  assert.ok(largeBinomialCdf <= 1 && largeBinomialCdf > 1 - 1e-12, 'large binomial CDF stays normalized at its upper endpoint');
  assert.equal(evaluateFormula('BINOM.DIST(2,1,0.5,FALSE)'), '#NUM!');
  assert.equal(evaluateFormula('BINOM.DIST(1.5,4,0.5,FALSE)'), '#NUM!');
  assert.ok(Math.abs(evaluateFormula('POISSON.DIST(3,2,FALSE)') - 0.18044704431548356) < 1e-14);
  assert.ok(Math.abs(evaluateFormula('POISSON.DIST(3,2,TRUE)') - 0.857123460498547) < 1e-14);
  assert.ok(Math.abs(evaluateFormula('POISSON.DIST(100000,100000,FALSE)') - 0.00126156520970530056294685214358) < 1e-16, 'large-mean Poisson PMF uses an accurate mode probability');
  assert.ok(Math.abs(evaluateFormula('POISSON.DIST(100000,100000,TRUE)') - 0.5008410430993401238696342731) < 2e-15, 'large-mean Poisson CDF accumulates probabilities accurately');
  assert.ok(Math.abs(evaluateFormula('POISSON.DIST(99999,100000,TRUE)') - 0.4995794778896348233066874210) < 2e-15, 'large-mean Poisson CDF remains accurate just below the mode');
  assert.equal(evaluateFormula('POISSON.DIST(0,0,TRUE)'), 1);
  assert.equal(evaluateFormula('POISSON.DIST(1,0,FALSE)'), 0);
  assert.equal(evaluateFormula('POISSON.DIST(-1,2,FALSE)'), '#NUM!');
  assert.equal(evaluateFormula('POISSON.DIST(1,100001,TRUE)'), '#NUM!');
  assert.equal(evaluateFormula('RAND()', { random: () => 0.25 }), 0.25);
  assert.equal(evaluateFormula('RAND(1)', { random: () => 0.25 }), '#VALUE!');
  assert.equal(evaluateFormula('RANDBETWEEN(1,6)', { random: () => 0.999 }), 6);
  assert.equal(evaluateFormula('RANDBETWEEN(-3.9,-1.2)', { random: () => 0.5 }), -2, 'RANDBETWEEN truncates fractional bounds');
  assert.equal(evaluateFormula('RANDBETWEEN(5,1)', { random: () => 0.5 }), '#NUM!');
  const randomValues = [0, 0.1, 0.5, 0.99, 0.3, 0.8];
  let randomIndex = 0;
  assert.deepEqual(evaluateFormula('RANDARRAY(2,3,1,10,TRUE)', { random: () => randomValues[randomIndex++], allowArrayResult: true }), [[1, 2, 6], [10, 4, 9]]);
  assert.deepEqual(evaluateFormula('RANDARRAY(2,2)', { random: () => 0.5, allowArrayResult: true }), [[0.5, 0.5], [0.5, 0.5]]);
  assert.equal(evaluateFormula('RANDARRAY(0)', { random: () => 0 }), '#VALUE!');
  assert.equal(evaluateFormula('RANDARRAY(101,100)', { random: () => 0 }), '#NUM!');
  const randomArrayCell = { address: 'A1', row: 0, column: 0, value: '', kind: 'number', formula: 'RANDARRAY(2,1,1,3,TRUE)', hasFormula: true };
  const randomArraySheet = { name: 'Sheet1', mergedCells: [], tables: [], cells: new Map([['A1', randomArrayCell]]) };
  const randomArrayDocument = { state: { formulas: true, definedNames: [], sheets: [randomArraySheet] }, getCell: (row, column) => row === 0 && column === 0 ? randomArrayCell : null };
  const randomArrayCalculator = createFormulaCalculator(randomArrayDocument);
  assert.ok(Number.isInteger(randomArrayCalculator.value(0, 0, 0)));
  assert.ok(Number.isInteger(randomArrayCalculator.value(0, 1, 0)), 'RANDARRAY values spill into worksheet cells');
  const modeContext = { getCell: row => [9, 9, 2, 2, 3][row] ?? null };
  assert.equal(evaluateFormula('MODE.SNGL(A1:A5)', modeContext), 9, 'single mode returns the first tied value');
  assert.deepEqual(evaluateFormula('MODE.MULT(A1:A5)', { ...modeContext, allowArrayResult: true }), [[9], [2]], 'multiple modes spill vertically in first-occurrence order');
  assert.equal(evaluateFormula('MODE.SNGL(A1:A3)', { getCell: row => row + 1 }), '#N/A');
  assert.equal(evaluateFormula('MODE.MULT(A1:A3)', { getCell: row => row + 1, allowArrayResult: true }), '#N/A');
  const frequencyContext = { getCell: (row, column) => ({ '0:0': 1, '1:0': 2, '2:0': 2, '3:0': 3, '4:0': 4, '5:0': 5, '6:0': 'text', '0:1': 2, '1:1': 3, '2:1': 4 })[row + ':' + column] ?? null, allowArrayResult: true };
  assert.deepEqual(evaluateFormula('FREQUENCY(A1:A7,B1:B3)', frequencyContext), [[3], [1], [1], [1]], 'each bin includes its upper boundary and the final bucket counts larger values');
  assert.deepEqual(evaluateFormula('FREQUENCY(A1:A7,C1:C1)', { ...frequencyContext, getCell: (row, column) => column === 2 && row === 0 ? null : frequencyContext.getCell(row, column) }), [[6]], 'an empty bin range returns one overflow bucket');
  assert.equal(evaluateFormula('FREQUENCY(1)'), '#VALUE!');
  const frequencyCells = new Map([
    ['A1', { address: 'A1', row: 0, column: 0, value: '', kind: 'number', formula: 'FREQUENCY(B1:B5,C1:C2)', hasFormula: true }],
    ...[1, 2, 2, 3, 5].map((value, row) => [`B${row + 1}`, { address: `B${row + 1}`, row, column: 1, value: String(value), kind: 'number' }]),
    ...[2, 3].map((value, row) => [`C${row + 1}`, { address: `C${row + 1}`, row, column: 2, value: String(value), kind: 'number' }])
  ]);
  const frequencyDocument = { state: { formulas: true, definedNames: [], sheets: [{ name: 'Sheet1', mergedCells: [], tables: [], cells: frequencyCells }] }, getCell: (row, column) => [...frequencyCells.values()].find(cell => cell.row === row && cell.column === column) || null };
  const frequencyCalculator = createFormulaCalculator(frequencyDocument);
  assert.deepEqual([frequencyCalculator.value(0, 0, 0), frequencyCalculator.value(0, 1, 0), frequencyCalculator.value(0, 2, 0)], [3, 1, 1], 'frequency buckets spill into worksheet cells');
  const multiModeDocument = { state: { formulas: true, definedNames: [], sheets: [{ name: 'Sheet1', mergedCells: [], tables: [], cells: new Map([
    ['A1', { address: 'A1', row: 0, column: 0, value: '', kind: 'number', formula: 'MODE.MULT(B1:B4)', hasFormula: true }],
    ['B1', { address: 'B1', row: 0, column: 1, value: '9', kind: 'number' }],
    ['B2', { address: 'B2', row: 1, column: 1, value: '9', kind: 'number' }],
    ['B3', { address: 'B3', row: 2, column: 1, value: '2', kind: 'number' }],
    ['B4', { address: 'B4', row: 3, column: 1, value: '2', kind: 'number' }]
  ]) }] }, getCell: (row, column) => [...multiModeDocument.state.sheets[0].cells.values()].find(cell => cell.row === row && cell.column === column) || null };
  const multiModeCalculator = createFormulaCalculator(multiModeDocument);
  assert.equal(multiModeCalculator.value(0, 0, 0), 9);
  assert.equal(multiModeCalculator.value(0, 1, 0), 2, 'MODE.MULT results spill into worksheet cells');
  assert.equal(evaluateFormula('VAR.S(A1:A1)', context), '#DIV/0!');
  assert.equal(evaluateFormula('"Jin"&"Sheet"'), 'JinSheet');
  assert.equal(evaluateFormula('ROUND(2.345,2)'), 2.35);
  assert.equal(evaluateFormula('ROUNDUP(1.01,1)'), 1.1);
  assert.equal(evaluateFormula('ROUNDUP(-1.01,1)'), -1.1);
  assert.equal(evaluateFormula('ROUNDUP(1234,-2)'), 1300);
  assert.equal(evaluateFormula('ROUNDDOWN(1.09,1)'), 1);
  assert.equal(evaluateFormula('ROUNDDOWN(-1.09,1)'), -1);
  assert.equal(evaluateFormula('ROUNDDOWN(1234,-2)'), 1200);
  assert.equal(evaluateFormula('CEILING.MATH(4.3)'), 5);
  assert.equal(evaluateFormula('CEILING.MATH(4.3,2)'), 6);
  assert.equal(evaluateFormula('CEILING.MATH(-4.3)'), -4);
  assert.equal(evaluateFormula('CEILING.MATH(-4.3,2,1)'), -6);
  assert.equal(evaluateFormula('CEILING.MATH(1.2,0)'), 0);
  assert.equal(evaluateFormula('CEILING.MATH(-4.3,,1)'), -5);
  assert.equal(evaluateFormula('CEILING.MATH(0.3,0.1)'), 0.3);
  assert.equal(evaluateFormula('CEILING(4.3,2)'), 6);
  assert.equal(evaluateFormula('CEILING(-4.3,-2)'), -6);
  assert.equal(evaluateFormula('FLOOR(4.3,2)'), 4);
  assert.equal(evaluateFormula('FLOOR(-4.3,-2)'), -4);
  assert.equal(evaluateFormula('CEILING(-4.3,2)'), '#NUM!');
  assert.equal(evaluateFormula('FLOOR(4.3,-2)'), '#NUM!');
  assert.equal(evaluateFormula('MROUND(10,3)'), 9);
  assert.equal(evaluateFormula('MROUND(11,3)'), 12);
  assert.equal(evaluateFormula('MROUND(-11,-3)'), -12);
  assert.equal(evaluateFormula('MROUND(10,-3)'), '#NUM!');
  assert.equal(evaluateFormula('MROUND(10,0)'), 0);
  assert.equal(evaluateFormula('FLOOR.MATH(4.3,2)'), 4);
  assert.equal(evaluateFormula('FLOOR.MATH(-4.3,2)'), -6);
  assert.equal(evaluateFormula('FLOOR.MATH(-4.3,2,1)'), -4);
  assert.equal(evaluateFormula('FLOOR.MATH(1.2,0)'), 0);
  assert.equal(evaluateFormula('FLOOR.MATH(0.3,0.1)'), 0.3);
  assert.equal(evaluateFormula('CEILING.MATH(1,2,1,4)'), '#VALUE!');
  assert.equal(evaluateFormula('EXP(1)'), Math.E);
  assert.equal(evaluateFormula('SIN(PI()/2)'), 1);
  assert.equal(evaluateFormula('COS(PI())'), -1);
  assert.equal(evaluateFormula('TAN(0)'), 0);
  assert.equal(evaluateFormula('ASIN(1)'), Math.PI / 2);
  assert.equal(evaluateFormula('ACOS(-1)'), Math.PI);
  assert.equal(evaluateFormula('ATAN(1)'), Math.PI / 4);
  assert.equal(evaluateFormula('ATAN2(1,1)'), Math.PI / 4, 'ATAN2 takes x first and y second');
  assert.equal(evaluateFormula('ATAN2(0,0)'), '#DIV/0!');
  assert.equal(evaluateFormula('DEGREES(PI())'), 180);
  assert.equal(evaluateFormula('RADIANS(180)'), Math.PI);
  assert.equal(evaluateFormula('SINH(0)'), 0);
  assert.equal(evaluateFormula('COSH(0)'), 1);
  assert.equal(evaluateFormula('TANH(0)'), 0);
  assert.equal(evaluateFormula('ASINH(0)'), 0);
  assert.equal(evaluateFormula('ACOSH(1)'), 0);
  assert.equal(evaluateFormula('ATANH(0)'), 0);
  assert.equal(evaluateFormula('ASIN(2)'), '#NUM!');
  assert.equal(evaluateFormula('ACOSH(0.5)'), '#NUM!');
  assert.equal(evaluateFormula('ATANH(1)'), '#NUM!');
  assert.equal(evaluateFormula('SIN(1,2)'), '#VALUE!');
  assert.equal(evaluateFormula('FACT(5)'), 120);
  assert.equal(evaluateFormula('FACT(5.9)'), 120, 'FACT truncates fractional inputs');
  assert.equal(evaluateFormula('FACT(171)'), '#NUM!');
  assert.equal(evaluateFormula('FACT(-1)'), '#NUM!');
  assert.equal(evaluateFormula('FACTDOUBLE(7)'), 105);
  assert.equal(evaluateFormula('FACTDOUBLE(8)'), 384);
  assert.equal(evaluateFormula('COMBIN(5,2)'), 10);
  assert.equal(evaluateFormula('COMBIN(5.9,2.9)'), 10, 'COMBIN truncates fractional inputs');
  assert.equal(evaluateFormula('COMBIN(2,3)'), '#NUM!');
  assert.equal(evaluateFormula('PERMUT(5,2)'), 20);
  assert.equal(evaluateFormula('GCD(24,36)'), 12);
  assert.equal(evaluateFormula('GCD(-24,36.9)'), 12);
  assert.equal(evaluateFormula('GCD(A1:A2)', context), 2, 'GCD accepts a sparse cell range');
  assert.equal(evaluateFormula('LCM(4,6)'), 12);
  assert.equal(evaluateFormula('LCM(0,6)'), 0);
  assert.equal(evaluateFormula('MULTINOMIAL(1,2,3)'), 60);
  assert.equal(evaluateFormula('LCM(9007199254740991,2)'), '#NUM!');
  assert.equal(evaluateFormula('MULTINOMIAL(1,-1)'), '#NUM!');
  assert.equal(evaluateFormula('LN(EXP(1))'), 1);
  assert.equal(evaluateFormula('LN(0)'), '#NUM!');
  assert.equal(evaluateFormula('LOG(100)'), 2);
  assert.equal(evaluateFormula('LOG(8,2)'), 3);
  assert.equal(evaluateFormula('LOG(10,1)'), '#DIV/0!');
  assert.equal(evaluateFormula('LOG10(1000)'), 3);
  assert.equal(evaluateFormula('PI()'), Math.PI);
  assert.equal(evaluateFormula('TRUNC(-8.9)'), -8);
  assert.equal(evaluateFormula('TRUNC(1234,-2)'), 1200);
  assert.equal(evaluateFormula('TRUNC(1.239,2)'), 1.23);
  assert.equal(evaluateFormula('TRUNC(1,101)'), '#NUM!');
  assert.equal(evaluateFormula('ROUNDUP(1)'), '#VALUE!');
  assert.equal(evaluateFormula('ROUNDDOWN(1,101)'), '#NUM!');
  assert.equal(evaluateFormula('IF(1=1,"yes",1/0)'), 'yes');
  assert.equal(evaluateFormula('IF(FALSE,1/0)'), false);
  assert.equal(evaluateFormula('CHOOSE(2,"red","blue","green")'), 'blue');
  assert.equal(evaluateFormula('CHOOSE(2.9,"red","blue")'), 'blue', 'CHOOSE truncates its index to an integer');
  assert.equal(evaluateFormula('CHOOSE(0,"red")'), '#VALUE!');
  assert.equal(evaluateFormula('IFS(1=0,"no",2>1,"yes",TRUE,"fallback")'), 'yes');
  assert.equal(evaluateFormula('IFS(FALSE,"no")'), '#N/A');
  assert.equal(evaluateFormula('IFS(TRUE,"yes",FALSE)'), '#VALUE!');
  assert.equal(evaluateFormula('SWITCH("Green","red",1,"green",2,0)'), 2);
  assert.equal(evaluateFormula('SWITCH(4,1,"one",2,"two")'), '#N/A');
  assert.equal(evaluateFormula('SWITCH(4,1,"one",2,"two",0)'), 0);
  assert.equal(evaluateFormula('IFERROR(1/0,9)'), 9);
  assert.equal(evaluateFormula('IFNA(MATCH(999,A1:A1,0),"missing")', { getCell: () => 1 }), 'missing');
  assert.equal(evaluateFormula('IFNA(1/0,"missing")'), '#DIV/0!', 'IFNA handles only #N/A and leaves other errors unchanged');
  assert.equal(evaluateFormula('IFNA(7,"missing")'), 7);
  assert.equal(evaluateFormula('IFNA(1)'), '#VALUE!');
  assert.equal(evaluateFormula('AND(TRUE,1>0)'), true);
assert.equal(evaluateFormula('OR(FALSE,0,TRUE)'), true);
assert.equal(evaluateFormula('XOR(TRUE,FALSE)'), true);
assert.equal(evaluateFormula('XOR(TRUE,1,0)'), false);
assert.equal(evaluateFormula('XOR(TRUE,FALSE,TRUE)'), false);
assert.equal(evaluateFormula('XOR(TRUE,1/0)'), '#DIV/0!');
assert.equal(evaluateFormula('XOR()'), '#VALUE!');
assert.equal(evaluateFormula('NOT(FALSE)'), true);
  const leapDay = evaluateFormula('DATE(2024,2,29)');
  assert.equal(evaluateFormula(`YEAR(${leapDay})`), 2024);
  assert.equal(evaluateFormula(`MONTH(${leapDay})`), 2);
  assert.equal(evaluateFormula(`DAY(${leapDay})`), 29);
  assert.equal(evaluateFormula('DATE(1900,2,29)'), 60);
  assert.equal(evaluateFormula('DATE(1900,3,1)'), 61);
  assert.equal(evaluateFormula('TODAY()', { todaySerial: 45351 }), 45351);
  assert.equal(evaluateFormula('HOUR(TIME(25,2,3))'), 1);
  assert.equal(evaluateFormula('MINUTE(TIME(25,2,3))'), 2);
  assert.equal(evaluateFormula('SECOND(TIME(25,2,3))'), 3);
  assert.equal(evaluateFormula('COUNTIF(A1:B2,">3")', context), 2);
  assert.equal(evaluateFormula('SUMIF(A1:B2,">3")', context), 9);
  assert.equal(evaluateFormula('AVERAGEIF(A1:B2,">3")', context), 4.5);
  const conditionContext = { getCell: (row, column) => [[1, 5, 'apple'], [2, '7', 'apricot'], [0, 9, 'pear']][row]?.[column] ?? null };
  assert.equal(evaluateFormula('SUMIF(A1:A3,">0",B1:B3)', conditionContext), 5, 'SUMIF ignores text cells in the sum range');
  const mismatchedSumContext = { getCell: (row, column) => [[10, 1], [20, 0], [30, 1]][row]?.[column] ?? null };
  assert.equal(evaluateFormula('SUMIF(B1:B3,1,A1:A2)', mismatchedSumContext), 10, 'SUMIF aligns a smaller sum range from its top-left cell and treats missing aligned cells as zero');
  assert.equal(evaluateFormula('SUMIF(B1:B2,1,A1:A3)', mismatchedSumContext), 10, 'SUMIF ignores sum-range cells outside the criteria range dimensions');
  assert.equal(evaluateFormula('COUNTIF(C1:C3,"ap*")', conditionContext), 2);
  const criterionTypeContext = { getCell: (row, column) => [[null], [0], [''], [false], ['0'], ['text']][row]?.[column] ?? null };
  assert.equal(evaluateFormula('COUNTIF(A1:A6,"0")', criterionTypeContext), 3, 'numeric zero criteria use spreadsheet comparison semantics for numeric, logical, and numeric-text cells');
  assert.equal(evaluateFormula('COUNTIF(A1:A6,"=")', criterionTypeContext), 2, 'an empty criterion matches physically blank and empty-text cells');
  assert.equal(evaluateFormula('COUNTIF(A1:A6,"<>0")', criterionTypeContext), 3, 'not-equal criteria include blank and empty-text cells');
  const escapedWildcardContext = { getCell: (row, column) => [['a*'], ['abc'], ['a?']][row]?.[column] ?? null };
  assert.equal(evaluateFormula('COUNTIF(A1:A3,"a~*")', escapedWildcardContext), 1, 'tilde escapes a wildcard in criteria');
  const wildcardTypeContext = { getCell: (row, column) => [['#N/A'], [42], [true], ['text']][row]?.[column] ?? null };
  assert.equal(evaluateFormula('COUNTIF(A1:A4,"*")', wildcardTypeContext), 1, 'wildcard criteria match text only, not errors, numbers, or logical values');
  assert.equal(evaluateFormula('COUNTIF(A1:A4,"<>*")', wildcardTypeContext), 2, 'negative wildcard criteria count nonblank numeric and logical values');
  const multiConditionContext = { getCell: (row, column) => [
    ['East', 'Apples', 100], ['East', 'Pears', 200], ['West', 'Apples', 150], ['East', 'Apples', 'text']
  ][row]?.[column] ?? null };
  assert.equal(evaluateFormula('COUNTIFS(A1:A4,"East",B1:B4,"Apples")', multiConditionContext), 2);
  assert.equal(evaluateFormula('SUMIFS(C1:C4,A1:A4,"East",B1:B4,"Apples")', multiConditionContext), 100);
  assert.equal(evaluateFormula('AVERAGEIFS(C1:C4,A1:A4,"East",B1:B4,"Apples")', multiConditionContext), 100);
  assert.equal(evaluateFormula('COUNTIFS(A1:A4,"East",B1:B3,"Apples")', multiConditionContext), '#VALUE!');
  const overflowingCriteriaContext = { getCell: (row, column) => [[1, 1E308], [1, 1E308]][row]?.[column] ?? null };
  for (const formula of [
    'SUMIF(A1:A2,">0",B1:B2)', 'AVERAGEIF(A1:A2,">0",B1:B2)',
    'SUMIFS(B1:B2,A1:A2,">0")', 'AVERAGEIFS(B1:B2,A1:A2,">0")',
    'SUMIF(HSTACK(1,1),">0",HSTACK(1E308,1E308))', 'AVERAGEIF(HSTACK(1,1),">0",HSTACK(1E308,1E308))',
    'SUMIFS(HSTACK(1E308,1E308),HSTACK(1,1),">0")', 'AVERAGEIFS(HSTACK(1E308,1E308),HSTACK(1,1),">0")'
  ]) assert.equal(evaluateFormula(formula, overflowingCriteriaContext), '#NUM!', `${formula} reports aggregate overflow`);
  const productContext = { getCell: (row, column) => [[1, 5], [2, 7], ['text', 9]][row]?.[column] ?? null };
  assert.equal(evaluateFormula('SUMPRODUCT(A1:A2,B1:B2)', productContext), 19);
  assert.equal(evaluateFormula('SUMPRODUCT(A1:B2)', productContext), 15);
  assert.equal(evaluateFormula('SUMPRODUCT(A1:A2,B1:B3)', productContext), '#VALUE!');
  assert.equal(evaluateFormula('SUMPRODUCT(A1:A2,1)', productContext), '#VALUE!');
  const productErrorContext = { getCell: (row, column) => row === 1 ? '#N/A' : 2 };
  assert.equal(evaluateFormula('SUMPRODUCT(A1:A2,B1:B2)', productErrorContext), '#N/A');
  const tableContext = { getCell: (row, column) => [[101, 'Ada'], [102, 'Bo'], [103, 'Cy']][row]?.[column] ?? null };
  assert.equal(evaluateFormula('VLOOKUP(102,A1:B3,2,FALSE)', tableContext), 'Bo');
  assert.equal(evaluateFormula('VLOOKUP(102,A1:B3,2)', tableContext), 'Bo');
  assert.equal(evaluateFormula('HLOOKUP(102,A1:C2,2,FALSE)', { getCell: (row, column) => [[101, 102, 103], ['Ada', 'Bo', 'Cy']][row]?.[column] ?? null }), 'Bo');
  const duplicateApproximateContext = { getCell: (row, column) => [[1, 'one'], [2, 'first'], [2, 'last'], [3, 'three']][row]?.[column] ?? null };
  assert.equal(evaluateFormula('MATCH(2,A1:A4,1)', duplicateApproximateContext), 3, 'ascending approximate MATCH returns the last duplicate key');
  assert.equal(evaluateFormula('VLOOKUP(2,A1:B4,2)', duplicateApproximateContext), 'last', 'approximate VLOOKUP returns the last duplicate key');
  assert.equal(evaluateFormula('VLOOKUP(2,A1:B4,2,FALSE)', duplicateApproximateContext), 'first', 'exact VLOOKUP retains its first-match behavior');
  assert.equal(evaluateFormula('HLOOKUP(2,A1:D2,2)', { getCell: (row, column) => [[1, 2, 2, 3], ['one', 'first', 'last', 'three']][row]?.[column] ?? null }), 'last', 'approximate HLOOKUP returns the last duplicate key');
  assert.equal(evaluateFormula('MATCH(2,A1:A4,-1)', { getCell: row => [3, 2, 2, 1][row] }), 3, 'descending approximate MATCH returns the last duplicate key');
assert.equal(evaluateFormula('LOOKUP(102,A1:A3)', tableContext), 102, 'LOOKUP returns the matched item when no result vector is supplied');
assert.equal(evaluateFormula('LOOKUP(104,A1:A3,B1:B3)', tableContext), 'Cy', 'LOOKUP uses the greatest ascending key not greater than the target');
assert.equal(evaluateFormula('LOOKUP(100,A1:A3,B1:B3)', tableContext), '#N/A', 'LOOKUP returns #N/A below the smallest key');
assert.equal(evaluateFormula('LOOKUP(2,A1:A4,B1:B4)', { getCell: (row, column) => [[1, 'A'], [2, 'B'], [2, 'C'], [3, 'D']][row]?.[column] ?? null }), 'C', 'LOOKUP returns the final result for a repeated matched key');
assert.equal(evaluateFormula('LOOKUP("d",A1:A3,B1:B3)', { getCell: (row, column) => [['a', 1], ['c', 2], ['e', 3]][row]?.[column] ?? null }), 2, 'LOOKUP supports sorted text vectors');
assert.equal(evaluateFormula('LOOKUP(102,A1:A3,B1:B3)', { getCell: row => row === 0 ? '#N/A' : [101, 103][row - 1] }), 101, 'LOOKUP skips an error before numeric keys in the lookup vector');
assert.equal(evaluateFormula('MATCH(2,A1:A3,1)', { getCell: row => ['#N/A', 1, 3][row] }), 2, 'approximate MATCH skips an error before sorted numeric values');
assert.equal(evaluateFormula('VLOOKUP(2,A1:B3,2,TRUE)', { getCell: (row, column) => [['#N/A', 'error'], [1, 'one'], [3, 'three']][row]?.[column] ?? null }), 'one', 'approximate VLOOKUP skips an error before its lookup keys');
assert.equal(evaluateFormula('XLOOKUP(2,A1:A3,B1:B3,"missing",-1)', { getCell: (row, column) => [['#N/A', 'error'], [1, 'one'], [3, 'three']][row]?.[column] ?? null }), 'one', 'approximate XLOOKUP skips an error before its lookup keys');
assert.equal(evaluateFormula('LOOKUP(1,A1:A3,B1:B2)', tableContext), '#VALUE!', 'LOOKUP vectors must have matching lengths');
assert.equal(evaluateFormula('LOOKUP(1,A1:B2,B1:B2)', tableContext), '#VALUE!', 'LOOKUP vectors must be one-dimensional');
  assert.equal(evaluateFormula('MATCH(102,A1:A3,0)', tableContext), 2);
  assert.equal(evaluateFormula('INDEX(B1:B3,2)', tableContext), 'Bo');
  assert.equal(evaluateFormula('VLOOKUP(999,A1:B3,2,FALSE)', tableContext), '#N/A');
  assert.equal(evaluateFormula('XLOOKUP(102,A1:A3,B1:B3)', tableContext), 'Bo');
  assert.equal(evaluateFormula('XLOOKUP(2,B1,A1)', { getCell: (row, column) => row === 0 ? column === 0 ? 9 : column === 1 ? 2 : null : null }), 9, 'single-cell lookup and return references remain valid');
  assert.equal(evaluateFormula('XLOOKUP(104,A1:A3,B1:B3,"missing",-1)', tableContext), 'Cy');
  assert.equal(evaluateFormula('XLOOKUP(100,A1:A3,B1:B3,"missing",1)', tableContext), 'Ada');
  assert.equal(evaluateFormula('XLOOKUP(999,A1:A3,B1:B3,"missing")', tableContext), 'missing');
assert.equal(evaluateFormula('XLOOKUP(999,A1:A3,B1:B3)', tableContext), '#N/A');
assert.equal(evaluateFormula('XLOOKUP(102,A1:A3,B1:B3,"missing",0,2)', tableContext), 'Bo', 'XLOOKUP binary search supports ascending arrays');
assert.equal(evaluateFormula('XLOOKUP(102,A1:A3,B1:B3,"missing",0,-2)', { getCell: (row, column) => [[103, 'Cy'], [102, 'Bo'], [101, 'Ada']][row]?.[column] ?? null }), 'Bo', 'XLOOKUP binary search supports descending arrays');
assert.equal(evaluateFormula('XLOOKUP(100,A1:A3,B1:B3,"missing",1,2)', tableContext), 'Ada');
assert.equal(evaluateFormula('XLOOKUP(104,A1:A3,B1:B3,"missing",-1,-2)', { getCell: (row, column) => [[103, 'Cy'], [102, 'Bo'], [101, 'Ada']][row]?.[column] ?? null }), 'Cy');
assert.equal(evaluateFormula('XLOOKUP(102,A1:A3,B1:B3,"missing",0,7)'), '#VALUE!');
const wildcardContext = { getCell: (row, column) => [['apple', 'apricot', 'pear'][row], ['A', 'B', 'C'][row]][column] ?? null };
  assert.equal(evaluateFormula('XLOOKUP("ap*",A1:A3,B1:B3,"missing",2)', wildcardContext), 'A');
  const tildeWildcardContext = { getCell: row => ['apple', '~apple', '~apricot'][row] ?? null };
  assert.equal(evaluateFormula('MATCH("~ap*",A1:A3,0)', tildeWildcardContext), 2, 'tilde before an ordinary letter remains a literal tilde in wildcard patterns');
  assert.equal(evaluateFormula('XMATCH("~ap*",A1:A3,2)', tildeWildcardContext), 2, 'XMATCH preserves literal tildes while matching wildcards');
  assert.equal(evaluateFormula('COUNTIF(A1:A3,"~ap*")', tildeWildcardContext), 2, 'criteria wildcard matching treats a tilde before ordinary text literally');
  assert.equal(evaluateFormula('XLOOKUP("~ap*",A1:A3,A1:A3,"missing",2)', tildeWildcardContext), '~apple', 'XLOOKUP wildcard mode preserves literal tildes');
  const duplicateContext = { getCell: (row, column) => [["A", 1], ["A", 2]][row]?.[column] ?? null };
  assert.equal(evaluateFormula('XLOOKUP("A",A1:A2,B1:B2,"missing",0,-1)', duplicateContext), 2);
  assert.deepEqual(evaluateFormula('XLOOKUP(2,A1:A3,B1:C3)', { getCell: (row, column) => [[1, 'One', 10], [2, 'Two', 20], [3, 'Three', 30]][row]?.[column] ?? null, allowArrayResult: true }), [['Two', 20]], 'vertical lookup returns the full matching row');
  assert.deepEqual(evaluateFormula('XLOOKUP(2,A1:C1,A2:C3)', { getCell: (row, column) => [[1, 2, 3], ['A', 'B', 'C'], [10, 20, 30]][row]?.[column] ?? null, allowArrayResult: true }), [['B'], [20]], 'horizontal lookup returns the full matching column');
  assert.equal(evaluateFormula('XLOOKUP(2,A1:A2,B1:C3)', { getCell: (row, column) => [[1, 2, 3], [2, 4, 5], [3, 6, 7]][row]?.[column] ?? null }), '#VALUE!', 'lookup and return ranges must align on the searched dimension');
  assert.equal(evaluateFormula('LEFT("JinSheet",3)'), 'Jin');
  assert.equal(evaluateFormula('RIGHT("JinSheet",5)'), 'Sheet');
  assert.equal(evaluateFormula('MID("스프레드시트",2,3)'), '프레드');
  assert.equal(evaluateFormula('LEN("😀a")'), 2);
  assert.equal(evaluateFormula('TRIM("  Jin   Sheet  ")'), 'Jin Sheet');
  assert.equal(evaluateFormula('LOWER("JinSHEET")'), 'jinsheet');
  assert.equal(evaluateFormula('UPPER("JinSheet")'), 'JINSHEET');
  assert.equal(evaluateFormula('CONCAT("Jin",TRUE,1)'), 'JinTRUE1');
  assert.equal(evaluateFormula('CONCATENATE("Jin","Sheet")'), 'JinSheet');
  assert.equal(evaluateFormula('TEXTJOIN("-",TRUE,"a","","b")'), 'a-b');
  assert.equal(evaluateFormula('TEXTJOIN("/",FALSE,A1:A3)', { getCell: (row, column) => [null, 'x', 2][row] }), '/x/2');
  assert.equal(evaluateFormula('TEXTJOIN(",",1,"a",FALSE)'), 'a,FALSE');
  assert.equal(evaluateFormula('TEXTJOIN(",",TRUE,"a",1/0)'), '#DIV/0!');
  assert.equal(evaluateFormula('TEXTJOIN(",",TRUE)'), '#VALUE!');
  assert.equal(evaluateFormula('TEXTJOIN("",TRUE,"' + 'x'.repeat(32768) + '")'), '#VALUE!');
  assert.equal(evaluateFormula('TEXTJOIN("",TRUE,"' + '😀'.repeat(16384) + '")').length, 32768);
  assert.equal(evaluateFormula('LEFT("text",-1)'), '#VALUE!');
  assert.equal(evaluateFormula('MID("text",0,2)'), '#VALUE!');
  assert.equal(evaluateFormula('RIGHT("text",1.5)'), 't');
  assert.equal(evaluateFormula('MID("text",2.9,1.8)'), 'e');
  assert.equal(evaluateFormula('CONCAT("ok",1/0)'), '#DIV/0!');
  assert.equal(evaluateFormula('1/0'), '#DIV/0!');
  assert.equal(evaluateFormula('UNKNOWN(1)'), '#NAME?');
});

test('XMATCH 근사 이진 검색은 오름차순·내림차순 경계와 중복 키에서 선형 결과를 유지한다', () => {
  const sortedArrays = (prefix, remaining, minimum = 0) => {
    if (!remaining) return [prefix];
    const arrays = [];
    for (let value = minimum; value <= 3; value++) arrays.push(...sortedArrays([...prefix, value], remaining - 1, value));
    return arrays;
  };
  for (let length = 1; length <= 7; length++) for (const ascendingValues of sortedArrays([], length)) {
    for (const ascending of [true, false]) {
      const values = ascending ? ascendingValues : [...ascendingValues].reverse();
      const context = { getCell: row => values[row] ?? null };
      const binarySearchMode = ascending ? 2 : -2;
      for (let target = -1; target <= 4; target++) for (const matchMode of [-1, 1]) {
        const matches = values.filter(value => matchMode === -1 ? value <= target : value >= target);
        const expected = matches.length ? (matchMode === -1 ? Math.max(...matches) : Math.min(...matches)) : '#N/A';
        for (const searchMode of [1, -1, binarySearchMode]) {
          const result = evaluateFormula(`XMATCH(${target},A1:A${length},${matchMode},${searchMode})`, context);
          const actual = Number.isInteger(result) && result > 0 ? values[result - 1] : result;
          assert.equal(actual, expected, `${values} target=${target} match=${matchMode} search=${searchMode}`);
        }
      }
    }
  }
});

test('수식 계산기는 통합문서 시트명을 기준으로 공백·아포스트로피가 있는 교차 시트 참조를 계산한다', () => {
  const cell = (address, row, column, value, formula = null, kind = 'number') => ({ address, row, column, value: String(value), cachedValue: value, formula, kind });
  const sheets = [
    { name: 'Data', cells: new Map([
      ['A1', cell('A1', 0, 0, 0, "'Sales O''Brien'!A1*2")],
      ['B1', cell('B1', 0, 1, 0, "SUM('Sales O''Brien'!A1:A2)")],
      ['C1', cell('C1', 0, 2, 0, 'Missing!A1')],
      ['D1', cell('D1', 0, 3, 0, 'IF(FALSE,D1,7)')],
      ['E2', cell('E2', 1, 4, 0, 'ROW()+COLUMN()+ROWS(A1:B2)+COLUMNS(A:C)')],
      ['F2', cell('F2', 1, 5, '', 'ADDRESS(ROW(),COLUMN())')]
    ]) },
    { name: "Sales O'Brien", cells: new Map([
      ['A1', cell('A1', 0, 0, 21)], ['A2', cell('A2', 1, 0, 9)]
    ]) }
  ];
  const document = {
    state: { sheets },
    getCell(row, column, sheetIndex) {
      let letters = '';
      for (let value = column + 1; value > 0; value = Math.floor((value - 1) / 26)) letters = String.fromCharCode(65 + (value - 1) % 26) + letters;
      return sheets[sheetIndex]?.cells.get(letters + (row + 1)) || null;
    }
  };
  const calculator = createFormulaCalculator(document);
  assert.equal(calculator.value(0, 0, 0), 42);
  assert.equal(calculator.value(0, 0, 1), 30);
  assert.equal(calculator.value(0, 0, 2), '#REF!');
  assert.equal(calculator.value(0, 0, 3), 7, 'an unselected IF branch does not leak its circular-reference error');
  assert.equal(calculator.value(0, 1, 4), 12, 'ROW/COLUMN use the formula cell and ROWS/COLUMNS report range dimensions');
  assert.equal(calculator.value(0, 1, 5), '$F$2', 'ADDRESS composes the current formula cell coordinates');
});

test('수식 계산기는 Excel 표의 열·헤더·데이터·현재 행 구조화 참조를 계산한다', () => {
  const cells = new Map();
  const add = (address, row, column, value, kind = 'number', formula = null) => cells.set(address, { address, row, column, value: String(value), kind, formula, hasFormula: Boolean(formula) });
  add('A1', 0, 0, 'Name', 'string'); add('B1', 0, 1, 'Amount', 'string'); add('C1', 0, 2, 'Result', 'string');
  add('A2', 1, 0, 'Alpha', 'string'); add('B2', 1, 1, 10); add('C2', 1, 2, '0', 'number', '[@Amount]');
  add('A3', 2, 0, 'Beta', 'string'); add('B3', 2, 1, 20); add('C3', 2, 2, '0', 'number', 'Sales[[#This Row],[Amount]]');
  add('C4', 3, 2, '0', 'number', '[@Amount]');
  add('C5', 4, 2, '0', 'number', 'SUM(Sales[[#All],[Amount]])');
  add('C6', 5, 2, '0', 'number', 'SUM(Sales[Missing])');
  add('C7', 6, 2, '0', 'number', 'INDEX(Sales[Amount],2)');
  add('D2', 1, 3, '0', 'number', 'SUM(Sales[Amount])');
  add('D3', 2, 3, '0', 'number', 'COUNTA(Sales[#Headers])');
  add('E1', 0, 4, '0', 'number', 'SUM(Sales[[Amount]:[Result]])');
  add('E2', 1, 4, '0', 'number', 'SUM(Sales[[#Data],[Amount]:[Result]])');
  add('E3', 2, 4, '0', 'number', 'SUM(Sales[@[Amount]:[Result]])');
  const document = { state: { formulas: true, definedNames: [], sheets: [{ name: 'Data', cells, tables: [{ name: 'Sales', range: { top: 0, bottom: 2, left: 0, right: 2 }, columns: ['Name', 'Amount', 'Result'], headerRowCount: 1, totalsRowCount: 0 }] }] }, getCell: (row, column) => [...cells.values()].find(cell => cell.row === row && cell.column === column) || null };
  const calculator = createFormulaCalculator(document);
  assert.equal(calculator.value(0, 1, 2), 10);
  assert.equal(calculator.value(0, 2, 2), 20);
  assert.equal(calculator.value(0, 1, 3), 30);
  assert.equal(calculator.value(0, 2, 3), 3);
  assert.equal(calculator.value(0, 3, 2), '#VALUE!', 'a this-row reference outside the table data is invalid');
  assert.equal(calculator.value(0, 4, 2), 30);
  assert.equal(calculator.value(0, 5, 2), '#VALUE!');
  assert.equal(calculator.value(0, 6, 2), 20);
  assert.equal(calculator.value(0, 0, 4), 60, 'a contiguous structured column range is passed to aggregate functions');
  assert.equal(calculator.value(0, 1, 4), 60, 'a data selector can scope a structured column range');
  assert.equal(calculator.value(0, 2, 4), 40, 'a this-row structured column range returns the cells from the formula row');
});

test('INDIRECT는 A1·R1C1 셀 참조와 범위를 해석하고 배열 결과를 spill한다', () => {
  const cells = new Map([
    ['A1', { address: 'A1', row: 0, column: 0, value: '1', kind: 'number' }],
    ['B1', { address: 'B1', row: 0, column: 1, value: '2', kind: 'number' }],
    ['A2', { address: 'A2', row: 1, column: 0, value: '3', kind: 'number' }],
    ['B2', { address: 'B2', row: 1, column: 1, value: '4', kind: 'number' }],
    ['D2', { address: 'D2', row: 1, column: 3, value: '8', kind: 'number' }],
    ['E1', { address: 'E1', row: 0, column: 4, value: '', kind: 'number', formula: 'INDIRECT("A1:B2")', hasFormula: true }]
  ]);
  const document = { state: { formulas: true, sheets: [{ name: 'Data', cells, mergedCells: [], tables: [] }, { name: 'Other Sheet', cells: new Map([['B2', { address: 'B2', row: 1, column: 1, value: '22', kind: 'number' }]]), mergedCells: [], tables: [] }] }, getCell: (row, column, sheetIndex = 0) => [...document.state.sheets[sheetIndex].cells.values()].find(cell => cell.row === row && cell.column === column) || null };
  const calculator = createFormulaCalculator(document);
  assert.equal(evaluateFormula('INDIRECT("A1")', { getCell: (row, column) => row === 0 && column === 0 ? 1 : 0 }), 1);
  assert.equal(evaluateFormula('SUM(INDIRECT("A1:A2"))', { getCell: (row, column) => row < 2 && column === 0 ? row + 1 : 0 }), 3);
  assert.equal(evaluateFormula('INDIRECT("R[-1]C[1]",FALSE)', { formulaRow: 2, formulaColumn: 2, getCell: (row, column) => row === 1 && column === 3 ? 8 : 0 }), 8);
  assert.equal(evaluateFormula('INDIRECT("R1C2",FALSE)', { formulaRow: 2, formulaColumn: 2, getCell: (row, column) => row === 0 && column === 1 ? 2 : 0 }), 2);
  assert.equal(evaluateFormula('INDIRECT("[Book.xlsx]Sheet1!A1")'), '#REF!');
  assert.equal(evaluateFormula('INDIRECT("A1","bad")'), '#VALUE!');
  assert.equal(calculator.cell(0, 0, 4).value, '1');
  assert.deepEqual([[calculator.cell(0, 0, 4).value, calculator.cell(0, 0, 5).value], [calculator.cell(0, 1, 4).value, calculator.cell(0, 1, 5).value]], [['1', '2'], ['3', '4']]);
});

test('OFFSET는 A1 참조에서 이동한 셀·범위를 계산하고 배열 결과를 spill한다', () => {
  const values = new Map([['0:0', 1], ['0:1', 2], ['1:0', 3], ['1:1', 4], ['2:0', 5]]);
  const context = { getCell: (row, column) => values.get(row + ':' + column) ?? null };
  assert.equal(evaluateFormula('OFFSET(A1,1,1)', context), 4);
  assert.equal(evaluateFormula('SUM(OFFSET(A1,1,0,2,1))', context), 8);
  assert.deepEqual(evaluateFormula('OFFSET(A1,0,0,2,2)', { ...context, allowArrayResult: true }), [[1, 2], [3, 4]]);
  assert.deepEqual(evaluateFormula('OFFSET(A1,0,0,,2)', { ...context, allowArrayResult: true }), [[1, 2]]);
  assert.equal(evaluateFormula('OFFSET(A1,-1,0)', context), '#REF!');
  assert.equal(evaluateFormula('OFFSET(A1,0,0,0)', context), '#VALUE!');
  assert.equal(evaluateFormula('OFFSET(A1,0,0,101,100)', context), '#NUM!');
  assert.equal(evaluateFormula('OFFSET(1,0,0)', context), '#VALUE!');
  const cells = new Map([['A1', { address: 'A1', row: 0, column: 0, value: '1', kind: 'number' }], ['B1', { address: 'B1', row: 0, column: 1, value: '2', kind: 'number' }], ['A2', { address: 'A2', row: 1, column: 0, value: '3', kind: 'number' }], ['B2', { address: 'B2', row: 1, column: 1, value: '4', kind: 'number' }], ['D1', { address: 'D1', row: 0, column: 3, value: '', kind: 'number', formula: 'OFFSET(A1,0,0,2,2)', hasFormula: true }]]);
  const document = { state: { formulas: true, sheets: [{ name: 'Data', cells, mergedCells: [], tables: [] }] }, getCell: (row, column) => [...cells.values()].find(cell => cell.row === row && cell.column === column) || null };
  const calculator = createFormulaCalculator(document);
  assert.deepEqual([[calculator.cell(0, 0, 3).value, calculator.cell(0, 0, 4).value], [calculator.cell(0, 1, 3).value, calculator.cell(0, 1, 4).value]], [['1', '2'], ['3', '4']]);
});

test('TEXTBEFORE와 TEXTAFTER는 구분자 출현 횟수·방향·대소문자·대체값을 처리한다', () => {
  assert.equal(evaluateFormula('TEXTBEFORE("alpha-beta-gamma","-")'), 'alpha');
  assert.equal(evaluateFormula('TEXTBEFORE("alpha-beta-gamma","-",2)'), 'alpha-beta');
  assert.equal(evaluateFormula('TEXTBEFORE("alpha-beta-gamma","-",-1)'), 'alpha-beta');
  assert.equal(evaluateFormula('TEXTAFTER("alpha-beta-gamma","-",2)'), 'gamma');
  assert.equal(evaluateFormula('TEXTAFTER("alpha-beta-gamma","-",-2)'), 'beta-gamma');
  assert.equal(evaluateFormula('TEXTBEFORE("alphaXbeta","x",1,1)'), 'alpha');
  assert.equal(evaluateFormula('TEXTBEFORE("alpha","!",1,0,1)'), 'alpha');
  assert.equal(evaluateFormula('TEXTAFTER("alpha","!",1,0,1)'), '');
  assert.equal(evaluateFormula('TEXTBEFORE("alpha","!")'), '#N/A');
  assert.equal(evaluateFormula('TEXTAFTER("alpha","!",1,0,0,"missing")'), 'missing');
  assert.equal(evaluateFormula('TEXTBEFORE("aaa","aa",2)'), '#N/A', 'delimiter occurrences do not overlap');
  assert.equal(evaluateFormula('TEXTAFTER("한글😀끝","😀")'), '끝');
  assert.equal(evaluateFormula('TEXTBEFORE("alpha","",1)'), '#VALUE!');
  assert.equal(evaluateFormula('TEXTAFTER("alpha","-",0)'), '#VALUE!');
});

test('LibreOffice Calc 기준 파일의 wildcard 조건부 집계를 비교한다', () => {
  const fs = require('node:fs');
  const path = require('node:path');
  const reference = analyze(fs.readFileSync(path.join(__dirname, 'fixtures', 'LibreOffice-Criteria-Functions-Reference.xlsx')));
  assert.equal(reference.writable, true, reference.readOnlyReason);
  const cells = reference.sheets[0].cells;
  const cellValue = (row, column) => {
    let letters = '';
    for (let value = column + 1; value > 0; value = Math.floor((value - 1) / 26)) letters = String.fromCharCode(65 + (value - 1) % 26) + letters;
    const cell = cells.get(letters + (row + 1));
    if (!cell) return null;
    if (cell.kind === 'number') return Number(cell.value);
    if (cell.kind === 'boolean') return /^TRUE$/i.test(String(cell.value));
    return cell.value;
  };
  const context = { getCell: cellValue };
  for (const [address, expected] of [['D1', 30], ['D2', 1], ['D3', 20], ['D4', 1], ['D5', 0]]) {
    const cell = cells.get(address);
    assert.ok(cell?.formula, `${address} retains its reference formula`);
    assert.equal(Number(cell.cachedValue), expected, `${address} is the LibreOffice Calc cached result`);
    assert.equal(evaluateFormula(cell.formula, context), expected, `${cell.formula} matches the LibreOffice reference`);
  }
});

test('LibreOffice Calc 기준 파일로 lookup 중복 키와 오류 셀을 대조한다', () => {
  const fs = require('node:fs'), path = require('node:path');
  const reference = analyze(fs.readFileSync(path.join(__dirname, 'fixtures', 'LibreOffice-Lookup-Duplicate-Error-Reference.xlsx')));
  assert.equal(reference.writable, true, reference.readOnlyReason);
  const cells = reference.sheets[0].cells;
  const context = { getCell: (row, column) => {
    let letters = '';
    for (let value = column + 1; value > 0; value = Math.floor((value - 1) / 26)) letters = String.fromCharCode(65 + (value - 1) % 26) + letters;
    const cell = cells.get(letters + (row + 1));
    if (!cell) return null;
    if (cell.kind === 'number') return Number(cell.value);
    if (cell.kind === 'boolean') return /^TRUE$/i.test(String(cell.value));
    return cell.value;
  } };
  for (const [address, expected] of [['G9', 3], ['G10', 'last'], ['G11', 'last'], ['G12', 3]]) {
    const cell = cells.get(address);
    assert.ok(cell?.formula, `${address} retains its reference formula`);
    assert.equal(cell.cachedValue, String(expected), `${address} is the LibreOffice Calc cached result`);
    assert.equal(evaluateFormula(cell.formula, context), expected, `${cell.formula} matches the LibreOffice reference`);
  }
});

test('TEXTSPLIT은 열·행 구분자, 빈 값, 대소문자 모드와 불균일 행 패딩을 처리한다', () => {
  assert.deepEqual(evaluateFormula('TEXTSPLIT("a,b;c,d",",",";")', { allowArrayResult: true }), [['a', 'b'], ['c', 'd']]);
  assert.deepEqual(evaluateFormula('TEXTSPLIT("a,,b",",",,TRUE)', { allowArrayResult: true }), [['a', 'b']]);
  assert.deepEqual(evaluateFormula('TEXTSPLIT("aXb","x",,FALSE,1)', { allowArrayResult: true }), [['a', 'b']]);
  assert.deepEqual(evaluateFormula('TEXTSPLIT("a,b;c",",",";",FALSE,0,"-")', { allowArrayResult: true }), [['a', 'b'], ['c', '-']]);
  assert.deepEqual(evaluateFormula('TEXTSPLIT("a,b;c",HSTACK(",",";"))', { allowArrayResult: true }), [['a', 'b', 'c']]);
  assert.deepEqual(evaluateFormula('TEXTSPLIT("a::b;c",HSTACK(":","::"),";")', { allowArrayResult: true }), [['a', 'b'], ['c', '#N/A']]);
  assert.deepEqual(evaluateFormula('HSTACK("a","b")', { allowArrayResult: true }), [['a', 'b']]);
  assert.equal(evaluateFormula('TEXTSPLIT("a,,b","",,TRUE)'), '#VALUE!');
  assert.equal(evaluateFormula('TEXTSPLIT(",,",",",,TRUE)'), '#CALC!');
  assert.equal(evaluateFormula('TEXTSPLIT("a,b",",",,FALSE,2)'), '#VALUE!');
});

test('TEXTSPLIT 결과는 수식 셀에서 행·열로 spill한다', () => {
  const cells = new Map([['A1', { address: 'A1', row: 0, column: 0, value: '', kind: 'string', formula: 'TEXTSPLIT("a,b;c",",",";")', hasFormula: true }]]);
  const document = { state: { formulas: true, sheets: [{ name: 'Data', cells, mergedCells: [], tables: [] }] }, getCell: (row, column) => [...cells.values()].find(cell => cell.row === row && cell.column === column) || null };
  const calculator = createFormulaCalculator(document);
  assert.deepEqual([[calculator.cell(0, 0, 0).value, calculator.cell(0, 0, 1).value], [calculator.cell(0, 1, 0).value, calculator.cell(0, 1, 1).value]], [['a', 'b'], ['c', '#N/A']]);
  assert.equal(calculator.cell(0, 1, 1).isSpill, true);
});

test('UNICHAR와 UNICODE는 BMP·보충 평면 문자를 왕복하고 잘못된 코드 포인트를 거부한다', () => {
  assert.equal(evaluateFormula('UNICHAR(65)'), 'A');
  assert.equal(evaluateFormula('UNICHAR(8364)'), '€');
  assert.equal(evaluateFormula('UNICHAR(128512)'), '😀');
  assert.equal(evaluateFormula('UNICODE("A")'), 65);
  assert.equal(evaluateFormula('UNICODE("😀x")'), 128512, 'UNICODE reads the first Unicode code point, including a surrogate pair');
  assert.equal(evaluateFormula('UNICHAR(0)'), '#VALUE!');
  assert.equal(evaluateFormula('UNICHAR(55296)'), '#VALUE!', 'a lone surrogate is not a Unicode scalar value');
  assert.equal(evaluateFormula('UNICHAR(1114112)'), '#VALUE!');
  assert.equal(evaluateFormula('UNICHAR(1.5)'), '#VALUE!');
  assert.equal(evaluateFormula('UNICODE("")'), '#VALUE!');
});

test('SEQUENCE는 동적 배열 값을 만들고 점유 셀에서는 spill 오류를 반환한다', () => {
  const cells = new Map();
  const add = (address, row, column, value, formula = null) => cells.set(address, { address, row, column, value: String(value), kind: 'number', formula, hasFormula: Boolean(formula) });
  add('A1', 0, 0, 0, 'SEQUENCE(2,2,1,2)');
  add('E1', 0, 4, 0, 'SEQUENCE(3)');
  add('H1', 0, 7, 0, 'SUM(A1#)');
  add('J1', 0, 9, 0, 'A1#');
  add('D10', 9, 3, 10); add('E10', 9, 4, 20); add('F10', 9, 5, 30);
  add('D11', 10, 3, 40); add('E11', 10, 4, 50); add('F11', 10, 5, 60);
  add('L10', 9, 11, 0, 'TRANSPOSE(D10:F11)');
  const document = { state: { sheets: [{ name: 'Data', cells, mergedCells: [], tables: [] }] }, getCell: (row, column) => [...cells.values()].find(cell => cell.row === row && cell.column === column) || null };
  const calculator = createFormulaCalculator(document);
  assert.equal(calculator.value(0, 0, 0), 1, 'the anchor cell contains the top-left array value');
  assert.deepEqual([[calculator.cell(0, 0, 0).value, calculator.cell(0, 0, 1).value], [calculator.cell(0, 1, 0).value, calculator.cell(0, 1, 1).value]], [['1', '3'], ['5', '7']]);
  assert.equal(calculator.cell(0, 1, 1).isSpill, true);
  assert.equal(calculator.value(0, 0, 4), 1);
  assert.equal(calculator.cell(0, 2, 4).value, '3');
  assert.equal(calculator.value(0, 0, 7), 16, 'the spill operator exposes every value in a dynamic array');
  assert.equal(calculator.cell(0, 1, 9).value, '5', 'a spill reference can feed another dynamic array formula');
  assert.deepEqual([[calculator.cell(0, 9, 11).value, calculator.cell(0, 9, 12).value], [calculator.cell(0, 10, 11).value, calculator.cell(0, 10, 12).value], [calculator.cell(0, 11, 11).value, calculator.cell(0, 11, 12).value]], [['10', '40'], ['20', '50'], ['30', '60']], 'TRANSPOSE spills swapped rows and columns');
  assert.equal(evaluateFormula('SUM(SEQUENCE(3))'), 6, 'aggregate functions can consume an array result directly');
  assert.equal(evaluateFormula('TRANSPOSE(1,2)'), '#VALUE!');
  add('B2', 1, 1, 99);
  assert.equal(createFormulaCalculator(document).value(0, 0, 0), '#SPILL!');
  assert.equal(evaluateFormula('SEQUENCE(0)'), '#VALUE!');
  assert.equal(evaluateFormula('SEQUENCE(101,100)'), '#NUM!', 'array dimensions are bounded');
});

test('XLOOKUP은 일치 행의 여러 반환 열을 동적 배열로 spill한다', () => {
  const cells = new Map();
  const add = (row, column, value, kind = typeof value === 'number' ? 'number' : 'string', formula = null) => {
    const address = `${String.fromCharCode(65 + column)}${row + 1}`;
    cells.set(address, { address, row, column, value: String(value), kind, formula, hasFormula: Boolean(formula) });
  };
  add(0, 0, 1); add(1, 0, 2); add(2, 0, 3);
  add(0, 1, 'One'); add(1, 1, 'Two'); add(2, 1, 'Three');
  add(0, 2, 10); add(1, 2, 20); add(2, 2, 30);
  add(0, 4, 0, 'number', 'XLOOKUP(2,A1:A3,B1:C3)');
  const document = { state: { formulas: true, sheets: [{ name: 'Data', cells, mergedCells: [], tables: [] }] }, getCell: (row, column) => [...cells.values()].find(cell => cell.row === row && cell.column === column) || null };
  const calculator = createFormulaCalculator(document);
  assert.equal(calculator.cell(0, 0, 4).value, 'Two');
  assert.equal(calculator.cell(0, 0, 5).value, '20');
  assert.equal(calculator.cell(0, 0, 5).isSpill, true);
});

test('FILTER는 행·열 조건 배열로 원본 범위를 골라 spill하고 빈 결과 대안을 처리한다', () => {
  const cells = new Map();
  const add = (address, row, column, value, kind = 'number', formula = null) => cells.set(address, { address, row, column, value: String(value), kind, formula, hasFormula: Boolean(formula) });
  add('A1', 0, 0, 10); add('A2', 1, 0, 20); add('A3', 2, 0, 30); add('A4', 3, 0, 40);
  add('B1', 0, 1, 'TRUE', 'boolean'); add('B2', 1, 1, 'FALSE', 'boolean'); add('B3', 2, 1, 'TRUE', 'boolean'); add('B4', 3, 1, 'FALSE', 'boolean');
  add('C1', 0, 2, 'FALSE', 'boolean'); add('C2', 1, 2, 'FALSE', 'boolean'); add('C3', 2, 2, 'FALSE', 'boolean'); add('C4', 3, 2, 'FALSE', 'boolean');
  add('M1', 0, 12, 'FALSE', 'boolean'); add('N1', 0, 13, 'TRUE', 'boolean');
  add('E1', 0, 4, 0, 'number', 'FILTER(A1:A4,B1:B4)');
  add('G1', 0, 6, 0, 'number', 'FILTER(A1:A4,C1:C4,"empty")');
  add('I1', 0, 8, 0, 'number', 'FILTER(A1:A4,C1:C4)');
  add('O1', 0, 14, 0, 'number', 'FILTER(A1:B2,M1:N1)');
  const document = { state: { formulas: true, sheets: [{ name: 'Data', cells, mergedCells: [], tables: [] }] }, getCell: (row, column) => [...cells.values()].find(cell => cell.row === row && cell.column === column) || null };
  const calculator = createFormulaCalculator(document);
  assert.equal(calculator.value(0, 0, 4), 10);
  assert.equal(calculator.cell(0, 1, 4).value, '30');
  assert.equal(calculator.value(0, 0, 6), 'empty');
  assert.equal(calculator.value(0, 0, 8), '#CALC!');
  assert.equal(calculator.value(0, 0, 14), true, 'a horizontal include array selects matching columns from a two-dimensional source');
  assert.equal(calculator.cell(0, 1, 14).value, 'FALSE');
  assert.equal(evaluateFormula('SUM(FILTER(A1:A4,B1:B4))', { getCell: (row, column) => [[10, true], [20, false], [30, true], [40, false]][row]?.[column] ?? null }), 40);
});

test('SORT는 지정한 행·열 키로 안정 정렬한 동적 배열을 반환한다', () => {
  const cells = new Map();
  const add = (address, row, column, value, kind = 'number', formula = null) => cells.set(address, { address, row, column, value: String(value), kind, formula, hasFormula: Boolean(formula) });
  [[3, 'c'], [1, 'a'], [2, 'b'], [2, 'd']].forEach(([number, label], row) => { add(`A${row + 1}`, row, 0, number); add(`B${row + 1}`, row, 1, label, 'string'); });
  add('D1', 0, 3, 0, 'number', 'SORT(A1:B4,1,-1)');
  add('D7', 6, 3, 0, 'number', 'SORT(A1:B2,2,-1,TRUE)');
  const document = { state: { formulas: true, sheets: [{ name: 'Data', cells, mergedCells: [], tables: [] }] }, getCell: (row, column) => [...cells.values()].find(cell => cell.row === row && cell.column === column) || null };
  const calculator = createFormulaCalculator(document);
  assert.deepEqual([[calculator.cell(0, 0, 3).value, calculator.cell(0, 0, 4).value], [calculator.cell(0, 1, 3).value, calculator.cell(0, 1, 4).value], [calculator.cell(0, 2, 3).value, calculator.cell(0, 2, 4).value], [calculator.cell(0, 3, 3).value, calculator.cell(0, 3, 4).value]], [['3', 'c'], ['2', 'b'], ['2', 'd'], ['1', 'a']]);
  assert.deepEqual([[calculator.cell(0, 6, 3).value, calculator.cell(0, 6, 4).value], [calculator.cell(0, 7, 3).value, calculator.cell(0, 7, 4).value]], [['c', '3'], ['a', '1']]);
  assert.equal(evaluateFormula('SORT(A1:A2,2)'), '#VALUE!');
  assert.equal(evaluateFormula('SORT(1,1,0)'), '#VALUE!');
});

test('SORTBY는 다중 정렬 키와 행·열 방향을 처리한다', () => {
  const cells = new Map();
  const add = (address, row, column, value, kind = 'number', formula = null) => cells.set(address, { address, row, column, value: String(value), kind, formula, hasFormula: Boolean(formula) });
  [['a', 2], ['b', 1], ['c', 2], ['d', 1]].forEach(([label, amount], row) => { add(`A${row + 1}`, row, 0, label, 'string'); add(`B${row + 1}`, row, 1, amount); });
  [2, 1, 1, 1].forEach((value, row) => add(`C${row + 1}`, row, 2, value));
  [1, 2, 1, 1].forEach((value, row) => add(`D${row + 1}`, row, 3, value));
  [[1, 3], [2, 4]].forEach((values, row) => values.forEach((value, column) => add(`${String.fromCharCode(72 + column)}${row + 1}`, row, 7 + column, value)));
  add('H3', 2, 7, 2); add('I3', 2, 8, 1);
  add('F1', 0, 5, 0, 'number', 'SORTBY(A1:B4,C1:C4,1,D1:D4,-1)');
  add('K1', 0, 10, 0, 'number', 'SORTBY(H1:I2,H3:I3,1)');
  const document = { state: { formulas: true, sheets: [{ name: 'Data', cells, mergedCells: [], tables: [] }] }, getCell: (row, column) => [...cells.values()].find(cell => cell.row === row && cell.column === column) || null };
  const calculator = createFormulaCalculator(document);
  assert.deepEqual([0, 1, 2, 3].map(offset => [calculator.cell(0, offset, 5).value, calculator.cell(0, offset, 6).value]), [['b', '1'], ['c', '2'], ['d', '1'], ['a', '2']]);
  assert.deepEqual([[calculator.cell(0, 0, 10).value, calculator.cell(0, 0, 11).value], [calculator.cell(0, 1, 10).value, calculator.cell(0, 1, 11).value]], [['3', '1'], ['4', '2']]);
  assert.equal(evaluateFormula('SORTBY(A1:A2,B1:B2,0)'), '#VALUE!');
  assert.equal(evaluateFormula('SORTBY(1,2)'), '#VALUE!');
});

test('UNIQUE는 첫 출현 순서·정확히 한 번·행과 열 방향으로 중복을 제거한다', () => {
  const cells = new Map();
  const add = (address, row, column, value, kind = 'string', formula = null) => cells.set(address, { address, row, column, value: String(value), kind, formula, hasFormula: Boolean(formula) });
  ['alpha', 'beta', 'alpha', 'gamma', 'beta'].forEach((value, row) => add(`A${row + 1}`, row, 0, value));
  [[1, 2, 1], [4, 5, 4], [7, 6, 7]].forEach((values, row) => values.forEach((value, column) => add(`${String.fromCharCode(72 + column)}${row + 1}`, row, 7 + column, value, 'number')));
  add('D1', 0, 3, 0, 'number', 'UNIQUE(A1:A5)');
  add('F1', 0, 5, 0, 'number', 'UNIQUE(A1:A5,FALSE,TRUE)');
  add('L1', 0, 11, 0, 'number', 'UNIQUE(H1:J3,TRUE)');
  add('N1', 0, 13, 0, 'number', 'UNIQUE(H1:J3,TRUE,TRUE)');
  const document = { state: { formulas: true, sheets: [{ name: 'Data', cells, mergedCells: [], tables: [] }] }, getCell: (row, column) => [...cells.values()].find(cell => cell.row === row && cell.column === column) || null };
  const calculator = createFormulaCalculator(document);
  assert.deepEqual([0, 1, 2].map(offset => calculator.cell(0, offset, 3).value), ['alpha', 'beta', 'gamma']);
  assert.equal(calculator.value(0, 0, 5), 'gamma');
  assert.deepEqual([0, 1, 2].map(offset => [calculator.cell(0, offset, 11).value, calculator.cell(0, offset, 12).value]), [['1', '2'], ['4', '5'], ['7', '6']]);
  assert.deepEqual([0, 1, 2].map(offset => calculator.cell(0, offset, 13).value), ['2', '5', '6']);
  assert.equal(evaluateFormula('UNIQUE(A1:A2,2)'), '#VALUE!');
});

test('CHOOSECOLS와 CHOOSEROWS는 양수·음수 인덱스로 선택한 배열을 spill한다', () => {
  const cells = new Map();
  const add = (address, row, column, value, kind = 'number', formula = null) => cells.set(address, { address, row, column, value: String(value), kind, formula, hasFormula: Boolean(formula) });
  [[1, 2, 3], [4, 5, 6], [7, 8, 9]].forEach((values, row) => values.forEach((value, column) => add(`${String.fromCharCode(65 + column)}${row + 1}`, row, column, value)));
  add('E1', 0, 4, 0, 'number', 'CHOOSECOLS(A1:C3,3,1)');
  add('H1', 0, 7, 0, 'number', 'CHOOSEROWS(A1:C3,-1,1)');
  add('K1', 0, 10, 0, 'number', 'CHOOSECOLS(A1:C3,4)');
  add('M1', 0, 12, 0, 'number', 'CHOOSEROWS(A1:C3,0)');
  const document = { state: { formulas: true, sheets: [{ name: 'Data', cells, mergedCells: [], tables: [] }] }, getCell: (row, column) => [...cells.values()].find(cell => cell.row === row && cell.column === column) || null };
  const calculator = createFormulaCalculator(document);
  assert.deepEqual([0, 1, 2].map(row => [calculator.cell(0, row, 4).value, calculator.cell(0, row, 5).value]), [['3', '1'], ['6', '4'], ['9', '7']]);
  assert.deepEqual([0, 1, 2].map(column => [calculator.cell(0, 0, 7 + column).value, calculator.cell(0, 1, 7 + column).value]), [['7', '1'], ['8', '2'], ['9', '3']]);
  assert.equal(calculator.value(0, 0, 10), '#VALUE!');
  assert.equal(calculator.value(0, 0, 12), '#VALUE!');
  assert.equal(evaluateFormula('CHOOSECOLS(A1:C2,1.5)', { getCell: () => 1 }), '#VALUE!');
});

test('TAKE와 DROP은 배열의 앞·뒤 행과 열을 선택하거나 제외해 spill한다', () => {
  const cells = new Map();
  const add = (address, row, column, value, kind = 'number', formula = null) => cells.set(address, { address, row, column, value: String(value), kind, formula, hasFormula: Boolean(formula) });
  [[1, 2, 3], [4, 5, 6], [7, 8, 9]].forEach((values, row) => values.forEach((value, column) => add(`${String.fromCharCode(65 + column)}${row + 1}`, row, column, value)));
  add('E1', 0, 4, 0, 'number', 'TAKE(A1:C3,2,2)');
  add('H1', 0, 7, 0, 'number', 'TAKE(A1:C3,-2,-1)');
  add('J1', 0, 9, 0, 'number', 'DROP(A1:C3,1,1)');
  add('M1', 0, 12, 0, 'number', 'DROP(A1:C3,-1,-1)');
  add('P1', 0, 15, 0, 'number', 'TAKE(A1:C3,4)');
  add('R1', 0, 17, 0, 'number', 'DROP(A1:C3,3)');
  const document = { state: { formulas: true, sheets: [{ name: 'Data', cells, mergedCells: [], tables: [] }] }, getCell: (row, column) => [...cells.values()].find(cell => cell.row === row && cell.column === column) || null };
  const calculator = createFormulaCalculator(document);
  assert.deepEqual([[calculator.cell(0, 0, 4).value, calculator.cell(0, 0, 5).value], [calculator.cell(0, 1, 4).value, calculator.cell(0, 1, 5).value]], [['1', '2'], ['4', '5']]);
  assert.deepEqual([calculator.cell(0, 0, 7).value, calculator.cell(0, 1, 7).value], ['6', '9']);
  assert.deepEqual([[calculator.cell(0, 0, 9).value, calculator.cell(0, 0, 10).value], [calculator.cell(0, 1, 9).value, calculator.cell(0, 1, 10).value]], [['5', '6'], ['8', '9']]);
  assert.deepEqual([[calculator.cell(0, 0, 12).value, calculator.cell(0, 0, 13).value], [calculator.cell(0, 1, 12).value, calculator.cell(0, 1, 13).value]], [['1', '2'], ['4', '5']]);
  assert.equal(calculator.value(0, 0, 15), '#VALUE!');
  assert.equal(calculator.value(0, 0, 17), '#CALC!');
  assert.equal(evaluateFormula('TAKE(A1:C2,0)', { getCell: () => 1 }), '#VALUE!');
});

test('EXPAND는 배열을 지정 크기로 늘리고 채움 값을 적용한다', () => {
  const context = { getCell: (row, column) => [[1, 2], [3, 4]][row]?.[column] ?? null, allowArrayResult: true };
  assert.deepEqual(evaluateFormula('EXPAND(A1:B2,3,4,0)', context), [[1, 2, 0, 0], [3, 4, 0, 0], [0, 0, 0, 0]]);
  assert.deepEqual(evaluateFormula('EXPAND(A1:B2,3)', context), [[1, 2], [3, 4], ['#N/A', '#N/A']]);
  assert.deepEqual(evaluateFormula('EXPAND(1,2,3,"x")', { allowArrayResult: true }), [[1, 'x', 'x'], ['x', 'x', 'x']]);
  assert.equal(evaluateFormula('EXPAND(A1:B2,1,4)', context), '#VALUE!');
  assert.equal(evaluateFormula('EXPAND(A1:B2,3,1)', context), '#VALUE!');
  assert.equal(evaluateFormula('EXPAND(A1:B2,0)', context), '#VALUE!');
  assert.equal(evaluateFormula('EXPAND(1,101,100)'), '#NUM!');
});

test('TOCOL과 TOROW는 배열을 행·열로 펼치고 빈 셀·오류 제외 및 순서를 적용한다', () => {
  const cells = new Map();
  const add = (address, row, column, value, kind = 'number', formula = null) => cells.set(address, { address, row, column, value: String(value), kind, formula, hasFormula: Boolean(formula) });
  add('A1', 0, 0, 1); add('C1', 0, 2, '#N/A', 'error');
  add('A2', 1, 0, 4); add('B2', 1, 1, 5); add('C2', 1, 2, 6);
  add('E1', 0, 4, 0, 'number', 'TOCOL(A1:C2)');
  add('G1', 0, 6, 0, 'number', 'TOCOL(A1:C2,3)');
  add('I1', 0, 8, 0, 'number', 'TOROW(A1:C2,0,TRUE)');
  add('P1', 0, 15, 0, 'number', 'TOCOL(B1:B1,1)');
  const document = { state: { formulas: true, sheets: [{ name: 'Data', cells, mergedCells: [], tables: [] }] }, getCell: (row, column) => [...cells.values()].find(cell => cell.row === row && cell.column === column) || null };
  const calculator = createFormulaCalculator(document);
  assert.deepEqual([0, 1, 2, 3, 4, 5].map(row => calculator.cell(0, row, 4).value), ['1', '', '#N/A', '4', '5', '6']);
  assert.deepEqual([0, 1, 2, 3].map(row => calculator.cell(0, row, 6).value), ['1', '4', '5', '6']);
  assert.deepEqual([0, 1, 2, 3, 4, 5].map(column => calculator.cell(0, 0, 8 + column).value), ['1', '4', '', '5', '#N/A', '6']);
  assert.equal(calculator.value(0, 0, 15), '#CALC!');
  assert.equal(evaluateFormula('TOCOL(A1:A2,4)', { getCell: () => 1 }), '#VALUE!');
  assert.equal(evaluateFormula('TOROW(A1:A2,0,1)', { getCell: () => 1 }), '#VALUE!');
});

test('WRAPROWS와 WRAPCOLS는 행 우선·열 우선으로 배열을 감싸고 마지막을 채운다', () => {
  const cells = new Map();
  const add = (address, row, column, value, kind = 'number', formula = null) => cells.set(address, { address, row, column, value: String(value), kind, formula, hasFormula: Boolean(formula) });
  [[1, 2, 3], [4, 5, 6]].forEach((values, row) => values.forEach((value, column) => add(`${String.fromCharCode(65 + column)}${row + 1}`, row, column, value)));
  add('E1', 0, 4, 0, 'number', 'WRAPROWS(A1:C2,4)');
  add('J1', 0, 9, 0, 'number', 'WRAPROWS(A1:C2,4,"pad")');
  add('O1', 0, 14, 0, 'number', 'WRAPCOLS(A1:C2,4)');
  add('T1', 0, 19, 0, 'number', 'WRAPROWS(A1:C2,0)');
  add('W1', 0, 22, 0, 'number', 'WRAPCOLS(A1:C2,10001)');
  const document = { state: { formulas: true, sheets: [{ name: 'Data', cells, mergedCells: [], tables: [] }] }, getCell: (row, column) => [...cells.values()].find(cell => cell.row === row && cell.column === column) || null };
  const calculator = createFormulaCalculator(document);
  assert.deepEqual([0, 1, 2, 3].map(column => calculator.cell(0, 0, 4 + column).value), ['1', '2', '3', '4']);
  assert.deepEqual([0, 1, 2, 3].map(column => calculator.cell(0, 1, 4 + column).value), ['5', '6', '#N/A', '#N/A']);
  assert.deepEqual([0, 1, 2, 3].map(column => calculator.cell(0, 1, 9 + column).value), ['5', '6', 'pad', 'pad']);
  assert.deepEqual([0, 1, 2, 3].map(row => [calculator.cell(0, row, 14).value, calculator.cell(0, row, 15).value]), [['1', '3'], ['4', '6'], ['2', '#N/A'], ['5', '#N/A']]);
  assert.equal(calculator.value(0, 0, 19), '#NUM!');
  assert.equal(calculator.value(0, 0, 22), '#NUM!');
});

test('HSTACK과 VSTACK은 서로 다른 크기의 범위를 결합하고 부족한 자리를 #N/A로 채운다', () => {
  const cells = new Map();
  const add = (address, row, column, value, kind = 'number', formula = null) => cells.set(address, { address, row, column, value: String(value), kind, formula, hasFormula: Boolean(formula) });
  [[1, 2], [3, 4]].forEach((values, row) => values.forEach((value, column) => add(`${String.fromCharCode(65 + column)}${row + 1}`, row, column, value)));
  [5, 6, 7].forEach((value, row) => add(`D${row + 1}`, row, 3, value));
  add('F1', 0, 5, 0, 'number', 'HSTACK(A1:B2,D1:D3)');
  add('K1', 0, 10, 0, 'number', 'VSTACK(A1:B2,D1:D3)');
  add('N1', 0, 13, 0, 'number', 'HSTACK(A1:B2,D1:D3,A1:A2)');
  const document = { state: { formulas: true, sheets: [{ name: 'Data', cells, mergedCells: [], tables: [] }] }, getCell: (row, column) => [...cells.values()].find(cell => cell.row === row && cell.column === column) || null };
  const calculator = createFormulaCalculator(document);
  assert.deepEqual([0, 1, 2].map(row => [0, 1, 2].map(column => calculator.cell(0, row, 5 + column).value)), [['1', '2', '5'], ['3', '4', '6'], ['#N/A', '#N/A', '7']]);
  assert.deepEqual([0, 1, 2, 3, 4].map(row => [calculator.cell(0, row, 10).value, calculator.cell(0, row, 11).value]), [['1', '2'], ['3', '4'], ['5', '#N/A'], ['6', '#N/A'], ['7', '#N/A']]);
  assert.deepEqual([0, 1].map(row => [0, 1, 2, 3].map(column => calculator.cell(0, row, 13 + column).value)), [['1', '2', '5', '1'], ['3', '4', '6', '3']]);
  assert.equal(evaluateFormula('HSTACK(SEQUENCE(100,100),SEQUENCE(100))'), '#NUM!');
});

test('SUBTOTAL은 필터된 행과 범위 안의 중첩 SUBTOTAL을 제외한다', () => {
  const cell = (address, row, column, value, formula = null, kind = 'number') => ({ address, row, column, value: String(value), formula, kind });
  const cells = new Map([
    ['A1', cell('A1', 0, 0, 'Group', null, 'text')],
    ['B1', cell('B1', 0, 1, 'Value', null, 'text')],
    ['A2', cell('A2', 1, 0, 'keep', null, 'text')],
    ['B2', cell('B2', 1, 1, 10)],
    ['A3', cell('A3', 2, 0, 'hide', null, 'text')],
    ['B3', cell('B3', 2, 1, 20)],
    ['A4', cell('A4', 3, 0, 'keep', null, 'text')],
    ['B4', cell('B4', 3, 1, '', 'SUBTOTAL(9,B2:B3)')],
    ['E1', cell('E1', 0, 4, '', 'SUBTOTAL(9,B2:B4)')],
    ['E2', cell('E2', 1, 4, '', 'SUBTOTAL(109,B2:B4)')],
    ['E3', cell('E3', 2, 4, '', 'SUBTOTAL(2,B2:B4)')],
    ['E4', cell('E4', 3, 4, '', 'SUBTOTAL(9,B:B)')],
    ['E5', cell('E5', 4, 4, '', 'SUBTOTAL(102,B2:B4)')]
  ]);
  const document = {
    state: { sheets: [{ name: 'Data', cells, hiddenRows: [1], autoFilter: { top: 0, bottom: 3, left: 0, right: 1, columns: [{ column: 0, values: ['keep'], blanks: false }] } }] },
    drafts: new Map(), formulaDrafts: new Map(),
    getCell(row, column) {
      let letters = '';
      for (let value = column + 1; value > 0; value = Math.floor((value - 1) / 26)) letters = String.fromCharCode(65 + (value - 1) % 26) + letters;
      return cells.get(letters + (row + 1)) || null;
    }
  };
  const calculator = createFormulaCalculator(document);
  assert.equal(calculator.value(0, 0, 4), 10, 'ordinary cell ranges omit filtered rows and nested subtotals');
  assert.equal(calculator.value(0, 1, 4), 0, '101-based function codes also omit manually hidden rows');
  assert.equal(calculator.value(0, 2, 4), 1, 'COUNT counts only visible numeric values');
  assert.equal(calculator.value(0, 3, 4), 10, 'sparse full-column references keep the formula sheet context');
  assert.equal(calculator.value(0, 4, 4), 0, '102 ignores manually hidden numeric cells');
  assert.equal(evaluateFormula('SUBTOTAL(12,1,2)'), '#VALUE!');
});

test('수식 계산기는 희소 셀만 순회해 전체 행·열 참조를 집계한다', () => {
  const cells = new Map([
    ['A1', { address: 'A1', row: 0, column: 0, value: '2', kind: 'number' }],
    ['B1', { address: 'B1', row: 0, column: 1, value: '3', kind: 'number' }],
    ['A1000000', { address: 'A1000000', row: 999999, column: 0, value: '7', kind: 'number' }],
    ['B1000000', { address: 'B1000000', row: 999999, column: 1, value: '4', kind: 'number' }],
    ['C2', { address: 'C2', row: 1, column: 2, value: '', kind: 'number', formula: 'SUM(A:A)' }],
    ['D2', { address: 'D2', row: 1, column: 3, value: '', kind: 'number', formula: 'SUM(1:1)' }],
    ['E2', { address: 'E2', row: 1, column: 4, value: '', kind: 'number', formula: 'COUNTBLANK(A:A)' }],
    ['F2', { address: 'F2', row: 1, column: 5, value: '', kind: 'number', formula: 'COUNT(A:A)' }],
    ['G2', { address: 'G2', row: 1, column: 6, value: '', kind: 'number', formula: 'COUNTIF(A:A,">0")' }],
    ['H2', { address: 'H2', row: 1, column: 7, value: '', kind: 'number', formula: 'COUNTIF(A:A,"")' }],
    ['I2', { address: 'I2', row: 1, column: 8, value: '', kind: 'number', formula: 'SUMIF(A:A,">0",B:B)' }],
    ['J2', { address: 'J2', row: 1, column: 9, value: '', kind: 'number', formula: 'SUMIFS(B:B,A:A,">0")' }],
    ['K2', { address: 'K2', row: 1, column: 10, value: '', kind: 'number', formula: 'AVERAGEIFS(B:B,A:A,">0")' }],
    ['L2', { address: 'L2', row: 1, column: 11, value: '', kind: 'number', formula: 'MAXIFS(B:B,A:A,">0")' }],
    ['M2', { address: 'M2', row: 1, column: 12, value: '', kind: 'number', formula: 'MINIFS(B:B,A:A,">0")' }],
    ['N2', { address: 'N2', row: 1, column: 13, value: '', kind: 'number', formula: 'COUNTIFS(A:A,">0",B:B,">3")' }],
    ['O2', { address: 'O2', row: 1, column: 14, value: '', kind: 'number', formula: 'COUNTIFS(A:A,"",B:B,"")' }],
    ['P2', { address: 'P2', row: 1, column: 15, value: '', kind: 'number', formula: 'SUMIF(A:A,"",B:B)' }],
    ['Q2', { address: 'Q2', row: 1, column: 16, value: '', kind: 'number', formula: 'SUMIF(A:A,">0",B1:B3)' }]
  ]);
  const document = {
    state: { sheets: [{ name: 'Data', cells }] },
    drafts: new Map(), formulaDrafts: new Map(),
    getCell(row, column) {
      let letters = '';
      for (let value = column + 1; value > 0; value = Math.floor((value - 1) / 26)) letters = String.fromCharCode(65 + (value - 1) % 26) + letters;
      return cells.get(letters + (row + 1)) || null;
    }
  };
  const calculator = createFormulaCalculator(document);
  assert.equal(calculator.value(0, 1, 2), 9, 'SUM(A:A) includes distant sparse cells without allocating a million-element array');
  assert.equal(calculator.value(0, 1, 3), 5, 'a full-row reference sums its nonempty cells');
  assert.equal(calculator.value(0, 1, 4), 1048574, 'COUNTBLANK uses the worksheet row count and sparse nonblank cells');
  assert.equal(calculator.value(0, 1, 5), 2, 'COUNT counts numeric values in a full column');
  assert.equal(calculator.value(0, 1, 6), 2, 'COUNTIF counts sparse full-column criteria matches');
  assert.equal(calculator.value(0, 1, 7), 1048574, 'COUNTIF includes blank cells without enumerating the full column');
  assert.equal(calculator.value(0, 1, 8), 7, 'SUMIF aligns sparse criterion and sum ranges');
  assert.equal(calculator.value(0, 1, 9), 7, 'SUMIFS accepts matching sparse full-column ranges');
  assert.equal(calculator.value(0, 1, 10), 3.5, 'AVERAGEIFS includes matching nonempty sparse sums');
  assert.equal(calculator.value(0, 1, 11), 4);
  assert.equal(calculator.value(0, 1, 12), 3);
  assert.equal(calculator.value(0, 1, 13), 1, 'COUNTIFS evaluates multiple aligned sparse criteria ranges');
  assert.equal(calculator.value(0, 1, 14), 1048574, 'COUNTIFS counts rows blank in every sparse full-column criteria range');
  assert.equal(calculator.value(0, 1, 15), 0, 'SUMIF does not shift sums onto nonmatching blank criterion rows');
  assert.equal(calculator.value(0, 1, 16), 3, 'SUMIF aligns a bounded sum range to a full-column criteria range without enumerating the entire column');
});

function fixture({ formula = false, extra = undefined, workbookXml = undefined, formatted = false, columnDefinitions = '', sheetViews = '', cellExtension = '', autoFilter = '' } = {}) {
  const files = {
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': workbookXml || `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Hidden" sheetId="2" state="hidden" r:id="rId2"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>${formatted ? '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' : ''}</Relationships>`,
    'xl/worksheets/sheet1.xml': `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="${formatted ? 'A1:D1' : 'A1:C1'}"/>${sheetViews}${columnDefinitions}<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>before</t></is>${cellExtension}</c>${formatted ? '<c r="B1" s="1"><v>45292</v></c><c r="C1" s="2"><v>0.125</v></c><c r="D1" s="3"><v>45292</v></c>' : '<c r="B1"><v>42</v></c><c r="C1" t="b"><v>1</v></c>'}</row>${formula ? '<row r="2"><c r="A2"><f>1+1</f><v>2</v></c></row>' : ''}</sheetData>${autoFilter}</worksheet>`,
    'xl/worksheets/sheet2.xml': `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>9</v></c></row></sheetData></worksheet>`
  };
  if (formatted) {
    files['xl/styles.xml'] = '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts><fonts count="2"><font/><font><b/><i/><color rgb="FF123456"/><sz val="14"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFEEAA"/></patternFill></fill></fills><borders count="2"><border/><border><left style="thin"><color rgb="FFABCDEF"/></left></border></borders><cellXfs count="4"><xf numFmtId="0"/><xf numFmtId="14"/><xf numFmtId="10"/><xf numFmtId="164" fontId="1" fillId="2" borderId="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf></cellXfs></styleSheet>';
    files['[Content_Types].xml'] = files['[Content_Types].xml'].replace('</Types>', '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
  }
  if (extra) files[extra.name] = extra.content;
  return zipSync(Object.fromEntries(Object.entries(files).map(([name, content]) => [name, new TextEncoder().encode(content)])), { level: 6 });
}

function sharedFormulaFixture() {
  const files = unzipSync(fixture({ formula: true }));
  const xml = new TextDecoder().decode(files['xl/worksheets/sheet1.xml'])
    .replace('<dimension ref="A1:C1"/>', '<dimension ref="A1:C3"/>')
    .replace('<row r="2"><c r="A2"><f>1+1</f><v>2</v></c></row>', '<row r="2"><c r="A2"><f t="shared" ref="A2:A3" si="0">B1+1</f><v>43</v></c></row><row r="3"><c r="A3"><f t="shared" si="0"/><v>1</v></c></row>');
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(xml);
  return zipSync(files, { level: 0 });
}

function fixtureWithCachedChart(title = 'Revenue & Growth') {
  const files = unzipSync(fixture());
  const decode = name => new TextDecoder().decode(files[name]);
  const encode = value => new TextEncoder().encode(value);
  files['xl/worksheets/sheet1.xml'] = encode(decode('xl/worksheets/sheet1.xml').replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ').replace('</worksheet>', '<drawing r:id="rId3"/></worksheet>'));
  files['xl/worksheets/_rels/sheet1.xml.rels'] = encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>');
  files['xl/drawings/drawing1.xml'] = encode('<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:graphicFrame><xdr:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></xdr:graphic></xdr:graphicFrame></xdr:twoCellAnchor></xdr:wsDr>');
  files['xl/drawings/_rels/drawing1.xml.rels'] = encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>');
  const escapedTitle = title.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
  files['xl/charts/chart1.xml'] = encode(`<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>${escapedTitle}</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:tx><c:strRef><c:f>Data!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx><c:cat><c:strRef><c:f>Data!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Jan</c:v></c:pt><c:pt idx="1"><c:v>Feb</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>Data!$B$2:$B$3</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>4</c:v></c:pt><c:pt idx="1"><c:v>6</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`);
  files['[Content_Types].xml'] = encode(decode('[Content_Types].xml').replace('</Types>', '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>'));
  return zipSync(files, { level: 6 });
}

test('차트 인쇄 앵커는 셀 두 점·셀과 크기·절대 좌표·인쇄 제외를 보존한다', () => {
  const marker = name => `<xdr:${name}><xdr:col>2</xdr:col><xdr:colOff>9525</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>19050</xdr:rowOff></xdr:${name}>`;
  const cases = [
    ['twoCellAnchor', marker('from') + marker('to'), { printable: true, type: 'twoCellAnchor', from: { column: 2, row: 3, x: 1, y: 2 }, to: { column: 2, row: 3, x: 1, y: 2 } }],
    ['oneCellAnchor', marker('from') + '<xdr:ext cx="952500" cy="476250"/>', { printable: true, type: 'oneCellAnchor', from: { column: 2, row: 3, x: 1, y: 2 }, width: 100, height: 50 }],
    ['absoluteAnchor', '<xdr:pos x="95250" y="190500"/><xdr:ext cx="952500" cy="476250"/>', { printable: true, type: 'absoluteAnchor', x: 10, y: 20, width: 100, height: 50 }],
    ['twoCellAnchor', '<xdr:clientData fPrintsWithSheet="false"/>', { printable: false }],
    ['twoCellAnchor', '', null],
    ['twoCellAnchor', marker('from').replace('<xdr:row>3', '<xdr:row>-1') + marker('to'), null],
    ['twoCellAnchor', marker('from') + marker('to') + '<xdr:xfrm rot="60000"/>', null]
  ];
  for (const [type, geometry, expected] of cases) {
    const files = unzipSync(fixtureWithCachedChart()), key = 'xl/drawings/drawing1.xml';
    files[key] = new TextEncoder().encode(new TextDecoder().decode(files[key]).replace('<xdr:twoCellAnchor>', `<xdr:${type}>${geometry}`).replace('</xdr:twoCellAnchor>', `</xdr:${type}>`));
    const model = analyze(zipSync(files));
    assert.equal(model.writable, true, model.readOnlyReason); assert.deepEqual(model.sheets[0].chartAnchors, [expected]);
    const changed = patchCells(model, [{ sheetIndex: 0, address: 'D1', value: 'changed' }]).model;
    assert.deepEqual(changed.files[key], files[key], 'cell edits preserve drawing bytes');
  }
  const fixedFiles = unzipSync(fixtureWithCachedChart()), fixedKey = 'xl/drawings/drawing1.xml';
  fixedFiles[fixedKey] = new TextEncoder().encode(new TextDecoder().decode(fixedFiles[fixedKey]).replace('<xdr:twoCellAnchor>', '<xdr:twoCellAnchor editAs="oneCell">' + marker('from') + marker('to')));
  const fixed = analyze(zipSync(fixedFiles)); assert.equal(fixed.writable, true); assert.equal(fixed.sheets[0].chartAnchors[0].editAs, 'oneCell');
});

function fixtureWithPositionedChart(editAs = 'twoCell', hidden = false) {
  const files = unzipSync(fixtureWithCachedChart()), key = 'xl/drawings/drawing1.xml';
  const marker = (name, column, row) => `<xdr:${name}><xdr:col>${column}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:${name}>`;
  for (const name of ['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml']) files[name] = new TextEncoder().encode(new TextDecoder().decode(files[name]).replace('<sheetData>', '<sheetFormatPr defaultColWidth="20.5703125" defaultRowHeight="19.5"/><sheetData>'));
  let xml = new TextDecoder().decode(files[key]);
  xml = xml.replace('<xdr:twoCellAnchor>', `<xdr:twoCellAnchor editAs="${editAs}">` + marker('from', 2, 2) + marker('to', 6, 8));
  xml = xml.replace('<xdr:graphicFrame>', `<xdr:graphicFrame><xdr:xfrm><a:off x="${288 * 9525}" y="${52 * 9525}"/><a:ext cx="${576 * 9525}" cy="${156 * 9525}"/></xdr:xfrm>`);
  if (hidden) xml = xml.replace('</xdr:twoCellAnchor>', '<xdr:clientData fPrintsWithSheet="0"/></xdr:twoCellAnchor>');
  files[key] = new TextEncoder().encode(xml); return zipSync(files);
}

test('숫자·텍스트 조건 필터는 읽기·저장·구조 변경에서 조건과 다른 부품을 보존한다', () => {
  const custom = { column: 1, custom: { and: true, conditions: [{ operator: 'greaterThanOrEqual', value: '10' }, { operator: 'lessThan', value: '20' }] } };
  const original = analyze(fixture()), filter = { top: 0, bottom: 4, left: 0, right: 2, columns: [custom] };
  const model = patchAutoFilters(original, [{ sheetIndex: 0, filter }]).model; assert.deepEqual(model.sheets[0].autoFilter, filter); assert.equal(model.writable, true);
  for (const [name, bytes] of Object.entries(original.files)) if (name !== original.sheets[0].path) assert.deepEqual(model.files[name], bytes, name);
  const inserted = transformStructure(model, { sheetIndex: 0, axis: 'column', action: 'insert', index: 0 }).model;
  assert.equal(inserted.sheets[0].autoFilter.columns[0].column, 2); assert.deepEqual(inserted.sheets[0].autoFilter.columns[0].custom, custom.custom);
  const shifted = transformStructure(inserted, { sheetIndex: 0, axis: 'row', action: 'insert', index: 2 }).model; assert.equal(shifted.sheets[0].autoFilter.bottom, 5); assert.deepEqual(shifted.sheets[0].autoFilter.columns[0].custom, custom.custom);
  const text = { ...filter, columns: [{ column: 0, custom: { and: false, conditions: [{ operator: 'equal', value: '*a&"\n\t~?*' }] } }] };
  const escaped = patchAutoFilters(model, [{ sheetIndex: 0, filter: text }]).model; assert.deepEqual(escaped.sheets[0].autoFilter, text); assert.match(escaped.sheets[0].xml, /&amp;&quot;&#10;&#9;/);
});

test('잘못된 조건 필터는 원본을 보존해 읽기 전용으로 열고 새 적용도 거부한다', () => {
  const forms = ['<customFilters/>', '<customFilters and="2"><customFilter val="x"/></customFilters>', '<customFilters><customFilter operator="unknown" val="x"/></customFilters>', '<customFilters><customFilter operator="" val="x"/></customFilters>', '<customFilters><customFilter val="x"/><customFilter val="y"/><customFilter val="z"/></customFilters>', '<customFilters><customFilter val="' + '*a'.repeat(513) + '"/></customFilters>'];
  for (const form of forms) {
    const bytes = fixture({ autoFilter: '<autoFilter ref="A1:B4"><filterColumn colId="0">' + form + '</filterColumn></autoFilter>' });
    const model = analyze(bytes); assert.equal(model.writable, false, form.slice(0, 80)); assert.deepEqual(Buffer.from(model.bytes), Buffer.from(bytes));
  }
  const original = analyze(fixture()); assert.throws(() => patchAutoFilters(original, [{ sheetIndex: 0, filter: { top: 0, bottom: 3, left: 0, right: 1, columns: [{ column: 0, custom: { and: true, conditions: [{ operator: 'bad', value: 'x' }] } }] } }]), /유효하지/);
});

test('XLSX 기본 치수와 명시된 기본 너비·행 높이는 원본 단위로 해석한다', () => {
  const plain = analyze(fixture()); assert.equal(plain.sheets[0].defaultColumnWidth, 64); assert.equal(plain.sheets[0].defaultRowHeight, 20);
  const files = unzipSync(fixture()), key = 'xl/worksheets/sheet1.xml', xml = new TextDecoder().decode(files[key]);
  files[key] = new TextEncoder().encode(xml.replace('<sheetData>', '<sheetFormatPr defaultColWidth="12" defaultRowHeight="18"/><cols><col min="1" max="1" width="0"/><col min="2" max="2" width="255"/></cols><sheetData>'));
  const custom = analyze(zipSync(files)); assert.equal(custom.sheets[0].defaultColumnWidth, 84); assert.equal(custom.sheets[0].defaultRowHeight, 24); assert.deepEqual(custom.sheets[0].hiddenColumns, [0]);
  const shown = patchHiddenColumns(custom, [{ sheetIndex: 0, columns: [0], hidden: false }]).model; assert.deepEqual(shown.sheets[0].hiddenColumns, []); assert.equal(shown.sheets[0].columnWidths[0], undefined); assert.equal(shown.sheets[0].defaultColumnWidth, 84);
  files[key] = new TextEncoder().encode(new TextDecoder().decode(custom.files[key]).replace('width="0"', 'width="0.05"'));
  const subpixel = analyze(zipSync(files)); assert.deepEqual(subpixel.sheets[0].hiddenColumns, [0]);
  assert.deepEqual(patchHiddenColumns(subpixel, [{ sheetIndex: 0, columns: [0], hidden: false }]).model.sheets[0].hiddenColumns, []);
  for (const value of ['', '-1', '256', 'NaN']) {
    files[key] = new TextEncoder().encode(xml.replace('<sheetData>', `<sheetFormatPr defaultColWidth="${value}"/><sheetData>`));
    assert.throws(() => analyze(zipSync(files)), /기본 열 너비/);
  }
});

test('기본 열 너비 0인 시트는 명시한 열만 표시하고 숨김 해제한 열의 너비를 복구한다', () => {
  const files = unzipSync(fixture()), key = 'xl/worksheets/sheet1.xml';
  files[key] = new TextEncoder().encode(new TextDecoder().decode(files[key]).replace('<sheetData>', '<sheetFormatPr defaultColWidth="0"/><cols><col min="2" max="2" width="12"/><col min="4" max="4"/></cols><sheetData>'));
  const model = analyze(zipSync(files)); assert.equal(model.sheets[0].defaultColumnWidth, 0); assert.equal(model.sheets[0].hiddenColumns.length, 16383); assert.equal(model.sheets[0].hiddenColumns.includes(1), false);
  const shown = patchHiddenColumns(model, [{ sheetIndex: 0, columns: [0, 3], hidden: false }]).model;
  assert.equal(shown.sheets[0].hiddenColumns.includes(0), false); assert.equal(shown.sheets[0].hiddenColumns.includes(3), false); assert.equal(shown.sheets[0].columnWidths[0], 9.140625); assert.equal(shown.sheets[0].columnWidths[3], 9.140625);
});

test('차트 배치 세 모드는 열 너비·행 높이·자동 맞춤·숨김 변경과 함께 유지된다', () => {
  const box = model => chartPrintGeometry(model.sheets[0], model.sheets[0].chartAnchors).boxes[0];
  for (const mode of ['twoCell', 'oneCell', 'absolute']) {
    const initial = analyze(fixtureWithPositionedChart(mode)); assert.equal(initial.writable, true, initial.readOnlyReason);
    const original = box(initial); assert.deepEqual([original.x, original.y, original.width, original.height], [288, 52, 576, 156]);
    let model = patchColumnWidths(initial, [{ sheetIndex: 0, column: 3, width: 40 }]).model;
    assert.ok(Math.abs(box(model).width - (mode === 'twoCell' ? 712 : 576)) < .001);
    model = patchColumnWidths(model, [{ sheetIndex: 0, column: 0, width: 40 }]).model;
    assert.ok(Math.abs(box(model).x - (mode === 'absolute' ? 288 : 424)) < .001);
    if (mode !== 'twoCell') {
      const drawing = new TextDecoder().decode(model.files[model.sheets[0].chartDrawingPath]);
      assert.ok(drawing.includes(`<a:off x="${Math.round(box(model).x * 9525)}"`)); assert.ok(drawing.includes(`<a:ext cx="${576 * 9525}" cy="${156 * 9525}"`));
    }
    model = patchRowHeights(model, 0, 3, 3, 39).model;
    assert.ok(Math.abs(box(model).height - (mode === 'twoCell' ? 182 : 156)) < .001);
    model = patchRowHeights(model, 0, 0, 0, 39).model;
    assert.ok(Math.abs(box(model).y - (mode === 'absolute' ? 52 : 78)) < .001);
    const full = patchRowHeights(initial, 0, 0, 1048575, 30).model;
    assert.ok(Math.abs(box(full).height - (mode === 'twoCell' ? 240 : 156)) < .001);
    const fitted = patchRowHeightValues(initial, 0, [{ row: 3, height: 39 }]).model;
    assert.ok(Math.abs(box(fitted).height - (mode === 'twoCell' ? 182 : 156)) < .001);
    model = patchHiddenColumns(initial, [{ sheetIndex: 0, columns: [3], hidden: true }]).model;
    assert.ok(Math.abs(box(model).width - (mode === 'twoCell' ? 432 : 576)) < .001);
    model = patchHiddenColumns(model, [{ sheetIndex: 0, columns: [0], hidden: true }]).model;
    assert.ok(Math.abs(box(model).x - (mode === 'absolute' ? 288 : 144)) < .001);
    model = patchHiddenRows(model, [{ sheetIndex: 0, rows: [3], hidden: true }]).model;
    assert.ok(Math.abs(box(model).height - (mode === 'twoCell' ? 130 : 156)) < .001);
    model = patchHiddenRows(model, [{ sheetIndex: 0, rows: [0], hidden: true }]).model;
    assert.ok(Math.abs(box(model).y - (mode === 'absolute' ? 52 : 26)) < .001);
    model = patchHiddenColumns(model, [{ sheetIndex: 0, columns: [0, 3], hidden: false }]).model;
    model = patchHiddenRows(model, [{ sheetIndex: 0, rows: [0, 3], hidden: false }]).model;
    const restored = box(model); for (const key of ['x', 'y', 'width', 'height']) assert.ok(Math.abs(restored[key] - original[key]) < .001, mode + ':' + key);
    for (const [name, bytes] of Object.entries(initial.files)) if (![initial.sheets[0].path, initial.sheets[0].chartDrawingPath].includes(name)) assert.deepEqual(model.files[name], bytes, name);
  }
});

test('공유 그리기의 고정 차트 크기를 바꾸면 대상 시트만 그리기를 분리한다', () => {
  const files = unzipSync(fixtureWithPositionedChart('oneCell')), key = 'xl/worksheets/sheet2.xml';
  files[key] = new TextEncoder().encode(new TextDecoder().decode(files[key]).replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ').replace('</worksheet>', '<drawing r:id="rId1"/></worksheet>'));
  files['xl/worksheets/_rels/sheet2.xml.rels'] = new TextEncoder().encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>');
  const original = analyze(zipSync(files)); assert.equal(original.writable, true, original.readOnlyReason);
  const changed = patchColumnWidths(original, [{ sheetIndex: 0, column: 3, width: 40 }]).model;
  assert.notEqual(changed.sheets[0].chartDrawingPath, changed.sheets[1].chartDrawingPath);
  assert.deepEqual(changed.files[original.sheets[0].chartDrawingPath], original.files[original.sheets[0].chartDrawingPath]);
  for (const sheet of changed.sheets) assert.ok(Math.abs(chartPrintGeometry(sheet, sheet.chartAnchors).boxes[0].width - 576) < .001);
  const both = patchColumnWidths(original, [{ sheetIndex: 0, column: 3, width: 40 }, { sheetIndex: 1, column: 4, width: 30 }]).model;
  assert.equal(both.writable, true); assert.notEqual(both.sheets[0].chartDrawingPath, both.sheets[1].chartDrawingPath);
  const placed = patchChartPlacement(original, 0, 0, 'absolute').model;
  assert.equal(placed.sheets[0].chartPlacements[0], 'absolute'); assert.equal(placed.sheets[1].chartPlacements[0], 'oneCell');
  assert.notEqual(placed.sheets[0].chartDrawingPath, placed.sheets[1].chartDrawingPath);
  assert.deepEqual(placed.files[original.sheets[1].chartDrawingPath], original.files[original.sheets[1].chartDrawingPath]);
});

test('인쇄 제외인 고정 차트도 좌표를 보존하고 손상된 고정 좌표는 변경 전에 거부한다', () => {
  const model = analyze(fixtureWithPositionedChart('oneCell', true)), drawing = model.sheets[0].chartDrawingPath;
  const changed = patchColumnWidths(model, [{ sheetIndex: 0, column: 3, width: 40 }]).model;
  assert.deepEqual(changed.sheets[0].chartAnchors, [{ printable: false }]); assert.notDeepEqual(changed.files[drawing], model.files[drawing]);
  const files = { ...model.files }; files[drawing] = new TextEncoder().encode(new TextDecoder().decode(files[drawing]).replace('<xdr:col>6</xdr:col>', '<xdr:col>oops</xdr:col>'));
  const invalid = analyze(zipSync(files)); assert.equal(invalid.writable, true);
  assert.throws(() => patchColumnWidths(invalid, [{ sheetIndex: 0, column: 3, width: 40 }]), /고정 차트/); assert.deepEqual(invalid.files[drawing], files[drawing]);
});

test('차트 배치 변경은 세 앵커 유형의 위치·크기·메타데이터를 보존한다', () => {
  for (const kind of ['twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor']) {
    const files = unzipSync(fixtureWithPositionedChart()), key = 'xl/drawings/drawing1.xml';
    let xml = new TextDecoder().decode(files[key]);
    if (kind !== 'twoCellAnchor') {
      xml = xml.replace(/<xdr:twoCellAnchor[^>]*>/, '<xdr:' + kind + '>').replace('</xdr:twoCellAnchor>', '</xdr:' + kind + '>');
      xml = xml.replace(/<xdr:to>[\s\S]*?<\/xdr:to>/, `<xdr:ext cx="${576 * 9525}" cy="${156 * 9525}"/>`);
      if (kind === 'absoluteAnchor') xml = xml.replace(/<xdr:from>[\s\S]*?<\/xdr:from>/, `<xdr:pos x="${288 * 9525}" y="${52 * 9525}"/>`);
    }
    files[key] = new TextEncoder().encode(xml); let model = analyze(zipSync(files)); assert.equal(model.writable, true, model.readOnlyReason);
    const cache = Buffer.from(model.files['xl/charts/chart1.xml']);
    for (const mode of ['absolute', 'oneCell', 'twoCell']) {
      model = patchChartPlacement(model, 0, 0, mode).model;
      const box = chartPrintGeometry(model.sheets[0], model.sheets[0].chartAnchors).boxes[0];
      for (const [key, expected] of Object.entries({ x: 288, y: 52, width: 576, height: 156 })) assert.ok(Math.abs(box[key] - expected) < .001);
      assert.equal(model.sheets[0].chartPlacements[0], mode); assert.deepEqual(Buffer.from(model.files['xl/charts/chart1.xml']), cache);
      assert.equal(patchChartPlacement(model, 0, 0, mode).model, model);
    }
    assert.throws(() => patchChartPlacement(model, 0, 0, 'other'), /차트 배치/);
    assert.throws(() => patchChartPlacement(model, 0, 99, 'absolute'), /차트 배치/);
  }
});

test('XML 선언으로 시작하는 워크시트에도 차트를 추가하고 선언을 보존한다', () => {
  const files = unzipSync(fixture()), key = 'xl/worksheets/sheet1.xml';
  files[key] = new TextEncoder().encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + new TextDecoder().decode(files[key]));
  const source = patchCells(analyze(zipSync(files)), [{ sheetIndex: 0, address: 'B2', value: '5', kind: 'number' }]).model;
  const result = patchAddChart(source, 0, { top: 0, bottom: 2, left: 0, right: 1 }).model;
  assert.equal(result.sheets[0].charts.length, 1); assert.equal(result.sheets[0].chartAnchors[0].type, 'twoCellAnchor');
  assert.match(new TextDecoder().decode(result.files[key]), /^<\?xml version="1.0"/);
});

test('외부 하이퍼링크가 있는 XLSX는 셀 편집 후 링크 관계를 보존한다', () => {
  const files = unzipSync(fixture());
  const decode = name => new TextDecoder().decode(files[name]);
  const encode = value => new TextEncoder().encode(value);
  files['xl/worksheets/sheet1.xml'] = encode(decode('xl/worksheets/sheet1.xml')
    .replace('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
    .replace('</sheetData>', '</sheetData><hyperlinks><hyperlink ref="A1" r:id="rId1" tooltip="Open source" display="Source label"/><hyperlink ref="B1" location="Hidden!A1" display="Internal"/></hyperlinks>'));
  files['xl/worksheets/_rels/sheet1.xml.rels'] = encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/source" TargetMode="External"/></Relationships>');
  const original = analyze(zipSync(files));
  assert.equal(original.writable, true, original.readOnlyReason);
  assert.equal(original.sheets[0].hyperlinks.length, 2);
  assert.deepEqual(original.sheets[0].hyperlinks[0], { top: 0, bottom: 0, left: 0, right: 0, relationshipId: 'rId1', location: null, display: 'Source label', tooltip: 'Open source', target: 'https://example.test/source' });
  assert.deepEqual(original.sheets[0].hyperlinks[1], { top: 0, bottom: 0, left: 1, right: 1, relationshipId: null, location: 'Hidden!A1', display: 'Internal', tooltip: null, target: null });
  const shifted = transformStructure(original, { sheetIndex: 0, axis: 'row', action: 'insert', index: 0 }).model;
  assert.deepEqual(shifted.sheets[0].hyperlinks.map(({ top, bottom, left, right }) => ({ top, bottom, left, right })), [
    { top: 1, bottom: 1, left: 0, right: 0 }, { top: 1, bottom: 1, left: 1, right: 1 }
  ]);
  assert.equal(shifted.sheets[0].hyperlinks[1].location, 'Hidden!A1');
  assert.match(new TextDecoder().decode(shifted.files['xl/worksheets/sheet1.xml']), /<hyperlink ref="A2" r:id="rId1"/);
  const removed = transformStructure(original, { sheetIndex: 0, axis: 'row', action: 'delete', index: 0 }).model;
  assert.equal(removed.sheets[0].hyperlinks.length, 0);
  assert.doesNotMatch(new TextDecoder().decode(removed.files['xl/worksheets/_rels/sheet1.xml.rels']), /rId1/);
  const targetFiles = { ...files };
  targetFiles['xl/worksheets/sheet2.xml'] = encode(new TextDecoder().decode(targetFiles['xl/worksheets/sheet2.xml']).replace('</worksheet>', '<hyperlinks><hyperlink ref="C1" location="Data!$A$1"/></hyperlinks></worksheet>'));
  const internalTarget = analyze(zipSync(targetFiles));
  const movedTarget = transformStructure(internalTarget, { sheetIndex: 0, axis: 'row', action: 'insert', index: 0 }).model;
  assert.equal(movedTarget.sheets[1].hyperlinks[0].location, 'Data!$A$2');
  assert.throws(() => transformStructure(internalTarget, { sheetIndex: 0, axis: 'row', action: 'delete', index: 0 }), /내부 하이퍼링크/);
  assert.throws(() => patchSheetName(original, 0, 'Renamed'), /내부 하이퍼링크/);
  assert.throws(() => patchDeleteSheet(original, 1), /내부 하이퍼링크/);
  const edited = patchCell(original, 0, 'A1', 'updated');
  assert.equal(edited.model.writable, true, edited.model.readOnlyReason);
  assert.match(new TextDecoder().decode(edited.model.files['xl/worksheets/sheet1.xml']), /<hyperlink ref="A1" r:id="rId1" tooltip="Open source" display="Source label"\/>/);
  assert.match(new TextDecoder().decode(edited.model.files['xl/worksheets/sheet1.xml']), /<hyperlink ref="B1" location="Hidden!A1" display="Internal"\/>/);
  assert.deepEqual(Buffer.from(edited.model.files['xl/worksheets/_rels/sheet1.xml.rels']), Buffer.from(files['xl/worksheets/_rels/sheet1.xml.rels']));
  assert.equal(edited.model.sheets[0].cells.get('A1').value, 'updated');
});

test('잘못된 하이퍼링크 관계가 있는 XLSX는 읽기 전용으로 유지한다', () => {
  const files = unzipSync(fixture());
  const decode = name => new TextDecoder().decode(files[name]);
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(decode('xl/worksheets/sheet1.xml')
    .replace('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
    .replace('</sheetData>', '</sheetData><hyperlinks><hyperlink ref="A1" r:id="missing"/></hyperlinks>'));
  const model = analyze(zipSync(files));
  assert.equal(model.writable, false);
  assert.match(model.readOnlyReason, /하이퍼링크/);
});

test('시트 이름 변경은 통합 문서·수식·정의된 이름을 함께 갱신하고 검증한다', () => {
  const files = unzipSync(fixture({ formula: true }));
  const decode = name => new TextDecoder().decode(files[name]);
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(decode('xl/worksheets/sheet1.xml').replace('<f>1+1</f>', '<f>Hidden!A1+\'Data\'!A1+Data!B1+"Data!A1"</f>'));
  files['xl/worksheets/sheet2.xml'] = new TextEncoder().encode(decode('xl/worksheets/sheet2.xml').replace('</sheetData>', '<row r="2"><c r="A2"><f>Data!A1</f><v>42</v></c></row></sheetData>'));
  files['xl/workbook.xml'] = new TextEncoder().encode(decode('xl/workbook.xml').replace('</workbook>', '<definedNames><definedName name="NameRef">\'Data\'!$A$1</definedName></definedNames></workbook>'));
  const original = analyze(zipSync(files)), renamed = patchSheetName(original, 0, 'Quarter 2');
  assert.equal(renamed.model.writable, true, renamed.model.readOnlyReason);
  assert.equal(renamed.model.sheets[0].name, 'Quarter 2');
  assert.equal(renamed.model.sheets[0].cells.get('A2').formula, "Hidden!A1+'Quarter 2'!A1+'Quarter 2'!B1+\"Data!A1\"");
  assert.equal(renamed.model.sheets[1].cells.get('A2').formula, "'Quarter 2'!A1");
  assert.match(decode('xl/workbook.xml'), /definedName name="NameRef">'Data'/);
  const workbook = new TextDecoder().decode(renamed.model.files['xl/workbook.xml']);
  assert.match(workbook, /name="Quarter 2"/);
  assert.match(workbook, /definedName name="NameRef">&apos;Quarter 2&apos;!\$A\$1/);
  assert.throws(() => patchSheetName(original, 0, 'Hidden'), /이미 있습니다/);
  for (const invalid of ['', "'bad'", 'bad/name', 'x'.repeat(32)]) assert.throws(() => patchSheetName(original, 0, invalid));
});

test('차트가 참조하는 시트 이름을 변경하면 차트 수식과 참조 캐시를 보존한다', () => {
  const original = analyze(fixtureWithCachedChart());
  assert.equal(original.writable, true, original.readOnlyReason);
  assert.equal(original.sheets[0].charts.length, 1);
  const chartPath = original.sheets[0].chartPaths[0], originalChartXml = new TextDecoder().decode(original.files[chartPath]);
  assert.match(originalChartXml, /Data!/);
  const renamed = patchSheetName(original, 0, 'Quarter 2').model;
  const chartXml = new TextDecoder().decode(renamed.files[chartPath]);
  assert.match(chartXml, /Quarter 2/);
  assert.doesNotMatch(chartXml, /Data!/);
  assert.deepEqual(renamed.sheets[0].charts, original.sheets[0].charts);
  assert.deepEqual(renamed.files['xl/drawings/drawing1.xml'], original.files['xl/drawings/drawing1.xml']);
});

test('새 워크시트는 관계·콘텐츠 형식·활성 탭을 갱신하고 다시 분석할 수 있다', () => {
  const original = analyze(fixture()), added = patchAddSheet(original);
  assert.equal(added.model.writable, true, added.model.readOnlyReason);
  assert.equal(added.sheetIndex, 2);
  assert.equal(added.model.sheets[2].name, 'Sheet3');
  assert.equal(added.model.sheets[2].cells.size, 0);
  assert.equal(added.model.sheets[2].path, 'xl/worksheets/sheet3.xml');
  const workbook = new TextDecoder().decode(added.model.files['xl/workbook.xml']);
  assert.match(workbook, /<workbookView activeTab="2"\/>/);
  assert.match(new TextDecoder().decode(added.model.files['xl/_rels/workbook.xml.rels']), /Target="worksheets\/sheet3\.xml"/);
  assert.match(new TextDecoder().decode(added.model.files['[Content_Types].xml']), /PartName="\/xl\/worksheets\/sheet3\.xml"/);
  const extraFiles = Object.fromEntries(Object.entries(unzipSync(fixture())).map(([path, bytes]) => [path, new TextDecoder().decode(bytes)]));
  extraFiles['xl/workbook.xml'] = extraFiles['xl/workbook.xml'].replace('</sheets>', '<sheet name="Sheet3" sheetId="3" r:id="rId3"/></sheets>');
  extraFiles['xl/_rels/workbook.xml.rels'] = extraFiles['xl/_rels/workbook.xml.rels'].replace('</Relationships>', '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/></Relationships>');
  extraFiles['[Content_Types].xml'] = extraFiles['[Content_Types].xml'].replace('</Types>', '<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  extraFiles['xl/worksheets/sheet3.xml'] = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>';
  const occupied = analyze(zipSync(Object.fromEntries(Object.entries(extraFiles).map(([path, value]) => [path, new TextEncoder().encode(value)]))));
  assert.equal(patchAddSheet(occupied).model.sheets[3].name, 'Sheet4');
});

test('시트 순서 변경은 정의된 이름 범위와 활성 탭을 갱신하고 시트 데이터를 보존한다', () => {
  const files = Object.fromEntries(Object.entries(unzipSync(fixture())).map(([path, bytes]) => [path, new TextDecoder().decode(bytes)]));
  files['xl/workbook.xml'] = files['xl/workbook.xml'].replace('<sheets>', '<bookViews><workbookView activeTab="0"/></bookViews><sheets>').replace('</workbook>', '<definedNames><definedName name="LocalRef" localSheetId="0">$A$1</definedName></definedNames></workbook>');
  const original = analyze(zipSync(Object.fromEntries(Object.entries(files).map(([path, value]) => [path, new TextEncoder().encode(value)]))));
  const before = new TextDecoder().decode(original.files['xl/worksheets/sheet1.xml']);
  const moved = patchMoveSheet(original, 0, 1);
  assert.deepEqual(moved.model.sheets.map(sheet => sheet.name), ['Hidden', 'Data']);
  const workbook = new TextDecoder().decode(moved.model.files['xl/workbook.xml']);
  assert.match(workbook, /activeTab="1"/);
  assert.match(workbook, /localSheetId="1"/);
  assert.equal(new TextDecoder().decode(moved.model.files['xl/worksheets/sheet1.xml']), before);
  assert.equal(patchMoveSheet(moved.model, 0, 0).model, moved.model);
});

test('차트가 있는 통합 문서는 시트 순서를 바꿔도 차트 수식·캐시와 관계를 보존한다', () => {
  const original = analyze(fixtureWithCachedChart());
  assert.equal(original.writable, true, original.readOnlyReason);
  const chartPath = original.sheets[0].chartPaths[0], drawingPath = original.sheets[0].chartDrawingPath;
  const chartXml = Buffer.from(original.files[chartPath]), drawingXml = Buffer.from(original.files[drawingPath]);
  const originalCharts = original.sheets[0].charts;
  const moved = patchMoveSheet(original, 0, 1);
  assert.deepEqual(moved.model.sheets.map(sheet => sheet.name), ['Hidden', 'Data']);
  assert.deepEqual(moved.model.sheets[1].charts, originalCharts);
  assert.deepEqual(Buffer.from(moved.model.files[chartPath]), chartXml);
  assert.deepEqual(Buffer.from(moved.model.files[drawingPath]), drawingXml);
  assert.deepEqual(moved.model.sheets[1].chartPaths, original.sheets[0].chartPaths);
});

test('기본 XLSX 표가 있는 통합문서는 셀 편집을 허용하고 표 패키지를 보존한다', () => {
  const files = Object.fromEntries(Object.entries(unzipSync(fixture())).map(([path, bytes]) => [path, new TextDecoder().decode(bytes)]));
  files['[Content_Types].xml'] = files['[Content_Types].xml'].replace('</Types>', '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>');
  files['xl/worksheets/sheet1.xml'] = files['xl/worksheets/sheet1.xml'].replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ').replace('</worksheet>', '<tableParts count="1"><tablePart r:id="table1"/></tableParts></worksheet>');
  files['xl/worksheets/_rels/sheet1.xml.rels'] = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="table1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>';
  files['xl/tables/table1.xml'] = '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="SalesTable" displayName="SalesTable" ref="A1:C2000"><autoFilter ref="A1:C2000"/><tableColumns count="3"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Amount"/><tableColumn id="3" name="Active"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>';
  const original = analyze(zipSync(Object.fromEntries(Object.entries(files).map(([path, value]) => [path, new TextEncoder().encode(value)]))));
  assert.equal(original.writable, true, original.readOnlyReason);
  assert.equal(original.sheets[0].tables[0].name, 'SalesTable');
  assert.deepEqual(original.sheets[0].tables[0].styleInfo, { name: 'TableStyleMedium2', showFirstColumn: false, showLastColumn: false, showRowStripes: true, showColumnStripes: false });
  const styledTable = patchTableStyles(original, [{ sheetIndex: 0, tableName: 'SalesTable', styleInfo: { name: 'TableStyleLight9', showFirstColumn: true, showLastColumn: true, showRowStripes: false, showColumnStripes: true } }]);
  assert.equal(styledTable.model.writable, true, styledTable.model.readOnlyReason);
  assert.deepEqual(styledTable.model.sheets[0].tables[0].styleInfo, { name: 'TableStyleLight9', showFirstColumn: true, showLastColumn: true, showRowStripes: false, showColumnStripes: true });
  const reopenedStyledTable = analyze(styledTable.bytes);
  assert.equal(reopenedStyledTable.writable, true, reopenedStyledTable.readOnlyReason);
  assert.deepEqual(reopenedStyledTable.sheets[0].tables[0].styleInfo, styledTable.model.sheets[0].tables[0].styleInfo);
  const unstyledTable = patchTableStyles(styledTable.model, [{ sheetIndex: 0, tableName: 'SalesTable', styleInfo: { name: null, showFirstColumn: false, showLastColumn: false, showRowStripes: false, showColumnStripes: false } }]);
  assert.equal(unstyledTable.model.sheets[0].tables[0].styleInfo, null);
  assert.doesNotMatch(new TextDecoder().decode(unstyledTable.model.files['xl/tables/table1.xml']), /<tableStyleInfo\b/);
  const insertedStyle = patchTableStyles(unstyledTable.model, [{ sheetIndex: 0, tableName: 'SalesTable', styleInfo: { name: 'TableStyleDark1', showFirstColumn: false, showLastColumn: true, showRowStripes: true, showColumnStripes: false } }]);
  assert.equal(insertedStyle.model.writable, true, insertedStyle.model.readOnlyReason);
  assert.equal(analyze(insertedStyle.bytes).sheets[0].tables[0].styleInfo.name, 'TableStyleDark1');
  const insertedTableRow = transformStructure(original, { sheetIndex: 0, axis: 'row', action: 'insert', index: 1 });
  assert.equal(insertedTableRow.model.writable, true, insertedTableRow.model.readOnlyReason);
  assert.equal(insertedTableRow.model.sheets[0].tables[0].range.bottom, 2000);
  assert.match(new TextDecoder().decode(insertedTableRow.model.files['xl/tables/table1.xml']), /<table[^>]*ref="A1:C2001"/);
  assert.match(new TextDecoder().decode(insertedTableRow.model.files['xl/tables/table1.xml']), /<autoFilter ref="A1:C2001"/);
  const deletedTableRow = transformStructure(insertedTableRow.model, { sheetIndex: 0, axis: 'row', action: 'delete', index: 1 });
  assert.equal(deletedTableRow.model.sheets[0].tables[0].range.bottom, 1999);
  assert.match(new TextDecoder().decode(deletedTableRow.model.files['xl/tables/table1.xml']), /<table[^>]*ref="A1:C2000"/);
  const shiftedTable = transformStructure(original, { sheetIndex: 0, axis: 'row', action: 'insert', index: 0 });
  assert.deepEqual(shiftedTable.model.sheets[0].tables[0].range, { top: 1, bottom: 2000, left: 0, right: 2 });
  assert.match(new TextDecoder().decode(shiftedTable.model.files['xl/tables/table1.xml']), /ref="A2:C2001"/);
  const insertedTableColumn = transformStructure(original, { sheetIndex: 0, axis: 'column', action: 'insert', index: 1 });
  assert.equal(insertedTableColumn.model.writable, true, insertedTableColumn.model.readOnlyReason);
  assert.deepEqual(insertedTableColumn.model.sheets[0].tables[0].columns, ['Name', 'Column2', 'Amount', 'Active']);
  assert.equal(insertedTableColumn.model.sheets[0].tables[0].range.right, 3);
  const totalsFiles = Object.fromEntries(Object.entries(unzipSync(fixture())).map(([path, bytes]) => [path, new TextDecoder().decode(bytes)]));
  totalsFiles['[Content_Types].xml'] = totalsFiles['[Content_Types].xml'].replace('</Types>', '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>');
  totalsFiles['xl/worksheets/sheet1.xml'] = totalsFiles['xl/worksheets/sheet1.xml'].replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ').replace('</worksheet>', '<tableParts count="1"><tablePart r:id="table1"/></tableParts></worksheet>');
  totalsFiles['xl/worksheets/_rels/sheet1.xml.rels'] = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="table1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>';
  totalsFiles['xl/tables/table1.xml'] = '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="SalesTable" displayName="SalesTable" ref="A1:C3" totalsRowCount="1"><autoFilter ref="A1:C3"/><tableColumns count="3"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Amount"/><tableColumn id="3" name="Active"/></tableColumns></table>';
  const totalsModel = analyze(zipSync(Object.fromEntries(Object.entries(totalsFiles).map(([path, value]) => [path, new TextEncoder().encode(value)]))));
  assert.equal(totalsModel.writable, true, totalsModel.readOnlyReason);
  const rowBeforeTotals = transformStructure(totalsModel, { sheetIndex: 0, axis: 'row', action: 'insert', index: 2 }).model;
  assert.match(new TextDecoder().decode(rowBeforeTotals.files['xl/tables/table1.xml']), /ref="A1:C4"/);
  assert.throws(() => transformStructure(totalsModel, { sheetIndex: 0, axis: 'row', action: 'delete', index: 2 }), /머리글·합계 행 삭제/);
  const filtered = patchAutoFilters(original, [{ sheetIndex: 0, filter: { top: 0, bottom: 1999, left: 0, right: 2, columns: [{ column: 1, values: ['42'], blanks: false }] } }]);
  assert.deepEqual(filtered.model.sheets[0].tables[0].autoFilter.columns, [{ column: 1, values: ['42'], blanks: false }]);
  assert.match(new TextDecoder().decode(filtered.model.files['xl/tables/table1.xml']), /filterColumn colId="1"/);
  const saved = patchCell(original, 0, 'A1', 'Changed');
  assert.equal(saved.model.writable, true, saved.model.readOnlyReason);
  assert.deepEqual(Buffer.from(saved.model.files['xl/tables/table1.xml']), Buffer.from(original.files['xl/tables/table1.xml']));
  assert.deepEqual(Buffer.from(saved.model.files['xl/worksheets/_rels/sheet1.xml.rels']), Buffer.from(original.files['xl/worksheets/_rels/sheet1.xml.rels']));
  const deleted = patchDeleteSheet(original, 0);
  assert.equal(Object.hasOwn(deleted.model.files, 'xl/tables/table1.xml'), false);
  assert.equal(Object.hasOwn(deleted.model.files, 'xl/worksheets/_rels/sheet1.xml.rels'), false);
});

test('워크시트 삭제는 수식 참조를 #REF로 바꾸고 유효성 참조는 안전하게 거부한다', () => {
  const originalFiles = Object.fromEntries(Object.entries(unzipSync(fixture())).map(([path, bytes]) => [path, new TextDecoder().decode(bytes)]));
  originalFiles['xl/workbook.xml'] = originalFiles['xl/workbook.xml'].replace('<sheets>', '<bookViews><workbookView activeTab="1"/></bookViews><sheets>').replace('</workbook>', '<definedNames><definedName name="LocalRef" localSheetId="1">$A$1</definedName></definedNames></workbook>');
  const original = analyze(zipSync(Object.fromEntries(Object.entries(originalFiles).map(([path, value]) => [path, new TextEncoder().encode(value)]))));
  const deleted = patchDeleteSheet(original, 1);
  assert.equal(deleted.model.sheets.length, 1);
  assert.equal(deleted.model.sheets[0].name, 'Data');
  assert.equal(Object.hasOwn(deleted.model.files, 'xl/worksheets/sheet2.xml'), false);
  assert.doesNotMatch(new TextDecoder().decode(deleted.model.files['xl/_rels/workbook.xml.rels']), /rId2/);
  assert.doesNotMatch(new TextDecoder().decode(deleted.model.files['[Content_Types].xml']), /sheet2\.xml/);
  assert.match(new TextDecoder().decode(deleted.model.files['xl/workbook.xml']), /activeTab="0"/);
  assert.doesNotMatch(new TextDecoder().decode(deleted.model.files['xl/workbook.xml']), /LocalRef/);
  assert.throws(() => patchDeleteSheet(deleted.model, 0), /마지막 시트/);
  const dependentFiles = Object.fromEntries(Object.entries(unzipSync(fixture())).map(([path, bytes]) => [path, new TextDecoder().decode(bytes)]));
  dependentFiles['xl/worksheets/sheet2.xml'] = dependentFiles['xl/worksheets/sheet2.xml'].replace('</sheetData>', '<row r="2"><c r="A2"><f>Data!A1</f><v>42</v></c></row></sheetData>');
  const dependent = analyze(zipSync(Object.fromEntries(Object.entries(dependentFiles).map(([path, value]) => [path, new TextEncoder().encode(value)]))));
  const deletedReference = patchDeleteSheet(dependent, 0);
  assert.match(new TextDecoder().decode(deletedReference.model.files['xl/worksheets/sheet2.xml']), /<f>#REF!<\/f>/);
  const validationFiles = Object.fromEntries(Object.entries(unzipSync(fixture())).map(([path, bytes]) => [path, new TextDecoder().decode(bytes)]));
  validationFiles['xl/worksheets/sheet2.xml'] = validationFiles['xl/worksheets/sheet2.xml'].replace('</worksheet>', '<dataValidations count="1"><dataValidation type="list" sqref="A1"><formula1>Data!$A$1:$A$3</formula1></dataValidation></dataValidations></worksheet>');
  const validationModel = analyze(zipSync(Object.fromEntries(Object.entries(validationFiles).map(([path, value]) => [path, new TextEncoder().encode(value)]))));
  assert.throws(() => patchDeleteSheet(validationModel, 0), /데이터 유효성·조건부 서식/);
});

test('수식 복사는 상대 참조만 이동하고 절대 참조·문자열·표 구조화 참조를 보존한다', () => {
  assert.equal(
    translateFormulaReferences('SUM(A1:B2)+$C3+D$4+$E$5+"A1"&\'Sales A1\'!A1+Table1[A1]+LOG10(A1)', 1, 2),
    'SUM(C2:D3)+$C4+F$4+$E$5+"A1"&\'Sales A1\'!C2+Table1[A1]+LOG10(C2)'
  );
  assert.equal(translateFormulaReferences('A1+$B2+C$3+$D$4', -1, -1), '#REF!+$B1+B$3+$D$4');
});

test('XLSX 행 높이·기본 높이·줄 바꿈 서식을 읽고 셀 저장 후 보존한다', () => {
  const files = unzipSync(fixture({ formatted: true })), sheetPath = 'xl/worksheets/sheet1.xml';
  files[sheetPath] = new TextEncoder().encode(new TextDecoder().decode(files[sheetPath]).replace('<sheetData>', '<sheetFormatPr defaultRowHeight="18"/><sheetData>').replace('<row r="1">', '<row r="1" ht="45" customHeight="1">'));
  files['xl/styles.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/styles.xml']).replace('<xf numFmtId="14"/>', '<xf numFmtId="14" applyAlignment="1"><alignment wrapText="1"/></xf>'));
  const model = analyze(zipSync(files));
  assert.equal(model.writable, true, model.readOnlyReason);
  assert.deepEqual(model.sheets[0].rowHeights, { 0: 60 }); assert.equal(model.sheets[0].defaultRowHeight, 24);
  assert.equal(model.sheets[0].cells.get('B1').styleProperties.wrapText, true);
  const edited = patchCell(model, 0, 'A1', 'changed');
  assert.deepEqual(edited.model.sheets[0].rowHeights, { 0: 60 });
  assert.match(edited.model.sheets[0].xml, /ht="45" customHeight="1"/);
  assert.deepEqual(Buffer.from(edited.model.files['xl/styles.xml']), Buffer.from(model.files['xl/styles.xml']));
  const invalid = { ...files, [sheetPath]: new TextEncoder().encode(new TextDecoder().decode(files[sheetPath]).replace('ht="45"', 'ht="500"')) };
  assert.equal(analyze(zipSync(invalid)).writable, false);
  const zero = { ...files, [sheetPath]: new TextEncoder().encode(new TextDecoder().decode(files[sheetPath]).replace('ht="45"', 'ht="0"')) };
  const zeroModel = analyze(zipSync(zero));
  assert.deepEqual(zeroModel.sheets[0].hiddenRows, [0]);
  const shown = patchHiddenRows(zeroModel, [{ sheetIndex: 0, rows: [0], hidden: false }]).model;
  assert.deepEqual(shown.sheets[0].hiddenRows, []); assert.deepEqual(shown.sheets[0].rowHeights, {});
});

test('행 높이를 선택 범위와 빈 행에 적용하고 기본 높이로 복원한다', () => {
  const original = analyze(fixture());
  const changed = patchRowHeights(original, 0, 0, 2, 30);
  assert.deepEqual(changed.model.sheets[0].rowHeights, { 0: 40, 1: 40, 2: 40 });
  assert.equal(changed.model.sheets[0].cells.get('A1').value, 'before');
  for (const [name, bytes] of Object.entries(original.files)) if (name !== 'xl/worksheets/sheet1.xml') assert.deepEqual(Buffer.from(changed.model.files[name]), Buffer.from(bytes));
  assert.equal(patchRowHeights(changed.model, 0, 0, 2, 30).model, changed.model);
  const reset = patchRowHeights(changed.model, 0, 1, 2, null);
  assert.deepEqual(analyze(reset.bytes).sheets[0].rowHeights, { 0: 40 });
  const files = { ...original.files, 'xl/worksheets/sheet1.xml': new TextEncoder().encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>') };
  assert.deepEqual(patchRowHeights(analyze(zipSync(files)), 0, 2, 3, 45).model.sheets[0].rowHeights, { 2: 60, 3: 60 });
  assert.throws(() => patchRowHeights(original, 0, 0, 10000, 20), /10,000행/);
  for (const height of [0, -1, 410, Infinity, NaN]) assert.throws(() => patchRowHeights(original, 0, 0, 0, height), /행 높이/);
});

test('전체 행 높이는 기본 높이로 저장하고 숨김·기존 셀을 보존한다', () => {
  const files = unzipSync(fixture());
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('<row r="1">', '<row r="1" ht="0" customHeight="1">'));
  const original = analyze(zipSync(files));
  const changed = patchRowHeights(original, 0, 0, 1048575, 24).model;
  assert.equal(changed.sheets[0].defaultRowHeight, 32); assert.deepEqual(changed.sheets[0].rowHeights, {});
  assert.deepEqual(changed.sheets[0].hiddenRows, [0]); assert.equal(changed.sheets[0].cells.get('B1').value, '42');
  assert.equal((changed.sheets[0].xml.match(/<row\b/g) || []).length, 1, 'full-axis formatting does not materialize a million row elements');
});

test('행 자동 맞춤 높이는 희소 행을 일괄 저장하고 수동 높이 표시를 해제한다', () => {
  const model = analyze(fixture());
  const patched = patchRowHeightValues(model, 0, [{ row: 0, height: 45 }, { row: 999999, height: 60 }]);
  assert.deepEqual(patched.model.sheets[0].rowHeights, { 0: 60, 999999: 80 });
  assert.deepEqual(patched.model.sheets[0].customRowHeights, []);
  assert.equal((patched.model.sheets[0].xml.match(/<row\b/g) || []).length, 2);
  assert.throws(() => patchRowHeightValues(model, 0, [{ row: 1, height: 30 }, { row: 0, height: 30 }]), /자동 맞춤 결과/);
});

test('줄 바꿈 서식은 기본 스타일 파일을 추가하고 다른 맞춤·글꼴을 보존한다', () => {
  const original = analyze(fixture());
  const added = patchCells(original, [{ sheetIndex: 0, address: 'A1', value: 'before', alignment: { wrapText: true } }]);
  assert.ok(added.model.stylesPath); assert.equal(added.model.sheets[0].cells.get('A1').styleProperties.wrapText, true);
  assert.deepEqual(Buffer.from(added.model.files['xl/worksheets/sheet2.xml']), Buffer.from(original.files['xl/worksheets/sheet2.xml']));
  const styled = analyze(fixture({ formatted: true }));
  const changed = patchCells(styled, [{ sheetIndex: 0, address: 'D1', value: '45292', alignment: { wrapText: true } }]).model;
  const properties = changed.sheets[0].cells.get('D1').styleProperties;
  assert.equal(properties.wrapText, true); assert.equal(properties.bold, true); assert.equal(properties.horizontal, 'center');
  const cleared = patchCells(changed, [{ sheetIndex: 0, address: 'D1', value: '45292', alignment: { wrapText: false } }]).model;
  assert.equal(Boolean(cleared.sheets[0].cells.get('D1').styleProperties.wrapText), false);
});

test('XLSX freeze panes are displayed and preserved when saving an edited cell', () => {
  const paneXml = '<sheetViews><sheetView workbookViewId="0"><pane xSplit="2" ySplit="3" topLeftCell="C4" activePane="bottomRight" state="frozen"/><selection pane="bottomRight" activeCell="C4" sqref="C4"/></sheetView></sheetViews>';
  const model = analyze(fixture({ sheetViews: paneXml }));
  assert.equal(model.writable, true, model.readOnlyReason);
  assert.deepEqual(model.sheets[0].freezePane, { rows: 3, columns: 2 });
  const saved = patchCell(model, 0, 'A1', 'changed');
  const reopened = analyze(saved.bytes);
  assert.equal(reopened.writable, true, reopened.readOnlyReason);
  assert.deepEqual(reopened.sheets[0].freezePane, { rows: 3, columns: 2 });
  assert.ok(reopened.sheets[0].xml.includes(paneXml));
  assert.deepEqual(Buffer.from(reopened.files['xl/worksheets/sheet2.xml']), Buffer.from(model.files['xl/worksheets/sheet2.xml']));
});

test('XLSX 고정 창은 기존 시트 보기를 보존하며 적용·해제한다', () => {
  const views = '<sheetViews><sheetView workbookViewId="0" showGridLines="0"><selection activeCell="B2" sqref="B2"/></sheetView></sheetViews>';
  const model = analyze(fixture({ sheetViews: views }));
  const frozen = patchFreezePanes(model, [{ sheetIndex: 0, freezePane: { rows: 2, columns: 1 } }]);
  assert.deepEqual(frozen.model.sheets[0].freezePane, { rows: 2, columns: 1 });
  assert.match(frozen.model.sheets[0].xml, /<sheetView workbookViewId="0" showGridLines="0"><pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"\/><selection pane="bottomRight" activeCell="B3" sqref="B3"\/>/);
  const unfrozen = patchFreezePanes(frozen.model, [{ sheetIndex: 0, freezePane: { rows: 0, columns: 0 } }]);
  assert.deepEqual(unfrozen.model.sheets[0].freezePane, { rows: 0, columns: 0 });
  assert.doesNotMatch(unfrozen.model.sheets[0].xml, /<pane\b|<selection\b/);
  const created = patchFreezePanes(analyze(fixture()), [{ sheetIndex: 0, freezePane: { rows: 1, columns: 0 } }]);
  assert.deepEqual(created.model.sheets[0].freezePane, { rows: 1, columns: 0 });
  assert.throws(() => patchFreezePanes(model, [{ sheetIndex: 0, freezePane: { rows: 51, columns: 0 } }]), /요청이 유효하지 않습니다/);
});

test('XLSX 수동 숨김 행을 파싱하고 셀 편집 저장에서 보존한다', () => {
  const files = unzipSync(fixture({ formula: true }));
  let xml = new TextDecoder().decode(files['xl/worksheets/sheet1.xml']);
  xml = xml.replace('<row r="1">', '<row r="1" hidden="1">').replace('<row r="2">', '<row r="2" hidden="true">');
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(xml);
  const model = analyze(zipSync(files));
  assert.equal(model.writable, true, model.readOnlyReason);
  assert.deepEqual(model.sheets[0].hiddenRows, [0, 1]);
  const saved = patchCell(model, 0, 'B1', 'changed');
  assert.deepEqual(saved.model.sheets[0].hiddenRows, [0, 1]);
  assert.match(saved.model.sheets[0].xml, /<row r="1" hidden="1">/);
  assert.match(saved.model.sheets[0].xml, /<row r="2" hidden="true">/);

  const malformed = { ...files, 'xl/worksheets/sheet1.xml': new TextEncoder().encode(xml.replace('hidden="1"', 'hidden="maybe"')) };
  assert.equal(analyze(zipSync(malformed)).writable, false, 'unknown hidden-row values are not silently misinterpreted');
});

test('XLSX 행·열 개요 수준과 접힌 그룹 범위를 읽는다', () => {
  const files = unzipSync(fixture({
    sheetViews: '<sheetPr><outlinePr summaryBelow="1" summaryRight="1"/></sheetPr>',
    columnDefinitions: '<cols><col min="1" max="2" hidden="1" outlineLevel="1"/><col min="3" max="3" collapsed="1"/></cols>'
  }));
  let xml = new TextDecoder().decode(files['xl/worksheets/sheet1.xml']);
  xml = xml.replace('<row r="1">', '<row r="1" hidden="1" outlineLevel="1">')
    .replace('</row></sheetData>', '</row><row r="2" hidden="1" outlineLevel="1"><c r="A2"><v>2</v></c></row><row r="3" collapsed="1"><c r="A3"><v>3</v></c></row></sheetData>');
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(xml);
  const model = analyze(zipSync(files));
  assert.equal(model.writable, true, model.readOnlyReason);
  assert.deepEqual(model.sheets[0].rowOutlineLevels, { 0: 1, 1: 1 });
  assert.deepEqual(model.sheets[0].columnOutlineLevels, { 0: 1, 1: 1 });
  assert.deepEqual(model.sheets[0].rowGroups, [{ start: 0, end: 1, level: 1, summary: 2, collapsed: true }]);
  assert.deepEqual(model.sheets[0].columnGroups, [{ start: 0, end: 1, level: 1, summary: 2, collapsed: true }]);
  assert.deepEqual(model.sheets[0].outlineHiddenRows, [0, 1]);
  assert.deepEqual(model.sheets[0].outlineHiddenColumns, [0, 1]);
  assert.deepEqual(model.sheets[0].hiddenColumns, [], 'outline-hidden columns are not treated as manually hidden');
});

test('행 그룹은 접기·펼치기·해제와 수동 숨김을 구분해 저장한다', () => {
  const manuallyHidden = patchHiddenRows(analyze(fixture()), [{ sheetIndex: 0, rows: [0], hidden: true }]).model;
  const grouped = patchOutlineGroups(manuallyHidden, [{ sheetIndex: 0, axis: 'row', action: 'group', start: 0, end: 1 }]).model;
  assert.deepEqual(grouped.sheets[0].rowGroups, [{ start: 0, end: 1, level: 1, summary: 2, collapsed: false }]);
  assert.deepEqual(grouped.sheets[0].hiddenRows, [0], 'grouping preserves an existing manual hide');

  const collapsed = patchOutlineGroups(grouped, [{ sheetIndex: 0, axis: 'row', action: 'collapse', start: 0, end: 1 }]);
  assert.deepEqual(collapsed.model.sheets[0].outlineHiddenRows, [0, 1]);
  assert.deepEqual(collapsed.model.sheets[0].hiddenRows, [0], 'the overlapping manual hide remains distinguishable');
  const expanded = patchOutlineGroups(collapsed.model, [{ sheetIndex: 0, axis: 'row', action: 'expand', start: 0, end: 1 }]);
  assert.deepEqual(expanded.model.sheets[0].outlineHiddenRows, []);
  assert.deepEqual(expanded.model.sheets[0].hiddenRows, [0]);
  const ungrouped = patchOutlineGroups(expanded.model, [{ sheetIndex: 0, axis: 'row', action: 'ungroup', start: 0, end: 1 }]);
  assert.deepEqual(ungrouped.model.sheets[0].rowGroups, []);
  assert.deepEqual(ungrouped.model.sheets[0].hiddenRows, [0]);
  const reopened = analyze(ungrouped.bytes);
  assert.equal(reopened.writable, true, reopened.readOnlyReason);
  assert.deepEqual(reopened.sheets[0].hiddenRows, [0]);
  assert.deepEqual(reopened.sheets[0].rowOutlineLevels, {});
});

test('열 그룹을 접고 펼쳐도 수동 숨김과 기존 접힌 그룹 표시를 분리한다', () => {
  const files = unzipSync(fixture({ columnDefinitions: '<cols><col min="1" max="3" width="10" customWidth="1"/></cols>' }));
  const hidden = patchHiddenColumns(analyze(zipSync(files)), [{ sheetIndex: 0, columns: [2], hidden: true }]).model;
  const grouped = patchOutlineGroups(hidden, [{ sheetIndex: 0, axis: 'column', action: 'group', start: 0, end: 1 }]).model;
  assert.deepEqual(grouped.sheets[0].hiddenColumns, [2]);
  const collapsed = patchOutlineGroups(grouped, [{ sheetIndex: 0, axis: 'column', action: 'collapse', start: 0, end: 1 }]);
  assert.deepEqual(collapsed.model.sheets[0].outlineHiddenColumns, [0, 1]);
  assert.deepEqual(collapsed.model.sheets[0].hiddenColumns, [2]);
  assert.deepEqual(collapsed.model.sheets[0].nativeHiddenColumns, [0, 1, 2]);
  const expanded = patchOutlineGroups(collapsed.model, [{ sheetIndex: 0, axis: 'column', action: 'expand', start: 0, end: 1 }]);
  assert.deepEqual(expanded.model.sheets[0].outlineHiddenColumns, []);
  assert.deepEqual(expanded.model.sheets[0].hiddenColumns, [2]);
  assert.deepEqual(expanded.model.sheets[0].nativeHiddenColumns, [2]);
  const reopened = analyze(expanded.bytes);
  assert.deepEqual(reopened.sheets[0].hiddenColumns, [2]);
  assert.deepEqual(reopened.sheets[0].columnOutlineLevels, { 0: 1, 1: 1 });
});

test('같은 요약 행·열을 공유하는 중첩 그룹은 함께 접고 펼쳐 XLSX에 보존한다', () => {
  let model = analyze(fixture());
  for (const axis of ['row', 'column']) {
    model = patchOutlineGroups(model, [{ sheetIndex: 0, axis, action: 'group', start: 0, end: 1 }]).model;
    model = patchOutlineGroups(model, [{ sheetIndex: 0, axis, action: 'group', start: 1, end: 1 }]).model;
  }
  const expectedExpandedGroups = [
    { start: 0, end: 1, level: 1, summary: 2, collapsed: false },
    { start: 1, end: 1, level: 2, summary: 2, collapsed: false }
  ];
  assert.deepEqual(model.sheets[0].rowGroups, expectedExpandedGroups);
  assert.deepEqual(model.sheets[0].columnGroups, expectedExpandedGroups);

  const rowsCollapsed = patchOutlineGroups(model, [{ sheetIndex: 0, axis: 'row', action: 'collapse-summary', summary: 2 }]);
  const bothCollapsed = patchOutlineGroups(rowsCollapsed.model, [{ sheetIndex: 0, axis: 'column', action: 'collapse-summary', summary: 2 }]);
  assert.deepEqual(bothCollapsed.model.sheets[0].rowGroups.map(group => group.collapsed), [true, true]);
  assert.deepEqual(bothCollapsed.model.sheets[0].columnGroups.map(group => group.collapsed), [true, true]);
  assert.deepEqual(bothCollapsed.model.sheets[0].outlineHiddenRows, [0, 1]);
  assert.deepEqual(bothCollapsed.model.sheets[0].outlineHiddenColumns, [0, 1]);

  const reopened = analyze(bothCollapsed.bytes);
  assert.deepEqual(reopened.sheets[0].rowGroups.map(group => group.collapsed), [true, true]);
  assert.deepEqual(reopened.sheets[0].columnGroups.map(group => group.collapsed), [true, true]);
  const rowsExpanded = patchOutlineGroups(reopened, [{ sheetIndex: 0, axis: 'row', action: 'expand-summary', summary: 2 }]);
  const bothExpanded = patchOutlineGroups(rowsExpanded.model, [{ sheetIndex: 0, axis: 'column', action: 'expand-summary', summary: 2 }]);
  assert.deepEqual(bothExpanded.model.sheets[0].rowGroups, expectedExpandedGroups);
  assert.deepEqual(bothExpanded.model.sheets[0].columnGroups, expectedExpandedGroups);
});

test('행·열 표시 수준은 해당 수준까지의 그룹 상태를 저장한다', () => {
  let model = analyze(fixture());
  for (const axis of ['row', 'column']) {
    model = patchOutlineGroups(model, [{ sheetIndex: 0, axis, action: 'group', start: 0, end: 2 }]).model;
    model = patchOutlineGroups(model, [{ sheetIndex: 0, axis, action: 'group', start: 0, end: 1 }]).model;
  }
  const rowsLevel1 = patchOutlineGroups(model, [{ sheetIndex: 0, axis: 'row', action: 'show-level', level: 1 }]);
  const bothLevel1 = patchOutlineGroups(rowsLevel1.model, [{ sheetIndex: 0, axis: 'column', action: 'show-level', level: 1 }]);
  assert.deepEqual(bothLevel1.model.sheets[0].rowGroups.map(group => group.collapsed), [true, true]);
  assert.deepEqual(bothLevel1.model.sheets[0].columnGroups.map(group => group.collapsed), [true, true]);
  assert.deepEqual(bothLevel1.model.sheets[0].outlineHiddenRows, [0, 1, 2]);
  assert.deepEqual(bothLevel1.model.sheets[0].outlineHiddenColumns, [0, 1, 2]);

  const rowsLevel2 = patchOutlineGroups(bothLevel1.model, [{ sheetIndex: 0, axis: 'row', action: 'show-level', level: 2 }]);
  const bothLevel2 = patchOutlineGroups(rowsLevel2.model, [{ sheetIndex: 0, axis: 'column', action: 'show-level', level: 2 }]);
  const collapseStateByLevel = groups => Object.fromEntries(groups.map(group => [group.level, group.collapsed]));
  assert.deepEqual(collapseStateByLevel(bothLevel2.model.sheets[0].rowGroups), { 1: false, 2: true });
  assert.deepEqual(collapseStateByLevel(bothLevel2.model.sheets[0].columnGroups), { 1: false, 2: true });
  assert.deepEqual(bothLevel2.model.sheets[0].outlineHiddenRows, [0, 1]);
  assert.deepEqual(bothLevel2.model.sheets[0].outlineHiddenColumns, [0, 1]);

  const rowsLevel3 = patchOutlineGroups(bothLevel2.model, [{ sheetIndex: 0, axis: 'row', action: 'show-level', level: 3 }]);
  const bothLevel3 = patchOutlineGroups(rowsLevel3.model, [{ sheetIndex: 0, axis: 'column', action: 'show-level', level: 3 }]);
  assert.deepEqual(collapseStateByLevel(bothLevel3.model.sheets[0].rowGroups), { 1: false, 2: false });
  assert.deepEqual(collapseStateByLevel(bothLevel3.model.sheets[0].columnGroups), { 1: false, 2: false });
  assert.deepEqual(bothLevel3.model.sheets[0].outlineHiddenRows, []);
  assert.deepEqual(bothLevel3.model.sheets[0].outlineHiddenColumns, []);
  assert.equal(patchOutlineGroups(bothLevel3.model, [{ sheetIndex: 0, axis: 'row', action: 'show-level', level: 3 }]).model, bothLevel3.model, 'reselecting the current level is a no-op');
  assert.throws(() => patchOutlineGroups(model, [{ sheetIndex: 0, axis: 'row', action: 'show-level', level: 9 }]), /유효하지 않습니다/);
});

test('XLSX 병합 셀은 표준 위치에 기록하고 겹침을 거부하며 해제도 지원한다', () => {
  const base = analyze(fixture()), xml = new TextDecoder().decode(base.files[base.sheets[0].path]);
  const withFilter = analyze(zipSync({ ...base.files, [base.sheets[0].path]: new TextEncoder().encode(xml.replace('</sheetData>', '</sheetData><autoFilter ref="A1:C2"/>')) }));
  const merged = patchMergedCells(withFilter, [{ sheetIndex: 0, ranges: [{ top: 0, bottom: 0, left: 0, right: 1 }] }]);
  assert.equal(merged.model.writable, true, merged.model.readOnlyReason);
  assert.deepEqual(merged.model.sheets[0].mergedCells.map(({ top, bottom, left, right }) => ({ top, bottom, left, right })), [{ top: 0, bottom: 0, left: 0, right: 1 }]);
  assert.ok(merged.model.sheets[0].xml.indexOf('<autoFilter') < merged.model.sheets[0].xml.indexOf('<mergeCells'));
  assert.match(merged.model.sheets[0].xml, /<mergeCells count="1"><mergeCell ref="A1:B1"\/><\/mergeCells>/);
  assert.throws(() => patchMergedCells(withFilter, [{ sheetIndex: 0, ranges: [{ top: 0, bottom: 1, left: 0, right: 1 }, { top: 1, bottom: 2, left: 1, right: 2 }] }]), /서로 겹칩니다/);
  const unmerged = patchMergedCells(merged.model, [{ sheetIndex: 0, ranges: [] }]);
  assert.equal(unmerged.model.sheets[0].mergedCells.length, 0);
  assert.doesNotMatch(unmerged.model.sheets[0].xml, /<mergeCells\b/);
});

test('XLSX 병합 셀은 구조 변경 위치에 따라 참조를 이동하거나 범위 교차를 거부한다', () => {
  const files = unzipSync(fixture()), source = new TextDecoder().decode(files['xl/worksheets/sheet1.xml']);
  const mergedXml = '<mergeCells count="2"><mergeCell ref="A1:B2"/><mergeCell ref="$D$4:$E$4"/></mergeCells>';
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(source.replace('</worksheet>', mergedXml + '</worksheet>'));
  const model = analyze(zipSync(files));
  assert.equal(model.writable, true, model.readOnlyReason);
  assert.deepEqual(model.sheets[0].mergedCells, [
    { top: 0, bottom: 1, left: 0, right: 1, ref: 'A1:B2' },
    { top: 3, bottom: 3, left: 3, right: 4, ref: '$D$4:$E$4' }
  ]);
  const saved = patchCell(model, 0, 'A1', 'merged anchor');
  assert.equal(saved.model.writable, true, saved.model.readOnlyReason);
  assert.deepEqual(saved.model.sheets[0].mergedCells, model.sheets[0].mergedCells);
  assert.ok(saved.model.sheets[0].xml.includes(mergedXml));
  const rowInserted = transformStructure(saved.model, { sheetIndex: 0, axis: 'row', action: 'insert', index: 0 }).model;
  assert.deepEqual(rowInserted.sheets[0].mergedCells, [
    { top: 1, bottom: 2, left: 0, right: 1, ref: 'A2:B3' },
    { top: 4, bottom: 4, left: 3, right: 4, ref: '$D$5:$E$5' }
  ]);
  assert.throws(() => transformStructure(saved.model, { sheetIndex: 0, axis: 'row', action: 'insert', index: 1 }), /병합 범위를 가로지르는/);
  assert.throws(() => transformStructure(saved.model, { sheetIndex: 0, axis: 'column', action: 'delete', index: 1 }), /병합 범위를 가로지르는/);
  const restored = transformStructure(rowInserted, { sheetIndex: 0, axis: 'row', action: 'delete', index: 0 }).model;
  assert.deepEqual(restored.sheets[0].mergedCells, saved.model.sheets[0].mergedCells);

  for (const invalid of [
    '<mergeCells count="2"><mergeCell ref="A1:B2"/></mergeCells>',
    '<mergeCells count="1"><mergeCell ref="B2:A1"/></mergeCells>',
    '<mergeCells count="1"><mergeCell ref="A1"/></mergeCells>',
    '<mergeCells count="2"><mergeCell ref="A1:B2"/><mergeCell ref="B2:C3"/></mergeCells>'
  ]) {
    const badFiles = { ...files, 'xl/worksheets/sheet1.xml': new TextEncoder().encode(source.replace('</worksheet>', invalid + '</worksheet>')) };
    const unsupported = analyze(zipSync(badFiles));
    assert.equal(unsupported.writable, false, invalid);
    assert.match(unsupported.readOnlyReason, /병합 셀/);
  }
});

test('XLSX cell extensions and rich inline strings are read-only before an edit can discard them', () => {
  const extensionModel = analyze(fixture({ cellExtension: '<extLst><ext uri="urn:test"/></extLst>' }));
  assert.equal(extensionModel.writable, false);
  assert.match(extensionModel.readOnlyReason, /지원하지 않는 셀 구성요소/);
  assert.throws(() => patchCell(extensionModel, 0, 'A1', 'changed'), /지원하지 않는 셀 구성요소/);

  const files = unzipSync(fixture());
  const worksheet = new TextDecoder().decode(files['xl/worksheets/sheet1.xml']);
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(worksheet.replace('<is><t>before</t></is>', '<is><r><rPr><b/></rPr><t>before</t></r></is>'));
  const richTextModel = analyze(zipSync(files));
  assert.equal(richTextModel.writable, false);
  assert.match(richTextModel.readOnlyReason, /서식 또는 확장이 포함된 인라인 문자열/);
});

test('XLSX split panes, malformed freeze panes and oversized frozen regions are read-only', () => {
  const cases = [
    ['<sheetViews><sheetView><pane xSplit="1" ySplit="0" topLeftCell="B1" state="split"/></sheetView></sheetViews>', /분할 창은 고정 창/],
    ['<sheetViews><sheetView><pane xSplit="1.5" ySplit="0" topLeftCell="B1" state="frozen"/></sheetView></sheetViews>', /소수 또는 잘못된 고정 창 범위/],
    ['<sheetViews><sheetView><pane xSplit="0" ySplit="51" topLeftCell="A52" state="frozen"/></sheetView></sheetViews>', /화면 렌더링 한도/],
    ['<sheetViews><sheetView><pane xSplit="1" ySplit="0" topLeftCell="?1" state="frozen"/></sheetView></sheetViews>', /시작 셀이 시트 범위 또는 고정 영역/],
    ['<sheetViews><sheetView><pane xSplit="1" ySplit="0" topLeftCell="A1" state="frozen"/></sheetView></sheetViews>', /시작 셀이 시트 범위 또는 고정 영역/],
    ['<sheetViews><sheetView><pane xSplit="1" ySplit="0" topLeftCell="XFE1" state="frozen"/></sheetView></sheetViews>', /시작 셀이 시트 범위 또는 고정 영역/] ,
    ['<sheetViews><sheetView><pane xSplit="0" ySplit="1" topLeftCell="A1048577" state="frozen"/></sheetView></sheetViews>', /시작 셀이 시트 범위 또는 고정 영역/],
    ['<sheetViews><sheetView><pane xSplit="1" ySplit="0" topLeftCell="B1" state="frozen"/><pane xSplit="2" ySplit="0" topLeftCell="C1" state="frozen"/></sheetView></sheetViews>', /여러 분할 창/]
  ];
  for (const [sheetViews, reason] of cases) {
    const model = analyze(fixture({ sheetViews }));
    assert.equal(model.writable, false);
    assert.match(model.readOnlyReason, reason);
  }
});

test('XLSX column width spans are parsed as bounded zero-based metadata', () => {
  const model = analyze(fixture({ columnDefinitions: '<cols><col min="1" max="2" width="24" customWidth="1"/><col min="4" max="4" width="9.5"/></cols>' }));
  assert.deepEqual(model.sheets[0].columnWidths, { 0: 24, 1: 24, 3: 9.5 });
  assert.equal(model.writable, true, model.readOnlyReason);
  const saved = patchCell(model, 0, 'A1', 'after');
  const reopened = analyze(saved.bytes);
  assert.deepEqual(reopened.sheets[0].columnWidths, { 0: 24, 1: 24, 3: 9.5 });
  assert.match(reopened.sheets[0].xml, /<cols><col min="1" max="2" width="24"/);
  assert.throws(() => analyze(fixture({ columnDefinitions: '<cols><col min="0" max="1" width="12"/></cols>' })), /열 너비 범위/);
  assert.throws(() => analyze(fixture({ columnDefinitions: '<cols><col min="1" max="1" width="256"/></cols>' })), /열 너비 값/);
  assert.throws(() => analyze(fixture({ columnDefinitions: '<cols><col min="1" max="1" width=""/></cols>' })), /열 너비 값/);
});

test('XLSX 열 삽입·삭제는 저장된 열 너비·숨김·스타일 범위를 함께 갱신한다', () => {
  const original = analyze(fixture({ columnDefinitions: '<cols><col min="1" max="3" width="24" customWidth="1" hidden="1" style="0"/><col min="5" max="5" width="9.5"/></cols>' }));
  assert.equal(original.writable, true, original.readOnlyReason);
  const insertedInside = transformStructure(original, { sheetIndex: 0, axis: 'column', action: 'insert', index: 1 }).model;
  assert.equal(insertedInside.writable, true, insertedInside.readOnlyReason);
  assert.deepEqual(insertedInside.sheets[0].columnWidths, { 0: 24, 1: 24, 2: 24, 3: 24, 5: 9.5 });
  assert.match(insertedInside.sheets[0].xml, /<col min="1" max="4" width="24" customWidth="1" hidden="1" style="0"\/>/);
  assert.match(insertedInside.sheets[0].xml, /<col min="6" max="6" width="9\.5"\/>/);
  const deleted = transformStructure(insertedInside, { sheetIndex: 0, axis: 'column', action: 'delete', index: 1 }).model;
  assert.deepEqual(deleted.sheets[0].columnWidths, original.sheets[0].columnWidths);
  const shifted = transformStructure(original, { sheetIndex: 0, axis: 'column', action: 'insert', index: 0 }).model;
  assert.deepEqual(shifted.sheets[0].columnWidths, { 1: 24, 2: 24, 3: 24, 5: 9.5 });
  assert.match(shifted.sheets[0].xml, /<col min="2" max="4" width="24" customWidth="1" hidden="1" style="0"\/>/);
  const removedDefinition = transformStructure(original, { sheetIndex: 0, axis: 'column', action: 'delete', index: 4 }).model;
  assert.deepEqual(removedDefinition.sheets[0].columnWidths, { 0: 24, 1: 24, 2: 24 });
  assert.doesNotMatch(removedDefinition.sheets[0].xml, /<col min="5"/);
});

test('XLSX 숫자 비교 조건부 서식은 표시 스타일을 계산하고 셀 편집에서 규칙 XML을 보존한다', () => {
  const files = unzipSync(fixture({ formatted: true }));
  files['xl/styles.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/styles.xml']).replace('</styleSheet>', '<dxfs count="2"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></dxf><dxf><font><b/><color rgb="FF008000"/></font></dxf></dxfs></styleSheet>'));
  const conditions = '<conditionalFormatting sqref="B1:C1"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>100</formula></cfRule><cfRule type="cellIs" dxfId="1" priority="2" operator="lessThan"><formula>1</formula></cfRule></conditionalFormatting>';
  let sheetXml = new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</sheetData>', '</sheetData>' + conditions);
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(sheetXml);
  const original = analyze(zipSync(files));
  assert.equal(original.writable, true, original.readOnlyReason);
  assert.equal(original.sheets[0].conditionalFormats.length, 2);
  assert.deepEqual(require('../src/xlsx').conditionalStyleForCell(original.sheets[0], 0, 1, original.sheets[0].cells.get('B1')), { fill: '#FF0000' });
  assert.deepEqual(require('../src/xlsx').conditionalStyleForCell(original.sheets[0], 0, 2, original.sheets[0].cells.get('C1')), { bold: true, color: '#008000' });
  assert.equal(require('../src/xlsx').conditionalStyleForCell(original.sheets[0], 0, 0, original.sheets[0].cells.get('A1')), null);
  const saved = patchCell(original, 0, 'A1', 'after');
  const reopened = analyze(saved.bytes);
  assert.equal(reopened.writable, true, reopened.readOnlyReason);
  assert.equal(reopened.sheets[0].conditionalFormats.length, 2);
  assert.ok(reopened.sheets[0].xml.includes(conditions));
  const inserted = transformStructure(reopened, { sheetIndex: 0, axis: 'row', action: 'insert', index: 0 });
  assert.equal(inserted.model.sheets[0].conditionalFormats.length, 2);
  assert.deepEqual(inserted.model.sheets[0].conditionalFormats[0].ranges, [{ top: 1, bottom: 1, left: 1, right: 2 }]);
  const removedRange = transformStructure(inserted.model, { sheetIndex: 0, axis: 'row', action: 'delete', index: 1 });
  assert.equal(removedRange.model.sheets[0].conditionalFormats.length, 0, 'deleting the only applied row removes its conditional rules');
  const expressionXml = sheetXml.replace('</conditionalFormatting>', '</conditionalFormatting><conditionalFormatting sqref="B1:C1"><cfRule type="expression" dxfId="0" priority="3"><formula>B1&gt;10</formula></cfRule></conditionalFormatting>');
  const expressionModel = analyze(zipSync({ ...files, 'xl/worksheets/sheet1.xml': new TextEncoder().encode(expressionXml) }));
  assert.equal(expressionModel.writable, true, expressionModel.readOnlyReason);
  const expressionStyle = column => require('../src/xlsx').conditionalStyleForCell(expressionModel.sheets[0], 0, column, expressionModel.sheets[0].cells.get(`${String.fromCharCode(65 + column)}1`) || null, (row, col) => col === 1 ? 42 : 1);
  assert.deepEqual(expressionStyle(1), { fill: '#FF0000' }, 'relative reference is evaluated at the range anchor without treating blank values as zero');
  assert.deepEqual(expressionStyle(2), { bold: true, color: '#008000' }, 'relative reference shifts with the target cell and the matching numeric rule keeps its differential style');
  const unsupportedExpression = analyze(zipSync({ ...files, 'xl/worksheets/sheet1.xml': new TextEncoder().encode(sheetXml.replace('</conditionalFormatting>', '</conditionalFormatting><conditionalFormatting sqref="B1"><cfRule type="expression" dxfId="0" priority="3"><formula>UNSUPPORTED(B1)</formula></cfRule></conditionalFormatting>')) }));
  assert.equal(unsupportedExpression.writable, false);
  assert.match(unsupportedExpression.readOnlyReason, /조건부 서식/);
});

test('XLSX 행·열 구조 변경은 조건부 서식 범위와 수식 참조를 함께 갱신한다', () => {
  const files = unzipSync(fixture());
  const worksheet = new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</sheetData>', '</sheetData><conditionalFormatting sqref="B1:C2"><cfRule type="expression" dxfId="0" priority="1"><formula>B1&gt;10</formula></cfRule></conditionalFormatting>');
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(worksheet);
  files['xl/styles.xml'] = new TextEncoder().encode('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs><dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs></styleSheet>');
  files['[Content_Types].xml'] = new TextEncoder().encode(new TextDecoder().decode(files['[Content_Types].xml']).replace('</Types>', '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'));
  files['xl/_rels/workbook.xml.rels'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/_rels/workbook.xml.rels']).replace('</Relationships>', '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'));
  const original = analyze(zipSync(files));
  assert.equal(original.writable, true, original.readOnlyReason);
  const inserted = transformStructure(original, { sheetIndex: 0, axis: 'row', action: 'insert', index: 0 });
  assert.equal(inserted.model.writable, true, inserted.model.readOnlyReason);
  assert.deepEqual(inserted.model.sheets[0].conditionalFormats[0].ranges, [{ top: 1, bottom: 2, left: 1, right: 2 }]);
  assert.equal(inserted.model.sheets[0].conditionalFormats[0].formula, 'B2>10');
  const reopened = analyze(inserted.bytes);
  assert.equal(reopened.writable, true, reopened.readOnlyReason);
  assert.equal(reopened.sheets[0].conditionalFormats[0].formula, 'B2>10');
  const deletedReference = transformStructure(original, { sheetIndex: 0, axis: 'row', action: 'delete', index: 0 });
  assert.equal(deletedReference.model.writable, true, deletedReference.model.readOnlyReason);
  assert.deepEqual(deletedReference.model.sheets[0].conditionalFormats[0].ranges, [{ top: 0, bottom: 0, left: 1, right: 2 }]);
  assert.equal(deletedReference.model.sheets[0].conditionalFormats[0].formula, '#REF!>10');
  assert.equal(analyze(deletedReference.bytes).writable, true, 'a deleted conditional-formula reference remains a valid saved workbook');
  const columnInserted = transformStructure(original, { sheetIndex: 0, axis: 'column', action: 'insert', index: 0 });
  assert.equal(columnInserted.model.writable, true, columnInserted.model.readOnlyReason);
  assert.deepEqual(columnInserted.model.sheets[0].conditionalFormats[0].ranges, [{ top: 0, bottom: 1, left: 2, right: 3 }]);
  assert.equal(columnInserted.model.sheets[0].conditionalFormats[0].formula, 'C1>10');
});

test('XLSX 상위·하위 조건부 서식은 순위·백분율·동점을 계산하고 저장 후 유지한다', () => {
  const files = unzipSync(fixture({ formatted: true }));
  files['xl/styles.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/styles.xml']).replace('</styleSheet>', '<dxfs count="2"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></dxf><dxf><font><b/></font></dxf></dxfs></styleSheet>'));
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A5"/><sheetData><row r="1"><c r="A1"><v>10</v></c></row><row r="2"><c r="A2"><v>20</v></c></row><row r="3"><c r="A3"><v>20</v></c></row><row r="4"><c r="A4"><f>A1+20</f><v>1</v></c></row><row r="5"><c r="A5" t="inlineStr"><is><t>text</t></is></c></row></sheetData><conditionalFormatting sqref="A1:A5"><cfRule type="top10" dxfId="0" priority="1" rank="2"/><cfRule type="top10" dxfId="1" priority="2" rank="25" percent="1" bottom="1"/></conditionalFormatting></worksheet>');
  const original = analyze(zipSync(files));
  assert.equal(original.writable, true, original.readOnlyReason);
  assert.deepEqual(original.sheets[0].conditionalFormats.map(rule => rule.top10), [{ rank: 2, percent: false, bottom: false }, { rank: 25, percent: true, bottom: true }]);
  const values = [10, 20, 20, 30, null];
  const valueAt = row => values[row];
  const styleAt = row => {
    const cell = original.sheets[0].cells.get('A' + (row + 1)) || null;
    return require('../src/xlsx').conditionalStyleForCell(original.sheets[0], row, 0, row === 3 ? { ...cell, kind: 'number', value: '30' } : cell, valueAt);
  };
  assert.deepEqual([0, 1, 2, 3, 4].map(row => styleAt(row)), [
    { bold: true }, { fill: '#FF0000' }, { fill: '#FF0000' }, { fill: '#FF0000' }, null
  ]);
  const saved = patchCell(original, 0, 'A1', 'changed');
  const reopened = analyze(saved.bytes);
  assert.equal(reopened.writable, true, reopened.readOnlyReason);
  assert.deepEqual(reopened.sheets[0].conditionalFormats.map(rule => rule.top10), original.sheets[0].conditionalFormats.map(rule => rule.top10));
  assert.match(reopened.sheets[0].xml, /type="top10" dxfId="0" priority="1" rank="2"/);
  assert.match(reopened.sheets[0].xml, /type="top10" dxfId="1" priority="2" rank="25" percent="1" bottom="1"/);
  const edited = patchEditConditionalFormatRule(reopened, 0, 1, { kind: 'condition', top10: { rank: 50, percent: true, bottom: true } });
  assert.equal(edited.model.writable, true, edited.model.readOnlyReason);
  assert.deepEqual(edited.model.sheets[0].conditionalFormats[0].top10, { rank: 50, percent: true, bottom: true });
  assert.match(edited.model.sheets[0].xml, /type="top10"[^>]*rank="50"[^>]*percent="1"[^>]*bottom="1"/);
  assert.equal(analyze(edited.bytes).sheets[0].conditionalFormats[0].top10.rank, 50);
  assert.equal(patchEditConditionalFormatRule(edited.model, 0, 1, { kind: 'condition', top10: { rank: 50, percent: true, bottom: true } }).model, edited.model);
  for (const top10 of [{ rank: 0, percent: false, bottom: false }, { rank: 1001, percent: false, bottom: false }, { rank: 101, percent: true, bottom: false }, { rank: 2, percent: 1, bottom: false }, { rank: 2, percent: false, bottom: false, extra: true }]) assert.throws(() => patchEditConditionalFormatRule(reopened, 0, 1, { kind: 'condition', top10 }), /순위 설정이 유효하지 않습니다/);
  const malformed = new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('rank="2"', 'rank="0"');
  assert.equal(analyze(zipSync({ ...files, 'xl/worksheets/sheet1.xml': new TextEncoder().encode(malformed) })).writable, false);
  const misplaced = new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('<cfRule type="top10" dxfId="0" priority="1" rank="2"/>', '<cfRule type="expression" dxfId="0" priority="1" rank="2"><formula>A1&gt;0</formula></cfRule>');
  assert.equal(analyze(zipSync({ ...files, 'xl/worksheets/sheet1.xml': new TextEncoder().encode(misplaced) })).writable, false);
});

test('XLSX 중복 값 조건부 서식은 대소문자와 빈 셀을 처리하고 저장·재열기에서 유지한다', () => {
  const files = unzipSync(fixture()), decode = bytes => new TextDecoder().decode(bytes), encode = text => new TextEncoder().encode(text);
  let worksheet = decode(files['xl/worksheets/sheet1.xml']).replace('<t>before</t>', '<t>Apple</t>').replace('<c r="B1"><v>42</v></c>', '<c r="B1" t="inlineStr"><is><t>apple</t></is></c>').replace('<c r="C1" t="b"><v>1</v></c>', '<c r="C1" t="inlineStr"><is><t>APPLE</t></is></c>');
  files['xl/worksheets/sheet1.xml'] = encode(worksheet);
  const original = analyze(zipSync(files));
  assert.equal(original.writable, true, original.readOnlyReason);
  const result = patchDuplicateValuesConditionalFormat(original, 0, { top: 0, bottom: 0, left: 0, right: 3 });
  assert.equal(result.model.writable, true, result.model.readOnlyReason);
  const rule = result.model.sheets[0].conditionalFormats[0];
  assert.equal(rule.type, 'expression');
  assert.equal(rule.formula, 'AND(A1<>"",COUNTIF($A$1:$D$1,A1)>1)');
  assert.deepEqual(rule.ranges, [{ top: 0, bottom: 0, left: 0, right: 3 }]);
  assert.deepEqual(rule.style, { fill: '#F4CCCC' });
  const getCellValue = (row, column) => result.model.sheets[0].cells.get(String.fromCharCode(65 + column) + (row + 1))?.value ?? null;
  const styleAt = column => require('../src/xlsx').conditionalStyleForCell(result.model.sheets[0], 0, column, result.model.sheets[0].cells.get(String.fromCharCode(65 + column) + '1') || null, getCellValue, evaluateFormula);
  assert.deepEqual(styleAt(0), { fill: '#F4CCCC' });
  assert.deepEqual(styleAt(1), { fill: '#F4CCCC' });
  assert.deepEqual(styleAt(2), { fill: '#F4CCCC' });
  assert.equal(styleAt(3), null, 'blank cells are not highlighted as duplicates');
  const edited = patchCell(result.model, 0, 'A1', 'Pear');
  const reopened = analyze(edited.bytes);
  assert.equal(reopened.writable, true, reopened.readOnlyReason);
  assert.equal(reopened.sheets[0].conditionalFormats[0].formula, rule.formula);
  assert.deepEqual(require('../src/xlsx').conditionalStyleForCell(reopened.sheets[0], 0, 0, reopened.sheets[0].cells.get('A1'), (row, column) => reopened.sheets[0].cells.get(String.fromCharCode(65 + column) + (row + 1))?.value ?? null, evaluateFormula), null, 'changing a value refreshes the duplicate check');
  assert.throws(() => patchDuplicateValuesConditionalFormat(original, 0, { top: 0, bottom: 100000, left: 0, right: 0 }), /100,000셀/);
  assert.throws(() => patchDuplicateValuesConditionalFormat(original, 0, { top: 0, bottom: 0, left: 0, right: 0 }, 'red'), /색상/);
});

test('XLSX 조건부 서식 차등 글꼴·채우기 색상은 공유 스타일을 복제해 대상 규칙에만 적용한다', () => {
  const files = unzipSync(fixture({ formatted: true }));
  const stylesPath = 'xl/styles.xml';
  const originalDxf = '<dxf><font><b/><color rgb="FF008000"/></font><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></dxf>';
  const otherDxf = '<dxf><font><i/></font></dxf>';
  files[stylesPath] = new TextEncoder().encode(new TextDecoder().decode(files[stylesPath]).replace('</styleSheet>', '<dxfs count="2">' + originalDxf + otherDxf + '</dxfs></styleSheet>'));
  const rules = '<conditionalFormatting sqref="A1:C1"><cfRule type="cellIs" operator="greaterThan" dxfId="0" priority="1"><formula>10</formula></cfRule><cfRule type="expression" dxfId="0" priority="2"><formula>A1&gt;10</formula></cfRule><cfRule type="cellIs" operator="lessThan" dxfId="1" priority="3"><formula>0</formula></cfRule></conditionalFormatting>';
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</sheetData>', '</sheetData>' + rules));
  const original = analyze(zipSync(files));
  assert.equal(original.writable, true, original.readOnlyReason);
  assert.equal(original.sheets[0].conditionalFormats[0].dxfId, 0);
  assert.deepEqual(original.sheets[0].conditionalFormats[0].style, { bold: true, color: '#008000', fill: '#FF0000' });

  const fontChanged = patchEditConditionalFormatRule(original, 0, 1, { kind: 'differential-color', property: 'fontColor', color: '#123456' });
  assert.equal(fontChanged.model.sheets[0].conditionalFormats[0].dxfId, 2);
  assert.equal(fontChanged.model.sheets[0].conditionalFormats[1].dxfId, 0, 'the other rule keeps using the shared original differential style');
  assert.equal(fontChanged.model.sheets[0].conditionalFormats[2].dxfId, 1, 'unrelated differential styles remain in place');
  assert.deepEqual(fontChanged.model.sheets[0].conditionalFormats[0].style, { bold: true, color: '#123456', fill: '#FF0000' });

  const fillChanged = patchEditConditionalFormatRule(fontChanged.model, 0, 1, { kind: 'differential-color', property: 'fillColor', color: '#654321' });
  assert.equal(fillChanged.model.sheets[0].conditionalFormats[0].dxfId, 3);
  assert.deepEqual(fillChanged.model.sheets[0].conditionalFormats[0].style, { bold: true, color: '#123456', fill: '#654321' });
  const colorRemoved = patchEditConditionalFormatRule(fillChanged.model, 0, 1, { kind: 'differential-color', property: 'fontColor', color: null });
  assert.deepEqual(colorRemoved.model.sheets[0].conditionalFormats[0].style, { bold: true, fill: '#654321' });
  assert.deepEqual(colorRemoved.model.sheets[0].conditionalFormats[1].style, { bold: true, color: '#008000', fill: '#FF0000' });
  assert.deepEqual(colorRemoved.model.sheets[0].conditionalFormats[2].style, { italic: true });
  const reopened = analyze(colorRemoved.bytes);
  assert.equal(reopened.writable, true, reopened.readOnlyReason);
  assert.deepEqual(reopened.sheets[0].conditionalFormats[0].style, { bold: true, fill: '#654321' });
  assert.equal(reopened.sheets[0].conditionalFormats[1].dxfId, 0);
  assert.equal(reopened.sheets[0].conditionalFormats[2].dxfId, 1);
  const savedStyles = new TextDecoder().decode(reopened.files[stylesPath]);
  assert.match(savedStyles, /<dxfs count="5">/);
  assert.ok(savedStyles.includes(originalDxf), 'the shared original differential style is preserved');
  assert.throws(() => patchEditConditionalFormatRule(reopened, 0, 1, { kind: 'differential-color', property: 'fillColor', color: '#12345' }), /조건부 서식 색상 값/);
});

test('XLSX 조건부 서식 지우기는 선택 범위만 제거하고 나머지 참조와 우선순위를 보존한다', () => {
  const files = unzipSync(fixture({ formatted: true }));
  const dataBarRule = '<cfRule type="dataBar" priority="7"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF4472C4"/></dataBar></cfRule>';
  const scaleRule = '<cfRule type="colorScale" priority="2"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FF00FF00"/></colorScale></cfRule>';
  const xml = new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</sheetData>', '</sheetData><conditionalFormatting sqref="B1:D3 F1:F2">' + dataBarRule + '</conditionalFormatting><conditionalFormatting sqref="A1">' + scaleRule + '</conditionalFormatting>');
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(xml);
  const original = analyze(zipSync(files));
  assert.equal(original.writable, true, original.readOnlyReason);
  const changed = patchClearConditionalFormats(original, 0, { top: 1, bottom: 1, left: 2, right: 2 });
  assert.equal(original.sheets[0].xml, xml, 'the source model remains unchanged');
  assert.equal(changed.model.writable, true, changed.model.readOnlyReason);
  const rules = changed.model.sheets[0].conditionalFormats;
  assert.deepEqual(rules.map(rule => rule.priority), [2, 7]);
  assert.deepEqual(rules.find(rule => rule.priority === 7).ranges, [
    { top: 0, bottom: 0, left: 1, right: 3 },
    { top: 2, bottom: 2, left: 1, right: 3 },
    { top: 1, bottom: 1, left: 1, right: 1 },
    { top: 1, bottom: 1, left: 3, right: 3 },
    { top: 0, bottom: 1, left: 5, right: 5 }
  ]);
  assert.ok(changed.model.sheets[0].xml.includes(dataBarRule), 'the rule body and priority are preserved');
  assert.ok(changed.model.sheets[0].xml.includes('<conditionalFormatting sqref="B1:D1 B3:D3 B2 D2 F1:F2">'));

  const wideClear = patchClearConditionalFormats(changed.model, 0, { top: 0, bottom: 9, left: 0, right: 10 });
  assert.equal(wideClear.model.sheets[0].conditionalFormats.length, 0);
  assert.doesNotMatch(wideClear.model.sheets[0].xml, /<conditionalFormatting\b/);

  const manyReferences = ['B1:D3', ...Array.from({ length: 63 }, (_, index) => 'A' + (index + 10))];
  const manyRangesModel = analyze(zipSync({ ...files, 'xl/worksheets/sheet1.xml': new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace(/sqref="B1:D3 F1:F2"/, 'sqref="' + manyReferences.join(' ') + '"').replace(/<conditionalFormatting sqref="A1">[\s\S]*?<\/conditionalFormatting>/, '')) }));
  assert.equal(manyRangesModel.writable, true, manyRangesModel.readOnlyReason);
  assert.throws(() => patchClearConditionalFormats(manyRangesModel, 0, { top: 1, bottom: 1, left: 2, right: 2 }), /64개를 초과/);
});

test('XLSX 조건부 서식 규칙 관리는 한 규칙만 삭제하고 형제 규칙을 보존한다', () => {
  const files = unzipSync(fixture({ formatted: true }));
  const dataBar = '<cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF4472C4"/></dataBar></cfRule>';
  const colorScale = '<cfRule type="colorScale" priority="2"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FF00FF00"/></colorScale></cfRule>';
  const block = '<conditionalFormatting sqref="A1:C3">' + dataBar + colorScale + '</conditionalFormatting>';
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</sheetData>', '</sheetData>' + block));
  const original = analyze(zipSync(files));
  assert.equal(original.writable, true, original.readOnlyReason);
  const deleted = patchDeleteConditionalFormatRule(original, 0, 1);
  assert.equal(deleted.model.writable, true, deleted.model.readOnlyReason);
  assert.deepEqual(deleted.model.sheets[0].conditionalFormats.map(rule => rule.priority), [2]);
  assert.ok(deleted.model.sheets[0].xml.includes('<conditionalFormatting sqref="A1:C3">' + colorScale + '</conditionalFormatting>'));
  assert.equal(original.sheets[0].xml.includes(block), true, 'the source model remains unchanged');
  assert.throws(() => patchDeleteConditionalFormatRule(deleted.model, 0, 1), /찾을 수 없습니다/);
  const removedLast = patchDeleteConditionalFormatRule(deleted.model, 0, 2);
  assert.equal(removedLast.model.sheets[0].conditionalFormats.length, 0);
  assert.doesNotMatch(removedLast.model.sheets[0].xml, /<conditionalFormatting\b/);
});

test('XLSX 조건부 서식 규칙 우선순위 이동은 인접 규칙과 순서만 맞바꾼다', () => {
  const files = unzipSync(fixture({ formatted: true }));
  const dataBar = '<cfRule type="dataBar" priority="7"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF4472C4"/></dataBar></cfRule>';
  const colorScale = '<cfRule type="colorScale" priority="2"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FF00FF00"/></colorScale></cfRule>';
  const otherRule = '<cfRule type="dataBar" priority="10"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF70AD47"/></dataBar></cfRule>';
  const xml = new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</sheetData>', '</sheetData><conditionalFormatting sqref="A1">' + dataBar + '</conditionalFormatting><conditionalFormatting sqref="B1">' + colorScale + '</conditionalFormatting><conditionalFormatting sqref="C1">' + otherRule + '</conditionalFormatting>');
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(xml);
  const original = analyze(zipSync(files));
  assert.equal(original.writable, true, original.readOnlyReason);

  const movedUp = patchMoveConditionalFormatRule(original, 0, 7, 'up');
  assert.equal(movedUp.model.writable, true, movedUp.model.readOnlyReason);
  assert.match(movedUp.model.sheets[0].xml, /<cfRule type="dataBar" priority="2">/);
  assert.match(movedUp.model.sheets[0].xml, /<cfRule type="colorScale" priority="7">/);
  assert.match(movedUp.model.sheets[0].xml, /<cfRule type="dataBar" priority="10">/);
  assert.ok(movedUp.model.sheets[0].xml.includes('<conditionalFormatting sqref="A1"><cfRule type="dataBar" priority="2">'));
  assert.ok(movedUp.model.sheets[0].xml.includes('<conditionalFormatting sqref="B1"><cfRule type="colorScale" priority="7">'));
  assert.equal(original.sheets[0].xml, xml, 'the source model remains unchanged');
  assert.strictEqual(patchMoveConditionalFormatRule(movedUp.model, 0, 2, 'up').model, movedUp.model, 'moving the highest-priority rule up is a no-op');

  const movedDown = patchMoveConditionalFormatRule(original, 0, 2, 'down');
  assert.match(movedDown.model.sheets[0].xml, /<cfRule type="dataBar" priority="2">/);
  assert.match(movedDown.model.sheets[0].xml, /<cfRule type="colorScale" priority="7">/);
  assert.throws(() => patchMoveConditionalFormatRule(original, 0, 7, 'left'), /선택이 유효하지 않습니다/);
  assert.throws(() => patchMoveConditionalFormatRule(original, 0, 9, 'up'), /찾을 수 없습니다/);
});

test('XLSX 셀 비교·수식 조건을 편집하면서 범위와 차등 서식을 보존한다', () => {
  const files = unzipSync(fixture({ formatted: true }));
  files['xl/styles.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/styles.xml']).replace('</styleSheet>', '<dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs></styleSheet>'));
  const cellRule = '<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>10</formula></cfRule>';
  const expressionRule = '<cfRule type="expression" dxfId="0" priority="2"><formula>B1&gt;10</formula></cfRule>';
  const xml = new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</sheetData>', '</sheetData><conditionalFormatting sqref="A1:B2">' + cellRule + '</conditionalFormatting><conditionalFormatting sqref="C1:C2">' + expressionRule + '</conditionalFormatting>');
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(xml);
  const original = analyze(zipSync(files));
  assert.equal(original.writable, true, original.readOnlyReason);

  const cellChanged = patchEditConditionalFormatRule(original, 0, 1, { kind: 'condition', operator: 'between', values: ['12.5', '20'] });
  assert.equal(cellChanged.model.writable, true, cellChanged.model.readOnlyReason);
  assert.deepEqual(cellChanged.model.sheets[0].conditionalFormats[0].values, [12.5, 20]);
  assert.match(cellChanged.model.sheets[0].xml, /operator="between"><formula>12\.5<\/formula><formula>20<\/formula>/);
  assert.match(cellChanged.model.sheets[0].xml, /<conditionalFormatting sqref="A1:B2">/);
  assert.match(cellChanged.model.sheets[0].xml, /dxfId="0"/);
  const singleValue = patchEditConditionalFormatRule(cellChanged.model, 0, 1, { kind: 'condition', operator: 'lessThan', values: ['7'] });
  assert.deepEqual(singleValue.model.sheets[0].conditionalFormats[0].values, [7]);
  assert.match(singleValue.model.sheets[0].xml, /operator="lessThan"><formula>7<\/formula><\/cfRule>/);

  const formulaChanged = patchEditConditionalFormatRule(cellChanged.model, 0, 2, { kind: 'condition', formula: '=B1>15' });
  assert.equal(formulaChanged.model.writable, true, formulaChanged.model.readOnlyReason);
  assert.equal(formulaChanged.model.sheets[0].conditionalFormats.find(rule => rule.type === 'expression').formula, 'B1>15');
  assert.match(formulaChanged.model.sheets[0].xml, /<formula>B1&gt;15<\/formula>/);
  assert.throws(() => patchEditConditionalFormatRule(original, 0, 1, { kind: 'condition', operator: 'between', values: ['12', 'bad'] }), /비교값이 유효하지 않습니다/);
  assert.throws(() => patchEditConditionalFormatRule(original, 0, 2, { kind: 'condition', formula: 'UNSUPPORTED(A1)' }), /수식이 유효하지 않거나 지원되지 않습니다/);
  assert.equal(original.sheets[0].xml, xml, 'the source model remains unchanged');
});

test('XLSX 데이터 막대·색조 규칙 색 편집은 기존 기준과 적용 범위를 보존한다', () => {
  const files = unzipSync(fixture({ formatted: true }));
  const dataBar = '<cfRule type="dataBar" priority="1"><dataBar minLength="5" maxLength="90"><cfvo type="min"/><cfvo type="max"/><color rgb="FF4472C4"/></dataBar></cfRule>';
  const colorScale = '<cfRule type="colorScale" priority="2"><colorScale><cfvo type="min"/><cfvo type="num" val="50"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FFFFFF00"/><color rgb="FF00B050"/></colorScale></cfRule>';
  const xml = new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</sheetData>', '</sheetData><conditionalFormatting sqref="A1:A3">' + dataBar + '</conditionalFormatting><conditionalFormatting sqref="B1:B3">' + colorScale + '</conditionalFormatting>');
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(xml);
  const original = analyze(zipSync(files));
  assert.equal(original.writable, true, original.readOnlyReason);
  const barChanged = patchEditConditionalFormatRule(original, 0, 1, { kind: 'colors', colors: ['#112233'] });
  assert.equal(barChanged.model.sheets[0].conditionalFormats[0].dataBar.color, '#112233');
  assert.match(barChanged.model.sheets[0].xml, /dataBar minLength="5" maxLength="90">[\s\S]*<color rgb="FF112233"\/>/);
  const scaleChanged = patchEditConditionalFormatRule(barChanged.model, 0, 2, { kind: 'colors', colors: ['#010203', '#A0B0C0', '#FEDCBA'] });
  const scale = scaleChanged.model.sheets[0].conditionalFormats.find(rule => rule.type === 'colorScale');
  assert.deepEqual(scale.colorScale.colors, ['#010203', '#A0B0C0', '#FEDCBA']);
  assert.deepEqual(scale.colorScale.stops, [{ type: 'min', value: null }, { type: 'num', value: 50 }, { type: 'max', value: null }]);
  assert.match(scaleChanged.model.sheets[0].xml, /<conditionalFormatting sqref="B1:B3"><cfRule type="colorScale" priority="2">/);
  const thresholdChanged = patchEditConditionalFormatRule(scaleChanged.model, 0, 2, { kind: 'threshold', stop: { type: 'percentile', value: '75.5' } });
  const thresholdRule = thresholdChanged.model.sheets[0].conditionalFormats.find(rule => rule.type === 'colorScale');
  assert.deepEqual(thresholdRule.colorScale.stops, [{ type: 'min', value: null }, { type: 'percentile', value: 75.5 }, { type: 'max', value: null }]);
  assert.deepEqual(thresholdRule.colorScale.colors, ['#010203', '#A0B0C0', '#FEDCBA']);
  assert.match(thresholdChanged.model.sheets[0].xml, /<cfvo type="percentile" val="75\.5"\/>/);
  assert.deepEqual(thresholdChanged.model.sheets[0].conditionalFormats.find(rule => rule.type === 'dataBar').dataBar, { color: '#112233', minLength: 5, maxLength: 90, showValue: true, gradient: true });
  const scientificThreshold = patchEditConditionalFormatRule(thresholdChanged.model, 0, 2, { kind: 'threshold', stop: { type: 'num', value: '1.25e2' } });
  assert.deepEqual(scientificThreshold.model.sheets[0].conditionalFormats.find(rule => rule.type === 'colorScale').colorScale.stops[1], { type: 'num', value: 125 });
  assert.match(scientificThreshold.model.sheets[0].xml, /<cfvo type="num" val="1\.25e2"\/>/);
  assert.throws(() => patchEditConditionalFormatRule(original, 0, 1, { kind: 'colors', colors: ['red'] }), /색상 값이 유효하지 않습니다/);
  assert.throws(() => patchEditConditionalFormatRule(original, 0, 2, { kind: 'colors', colors: ['#112233'] }), /규칙을 편집할 수 없습니다/);
  assert.throws(() => patchEditConditionalFormatRule(original, 0, 2, { kind: 'threshold', stop: { type: 'percent', value: '101' } }), /기준값이 유효하지 않습니다/);
  assert.throws(() => patchEditConditionalFormatRule(original, 0, 2, { kind: 'threshold', stop: { type: 'formula', value: '=A1' } }), /규칙을 편집할 수 없습니다/);
  assert.equal(original.sheets[0].xml, xml, 'the source model remains unchanged');
});

test('XLSX 데이터 막대 표시 편집은 길이·값 표시·채우기를 검증하고 다른 규칙을 보존한다', () => {
  const files = unzipSync(fixture({ formatted: true }));
  const dataBar = '<cfRule type="dataBar" priority="1"><dataBar minLength="5" maxLength="90" showValue="0" gradient="0"><cfvo type="min"/><cfvo type="max"/><color rgb="FF4472C4"/></dataBar></cfRule>';
  const colorScale = '<cfRule type="colorScale" priority="2"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FF00FF00"/></colorScale></cfRule>';
  const xml = new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</sheetData>', '</sheetData><conditionalFormatting sqref="A1:A3">' + dataBar + '</conditionalFormatting><conditionalFormatting sqref="B1:B3">' + colorScale + '</conditionalFormatting>');
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(xml);
  const original = analyze(zipSync(files));
  assert.equal(original.writable, true, original.readOnlyReason);
  const unchanged = patchEditConditionalFormatRule(original, 0, 1, { kind: 'data-bar-display', dataBar: { minLength: 5, maxLength: 90, showValue: false, gradient: false } });
  assert.strictEqual(unchanged.model, original, 'unchanged settings are a no-op');
  const changed = patchEditConditionalFormatRule(original, 0, 1, { kind: 'data-bar-display', dataBar: { minLength: 25, maxLength: 80, showValue: true, gradient: true } });
  const rule = changed.model.sheets[0].conditionalFormats.find(item => item.type === 'dataBar');
  assert.deepEqual(rule.dataBar, { color: '#4472C4', minLength: 25, maxLength: 80, showValue: true, gradient: true });
  assert.match(changed.model.sheets[0].xml, /<dataBar minLength="25" maxLength="80" showValue="1" gradient="1"><cfvo type="min"\/><cfvo type="max"\/><color rgb="FF4472C4"\/><\/dataBar>/);
  assert.deepEqual(changed.model.sheets[0].conditionalFormats.find(item => item.type === 'colorScale').colorScale.colors, ['#FF0000', '#00FF00']);
  assert.ok(changed.model.sheets[0].xml.includes('<conditionalFormatting sqref="A1:A3">'));
  assert.equal(original.sheets[0].xml, xml, 'the source model remains unchanged');
  for (const display of [
    { minLength: -1, maxLength: 80, showValue: true, gradient: true },
    { minLength: 25, maxLength: 101, showValue: true, gradient: true },
    { minLength: 81, maxLength: 80, showValue: true, gradient: true },
    { minLength: 25.5, maxLength: 80, showValue: true, gradient: true },
    { minLength: 25, maxLength: 80, showValue: '1', gradient: true },
    { minLength: 25, maxLength: 80, showValue: true, gradient: true, color: '#FFFFFF' }
  ]) assert.throws(() => patchEditConditionalFormatRule(original, 0, 1, { kind: 'data-bar-display', dataBar: display }), /데이터 막대 표시 설정이 유효하지 않습니다/);
  assert.throws(() => patchEditConditionalFormatRule(original, 0, 2, { kind: 'data-bar-display', dataBar: { minLength: 25, maxLength: 80, showValue: true, gradient: true } }), /데이터 막대 표시 설정이 유효하지 않습니다/);
});

test('XLSX 최소·최대 데이터 막대는 값 범위에 맞춰 표시하고 규칙 XML을 보존한다', () => {
  const files = unzipSync(fixture({ formatted: true }));
  const worksheet = new TextDecoder().decode(files['xl/worksheets/sheet1.xml'])
    .replace('<v>45292</v></c><c r="C1"', '<v>1</v></c><c r="C1"')
    .replace('<v>0.125</v></c><c r="D1"', '<v>2</v></c><c r="D1"')
    .replace('<v>45292</v></c></row>', '<v>3</v></c></row>')
    .replace('</sheetData>', '</sheetData><conditionalFormatting sqref="B1:D1"><cfRule type="dataBar" priority="1"><dataBar minLength="0" maxLength="100"><cfvo type="min" val="0"/><cfvo type="max" val="0"/><color rgb="FF638EC6"/></dataBar></cfRule></conditionalFormatting>');
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(worksheet);
  const model = analyze(zipSync(files));
  assert.equal(model.writable, true, model.readOnlyReason);
  const styleAt = column => require('../src/xlsx').conditionalStyleForCell(model.sheets[0], 0, column, model.sheets[0].cells.get(`${String.fromCharCode(65 + column)}1`), (row, targetColumn) => Number(model.sheets[0].cells.get(`${String.fromCharCode(65 + targetColumn)}1`)?.value));
  assert.deepEqual(styleAt(1), { dataBar: { color: '#638EC6', width: 0, showValue: true, gradient: true } });
  assert.deepEqual(styleAt(2), { dataBar: { color: '#638EC6', width: 50, showValue: true, gradient: true } });
  assert.deepEqual(styleAt(3), { dataBar: { color: '#638EC6', width: 100, showValue: true, gradient: true } });
  const saved = patchCell(model, 0, 'A1', 'after');
  const reopened = analyze(saved.bytes);
  assert.equal(reopened.writable, true, reopened.readOnlyReason);
  assert.ok(reopened.sheets[0].xml.includes('<cfRule type="dataBar" priority="1">'));
});

test('음수 데이터 막대는 혼합 범위에서 0 축을 기준으로 양쪽에 표시한다', () => {
  const files = unzipSync(fixture({ formatted: true }));
  const worksheet = new TextDecoder().decode(files['xl/worksheets/sheet1.xml'])
    .replace('<v>45292</v></c><c r="C1"', '<v>-10</v></c><c r="C1"')
    .replace('<v>0.125</v></c><c r="D1"', '<v>0</v></c><c r="D1"')
    .replace('<v>45292</v></c></row>', '<v>10</v></c></row>')
    .replace('</sheetData>', '</sheetData><conditionalFormatting sqref="B1:D1"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule></conditionalFormatting>');
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(worksheet);
  const model = analyze(zipSync(files));
  assert.equal(model.writable, true, model.readOnlyReason);
  const styleAt = column => require('../src/xlsx').conditionalStyleForCell(model.sheets[0], 0, column, model.sheets[0].cells.get(`${String.fromCharCode(65 + column)}1`), (row, targetColumn) => Number(model.sheets[0].cells.get(`${String.fromCharCode(65 + targetColumn)}1`)?.value));
  assert.deepEqual(styleAt(1).dataBar, { color: '#638EC6', width: 50, start: 0, end: 50, axis: 50, negative: true, showValue: true, gradient: true });
  assert.deepEqual(styleAt(2).dataBar, { color: '#638EC6', width: 0, start: 50, end: 50, axis: 50, negative: false, showValue: true, gradient: true });
  assert.deepEqual(styleAt(3).dataBar, { color: '#638EC6', width: 50, start: 50, end: 100, axis: 50, negative: false, showValue: true, gradient: true });

  const negativeFiles = { ...files };
  const negativeWorksheet = new TextDecoder().decode(negativeFiles['xl/worksheets/sheet1.xml'])
    .replace('<v>0</v>', '<v>-5</v>').replace('<v>10</v>', '<v>-1</v>');
  negativeFiles['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(negativeWorksheet);
  const allNegative = analyze(zipSync(negativeFiles));
  const negativeStyle = column => require('../src/xlsx').conditionalStyleForCell(allNegative.sheets[0], 0, column, allNegative.sheets[0].cells.get(`${String.fromCharCode(65 + column)}1`), (row, targetColumn) => Number(allNegative.sheets[0].cells.get(`${String.fromCharCode(65 + targetColumn)}1`)?.value));
  assert.deepEqual(negativeStyle(1).dataBar, { color: '#638EC6', width: 100, start: 0, end: 100, axis: 100, negative: true, showValue: true, gradient: true });
  assert.deepEqual(negativeStyle(2).dataBar, { color: '#638EC6', width: 50, start: 50, end: 100, axis: 100, negative: true, showValue: true, gradient: true });
  assert.deepEqual(negativeStyle(3).dataBar, { color: '#638EC6', width: 10, start: 90, end: 100, axis: 100, negative: true, showValue: true, gradient: true });
});

test('XLSX 3색 색조는 최소·백분위·최대 기준으로 셀 색을 보간하고 규칙 XML을 보존한다', () => {
  const files = unzipSync(fixture({ formatted: true }));
  const worksheet = new TextDecoder().decode(files['xl/worksheets/sheet1.xml'])
    .replace('<v>45292</v></c><c r="C1"', '<v>1</v></c><c r="C1"')
    .replace('<v>0.125</v></c><c r="D1"', '<v>2</v></c><c r="D1"')
    .replace('<v>45292</v></c></row>', '<v>3</v></c></row>')
    .replace('</sheetData>', '</sheetData><conditionalFormatting sqref="B1:D1"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FFFFFF00"/><color rgb="FF00FF00"/></colorScale></cfRule></conditionalFormatting>');
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(worksheet);
  const model = analyze(zipSync(files));
  assert.equal(model.writable, true, model.readOnlyReason);
  const styleAt = column => require('../src/xlsx').conditionalStyleForCell(model.sheets[0], 0, column, model.sheets[0].cells.get(`${String.fromCharCode(65 + column)}1`), (row, targetColumn) => Number(model.sheets[0].cells.get(`${String.fromCharCode(65 + targetColumn)}1`)?.value));
  assert.deepEqual([styleAt(1).fill, styleAt(2).fill, styleAt(3).fill], ['#FF0000', '#FFFF00', '#00FF00']);
  const saved = patchCell(model, 0, 'A1', 'after'), reopened = analyze(saved.bytes);
  assert.equal(reopened.writable, true, reopened.readOnlyReason);
  assert.ok(reopened.sheets[0].xml.includes('<cfRule type="colorScale" priority="1">'));
});

test('XLSX 아이콘 집합을 표시하고 아이콘 필터를 저장·재개방한다', () => {
  const files = unzipSync(fixture());
  const xml = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A4"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Score</t></is></c></row><row r="2"><c r="A2"><v>1</v></c></row><row r="3"><c r="A3"><v>5</v></c></row><row r="4"><c r="A4"><v>9</v></c></row></sheetData><conditionalFormatting sqref="A2:A4"><cfRule type="iconSet" priority="1"><iconSet iconSet="3TrafficLights1"><cfvo type="min"/><cfvo type="formula" val="A2*5"/><cfvo type="max"/></iconSet></cfRule></conditionalFormatting></worksheet>';
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(xml);
  const original = analyze(zipSync(files));
  assert.equal(original.writable, true, original.readOnlyReason);
  const styleAt = row => require('../src/xlsx').conditionalStyleForCell(original.sheets[0], row, 0, original.sheets[0].cells.get('A' + (row + 1)), (targetRow, column) => Number(original.sheets[0].cells.get('A' + (targetRow + 1))?.value), evaluateFormula);
  assert.deepEqual([styleAt(1).icon.id, styleAt(2).icon.id, styleAt(3).icon.id], [2, 1, 0]);
  const filter = { top: 0, bottom: 3, left: 0, right: 0, columns: [{ column: 0, iconFilter: { iconSet: '3TrafficLights1', iconId: 0 } }] };
  const saved = patchAutoFilters(original, [{ sheetIndex: 0, filter }]);
  assert.deepEqual(saved.model.sheets[0].nativeHiddenRows, [1, 2]);
  assert.match(saved.model.sheets[0].xml, /<iconFilter iconSet="3TrafficLights1" iconId="0"\/>/);
  const reopened = analyze(saved.bytes);
  assert.equal(reopened.writable, true, reopened.readOnlyReason);
  assert.deepEqual(reopened.sheets[0].autoFilter.columns, filter.columns);
  assert.deepEqual(reopened.sheets[0].nativeHiddenRows, [1, 2]);
  const filterOnly = analyze(fixture({ autoFilter: '<autoFilter ref="A1:C2"><filterColumn colId="1"><iconFilter iconSet="3Arrows" iconId="1"/></filterColumn></autoFilter>' }));
  const shifted = transformStructure(filterOnly, { sheetIndex: 0, axis: 'column', action: 'insert', index: 0 }).model;
  assert.deepEqual(shifted.sheets[0].autoFilter.columns, [{ column: 2, iconFilter: { iconSet: '3Arrows', iconId: 1 } }]);
});

test('XLSX 사용자 지정 아이콘 매핑을 표시·필터링하고 편집 시 보존한다', () => {
  const files=unzipSync(fixture());
  const xml='<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A4"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Score</t></is></c></row><row r="2"><c r="A2"><v>1</v></c></row><row r="3"><c r="A3"><v>5</v></c></row><row r="4"><c r="A4"><v>9</v></c></row></sheetData><conditionalFormatting sqref="A2:A4"><cfRule type="iconSet" priority="1"><iconSet iconSet="3TrafficLights1" custom="1"><cfvo type="min"/><cfvo type="num" val="4"/><cfvo type="max"/><cfIcon iconSet="3Arrows" iconId="2"/><cfIcon iconSet="3TrafficLights1" iconId="0"/><cfIcon iconSet="3Symbols2" iconId="2"/></iconSet></cfRule></conditionalFormatting></worksheet>';
  files['xl/worksheets/sheet1.xml']=new TextEncoder().encode(xml);
  const original=analyze(zipSync(files));assert.equal(original.writable,true,original.readOnlyReason);
  const rule=original.sheets[0].conditionalFormats[0];
  assert.deepEqual(rule.iconSet.customIcons,[{iconSet:'3Arrows',iconId:2},{iconSet:'3TrafficLights1',iconId:0},{iconSet:'3Symbols2',iconId:2}]);
  const styleAt=row=>require('../src/xlsx').conditionalStyleForCell(original.sheets[0],row,0,original.sheets[0].cells.get('A'+(row+1)),(targetRow,column)=>Number(original.sheets[0].cells.get('A'+(targetRow+1))?.value),evaluateFormula);
  assert.deepEqual([styleAt(1).icon,styleAt(2).icon,styleAt(3).icon].map(({set,id})=>({set,id})),[{set:'3Symbols2',id:2},{set:'3TrafficLights1',id:0},{set:'3Arrows',id:2}]);
  const originalThreshold=patchEditConditionalFormatRule(original,0,1,{kind:'icon-set-threshold',index:1,stop:{type:'num',value:6,gte:true}});
  assert.deepEqual(originalThreshold.model.sheets[0].conditionalFormats[0].iconSet.customIcons,rule.iconSet.customIcons);
  assert.deepEqual(originalThreshold.model.sheets[0].conditionalFormats[0].iconSet.thresholds[1],{type:'num',value:6,gte:true});
  const display=patchEditConditionalFormatRule(original,0,1,{kind:'icon-set-display',iconSet:{showValue:false,reverse:false}});
  assert.deepEqual(display.model.sheets[0].conditionalFormats[0].iconSet.customIcons,rule.iconSet.customIcons);
  assert.equal(display.model.sheets[0].conditionalFormats[0].iconSet.showValue,false);
  const remapped=patchEditConditionalFormatRule(original,0,1,{kind:'icon-set-mapping',customIcons:[{iconSet:'3ArrowsGray',iconId:1},{iconSet:'3TrafficLights2',iconId:2},{iconSet:'3Symbols',iconId:0}]});
  assert.deepEqual(remapped.model.sheets[0].conditionalFormats[0].iconSet.customIcons,[{iconSet:'3ArrowsGray',iconId:1},{iconSet:'3TrafficLights2',iconId:2},{iconSet:'3Symbols',iconId:0}]);
  assert.match(new TextDecoder().decode(remapped.model.files['xl/worksheets/sheet1.xml']),/<iconSet iconSet="3TrafficLights1" custom="1"><cfvo type="min"\/><cfvo type="num" val="4"\/><cfvo type="max"\/><cfIcon iconSet="3ArrowsGray" iconId="1"\/><cfIcon iconSet="3TrafficLights2" iconId="2"\/><cfIcon iconSet="3Symbols" iconId="0"\/><\/iconSet>/);
  assert.deepEqual(analyze(remapped.bytes).sheets[0].conditionalFormats[0].iconSet.customIcons,[{iconSet:'3ArrowsGray',iconId:1},{iconSet:'3TrafficLights2',iconId:2},{iconSet:'3Symbols',iconId:0}]);
  const remappedThreshold=patchEditConditionalFormatRule(remapped.model,0,1,{kind:'icon-set-threshold',index:1,stop:{type:'num',value:6,gte:false}});
  assert.deepEqual(remappedThreshold.model.sheets[0].conditionalFormats[0].iconSet.thresholds[1],{type:'num',value:6,gte:false});
  assert.deepEqual(remappedThreshold.model.sheets[0].conditionalFormats[0].iconSet.customIcons,remapped.model.sheets[0].conditionalFormats[0].iconSet.customIcons);
  assert.match(new TextDecoder().decode(remappedThreshold.model.files['xl/worksheets/sheet1.xml']),/<cfvo type="num" val="6" gte="0"\/><cfvo type="max"\/><cfIcon iconSet="3ArrowsGray" iconId="1"\/>/);
  assert.deepEqual(analyze(remappedThreshold.bytes).sheets[0].conditionalFormats[0].iconSet.customIcons,remapped.model.sheets[0].conditionalFormats[0].iconSet.customIcons);
  assert.throws(()=>patchEditConditionalFormatRule(original,0,1,{kind:'icon-set-mapping',customIcons:[{iconSet:'3Arrows',iconId:3},{iconSet:'3TrafficLights1',iconId:0},{iconSet:'3Symbols2',iconId:2}]}),/아이콘/);
  const filter=patchAutoFilters(original,[{sheetIndex:0,filter:{top:0,bottom:3,left:0,right:0,columns:[{column:0,iconFilter:{iconSet:'3TrafficLights1',iconId:0}}]}}]);
  assert.deepEqual(filter.model.sheets[0].nativeHiddenRows,[1,3]);
  const saved=patchCell(original,0,'A1','Changed');
  assert.match(new TextDecoder().decode(saved.model.files['xl/worksheets/sheet1.xml']),/<cfIcon iconSet="3Arrows" iconId="2"\/><cfIcon iconSet="3TrafficLights1" iconId="0"\/><cfIcon iconSet="3Symbols2" iconId="2"\/>/);
  const invalid=unzipSync(fixture());invalid['xl/worksheets/sheet1.xml']=new TextEncoder().encode(xml.replace('iconId="2"/><cfIcon iconSet="3TrafficLights1"','iconId="9"/><cfIcon iconSet="3TrafficLights1"'));
  assert.equal(analyze(zipSync(invalid)).writable,false);
});

test('XLSX 아이콘 집합 표시와 기준을 편집해 저장하고 미지원 패키지 항목을 보존한다', () => {
  const files = unzipSync(fixture());
  const xml = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A3"/><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>5</v></c></row><row r="3"><c r="A3"><v>9</v></c></row></sheetData><conditionalFormatting sqref="A1:A3"><cfRule type="iconSet" priority="1"><iconSet iconSet="3TrafficLights1" percent="1" showValue="0"><cfvo type="min"/><cfvo type="num" val="4" gte="0"/><cfvo type="max"/></iconSet></cfRule><cfRule type="colorScale" priority="2"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FF00FF00"/></colorScale></cfRule></conditionalFormatting></worksheet>';
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(xml);
  const original = analyze(zipSync(files));
  assert.equal(original.writable, true, original.readOnlyReason);
  const displayed = patchEditConditionalFormatRule(original, 0, 1, { kind: 'icon-set-display', iconSet: { showValue: true, reverse: true } });
  assert.deepEqual(displayed.model.sheets[0].conditionalFormats[0].iconSet, { name: '3TrafficLights1', showValue: true, reverse: true, thresholds: [{ type: 'min', value: null, gte: true }, { type: 'num', value: 4, gte: false }, { type: 'max', value: null, gte: true }] });
  assert.match(displayed.model.sheets[0].xml, /<iconSet iconSet="3TrafficLights1" percent="1" showValue="1" reverse="1">/);
  const threshold = patchEditConditionalFormatRule(displayed.model, 0, 1, { kind: 'icon-set-threshold', index: 1, stop: { type: 'formula', formula: 'A1*5', gte: true } });
  assert.deepEqual(threshold.model.sheets[0].conditionalFormats[0].iconSet.thresholds[1], { type: 'formula', formula: 'A1*5', value: null, gte: true });
  assert.match(threshold.model.sheets[0].xml, /<cfvo type="formula" val="A1\*5"\/>/);
  const reopened = analyze(threshold.bytes);
  assert.equal(reopened.writable, true, reopened.readOnlyReason);
  assert.deepEqual(reopened.sheets[0].conditionalFormats[0].iconSet, threshold.model.sheets[0].conditionalFormats[0].iconSet);
  assert.deepEqual(threshold.model.sheets[0].conditionalFormats[1], original.sheets[0].conditionalFormats[1]);
  for (const [name, before] of Object.entries(original.files)) if (name !== original.sheets[0].path) assert.deepEqual(threshold.model.files[name], before, name);
  assert.equal(patchEditConditionalFormatRule(threshold.model, 0, 1, { kind: 'icon-set-threshold', index: 1, stop: { type: 'formula', formula: 'A1*5', gte: true } }).model, threshold.model);
  const sameCountStyle = patchEditConditionalFormatRule(original, 0, 1, { kind: 'icon-set-style', name: '3TrafficLights2' });
  assert.equal(sameCountStyle.model.writable, true, sameCountStyle.model.readOnlyReason);
  assert.equal(sameCountStyle.model.sheets[0].conditionalFormats[0].iconSet.name, '3TrafficLights2');
  assert.deepEqual(sameCountStyle.model.sheets[0].conditionalFormats[0].iconSet.thresholds, original.sheets[0].conditionalFormats[0].iconSet.thresholds);
  assert.match(sameCountStyle.model.sheets[0].xml, /<iconSet iconSet="3TrafficLights2" percent="1" showValue="0">/);
  const differentCountStyle = patchEditConditionalFormatRule(sameCountStyle.model, 0, 1, { kind: 'icon-set-style', name: '5Quarters' });
  assert.equal(differentCountStyle.model.writable, true, differentCountStyle.model.readOnlyReason);
  assert.deepEqual(differentCountStyle.model.sheets[0].conditionalFormats[0].iconSet.thresholds, [0, 20, 40, 60, 80].map(value => ({ type: 'percent', value, gte: true })));
  assert.match(differentCountStyle.model.sheets[0].xml, /<iconSet iconSet="5Quarters" percent="1" showValue="0">/);
  const styleReopened = analyze(differentCountStyle.bytes);
  assert.equal(styleReopened.writable, true, styleReopened.readOnlyReason);
  assert.deepEqual(styleReopened.sheets[0].conditionalFormats[0].iconSet, differentCountStyle.model.sheets[0].conditionalFormats[0].iconSet);
  assert.throws(() => patchEditConditionalFormatRule(original, 0, 1, { kind: 'icon-set-style', name: 'customIcon' }), /아이콘/);
  for (const stop of [
    { type: 'percent', value: 101, gte: true },
    { type: 'percentile', value: -1, gte: false },
    { type: 'formula', formula: 'Other!A1', gte: true },
    { type: 'num', value: 'NaN', gte: true },
    { type: 'min', value: 0, gte: true }
  ]) assert.throws(() => patchEditConditionalFormatRule(original, 0, 1, { kind: 'icon-set-threshold', index: 1, stop }), /아이콘/);
  assert.throws(() => patchEditConditionalFormatRule(original, 0, 1, { kind: 'icon-set-threshold', index: 9, stop: { type: 'min', gte: true } }), /아이콘/);
});

test('XLSX 표준 아이콘 집합 17종을 선택 범위에 저장하고 기존 규칙 순서를 보존한다', () => {
  const files = unzipSync(fixture());
  const sheetXml = new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</worksheet>', '<conditionalFormatting sqref="A1"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FF00FF00"/></colorScale></cfRule></conditionalFormatting></worksheet>');
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(sheetXml);
  const original = analyze(zipSync(files));
  assert.equal(original.writable, true, original.readOnlyReason);
  assert.equal(Object.keys(ICON_SET_LENGTHS).length, 17);
  for (const [name, count] of Object.entries(ICON_SET_LENGTHS)) {
    const added = patchIconSetConditionalFormat(original, 0, { top: 0, bottom: 2, left: 0, right: 0 }, name);
    assert.equal(added.model.writable, true, name + ': ' + added.model.readOnlyReason);
    const rules = added.model.sheets[0].conditionalFormats;
    assert.equal(rules[0].type, 'iconSet', name);
    assert.equal(rules[0].priority, 1, name);
    assert.equal(rules[0].iconSet.name, name);
    assert.equal(rules[0].iconSet.thresholds.length, count);
    assert.deepEqual(rules[0].iconSet.thresholds.map(stop => stop.value), Array.from({ length: count }, (_item, index) => Math.round(index * 100 / count)));
    assert.equal(rules[1].priority, 2, name);
    assert.equal(rules[1].type, 'colorScale', name);
    const reopened = analyze(added.bytes);
    assert.equal(reopened.writable, true, name + ': ' + reopened.readOnlyReason);
    assert.equal(reopened.sheets[0].conditionalFormats[0].iconSet.name, name);
  }
  assert.throws(() => patchIconSetConditionalFormat(original, 0, { top: 0, bottom: 0, left: 0, right: 0 }, 'customIcon'), /아이콘/);
});

test('XLSX 인라인 목록 데이터 유효성은 셀 편집과 행 구조 변경에서 보존된다', () => {
  const validation = '<dataValidations count="1"><dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="A1:A3"><formula1>"Open,In Progress,Closed"</formula1></dataValidation></dataValidations>';
  const files = unzipSync(fixture());
  const worksheet = new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</worksheet>', validation + '</worksheet>');
  const model = analyze(zipSync({ ...files, 'xl/worksheets/sheet1.xml': new TextEncoder().encode(worksheet) }));
  assert.equal(model.writable, true, model.readOnlyReason);
  assert.deepEqual(model.sheets[0].dataValidations[0].values, ['Open', 'In Progress', 'Closed']);
  const saved = patchCell(model, 0, 'B1', 'unchanged validation');
  assert.ok(saved.model.sheets[0].xml.includes(validation));
  const inserted = transformStructure(saved.model, { sheetIndex: 0, axis: 'row', action: 'insert', index: 0 }).model;
  assert.equal(inserted.writable, true, inserted.readOnlyReason);
  assert.deepEqual(inserted.sheets[0].dataValidations[0].ranges, [{ top: 1, bottom: 3, left: 0, right: 0 }]);
  assert.match(inserted.sheets[0].xml, /sqref="A2:A4"><formula1>"Open,In Progress,Closed"/);
  const unsupportedWorksheet = worksheet.replace('"Open,In Progress,Closed"', '=$A$1:$A$3');
  const referenceWorksheet = worksheet.replace('"Open,In Progress,Closed"', '=$B$1:$B$3');
  const referenceModel = analyze(zipSync({ ...files, 'xl/worksheets/sheet1.xml': new TextEncoder().encode(referenceWorksheet) }));
  assert.equal(referenceModel.writable, true, referenceModel.readOnlyReason);
  assert.deepEqual(referenceModel.sheets[0].dataValidations[0].source, { top: 0, bottom: 2, left: 1, right: 1 });
  const movedReference = transformStructure(referenceModel, { sheetIndex: 0, axis: 'row', action: 'insert', index: 0 }).model;
  assert.deepEqual(movedReference.sheets[0].dataValidations[0].source, { top: 1, bottom: 3, left: 1, right: 1 });
  assert.match(movedReference.sheets[0].xml, /<formula1>=\$B\$2:\$B\$4<\/formula1>/);
  const unsupported = analyze(zipSync({ ...files, 'xl/worksheets/sheet1.xml': new TextEncoder().encode(unsupportedWorksheet.replace('=$A$1:$A$3', '=StatusList')) }));
  assert.equal(unsupported.writable, false);
  assert.match(unsupported.readOnlyReason, /데이터 유효성/);
  const promptWorksheet = worksheet.replace('type="list"', 'type="list" prompt="Pick a status" showInputMessage="1"');
  const promptModel = analyze(zipSync({ ...files, 'xl/worksheets/sheet1.xml': new TextEncoder().encode(promptWorksheet) }));
  assert.equal(promptModel.writable, true, promptModel.readOnlyReason);
  assert.equal(promptModel.sheets[0].dataValidations[0].prompt, 'Pick a status');
  assert.equal(promptModel.sheets[0].dataValidations[0].showInputMessage, true);
  const messages = worksheet.replace('type="list"', 'type="list" errorTitle="Invalid status" error="Choose an existing status" promptTitle="Status" prompt="Pick a status"');
  const messageModel = analyze(zipSync({ ...files, 'xl/worksheets/sheet1.xml': new TextEncoder().encode(messages) }));
  assert.equal(messageModel.writable, true, messageModel.readOnlyReason);
  const { errorTitle, error, promptTitle, prompt } = messageModel.sheets[0].dataValidations[0];
  assert.deepEqual({ errorTitle, error, promptTitle, prompt }, { errorTitle: 'Invalid status', error: 'Choose an existing status', promptTitle: 'Status', prompt: 'Pick a status' });
  const tooLongMessage = worksheet.replace('type="list"', `type="list" error="${'x'.repeat(256)}"`);
  assert.equal(analyze(zipSync({ ...files, 'xl/worksheets/sheet1.xml': new TextEncoder().encode(tooLongMessage) })).writable, false);
});

test('XLSX 경고·정보 알림 규칙은 파싱·저장 후 종류와 원본 XML을 유지한다', () => {
  const files = unzipSync(fixture());
  const source = new TextDecoder().decode(files['xl/worksheets/sheet1.xml']);
  const validations = '<dataValidations count="2"><dataValidation type="list" errorStyle="warning" sqref="G1"><formula1>"Yes,No"</formula1></dataValidation><dataValidation type="whole" errorStyle="information" operator="greaterThan" sqref="G2"><formula1>0</formula1></dataValidation></dataValidations>';
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(source.replace('</worksheet>', validations + '</worksheet>'));
  const model = analyze(zipSync(files));
  assert.equal(model.writable, true, model.readOnlyReason);
  assert.deepEqual(model.sheets[0].dataValidations.map(rule => rule.errorStyle), ['warning', 'information']);
  const saved = patchCell(model, 0, 'A1', 'edited');
  const reopened = analyze(saved.bytes);
  assert.equal(reopened.writable, true, reopened.readOnlyReason);
  assert.deepEqual(reopened.sheets[0].dataValidations.map(rule => rule.errorStyle), ['warning', 'information']);
  assert.ok(reopened.sheets[0].xml.includes(validations));
});

test('XLSX 다른 시트 범위 목록 원본을 파싱하고 없는 시트 참조를 거부한다', () => {
  const files = unzipSync(fixture());
  const validation = '<dataValidations count="1"><dataValidation type="list" sqref="A1"><formula1>\'Hidden\'!$A$1:$A$1</formula1></dataValidation></dataValidations>';
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</worksheet>', validation + '</worksheet>'));
  const model = analyze(zipSync(files));
  assert.equal(model.writable, true, model.readOnlyReason);
  assert.deepEqual(model.sheets[0].dataValidations[0].source, { sheetName: 'Hidden', top: 0, bottom: 0, left: 0, right: 0 });
  const movedSource = transformStructure(model, { sheetIndex: 1, axis: 'row', action: 'insert', index: 0 }).model;
  assert.deepEqual(movedSource.sheets[0].dataValidations[0].source, { sheetName: 'Hidden', top: 1, bottom: 1, left: 0, right: 0 });
  assert.match(movedSource.sheets[0].xml, /<formula1>&apos;Hidden&apos;!\$A\$2:\$A\$2<\/formula1>/);
  const missingSheet = new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace("'Hidden'!", "'Missing'!");
  const unsupported = analyze(zipSync({ ...files, 'xl/worksheets/sheet1.xml': new TextEncoder().encode(missingSheet) }));
  assert.equal(unsupported.writable, false);
  assert.match(unsupported.readOnlyReason, /목록 원본 시트를 찾을 수 없습니다/);
});

test('데이터 유효성 적용 범위와 원본 수식은 행·열 편집을 따라가고 끊긴 참조를 보존한다', () => {
  const files = unzipSync(fixture()), encode = value => new TextEncoder().encode(value), decode = name => new TextDecoder().decode(files[name]);
  const targetValidations = '<dataValidations count="2"><dataValidation type="whole" operator="between" sqref="B1:B3"><formula1>A1</formula1><formula2>A1+2</formula2></dataValidation><dataValidation type="list" sqref="C1 D2:D3"><formula1>"Yes,No"</formula1></dataValidation></dataValidations>';
  const sourceValidation = '<dataValidations count="1"><dataValidation type="list" sqref="A1"><formula1>Data!$A$1</formula1></dataValidation></dataValidations>';
  files['xl/worksheets/sheet1.xml'] = encode(decode('xl/worksheets/sheet1.xml').replace('</worksheet>', targetValidations + '</worksheet>'));
  files['xl/worksheets/sheet2.xml'] = encode(decode('xl/worksheets/sheet2.xml').replace('</worksheet>', sourceValidation + '</worksheet>'));
  const original = analyze(zipSync(files));
  assert.equal(original.writable, true, original.readOnlyReason);

  const insertedRows = transformStructure(original, { sheetIndex: 0, axis: 'row', action: 'insert', index: 0 }).model;
  assert.equal(insertedRows.writable, true, insertedRows.readOnlyReason);
  assert.deepEqual(insertedRows.sheets[0].dataValidations[0].ranges, [{ top: 1, bottom: 3, left: 1, right: 1 }]);
  assert.deepEqual(insertedRows.sheets[0].dataValidations[0].formulas, ['A2', 'A2+2']);
  assert.deepEqual(insertedRows.sheets[0].dataValidations[1].ranges, [{ top: 1, bottom: 1, left: 2, right: 2 }, { top: 2, bottom: 3, left: 3, right: 3 }]);
  assert.deepEqual(insertedRows.sheets[1].dataValidations[0].source, { sheetName: 'Data', top: 1, bottom: 1, left: 0, right: 0 });
  assert.match(insertedRows.sheets[1].xml, /<formula1>Data!\$A\$2<\/formula1>/);

  const deletedColumn = transformStructure(insertedRows, { sheetIndex: 0, axis: 'column', action: 'delete', index: 0 }).model;
  assert.equal(deletedColumn.writable, true, deletedColumn.readOnlyReason);
  assert.deepEqual(deletedColumn.sheets[0].dataValidations[0].ranges, [{ top: 1, bottom: 3, left: 0, right: 0 }]);
  assert.deepEqual(deletedColumn.sheets[0].dataValidations[0].formulas, ['#REF!', '#REF!+2']);
  assert.deepEqual(deletedColumn.sheets[0].dataValidations[1].ranges, [{ top: 1, bottom: 1, left: 1, right: 1 }, { top: 2, bottom: 3, left: 2, right: 2 }]);
  assert.equal(deletedColumn.sheets[1].dataValidations[0].sourceFormula, '#REF!');
  assert.equal(deletedColumn.sheets[1].dataValidations[0].source, null);

  const singleFiles = unzipSync(fixture());
  singleFiles['xl/worksheets/sheet1.xml'] = encode(new TextDecoder().decode(singleFiles['xl/worksheets/sheet1.xml']).replace('</worksheet>', '<dataValidations count="1"><dataValidation type="list" sqref="A1"><formula1>"Yes,No"</formula1></dataValidation></dataValidations></worksheet>'));
  const removed = transformStructure(analyze(zipSync(singleFiles)), { sheetIndex: 0, axis: 'row', action: 'delete', index: 0 }).model;
  assert.deepEqual(removed.sheets[0].dataValidations, []);
  assert.doesNotMatch(removed.sheets[0].xml, /<dataValidations\b/);
});

test('XLSX 이름 정의 범위 목록은 조회·저장하고 참조 이름의 구조 변경을 지원한다', () => {
  const files = unzipSync(fixture());
  files['xl/workbook.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/workbook.xml']).replace('</workbook>', '<definedNames><definedName name="StatusList">\'Hidden\'!$A$1:$A$1</definedName></definedNames></workbook>'));
  const validation = '<dataValidations count="1"><dataValidation type="list" sqref="A1"><formula1>StatusList</formula1></dataValidation></dataValidations>';
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</worksheet>', validation + '</worksheet>'));
  const model = analyze(zipSync(files));
  assert.equal(model.writable, true, model.readOnlyReason);
  assert.deepEqual(model.definedNames, [{ name: 'StatusList', localSheetId: null, source: { sheetName: 'Hidden', top: 0, bottom: 0, left: 0, right: 0 } }]);
  const saved = patchCell(model, 0, 'B1', 'edited');
  const reopened = analyze(saved.bytes);
  assert.equal(reopened.writable, true, reopened.readOnlyReason);
  assert.deepEqual(reopened.definedNames, model.definedNames);
  assert.ok(reopened.sheets[0].xml.includes(validation));
  const insertedValidation = transformStructure(reopened, { sheetIndex: 0, axis: 'row', action: 'insert', index: 0 }).model;
  assert.deepEqual(insertedValidation.sheets[0].dataValidations[0].ranges, [{ top: 1, bottom: 1, left: 0, right: 0 }]);
  assert.deepEqual(insertedValidation.sheets[0].dataValidations[0].source, { sheetName: 'Hidden', top: 0, bottom: 0, left: 0, right: 0 });
  const unreferencedNames = unzipSync(fixture());
  unreferencedNames['xl/workbook.xml'] = new TextEncoder().encode(new TextDecoder().decode(unreferencedNames['xl/workbook.xml']).replace('</workbook>', '<definedNames><definedName name="StatusList">\'Hidden\'!$A$1:$A$1</definedName></definedNames></workbook>'));
  const namesOnlyModel = analyze(zipSync(unreferencedNames));
  assert.equal(namesOnlyModel.writable, true, namesOnlyModel.readOnlyReason);
  const namesOnlyInserted = transformStructure(namesOnlyModel, { sheetIndex: 0, axis: 'row', action: 'insert', index: 0 }).model;
  assert.equal(namesOnlyInserted.writable, true, namesOnlyInserted.readOnlyReason);
  assert.deepEqual(namesOnlyInserted.definedNames[0].source, { sheetName: 'Hidden', top: 0, bottom: 0, left: 0, right: 0 });
  const scopedNames = unzipSync(zipSync(files));
  scopedNames['xl/workbook.xml'] = new TextEncoder().encode(new TextDecoder().decode(scopedNames['xl/workbook.xml']).replace('</definedNames>', '<definedName name="StatusList" localSheetId="0">$B$1:$B$1</definedName></definedNames>'));
  const scopedModel = analyze(zipSync(scopedNames));
  assert.equal(scopedModel.writable, true, scopedModel.readOnlyReason);
  assert.deepEqual(scopedModel.sheets[0].dataValidations[0].source, { sheetName: 'Data', top: 0, bottom: 0, left: 1, right: 1 }, 'sheet-scoped definitions take precedence over workbook-scoped definitions');
  const formulaWorkbook = new TextDecoder().decode(files['xl/workbook.xml']).replace('</workbook>', '</workbook>');
  const formulaFiles = { ...files, 'xl/workbook.xml': new TextEncoder().encode(formulaWorkbook.replace('</definedNames>', '</definedNames>')) };
  formulaFiles['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</sheetData>', '<row r="2"><c r="A2"><f>StatusList</f><v>9</v></c></row></sheetData>'));
  const namedFormula = analyze(zipSync(formulaFiles));
  assert.equal(namedFormula.writable, true, namedFormula.readOnlyReason);
  const namedDocument = { state: namedFormula, getCell: (row, column, sheetIndex = 0) => [...namedFormula.sheets[sheetIndex].cells.values()].find(cell => cell.row === row && cell.column === column) || null };
  assert.equal(createFormulaCalculator(namedDocument).value(0, 1, 0), 9);
});

test('정의된 이름은 지역 범위 우선 순위와 범위 spill을 계산한다', () => {
  const cells = new Map([
    ['A1', { address: 'A1', row: 0, column: 0, value: '1', kind: 'number' }],
    ['A2', { address: 'A2', row: 1, column: 0, value: '2', kind: 'number' }],
    ['B1', { address: 'B1', row: 0, column: 1, value: '10', kind: 'number' }],
    ['B2', { address: 'B2', row: 1, column: 1, value: '20', kind: 'number' }],
    ['C1', { address: 'C1', row: 0, column: 2, value: '', kind: 'number', formula: 'SUM(Amount)', hasFormula: true }],
    ['D1', { address: 'D1', row: 0, column: 3, value: '', kind: 'number', formula: 'GlobalAmount', hasFormula: true }]
  ]);
  const document = { state: { formulas: true, definedNames: [
    { name: 'Amount', localSheetId: null, source: { sheetName: 'Data', top: 0, bottom: 1, left: 0, right: 0 } },
    { name: 'Amount', localSheetId: 0, source: { sheetName: 'Data', top: 0, bottom: 1, left: 1, right: 1 } },
    { name: 'GlobalAmount', localSheetId: null, source: { sheetName: 'Data', top: 0, bottom: 1, left: 0, right: 0 } }
  ], sheets: [{ name: 'Data', cells, mergedCells: [], tables: [] }] }, getCell: (row, column) => [...cells.values()].find(cell => cell.row === row && cell.column === column) || null };
  const calculator = createFormulaCalculator(document);
  assert.equal(calculator.value(0, 0, 2), 30, 'a sheet-local name takes precedence over a workbook-scoped name');
  assert.deepEqual([calculator.cell(0, 0, 3).value, calculator.cell(0, 1, 3).value], ['1', '2']);
});

test('LET은 순차 변수, 중첩 범위, 정의된 이름 가림과 배열 spill을 계산한다', () => {
  assert.equal(evaluateFormula('LET(x,2,x+3)'), 5);
  assert.equal(evaluateFormula('LET(x,2,y,x*3,y+x)'), 8, 'later bindings can use earlier values');
  assert.equal(evaluateFormula('LET(x,2,LET(x,4,x)+x)'), 6, 'nested bindings restore the outer scope');
  assert.equal(evaluateFormula('LET(Amount,7,Amount)', { getDefinedName: () => 99 }), 7, 'local bindings take precedence over defined names');
  assert.equal(evaluateFormula('LET(x,2,X,3,X)'), '#VALUE!', 'duplicate names are case-insensitive');
  assert.equal(evaluateFormula('LET(A1,2,A1)'), '#VALUE!', 'cell references cannot be binding names');
  assert.equal(evaluateFormula('LET(x,2)'), '#VALUE!', 'a value and final expression are required');
  assert.deepEqual(evaluateFormula('LET(values,SEQUENCE(2),values)', { allowArrayResult: true }), [[1], [2]]);
  const document = { state: { formulas: true, definedNames: [], sheets: [{ name: 'Sheet1', cells: new Map([
    ['A1', { address: 'A1', row: 0, column: 0, value: '', kind: 'number', formula: 'LET(values,SEQUENCE(2),values)', hasFormula: true }]
  ]), mergedCells: [], tables: [] }] }, getCell: (row, column) => [...document.state.sheets[0].cells.values()].find(cell => cell.row === row && cell.column === column) || null };
  const calculator = createFormulaCalculator(document);
  assert.equal(calculator.value(0, 0, 0), 1);
  assert.equal(calculator.value(0, 1, 0), 2, 'LET array results spill into following cells');
});

test('XLSX 숫자·날짜·시간·텍스트 길이 데이터 유효성을 파싱하고 수식 기준도 검사한다', () => {
  const validations = '<dataValidations count="4"><dataValidation type="whole" operator="between" sqref="A1"><formula1>1</formula1><formula2>10</formula2></dataValidation><dataValidation type="decimal" operator="greaterThanOrEqual" sqref="B1"><formula1>0.5</formula1></dataValidation><dataValidation type="date" operator="lessThan" sqref="C1"><formula1>46000</formula1></dataValidation><dataValidation type="textLength" operator="equal" sqref="D1"><formula1>3</formula1></dataValidation></dataValidations>';
  const files = unzipSync(fixture()), worksheet = new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</worksheet>', validations + '</worksheet>');
  const model = analyze(zipSync({ ...files, 'xl/worksheets/sheet1.xml': new TextEncoder().encode(worksheet) }));
  assert.equal(model.writable, true, model.readOnlyReason);
  assert.deepEqual(model.sheets[0].dataValidations.map(rule => [rule.type, rule.operator, rule.values]), [
    ['whole', 'between', [1, 10]], ['decimal', 'greaterThanOrEqual', [0.5]], ['date', 'lessThan', [46000]], ['textLength', 'equal', [3]]
  ]);
  const unsupportedWorksheet = worksheet.replace('<formula1>0.5</formula1>', '<formula1>=A1</formula1>');
  const referenceModel = analyze(zipSync({ ...files, 'xl/worksheets/sheet1.xml': new TextEncoder().encode(unsupportedWorksheet) }));
  assert.equal(referenceModel.writable, true, referenceModel.readOnlyReason);
  assert.deepEqual(referenceModel.sheets[0].dataValidations[1].formulas, ['A1']);
  const unknownFormulaWorksheet = worksheet.replace('<formula1>0.5</formula1>', '<formula1>UNKNOWN_FUNCTION(A1)</formula1>');
  const unsupported = analyze(zipSync({ ...files, 'xl/worksheets/sheet1.xml': new TextEncoder().encode(unknownFormulaWorksheet) }));
  assert.equal(unsupported.writable, false);
  assert.match(unsupported.readOnlyReason, /데이터 유효성/);
});

test('XLSX 데이터 유효성 비교 기준은 지원되는 셀 참조 수식을 보존하고 미지원 수식을 거부한다', () => {
  const validation = '<dataValidations count="1"><dataValidation type="whole" operator="between" sqref="A1:A2"><formula1>B1</formula1><formula2>B1+2</formula2></dataValidation></dataValidations>';
  const files = unzipSync(fixture());
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</worksheet>', validation + '</worksheet>'));
  const model = analyze(zipSync(files));
  assert.equal(model.writable, true, model.readOnlyReason);
  assert.deepEqual(model.sheets[0].dataValidations[0].formulas, ['B1', 'B1+2']);
  const unsupportedXml = validation.replace('B1+2', 'UNKNOWN_VALIDATION_FUNCTION(B1)');
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace(validation, unsupportedXml));
  const unsupported = analyze(zipSync(files));
  assert.equal(unsupported.writable, false);
  assert.match(unsupported.readOnlyReason, /지원하지 않는 데이터 유효성 기준 수식/);
});

test('XLSX supports the final Excel worksheet row for reading and editing', () => {
  const files = unzipSync(fixture());
  let sheet = new TextDecoder().decode(files['xl/worksheets/sheet1.xml']);
  sheet = sheet.replaceAll('A1', 'A1048576').replaceAll('B1', 'B1048576').replaceAll('C1', 'C1048576').replace('r="1"', 'r="1048576"');
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(sheet);
  const model = analyze(zipSync(files));
  assert.equal(model.writable, true);
  assert.equal(model.sheets[0].cells.get('A1048576').value, 'before');
  const saved = patchCell(model, 0, 'A1048576', 'after');
  assert.equal(analyze(saved.bytes).sheets[0].cells.get('A1048576').value, 'after');
});

test('XLSX 열 너비 저장은 범위 정의를 나누고 숨김·스타일 속성과 다른 패키지 항목을 보존한다', () => {
  const original = fixture({ columnDefinitions: '<cols><col min="1" max="3" width="24" customWidth="1" hidden="1" style="0"/><col min="5" max="5" width="9.5"/></cols>' });
  const model = analyze(original);
  const result = patchColumnWidths(model, [{ sheetIndex: 0, column: 1, width: 12.25 }, { sheetIndex: 0, column: 3, width: 18 }]);
  const reopened = analyze(result.bytes);
  assert.deepEqual(reopened.sheets[0].columnWidths, { 0: 24, 1: 12.25, 2: 24, 3: 18, 4: 9.5 });
  assert.match(reopened.sheets[0].xml, /<col min="1" max="1" width="24" customWidth="1" hidden="1" style="0"\/>/);
  assert.match(reopened.sheets[0].xml, /<col min="2" max="2" width="12.25" customWidth="1" hidden="1" style="0"\/>/);
  assert.match(reopened.sheets[0].xml, /<col min="3" max="3" width="24" customWidth="1" hidden="1" style="0"\/>/);
  assert.match(reopened.sheets[0].xml, /<col min="4" max="4" width="18" customWidth="1"\/>/);
  assert.deepEqual(Buffer.from(reopened.files['xl/worksheets/sheet2.xml']), Buffer.from(model.files['xl/worksheets/sheet2.xml']));
  assert.deepEqual(Buffer.from(reopened.files['xl/styles.xml'] || []), Buffer.from(model.files['xl/styles.xml'] || []));
});

test('XLSX 열 너비 저장은 없는 정의를 만들고 범위·형식 오류에서 저장을 거부한다', () => {
  const model = analyze(fixture());
  const result = patchColumnWidths(model, [{ sheetIndex: 0, column: 2, width: 20 }]);
  const reopened = analyze(result.bytes);
  assert.deepEqual(reopened.sheets[0].columnWidths, { 2: 20 });
  assert.match(reopened.sheets[0].xml, /<cols><col min="3" max="3" width="20" customWidth="1"\/><\/cols><sheetData>/);
  assert.throws(() => patchColumnWidths(model, [{ sheetIndex: 0, column: 16384, width: 12 }]), /요청이 유효하지 않습니다/);
  assert.equal(patchColumnWidths(analyze(fixture({ formula: true })), [{ sheetIndex: 0, column: 0, width: 12 }]).model.formulas, true);
});

test('XLSX 행 숨김은 기존 행 속성을 보존하고 빈 행도 생성·해제한다', () => {
  const model = analyze(fixture({ formula: true }));
  const hidden = patchHiddenRows(model, [{ sheetIndex: 0, rows: [1, 2], hidden: true }]);
  assert.deepEqual(hidden.model.sheets[0].hiddenRows, [1, 2]);
  assert.match(hidden.model.sheets[0].xml, /<row r="2" hidden="1">/);
  assert.match(hidden.model.sheets[0].xml, /<row r="3" hidden="1"\/>/);
  const shown = patchHiddenRows(hidden.model, [{ sheetIndex: 0, rows: [1, 2], hidden: false }]);
  assert.deepEqual(shown.model.sheets[0].hiddenRows, []);
  assert.match(shown.model.sheets[0].xml, /<row r="2">/);
  assert.doesNotMatch(shown.model.sheets[0].xml, /<row r="3"/);
  assert.throws(() => patchHiddenRows(model, [{ sheetIndex: 0, rows: [2, 1], hidden: true }]), /요청이 유효하지 않습니다/);
});

test('XLSX 열 숨김은 범위 정의를 분리하고 빈 열도 추가·해제한다', () => {
  const source = fixture({ formula: true });
  const files = unzipSync(source);
  const path = 'xl/worksheets/sheet1.xml';
  files[path] = new TextEncoder().encode(new TextDecoder().decode(files[path]).replace('<sheetData>', '<cols><col min="1" max="4" width="12" customWidth="1"/></cols><sheetData>'));
  const model = analyze(zipSync(files));
  const hidden = patchHiddenColumns(model, [{ sheetIndex: 0, columns: [1, 5], hidden: true }]);
  assert.deepEqual(hidden.model.sheets[0].hiddenColumns, [1, 5]);
  assert.match(hidden.model.sheets[0].xml, /<col min="1" max="1" width="12" customWidth="1"\/>/);
  assert.match(hidden.model.sheets[0].xml, /<col min="2" max="2" width="12" customWidth="1" hidden="1"\/>/);
  assert.match(hidden.model.sheets[0].xml, /<col min="3" max="4" width="12" customWidth="1"\/>/);
  assert.match(hidden.model.sheets[0].xml, /<col min="6" max="6" hidden="1"\/>/);
  const shown = patchHiddenColumns(hidden.model, [{ sheetIndex: 0, columns: [1, 5], hidden: false }]);
  assert.deepEqual(shown.model.sheets[0].hiddenColumns, []);
  assert.doesNotMatch(shown.model.sheets[0].xml, /<col[^>]*hidden="1"/);
  assert.throws(() => patchHiddenColumns(model, [{ sheetIndex: 0, columns: [2, 1], hidden: true }]), /요청이 유효하지 않습니다/);
});

test('XLSX AutoFilter 조건은 XML 이스케이프·재파싱·삭제를 보존한다', () => {
  const model = analyze(fixture({ autoFilter: '<autoFilter ref="A1:C2"/>' }));
  const filter = { top: 0, bottom: 1, left: 0, right: 2, columns: [
    { column: 0, values: ['a&b', 'x<y'], blanks: true },
    { column: 2, values: ['TRUE'], blanks: false }
  ] };
  const patched = patchAutoFilters(model, [{ sheetIndex: 0, filter }]);
  assert.deepEqual(patched.model.sheets[0].autoFilter, filter);
  const xml = new TextDecoder().decode(patched.model.files['xl/worksheets/sheet1.xml']);
  assert.match(xml, /<autoFilter ref="A1:C2"><filterColumn colId="0"><filters blank="1"><filter val="a&amp;b"\/><filter val="x&lt;y"\/><\/filters><\/filterColumn><filterColumn colId="2"><filters><filter val="TRUE"\/><\/filters><\/filterColumn><\/autoFilter>/);
  const cleared = patchAutoFilters(patched.model, [{ sheetIndex: 0, filter: null }]);
  assert.equal(cleared.model.sheets[0].autoFilter, null);
  assert.doesNotMatch(new TextDecoder().decode(cleared.model.files['xl/worksheets/sheet1.xml']), /<autoFilter\b/);
  const noneSelected = patchAutoFilters(model, [{ sheetIndex: 0, filter: { ...filter, columns: [{ column: 0, values: [], blanks: false }] } }]);
  assert.deepEqual(noneSelected.model.sheets[0].autoFilter.columns, [{ column: 0, values: [], blanks: false }], 'an empty checklist is a valid filter that hides all data rows');
  const created = patchAutoFilters(analyze(fixture()), [{ sheetIndex: 0, filter }]);
  assert.deepEqual(created.model.sheets[0].autoFilter, filter, 'the filter XML can be inserted after sheetData when none existed');
  const manyValues = Array.from({ length: 1101 }, (_, index) => 'value-' + index);
  const many = patchAutoFilters(model, [{ sheetIndex: 0, filter: { ...filter, columns: [{ column: 0, values: manyValues, blanks: false }] } }]);
  assert.equal(many.model.sheets[0].autoFilter.columns[0].values.length, 1101, 'the XML writer persists filter lists larger than 1,000 values');
});

test('XLSX standard date/percent and common custom formats affect display only', () => {
  const model = analyze(fixture({ formatted: true }));
  const date = model.sheets[0].cells.get('B1'), percent = model.sheets[0].cells.get('C1');
  assert.equal(model.writable, true);
  assert.equal(date.value, '45292');
  assert.equal(date.numFmtId, 14);
  assert.equal(date.styleProperties, null);
  assert.equal(formatCellValue(date, 'en'), '1/1/24');
  assert.equal(formatCellValue(date, 'ko'), '2024-01-01');
  assert.equal(formatCellValue({ value: '60', kind: 'number', numFmtId: 14 }, 'en'), '2/29/00');
  assert.equal(percent.value, '0.125');
  assert.equal(formatCellValue(percent, 'en'), '12.50%');
  assert.equal(formatCellValue({ value: '1234.5', kind: 'number', numFmtId: 164, formatCode: '#,##0.00' }, 'en'), '1,234.50');
  assert.equal(formatCellValue({ value: '1234.5', kind: 'number', numFmtId: 165, formatCode: '"€"#,##0.00' }, 'en'), '€1,234.50');
  assert.equal(formatCellValue({ value: '1234', kind: 'number', numFmtId: 166, formatCode: '"₩"#,##0' }, 'ko'), '₩1,234');
  assert.equal(formatCellValue({ value: '1234.5', kind: 'number', numFmtId: 5 }, 'en'), '$1,235');
  assert.equal(formatCellValue({ value: '1234.5', kind: 'number', numFmtId: 6 }, 'en'), '$1,235');
  assert.equal(formatCellValue({ value: '1234.5', kind: 'number', numFmtId: 7 }, 'en'), '$1,234.50');
  assert.equal(formatCellValue({ value: '1234.5', kind: 'number', numFmtId: 8 }, 'en'), '$1,234.50');
  assert.equal(formatCellValue({ value: '-1234.5', kind: 'number', numFmtId: 39 }, 'en'), '(1,234.50)');
  assert.equal(formatCellValue({ value: '1234.5', kind: 'number', numFmtId: 166, formatCode: '0 \"days\"' }, 'en'), '1234.5');
  const customDate = model.sheets[0].cells.get('D1');
  assert.equal(customDate.numFmtId, 164);
  assert.equal(customDate.formatCode, 'yyyy-mm-dd');
  assert.deepEqual(customDate.styleProperties, { bold: true, italic: true, color: '#123456', fontSize: 14, fill: '#FFEEAA', horizontal: 'center', vertical: 'center', borders: { left: { color: '#ABCDEF', line: 'thin' } } });
  const currency = patchCells(analyze(fixture({ formatted: true })), [{ sheetIndex: 0, address: 'A1', value: '1234.5', kind: 'number', numberFormatId: 165, numberFormatCode: '"€"#,##0.00' }]);
  assert.equal(currency.model.sheets[0].cells.get('A1').formatCode, '"€"#,##0.00');
  assert.equal(formatCellValue(currency.model.sheets[0].cells.get('A1'), 'en'), '€1,234.50');
  const currencyStyles = new TextDecoder().decode(currency.model.files['xl/styles.xml']);
  assert.match(currencyStyles, /numFmtId="165" formatCode="&quot;€&quot;#,##0\.00"/);
  assert.equal(formatCellValue(customDate, 'en'), '2024-01-01');
  const editedModel = patchCell(model, 0, 'C1', '0.25').model;
  const edited = editedModel.sheets[0].cells.get('C1');
  assert.equal(edited.value, '0.25');
  assert.equal(formatCellValue(edited, 'en'), '25.00%');
  assert.deepEqual(Buffer.from(editedModel.files['xl/styles.xml']), Buffer.from(model.files['xl/styles.xml']));
});

test('명시적 숫자 표시 변경은 기존 사용자 지정 표시 형식을 새 서식으로 교체하고 다른 서식 속성을 보존한다', () => {
  const original = analyze(fixture({ formatted: true }));
  const before = original.sheets[0].cells.get('D1');
  assert.equal(before.numFmtId, 164);
  const changed = patchCells(original, [{ sheetIndex: 0, address: 'D1', value: before.value, kind: 'number', numberFormatId: 20, forceNumberFormat: true }]);
  const after = changed.model.sheets[0].cells.get('D1');
  assert.equal(after.numFmtId, 20);
  assert.deepEqual(after.styleProperties, before.styleProperties, 'font, fill, border and alignment survive a number-format change');
  assert.equal(analyze(changed.bytes).sheets[0].cells.get('D1').numFmtId, 20);
  const currency = patchCells(original, [{ sheetIndex: 0, address: 'D1', value: before.value, kind: 'number', numberFormatId: 165, numberFormatCode: '"$"#,##0.00', forceNumberFormat: true }]);
  assert.equal(currency.model.sheets[0].cells.get('D1').formatCode, '"$"#,##0.00');
  assert.deepEqual(currency.model.sheets[0].cells.get('D1').styleProperties, before.styleProperties);
  const number = patchCells(original, [{ sheetIndex: 0, address: 'C1', value: '0.125', kind: 'number', numberFormatId: 4, forceNumberFormat: true }]);
  assert.equal(number.model.sheets[0].cells.get('C1').formatCode, '#,##0.00');
});

test('100,000 numeric-cell workbook survives distant multi-cell edit, save and reopen', () => {
  const files = unzipSync(fixture());
  const rows = [];
  for (let row = 1; row <= 10000; row++) {
    let cells = '';
    for (let column = 1; column <= 10; column++) {
      const address = String.fromCharCode(64 + column) + row;
      cells += `<c r="${address}"><v>${row * column}</v></c>`;
    }
    rows.push(`<row r="${row}">${cells}</row>`);
  }
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:J10000"/><sheetData>${rows.join('')}</sheetData></worksheet>`);
  files['xl/worksheets/sheet2.xml'] = new TextEncoder().encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>');
  const model = analyze(zipSync(files, { level: 0 }));
  assert.equal(model.sheets[0].cells.size, 100000);
  assert.equal(model.writable, true);
  const untouchedEntries = Object.fromEntries(Object.entries(model.files).filter(([name]) => name !== model.sheets[0].path));
  const output = patchCells(model, [
    { sheetIndex: 0, address: 'A1', value: 'first-edited' },
    { sheetIndex: 0, address: 'J10000', value: 'last-edited' }
  ]).bytes;
  const reopened = analyze(output);
  assert.equal(reopened.writable, true, reopened.readOnlyReason);
  assert.equal(reopened.sheets[0].cells.size, 100000);
  assert.equal(reopened.sheets[0].cells.get('A1').value, 'first-edited');
  assert.equal(reopened.sheets[0].cells.get('J10000').value, 'last-edited');
  assert.equal(reopened.sheets[0].cells.get('E5678').value, String(5678 * 5));
  for (const [name, bytes] of Object.entries(untouchedEntries)) assert.deepEqual(Buffer.from(reopened.files[name]), Buffer.from(bytes), `untouched package entry changed: ${name}`);
});

test("빈 셀에 새 값을 입력하면 셀·행 순서와 사용 범위가 갱신되고 값을 비울 수 있다", () => {
  const model = analyze(fixture());
  const addedCell = patchCell(model, 0, "D1", "12").model;
  assert.equal(addedCell.sheets[0].cells.get("D1").kind, "number");
  assert.equal(addedCell.sheets[0].cells.get("D1").value, "12");
  assert.match(addedCell.sheets[0].xml, /<dimension ref="A1:D1"\/>/);
  const addedRow = patchCell(model, 0, "B2", "new row").model;
  assert.equal(addedRow.sheets[0].cells.get("B2").value, "new row");
  assert.match(addedRow.sheets[0].xml, /<dimension ref="A1:C2"\/>/);
  const multiple = patchCells(model, [
    { sheetIndex: 0, address: "E1", value: "last" },
    { sheetIndex: 0, address: "D1", value: "12" },
    { sheetIndex: 0, address: "B2", value: "second row" },
    { sheetIndex: 0, address: "A2", value: "first cell" }
  ]).model;
  assert.deepEqual([...multiple.sheets[0].cells.values()].filter(cell => cell.row === 0).sort((a, b) => a.column - b.column).map(cell => cell.address), ["A1", "B1", "C1", "D1", "E1"]);
  assert.equal(multiple.sheets[0].cells.get("A2").value, "first cell");
  assert.equal(multiple.sheets[0].cells.get("B2").value, "second row");
  const cleared = patchCell(model, 0, "A1", "").model;
  assert.equal(cleared.sheets[0].cells.get("A1").value, "");
});

test("단순 XLSX에서 행·열 삽입·삭제가 셀 주소와 사용 범위를 함께 이동한다", () => {
  const original = fixture(), model = analyze(original);
  const insertedRow = transformStructure(model, { sheetIndex: 0, axis: "row", action: "insert", index: 0 });
  assert.equal(insertedRow.model.sheets[0].cells.get("A2").value, "before");
  assert.equal(insertedRow.model.sheets[0].cells.get("B2").value, "42");
  assert.match(insertedRow.model.sheets[0].xml, /<dimension ref="A1:C2"\/>/);
  const removedRow = transformStructure(insertedRow.model, { sheetIndex: 0, axis: "row", action: "delete", index: 0 });
  assert.equal(removedRow.model.sheets[0].cells.get("A1").value, "before");
  assert.equal(removedRow.model.sheets[0].cells.has("A2"), false);
  const insertedColumn = transformStructure(model, { sheetIndex: 0, axis: "column", action: "insert", index: 1 });
  assert.equal(insertedColumn.model.sheets[0].cells.get("A1").value, "before");
  assert.equal(insertedColumn.model.sheets[0].cells.get("C1").value, "42");
  assert.equal(insertedColumn.model.sheets[0].cells.get("D1").value, "TRUE");
  assert.match(insertedColumn.model.sheets[0].xml, /<dimension ref="A1:D1"\/>/);
  const removedColumn = transformStructure(insertedColumn.model, { sheetIndex: 0, axis: "column", action: "delete", index: 1 });
  assert.equal(removedColumn.model.sheets[0].cells.get("B1").value, "42");
  assert.equal(removedColumn.model.sheets[0].cells.get("C1").value, "TRUE");
  for (const name of Object.keys(model.files)) {
    if (name !== model.sheets[0].path && name !== 'xl/workbook.xml') assert.deepEqual(Buffer.from(insertedRow.model.files[name]), Buffer.from(model.files[name]));
  }
  const formulaShift = transformStructure(analyze(fixture({ formula: true })), { sheetIndex: 0, axis: "row", action: "insert", index: 0 });
  assert.equal(formulaShift.model.sheets[0].cells.get('A3').formula, '1+1');
});

test('행·열 구조 변경은 단순 AutoFilter 범위와 필터 열 조건을 함께 이동한다', () => {
  const filterXml = '<autoFilter ref="A1:C3"><filterColumn colId="0"><filters><filter val="apple"/></filters></filterColumn><filterColumn colId="2"><filters blank="1"><filter val="TRUE"/></filters></filterColumn></autoFilter>';
  const original = analyze(fixture({ autoFilter: filterXml }));
  const insertedRow = transformStructure(original, { sheetIndex: 0, axis: 'row', action: 'insert', index: 1 }).model;
  assert.deepEqual(insertedRow.sheets[0].autoFilter, { top: 0, bottom: 3, left: 0, right: 2, columns: [
    { column: 0, values: ['apple'], blanks: false }, { column: 2, values: ['TRUE'], blanks: true }
  ] });
  assert.match(insertedRow.sheets[0].xml, /<autoFilter ref="A1:C4"/);

  const shiftedRow = transformStructure(original, { sheetIndex: 0, axis: 'row', action: 'insert', index: 0 }).model;
  assert.deepEqual(shiftedRow.sheets[0].autoFilter, { top: 1, bottom: 3, left: 0, right: 2, columns: [
    { column: 0, values: ['apple'], blanks: false }, { column: 2, values: ['TRUE'], blanks: true }
  ] });

  const insertedColumn = transformStructure(original, { sheetIndex: 0, axis: 'column', action: 'insert', index: 1 }).model;
  assert.equal(insertedColumn.sheets[0].autoFilter.right, 3);
  assert.deepEqual(insertedColumn.sheets[0].autoFilter.columns.map(item => item.column), [0, 3]);
  assert.match(insertedColumn.sheets[0].xml, /<autoFilter ref="A1:D3"/);

  const deletedColumn = transformStructure(original, { sheetIndex: 0, axis: 'column', action: 'delete', index: 1 }).model;
  assert.equal(deletedColumn.sheets[0].autoFilter.right, 1);
  assert.deepEqual(deletedColumn.sheets[0].autoFilter.columns.map(item => item.column), [0, 1]);
  assert.match(deletedColumn.sheets[0].xml, /<autoFilter ref="A1:B3"/);

  const twoRowFilter = analyze(fixture({ autoFilter: filterXml.replace('A1:C3', 'A1:C2') }));
  const removedFilter = transformStructure(twoRowFilter, { sheetIndex: 0, axis: 'row', action: 'delete', index: 1 }).model;
  assert.equal(removedFilter.sheets[0].autoFilter, null, 'deleting the only data row removes an invalid one-row filter range');
  assert.doesNotMatch(removedFilter.sheets[0].xml, /<autoFilter\b/);
});

test("셀 값 편집은 숫자와 불리언 데이터 형식을 보존한다", () => {
  const model = analyze(fixture());
  const numeric = patchCell(model, 0, "B1", "58").model.sheets[0].cells.get("B1");
  assert.equal(numeric.kind, "number");
  assert.equal(numeric.value, "58");
  const boolean = patchCell(model, 0, "C1", "FALSE").model.sheets[0].cells.get("C1");
  assert.equal(boolean.kind, "boolean");
  assert.equal(boolean.value, "FALSE");
  const text = patchCell(model, 0, "B1", "문자").model.sheets[0].cells.get("B1");
  assert.equal(text.kind, "string");
  assert.equal(text.value, "문자");
});

test('표준 단순 XLSX에서 여러 시트·값을 읽고 단일 셀만 수정한다', () => {
  const original = fixture();
  const model = analyze(original);
  assert.equal(model.writable, true);
  assert.equal(model.sheets.length, 2);
  assert.equal(model.sheets[0].cells.get('A1').value, 'before');
  const result = patchCell(model, 0, 'A1', 'after < &');
  assert.equal(result.model.sheets[0].cells.get('A1').value, 'after < &');
  for (const name of Object.keys(model.files)) {
    if (name !== 'xl/worksheets/sheet1.xml') assert.deepEqual(Buffer.from(result.model.files[name]), Buffer.from(model.files[name]));
  }
  assert.equal(result.model.sheets[0].cells.get('B1').value, '42');
});

test('지원되는 차트가 있는 XLSX는 원본 셀 편집 시 차트 캐시를 갱신하고 차트 파트를 보존한다', () => {
  const model = analyze(fixtureWithPositionedChart());
  assert.equal(model.writable, true, model.readOnlyReason);
  assert.equal(model.hasCharts, true);
  assert.deepEqual(model.chartSourceRanges.map(({ sheetIndex, top, bottom, left, right }) => ({ sheetIndex, top, bottom, left, right })), [
    { sheetIndex: 0, top: 0, bottom: 0, left: 1, right: 1 },
    { sheetIndex: 0, top: 1, bottom: 2, left: 0, right: 0 },
    { sheetIndex: 0, top: 1, bottom: 2, left: 1, right: 1 }
  ]);
  assert.deepEqual(model.sheets[0].charts, [{ type: 'bar', direction: 'vertical', title: 'Revenue & Growth', series: [{ name: 'Revenue', categories: ['Jan', 'Feb'], values: [4, 6] }] }]);
  const edited = patchCells(model, [
    { sheetIndex: 0, address: 'B1', value: 'Net sales', kind: 'string' },
    { sheetIndex: 0, address: 'B2', value: '9', kind: 'number' },
    { sheetIndex: 0, address: 'D1', value: 'safe', kind: 'string' }
  ]).model;
  assert.equal(edited.sheets[0].cells.get('D1').value, 'safe');
  assert.equal(edited.sheets[0].charts[0].series[0].name, 'Net sales');
  assert.deepEqual(edited.sheets[0].charts[0].series[0].categories, ['Jan', 'Feb']);
  assert.deepEqual(edited.sheets[0].charts[0].series[0].values, [9, null]);
  assert.notDeepEqual(Buffer.from(edited.files['xl/charts/chart1.xml']), Buffer.from(model.files['xl/charts/chart1.xml']));
  assert.deepEqual(Buffer.from(edited.files['xl/drawings/drawing1.xml']), Buffer.from(model.files['xl/drawings/drawing1.xml']));
});

test('XLSX 행·열 삽입과 삭제는 차트 원본 수식·캐시와 셀 기준 배치를 함께 갱신한다', () => {
  for (const editAs of ['twoCell', 'oneCell', 'absolute']) {
    const original = analyze(fixtureWithPositionedChart(editAs));
    const populated = patchCells(original, [
      { sheetIndex: 0, address: 'A2', value: 'Jan', kind: 'string' },
      { sheetIndex: 0, address: 'A3', value: 'Feb', kind: 'string' },
      { sheetIndex: 0, address: 'B2', value: '4', kind: 'number' },
      { sheetIndex: 0, address: 'B3', value: '6', kind: 'number' }
    ]).model;
    const before = populated.sheets[0].chartAnchors[0];
    const inserted = transformStructure(populated, { sheetIndex: 0, axis: 'row', action: 'insert', index: 2 });
    assert.equal(inserted.model.writable, true, inserted.model.readOnlyReason);
    assert.match(new TextDecoder().decode(inserted.model.files['xl/charts/chart1.xml']), /Data!\$A\$2:\$A\$4/);
    assert.match(new TextDecoder().decode(inserted.model.files['xl/charts/chart1.xml']), /Data!\$B\$2:\$B\$4/);
    assert.deepEqual(inserted.model.sheets[0].charts[0].series[0].categories, ['Jan', '', 'Feb']);
    assert.deepEqual(inserted.model.sheets[0].charts[0].series[0].values, [4, null, 6]);
    if (editAs === 'twoCell') {
      assert.equal(inserted.model.sheets[0].chartAnchors[0].from.row, before.from.row + 1);
      assert.equal(inserted.model.sheets[0].chartAnchors[0].to.row, before.to.row + 1);
      const geometry = require('../src/chart-print').chartPrintGeometry(inserted.model.sheets[0], inserted.model.sheets[0].chartAnchors).boxes[0];
      assert.deepEqual([geometry.y, geometry.height], [78, 156]);
    } else {
      const geometry = require('../src/chart-print').chartPrintGeometry(inserted.model.sheets[0], inserted.model.sheets[0].chartAnchors).boxes[0];
      assert.deepEqual([geometry.x, geometry.y, geometry.width, geometry.height], [288, editAs === 'absolute' ? 52 : 78, 576, 156]);
    }
    const restored = transformStructure(inserted.model, { sheetIndex: 0, axis: 'row', action: 'delete', index: 2 });
    assert.equal(restored.model.writable, true, restored.model.readOnlyReason);
    assert.deepEqual(restored.model.sheets[0].charts[0].series[0].categories, ['Jan', 'Feb']);
    assert.deepEqual(restored.model.sheets[0].charts[0].series[0].values, [4, 6]);
    assert.deepEqual(restored.model.sheets[0].chartAnchors, before ? populated.sheets[0].chartAnchors : []);
    const insertedColumn = transformStructure(populated, { sheetIndex: 0, axis: 'column', action: 'insert', index: 1 });
    assert.equal(insertedColumn.model.writable, true, insertedColumn.model.readOnlyReason);
    assert.match(new TextDecoder().decode(insertedColumn.model.files['xl/charts/chart1.xml']), /Data!\$C\$1/);
    assert.match(new TextDecoder().decode(insertedColumn.model.files['xl/charts/chart1.xml']), /Data!\$C\$2:\$C\$3/);
    assert.deepEqual(insertedColumn.model.sheets[0].charts[0].series[0].values, [4, 6]);
    const restoredColumn = transformStructure(insertedColumn.model, { sheetIndex: 0, axis: 'column', action: 'delete', index: 1 });
    assert.equal(restoredColumn.model.writable, true, restoredColumn.model.readOnlyReason);
    assert.deepEqual(restoredColumn.model.sheets[0].charts[0].series[0].values, [4, 6]);
  }
});

test('차트가 공유한 그리기 파일은 구조 변경한 워크시트에만 분리한다', () => {
  const files = unzipSync(fixtureWithPositionedChart()), drawingPath = 'xl/drawings/drawing1.xml';
  files['xl/worksheets/sheet2.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet2.xml']).replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ').replace('</worksheet>', '<drawing r:id="rId1"/></worksheet>'));
  files['xl/worksheets/_rels/sheet2.xml.rels'] = new TextEncoder().encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>');
  const original = analyze(zipSync(files));
  assert.equal(original.writable, true, original.readOnlyReason);
  const changed = transformStructure(original, { sheetIndex: 0, axis: 'row', action: 'insert', index: 2 }).model;
  assert.notEqual(changed.sheets[0].chartDrawingPath, changed.sheets[1].chartDrawingPath);
  assert.deepEqual(changed.files[drawingPath], original.files[drawingPath]);
  assert.equal(changed.sheets[0].chartAnchors[0].from.row, original.sheets[0].chartAnchors[0].from.row + 1);
  assert.equal(changed.sheets[1].chartAnchors[0].from.row, original.sheets[1].chartAnchors[0].from.row);
  assert.deepEqual(changed.sheets[0].chartPaths, original.sheets[0].chartPaths);
  assert.deepEqual(changed.sheets[1].chartPaths, original.sheets[1].chartPaths);
});

test('차트 외 그리기 개체가 포함된 XLSX는 쓰기 보호를 유지한다', () => {
  const files = unzipSync(fixtureWithCachedChart());
  const path = 'xl/drawings/drawing1.xml';
  files[path] = new TextEncoder().encode(new TextDecoder().decode(files[path]).replace('</xdr:twoCellAnchor>', '<xdr:sp><xdr:nvSpPr/><xdr:spPr/></xdr:sp></xdr:twoCellAnchor>'));
  const model = analyze(zipSync(files));
  assert.equal(model.writable, false);
  assert.match(model.readOnlyReason, /차트 외 그리기 개체/);
});

test('차트 원본 수식 셀을 수정하면 재계산한 수식을 캐시에 반영한다', () => {
  const files = unzipSync(fixtureWithCachedChart());
  const path = 'xl/worksheets/sheet1.xml';
  files[path] = new TextEncoder().encode(new TextDecoder().decode(files[path]).replace('</sheetData>', '<row r="2"><c r="B2"><f>2+2</f><v>4</v></c></row></sheetData>'));
  const model = analyze(zipSync(files));
  const edited = patchCells(model, [{ sheetIndex: 0, address: 'B2', value: '', kind: 'number', formula: '6+3' }]).model;
  assert.equal(edited.sheets[0].cells.get('B2').formula, '6+3');
  assert.deepEqual(edited.sheets[0].charts[0].series[0].values, [9, null]);
});

test('차트 밖의 수식 의존 셀을 바꾸면 차트 원본 수식을 재계산해 저장한다', () => {
  const data = patchCells(analyze(fixture()), [
    { sheetIndex: 0, address: 'A1', value: 'Category' }, { sheetIndex: 0, address: 'B1', value: 'Computed' },
    { sheetIndex: 0, address: 'A2', value: 'One' }, { sheetIndex: 0, address: 'B2', value: '', formula: 'D2*2' },
    { sheetIndex: 0, address: 'D2', value: '', formula: 'Hidden!A1+1' }
  ]).model;
  const chart = patchAddChart(data, 0, { top: 0, bottom: 1, left: 0, right: 1 }).model;
  assert.deepEqual(chart.sheets[0].charts[0].series[0].values, [20]);
  const edited = patchCell(chart, 1, 'A1', '11');
  assert.deepEqual(analyze(edited.bytes).sheets[0].charts[0].series[0].values, [24]);
  const error = patchCells(chart, [{ sheetIndex: 1, address: 'A1', value: '', formula: '1/0' }]);
  assert.deepEqual(error.model.sheets[0].charts[0].series[0].values, [null]);
});

test('혼합 차트와 해석할 수 없는 차트 원본은 읽기 전용으로 유지한다', () => {
  const files = unzipSync(fixtureWithCachedChart());
  const path = 'xl/charts/chart1.xml';
  files[path] = new TextEncoder().encode(new TextDecoder().decode(files[path]).replace('</c:plotArea>', '<c:lineChart><c:ser/></c:lineChart></c:plotArea>'));
  let model = analyze(zipSync(files));
  assert.equal(model.writable, false);
  assert.match(model.readOnlyReason, /차트 형식 또는 원본 범위/);

  const invalid = unzipSync(fixtureWithCachedChart());
  invalid[path] = new TextEncoder().encode(new TextDecoder().decode(invalid[path]).replace('Data!$B$2:$B$3', 'Data!$B:$B'));
  model = analyze(zipSync(invalid));
  assert.equal(model.writable, false);
  assert.match(model.readOnlyReason, /차트 형식 또는 원본 범위/);
});

test('XLSX 원형 차트는 범주별 캐시와 도넛 종류를 미리보기 모델로 읽는다', () => {
  const files = unzipSync(fixtureWithCachedChart());
  const chartPath = 'xl/charts/chart1.xml';
  files[chartPath] = new TextEncoder().encode(new TextDecoder().decode(files[chartPath]).replace('<c:barChart><c:barDir val="col"/>', '<c:doughnutChart>').replace('</c:barChart>', '</c:doughnutChart>'));
  const model = analyze(zipSync(files));
  assert.equal(model.writable, true);
  assert.equal(model.sheets[0].charts[0].type, 'doughnut');
  assert.deepEqual(model.sheets[0].charts[0].series[0].categories, ['Jan', 'Feb']);
  assert.deepEqual(model.sheets[0].charts[0].series[0].values, [4, 6]);
});

test('XLSX 영역형 차트 캐시를 지원 미리보기 모델로 읽는다', () => {
  const files = unzipSync(fixtureWithCachedChart());
  const chartPath = 'xl/charts/chart1.xml';
  files[chartPath] = new TextEncoder().encode(new TextDecoder().decode(files[chartPath]).replace('<c:barChart><c:barDir val="col"/>', '<c:areaChart>').replace('</c:barChart>', '</c:areaChart>'));
  const model = analyze(zipSync(files));
  assert.equal(model.writable, true);
  assert.equal(model.sheets[0].charts[0].type, 'area');
  assert.deepEqual(model.sheets[0].charts[0].series[0].values, [4, 6]);
});

test('XLSX 분산형 차트는 숫자 X·Y 캐시를 쌍으로 읽는다', () => {
  const files = unzipSync(fixtureWithCachedChart());
  const chartPath = 'xl/charts/chart1.xml';
  let xml = new TextDecoder().decode(files[chartPath]).replace('<c:barChart><c:barDir val="col"/>', '<c:scatterChart>').replace('</c:barChart>', '</c:scatterChart>');
  xml = xml.replace(/<c:cat>.*?<\/c:cat>/, '<c:xVal><c:numRef><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:xVal>');
  xml = xml.replace(/<c:val>.*?<\/c:val>/, '<c:yVal><c:numRef><c:numCache><c:pt idx="0"><c:v>4</c:v></c:pt><c:pt idx="1"><c:v>6</c:v></c:pt></c:numCache></c:numRef></c:yVal>');
  files[chartPath] = new TextEncoder().encode(xml);
  const model = analyze(zipSync(files));
  assert.equal(model.writable, true);
  assert.equal(model.sheets[0].charts[0].type, 'scatter');
  assert.deepEqual(model.sheets[0].charts[0].series[0].xValues, [1, 2]);
  assert.deepEqual(model.sheets[0].charts[0].series[0].values, [4, 6]);
});

test('수식 통합문서는 일반 셀 편집을 허용하고 수식 셀·수식 참조 변경 작업은 보호한다', () => {
  const model = analyze(fixture({ formula: true }));
  assert.equal(model.writable, true);
  assert.equal(model.formulas, true);
  assert.equal(model.sheets[0].cells.get('A2').formula, '1+1');
  assert.equal(model.sheets[0].cells.get('A2').cachedValue, '2');
  const edited = patchCell(model, 0, 'A1', 'updated');
  assert.equal(edited.model.sheets[0].cells.get('A1').value, 'updated');
  assert.equal(edited.model.sheets[0].cells.get('A2').formula, '1+1');
  assert.equal(edited.model.sheets[0].cells.get('A2').cachedValue, '2');
  assert.match(new TextDecoder().decode(edited.model.files['xl/workbook.xml']), /<calcPr[^>]*fullCalcOnLoad="1"[^>]*forceFullCalc="1"/);
  assert.throws(() => patchCell(model, 0, 'A2', 'blocked'), /기존의 단순 값 셀만 수정/);
  const structureFiles = unzipSync(fixture({ formula: true }));
  let firstSheet = new TextDecoder().decode(structureFiles['xl/worksheets/sheet1.xml']).replace('<f>1+1</f><v>2</v>', '<f>SUM(A1:B1)</f><v>44</v>');
  structureFiles['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(firstSheet);
  structureFiles['xl/worksheets/sheet2.xml'] = new TextEncoder().encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>9</v></c></row><row r="2"><c r="B2"><f>Data!A1</f><v>1</v></c></row></sheetData></worksheet>');
  const formulaStructure = analyze(zipSync(structureFiles));
  const insertedWithFormulas = transformStructure(formulaStructure, { sheetIndex: 0, axis: 'row', action: 'insert', index: 0 });
  const formulaReopened = analyze(insertedWithFormulas.bytes);
  assert.equal(formulaReopened.sheets[0].cells.get('A3').formula, 'SUM(A2:B2)');
  assert.equal(formulaReopened.sheets[0].cells.get('A3').cachedValue, '');
  assert.equal(formulaReopened.sheets[1].cells.get('B2').formula, 'Data!A2');
  assert.equal(formulaReopened.writable, true, formulaReopened.readOnlyReason);
  assert.match(new TextDecoder().decode(formulaReopened.files['xl/workbook.xml']), /fullCalcOnLoad="1"/);
  const deletedWithFormulas = transformStructure(formulaReopened, { sheetIndex: 0, axis: 'row', action: 'delete', index: 1 });
  assert.equal(deletedWithFormulas.model.sheets[1].cells.get('B2').formula, '#REF!');
  const sharedFilesStructure = unzipSync(sharedFormulaFixture());
  const sharedStructure = analyze(zipSync(sharedFilesStructure));
  for (const [operation, expected] of [
    [{ sheetIndex: 0, axis: 'row', action: 'insert', index: 0 }, [['A3', 'B2+1'], ['A4', 'B3+1']]],
    [{ sheetIndex: 0, axis: 'column', action: 'insert', index: 0 }, [['B2', 'C1+1'], ['B3', 'C2+1']]]
  ]) {
    const transformed = transformStructure(sharedStructure, operation), transformedSheet = transformed.model.sheets[0];
    assert.equal(transformed.model.writable, true, transformed.model.readOnlyReason);
    assert.equal(transformedSheet.sharedFormulaGroups.length, 0);
    assert.doesNotMatch(transformedSheet.xml, /t="shared"/);
    for (const [address, formula] of expected) {
      assert.equal(transformedSheet.cells.get(address).formula, formula);
      assert.equal(transformedSheet.cells.get(address).formulaType, null);
      assert.equal(transformedSheet.cells.get(address).cachedValue, '');
    }
    const letters = column => String.fromCharCode(65 + column);
    const calculator = createFormulaCalculator({
      state: transformed.model,
      getCell: (row, column, sheetIndex = 0) => transformed.model.sheets[sheetIndex]?.cells.get(letters(column) + (row + 1)) || null
    });
    if (operation.axis === 'row') {
      assert.equal(calculator.value(0, 2, 0), 43);
      assert.equal(calculator.value(0, 3, 0), 1);
    } else {
      assert.equal(calculator.value(0, 1, 1), 43);
      assert.equal(calculator.value(0, 2, 1), 1);
    }
  }
  const arrayFiles = unzipSync(fixture({ formula: true }));
  arrayFiles['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(new TextDecoder().decode(arrayFiles['xl/worksheets/sheet1.xml']).replace('<f>1+1</f>', '<f t="array" ref="A2">1+1</f>'));
  assert.throws(() => transformStructure(analyze(zipSync(arrayFiles)), { sheetIndex: 0, axis: 'row', action: 'insert', index: 0 }), /배열 수식/);
  const widthChanged = patchColumnWidths(model, [{ sheetIndex: 0, column: 0, width: 12 }]);
  assert.equal(widthChanged.model.sheets[0].cells.get('A2').formula, '1+1');
  const sharedFormula = analyze(sharedFormulaFixture());
  assert.equal(sharedFormula.writable, true);
  assert.equal(sharedFormula.sheets[0].cells.get('A2').hasFormula, true);
  assert.throws(() => patchCell(sharedFormula, 0, 'A2', 'blocked'), /기존의 단순 값 셀만 수정/);
  const editedSharedFormula = patchCells(sharedFormula, [{ sheetIndex: 0, address: 'A3', value: '0', formula: 'B2+10' }]);
  assert.equal(editedSharedFormula.model.sheets[0].cells.get('A2').formula, 'B1+1');
  assert.equal(editedSharedFormula.model.sheets[0].cells.get('A2').formulaType, null);
  assert.equal(editedSharedFormula.model.sheets[0].cells.get('A3').formula, 'B2+10');
  assert.equal(editedSharedFormula.model.sheets[0].cells.get('A3').formulaType, null);
  assert.doesNotMatch(editedSharedFormula.model.sheets[0].xml, /t="shared"/);
  const restructuredShared = transformStructure(editedSharedFormula.model, { sheetIndex: 0, axis: 'row', action: 'insert', index: 0 });
  assert.equal(restructuredShared.model.sheets[0].cells.get('A4').formula, 'B3+10');
  const overwrittenSharedFormula = patchCells(sharedFormula, [{ sheetIndex: 0, address: 'A2', value: 'replaced', formula: null, kind: 'string' }]);
  assert.equal(overwrittenSharedFormula.model.sheets[0].cells.get('A2').hasFormula, false);
  assert.equal(overwrittenSharedFormula.model.sheets[0].cells.get('A3').formula, 'B2+1');
  assert.doesNotMatch(overwrittenSharedFormula.model.sheets[0].xml, /t="shared"/);
  const errorFiles = unzipSync(fixture({ formula: true }));
  errorFiles['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(new TextDecoder().decode(errorFiles['xl/worksheets/sheet1.xml']).replace('<c r="A2"><f>1+1</f><v>2</v></c>', '<c r="A2" t="e"><f>1+1</f><v>#N/A</v></c>'));
  const errorFormula = analyze(zipSync(errorFiles));
  assert.equal(errorFormula.writable, true);
  assert.equal(errorFormula.sheets[0].cells.get('A2').kind, 'error');
  assert.equal(patchCell(errorFormula, 0, 'A1', 'ordinary edit').model.sheets[0].cells.get('A2').value, '#N/A');
  assert.throws(() => patchCell(errorFormula, 0, 'A2', 'blocked'), /기존의 단순 값 셀만 수정/);
  const calcFiles = unzipSync(fixture({ formula: true }));
  calcFiles['xl/workbook.xml'] = new TextEncoder().encode(new TextDecoder().decode(calcFiles['xl/workbook.xml']).replace('</workbook>', '<calcPr calcId="191029" calcMode="manual"/></workbook>'));
  const calcModel = analyze(zipSync(calcFiles));
  const calcEdited = patchCell(calcModel, 0, 'A1', 'updated');
  const calcXml = new TextDecoder().decode(calcEdited.model.files['xl/workbook.xml']);
  assert.match(calcXml, /calcId="191029"/);
  assert.match(calcXml, /calcMode="manual"/);
  assert.equal((calcXml.match(/<calcPr/g) || []).length, 1, 'existing calculation settings are updated instead of duplicated');
  const pastedFormula = patchCells(analyze(fixture()), [{ sheetIndex: 0, address: 'D3', value: '', formula: 'A1+$B$1' }]).model;
  assert.equal(pastedFormula.formulas, true);
  assert.equal(pastedFormula.sheets[0].cells.get('D3').formula, 'A1+$B$1');
  assert.match(new TextDecoder().decode(pastedFormula.files['xl/workbook.xml']), /fullCalcOnLoad="1"/);
});

test('XLSX 공유 수식은 각 셀의 상대 참조로 계산하고 저장 중 원본 그룹을 보존한다', () => {
  const model = analyze(sharedFormulaFixture());
  assert.equal(model.writable, true, model.readOnlyReason);
  assert.equal(model.sheets[0].cells.get('A2').formula, 'B1+1');
  assert.equal(model.sheets[0].cells.get('A3').formula, 'B2+1');
  const document = {
    state: model,
    getCell(row, column, sheetIndex = 0) {
      let letters = '';
      for (let value = column + 1; value > 0; value = Math.floor((value - 1) / 26)) letters = String.fromCharCode(65 + (value - 1) % 26) + letters;
      return model.sheets[sheetIndex]?.cells.get(letters + (row + 1)) || null;
    }
  };
  const calculator = createFormulaCalculator(document);
  assert.equal(calculator.value(0, 1, 0), 43);
  assert.equal(calculator.value(0, 2, 0), 1);

  const edited = patchCell(model, 0, 'B1', '7');
  const xml = edited.model.sheets[0].xml;
  assert.match(xml, /<f t="shared" ref="A2:A3" si="0">B1\+1<\/f>/);
  const reopened = analyze(edited.bytes);
  assert.equal(reopened.sheets[0].cells.get('A2').formula, 'B1+1');
  assert.equal(reopened.sheets[0].cells.get('A3').formula, 'B2+1');
  const reopenedDocument = { ...document, state: reopened };
  reopenedDocument.getCell = (row, column, sheetIndex = 0) => {
    let letters = '';
    for (let value = column + 1; value > 0; value = Math.floor((value - 1) / 26)) letters = String.fromCharCode(65 + (value - 1) % 26) + letters;
    return reopened.sheets[sheetIndex]?.cells.get(letters + (row + 1)) || null;
  };
  const reopenedCalculator = createFormulaCalculator(reopenedDocument);
  assert.equal(reopenedCalculator.value(0, 1, 0), 8);
  assert.equal(reopenedCalculator.value(0, 2, 0), 1);

  const malformedFiles = unzipSync(sharedFormulaFixture());
  malformedFiles['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(new TextDecoder().decode(malformedFiles['xl/worksheets/sheet1.xml']).replace('>B1+1</f>', '></f>'));
  const malformed = analyze(zipSync(malformedFiles));
  assert.equal(malformed.writable, false);
  assert.match(malformed.readOnlyReason, /공유 수식 그룹.*기준 수식/);
});

test('알 수 없는 패키지 구성요소는 표시를 허용하되 저장을 막는다', () => {
  const model = analyze(fixture({ extra: { name: 'xl/customXml/item1.xml', content: '<custom/>' } }));
  assert.equal(model.sheets[0].cells.get('A1').value, 'before');
  assert.equal(model.writable, false);
  assert.match(model.readOnlyReason, /지원하지 않는 OOXML/);
});

test('구조 편집에서 참조를 갱신할 수 없는 워크시트 메타데이터는 읽기 전용으로 보호한다', () => {
  const features = [
    ['scenarios', '<scenarios><scenario name="Case1" locked="1"><inputCells><inputCell r="B2" val="1"/></inputCells></scenario></scenarios>'],
    ['dataConsolidate', '<dataConsolidate ref="A1:B2"/>'],
    ['customSheetViews', '<customSheetViews><customSheetView guid="{00000000-0000-0000-0000-000000000001}"/></customSheetViews>'],
    ['phoneticPr', '<phoneticPr fontId="0" type="noConversion"/>'],
    ['customProperties', '<customProperties><customProperty name="Record" value="A1"/></customProperties>'],
    ['cellWatches', '<cellWatches><cellWatch r="A1"/></cellWatches>'],
    ['ignoredErrors', '<ignoredErrors><ignoredError sqref="A1" numberStoredAsText="1"/></ignoredErrors>'],
    ['smartTags', '<smartTags><smartTag r="A1" type="urn:example"/></smartTags>'],
    ['legacyDrawingHF', '<legacyDrawingHF/>'],
    ['picture', '<picture/>'],
    ['webPublishItems', '<webPublishItems/>']
  ];
  for (const [feature, xml] of features) {
    const files = unzipSync(fixture());
    const worksheet = new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</worksheet>', xml + '</worksheet>');
    files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(worksheet);
    const bytes = zipSync(files), model = analyze(bytes);
    assert.equal(model.sheets[0].cells.get('A1').value, 'before', feature + ' remains readable');
    assert.equal(model.writable, false, feature + ' must not be accepted for editing');
    assert.match(model.readOnlyReason, new RegExp('워크시트 기능 ' + feature));
    assert.throws(() => patchCell(model, 0, 'A1', 'changed'), /현재 저장을 지원하지 않습니다/);
    assert.deepEqual(Buffer.from(model.bytes), Buffer.from(bytes), feature + ' keeps the source package unchanged');
  }
});

test('차트 삽입은 선택한 표 데이터를 차트로 만들고 XLSX 재파싱에서 유지한다', () => {
  const base = analyze(fixture());
  const data = patchCells(base, [
    { sheetIndex: 0, address: 'A1', value: 'Month' }, { sheetIndex: 0, address: 'B1', value: 'Sales' },
    { sheetIndex: 0, address: 'A2', value: 'Jan' }, { sheetIndex: 0, address: 'B2', value: '10' },
    { sheetIndex: 0, address: 'A3', value: 'Feb' }, { sheetIndex: 0, address: 'B3', value: '20' }
  ]).model;
  const added = patchAddChart(data, 0, { top: 0, bottom: 2, left: 0, right: 1 });
  assert.equal(added.model.writable, true, added.model.readOnlyReason);
  assert.deepEqual(added.model.sheets[0].charts, [{ type: 'bar', direction: 'vertical', title: 'Chart', series: [{ name: 'Sales', categories: ['Jan', 'Feb'], values: [10, 20] }] }]);
  assert.ok(added.model.files['xl/charts/chart1.xml']);
  assert.ok(added.model.files['xl/drawings/drawing1.xml']);
  assert.equal(analyze(added.bytes).sheets[0].charts[0].series[0].values[1], 20);
  assert.deepEqual(Buffer.from(added.model.files['xl/workbook.xml']), Buffer.from(data.files['xl/workbook.xml']));
});

test('차트 제목과 계열 이름을 수정하고 빈 제목 삭제·계열 이름 변경을 재파싱한다', () => {
  const model = analyze(fixtureWithCachedChart());
  const edited = patchChartMetadata(model, 0, 0, { title: 'Quarterly <Revenue> & cost', seriesNames: ['Net revenue'] });
  assert.equal(edited.model.sheets[0].charts[0].title, 'Quarterly <Revenue> & cost');
  assert.equal(edited.model.sheets[0].charts[0].series[0].name, 'Net revenue');
  assert.equal(analyze(edited.bytes).sheets[0].charts[0].title, 'Quarterly <Revenue> & cost');
  assert.deepEqual(Buffer.from(edited.model.files['xl/drawings/drawing1.xml']), Buffer.from(model.files['xl/drawings/drawing1.xml']));
  const cleared = patchChartMetadata(edited.model, 0, 0, { title: '', seriesNames: [''] });
  assert.equal(cleared.model.sheets[0].charts[0].title, '');
  assert.equal(cleared.model.sheets[0].charts[0].series[0].name, 'Series 1');
  assert.throws(() => patchChartMetadata(model, 0, 0, { title: 'x'.repeat(513) }), /512자/);
  assert.throws(() => patchChartMetadata(model, 0, 0, { seriesNames: [] }), /계열 이름/);
});

test('차트 제목 저장은 DrawingML 텍스트 네임스페이스와 필수 본문 요소를 사용한다', () => {
  for (const defaultNamespace of [false, true]) {
    const files = unzipSync(fixtureWithCachedChart());
    if (defaultNamespace) {
      const xml = new TextDecoder().decode(files['xl/charts/chart1.xml']).replaceAll('xmlns:c=', 'xmlns=').replaceAll('<c:', '<').replaceAll('</c:', '</');
      files['xl/charts/chart1.xml'] = new TextEncoder().encode(xml);
    }
    const edited = patchChartMetadata(analyze(zipSync(files)), 0, 0, { title: '  한글 & Revenue  ' });
    const tags = [], parser = new SaxesParser({ xmlns: true });
    parser.on('opentag', tag => { if (['bodyPr', 'lstStyle', 'p', 'r', 't'].includes(tag.local)) tags.push(tag); });
    parser.write(new TextDecoder().decode(edited.model.files['xl/charts/chart1.xml'])).close();
    assert.deepEqual(tags.map(tag => tag.local), ['bodyPr', 'lstStyle', 'p', 'r', 't']);
    assert.ok(tags.every(tag => tag.uri === 'http://schemas.openxmlformats.org/drawingml/2006/main'));
    assert.equal(edited.model.sheets[0].charts[0].title, '  한글 & Revenue  ');
  }
});

test('제목만 바꾸면 계열 이름의 셀 연결을 보존하고 동일한 적용은 변경을 만들지 않는다', () => {
  const model = analyze(fixtureWithCachedChart());
  const noChange = patchChartMetadata(model, 0, 0, { title: model.sheets[0].charts[0].title, seriesNames: ['Revenue'] });
  assert.equal(noChange.model, model);
  const titleOnly = patchChartMetadata(model, 0, 0, { title: 'New title', seriesNames: ['Revenue'] });
  const afterSourceEdit = patchCell(titleOnly.model, 0, 'B1', 'Renamed at source');
  assert.equal(afterSourceEdit.model.sheets[0].charts[0].series[0].name, 'Renamed at source');
  const cleared = patchChartMetadata(titleOnly.model, 0, 0, { title: '' });
  assert.equal(patchChartMetadata(cleared.model, 0, 0, { title: 'Restored' }).model.sheets[0].charts[0].title, 'Restored');
});

test('차트 원본 범위 변경은 계열 증감·캐시를 갱신하고 제목·서식·다른 차트를 보존한다', () => {
  const data = patchCells(analyze(fixture()), [
    { sheetIndex: 0, address: 'A1', value: 'Month' }, { sheetIndex: 0, address: 'B1', value: 'Sales' }, { sheetIndex: 0, address: 'C1', value: 'Costs' },
    { sheetIndex: 0, address: 'A2', value: 'Jan' }, { sheetIndex: 0, address: 'B2', value: '10', kind: 'number' }, { sheetIndex: 0, address: 'C2', value: '4', kind: 'number' },
    { sheetIndex: 0, address: 'A3', value: 'Feb' }, { sheetIndex: 0, address: 'B3', value: '20', kind: 'number' }, { sheetIndex: 0, address: 'C3', value: '8', kind: 'number' }
  ]).model;
  for (const type of ['bar', 'line', 'area', 'pie', 'doughnut']) {
    for (const defaultNamespace of [false, true]) {
      let model = patchAddChart(data, 0, { top: 0, bottom: 1, left: 0, right: 1 }, type).model;
      model = patchAddChart(model, 0, { top: 0, bottom: 1, left: 0, right: 1 }).model;
      const files = { ...model.files }, path = 'xl/charts/chart1.xml';
      let xml = new TextDecoder().decode(files[path]).replace('</c:tx><c:cat>', '</c:tx><c:spPr><a:solidFill><a:srgbClr val="AABBCC"/></a:solidFill></c:spPr><c:cat>');
      if (defaultNamespace) xml = xml.replaceAll('xmlns:c=', 'xmlns=').replaceAll('<c:', '<').replaceAll('</c:', '</');
      files[path] = new TextEncoder().encode(xml);
      model = analyze(zipSync(files));
      const single = ['pie', 'doughnut'].includes(type), range = { top: 0, bottom: 2, left: 0, right: single ? 1 : 2 };
      const edited = patchChartSource(model, 0, 0, range);
      const chart = analyze(edited.bytes).sheets[0].charts[0];
      assert.equal(chart.type, type); assert.equal(chart.title, 'Chart');
      assert.equal(chart.series.length, single ? 1 : 2);
      assert.deepEqual(chart.series[0].categories, ['Jan', 'Feb']);
      assert.deepEqual(chart.series[0].values, [10, 20]);
      assert.match(new TextDecoder().decode(edited.model.files[path]), /srgbClr val="AABBCC"/);
      for (const [name, bytes] of Object.entries(model.files)) if (name !== path) assert.deepEqual(Buffer.from(edited.model.files[name]), Buffer.from(bytes), name);
      const refreshed = patchCell(edited.model, 0, 'B3', '30').model;
      assert.equal(refreshed.sheets[0].charts[0].series[0].values[1], 30);
      const shrunk = patchChartSource(edited.model, 0, 0, { top: 0, bottom: 1, left: 0, right: 1 }).model;
      assert.equal(shrunk.sheets[0].charts[0].series.length, 1);
      assert.deepEqual(shrunk.sheets[0].charts[0].series[0].values, [10]);
      if (single) assert.throws(() => patchChartSource(model, 0, 0, { ...range, right: 2 }), /계열 하나/);
      assert.throws(() => patchChartSource(model, 0, 0, { ...range, bottom: 1001 }), /1,000셀/);
    }
  }
});

test('분산형 차트 원본 변경은 숫자 X·수식 Y 범위를 저장하고 잘못된 X 값을 거부한다', () => {
  const files = unzipSync(fixtureWithCachedChart()), path = 'xl/charts/chart1.xml';
  let xml = new TextDecoder().decode(files[path]).replace('<c:barChart><c:barDir val="col"/>', '<c:scatterChart>').replace('</c:barChart>', '</c:scatterChart>');
  xml = xml.replace(/<c:cat>.*?<\/c:cat>/, '<c:xVal><c:numRef><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:xVal>');
  xml = xml.replace(/<c:val>.*?<\/c:val>/, '<c:yVal><c:numRef><c:numCache><c:pt idx="0"><c:v>4</c:v></c:pt></c:numCache></c:numRef></c:yVal>');
  files[path] = new TextEncoder().encode(xml);
  const data = patchCells(analyze(zipSync(files)), [
    { sheetIndex: 0, address: 'A2', value: '2', kind: 'number' }, { sheetIndex: 0, address: 'A3', value: '5', kind: 'number' },
    { sheetIndex: 0, address: 'B2', value: '', formula: 'A2*3' }, { sheetIndex: 0, address: 'B3', value: '15', kind: 'number' }
  ]).model;
  const edited = patchChartSource(data, 0, 0, { top: 0, bottom: 2, left: 0, right: 1 });
  assert.deepEqual(edited.model.sheets[0].charts[0].series[0].xValues, [2, 5]);
  assert.deepEqual(edited.model.sheets[0].charts[0].series[0].values, [6, 15]);
  const invalid = patchCell(data, 0, 'A2', 'text').model;
  assert.throws(() => patchChartSource(invalid, 0, 0, { top: 0, bottom: 2, left: 0, right: 1 }), /숫자 X/);
});

test('분산형 차트 삽입은 공통 X·복수 Y 계열과 숫자 축·수식·빈 좌표를 저장한다', () => {
  const data = patchCells(analyze(fixture()), [
    { sheetIndex: 0, address: 'A1', value: 'X' }, { sheetIndex: 0, address: 'B1', value: 'Y one' }, { sheetIndex: 0, address: 'C1', value: 'Y two' },
    { sheetIndex: 0, address: 'A2', value: '0', kind: 'number' }, { sheetIndex: 0, address: 'A3', value: '', formula: '5*2' }, { sheetIndex: 0, address: 'A4', value: '-2', kind: 'number' },
    { sheetIndex: 0, address: 'B2', value: '', formula: 'A2+1' }, { sheetIndex: 0, address: 'B3', value: '5', kind: 'number' }, { sheetIndex: 0, address: 'B4', value: '-3', kind: 'number' }, { sheetIndex: 0, address: 'B5', value: '8', kind: 'number' },
    { sheetIndex: 0, address: 'C2', value: '2', kind: 'number' }, { sheetIndex: 0, address: 'C3', value: '7', kind: 'number' }, { sheetIndex: 0, address: 'C4', value: '-6', kind: 'number' }
  ]).model;
  const added = patchAddChart(data, 0, { top: 0, bottom: 4, left: 0, right: 2 }, 'scatter');
  const chart = analyze(added.bytes).sheets[0].charts[0];
  assert.equal(chart.type, 'scatter'); assert.equal(chart.series.length, 2);
  assert.deepEqual(chart.series[0].xValues, [0, 10, -2, null]);
  assert.deepEqual(chart.series[0].values, [1, 5, -3, 8]);
  assert.deepEqual(chart.series[1].xValues, [0, 10, -2, null]);
  assert.deepEqual(chart.series[1].values, [2, 7, -6, null]);
  const xml = new TextDecoder().decode(added.model.files['xl/charts/chart1.xml']);
  assert.equal((xml.match(/<c:valAx>/g) || []).length, 2);
  assert.doesNotMatch(xml, /<c:catAx>/);
  assert.equal((xml.match(/<c:symbol val="circle"/g) || []).length, 2);
  assert.match(xml, /<c:scatterStyle val="marker"/);
  const changed = patchCells(added.model, [{ sheetIndex: 0, address: 'A3', value: '', formula: '6*2' }]).model;
  assert.equal(changed.sheets[0].charts[0].series[0].xValues[1], 12);
  const invalid = patchCell(data, 0, 'A2', 'text').model;
  assert.throws(() => patchAddChart(invalid, 0, { top: 0, bottom: 2, left: 0, right: 1 }, 'scatter'), /숫자 X/);
  assert.throws(() => patchAddChart(data, 0, { top: 0, bottom: 1, left: 3, right: 4 }, 'scatter'), /숫자 X/);
});

test('차트 삽입은 막대·선·영역·원형·도넛 유형을 재파싱하고 원형 계열을 제한한다', () => {
  const base = analyze(fixture());
  const data = patchCells(base, [
    { sheetIndex: 0, address: 'A1', value: 'Month' }, { sheetIndex: 0, address: 'B1', value: 'Sales' }, { sheetIndex: 0, address: 'C1', value: 'Costs' },
    { sheetIndex: 0, address: 'A2', value: 'Jan' }, { sheetIndex: 0, address: 'B2', value: '10' }, { sheetIndex: 0, address: 'C2', value: '4' },
    { sheetIndex: 0, address: 'A3', value: 'Feb' }, { sheetIndex: 0, address: 'B3', value: '20' }, { sheetIndex: 0, address: 'C3', value: '8' }
  ]).model;
  for (const type of ['bar', 'line', 'area', 'pie', 'doughnut']) {
    const range = { top: 0, bottom: 2, left: 0, right: ['pie', 'doughnut'].includes(type) ? 1 : 2 };
    const added = patchAddChart(data, 0, range, type);
    assert.equal(added.model.sheets[0].charts[0].type, type);
  }
  assert.throws(() => patchAddChart(data, 0, { top: 0, bottom: 2, left: 0, right: 2 }, 'pie'), /계열 하나/);
  assert.throws(() => patchAddChart(data, 0, { top: 0, bottom: 2, left: 0, right: 1 }, 'unknown'), /차트 유형/);
});

test('기존 차트가 있는 시트에 새 차트를 추가해도 그리기와 관계를 보존한다', () => {
  const existing = analyze(fixtureWithCachedChart());
  assert.equal(existing.writable, true, existing.readOnlyReason);
  const data = patchCells(existing, [
    { sheetIndex: 0, address: 'C1', value: 'Quarter' }, { sheetIndex: 0, address: 'D1', value: 'Profit' },
    { sheetIndex: 0, address: 'C2', value: 'Q1' }, { sheetIndex: 0, address: 'D2', value: '5' },
    { sheetIndex: 0, address: 'C3', value: 'Q2' }, { sheetIndex: 0, address: 'D3', value: '8' }
  ]).model;
  const added = patchAddChart(data, 0, { top: 0, bottom: 2, left: 2, right: 3 });
  assert.equal(added.model.writable, true, added.model.readOnlyReason);
  assert.equal(added.model.sheets[0].charts.length, 2);
  assert.equal(added.model.sheets[0].charts[0].title, 'Revenue & Growth');
  assert.deepEqual(added.model.sheets[0].charts[1].series[0].values, [5, 8]);
  assert.match(new TextDecoder().decode(added.model.files['xl/drawings/_rels/drawing1.xml.rels']), /rId2/);
  assert.equal(analyze(added.bytes).sheets[0].charts.length, 2);
});

test('차트 삭제는 선택 차트만 제거하고 마지막 삭제 후 패키지를 정리한다', () => {
  const data = patchCells(analyze(fixtureWithCachedChart()), [
    { sheetIndex: 0, address: 'A2', value: 'Mar' }, { sheetIndex: 0, address: 'B2', value: '5', kind: 'number' },
    { sheetIndex: 0, address: 'A3', value: 'Apr' }, { sheetIndex: 0, address: 'B3', value: '8', kind: 'number' }
  ]).model;
  const original = patchAddChart(data, 0, { top: 0, bottom: 2, left: 0, right: 1 }).model;
  for (const index of [0, 1]) {
    const removedPath = original.sheets[0].chartPaths[index], retainedPath = original.sheets[0].chartPaths[1 - index];
    const first = patchDeleteChart(original, 0, index);
    assert.equal(first.model.sheets[0].charts.length, 1);
    assert.equal(first.model.files[removedPath], undefined);
    assert.deepEqual(Buffer.from(first.model.files[retainedPath]), Buffer.from(original.files[retainedPath]));
    assert.deepEqual(Buffer.from(first.model.files['xl/worksheets/sheet1.xml']), Buffer.from(original.files['xl/worksheets/sheet1.xml']));
    const last = patchDeleteChart(first.model, 0, 0);
    assert.equal(analyze(last.bytes).hasCharts, false);
    for (const name of ['xl/charts/chart1.xml', 'xl/charts/chart2.xml', 'xl/drawings/drawing1.xml', 'xl/drawings/_rels/drawing1.xml.rels', 'xl/worksheets/_rels/sheet1.xml.rels']) assert.equal(last.model.files[name], undefined, name);
    assert.doesNotMatch(last.model.sheets[0].xml, /<drawing\b/);
    assert.doesNotMatch(new TextDecoder().decode(last.model.files['[Content_Types].xml']), /PartName="\/xl\/(?:charts|drawings)\//);
    assert.equal(last.model.sheets[0].cells.get('B2').value, '5');
    assert.equal(patchAddChart(last.model, 0, { top: 0, bottom: 2, left: 0, right: 1 }).model.sheets[0].charts.length, 1);
  }
  assert.throws(() => patchDeleteChart(original, 0, 99), /삭제할 차트/);
  assert.throws(() => patchDeleteChart({ ...original, writable: false, readOnlyReason: '읽기 전용' }, 0, 0), /읽기 전용/);
});

test('공유된 차트 파일과 그리기는 다른 참조가 남아 있으면 삭제하지 않는다', () => {
  const data = patchCells(analyze(fixtureWithCachedChart()), [{ sheetIndex: 0, address: 'A2', value: 'Jan' }, { sheetIndex: 0, address: 'B2', value: '5', kind: 'number' }]).model;
  const two = patchAddChart(data, 0, { top: 0, bottom: 1, left: 0, right: 1 }).model;
  const files = { ...two.files }, encode = value => new TextEncoder().encode(value), decode = name => new TextDecoder().decode(files[name]);
  files['xl/worksheets/sheet2.xml'] = encode(decode('xl/worksheets/sheet2.xml').replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ').replace('</worksheet>', '<drawing r:id="rId1"/></worksheet>'));
  files['xl/worksheets/_rels/sheet2.xml.rels'] = encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>');
  const shared = analyze(zipSync(files)); assert.equal(shared.writable, true, shared.readOnlyReason);
  const edited = patchDeleteChart(shared, 0, 0).model;
  assert.equal(edited.sheets[0].charts.length, 1); assert.equal(edited.sheets[1].charts.length, 2);
  assert.deepEqual(Buffer.from(edited.files['xl/drawings/drawing1.xml']), Buffer.from(shared.files['xl/drawings/drawing1.xml']));
  assert.ok(edited.files['xl/drawings/drawing2.xml']); assert.ok(edited.files['xl/charts/chart1.xml']);
  const detached = patchDeleteChart(edited, 0, 0).model;
  assert.equal(detached.files['xl/drawings/drawing2.xml'], undefined);
  assert.equal(detached.sheets[1].charts.length, 2);
  const oneLeft = patchDeleteChart(detached, 1, 0).model;
  assert.ok(oneLeft.files['xl/charts/chart2.xml']);
  assert.equal(patchDeleteChart(oneLeft, 1, 0).model.hasCharts, false);

  const duplicate = { ...two.files };
  duplicate['xl/drawings/_rels/drawing1.xml.rels'] = encode(new TextDecoder().decode(duplicate['xl/drawings/_rels/drawing1.xml.rels']).replace('chart2.xml', 'chart1.xml'));
  delete duplicate['xl/charts/chart2.xml'];
  duplicate['[Content_Types].xml'] = encode(new TextDecoder().decode(duplicate['[Content_Types].xml']).replace(/<Override PartName="\/xl\/charts\/chart2.xml"[^>]*\/>/, ''));
  const sharedPart = analyze(zipSync(duplicate)); assert.equal(sharedPart.writable, true, sharedPart.readOnlyReason);
  const remaining = patchDeleteChart(sharedPart, 0, 0).model;
  assert.ok(remaining.files['xl/charts/chart1.xml']);
  assert.equal(patchDeleteChart(remaining, 0, 0).model.files['xl/charts/chart1.xml'], undefined);
});

test('외부 엔터티와 DTD가 있는 XML을 거부한다', () => {
  const bytes = fixture();
  assert.throws(() => analyze(bytes.subarray(0, 10)), /ZIP/);
  const hostile = fixture({ workbookXml: '<!DOCTYPE x [<!ENTITY a SYSTEM "file:///etc/passwd">]><workbook>&a;</workbook>' });
  assert.throws(() => analyze(hostile), /XML .* 거부/);
});

test('중복·상위 경로 ZIP 항목, 크기 변조와 CRC 오류를 거부한다', () => {
  assert.throws(() => analyze(zipSync({ '../escape.xml': new Uint8Array([1]) })), /경로/);
  const sizeTampered = Buffer.from(zipSync({ 'part.xml': new TextEncoder().encode('<x/>') }));
  const directory = sizeTampered.indexOf(Buffer.from([0x50, 0x4b, 0x01, 0x02]));
  sizeTampered.writeUInt32LE(64 * 1024 * 1024 + 1, directory + 24);
  assert.throws(() => inspectZip(sizeTampered), /크기 제한/);
  const crcTampered = Buffer.from(zipSync({ 'part.xml': new TextEncoder().encode('<x/>') }));
  const crcDirectory = crcTampered.indexOf(Buffer.from([0x50, 0x4b, 0x01, 0x02]));
  crcTampered.writeUInt32LE((crcTampered.readUInt32LE(crcDirectory + 16) ^ 0xffffffff) >>> 0, crcDirectory + 16);
  assert.throws(() => inspectZip(crcTampered), /CRC/);
});


test("실제 JinSheet XLSX 샘플은 편집·행 삽입 후 재파싱되고 비수정 패키지 항목을 보존한다", () => {
  const fs = require("node:fs");
  const path = require("node:path");
  const originalBytes = fs.readFileSync(path.join(__dirname, "fixtures", "JinSheet-Sample.xlsx"));
  const original = analyze(originalBytes);
  assert.equal(original.writable, true, original.readOnlyReason);
  assert.equal(original.sheets[0].cells.get("B2").value, "12");

  const edited = patchCell(original, 0, "B2", "13").model;
  const changed = transformStructure(edited, { sheetIndex: 0, axis: "row", action: "insert", index: 0 });
  const reopened = analyze(changed.bytes);
  assert.equal(reopened.writable, true, reopened.readOnlyReason);
  assert.equal(reopened.sheets[0].cells.get("A2").value, "상품");
  assert.equal(reopened.sheets[0].cells.get("B3").value, "13");
  assert.equal(reopened.sheets[0].cells.get("C3").value, "1500");

  for (const [name, bytes] of Object.entries(original.files)) {
    if (name === original.sheets[0].path || name === 'xl/workbook.xml') continue;
    assert.deepEqual(Buffer.from(reopened.files[name]), Buffer.from(bytes), `untouched package entry changed: ${name}`);
  }
});

test('LibreOffice가 재저장한 기본 차트 XLSX는 편집·차트 캐시 갱신 후에도 다시 열 수 있다', () => {
  const fs = require('node:fs');
  const path = require('node:path');
  const bytes = fs.readFileSync(path.join(__dirname, 'fixtures', 'LibreOffice-Charts-Roundtrip.xlsx'));
  const model = analyze(bytes);
  assert.equal(model.writable, true, model.readOnlyReason);
  assert.deepEqual(model.sheets[0].charts.map(chart => chart.type), ['bar', 'line', 'area', 'pie', 'doughnut']);
  for (const chart of model.sheets[0].charts) {
    assert.deepEqual(chart.series[0].categories, ['사과', '배']);
    assert.deepEqual(chart.series[0].values, [12, 7]);
  }

  const edited = patchCell(model, 0, 'B2', '13').model;
  const reopened = analyze(edited.bytes);
  assert.equal(reopened.writable, true, reopened.readOnlyReason);
  assert.equal(reopened.sheets[0].cells.get('B2').value, '13');
  assert.deepEqual(reopened.sheets[0].charts.map(chart => chart.type), ['bar', 'line', 'area', 'pie', 'doughnut']);
  for (const chart of reopened.sheets[0].charts) assert.deepEqual(chart.series[0].values, [13, 7]);
});

test('LibreOffice Calc가 다시 저장한 서식 XLSX는 편집 후에도 서식과 패키지 구성요소를 보존한다', () => {
  const fs = require('node:fs');
  const path = require('node:path');
  const source = analyze(fs.readFileSync(path.join(__dirname, 'fixtures', 'LibreOffice-Styled-Roundtrip.xlsx')));
  assert.equal(source.writable, true, source.readOnlyReason);
  assert.equal(source.sheets[0].cells.get('A1').styleProperties.fill, '#1F4E78');
  assert.equal(source.sheets[0].cells.get('B2').formatCode, '#,##0');

  const reopened = analyze(patchCell(source, 0, 'B2', '13').bytes);
  assert.equal(reopened.writable, true, reopened.readOnlyReason);
  assert.equal(reopened.sheets[0].cells.get('B2').value, '13');
  assert.equal(reopened.sheets[0].cells.get('B2').formatCode, '#,##0');
  assert.deepEqual(reopened.sheets[0].cells.get('A1').styleProperties, source.sheets[0].cells.get('A1').styleProperties);
  assert.deepEqual(reopened.sheets[0].rowHeights, source.sheets[0].rowHeights);
  for (const [name, bytes] of Object.entries(source.files)) {
    if (name === source.sheets[0].path || name === 'xl/workbook.xml') continue;
    assert.deepEqual(Buffer.from(reopened.files[name]), Buffer.from(bytes), `untouched LibreOffice package entry changed: ${name}`);
  }
});

test('LibreOffice Calc가 저장한 복합 XLSX의 표·유효성·조건부 서식·차트를 편집 후에도 유지한다', () => {
  const fs = require('node:fs');
  const path = require('node:path');
  const source = analyze(fs.readFileSync(path.join(__dirname, 'fixtures', 'LibreOffice-Mixed-Features-Roundtrip.xlsx')));
  assert.equal(source.writable, true, source.readOnlyReason);
  assert.equal(source.sheets[0].tables[0].name, 'Inventory');
  assert.equal(source.sheets[0].tables[0].totalsRowCount, 1);
  assert.equal(source.sheets[0].tables[0].autoFilter.columns[0].values[0], '12');
  assert.equal(source.sheets[0].dataValidations.length, 1);
  assert.deepEqual(source.sheets[0].dataValidations[0].values, [1, 100]);
  assert.equal(source.sheets[0].conditionalFormats.length, 1);
  assert.deepEqual(source.sheets[0].conditionalFormats[0].style, { fill: '#F4CCCC' });
  assert.deepEqual(source.sheets[0].mergedCells.map(({ top, bottom, left, right }) => ({ top, bottom, left, right })), [{ top: 0, bottom: 0, left: 4, right: 5 }]);
  assert.equal(source.sheets[0].cells.get('D1').formula, 'SUM(B2:B3)');
  assert.equal(source.sheets[0].charts.length, 1);
  assert.deepEqual(source.sheets[0].freezePane, { rows: 1, columns: 1 });

  const reopened = analyze(patchCell(source, 0, 'B2', '13').bytes);
  assert.equal(reopened.writable, true, reopened.readOnlyReason);
  assert.equal(reopened.sheets[0].cells.get('B2').value, '13');
  assert.equal(reopened.sheets[0].tables[0].name, 'Inventory');
  assert.deepEqual(reopened.sheets[0].tables[0].autoFilter.columns[0].values, ['12']);
  assert.equal(reopened.sheets[0].dataValidations.length, 1);
  assert.equal(reopened.sheets[0].conditionalFormats.length, 1);
  assert.equal(reopened.sheets[0].cells.get('B4').formula, 'SUBTOTAL(109,[[#data],[수량]])');
  assert.deepEqual(reopened.sheets[0].charts[0].series[0].values, [13, 7]);
  assert.deepEqual(reopened.sheets[0].mergedCells.map(({ top, bottom, left, right }) => ({ top, bottom, left, right })), [{ top: 0, bottom: 0, left: 4, right: 5 }]);
  assert.equal(reopened.sheets[0].charts.length, 1);
  assert.deepEqual(reopened.sheets[0].freezePane, { rows: 1, columns: 1 });
});

test('LibreOffice 복합 XLSX의 사용자 지정 문서 속성과 행·열 그룹을 편집 후 보존한다', () => {
  const fs = require('node:fs');
  const path = require('node:path');
  const source = analyze(fs.readFileSync(path.join(__dirname, 'fixtures', 'LibreOffice-Mixed-Outline-Roundtrip.xlsx')));
  assert.equal(source.writable, true, source.readOnlyReason);
  assert.deepEqual(source.sheets[0].rowGroups, [{ start: 1, end: 2, level: 1, summary: 3, collapsed: false }]);
  assert.deepEqual(source.sheets[0].columnGroups, [{ start: 1, end: 2, level: 1, summary: 3, collapsed: false }]);
  assert.ok(source.files['docProps/custom.xml']);

  const edited = analyze(patchCell(source, 0, 'B2', '1600').bytes);
  assert.equal(edited.writable, true, edited.readOnlyReason);
  assert.equal(edited.sheets[0].cells.get('B2').value, '1600');
  assert.deepEqual(edited.files['docProps/custom.xml'], source.files['docProps/custom.xml']);
  assert.deepEqual(edited.sheets[0].rowGroups, source.sheets[0].rowGroups);
  assert.deepEqual(edited.sheets[0].columnGroups, source.sheets[0].columnGroups);
  assert.equal(edited.sheets[0].tables[0].name, 'Inventory');
  assert.equal(edited.sheets[0].charts.length, 1);

  const missingProperties = { ...source.files };
  delete missingProperties['docProps/custom.xml'];
  assert.throws(() => analyze(zipSync(missingProperties)), /사용자 지정 문서 속성 관계 대상을 찾을 수 없습니다/);
});

test('XLSX 미리보기 썸네일은 편집 가능 상태와 저장 후 패키지 보존을 유지한다', () => {
  const fs = require('node:fs');
  const path = require('node:path');
  const files = { ...analyze(fixture()).files };
  const rootRels = new TextDecoder().decode(files['_rels/.rels']).replace('</Relationships>', '<Relationship Id="rThumb" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" Target="docProps/thumbnail.png"/></Relationships>');
  const contentTypes = new TextDecoder().decode(files['[Content_Types].xml']).replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>');
  files['_rels/.rels'] = new TextEncoder().encode(rootRels);
  files['[Content_Types].xml'] = new TextEncoder().encode(contentTypes);
  files['docProps/thumbnail.png'] = fs.readFileSync(path.join(__dirname, '..', 'media', 'icon.png'));
  const source = analyze(zipSync(files));
  assert.equal(source.writable, true, source.readOnlyReason);

  const edited = analyze(patchCell(source, 0, 'A1', 'Edited').bytes);
  assert.equal(edited.writable, true, edited.readOnlyReason);
  assert.equal(edited.sheets[0].cells.get('A1').value, 'Edited');
  assert.equal(Buffer.compare(Buffer.from(edited.files['docProps/thumbnail.png']), Buffer.from(files['docProps/thumbnail.png'])), 0);
  const missingThumbnail = { ...files };
  delete missingThumbnail['docProps/thumbnail.png'];
  assert.throws(() => analyze(zipSync(missingThumbnail)), /문서 미리보기 썸네일 관계 대상을 찾을 수 없습니다/);
});

test('LibreOffice 호환 boolean 처리 뒤에도 활성 미지원 규칙과 잘못된 boolean은 편집을 차단한다', () => {
  const conditionalFiles = unzipSync(fixture({ formatted: true }));
  const worksheet = new TextDecoder().decode(conditionalFiles['xl/worksheets/sheet1.xml']);
  const styles = new TextDecoder().decode(conditionalFiles['xl/styles.xml']);
  conditionalFiles['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(worksheet.replace('</worksheet>', '<conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="0" priority="1" aboveAverage="1" operator="greaterThan"><formula>10</formula></cfRule></conditionalFormatting></worksheet>'));
  conditionalFiles['xl/styles.xml'] = new TextEncoder().encode(styles.replace('</styleSheet>', '<dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs></styleSheet>'));
  const activeRule = analyze(zipSync(conditionalFiles));
  assert.equal(activeRule.writable, false);
  assert.match(activeRule.readOnlyReason, /조건부 서식/);

  const validationFiles = unzipSync(fixture());
  const validationSheet = new TextDecoder().decode(validationFiles['xl/worksheets/sheet1.xml']);
  validationFiles['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(validationSheet.replace('</worksheet>', '<dataValidations count="1"><dataValidation type="whole" operator="between" allowBlank="yes" sqref="A1"><formula1>0</formula1><formula2>10</formula2></dataValidation></dataValidations></worksheet>'));
  const invalidBoolean = analyze(zipSync(validationFiles));
  assert.equal(invalidBoolean.writable, false);
  assert.match(invalidBoolean.readOnlyReason, /데이터 유효성/);
});

test('XLSX 표·유효성·병합·수식·조건부 서식·차트가 함께 있을 때 편집 후 전부 보존된다', () => {
  const fs = require('node:fs');
  const path = require('node:path');
  const source = analyze(fs.readFileSync(path.join(__dirname, 'fixtures', 'JinSheet-Sample.xlsx')));
  const files = { ...source.files };
  const worksheet = source.sheets[0].xml.replace('</sheetData>', '</sheetData><dataValidations count="1"><dataValidation type="whole" operator="between" sqref="B2:B3"><formula1>1</formula1><formula2>100</formula2></dataValidation></dataValidations>');
  files[source.sheets[0].path] = new TextEncoder().encode(worksheet);
  let model = analyze(zipSync(files));
  model = patchAddTable(model, { sheetIndex: 0, range: { top: 0, bottom: 2, left: 0, right: 2 }, name: 'Inventory' }).model;
  model = patchTableTotals(model, { sheetIndex: 0, tableName: 'Inventory', enabled: true, label: 'Total' }).model;
  model = patchTableTotals(model, { sheetIndex: 0, tableName: 'Inventory', column: 1, function: 'sum' }).model;
  const tableFilter = { ...model.sheets[0].tables[0].autoFilter, columns: [{ column: 1, values: ['12'], blanks: false }] };
  model = patchAutoFilters(model, [{ sheetIndex: 0, filter: tableFilter }]).model;
  model = patchFilterVisibility(model).model;
  model = patchCells(model, [{ sheetIndex: 0, address: 'D1', value: '', formula: 'SUM(B2:B3)' }]).model;
  model = patchCellValueConditionalFormat(model, 0, { top: 1, bottom: 2, left: 1, right: 2 }, 'greaterThan', [10]).model;
  model = patchAddChart(model, 0, { top: 0, bottom: 2, left: 0, right: 1 }, 'bar').model;
  model = patchMergedCells(model, [{ sheetIndex: 0, ranges: [{ top: 0, bottom: 0, left: 4, right: 5 }] }]).model;

  const original = analyze(model.bytes);
  assert.equal(original.writable, true, original.readOnlyReason);
  assert.equal(original.sheets[0].tables[0].name, 'Inventory');
  assert.equal(original.sheets[0].tables[0].totalsRowCount, 1);
  assert.equal(original.sheets[0].cells.get('B4').formula, 'SUBTOTAL(109,[[#Data],[수량]])');
  assert.deepEqual(original.sheets[0].tables[0].autoFilter.columns, tableFilter.columns);
  assert.deepEqual(original.sheets[0].nativeHiddenRows, [2]);
  assert.equal(original.sheets[0].dataValidations.length, 1);
  assert.deepEqual(original.sheets[0].mergedCells.map(({ top, bottom, left, right }) => ({ top, bottom, left, right })), [{ top: 0, bottom: 0, left: 4, right: 5 }]);
  assert.equal(original.sheets[0].cells.get('D1').formula, 'SUM(B2:B3)');
  assert.equal(original.sheets[0].conditionalFormats.length, 1);
  assert.equal(original.sheets[0].charts.length, 1);
  const originalCalculator = createFormulaCalculator({
    state: { formulas: true, sheets: original.sheets, definedNames: original.definedNames, date1904: original.date1904 },
    getCell(row, column, sheetIndex = 0) {
      let letters = '';
      for (let value = column + 1; value > 0; value = Math.floor((value - 1) / 26)) letters = String.fromCharCode(65 + (value - 1) % 26) + letters;
      return original.sheets[sheetIndex]?.cells.get(letters + (row + 1)) || null;
    }
  });
  assert.equal(originalCalculator.value(0, 0, 3), 19);
  assert.equal(originalCalculator.value(0, 3, 1), 12, 'the total row sums only the record left visible by the filter');

  const edited = patchCell(original, 0, 'B2', '13').model;
  const reopened = analyze(patchFilterVisibility(edited).bytes);
  assert.equal(reopened.writable, true, reopened.readOnlyReason);
  assert.equal(reopened.sheets[0].cells.get('B2').value, '13');
  assert.equal(reopened.sheets[0].tables[0].name, 'Inventory');
  assert.equal(reopened.sheets[0].tables[0].totalsRowCount, 1);
  assert.equal(reopened.sheets[0].cells.get('B4').formula, 'SUBTOTAL(109,[[#Data],[수량]])');
  assert.deepEqual(reopened.sheets[0].tables[0].autoFilter.columns, tableFilter.columns);
  assert.deepEqual(reopened.sheets[0].nativeHiddenRows, [1, 2]);
  assert.equal(reopened.sheets[0].dataValidations.length, 1);
  assert.deepEqual(reopened.sheets[0].mergedCells.map(({ top, bottom, left, right }) => ({ top, bottom, left, right })), [{ top: 0, bottom: 0, left: 4, right: 5 }]);
  assert.equal(reopened.sheets[0].cells.get('D1').formula, 'SUM(B2:B3)');
  assert.equal(reopened.sheets[0].conditionalFormats.length, 1);
  assert.deepEqual(reopened.sheets[0].charts[0].series[0].values, [13, 7]);
  const editedCalculator = createFormulaCalculator({
    state: { formulas: true, sheets: reopened.sheets, definedNames: reopened.definedNames, date1904: reopened.date1904 },
    getCell(row, column, sheetIndex = 0) {
      let letters = '';
      for (let value = column + 1; value > 0; value = Math.floor((value - 1) / 26)) letters = String.fromCharCode(65 + (value - 1) % 26) + letters;
      return reopened.sheets[sheetIndex]?.cells.get(letters + (row + 1)) || null;
    }
  });
  assert.equal(editedCalculator.value(0, 0, 3), 20);
  assert.equal(editedCalculator.value(0, 3, 1), 0, 'the total row recalculates after the edit makes every data row fail the filter');
});


test('실제 서식 샘플의 셀 색상·글꼴·정렬·테두리를 파싱하고 안전 편집을 허용한다', () => {
  const fs = require('node:fs');
  const path = require('node:path');
  const bytes = fs.readFileSync(path.join(__dirname, 'fixtures', 'JinSheet-Styled-Sample.xlsx'));
  const model = analyze(bytes);
  assert.equal(model.writable, true, model.readOnlyReason);
  assert.equal(model.sheets[0].cells.get('A1').styleProperties.fill, '#1F4E78');
  assert.equal(model.sheets[0].cells.get('A1').styleProperties.color, '#FFFFFF');
  assert.equal(model.sheets[0].cells.get('A1').styleProperties.bold, true);
  assert.equal(model.sheets[0].cells.get('A1').styleProperties.horizontal, 'center');
  assert.equal(model.sheets[0].cells.get('A1').styleProperties.borders.left.color, '#B4C7E7');
  assert.equal(model.sheets[0].cells.get('C2').styleProperties.horizontal, 'right');
  assert.equal(formatCellValue(model.sheets[0].cells.get('C2'), 'en'), '1,500');
});

test('페이지 설정은 용지·방향·배율·여백·눈금선과 맞춤 모드를 저장한다', () => {
  const model = analyze(fixture()), sheet = model.sheets[0];
  assert.equal(sheet.printSettings.paperSize, 9); assert.equal(sheet.printSettings.scale, 100);
  const settings = { ...sheet.printSettings, orientation: 'landscape', paperSize: 8, scale: 75, fitToPage: true, fitToWidth: 1, fitToHeight: 0, gridLines: true, headings: true, margins: { left: .4, right: .5, top: .6, bottom: .7, header: .2, footer: .2 } };
  const result = patchPrintSettings(model, 0, settings);
  assert.deepEqual(result.model.sheets[0].printSettings, settings);
  assert.ok(result.model.sheets[0].xml.indexOf('<sheetPr') < result.model.sheets[0].xml.indexOf('<dimension'));
  assert.ok(result.model.sheets[0].xml.indexOf('<printOptions') < result.model.sheets[0].xml.indexOf('<pageMargins'));
  assert.ok(result.model.sheets[0].xml.indexOf('<pageMargins') < result.model.sheets[0].xml.indexOf('<pageSetup'));
  for (const [name, bytes] of Object.entries(model.files)) if (name !== sheet.path) assert.deepEqual(result.model.files[name], bytes);
  assert.equal(patchPrintSettings(result.model, 0, settings).model, result.model, 'reapplying identical settings is a no-op');
});

test('페이지 설정은 기존 속성·머리글·바닥글·필터를 보존한다', () => {
  const model = analyze(fixture({ autoFilter: '<autoFilter ref="A1:C2"/><printOptions horizontalCentered="1"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="portrait" blackAndWhite="1"/><headerFooter><oddHeader>PRIVATE HEADER</oddHeader></headerFooter>', sheetViews: '<sheetPr codeName="Sheet1"><outlinePr summaryRight="0"/><pageSetUpPr autoPageBreaks="0"/></sheetPr>' }));
  const result = patchPrintSettings(model, 0, { ...model.sheets[0].printSettings, orientation: 'landscape' }).model;
  assert.match(result.sheets[0].xml, /codeName="Sheet1"/); assert.match(result.sheets[0].xml, /autoPageBreaks="0"/);
  assert.match(result.sheets[0].xml, /blackAndWhite="1"/); assert.match(result.sheets[0].xml, /horizontalCentered="1"/);
  assert.match(result.sheets[0].xml, /<oddHeader>PRIVATE HEADER<\/oddHeader>/); assert.deepEqual(result.sheets[0].autoFilter, model.sheets[0].autoFilter);
});

test('유효하지 않은 페이지 설정은 패키지를 변경하지 않는다', () => {
  const model = analyze(fixture()), base = model.sheets[0].printSettings;
  for (const update of [{ orientation: 'evil' }, { paperSize: 99 }, { scale: 0 }, { fitToWidth: -1 }, { fitToHeight: 1.5 }, { gridLines: 'yes' }, { margins: { ...base.margins, left: 5, right: 5 } }, { margins: null }]) assert.throws(() => patchPrintSettings(model, 0, { ...base, ...update }));
  assert.throws(() => patchPrintSettings(model, 99, base));
  assert.throws(() => patchPrintSettings({ ...model, writable: false }, 0, base));
  assert.equal(model.sheets[0].xml.includes('pageSetup'), false);
});

test('페이지 설정이 있는 시트도 구조 변경 후 설정과 관련 없는 부품을 보존한다', () => {
  const model = analyze(fixture());
  const configured = patchPrintSettings(model, 0, { ...model.sheets[0].printSettings, scale: 90, headings: true }).model;
  const result = transformStructure(configured, { sheetIndex: 0, axis: 'row', action: 'insert', index: 0, count: 1 }).model;
  assert.deepEqual(result.sheets[0].printSettings, configured.sheets[0].printSettings);
  assert.equal(result.sheets[0].cells.get('A2').value, 'before');
});

test('인쇄 영역은 시트별 이름 정의로 저장·교체·지우고 다른 부품을 보존한다', () => {
  const original = analyze(fixture());
  const area = { top: 1, bottom: 100, left: 1, right: 20 };
  const configured = patchPrintArea(original, 0, area).model;
  assert.deepEqual(configured.sheets[0].printAreas, [{ sheetName: 'Data', ...area }]);
  assert.deepEqual(configured.sheets[1].printAreas, []);
  assert.match(new TextDecoder().decode(configured.files['xl/workbook.xml']), /name="_xlnm.Print_Area" localSheetId="0"/);
  for (const [name, bytes] of Object.entries(original.files)) if (name !== 'xl/workbook.xml') assert.deepEqual(configured.files[name], bytes);
  assert.equal(patchPrintArea(configured, 0, area).model, configured);
  const other = patchPrintArea(configured, 1, { top: 0, bottom: 0, left: 0, right: 0 }).model;
  const cleared = patchPrintArea(other, 0, null).model;
  assert.deepEqual(cleared.sheets[0].printAreas, []); assert.equal(cleared.sheets[1].printAreas.length, 1);
  assert.equal(patchPrintArea(cleared, 0, null).model, cleared);
  assert.throws(() => patchPrintArea(original, 0, { ...area, bottom: 1048576 }), /인쇄/);
  assert.throws(() => patchPrintArea(original, 0, { ...area, left: -1 }), /인쇄/);
});

test('인쇄 영역은 행·열 삽입·삭제에 맞게 이동·확장·축소하고 전부 삭제되면 제거한다', () => {
  const original = analyze(fixture()), range = { top: 1, bottom: 3, left: 1, right: 2 };
  const configured = patchPrintArea(original, 0, range).model;
  const operate = (model, axis, action, index) => transformStructure(model, { sheetIndex: 0, axis, action, index }).model;
  assert.deepEqual(operate(configured, 'row', 'insert', 0).sheets[0].printAreas[0], { sheetName: 'Data', ...range, top: 2, bottom: 4 });
  assert.equal(operate(configured, 'row', 'insert', 2).sheets[0].printAreas[0].bottom, 4);
  assert.equal(operate(configured, 'row', 'delete', 2).sheets[0].printAreas[0].bottom, 2);
  assert.equal(operate(configured, 'column', 'insert', 1).sheets[0].printAreas[0].left, 2);
  assert.equal(operate(configured, 'column', 'delete', 1).sheets[0].printAreas[0].right, 1);
  const single = patchPrintArea(original, 0, { top: 0, bottom: 0, left: 0, right: 0 }).model;
  assert.deepEqual(operate(single, 'row', 'delete', 0).sheets[0].printAreas, []);
});

test('인쇄 영역은 시트 이름·순서 변경과 시트 삭제 후에도 올바른 시트에 연결된다', () => {
  const original = analyze(fixture());
  const configured = patchPrintArea(original, 0, { top: 0, bottom: 1, left: 0, right: 2 }).model;
  const renamed = patchSheetName(configured, 0, "O'Brien, Q1").model;
  assert.equal(renamed.sheets[0].printAreas[0].sheetName, "O'Brien, Q1");
  const moved = patchMoveSheet(renamed, 0, 1).model;
  assert.equal(moved.sheets[1].name, "O'Brien, Q1"); assert.equal(moved.sheets[1].printAreas.length, 1); assert.deepEqual(moved.sheets[0].printAreas, []);
  assert.equal(moved.definedNames[0].localSheetId, 1);
  const deleted = patchDeleteSheet(moved, 0).model;
  assert.equal(deleted.sheets[0].printAreas.length, 1); assert.equal(deleted.definedNames[0].localSheetId, 0);
});

test('기존 Excel의 여러 인쇄 영역과 쉼표·작은따옴표 시트 이름을 읽는다', () => {
  let model = patchSheetName(analyze(fixture()), 0, "O'Brien, Q1").model;
  model = patchPrintArea(model, 0, { top: 0, bottom: 0, left: 0, right: 0 }).model;
  const files = { ...model.files }, xml = new TextDecoder().decode(files['xl/workbook.xml']);
  files['xl/workbook.xml'] = new TextEncoder().encode(xml.replace(/(<definedName\b[^>]*>)[\s\S]*?(<\/definedName>)/, (element, opening, closing) => opening + "'O''Brien, Q1'!$A$1:$B$1,'O''Brien, Q1'!$C$1:$C$1" + closing));
  const read = analyze(zipSync(files));
  assert.equal(read.writable, true, read.readOnlyReason); assert.equal(read.sheets[0].printAreas.length, 2);
  const moved = transformStructure(read, { sheetIndex: 0, axis: 'row', action: 'insert', index: 0 }).model;
  assert.deepEqual(moved.sheets[0].printAreas.map(area => area.top), [1, 1]);
});

test('정의된 이름을 통합 문서 또는 시트 범위로 만들고 수정·삭제한다', () => {
  const original = patchSheetName(analyze(fixture()), 0, "O'Brien, Q1").model;
  const global = { action: 'set', name: 'Budget', localSheetId: null, source: { sheetName: "O'Brien, Q1", top: 0, bottom: 1, left: 0, right: 0 } };
  let model = patchDefinedName(original, global).model;
  assert.deepEqual(model.definedNames, [{ name: 'Budget', localSheetId: null, source: { sheetName: "O'Brien, Q1", top: 0, bottom: 1, left: 0, right: 0 } }]);
  assert.match(new TextDecoder().decode(model.files['xl/workbook.xml']), /name="Budget">&apos;O&apos;&apos;Brien, Q1&apos;!\$A\$1:\$A\$2/);
  for (const [name, bytes] of Object.entries(original.files)) if (name !== 'xl/workbook.xml') assert.deepEqual(model.files[name], bytes);

  model = patchDefinedName(model, { ...global, localSheetId: 0, source: { sheetName: "O'Brien, Q1", top: 0, bottom: 1, left: 1, right: 1 } }).model;
  assert.equal(model.definedNames.length, 2, 'local names can shadow workbook names');
  assert.deepEqual(model.definedNames.find(item => item.localSheetId === 0).source, { sheetName: "O'Brien, Q1", top: 0, bottom: 1, left: 1, right: 1 });
  model = patchDefinedName(model, { ...global, source: { sheetName: "O'Brien, Q1", top: 1, bottom: 3, left: 0, right: 0 } }).model;
  assert.deepEqual(model.definedNames.find(item => item.localSheetId === null).source, { sheetName: "O'Brien, Q1", top: 1, bottom: 3, left: 0, right: 0 });
  model = patchDefinedName(model, { action: 'delete', name: 'Budget' }).model;
  assert.equal(model.definedNames.length, 1); assert.equal(model.definedNames[0].localSheetId, 0);
  model = patchDefinedName(model, { action: 'delete', name: 'Budget', localSheetId: 0 }).model;
  assert.deepEqual(model.definedNames, []); assert.doesNotMatch(new TextDecoder().decode(model.files['xl/workbook.xml']), /<definedNames\b/);

  assert.throws(() => patchDefinedName(original, { ...global, name: 'A1' }), /유효하지 않습니다/);
  assert.throws(() => patchDefinedName(original, { ...global, name: 'HugeRange', source: { sheetName: "O'Brien, Q1", top: 0, bottom: 1000, left: 0, right: 0 } }), /1,000셀/);
  const duplicate = patchDefinedName(original, global).model;
  assert.strictEqual(patchDefinedName(duplicate, { ...global, name: 'budget' }).model, duplicate, 'name matching follows Excel case-insensitive semantics');
});

test('정의된 이름 참조는 행·열 구조 변경에 따라 이동하고 삭제된 이름은 #REF를 보존한다', () => {
  const files = unzipSync(fixture()), encode = value => new TextEncoder().encode(value), decode = name => new TextDecoder().decode(files[name]);
  const definitions = '<definedNames><definedName name="DataRange">Data!$A$1:$B$3</definedName><definedName name="LocalRows" localSheetId="0">$A$1:$A$3</definedName><definedName name="OtherRows" localSheetId="1">$A$1:$A$1</definedName><definedName name="SingleCell">Data!$A$1</definedName></definedNames>';
  files['xl/workbook.xml'] = encode(decode('xl/workbook.xml').replace('</sheets>', '</sheets>' + definitions));
  files['xl/worksheets/sheet2.xml'] = encode(decode('xl/worksheets/sheet2.xml').replace('</sheetData>', '<row r="2"><c r="A2"><f>SUM(DataRange)</f><v>42</v></c><c r="B2"><f>SUM(SingleCell)</f><v>0</v></c></row></sheetData>'));
  const original = analyze(zipSync(files));
  assert.equal(original.writable, true, original.readOnlyReason);

  const insertedRows = transformStructure(original, { sheetIndex: 0, axis: 'row', action: 'insert', index: 1 }).model;
  assert.equal(insertedRows.writable, true, insertedRows.readOnlyReason);
  assert.deepEqual(insertedRows.definedNames.find(item => item.name === 'DataRange').source, { sheetName: 'Data', top: 0, bottom: 3, left: 0, right: 1 });
  assert.deepEqual(insertedRows.definedNames.find(item => item.name === 'LocalRows').source, { sheetName: 'Data', top: 0, bottom: 3, left: 0, right: 0 });
  assert.deepEqual(insertedRows.definedNames.find(item => item.name === 'OtherRows').source, { sheetName: 'Hidden', top: 0, bottom: 0, left: 0, right: 0 });
  assert.equal(insertedRows.sheets[1].cells.get('A2').cachedValue, '');
  assert.equal(insertedRows.sheets[1].cells.get('B2').cachedValue, '');
  const columnInsert = transformStructure(insertedRows, { sheetIndex: 0, axis: 'column', action: 'insert', index: 1 }).model;
  assert.deepEqual(columnInsert.definedNames.find(item => item.name === 'DataRange').source, { sheetName: 'Data', top: 0, bottom: 3, left: 0, right: 2 });
  const workbookXml = new TextDecoder().decode(columnInsert.files['xl/workbook.xml']);
  assert.match(workbookXml, /name="DataRange">Data!\$A\$1:\$C\$4/);
  assert.match(workbookXml, /name="LocalRows" localSheetId="0">\$A\$1:\$A\$4/);

  const deleted = transformStructure(original, { sheetIndex: 0, axis: 'row', action: 'delete', index: 0 }).model;
  assert.equal(deleted.writable, true, deleted.readOnlyReason);
  assert.equal(deleted.definedNames.find(item => item.name === 'SingleCell').error, '#REF!');
  assert.match(new TextDecoder().decode(deleted.files['xl/workbook.xml']), /name="SingleCell">#REF!<\/definedName>/);
  assert.equal(deleted.sheets[1].cells.get('B2').cachedValue, '');
  const calculator = createFormulaCalculator({
    state: { formulas: true, sheets: deleted.sheets, definedNames: deleted.definedNames, date1904: deleted.date1904 },
    getCell: (row, column, sheetIndex = 0) => deleted.sheets[sheetIndex]?.cells.get(String.fromCharCode(65 + column) + (row + 1)) || null
  });
  assert.equal(calculator.value(1, 1, 1), '#REF!');
});

test('Excel의 전체 행·열 인쇄 영역은 범위로 읽고 이름 변경·구조 변경 후에도 보존한다', () => {
  const original = analyze(fixture()), files = { ...original.files };
  files['xl/workbook.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/workbook.xml']).replace('</sheets>', '</sheets><definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A:$C,Data!$5:$8</definedName></definedNames>'));
  const model = analyze(zipSync(files));
  assert.equal(model.writable, true, model.readOnlyReason);
  assert.deepEqual(model.sheets[0].printAreas, [{ sheetName: 'Data', top: 0, bottom: 1048575, left: 0, right: 2 }, { sheetName: 'Data', top: 4, bottom: 7, left: 0, right: 16383 }]);
  assert.equal(model.definedNames[0].source, null, 'multi-area names do not masquerade as a single formula range');
  const renamed = patchSheetName(model, 0, 'Printed Data').model;
  assert.ok(renamed.sheets[0].printAreas.every(area => area.sheetName === 'Printed Data'));
  const inserted = transformStructure(renamed, { sheetIndex: 0, axis: 'row', action: 'insert', index: 0 }).model;
  assert.equal(inserted.sheets[0].printAreas[0].bottom, 1048575); assert.equal(inserted.sheets[0].printAreas[1].top, 5);
});

test('잘못된 인쇄 영역과 과도한 영역 개수는 읽기 전용으로 보존한다', () => {
  const original = analyze(fixture());
  const invalid = ['Data!A1:B1,Hidden!A1:B1', 'Data!A0:B2', 'Data!B2:A1', Array.from({ length: 17 }, () => 'Data!A1:A1').join(','), "'Data!A1:A1"];
  for (const reference of invalid) {
    const files = { ...original.files };
    files['xl/workbook.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/workbook.xml']).replace('</sheets>', '</sheets><definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">' + reference + '</definedName></definedNames>'));
    const model = analyze(zipSync(files));
    assert.equal(model.writable, false, reference); assert.throws(() => patchPrintArea(model, 0, null));
  }
  const files = { ...original.files };
  files['xl/workbook.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/workbook.xml']).replace('</sheets>', '</sheets><definedNames><definedName name="NormalRange">Data!A1:A1001</definedName></definedNames>'));
  assert.equal(analyze(zipSync(files)).writable, false, 'ordinary formula names retain their 1,000-cell range limit');
});

test('반복 제목은 행·열 이름 정의로 저장하고 인쇄 영역과 다른 부품을 보존한다', () => {
  const original = analyze(fixture()), area = patchPrintArea(original, 0, { top: 2, bottom: 20, left: 2, right: 4 }).model;
  const titles = { rows: { start: 0, end: 1 }, columns: { start: 0, end: 0 } }, configured = patchPrintTitles(area, 0, titles).model;
  assert.deepEqual(configured.sheets[0].printTitles, titles); assert.deepEqual(configured.sheets[0].printAreas, area.sheets[0].printAreas);
  assert.equal(configured.definedNames.find(name => name.printTitles).source, null);
  const xml = new TextDecoder().decode(configured.files['xl/workbook.xml']);
  assert.match(xml, /name="_xlnm.Print_Titles" localSheetId="0"/); assert.match(xml, /\$1:\$2/); assert.match(xml, /\$A:\$A/);
  for (const [name, bytes] of Object.entries(area.files)) if (name !== 'xl/workbook.xml') assert.deepEqual(configured.files[name], bytes);
  assert.equal(patchPrintTitles(configured, 0, titles).model, configured);
  const cleared = patchPrintTitles(configured, 0, { rows: null, columns: null }).model;
  assert.deepEqual(cleared.sheets[0].printTitles, { rows: null, columns: null }); assert.equal(cleared.sheets[0].printAreas.length, 1);
});

test('반복 제목은 행·열 구조와 시트 이름·순서 변경에 맞춰 이동한다', () => {
  let model = patchPrintTitles(analyze(fixture()), 0, { rows: { start: 1, end: 2 }, columns: { start: 1, end: 1 } }).model;
  model = transformStructure(model, { sheetIndex: 0, axis: 'row', action: 'insert', index: 0 }).model;
  assert.deepEqual(model.sheets[0].printTitles.rows, { start: 2, end: 3 });
  model = transformStructure(model, { sheetIndex: 0, axis: 'row', action: 'insert', index: 3 }).model;
  assert.deepEqual(model.sheets[0].printTitles.rows, { start: 2, end: 4 });
  model = transformStructure(model, { sheetIndex: 0, axis: 'column', action: 'delete', index: 1 }).model;
  assert.equal(model.sheets[0].printTitles.columns, null);
  model = patchSheetName(model, 0, "O'Brien, Titles").model;
  assert.deepEqual(model.sheets[0].printTitles.rows, { start: 2, end: 4 });
  model = patchMoveSheet(model, 0, 1).model; assert.ok(model.sheets[1].printTitles.rows); assert.equal(model.definedNames.find(name => name.printTitles).localSheetId, 1);
  model = patchDeleteSheet(model, 0).model; assert.ok(model.sheets[0].printTitles.rows); assert.equal(model.definedNames.find(name => name.printTitles).localSheetId, 0);
});

test('기존 반복 제목을 읽고 잘못된 정의는 인쇄를 보호한다', () => {
  const original = analyze(fixture());
  for (const [reference, supported] of [['Data!$1:$2,Data!$A:$B', true], ['Data!$A:$B', true], ['Data!$A$1:$B$2', false], ['Data!$1:$2,Data!$3:$4', false], ['Hidden!$1:$2', false]]) {
    const files = { ...original.files };
    files['xl/workbook.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/workbook.xml']).replace('</sheets>', '</sheets><definedNames><definedName name="_xlnm.Print_Titles" localSheetId="0">' + reference + '</definedName></definedNames>'));
    const model = analyze(zipSync(files)); assert.equal(model.writable, supported, reference); assert.equal(model.sheets[0].printTitlesUnsupported, !supported);
  }
  for (const titles of [{ rows: { start: 0, end: 50 }, columns: null }, { rows: null, columns: { start: 0, end: 10 } }, { rows: { start: -1, end: 1 }, columns: null }, { rows: null }]) assert.throws(() => patchPrintTitles(original, 0, titles));
});

test('머리글·바닥글은 구역·첫 쪽·짝수 쪽·배율 설정을 저장하고 다른 부품을 보존한다', () => {
  const model = analyze(fixture()), config = JSON.parse(JSON.stringify(model.sheets[0].headerFooter));
  config.odd.header = { left: 'Report && data', center: '&BTitle&B', right: '&A' }; config.odd.footer.center = '&P / &N';
  config.first.header.center = 'First'; config.even.header.right = 'Even'; config.differentFirst = true; config.differentOddEven = true; config.scaleWithDoc = false;
  const configured = patchHeaderFooter(model, 0, config).model;
  assert.deepEqual(configured.sheets[0].headerFooter, config); assert.equal(configured.sheets[0].headerFooterUnsupported, false);
  assert.match(configured.sheets[0].xml, /xml:space="preserve"/); assert.match(configured.sheets[0].xml, /&amp;P \/ &amp;N/);
  for (const [name, bytes] of Object.entries(model.files)) if (name !== model.sheets[0].path) assert.deepEqual(configured.files[name], bytes);
  assert.equal(patchHeaderFooter(configured, 0, config).model, configured);
  const settings = { ...configured.sheets[0].printSettings, firstPageNumber: 10, useFirstPageNumber: true };
  const numbered = patchPrintSettings(configured, 0, settings).model;
  assert.deepEqual(numbered.sheets[0].headerFooter, config); assert.equal(numbered.sheets[0].printSettings.firstPageNumber, 10);
  assert.ok(numbered.sheets[0].xml.indexOf('<pageSetup') < numbered.sheets[0].xml.indexOf('<headerFooter'));
  assert.deepEqual(patchHeaderFooter(numbered, 0, config, true).model.sheets[0].headerFooter, model.sheets[0].headerFooter);
});

test('지원하지 않는 머리글은 셀 편집에서 보존하고 명시적으로 지운 뒤 새로 설정한다', () => {
  const model = analyze(fixture({ autoFilter: '<headerFooter><oddHeader>&amp;G</oddHeader></headerFooter>' }));
  assert.equal(model.writable, true); assert.equal(model.sheets[0].headerFooterUnsupported, true);
  const edited = patchCell(model, 0, 'A1', 'Edited').model;
  assert.match(edited.sheets[0].xml, /<oddHeader>&amp;G<\/oddHeader>/);
  assert.throws(() => patchHeaderFooter(edited, 0, edited.sheets[0].headerFooter), /지운 뒤/);
  const cleared = patchHeaderFooter(edited, 0, null, true).model;
  assert.equal(cleared.sheets[0].headerFooterUnsupported, false); assert.equal(cleared.sheets[0].cells.get('A1').value, 'Edited');
});

test('빈 머리글의 속성 순서와 시작 번호 0은 불필요한 변경 없이 처리한다', () => {
  const model = analyze(fixture()), config = model.sheets[0].headerFooter;
  const reordered = { odd: config.odd, even: config.even, first: config.first, alignWithMargins: true, scaleWithDoc: true, differentOddEven: false, differentFirst: false };
  assert.equal(patchHeaderFooter(model, 0, reordered).model, model);
  const numbered = patchPrintSettings(model, 0, { ...model.sheets[0].printSettings, useFirstPageNumber: true, firstPageNumber: 0 }).model;
  assert.equal(numbered.sheets[0].printSettings.firstPageNumber, 0);
});

test('수동 페이지 나누기는 행·열 경계와 개수를 저장하고 다른 부품을 보존한다', () => {
  const original = analyze(fixture());
  let model = patchPageBreaks(original, 0, { axis: 'row', action: 'add', index: 3 }).model;
  model = patchPageBreaks(model, 0, { axis: 'column', action: 'add', index: 2 }).model;
  assert.deepEqual(model.sheets[0].pageBreaks, { rows: [3], columns: [2], hasAny: true }); assert.equal(model.sheets[0].pageBreaksUnsupported, false);
  assert.match(model.sheets[0].xml, /<rowBreaks count="1" manualBreakCount="1"><brk id="3" min="0" max="16383" man="1"\/>/);
  assert.match(model.sheets[0].xml, /<colBreaks count="1" manualBreakCount="1"><brk id="2" min="0" max="1048575" man="1"\/>/);
  assert.ok(model.sheets[0].xml.indexOf('<rowBreaks') < model.sheets[0].xml.indexOf('<colBreaks'));
  for (const [name, bytes] of Object.entries(original.files)) if (name !== original.sheets[0].path) assert.deepEqual(model.files[name], bytes);
  assert.equal(patchPageBreaks(model, 0, { axis: 'row', action: 'add', index: 3 }).model, model);
  model = patchPageBreaks(model, 0, { axis: 'row', action: 'remove', index: 3 }).model; assert.deepEqual(model.sheets[0].pageBreaks.rows, []);
  model = patchPageBreaks(model, 0, { action: 'reset' }).model; assert.deepEqual(model.sheets[0].pageBreaks, { rows: [], columns: [], hasAny: false });
  assert.throws(() => patchPageBreaks(model, 0, { axis: 'row', action: 'add', index: 0 }), /첫 행/);
});

test('자동 나누기 정보는 보존하고 같은 경계를 수동으로 승격한다', () => {
  const original = analyze(fixture({ autoFilter: '<rowBreaks count="2" manualBreakCount="0"><brk id="3" min="0" max="16383"/><brk id="5" min="1" max="2" man="0"/></rowBreaks>' }));
  assert.equal(original.sheets[0].pageBreaksUnsupported, false); assert.deepEqual(original.sheets[0].pageBreaks.rows, []);
  const model = patchPageBreaks(original, 0, { axis: 'row', action: 'add', index: 3 }).model;
  assert.deepEqual(model.sheets[0].pageBreaks.rows, [3]); assert.match(model.sheets[0].xml, /count="2" manualBreakCount="1"/);
  assert.match(model.sheets[0].xml, /<brk id="5" min="1" max="2" man="0"\/>/);
  const removed = patchPageBreaks(model, 0, { axis: 'row', action: 'remove', index: 3 }).model;
  assert.match(removed.sheets[0].xml, /<brk id="5" min="1" max="2" man="0"\/>/); assert.deepEqual(removed.sheets[0].pageBreaks.rows, []);
});

test('나누기 참조는 구조 변경 후 이동하고 삭제로 겹친 경계는 합쳐진다', () => {
  let model = patchPageBreaks(analyze(fixture()), 0, { axis: 'row', action: 'add', index: 3 }).model;
  model = patchPageBreaks(model, 0, { axis: 'row', action: 'add', index: 4 }).model;
  model = patchPageBreaks(model, 0, { axis: 'column', action: 'add', index: 2 }).model;
  model = transformStructure(model, { sheetIndex: 0, axis: 'row', action: 'insert', index: 3 }).model;
  assert.deepEqual(model.sheets[0].pageBreaks.rows, [4, 5]); assert.deepEqual(model.sheets[0].pageBreaks.columns, [2]);
  model = transformStructure(model, { sheetIndex: 0, axis: 'row', action: 'delete', index: 4 }).model;
  assert.deepEqual(model.sheets[0].pageBreaks.rows, [4]); assert.equal(model.sheets[0].pageBreaksUnsupported, false);
  model = transformStructure(model, { sheetIndex: 0, axis: 'column', action: 'insert', index: 0 }).model; assert.deepEqual(model.sheets[0].pageBreaks.columns, [3]);
});

test('부분·손상된 나누기는 셀 편집에서 보존하고 전체 초기화로만 제거한다', () => {
  for (const raw of ['<rowBreaks count="1" manualBreakCount="1"><brk id="3" min="1" max="2" man="1"/></rowBreaks>', '<rowBreaks count="2"><brk id="3" min="0" max="16383" man="1"/></rowBreaks>', '<rowBreaks/><rowBreaks/>']) {
    const model = analyze(fixture({ autoFilter: raw }));
    assert.equal(model.writable, true); assert.equal(model.sheets[0].pageBreaksUnsupported, true);
    const edited = patchCell(model, 0, 'A1', 'Changed').model; assert.ok(edited.sheets[0].xml.includes(raw));
    assert.throws(() => patchPageBreaks(edited, 0, { axis: 'row', action: 'add', index: 4 }), /지운 뒤/);
    assert.throws(() => transformStructure(edited, { sheetIndex: 0, axis: 'row', action: 'insert', index: 0 }), /지운 뒤/);
    const reset = patchPageBreaks(edited, 0, { action: 'reset' }).model; assert.equal(reset.sheets[0].pageBreaksUnsupported, false); assert.equal(reset.sheets[0].cells.get('A1').value, 'Changed');
  }
});

test('병합 셀을 가르는 새 나누기는 거부한다', () => {
  const model = analyze(fixture({ autoFilter: '<mergeCells count="1"><mergeCell ref="A1:B2"/></mergeCells>' }));
  assert.throws(() => patchPageBreaks(model, 0, { axis: 'row', action: 'add', index: 1 }), /병합/);
  assert.throws(() => patchPageBreaks(model, 0, { axis: 'column', action: 'add', index: 1 }), /병합/);
  assert.deepEqual(patchPageBreaks(model, 0, { axis: 'row', action: 'add', index: 2 }).model.sheets[0].pageBreaks.rows, [2]);
});

test('구조 변경은 자동 나누기의 부분 범위를 이동하고 충돌 시 수동 경계를 우선한다', () => {
  const cached = analyze(fixture({ autoFilter: '<rowBreaks count="1" manualBreakCount="0"><brk id="4" min="1" max="2" man="0"/></rowBreaks>' }));
  let moved = transformStructure(cached, { sheetIndex: 0, axis: 'column', action: 'insert', index: 1 }).model;
  assert.match(moved.sheets[0].xml, /<brk id="4" min="2" max="3" man="0"\/>/); assert.equal(moved.sheets[0].pageBreaksUnsupported, false);
  moved = transformStructure(moved, { sheetIndex: 0, axis: 'column', action: 'delete', index: 2 }).model;
  assert.match(moved.sheets[0].xml, /<brk id="4" min="2" max="2" man="0"\/>/);
  moved = transformStructure(moved, { sheetIndex: 0, axis: 'column', action: 'delete', index: 2 }).model; assert.equal(moved.sheets[0].pageBreaks.hasAny, false);
  const collision = analyze(fixture({ autoFilter: '<rowBreaks count="2" manualBreakCount="1"><brk id="3" min="1" max="2" man="0"/><brk id="4" min="0" max="16383" man="1"/></rowBreaks>' }));
  const merged = transformStructure(collision, { sheetIndex: 0, axis: 'row', action: 'delete', index: 3 }).model;
  assert.deepEqual(merged.sheets[0].pageBreaks.rows, [3]); assert.match(merged.sheets[0].xml, /count="1" manualBreakCount="1"/); assert.equal(merged.sheets[0].pageBreaksUnsupported, false);
});


test('필터 행 표시는 native 숨김으로 저장하고 수동 숨김의 겹침을 재개방 후 보존한다', () => {
  const files = unzipSync(fixture());
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Header</t></is></c></row><row r="2" hidden="1"><c r="A2"><v>5</v></c></row><row r="3"><c r="A3"><f>5+10</f><v>0</v></c></row><row r="4"><c r="A4"><v>25</v></c></row><row r="5" hidden="1"><c r="A5"><v>18</v></c></row></sheetData></worksheet>');
  const original = analyze(zipSync(files));
  const filter = { top: 0, bottom: 5, left: 0, right: 0, columns: [{ column: 0, custom: { and: true, conditions: [{ operator: 'greaterThanOrEqual', value: '10' }, { operator: 'lessThan', value: '20' }] } }] };
  const patched = patchAutoFilters(original, [{ sheetIndex: 0, filter }]);
  let model = analyze(patched.bytes);
  assert.deepEqual(model.sheets[0].hiddenRows, [1, 4]);
  assert.deepEqual(model.sheets[0].nativeHiddenRows, [1, 3, 4, 5], 'formula is evaluated; sparse rejected blank row is materialized');
  assert.match(model.sheets[0].xml, /filterMode="1"/);
  assert.match(model.sheets[0].xml, /mc:Ignorable="jsv"/);
  assert.match(model.sheets[0].xml, /mc:PreserveAttributes="jsv:version jsv:manualHidden"/);
  assert.equal(Buffer.compare(Buffer.from(patchFilterVisibility(model).bytes), Buffer.from(model.bytes)), 0, 'visibility serialization is idempotent');
  model = patchCells(model, [{ sheetIndex: 0, address: 'A4', value: '12', kind: 'number' }]).model;
  model = patchFilterVisibility(model).model;
  assert.deepEqual(model.sheets[0].nativeHiddenRows, [1, 4, 5], 'changed cell causes previously filtered row to reappear');
  model = patchHiddenRows(model, [{ sheetIndex: 0, rows: [2], hidden: true }]).model;
  assert.deepEqual(model.sheets[0].hiddenRows, [1, 2, 4]);
  model = patchAutoFilters(analyze(model.bytes), [{ sheetIndex: 0, filter: null }]).model;
  assert.deepEqual(model.sheets[0].nativeHiddenRows, [1, 2, 4]);
  assert.deepEqual(model.sheets[0].hiddenRows, [1, 2, 4]);
  assert.match(model.sheets[0].xml, /filterMode="0"/);
  assert.doesNotMatch(model.sheets[0].xml, /<row r="6"/);
  assert.deepEqual(model.sheets[0].cells.get('A3').formula, '5+10');
  assert.deepEqual(model.files['xl/worksheets/sheet2.xml'], original.files['xl/worksheets/sheet2.xml']);
});

test('외부 native 필터 숨김은 필터 해제 때 복구하고 조건을 통과하는 수동 숨김은 유지한다', () => {
  const files = unzipSync(fixture());
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr filterMode="1"/><sheetData><row r="1"><c r="A1"><v>0</v></c></row><row r="2" hidden="1"><c r="A2"><v>5</v></c></row><row r="3" hidden="1"><c r="A3"><v>15</v></c></row></sheetData><autoFilter ref="A1:A3"><filterColumn colId="0"><customFilters><customFilter operator="greaterThan" val="10"/></customFilters></filterColumn></autoFilter></worksheet>');
  const model = analyze(zipSync(files));
  assert.deepEqual(model.sheets[0].hiddenRows, [2]);
  const edited = patchFilterVisibility(patchCells(model, [{ sheetIndex: 0, address: 'A2', value: '20', kind: 'number' }]).model).model;
  assert.deepEqual(edited.sheets[0].hiddenRows, [2]); assert.deepEqual(edited.sheets[0].nativeHiddenRows, [2], 'editing an imported excluded row to pass must not turn it into a manually hidden row');
  const manuallyHidden = patchHiddenRows(model, [{ sheetIndex: 0, rows: [1], hidden: true }]).model;
  assert.deepEqual(manuallyHidden.sheets[0].hiddenRows, [1, 2], 'manual hide on an already excluded native row is preserved');
  const changed = patchAutoFilters(model, [{ sheetIndex: 0, filter: { ...model.sheets[0].autoFilter, columns: [{ column: 0, values: ['5'], blanks: false }] } }]).model;
  assert.deepEqual(changed.sheets[0].hiddenRows, [2], 'changing criteria must not reclassify old filter-hidden rows as manual');
  assert.deepEqual(changed.sheets[0].nativeHiddenRows, [2]);
  const cleared = patchAutoFilters(analyze(changed.bytes), [{ sheetIndex: 0, filter: null }]).model;
  assert.deepEqual(cleared.sheets[0].nativeHiddenRows, [2]);
});

test('필터 숨김 메타데이터는 네임스페이스 충돌을 피하고 행 삽입·삭제에 맞춰 이동한다', () => {
  const files = unzipSync(fixture({ formula: true }));
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('<worksheet ', '<worksheet xmlns:jsv="urn:other" xmlns:compat="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:other="urn:another" compat:Ignorable="other" '));
  let model = patchHiddenRows(analyze(zipSync(files)), [{ sheetIndex: 0, rows: [1], hidden: true }]).model;
  const filter = { top: 0, bottom: 2, left: 0, right: 0, columns: [{ column: 0, values: [], blanks: false }] };
  model = patchAutoFilters(model, [{ sheetIndex: 0, filter }]).model;
  assert.match(model.sheets[0].xml, /xmlns:jsv1="https:\/\/jinsheet.dev\/ooxml\/row-visibility\/1"/);
  assert.match(model.sheets[0].xml, /compat:Ignorable="other jsv1"/);
  assert.deepEqual(model.sheets[0].hiddenRows, [1]);
  model = transformStructure(model, { sheetIndex: 0, axis: 'row', action: 'insert', index: 1 }).model;
  assert.deepEqual(model.sheets[0].hiddenRows, [2]);
  model = patchAutoFilters(model, [{ sheetIndex: 0, filter: null }]).model;
  assert.deepEqual(model.sheets[0].nativeHiddenRows, [2]);
});

test('필터 native 행 숨김은 고정 크기 차트의 저장 사각형을 유지한다', () => {
  const original = analyze(fixtureWithPositionedChart('oneCell'));
  const before = chartPrintGeometry(original.sheets[0], original.sheets[0].chartAnchors).boxes[0];
  const filter = { top: 0, bottom: 5, left: 0, right: 0, columns: [{ column: 0, values: [], blanks: false }] };
  const model = patchAutoFilters(original, [{ sheetIndex: 0, filter }]).model;
  const sheet = model.sheets[0], after = chartPrintGeometry({ ...sheet, hiddenRows: sheet.nativeHiddenRows }, sheet.chartAnchors).boxes[0];
  assert.ok(Math.abs(before.width - after.width) < .001); assert.ok(Math.abs(before.height - after.height) < .001);
  const repositioned = patchChartPlacement(model, 0, 0, 'absolute').model.sheets[0];
  const moved = chartPrintGeometry({ ...repositioned, hiddenRows: repositioned.nativeHiddenRows }, repositioned.chartAnchors).boxes[0];
  assert.ok(Math.abs(after.width - moved.width) < .001); assert.ok(Math.abs(after.height - moved.height) < .001);

  const cleared = patchAutoFilters(model, [{ sheetIndex: 0, filter: null }]).model;
  const restored = chartPrintGeometry(cleared.sheets[0], cleared.sheets[0].chartAnchors).boxes[0];
  assert.ok(Math.abs(before.height - restored.height) < .001);
});

function totalsFilterFixture(filterReference = 'A1:B3') {
  const files = unzipSync(fixture());
  files['[Content_Types].xml'] = new TextEncoder().encode(new TextDecoder().decode(files['[Content_Types].xml']).replace('</Types>', '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>'));
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c><c r="C1" t="inlineStr"><is><t>Extra</t></is></c></row><row r="2"><c r="B2"><v>10</v></c></row><row r="3"><c r="B3"><v>20</v></c></row><row r="4"><c r="A4" t="inlineStr"><is><t>Total</t></is></c><c r="B4"><f>SUBTOTAL(109,[Amount])</f><v>30</v></c></row></sheetData><tableParts count="1"><tablePart r:id="t1"/></tableParts></worksheet>');
  files['xl/worksheets/_rels/sheet1.xml.rels'] = new TextEncoder().encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="t1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>');
  files['xl/tables/table1.xml'] = new TextEncoder().encode('<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:B4" totalsRowCount="1"><autoFilter ref="' + filterReference + '"/><tableColumns count="2"><tableColumn id="1" name="Name" totalsRowLabel="Total"/><tableColumn id="2" name="Amount" totalsRowFunction="sum"/></tableColumns></table>');
  return zipSync(files);
}

test('총계 표의 필터 범위는 행·열 구조 변경과 열 크기 조정 후에도 총계를 제외한다', () => {
  const original = analyze(totalsFilterFixture()); assert.equal(original.writable, true, original.readOnlyReason);
  const table = original.sheets[0].tables[0], filter = { ...table.autoFilter, columns: [{ column: 1, values: ['20'], blanks: false }] };
  let model = patchAutoFilters(original, [{ sheetIndex: 0, filter }]).model;
  assert.match(new TextDecoder().decode(model.files['xl/workbook.xml']), /fullCalcOnLoad="1"/);
  model = transformStructure(model, { sheetIndex: 0, axis: 'row', action: 'insert', index: 3 }).model;
  assert.equal(model.sheets[0].tables[0].range.bottom, 4); assert.equal(model.sheets[0].tables[0].autoFilter.bottom, 3);
  assert.match(new TextDecoder().decode(model.files[table.path]), /<autoFilter ref="A1:B4"/);
  assert.equal(model.sheets[0].cells.get('B5').formula, 'SUBTOTAL(109,[Amount])');
  model = patchFilterVisibility(model).model; assert.deepEqual(model.sheets[0].nativeHiddenRows, [1, 3]);
  model = transformStructure(model, { sheetIndex: 0, axis: 'row', action: 'delete', index: 3 }).model;
  assert.equal(model.sheets[0].tables[0].range.bottom, 3); assert.equal(model.sheets[0].tables[0].autoFilter.bottom, 2);
  model = transformStructure(model, { sheetIndex: 0, axis: 'column', action: 'insert', index: 1 }).model;
  assert.equal(model.sheets[0].tables[0].autoFilter.columns[0].column, 2); assert.equal(model.sheets[0].tables[0].autoFilter.bottom, 2);
  assert.match(new TextDecoder().decode(model.files[table.path]), /<autoFilter ref="A1:C3"/);
  const resized = patchTableRanges(original, [{ sheetIndex: 0, tableName: 'Sales', range: { top: 0, bottom: 3, left: 0, right: 2 } }]).model;
  assert.equal(resized.sheets[0].tables[0].autoFilter.bottom, 2); assert.equal(resized.sheets[0].tables[0].autoFilter.right, 2);
  assert.match(new TextDecoder().decode(resized.files[table.path]), /<autoFilter ref="A1:C3"/);
  assert.match(new TextDecoder().decode(resized.files[table.path]), /totalsRowFunction="sum"/);
  assert.equal(patchTableRanges(original, [{ sheetIndex: 0, tableName: 'Sales', range: { top: 0, bottom: 5, left: 0, right: 1 } }]).model.sheets[0].cells.get('A6').value, 'Total');
  assert.throws(() => transformStructure(original, { sheetIndex: 0, axis: 'row', action: 'delete', index: 3 }), /머리글·합계 행 삭제/);
  assert.throws(() => patchAutoFilters(original, [{ sheetIndex: 0, filter: { ...filter, bottom: 3 } }]), /머리글과 데이터 행/);
});

test('총계 표의 잘못된 필터 범위는 읽기 전용이며 이전 전체 범위는 정규화한다', () => {
  for (const ref of ['A1:B2', 'A2:B3', 'A1:C3', 'A1:B5']) assert.equal(analyze(totalsFilterFixture(ref)).writable, false, ref);
  const legacy = analyze(totalsFilterFixture('A1:B4')); assert.equal(legacy.writable, true); assert.equal(legacy.sheets[0].tables[0].autoFilter.bottom, 2);
  const normalized = patchFilterVisibility(legacy).model; assert.equal(normalized.sheets[0].tables[0].filterRangeNeedsNormalization, false);
  assert.match(new TextDecoder().decode(normalized.files['xl/tables/table1.xml']), /<table[^>]*ref="A1:B4"/);
  assert.match(new TextDecoder().decode(normalized.files['xl/tables/table1.xml']), /<autoFilter ref="A1:B3"/);
});

test('총계 표 세로 크기 조정은 총계 셀·서식과 직접 참조·이름 정의를 함께 이동한다', () => {
  const files = unzipSync(totalsFilterFixture());
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</sheetData>', '<row r="5"><c r="B5"><v>7</v></c></row></sheetData>').replace('</row><row r="5">', '<c r="C4"><f>B4+$B$4</f><v>60</v></c><c r="D4"><f>SUM(A4:B4)</f><v>30</v></c><c r="E4"><f>SUM(B2:B4)</f><v>60</v></c><c r="F4"><f>"B4"</f><v>0</v></c></row><row r="5">'));
  files['xl/worksheets/sheet2.xml'] = new TextEncoder().encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>Data!B4</f><v>30</v></c><c r="B1"><f>TotalAmount</f><v>30</v></c></row></sheetData></worksheet>');
  files['xl/workbook.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/workbook.xml']).replace('</workbook>', '<definedNames><definedName name="TotalAmount">Data!$B$4</definedName></definedNames></workbook>'));
  const original = patchCells(analyze(zipSync(files)), [{ sheetIndex: 0, address: 'B4', value: '30', formula: 'SUBTOTAL(109,[Amount])', font: { bold: true, italic: false } }, { sheetIndex: 0, address: 'B3', value: '20', font: { bold: false, italic: true } }]).model, grown = patchTableRanges(original, [{ sheetIndex: 0, tableName: 'Sales', range: { top: 0, bottom: 5, left: 0, right: 1 } }]).model;
  const sheet = grown.sheets[0]; assert.equal(sheet.cells.get('A6').value, 'Total'); assert.equal(sheet.cells.get('B6').formula, 'SUBTOTAL(109,[Amount])'); assert.equal(sheet.cells.get('B4')?.value ?? '', '');
  assert.equal(sheet.cells.get('B6').styleProperties.bold, true); assert.equal(sheet.cells.get('B4').styleProperties.italic, true);
  assert.equal(sheet.cells.get('C4').formula, 'B6+$B$6'); assert.equal(sheet.cells.get('D4').formula, 'SUM(A6:B6)'); assert.equal(sheet.cells.get('E4').formula, 'SUM(B2:B4)'); assert.equal(sheet.cells.get('F4').formula, '"B4"');
  assert.equal(grown.sheets[1].cells.get('A1').formula, 'Data!B6'); assert.equal(grown.definedNames[0].source.top, 5);
  const calculator = createFormulaCalculator({ state: grown, getCell: (row, column, index = 0) => [...grown.sheets[index].cells.values()].find(cell => cell.row === row && cell.column === column) || null });
  assert.equal(calculator.value(0, 5, 1), 37); assert.equal(calculator.value(0, 3, 2), 74); assert.equal(calculator.value(1, 0, 1), 37);
  assert.equal(calculator.value(0, 3, 3), 37); assert.equal(calculator.value(0, 3, 4), 30); assert.equal(calculator.value(0, 3, 5), 'B4'); assert.deepEqual(grown.files['xl/styles.xml'], original.files['xl/styles.xml']);
  const shrunk = patchTableRanges(grown, [{ sheetIndex: 0, tableName: 'Sales', range: { top: 0, bottom: 3, left: 0, right: 1 } }]).model;
  assert.equal(shrunk.sheets[0].cells.get('B4').formula, 'SUBTOTAL(109,[Amount])'); assert.equal(shrunk.sheets[0].cells.get('B6')?.value ?? '', ''); assert.equal(shrunk.sheets[0].cells.get('B5').value, '7');
  assert.equal(shrunk.sheets[0].cells.get('C4').formula, 'B4+$B$4'); assert.equal(shrunk.definedNames[0].source.top, 3);
  assert.throws(() => patchTableRanges(original, [{ sheetIndex: 0, tableName: 'Sales', range: { top: 0, bottom: 2, left: 0, right: 1 } }]), /이동할 합계 행 위치에 데이터/);
});

test('합계 이동 참조는 한글 시트·작은따옴표·공백 범위와 부속 수식을 보존한다', () => {
  let model = patchSheetName(analyze(totalsFilterFixture()), 0, '데이터').model;
  let files = { ...model.files }, xml = model.sheets[0].xml;
  xml = xml.replace('</sheetData>', '<row r="7"><c r="C7"><f>SUM(B4 : B5)</f><v>30</v></c><c r="D7"><f>SUM(A4 : B4)</f><v>30</v></c></row></sheetData>').replace('</worksheet>', '<dataValidations count="1"><dataValidation type="list" sqref="C6"><formula1>데이터!$B$4</formula1></dataValidation></dataValidations><hyperlinks><hyperlink ref="D6" location="데이터!B4"/></hyperlinks></worksheet>');
  files[model.sheets[0].path] = new TextEncoder().encode(xml);
  files[model.sheets[1].path] = new TextEncoder().encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>데이터!B4+\'데이터\'!$B$4</f><v>60</v></c></row></sheetData></worksheet>');
  files['xl/tables/table1.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/tables/table1.xml']).replace('<tableColumn id="1" name="Name" totalsRowLabel="Total"/>', '<tableColumn id="1" name="Name" totalsRowLabel="Total"><calculatedColumnFormula>$B$4</calculatedColumnFormula></tableColumn>'));
  model = analyze(zipSync(files)); assert.equal(model.writable, true, model.readOnlyReason);
  model = patchTableRanges(model, [{ sheetIndex: 0, tableName: 'Sales', range: { top: 0, bottom: 5, left: 0, right: 1 } }]).model;
  assert.equal(model.sheets[0].cells.get('C7').formula, 'SUM(B4 : B5)'); assert.equal(model.sheets[0].cells.get('D7').formula, 'SUM(A6:B6)');
  assert.equal(model.sheets[1].cells.get('A1').formula, '데이터!B6+\'데이터\'!$B$6');
  assert.equal(model.sheets[0].dataValidations[0].source.top, 5); assert.equal(model.sheets[0].hyperlinks[0].location, '데이터!B6');
  assert.match(new TextDecoder().decode(model.files['xl/tables/table1.xml']), /<calculatedColumnFormula>\$B\$6<\/calculatedColumnFormula>/);
  const calc = createFormulaCalculator({ state: model, getCell: (row, column, index = 0) => [...model.sheets[index].cells.values()].find(cell => cell.row === row && cell.column === column) || null }); assert.equal(calc.value(1, 0, 0), 60);
});

test('합계 이동은 목적지 수식·부속 범위·3차원 참조에서 원본을 보존하며 거부한다', () => {
  for (const feature of ['formula', 'merge', 'validation', '3d']) {
    const files = unzipSync(totalsFilterFixture()); let xml = new TextDecoder().decode(files['xl/worksheets/sheet1.xml']);
    if (feature === 'formula') xml = xml.replace('</sheetData>', '<row r="6"><c r="B6"><f>""</f><v></v></c></row></sheetData>');
    if (feature === 'merge') xml = xml.replace('</worksheet>', '<mergeCells count="1"><mergeCell ref="A6:B6"/></mergeCells></worksheet>');
    if (feature === 'validation') xml = xml.replace('</worksheet>', '<dataValidations count="1"><dataValidation type="whole" operator="between" sqref="B6"><formula1>1</formula1><formula2>100</formula2></dataValidation></dataValidations></worksheet>');
    if (feature === '3d') xml = xml.replace('</sheetData>', '<row r="7"><c r="C7"><f>SUM(Data:Hidden!B4)</f><v>0</v></c></row></sheetData>');
    files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(xml);
    const model = analyze(zipSync(files)); assert.equal(model.writable, true, model.readOnlyReason); const original = Buffer.from(model.bytes);
    assert.throws(() => patchTableRanges(model, [{ sheetIndex: 0, tableName: 'Sales', range: { top: 0, bottom: 5, left: 0, right: 1 } }]), feature === 'formula' ? /이동할 합계 행 위치에 데이터/ : feature === '3d' ? /3차원/ : /병합·유효성/);
    assert.deepEqual(Buffer.from(model.bytes), original); assert.equal(model.sheets[0].cells.get('B4').formula, 'SUBTOTAL(109,[Amount])');
  }
  const model = patchCells(analyze(totalsFilterFixture()), [{ sheetIndex: 0, address: 'C7', value: '', formula: '"Data:Hidden!B4"' }, { sheetIndex: 0, address: 'D7', value: 'plain', formula: null, kind: 'string' }]).model;
  const moved = patchTableRanges(model, [{ sheetIndex: 0, tableName: 'Sales', range: { top: 0, bottom: 5, left: 0, right: 1 } }]).model;
  assert.equal(moved.sheets[0].cells.get('C7').formula, '"Data:Hidden!B4"'); assert.equal(moved.sheets[0].cells.get('D7').value, 'plain');
});


test('합계 이동은 가상 spill 목적지와 공유·배열 수식 레코드를 보호한다', () => {
  const plain = unzipSync(totalsFilterFixture()); plain['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(new TextDecoder().decode(plain['xl/worksheets/sheet1.xml']).replace('<f>SUBTOTAL', '<f t="normal">SUBTOTAL'));
  assert.equal(patchTableRanges(analyze(zipSync(plain)), [{ sheetIndex: 0, tableName: 'Sales', range: { top: 0, bottom: 5, left: 0, right: 1 } }]).model.sheets[0].cells.get('B6').formula, 'SUBTOTAL(109,[Amount])');

  for (const array of [false, true]) {
    const files = unzipSync(totalsFilterFixture());
    const row = array ? '<row r="7"><c r="D7"><f t="array" ref="D7:D8">SEQUENCE(2)</f><v>1</v></c></row>' : '<row r="5"><c r="B5"><f>SEQUENCE(2)</f><v>1</v></c></row>';
    files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</sheetData>', row + '</sheetData>'));
    const model = analyze(zipSync(files)); assert.equal(model.writable, true, model.readOnlyReason);
    assert.throws(() => patchTableRanges(model, [{ sheetIndex: 0, tableName: 'Sales', range: { top: 0, bottom: 5, left: 0, right: 1 } }]), array ? /공유·배열/ : /이동할 합계 행 위치에 데이터/);
  }
});

function basicTotalsControlFixture() {
  const files = unzipSync(totalsFilterFixture());
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace(/<row r="4">[\s\S]*?<\/row>/, ''));
  files['xl/tables/table1.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/tables/table1.xml']).replace('ref="A1:B4" totalsRowCount="1"', 'ref="A1:B3" totalsRowCount="0"').replace(' totalsRowLabel="Total"', '').replace(' totalsRowFunction="sum"', ''));
  return zipSync(files);
}

test('표 총계 제어는 기본 생성·집계 선택·None·숨김과 재표시를 native 메타데이터와 저장한다', () => {
  const { patchTableTotals } = require('../src/xlsx'); const original = analyze(basicTotalsControlFixture());
  let model = patchTableTotals(original, { sheetIndex: 0, tableName: 'Sales', enabled: true, label: '합계' }).model;
  const total = () => createFormulaCalculator({ state: model, getCell: (row, column) => [...model.sheets[0].cells.values()].find(cell => cell.row === row && cell.column === column) || null }).value(0, model.sheets[0].tables[0].range.bottom, 1);
  assert.equal(model.sheets[0].cells.get('A4').value, '합계'); assert.equal(total(), 30); assert.equal(model.sheets[0].tables[0].autoFilter.bottom, 2);
  for (const [fn, expected] of [['average', 15], ['count', 2], ['countNums', 2], ['min', 10], ['max', 20], ['stdDev', Math.sqrt(50)], ['var', 50], ['sum', 30]]) {
    model = patchTableTotals(model, { sheetIndex: 0, tableName: 'Sales', column: 1, function: fn }).model;
    assert.ok(Math.abs(total() - expected) < 1e-10, fn); assert.equal(model.sheets[0].tables[0].totalColumns[1].function, fn);
  }
  model = patchTableTotals(model, { sheetIndex: 0, tableName: 'Sales', column: 1, function: 'none' }).model; assert.equal(total(), null);
  model = patchTableTotals(model, { sheetIndex: 0, tableName: 'Sales', enabled: false }).model; assert.equal(model.sheets[0].tables[0].range.bottom, 2); assert.equal(model.sheets[0].cells.get('A4').value, '');
  model = patchTableTotals(model, { sheetIndex: 0, tableName: 'Sales', enabled: true }).model; assert.equal(model.sheets[0].cells.get('A4').value, '합계'); assert.equal(total(), null, 'explicit None is restored rather than defaulting to Sum');
  assert.equal(model.sheets[0].autoFilter, null); assert.deepEqual(model.files['xl/worksheets/sheet2.xml'], original.files['xl/worksheets/sheet2.xml']);
  const same = patchTableTotals(model, { sheetIndex: 0, tableName: 'Sales', enabled: true }); assert.equal(same.model, model);
});

test('표 총계는 특수문자 열 이름과 사용자 수식·여러 줄 라벨을 숨김 후 복구한다', () => {
  const { patchTableTotals } = require('../src/xlsx');
  for (const name of ['#금액', '@값', '[금액]', ']B4', '[B4', "O'Brien", '값:금액', ' 값 ']) {
    const files = unzipSync(basicTotalsControlFixture()); files['xl/tables/table1.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/tables/table1.xml']).replace('name="Amount"', 'name="' + name + '"'));
    const model = patchTableTotals(analyze(zipSync(files)), { sheetIndex: 0, tableName: 'Sales', enabled: true }).model;
    const calculator = createFormulaCalculator({ state: model, getCell: (row, column) => [...model.sheets[0].cells.values()].find(cell => cell.row === row && cell.column === column) || null }); assert.equal(calculator.value(0, 3, 1), 30, name);
  }
  let model = patchCells(analyze(totalsFilterFixture()), [{ sheetIndex: 0, address: 'A4', value: '합계\n금액', kind: 'string' }, { sheetIndex: 0, address: 'B4', value: '', formula: 'SUBTOTAL(109,B2:B2)*2' }]).model;
  model = patchTableTotals(model, { sheetIndex: 0, tableName: 'Sales', enabled: false }).model;
  assert.equal(model.sheets[0].tables[0].totalColumns[1].function, 'custom'); assert.equal(model.sheets[0].tables[0].totalColumns[0].label, '합계\n금액');
  model = patchTableTotals(analyze(model.bytes), { sheetIndex: 0, tableName: 'Sales', enabled: true }).model;
  assert.equal(model.sheets[0].cells.get('B4').formula, 'SUBTOTAL(109,B2:B2)*2'); assert.equal(model.sheets[0].cells.get('A4').value, '합계\n금액');
  const inserted = transformStructure(model, { sheetIndex: 0, axis: 'column', action: 'insert', index: 1 }).model;
  assert.equal(inserted.sheets[0].tables[0].columns.length, 3); assert.equal(inserted.sheets[0].tables[0].totalColumns[2].formula, 'SUBTOTAL(109,C2:C2)*2'); assert.equal(inserted.sheets[0].cells.get('C4').formula, 'SUBTOTAL(109,C2:C2)*2');

});

test('표 총계 제어는 충돌·유효하지 않은 요청·표 범위를 보호하고 None을 사용자 라벨에 적용한다', () => {
  const { patchTableTotals } = require('../src/xlsx'); const model = analyze(basicTotalsControlFixture()), request = { sheetIndex: 0, tableName: 'Sales', enabled: true };
  for (const invalid of [{ ...request, enabled: 'yes' }, { ...request, tableName: 'missing' }, { sheetIndex: 0, tableName: 'Sales', column: 1, function: 'bad' }, { sheetIndex: 0, tableName: 'Sales', column: 2, function: 'sum' }]) assert.throws(() => patchTableTotals(model, invalid), /유효하지/);
  assert.throws(() => patchTableTotals(model, { sheetIndex: 0, tableName: 'Sales', column: 1, function: 'sum' }), /먼저 켜세요/);
  const occupied = patchCells(model, [{ sheetIndex: 0, address: 'B4', value: '42', kind: 'number' }]).model;
  assert.throws(() => patchTableTotals(occupied, request), /이동할 합계 행 위치에 데이터/); assert.equal(occupied.sheets[0].cells.get('B4').value, '42');
  let totals = patchTableTotals(model, request).model; totals = patchTableTotals(totals, { sheetIndex: 0, tableName: 'Sales', column: 0, function: 'none' }).model;
  assert.equal(totals.sheets[0].cells.get('A4').value, ''); assert.equal(totals.sheets[0].tables[0].totalColumns[0].label, null);
  const structural = transformStructure(totals, { sheetIndex: 0, axis: 'column', action: 'insert', index: 1 }).model;
  assert.equal(structural.sheets[0].tables[0].totalColumns[2].function, 'sum'); assert.equal(structural.sheets[0].tables[0].columns.length, 3);
});

test('총계의 값 개수와 숫자 개수는 텍스트·수식·필터·수동 숨김을 구분한다', () => {
  const { patchTableTotals } = require('../src/xlsx');
  let model = patchCells(analyze(basicTotalsControlFixture()), [{ sheetIndex: 0, address: 'B2', value: '10', kind: 'string' }, { sheetIndex: 0, address: 'B3', value: '', formula: '5+15' }]).model;
  model = patchTableTotals(model, { sheetIndex: 0, tableName: 'Sales', enabled: true }).model;
  const calculate = () => createFormulaCalculator({ state: model, getCell: (row, column) => [...model.sheets[0].cells.values()].find(cell => cell.row === row && cell.column === column) || null }).value(0, 3, 1);
  model = patchTableTotals(model, { sheetIndex: 0, tableName: 'Sales', column: 1, function: 'count' }).model; assert.equal(calculate(), 2);
  model = patchTableTotals(model, { sheetIndex: 0, tableName: 'Sales', column: 1, function: 'countNums' }).model; assert.equal(calculate(), 1);
  model = patchHiddenRows(model, [{ sheetIndex: 0, rows: [2], hidden: true }]).model; assert.equal(calculate(), 0);
  model = patchTableTotals(model, { sheetIndex: 0, tableName: 'Sales', column: 1, function: 'count' }).model; assert.equal(calculate(), 1);
});

test('차트가 있는 표의 총계 제어도 원본 캐시를 갱신하고 차트 배치를 보존한다', () => {
  const { patchTableTotals } = require('../src/xlsx'), files = unzipSync(fixtureWithPositionedChart('absolute')), totals = unzipSync(basicTotalsControlFixture());
  const decode = bytes => new TextDecoder().decode(bytes), encode = value => new TextEncoder().encode(value);
  files['xl/worksheets/sheet1.xml'] = encode(decode(totals['xl/worksheets/sheet1.xml']).replace('</worksheet>', '<drawing r:id="rId3"/></worksheet>'));
  files['xl/worksheets/_rels/sheet1.xml.rels'] = encode(decode(files['xl/worksheets/_rels/sheet1.xml.rels']).replace('</Relationships>', '<Relationship Id="t1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>'));
  files['xl/tables/table1.xml'] = totals['xl/tables/table1.xml']; files['[Content_Types].xml'] = encode(decode(files['[Content_Types].xml']).replace('</Types>', '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>'));
  files['xl/charts/chart1.xml'] = encode(decode(files['xl/charts/chart1.xml']).replace('Data!$B$2:$B$3', 'Data!$B$2:$B$4').replace('Data!$A$2:$A$3', 'Data!$A$2:$A$4'));
  const original = analyze(zipSync(files)); assert.equal(original.writable, true, original.readOnlyReason);
  let model = patchTableTotals(original, { sheetIndex: 0, tableName: 'Sales', enabled: true }).model;
  assert.deepEqual(model.sheets[0].charts[0].series[0].values, [10, 20, 30]);
  model = patchTableTotals(model, { sheetIndex: 0, tableName: 'Sales', column: 1, function: 'average' }).model;
  assert.deepEqual(model.sheets[0].charts[0].series[0].values, [10, 20, 15]); assert.deepEqual(model.files['xl/drawings/drawing1.xml'], original.files['xl/drawings/drawing1.xml']);
});

test('모든 열의 명시적 None은 재표시 후에도 유지되고 켜기와 함수 선택을 함께 적용한다', () => {
  const { patchTableTotals } = require('../src/xlsx');
  let model = patchTableTotals(analyze(basicTotalsControlFixture()), { sheetIndex: 0, tableName: 'Sales', enabled: true, column: 1, function: 'average' }).model;
  assert.equal(model.sheets[0].cells.get('A4').value, 'Total'); assert.match(model.sheets[0].cells.get('B4').formula, /SUBTOTAL\(101,/);
  for (const column of [0, 1]) model = patchTableTotals(model, { sheetIndex: 0, tableName: 'Sales', column, function: 'none' }).model;
  model = patchTableTotals(model, { sheetIndex: 0, tableName: 'Sales', enabled: false }).model;
  model = patchTableTotals(analyze(model.bytes), { sheetIndex: 0, tableName: 'Sales', enabled: true }).model;
  assert.equal(model.sheets[0].cells.get('A4').value, ''); assert.equal(model.sheets[0].cells.get('B4').value, ''); assert.equal(model.sheets[0].cells.get('B4').formula, null);
});

function nativeCalculatedFixture() {
  const { patchCalculatedColumns } = require('../src/xlsx'), files = unzipSync(basicTotalsControlFixture());
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace(/<sheetData>[\s\S]*?<\/sheetData>/, '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c></row><row r="2"><c r="A2"><v>5</v></c><c r="B2"><f t="normal">A2*2</f><v>10</v></c></row><row r="3"><c r="A3"><v>10</v></c><c r="B3"><f>A3*2</f><v>20</v></c></row></sheetData>'));
  return patchCalculatedColumns(analyze(zipSync(files)), [{sheetIndex:0,tableName:'Sales',column:1,formula:'A2*2'}]).bytes;
}

test('계산 열의 첫 행·중간 행 삽입과 첫 행 삭제는 기준 수식과 새 셀을 올바르게 조정한다', () => {
  const original=analyze(nativeCalculatedFixture());
  for(const index of [1,2,0]) {
    const model=transformStructure(original,{sheetIndex:0,axis:'row',action:'insert',index}).model, table=model.sheets[0].tables[0], first=table.range.top+1;
    assert.equal(table.calculatedColumns[1],'A'+(first+1)+'*2', 'canonical anchor at '+index);
    for(let row=first;row<=table.range.bottom;row++) assert.equal(model.sheets[0].cells.get('B'+(row+1)).formula,'A'+(row+1)+'*2');
    if(index>0) assert.equal(model.sheets[0].cells.get('B'+(index+1)).hasFormula,true);
  }
  const deleted=transformStructure(original,{sheetIndex:0,axis:'row',action:'delete',index:1}).model;
  assert.equal(deleted.sheets[0].tables[0].calculatedColumns[1],'A2*2'); assert.equal(deleted.sheets[0].cells.get('B2').formula,'A2*2');
  assert.doesNotMatch(new TextDecoder().decode(deleted.files['xl/tables/table1.xml']), /#REF!/);
});

test('계산 열이 있는 표 확대는 새 빈 셀만 채우고 기존 예외·서식과 총계를 보존한다', () => {
  const {patchTableTotals}=require('../src/xlsx');
  let original=patchCells(analyze(nativeCalculatedFixture()),[{sheetIndex:0,address:'B2',value:'',formula:null},{sheetIndex:0,address:'A4',value:'7',kind:'number'},{sheetIndex:0,address:'A5',value:'9',kind:'number'},{sheetIndex:0,address:'B5',value:'999',kind:'number'},{sheetIndex:0,address:'B3',value:'20',formula:'A3*2',font:{bold:true,italic:false}}]).model;
  let grown=patchTableRanges(original,[{sheetIndex:0,tableName:'Sales',range:{top:0,bottom:4,left:0,right:1}}]).model;
  assert.equal(grown.sheets[0].cells.get('B2').formula,null); assert.equal(grown.sheets[0].cells.get('B4').formula,'A4*2'); assert.equal(grown.sheets[0].cells.get('B4').styleProperties.bold,true); assert.equal(grown.sheets[0].cells.get('B5').value,'999');
  let model=patchTableTotals(analyze(nativeCalculatedFixture()),{sheetIndex:0,tableName:'Sales',enabled:true}).model;
  model=patchTableRanges(model,[{sheetIndex:0,tableName:'Sales',range:{top:0,bottom:4,left:0,right:1}}]).model;
  assert.equal(model.sheets[0].cells.get('B4').formula,'A4*2'); assert.match(model.sheets[0].cells.get('B5').formula,/SUBTOTAL\(109/); assert.equal(model.sheets[0].tables[0].autoFilter.bottom,3);
});

test('계산 열 계획은 절대·상대 참조와 예외를 구분하고 메타데이터만 해제할 수 있다', () => {
  const {planCalculatedColumn,patchCalculatedColumns}=require('../src/xlsx'), model=analyze(nativeCalculatedFixture());
  const plan=planCalculatedColumn(model,{sheetIndex:0,row:2,column:1,formula:'A3+$A$2'});
  assert.deepEqual(plan.edits.map(edit=>edit.formula),['A2+$A$2','A3+$A$2']); assert.equal(plan.metadata.formula,'A2+$A$2');
  const removed=patchCalculatedColumns(model,[{sheetIndex:0,tableName:'Sales',column:1,formula:null}]).model;
  assert.equal(removed.sheets[0].tables[0].calculatedColumns[1],null); assert.equal(removed.sheets[0].cells.get('B2').formula,'A2*2');
  assert.deepEqual(removed.files['xl/worksheets/sheet1.xml'],model.files['xl/worksheets/sheet1.xml']);
});

test('첫 데이터 행 자동 채우기는 머리글 대신 인접 데이터 행 서식을 사용한다', () => {
  const model=patchCells(analyze(nativeCalculatedFixture()),[{sheetIndex:0,address:'B1',value:'Amount',font:{bold:true,italic:false}},{sheetIndex:0,address:'B2',value:'10',formula:'A2*2',font:{bold:false,italic:true}}]).model;
  const inserted=transformStructure(model,{sheetIndex:0,axis:'row',action:'insert',index:1}).model;
  assert.equal(inserted.sheets[0].cells.get('B2').styleProperties.italic,true);assert.notEqual(inserted.sheets[0].cells.get('B2').styleProperties.bold,true);
});

test('계산 열의 과도한 범위·배열 메타데이터·새 셀 유효성은 원본을 유지하며 거부한다', () => {
  const {planCalculatedColumn,patchCalculatedColumns}=require('../src/xlsx');
  const files=unzipSync(nativeCalculatedFixture());let xml=new TextDecoder().decode(files['xl/tables/table1.xml']);
  files['xl/tables/table1.xml']=new TextEncoder().encode(xml.replace('ref="A1:B3"','ref="A1:B100002"').replace(/<autoFilter\b[^>]*\/>/,''));let model=analyze(zipSync(files));assert.equal(model.writable,true,model.readOnlyReason);assert.throws(()=>planCalculatedColumn(model,{sheetIndex:0,row:1,column:1,formula:'A2*3'}),/100,000셀 이하/);
  files['xl/tables/table1.xml']=new TextEncoder().encode(xml.replace('<calculatedColumnFormula>','<calculatedColumnFormula array="1">'));model=analyze(zipSync(files));assert.equal(model.writable,true,model.readOnlyReason);assert.throws(()=>patchCalculatedColumns(model,[{sheetIndex:0,tableName:'Sales',column:1,formula:'A2*3'}]),/공유·배열/);
  files['xl/tables/table1.xml']=new TextEncoder().encode(xml);files['xl/worksheets/sheet1.xml']=new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</worksheet>','<dataValidations count="1"><dataValidation type="whole" operator="between" sqref="B4"><formula1>0</formula1><formula2>10</formula2></dataValidation></dataValidations></worksheet>'));
  model=analyze(zipSync(files));assert.equal(model.writable,true,model.readOnlyReason);const original=Buffer.from(model.bytes);assert.throws(()=>patchTableRanges(model,[{sheetIndex:0,tableName:'Sales',range:{top:0,bottom:3,left:0,right:1}}]),/자동 채우기를 중단/);assert.deepEqual(Buffer.from(model.bytes),original);
});

test('계산 열의 참조 이동은 셀 주소처럼 보이는 시트 이름과 한글 식별자를 바꾸지 않는다', () => {
  assert.equal(translateFormulaReferences('Q1!A2+JAN1!$B2+데이터A2!C$4+"Q1!A2"+Sales[@Amount]+한글A2',1,0),'Q1!A3+JAN1!$B3+데이터A2!C$4+"Q1!A2"+Sales[@Amount]+한글A2');
  const {planCalculatedColumn}=require('../src/xlsx'),model=analyze(nativeCalculatedFixture()),plan=planCalculatedColumn(model,{sheetIndex:0,row:1,column:1,formula:'Q1!A2+JAN1!$B2'});
  assert.deepEqual(plan.edits.map(edit=>edit.formula),['Q1!A2+JAN1!$B2','Q1!A3+JAN1!$B3']);
});

test('자동 표 확장 계획은 다른 표 소유 셀·병합·수식 머리글·정렬 상태와 충돌하지 않는다', () => {
  const {planTableExpansion,prepareTableExpansion}=require('../src/xlsx'),original=analyze(nativeCalculatedFixture());const bytes=Buffer.from(original.bytes);
  assert.equal(planTableExpansion(original,0,[{row:0,column:2,value:'',formula:'1+1'}]).requests.length,0);
  const files=unzipSync(original.bytes);files['xl/worksheets/sheet1.xml']=new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('</worksheet>','<mergeCells count="1"><mergeCell ref="A4:B4"/></mergeCells></worksheet>'));let model=analyze(zipSync(files));assert.equal(model.writable,true,model.readOnlyReason);assert.equal(planTableExpansion(model,0,[{row:3,column:0,value:'7'}]).requests.length,0);
  files['xl/worksheets/sheet1.xml']=original.files['xl/worksheets/sheet1.xml'];
  files['xl/tables/table1.xml']=new TextEncoder().encode(new TextDecoder().decode(files['xl/tables/table1.xml']).replace('</table>','<sortState ref="A1:B3"><sortCondition ref="B2:B3"/></sortState></table>'));
  model=analyze(zipSync(files));assert.equal(model.writable,true,model.readOnlyReason);
  const sortedSourceBytes=Buffer.from(model.bytes),sortedPlan=planTableExpansion(model,0,[{row:3,column:0,value:'7'}]);
  assert.equal(sortedPlan.requests.length,1);
  const sortedExpansion=patchTableRanges(prepareTableExpansion(model,sortedPlan.requests),sortedPlan.requests).model;
  const sortedXml=new TextDecoder().decode(sortedExpansion.files['xl/tables/table1.xml']);
  assert.match(sortedXml,/<sortState ref="A1:B4"><sortCondition ref="B2:B4"\/>/);
  assert.deepEqual(Buffer.from(model.bytes),sortedSourceBytes);assert.deepEqual(Buffer.from(original.bytes),bytes);
});

test('행·열 구조 변경은 Excel 표의 저장된 정렬 범위를 갱신한다', () => {
  const files = unzipSync(nativeCalculatedFixture());
  files['xl/tables/table1.xml'] = new TextEncoder().encode(new TextDecoder().decode(files['xl/tables/table1.xml']).replace('</table>', '<sortState ref="A1:B3"><sortCondition ref="B2:B3" descending="1"/></sortState></table>'));
  const original = analyze(zipSync(files));
  assert.equal(original.writable, true, original.readOnlyReason);

  const insertedRow = transformStructure(original, { sheetIndex: 0, axis: 'row', action: 'insert', index: 1 }).model;
  let tableXml = new TextDecoder().decode(insertedRow.files['xl/tables/table1.xml']);
  assert.match(tableXml, /<sortState ref="A1:B4"><sortCondition ref="B2:B4" descending="1"\/>/);
  const deletedRow = transformStructure(insertedRow, { sheetIndex: 0, axis: 'row', action: 'delete', index: 2 }).model;
  tableXml = new TextDecoder().decode(deletedRow.files['xl/tables/table1.xml']);
  assert.match(tableXml, /<sortState ref="A1:B3"><sortCondition ref="B2:B3" descending="1"\/>/);

  const insertedColumn = transformStructure(original, { sheetIndex: 0, axis: 'column', action: 'insert', index: 1 }).model;
  tableXml = new TextDecoder().decode(insertedColumn.files['xl/tables/table1.xml']);
  assert.match(tableXml, /<sortState ref="A1:C3"><sortCondition ref="C2:C3" descending="1"\/>/);

  const deletedSortColumn = transformStructure(original, { sheetIndex: 0, axis: 'column', action: 'delete', index: 1 }).model;
  tableXml = new TextDecoder().decode(deletedSortColumn.files['xl/tables/table1.xml']);
  assert.doesNotMatch(tableXml, /<sortState\b/);
  assert.equal(deletedSortColumn.sheets[0].tables[0].range.right, 0);

  const resizedRows = patchTableRanges(original, [{ sheetIndex: 0, tableName: 'Sales', range: { top: 0, bottom: 3, left: 0, right: 1 } }]).model;
  tableXml = new TextDecoder().decode(resizedRows.files['xl/tables/table1.xml']);
  assert.match(tableXml, /<sortState ref="A1:B4"><sortCondition ref="B2:B4" descending="1"\/>/);
  const shrunkRows = patchTableRanges(resizedRows, [{ sheetIndex: 0, tableName: 'Sales', range: { top: 0, bottom: 1, left: 0, right: 1 } }]).model;
  tableXml = new TextDecoder().decode(shrunkRows.files['xl/tables/table1.xml']);
  assert.match(tableXml, /<sortState ref="A1:B2"><sortCondition ref="B2" descending="1"\/>/);

  const withAdjacentHeader = patchCells(original, [{ sheetIndex: 0, address: 'C1', value: 'Extra', kind: 'string' }]).model;
  const widenedColumns = patchTableRanges(withAdjacentHeader, [{ sheetIndex: 0, tableName: 'Sales', range: { top: 0, bottom: 2, left: 0, right: 2 } }]).model;
  tableXml = new TextDecoder().decode(widenedColumns.files['xl/tables/table1.xml']);
  assert.match(tableXml, /<sortState ref="A1:C3"><sortCondition ref="B2:B3" descending="1"\/>/);
  const narrowedColumns = patchTableRanges(original, [{ sheetIndex: 0, tableName: 'Sales', range: { top: 0, bottom: 2, left: 0, right: 0 } }]).model;
  tableXml = new TextDecoder().decode(narrowedColumns.files['xl/tables/table1.xml']);
  assert.doesNotMatch(tableXml, /<sortState\b/);
});

test('자동 확장은 붙어 있는 표를 합치지 않고 겹치는 두 방향 확장을 생략한다', () => {
  const {planTableExpansion}=require('../src/xlsx'),files=unzipSync(nativeCalculatedFixture());
  files['xl/worksheets/sheet1.xml']=new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('<tableParts count="1">','<tableParts count="2">').replace('</tableParts>','<tablePart r:id="t2"/></tableParts>'));
  files['xl/worksheets/_rels/sheet1.xml.rels']=new TextEncoder().encode(new TextDecoder().decode(files['xl/worksheets/_rels/sheet1.xml.rels']).replace('</Relationships>','<Relationship Id="t2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table2.xml"/></Relationships>'));
  files['[Content_Types].xml']=new TextEncoder().encode(new TextDecoder().decode(files['[Content_Types].xml']).replace('</Types>','<Override PartName="/xl/tables/table2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>'));
  const table=ref=>'<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="2" name="Other" displayName="Other" ref="'+ref+'"><tableColumns count="2"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Value"/></tableColumns></table>';
  files['xl/tables/table2.xml']=new TextEncoder().encode(table('D1:E3'));let model=analyze(zipSync(files));assert.equal(model.writable,true,model.readOnlyReason);
  assert.equal(planTableExpansion(model,0,[{row:1,column:3,value:'Other data'}]).requests.length,0);
  const adjacent=planTableExpansion(model,0,[{row:1,column:2,value:'New'},{row:1,column:3,value:'Other data'}]);assert.equal(adjacent.requests.length,1);assert.equal(adjacent.requests[0].tableName,'Sales');assert.equal(adjacent.requests[0].range.right,2);
  files['xl/tables/table2.xml']=new TextEncoder().encode(table('C4:D6'));model=analyze(zipSync(files));assert.equal(model.writable,true,model.readOnlyReason);
  const overlap=planTableExpansion(model,0,[{row:2,column:2,value:'Right'},{row:3,column:1,value:'Below'}]);assert.equal(overlap.requests.length,0);assert.equal(model.sheets[0].tables[0].range.right,1);
});

test('자동 확장은 과도한 표 영역을 만들기 전에 거부한다', () => {
  const {planTableExpansion,prepareTableExpansion}=require('../src/xlsx'),files=unzipSync(nativeCalculatedFixture());files['xl/tables/table1.xml']=new TextEncoder().encode(new TextDecoder().decode(files['xl/tables/table1.xml']).replaceAll('A1:B3','A1:B50000'));
  const model=analyze(zipSync(files));assert.equal(model.writable,true,model.readOnlyReason);const plan=planTableExpansion(model,0,[{row:50000,column:0,value:'7'}]);assert.equal(plan.requests.length,1);assert.throws(()=>prepareTableExpansion(model,plan.requests),/100,000셀 이하/);assert.equal(model.sheets[0].tables[0].range.bottom,49999);
});

test('native top10 기본값과 캐시를 읽고 실제 숫자·수식으로 기준값을 갱신한다', () => {
  const files=unzipSync(fixture());files['xl/worksheets/sheet1.xml']=new TextEncoder().encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Rank</t></is></c></row><row r="2"><c r="A2"><v>10</v></c></row><row r="3"><c r="A3"><f>20+10</f><v>0</v></c></row><row r="4"><c r="A4"><v>20</v></c></row><row r="5"><c r="A5"><v>30</v></c></row><row r="6"><c r="A6" t="inlineStr"><is><t>9999</t></is></c></row><row r="7"><c r="A7" t="b"><v>1</v></c></row></sheetData><autoFilter ref="A1:A8"><filterColumn colId="0"><top10 val="2" filterVal="9999"/></filterColumn></autoFilter></worksheet>');
  const original=analyze(zipSync(files));assert.equal(original.writable,true,original.readOnlyReason);assert.deepEqual(original.sheets[0].autoFilter.columns[0].top10,{top:true,percent:false,value:2,filterValue:9999});const model=patchFilterVisibility(original).model;assert.deepEqual(model.sheets[0].nativeHiddenRows,[1,3,5,6,7]);assert.equal(model.sheets[0].autoFilter.columns[0].top10.filterValue,30);
  const shifted=transformStructure(model,{sheetIndex:0,axis:'column',action:'insert',index:0}).model;assert.equal(shifted.sheets[0].autoFilter.columns[0].column,1);assert.equal(shifted.sheets[0].autoFilter.columns[0].top10.value,2);assert.equal(shifted.sheets[0].autoFilter.columns[0].top10.filterValue,30);assert.deepEqual(shifted.sheets[0].nativeHiddenRows,[1,3,5,6,7]);
});

test('잘못된 순위 필터 XML과 혼합 조건은 원본을 변경하지 않고 거부한다', () => {
  for(const attrs of ['val="0"','val="501"','val="1.5"','val="101" percent="1"','val="2" top="yes"','val="2" filterVal="NaN"','val="0x10"','percent="1"']){
    const model=analyze(fixture({autoFilter:'<autoFilter ref="A1:C2"><filterColumn colId="0"><top10 '+attrs+'/></filterColumn></autoFilter>'}));assert.equal(model.writable,false,attrs);
  }
  const model=analyze(fixture());for(const column of [{column:0,top10:null},{column:0,top10:{top:true,percent:false,value:2},custom:{and:false,conditions:[{operator:'equal',value:'x'}]}},{column:0,top10:{top:true,percent:false,value:1},values:['x'],blanks:false}])assert.throws(()=>patchAutoFilters(model,[{sheetIndex:0,filter:{top:0,bottom:1,left:0,right:2,columns:[column]}}]));assert.equal(model.sheets[0].autoFilter,null);
});

test('숫자가 없는 순위 필터는 모든 데이터를 제외하고 오래된 filterVal을 제거한다', () => {
  const files=unzipSync(fixture());files['xl/worksheets/sheet1.xml']=new TextEncoder().encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Header</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>100</t></is></c></row></sheetData><autoFilter ref="A1:A3"><filterColumn colId="0"><top10 val="1" filterVal="100"/></filterColumn></autoFilter></worksheet>');const model=patchFilterVisibility(analyze(zipSync(files))).model;assert.deepEqual(model.sheets[0].nativeHiddenRows,[1,2]);assert.doesNotMatch(model.sheets[0].xml,/filterVal=/);assert.equal(model.sheets[0].autoFilter.columns[0].top10.filterValue,undefined);
});

test('native 날짜 그룹은 Gregorian·개별 값·빈 셀과 시간 정밀도를 보존한다', () => {
  const files=unzipSync(fixture());files['xl/worksheets/sheet1.xml']=new TextEncoder().encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Date</t></is></c></row><row r="2"><c r="A2"><v>45351.5</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>keep</t></is></c></row><row r="4"><c r="A4"><v>45352</v></c></row></sheetData><autoFilter ref="A1:A5"><filterColumn colId="0"><filters blank="true" calendarType="gregorian"><filter val="keep"/><dateGroupItem year="2024" month="2" day="29" hour="12" minute="0" dateTimeGrouping="minute"/></filters></filterColumn></autoFilter></worksheet>');const original=analyze(zipSync(files));assert.equal(original.writable,true,original.readOnlyReason);const model=patchFilterVisibility(original).model;assert.deepEqual(model.sheets[0].nativeHiddenRows,[3]);assert.equal(model.sheets[0].autoFilter.columns[0].dateGroups[0].minute,0);
  const shifted=transformStructure(model,{sheetIndex:0,axis:'column',action:'insert',index:0}).model;assert.equal(shifted.sheets[0].autoFilter.columns[0].column,1);assert.equal(shifted.sheets[0].autoFilter.columns[0].dateGroups[0].hour,12);assert.deepEqual(shifted.sheets[0].autoFilter.columns[0].values,['keep']);
});

test('잘못된 날짜 그룹·달력·과도한 그룹은 안전 편집을 허용하지 않는다', () => {
  for(const group of ['year="2023" month="2" day="29" dateTimeGrouping="day"','year="2024" month="13" dateTimeGrouping="month"','year="2024" month="2" day="29" hour="24" dateTimeGrouping="hour"','year="2024" dateTimeGrouping="quarter"','year="2024" month="2" day="1" dateTimeGrouping="month"','year="2024x" dateTimeGrouping="year"'])assert.equal(analyze(fixture({autoFilter:'<autoFilter ref="A1:C2"><filterColumn colId="0"><filters><dateGroupItem '+group+'/></filters></filterColumn></autoFilter>'})).writable,false,group);
  assert.equal(analyze(fixture({autoFilter:'<autoFilter ref="A1:C2"><filterColumn colId="0"><filters calendarType="hijri"><dateGroupItem year="2024" dateTimeGrouping="year"/></filters></filterColumn></autoFilter>'})).writable,false);
  const model=analyze(fixture());assert.throws(()=>patchAutoFilters(model,[{sheetIndex:0,filter:{top:0,bottom:1,left:0,right:2,columns:[{column:0,values:[],blanks:false,dateGroups:Array(10001).fill({dateTimeGrouping:'year',year:2024})}]}}]),/10,000/);assert.equal(model.sheets[0].autoFilter,null);
});

test('날짜 표시와 필터는 1904 체계·1900 윤일의 시간·일 경계 반올림을 공유한다', () => {
  const cell=(value,date1904=false)=>({value:String(value),kind:'number',numFmtId:164,formatCode:'yyyy-mm-dd hh:mm:ss',date1904});assert.equal(formatCellValue(cell(60.5)),'1900-02-29 12:00:00');assert.equal(formatCellValue(cell(0,true)),'1904-01-01 00:00:00');assert.equal(formatCellValue(cell(59.99999999999)),'1900-02-29 00:00:00');assert.equal(formatCellValue(cell(2958465.5)),'9999-12-31 12:00:00');
  const files=unzipSync(fixture({formatted:true}));files['xl/workbook.xml']=new TextEncoder().encode(new TextDecoder().decode(files['xl/workbook.xml']).replace('<sheets>','<workbookPr date1904="true"/><sheets>'));const model=analyze(zipSync(files));assert.equal(model.date1904,true);assert.equal(model.sheets[0].cells.get('A1').date1904,true);assert.equal(analyze(patchCell(model,0,'A1','changed').bytes).date1904,true);
});

test('1904 날짜 수식은 날짜 생성·분해·요일·TEXT와 현재 시각을 같은 체계로 계산한다', () => {
  const context={date1904:true},run=formula=>evaluateFormula(formula,context);
  for(const [formula,value] of [['DATE(1904,1,1)',0],['DATE(2024,2,29)',43889],['DATEVALUE("2024-02-29")',43889],['YEAR(0)',1904],['MONTH(0)',1],['DAY(0)',1],['WEEKDAY(0)',6],['WEEKNUM(0)',1],['ISOWEEKNUM(0)',53],['HOUR(43889.5)',12],['MINUTE(43889.5)',0],['TEXT(0,"yyyy-mm-dd")','1904-01-01'],['TEXT(DATE(2024,2,29)+0.5,"yyyy-mm-dd hh:mm")','2024-02-29 12:00'],['TEXT(123.45,"0.00")','123.45'],['DATE(2024,2,29)+1',43890]])assert.equal(run(formula),value,formula);
  const now=new Date(2024,1,29,12,30);assert.equal(evaluateFormula('TODAY()',{date1904:true,now}),43889);assert.equal(evaluateFormula('NOW()',{date1904:true,now}),43889+12.5/24);assert.equal(evaluateFormula('NOW()',{date1904:true,todaySerial:43889.25}),43889.25);assert.equal(evaluateFormula('DATE(2024,2,29)'),45351);
});

test('1904 날짜 이동·차이·근무일·금융 계산은 같은 실제 날짜의 1900 결과와 일치한다', () => {
  const dates=['EDATE(DATE(2024,1,31),1)','EOMONTH(DATE(2024,1,1),1)','WORKDAY(DATE(2024,2,29),1)','WORKDAY.INTL(DATE(2024,2,29),2,1,DATE(2024,3,1))','COUPNCD(DATE(2024,1,15),DATE(2026,12,31),2)','COUPPCD(DATE(2024,1,15),DATE(2026,12,31),2)'];
  for(const formula of dates){const normal=evaluateFormula(formula),epoch=evaluateFormula(formula,{date1904:true});assert.equal(typeof normal,'number',formula);assert.equal(epoch,normal-1462,formula)}
  const same=['DAYS(DATE(2024,3,4),DATE(2024,2,29))','DATEDIF(DATE(2024,2,29),DATE(2025,3,1),"Y")','DAYS360(DATE(2024,2,29),DATE(2025,3,1))','YEARFRAC(DATE(2024,2,29),DATE(2025,3,1),1)','NETWORKDAYS(DATE(2024,2,29),DATE(2024,3,4),DATE(2024,3,1))','NETWORKDAYS.INTL(DATE(2024,2,29),DATE(2024,3,4),"0000011",DATE(2024,3,1))','COUPDAYBS(DATE(2024,1,15),DATE(2026,12,31),2)','COUPDAYS(DATE(2024,1,15),DATE(2026,12,31),2)','COUPDAYSNC(DATE(2024,1,15),DATE(2026,12,31),2)','COUPNUM(DATE(2024,1,15),DATE(2026,12,31),2)','PRICE(DATE(2024,1,15),DATE(2026,12,31),0.05,0.06,100,2)','YIELD(DATE(2024,1,15),DATE(2026,12,31),0.05,98,100,2)','ACCRINTM(DATE(2024,1,1),DATE(2024,7,1),0.05,100,1)','DISC(DATE(2024,1,1),DATE(2024,7,1),98,100,1)','INTRATE(DATE(2024,1,1),DATE(2024,7,1),98,100,1)','RECEIVED(DATE(2024,1,1),DATE(2024,7,1),98,0.05,1)','PRICEDISC(DATE(2024,1,1),DATE(2024,7,1),0.05,100,1)','YIELDDISC(DATE(2024,1,1),DATE(2024,7,1),98,100,1)','XNPV(0.1,HSTACK(-100,110),HSTACK(DATE(2024,1,1),DATE(2025,1,1)))','XIRR(HSTACK(-100,110),HSTACK(DATE(2024,1,1),DATE(2025,1,1)))'];
  for(const formula of same){const normal=evaluateFormula(formula),epoch=evaluateFormula(formula,{date1904:true});assert.equal(typeof normal,'number',formula);assert.ok(Math.abs(epoch-normal)<1e-10,formula+': '+epoch+' vs '+normal)}
});

test('1904 날짜 수식은 잘못된 인수·음수·상한과 일반 숫자를 구분한다', () => {
  for(const [formula,value] of [['YEAR(-1)','#NUM!'],['YEAR(2957004)','#NUM!'],['YEAR(-1,2)','#VALUE!'],['DATE(1900,1,1)','#NUM!'],['WORKDAY(0,-1)','#NUM!'],['DATE(9999,12,31)',2957003],['YEAR(2957003)',9999],['SUM(1,2)',3],['TIME(12,0,0)',.5],['TIMEVALUE("12:00")',.5],['YEAR("no date")','#VALUE!']])assert.equal(evaluateFormula(formula,{date1904:true}),value,formula);
});

test('1904 희소 날짜 범위와 휴일은 원본을 유지하고 현금 흐름·근무일을 계산한다', () => {
  const files=unzipSync(fixture());files['xl/workbook.xml']=new TextEncoder().encode(new TextDecoder().decode(files['xl/workbook.xml']).replace('<sheets>','<workbookPr date1904="1"/><sheets>'));files['xl/worksheets/sheet1.xml']=new TextEncoder().encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>-100</v></c><c r="B1"><f>DATE(2024,1,1)</f><v>0</v></c><c r="C1"><f>DATE(2024,3,1)</f><v>0</v></c></row><row r="2"><c r="A2"><v>110</v></c><c r="B2"><f>DATE(2025,1,1)</f><v>0</v></c><c r="C2" t="inlineStr"><is><t>ignored</t></is></c></row><row r="4"><c r="D4"><f>XNPV(0.1,A1:A2,B1:B2)</f><v>0</v></c><c r="E4"><f>NETWORKDAYS(DATE(2024,2,29),DATE(2024,3,4),C1:C2)</f><v>0</v></c><c r="F4"><f>WORKDAY(DATE(2024,2,29),1,C1:C2)</f><v>0</v></c></row></sheetData></worksheet>');const model=analyze(zipSync(files)),calculator=createFormulaCalculator({state:model,getCell:(row,column,index=0)=>model.sheets[index].cells.get(String.fromCharCode(65+column)+(row+1))});assert.equal(calculator.value(0,3,4),2);assert.equal(calculator.value(0,3,5),43893);const expected=evaluateFormula('XNPV(0.1,HSTACK(-100,110),HSTACK(DATE(2024,1,1),DATE(2025,1,1)))');assert.ok(Math.abs(calculator.value(0,3,3)-expected)<1e-10);assert.equal(model.sheets[0].cells.get('B1').formula,'DATE(2024,1,1)');assert.equal(model.sheets[0].cells.get('C2').value,'ignored');
});

test('1904 마지막 날짜의 오후는 날짜 분해·TEXT·시간 함수에서 유효하다', () => {
  for(const date1904 of [false,true]){const serial=2958465-(date1904?1462:0)+.5,context={date1904};for(const [name,value] of [['YEAR',9999],['MONTH',12],['DAY',31],['HOUR',12],['MINUTE',0],['SECOND',0]])assert.equal(evaluateFormula(name+'('+serial+')',context),value,name);assert.equal(evaluateFormula('TEXT('+serial+',"yyyy-mm-dd hh:mm")',context),'9999-12-31 12:00');assert.equal(evaluateFormula('YEAR('+(Math.floor(serial)+1)+')',context),'#NUM!')}
});

test('native dynamicFilter는 오래된 기준값을 날짜·평균 판정으로 갱신하고 구조 변경을 보존한다', () => {
  const files=unzipSync(fixture());files['xl/worksheets/sheet1.xml']=new TextEncoder().encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Date</t></is></c></row><row r="2"><c r="A2"><v>45351.5</v></c></row><row r="3"><c r="A3"><v>45352</v></c></row><row r="4"><c r="A4"><f>DATE(2024,2,29)+0.75</f><v>0</v></c></row></sheetData><autoFilter ref="A1:A4"><filterColumn colId="0"><dynamicFilter type="today" val="1" maxVal="2"/></filterColumn></autoFilter></worksheet>');const original=analyze(zipSync(files));assert.equal(original.writable,true,original.readOnlyReason);const model=patchFilterVisibility(original,{now:new Date(2024,1,29)}).model;assert.deepEqual(model.sheets[0].nativeHiddenRows,[2]);assert.deepEqual(model.sheets[0].autoFilter.columns[0].dynamic,{type:'today',value:45351,maxValue:45352});const next=patchFilterVisibility(model,{now:new Date(2024,2,1)}).model;assert.deepEqual(next.sheets[0].nativeHiddenRows,[1,3]);assert.equal(next.sheets[0].autoFilter.columns[0].dynamic.value,45352);
  const averages=patchAutoFilters(original,[{sheetIndex:0,filter:{top:0,bottom:3,left:0,right:0,columns:[{column:0,dynamic:{type:'belowAverage'}}]}}]).model;assert.equal(averages.sheets[0].autoFilter.columns[0].dynamic.type,'belowAverage');assert.equal(averages.sheets[0].autoFilter.columns[0].dynamic.maxValue,undefined);const shifted=transformStructure(averages,{sheetIndex:0,axis:'column',action:'insert',index:0}).model;assert.equal(shifted.sheets[0].autoFilter.columns[0].column,1);assert.equal(shifted.sheets[0].autoFilter.columns[0].dynamic.type,'belowAverage');
});

test('잘못된 dynamicFilter 속성·값과 혼합 조건은 원본을 유지하며 거부한다', () => {
  for(const attrs of ['type="unknown"','type="today" val="NaN"','type="today" val="0x10"','type="today" maxVal="INF"','type="today" other="1"','val="1"'])assert.equal(analyze(fixture({autoFilter:'<autoFilter ref="A1:C2"><filterColumn colId="0"><dynamicFilter '+attrs+'/></filterColumn></autoFilter>'})).writable,false,attrs);
  const model=analyze(fixture());for(const column of [{column:0,dynamic:null},{column:0,dynamic:{type:'today'},values:[],blanks:false},{column:0,dynamic:{type:'today'},top10:{top:true,percent:false,value:1}}])assert.throws(()=>patchAutoFilters(model,[{sheetIndex:0,filter:{top:0,bottom:1,left:0,right:2,columns:[column]}}]));assert.equal(model.sheets[0].autoFilter,null);
});

test('색상 필터는 서식 없는 파일에 DXF만 추가하고 재사용·자동 글자색·구조 변경을 보존한다', () => {
  const original=analyze(fixture()),before=original.sheets[0].cells.get('A1'),filter={top:0,bottom:1,left:0,right:2,columns:[{column:0,colorFilter:{cellColor:true,color:null}}]};const first=patchAutoFilters(original,[{sheetIndex:0,filter}]).model;assert.equal(first.writable,true,first.readOnlyReason);assert.equal(first.sheets[0].cells.get('A1').value,before.value);assert.equal(first.sheets[0].cells.get('A1').style,before.style);assert.equal(first.sheets[0].autoFilter.columns[0].colorFilter.dxfId,0);const repeated=patchAutoFilters(first,[{sheetIndex:0,filter}]).model;assert.equal((new TextDecoder().decode(repeated.files[repeated.stylesPath]).match(/<dxf>/g)||[]).length,1);const font=patchAutoFilters(repeated,[{sheetIndex:0,filter:{...filter,columns:[{column:0,colorFilter:{cellColor:false,color:null}}]}}]).model;assert.equal(font.sheets[0].autoFilter.columns[0].colorFilter.cellColor,false);assert.equal(font.sheets[0].autoFilter.columns[0].colorFilter.color,null);const moved=transformStructure(font,{sheetIndex:0,axis:'column',action:'insert',index:0}).model;assert.equal(moved.sheets[0].autoFilter.columns[0].column,1);assert.equal(moved.sheets[0].autoFilter.columns[0].colorFilter.color,null);assert.equal(moved.writable,true,moved.readOnlyReason);
});

test('색상 필터의 잘못된 DXF·색·혼합 조건과 누락한 테마는 원본을 보존하며 거부한다', () => {
  for(const attrs of ['dxfId="-1"','dxfId="999"','dxfId="0" cellColor="yes"','cellColor="1"'])assert.equal(analyze(fixture({autoFilter:'<autoFilter ref="A1:C2"><filterColumn colId="0"><colorFilter '+attrs+'/></filterColumn></autoFilter>'})).writable,false,attrs);
  const original=analyze(fixture());for(const column of [{column:0,colorFilter:{cellColor:true,color:'red'}},{column:0,colorFilter:{cellColor:true,color:'#FF0000'},values:[],blanks:false}])assert.throws(()=>patchAutoFilters(original,[{sheetIndex:0,filter:{top:0,bottom:1,left:0,right:2,columns:[column]}}]));assert.equal(original.sheets[0].autoFilter,null);
  const valid=patchAutoFilters(original,[{sheetIndex:0,filter:{top:0,bottom:1,left:0,right:2,columns:[{column:0,colorFilter:{cellColor:true,color:'#FF0000'}}]}}]).model;const files={...valid.files};files[valid.stylesPath]=new TextEncoder().encode(new TextDecoder().decode(files[valid.stylesPath]).replace('<fgColor rgb="FFFF0000"/>','<fgColor theme="1"/>'));const unsupported=analyze(zipSync(files));assert.equal(unsupported.writable,false);assert.match(unsupported.readOnlyReason,/색상 필터의 단색 정의를 해석할 수 없습니다/);const defaultFiles={...valid.files};defaultFiles[valid.sheets[0].path]=new TextEncoder().encode(valid.sheets[0].xml.replace(' cellColor="1"',''));assert.equal(analyze(zipSync(defaultFiles)).sheets[0].autoFilter.columns[0].colorFilter.cellColor,true);
});

function themeColorFixture({ conditional = false } = {}) {
  const files=unzipSync(fixture({formatted:true})),encode=text=>new TextEncoder().encode(text),decode=bytes=>new TextDecoder().decode(bytes);
  files['xl/_rels/workbook.xml.rels']=encode(decode(files['xl/_rels/workbook.xml.rels']).replace('</Relationships>','<Relationship Id="rTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/custom.xml"/></Relationships>'));
  files['[Content_Types].xml']=encode(decode(files['[Content_Types].xml']).replace('</Types>','<Override PartName="/xl/theme/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>'));
  files['xl/theme/custom.xml']=encode('<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Custom"><a:themeElements><a:clrScheme name="Test"><a:dk1><a:sysClr val="windowText" lastClr="123456"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:accent1><a:srgbClr val="4F81BD"/></a:accent1></a:clrScheme></a:themeElements></a:theme>');
  let styles=decode(files['xl/styles.xml']).replace('<color rgb="FF123456"/>','<color theme="1"/>').replace('<fgColor rgb="FFFFEEAA"/>','<fgColor theme="4" tint="0.4"/>').replace('<color rgb="FFABCDEF"/>','<color indexed="1"/>');
  styles=styles.replace('</styleSheet>','<dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor theme="4" tint="0.4"/></patternFill></fill></dxf></dxfs><colors><indexedColors><rgbColor rgb="FF112233"/><rgbColor rgb="FFABCDEF"/></indexedColors></colors></styleSheet>');
  files['xl/styles.xml']=encode(styles);
  files['xl/worksheets/sheet1.xml']=encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="B1" t="inlineStr"><is><t>Amount</t></is></c></row><row r="2"><c r="B2" s="3"><v>10</v></c></row><row r="3"><c r="B3"><v>20</v></c></row><row r="4"><c r="C4"><f>SUBTOTAL(9,B2:B3)</f><v>0</v></c></row></sheetData><autoFilter ref="B1:B3"><filterColumn colId="0"><colorFilter dxfId="0"/></filterColumn></autoFilter>'+(conditional?'<conditionalFormatting sqref="B2:B3"><cfRule type="cellIs" operator="greaterThan" priority="1" dxfId="0"><formula>15</formula></cfRule></conditionalFormatting>':'')+'</worksheet>');
  return {files,bytes:zipSync(files)};
}

test('테마·밝기·사용자 인덱스 색은 화면과 native 필터·합계·저장에서 일치한다',()=>{
  const {bytes}=themeColorFixture(),original=analyze(bytes);assert.equal(original.writable,true,original.readOnlyReason);
  const properties=original.sheets[0].cells.get('B2').styleProperties;
  assert.equal(properties.fill,'#95B3D7');assert.equal(properties.color,'#123456');assert.equal(properties.borders.left.color,'#ABCDEF');
  assert.equal(original.sheets[0].autoFilter.columns[0].colorFilter.color,'#95B3D7');
  const model=patchFilterVisibility(original).model;assert.deepEqual(model.sheets[0].nativeHiddenRows,[2]);
  const calculator=createFormulaCalculator({state:model,getCell:(row,column,index=0)=>model.sheets[index].cells.get(String.fromCharCode(65+column)+(row+1))});
  assert.equal(calculator.value(0,3,2),10);
  const filter={...model.sheets[0].autoFilter,columns:[{column:1,colorFilter:{cellColor:true,color:'#95B3D7'}}]};
  const reapplied=patchAutoFilters(model,[{sheetIndex:0,filter}]).model;
  assert.equal(reapplied.sheets[0].autoFilter.columns[0].colorFilter.dxfId,0);
  assert.deepEqual(reapplied.files['xl/styles.xml'],original.files['xl/styles.xml']);
  assert.deepEqual(reapplied.files['xl/theme/custom.xml'],original.files['xl/theme/custom.xml']);
  assert.equal(analyze(reapplied.bytes).sheets[0].cells.get('B2').styleProperties.fill,'#95B3D7');
  const font=patchAutoFilters(reapplied,[{sheetIndex:0,filter:{...filter,columns:[{column:1,colorFilter:{cellColor:false,color:'#123456'}}]}}]).model;
  assert.equal(font.sheets[0].autoFilter.columns[0].colorFilter.color,'#123456');assert.deepEqual(font.sheets[0].nativeHiddenRows,[2]);
});

test('테마 조건부 서식의 표시색도 같은 색 필터로 판정한다',()=>{
  const {bytes}=themeColorFixture({conditional:true}),original=analyze(bytes);assert.equal(original.writable,true,original.readOnlyReason);
  const model=patchFilterVisibility(original).model;assert.deepEqual(model.sheets[0].nativeHiddenRows,[]);
  const calculator=createFormulaCalculator({state:model,getCell:(row,column,index=0)=>model.sheets[index].cells.get(String.fromCharCode(65+column)+(row+1))});
  assert.equal(calculator.value(0,3,2),30);assert.equal(calculator.filterStyle(0,2,1).fill,'#95B3D7');
});

test('인덱스 native 색 필터는 사용자 팔레트를 사용하고 잘못된 테마·밝기는 거부한다',()=>{
  const {files}=themeColorFixture(),encode=text=>new TextEncoder().encode(text),decode=bytes=>new TextDecoder().decode(bytes);
  const styles=decode(files['xl/styles.xml']);
  files['xl/styles.xml']=encode(styles.replaceAll('theme="4" tint="0.4"','indexed="1"'));
  const indexed=analyze(zipSync(files));assert.equal(indexed.writable,true,indexed.readOnlyReason);assert.equal(indexed.sheets[0].autoFilter.columns[0].colorFilter.color,'#ABCDEF');assert.equal(indexed.sheets[0].cells.get('B2').styleProperties.fill,'#ABCDEF');
  for(const attrs of ['theme="12"','theme="5"','theme="4" tint="1.1"','indexed="64"','theme="4" rgb="FF000000"']){
    files['xl/styles.xml']=encode(styles.replaceAll('theme="4" tint="0.4"',attrs));const bad=analyze(zipSync(files));assert.equal(bad.writable,false,attrs);
  }
});

test('테마·인덱스 색은 조건부 글꼴·테두리·색상 척도·데이터 막대에도 전달한다',()=>{
  const {files}=themeColorFixture(),encode=text=>new TextEncoder().encode(text),decode=bytes=>new TextDecoder().decode(bytes);
  const style=decode(files['xl/styles.xml']).replace('<dxf><fill><patternFill patternType="solid"><fgColor theme="4" tint="0.4"/></patternFill></fill></dxf>','<dxf><font><color theme="1"/></font><border><left style="thin"><color indexed="1"/></left></border></dxf>');
  files['xl/styles.xml']=encode(style);
  let xml=decode(files['xl/worksheets/sheet1.xml']).replace(/<autoFilter[\s\S]*?<\/autoFilter>/,'');
  xml=xml.replace('</worksheet>','<conditionalFormatting sqref="B2:B3"><cfRule type="cellIs" operator="greaterThan" priority="1" dxfId="0"><formula>0</formula></cfRule><cfRule type="colorScale" priority="2"><colorScale><cfvo type="min"/><cfvo type="max"/><color theme="4" tint="0.4"/><color indexed="1"/></colorScale></cfRule><cfRule type="dataBar" priority="3"><dataBar><cfvo type="min"/><cfvo type="max"/><color theme="4"/></dataBar></cfRule></conditionalFormatting></worksheet>');
  files['xl/worksheets/sheet1.xml']=encode(xml);const model=analyze(zipSync(files));assert.equal(model.writable,true,model.readOnlyReason);
  const rules=model.sheets[0].conditionalFormats;assert.equal(rules[0].style.color,'#123456');assert.equal(rules[0].style.borders.left.color,'#ABCDEF');assert.deepEqual(rules[1].colorScale.colors,['#95B3D7','#ABCDEF']);assert.equal(rules[2].dataBar.color,'#4F81BD');
});

function isoDateFixture(values, date1904=false) {
  const files=unzipSync(fixture({formatted:true})),encode=text=>new TextEncoder().encode(text);
  if(date1904)files['xl/workbook.xml']=encode(new TextDecoder().decode(files['xl/workbook.xml']).replace('<sheets>','<workbookPr date1904="1"/><sheets>'));
  files['xl/worksheets/sheet1.xml']=encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Date</t></is></c></row>'+values.map((value,index)=>'<row r="'+(index+2)+'"><c r="A'+(index+2)+'" s="3" t="d"><v>'+value+'</v></c></row>').join('')+'<row r="20"><c r="B20"><f>YEAR(A2)</f><v>0</v></c><c r="C20"><f>SUBTOTAL(9,A2:A'+(values.length+1)+')</f><v>0</v></c></row></sheetData></worksheet>');
  return zipSync(files);
}

test('ISO 날짜 셀은 1900/1904 날짜 체계의 숫자로 계산·표시한다',()=>{
  for(const date1904 of [false,true]){
    const model=analyze(isoDateFixture(['2024-02-29','2024-02-29T12:00:00'],date1904));assert.equal(model.writable,true,model.readOnlyReason);
    const first=model.sheets[0].cells.get('A2'),second=model.sheets[0].cells.get('A3'),base=date1904?43889:45351;
    assert.equal(first.kind,'number');assert.equal(Number(first.value),base);assert.equal(Number(second.value),base+.5);assert.equal(formatCellValue(first),'2024-02-29');
    const calculator=createFormulaCalculator({state:model,getCell:(row,column,index=0)=>model.sheets[index].cells.get(String.fromCharCode(65+column)+(row+1))});assert.equal(calculator.value(0,19,1),2024);assert.equal(calculator.value(0,19,2),base*2+.5);
  }
});

test('ISO 시간대·분수초·24시를 로컬 시간대와 무관하게 해석한다',()=>{
  const model=analyze(isoDateFixture(['2024-03-01T00:30:00+01:00','2024-02-29T23:30:00Z','2024-02-29T09:00:00-14:00','2024-02-29T23:00:00Z','2024-02-29T24:00:00','2024-03-01','2024-02-29T00:00:00.125','2024-02-29+14:00']));
  const value=index=>Number(model.sheets[0].cells.get('A'+(index+2)).value);
  assert.equal(value(0),value(1));assert.equal(value(2),value(3));assert.equal(value(4),value(5));assert.ok(Math.abs((value(6)-45351)*86400-.125)<1e-6);assert.ok(Math.abs(value(7)-(45351-14/24))<1e-10);
});

test('ISO 1900 연초·윤일 경계와 1904 기준일을 실제 달력으로 해석한다',()=>{
  const normal=analyze(isoDateFixture(['1900-01-01','1900-02-28','1900-03-01','1900-02-28T23:00:00-01:00']));
  assert.deepEqual([2,3,4,5].map(row=>Number(normal.sheets[0].cells.get('A'+row).value)),[1,59,61,61]);
  const epoch=analyze(isoDateFixture(['1904-01-01','1904-01-02'],true));assert.equal(Number(epoch.sheets[0].cells.get('A2').value),0);assert.equal(Number(epoch.sheets[0].cells.get('A3').value),1);
});

test('잘못된 ISO 날짜·윤일·시각·시간대는 원본을 변경하지 않고 거부한다',()=>{
  for(const value of ['','2024-02-30','1900-02-29','2023-02-29','0000-01-01','2024-00-01','2024-13-01','2024-01-00','2024-01-01T25:00:00','2024-01-01T24:00:01','2024-01-01T24:00:00.1','2024-01-01T12:60:00','2024-01-01T12:00:60','2024-01-01T12:00:00+14:01','2024-01-01T12:00:00+15:00','2024-01-01T12:00:00+01:60','9999-12-31T24:00:00','2024/01/01','45351','2024-01-01 trailing']){
    const bytes=isoDateFixture([value]),before=Buffer.from(bytes);assert.throws(()=>analyze(bytes),/ISO 날짜 셀/,value);assert.deepEqual(Buffer.from(bytes),before);
  }
});

test('ISO 셀은 날짜 필터·native 숨김·SUBTOTAL에서 숫자 날짜와 동일하다',()=>{
  for(const date1904 of [false,true]){
    const original=analyze(isoDateFixture(['2024-02-29T12:00:00','2024-03-01'],date1904));
    const model=patchAutoFilters(original,[{sheetIndex:0,filter:{top:0,bottom:2,left:0,right:0,columns:[{column:0,values:[],blanks:false,dateGroups:[{dateTimeGrouping:'day',year:2024,month:2,day:29}]}]}}]).model;
    assert.deepEqual(model.sheets[0].nativeHiddenRows,[2]);assert.equal(model.sheets[0].cells.get('A2').isoDate,'2024-02-29T12:00:00');
    const calculator=createFormulaCalculator({state:model,getCell:(row,column,index=0)=>model.sheets[index].cells.get(String.fromCharCode(65+column)+(row+1))});assert.equal(calculator.value(0,19,2),(date1904?43889:45351)+.5);
  }
});

test('ISO 표현은 무관한 수정·서식 편집에 보존되고 값 편집은 기존 날짜 서식을 유지한다',()=>{
  const original=analyze(isoDateFixture(['2024-03-01T00:30:00+01:00'])),raw=original.sheets[0].cells.get('A2').isoDate;
  const unrelated=patchCell(original,0,'A1','Changed').model;assert.equal(unrelated.sheets[0].cells.get('A2').isoDate,raw);
  const styled=patchCells(unrelated,[{sheetIndex:0,address:'A2',value:unrelated.sheets[0].cells.get('A2').value,font:{bold:true,italic:false}}]).model;
  assert.equal(styled.sheets[0].cells.get('A2').isoDate,raw);assert.equal(styled.sheets[0].cells.get('A2').styleProperties.bold,true);
  const edited=patchCell(styled,0,'A2','45352').model;assert.equal(edited.sheets[0].cells.get('A2').isoDate,undefined);assert.equal(formatCellValue(edited.sheets[0].cells.get('A2')),'2024-03-01');assert.equal(analyze(edited.bytes).sheets[0].cells.get('A2').value,'45352');
});

test('ISO 날짜는 숫자 표시 형식을 따르고 구조 이동·수식 캐시를 보존한다',()=>{
  const files=unzipSync(isoDateFixture(['2024-02-29T12:00:00Z'])),encode=text=>new TextEncoder().encode(text);let xml=new TextDecoder().decode(files['xl/worksheets/sheet1.xml']);
  files['xl/worksheets/sheet1.xml']=encode(xml.replace('s="3" t="d"','t="d"'));
  const general=analyze(zipSync(files));assert.equal(formatCellValue(general.sheets[0].cells.get('A2')),'45351.5');
  const moved=transformStructure(general,{sheetIndex:0,axis:'column',action:'insert',index:0}).model;assert.equal(moved.sheets[0].cells.get('B2').isoDate,'2024-02-29T12:00:00Z');assert.equal(moved.sheets[0].cells.get('C20').formula,'YEAR(B2)');
  files['xl/worksheets/sheet1.xml']=encode(xml.replace('<v>2024-02-29T12:00:00Z</v>','<f>DATE(2024,2,29)+0.5</f><v>2024-02-29T12:00:00Z</v>'));
  const model=analyze(zipSync(files)),calculator=createFormulaCalculator({state:model,getCell:(row,column,index=0)=>model.sheets[index].cells.get(String.fromCharCode(65+column)+(row+1))});assert.equal(calculator.value(0,1,0),45351.5);assert.equal(model.sheets[0].cells.get('A2').cachedValue,'45351.5');
});

test('선택 범위는 native 표·관계·형식을 추가하고 기존 셀 서식을 보존한다',()=>{
  const original=analyze(fixture({formatted:true})),result=patchAddTable(original,{sheetIndex:0,range:{top:0,bottom:1,left:0,right:3},name:'매출_2026'}).model;assert.equal(result.writable,true,result.readOnlyReason);const table=result.sheets[0].tables[0];assert.equal(table.name,'매출_2026');assert.deepEqual(table.range,{top:0,bottom:1,left:0,right:3});assert.deepEqual(table.columns,['before','1/1/24','12.50%','2024-01-01']);assert.equal(result.sheets[0].cells.get('D1').style,original.sheets[0].cells.get('D1').style);assert.equal(result.sheets[0].cells.get('D1').kind,'string');assert.deepEqual(result.files['xl/worksheets/sheet2.xml'],original.files['xl/worksheets/sheet2.xml']);assert.deepEqual(result.files['xl/styles.xml'],original.files['xl/styles.xml']);assert.equal(analyze(result.bytes).sheets[0].tables[0].name,'매출_2026');assert.equal(original.sheets[0].tables.length,0);
});

test('표 생성은 빈 머리글·중복·자동 이름과 같은 시트의 여러 표를 처리한다',()=>{
  let model=analyze(fixture());model=patchCells(model,[{sheetIndex:0,address:'A1',value:'Name'},{sheetIndex:0,address:'B1',value:'Name'},{sheetIndex:0,address:'C1',value:''}]).model;const first=patchAddTable(model,{sheetIndex:0,range:{top:0,bottom:0,left:0,right:2}}).model;assert.equal(first.sheets[0].tables[0].name,'Table1');assert.deepEqual(first.sheets[0].tables[0].columns,['Name','Name2','Column3']);const second=patchAddTable(first,{sheetIndex:0,range:{top:0,bottom:1,left:4,right:4}}).model;assert.equal(second.sheets[0].tables[1].name,'Table2');assert.equal(second.sheets[0].tables[1].columns[0],'Column1');assert.notEqual(second.sheets[0].tables[0].path,second.sheets[0].tables[1].path);assert.equal(analyze(second.bytes).sheets[0].tables.length,2);
});

test('기존 자동 필터와 같은 범위를 표로 만들면 조건과 native 숨김을 옮긴다',()=>{
  const original=analyze(fixture()),filter={top:0,bottom:2,left:0,right:2,columns:[{column:1,values:['42'],blanks:false}]},filtered=patchAutoFilters(original,[{sheetIndex:0,filter}]).model;const result=patchAddTable(filtered,{sheetIndex:0,range:{top:0,bottom:2,left:0,right:2},name:'Filtered'}).model;assert.equal(result.sheets[0].autoFilter,null);assert.deepEqual(result.sheets[0].tables[0].autoFilter.columns,filter.columns);assert.deepEqual(result.sheets[0].nativeHiddenRows,filtered.sheets[0].nativeHiddenRows);assert.equal(result.sheets[0].tables[0].autoFilter.bottom,2);
});

test('표 생성의 잘못된 범위·이름·충돌은 원본을 바꾸지 않고 거부한다',()=>{
  const original=analyze(fixture());for(const name of ['1bad','has space','A1','XFD1048576','R1C1','R','C','bad/name'])assert.throws(()=>patchAddTable(original,{sheetIndex:0,range:{top:0,bottom:1,left:0,right:2},name}),/표 이름/,name);for(const range of [{top:-1,bottom:1,left:0,right:1},{top:0,bottom:1048576,left:0,right:0},{top:0,bottom:100000,left:0,right:0},{top:1,bottom:0,left:0,right:0}])assert.throws(()=>patchAddTable(original,{sheetIndex:0,range}),/범위를 선택/);assert.equal(original.sheets[0].tables.length,0);const table=patchAddTable(original,{sheetIndex:0,range:{top:0,bottom:1,left:0,right:2},name:'Sales'}).model;assert.throws(()=>patchAddTable(table,{sheetIndex:0,range:{top:0,bottom:1,left:4,right:4},name:'sales'}),/표 이름/);assert.throws(()=>patchAddTable(table,{sheetIndex:0,range:{top:1,bottom:2,left:1,right:3},name:'Other'}),/겹치는/);
});

test('표 생성은 수식 머리글·활성 spill·부분 필터 겹침과 병합 영역을 거부한다',()=>{
  const original=analyze(fixture({formula:true}));assert.throws(()=>patchAddTable(original,{sheetIndex:0,range:{top:1,bottom:2,left:0,right:2}}),/수식 대신/);const spill=patchCells(analyze(fixture()),[{sheetIndex:0,address:'B2',value:'',formula:'SEQUENCE(1,2)'}]).model;assert.throws(()=>patchAddTable(spill,{sheetIndex:0,range:{top:0,bottom:1,left:0,right:1}}),/동적 배열/);const filtered=patchAutoFilters(analyze(fixture()),[{sheetIndex:0,filter:{top:0,bottom:2,left:0,right:2,columns:[]}}]).model;assert.throws(()=>patchAddTable(filtered,{sheetIndex:0,range:{top:0,bottom:1,left:0,right:2}}),/기존 자동 필터/);const merged=patchMergedCells(analyze(fixture()),[{sheetIndex:0,ranges:[{top:0,bottom:0,left:0,right:1}]}]).model;assert.throws(()=>patchAddTable(merged,{sheetIndex:0,range:{top:0,bottom:1,left:0,right:2}}),/병합/);
});

test('표 생성은 기존 차트의 관계·도형·캐시 바이트를 보존한다',()=>{
  const original=analyze(fixtureWithCachedChart()),model=patchAddTable(original,{sheetIndex:0,range:{top:0,bottom:2,left:3,right:4},name:'ChartData'}).model;assert.equal(model.writable,true,model.readOnlyReason);assert.equal(model.sheets[0].tables[0].name,'ChartData');assert.deepEqual(model.sheets[0].tables[0].range,{top:0,bottom:2,left:3,right:4});assert.equal(model.sheets[0].charts.length,1);for(const part of ['xl/drawings/drawing1.xml','xl/drawings/_rels/drawing1.xml.rels','xl/charts/chart1.xml'])assert.deepEqual(model.files[part],original.files[part],part);const relations=new TextDecoder().decode(model.files['xl/worksheets/_rels/sheet1.xml.rels']);assert.match(relations,/Id="rId3"[^>]+Type="[^"]+\/drawing"/);assert.match(relations,/Type="[^"]+\/table"/);
});

test('표 이름은 정의된 이름과 충돌하지 않고 머리글은 XML 특수문자·줄바꿈·최대 길이를 보존한다',()=>{
  const files=unzipSync(fixture()),encode=text=>new TextEncoder().encode(text);files['xl/workbook.xml']=encode(new TextDecoder().decode(files['xl/workbook.xml']).replace('</workbook>','<definedNames><definedName name="Table1">Data!$A$1</definedName></definedNames></workbook>'));let model=analyze(zipSync(files));assert.throws(()=>patchAddTable(model,{sheetIndex:0,range:{top:0,bottom:1,left:0,right:2},name:'table1'}),/표 이름/);model=patchCells(model,[{sheetIndex:0,address:'A1',value:'A & "B"\nnext'},{sheetIndex:0,address:'B1',value:'x'.repeat(255),kind:'string'},{sheetIndex:0,address:'C1',value:'x'.repeat(255),kind:'string'}]).model;const result=patchAddTable(model,{sheetIndex:0,range:{top:0,bottom:0,left:0,right:2}}).model;assert.equal(result.sheets[0].tables[0].name,'Table2');assert.deepEqual(result.sheets[0].tables[0].columns,['A & "B"\nnext','x'.repeat(255),'x'.repeat(254)+'2']);assert.equal(result.sheets[0].cells.get('A1').value,'A & "B"\nnext');
});
