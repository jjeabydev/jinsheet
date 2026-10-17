"use strict";
const assert = require("node:assert/strict");
const test = require("node:test");
const { languageOf, message, messagesFor, translateDiagnostic } = require("../src/i18n");
const { ICON_SET_LENGTHS } = require("../src/auto-filter");

test("Korean VS Code uses Korean UI; other display languages use English", () => {
  assert.equal(languageOf("ko"), "ko");
  assert.equal(languageOf("ko-KR"), "ko");
  assert.equal(languageOf("ko_KR"), "en");
  assert.equal(languageOf("en-US"), "en");
  assert.equal(languageOf("ja"), "en");
  assert.equal(languageOf("fr-CA"), "en");
  assert.equal(languageOf("zh-CN"), "en");
  assert.equal(languageOf("kok-IN"), "en");
  assert.equal(languageOf(undefined), "en");
  assert.equal(message("insert-row", "en"), "+ Row");
  assert.equal(message("insert-row", "ko"), "+ 행");
  assert.equal(message("chart-scatter", "ko"), "분산형 (X Y)");
  assert.equal(message("row-height", "ko"), "행 높이(pt)");
  assert.equal(message("wrap-text", "ko"), "텍스트 줄 바꿈");
  assert.match(translateDiagnostic("시트가 변경되어 행 자동 맞춤을 다시 실행해야 합니다.", "en"), /changed during measurement/);
  assert.equal(translateDiagnostic("행 높이 변경", "en"), "Change row height");
  assert.match(translateDiagnostic("행 높이는 한 번에 10,000행 이하 또는 전체 행에 적용하세요.", "en"), /10,000 rows/);
  assert.match(translateDiagnostic("분산형 차트의 첫 번째 열에는 숫자 X 값을 선택하세요.", "en"), /numeric X values/);
  assert.match(translateDiagnostic("분산형 차트에는 숫자 X 값이 필요합니다.", "en"), /numeric X values/);
  assert.equal(message("auto-fit-column", "en"), "Auto-fit column");
  assert.equal(message("auto-fit-column", "ko"), "열 너비 자동 맞춤");
  assert.equal(message("sort-ascending", "en"), "Sort A→Z");
  assert.equal(message("sort-ascending", "ko"), "오름차순");
  assert.equal(message("conditional-format-rule-move-up", "en"), "Move up");
  assert.equal(message("conditional-format-rule-move-up", "ko-KR"), "우선순위 올리기");
  assert.equal(message("conditional-format-move-down-history", "en"), "Lower conditional-formatting rule priority");
  assert.equal(message("conditional-format-rule-edit-condition", "en"), "Edit condition");
  assert.equal(message("conditional-format-rule-edit-colors", "ko-KR"), "색상 편집");
  assert.equal(message("conditional-format-rule-edit-font-color", "en-US"), "Edit font color");
  assert.equal(message("conditional-format-rule-edit-fill-color", "ko-KR"), "채우기 색 편집");
  assert.equal(message("conditional-format-rule-edit-data-bar-display", "en-US"), "Edit bar display");
  assert.equal(message("conditional-format-rule-edit-icon-set", "en-US"), "Edit icon set");
  assert.equal(message("conditional-format-icon-set-style", "en-US"), "Icon style");
  assert.equal(message("conditional-format-icon-style-preserve-thresholds", "fr-FR"), "Keep the current thresholds");
  assert.equal(message("conditional-format-icon-style-reset-thresholds", "ko-KR").replace("{count}", "5"), "기준 5개를 균등한 백분율로 재설정");
  assert.match(translateDiagnostic("차트 원본 셀을 모두 삭제하면 차트를 보존할 수 없습니다.", "en"), /chart could not be updated safely/);
  assert.equal(message("conditional-format-icon-style-5Quarters", "ko-KR"), "사분원 5개");
  assert.equal(message("conditional-format-icon-threshold-type-percentile", "ko-KR"), "백분위");
  assert.equal(message("conditional-format-data-bar-min-length-prompt", "ko-KR"), "최소 막대 길이를 입력하세요 (0~100%)");
  assert.equal(message("error-conditional-format-data-bar-display", "fr-CA"), "Enter integer lengths from 0 to 100; the minimum must not exceed the maximum.");
  assert.equal(translateDiagnostic("조건부 서식 비교값이 유효하지 않습니다.", "en"), "Enter a valid numeric value for this condition.");
  assert.match(translateDiagnostic("아이콘 기준값이 유효하지 않습니다.", "en"), /supported same-sheet formula or a valid number/);
  assert.match(translateDiagnostic("아이콘 기준 수식에 지원하지 않는 함수가 있습니다.", "en"), /supported same-sheet formula or a valid number/);
  assert.equal(translateDiagnostic("데이터 막대 표시 설정이 유효하지 않습니다.", "en"), "Enter integer lengths from 0 to 100; the minimum must not exceed the maximum.");
});

test("every existing UI message is available in English and Korean", () => {
  assert.deepEqual(Object.keys(messagesFor("ko")).sort(), Object.keys(messagesFor("en")).sort());
  for (const name of Object.keys(ICON_SET_LENGTHS)) {
    const key = "conditional-format-icon-style-" + name;
    assert.notEqual(messagesFor("en")[key], key, name + " has no English label");
    assert.notEqual(messagesFor("ko")[key], key, name + " has no Korean label");
  }
});

test("common read-only and parser diagnostics are localized without leaking Korean into English", () => {
  const formula = "수식이 포함되어 있어 저장할 수 없습니다. 수식과 캐시값은 보기 전용입니다.";
  assert.equal(translateDiagnostic(formula, "ko"), formula);
  assert.match(translateDiagnostic(formula, "en"), /contains formulas/);
  assert.doesNotMatch(translateDiagnostic("알 수 없는 한국어 진단", "en"), /[가-힣]/);
  assert.match(translateDiagnostic("구분자를 확실히 판별할 수 없습니다. 쉼표 또는 탭 형식으로 저장된 파일인지 확인하세요.", "en"), /delimiter/);
  assert.equal(translateDiagnostic("셀 값 편집", "en"), "Edit cell value");
  assert.equal(translateDiagnostic("선택 범위 값 지우기", "en"), "Clear selected cells");
  assert.equal(translateDiagnostic("선택 범위 값 지우기", "ko"), "선택 범위 값 지우기");
  assert.equal(translateDiagnostic("선택 범위 정렬", "en"), "Sort selected range");
  assert.equal(translateDiagnostic("열 너비 변경", "en"), "Resize column");
  assert.match(translateDiagnostic("클립보드 범위와 선택 영역의 크기가 맞지 않습니다.", "en"), /copied block does not fit/);
  assert.match(translateDiagnostic("백만 행 전체 열 복사는 같은 XLSX 통합문서 안에서만 붙여넣을 수 있습니다.", "en"), /same XLSX workbook/);
  assert.equal(translateDiagnostic("클립보드 범위와 선택 영역의 크기가 맞지 않습니다.", "ko"), "클립보드 범위와 선택 영역의 크기가 맞지 않습니다.");
});

test('totals relocation diagnostics explain destination and reference restrictions', () => {
  assert.match(translateDiagnostic('이동할 합계 행 위치에 데이터가 있습니다. 빈 행을 선택하거나 데이터 행 삽입을 사용하세요.', 'en'), /totals row contains data/);
  assert.match(translateDiagnostic('합계 이동 범위의 병합·유효성·조건부 서식·하이퍼링크를 먼저 해제하세요.', 'en'), /validation/);
  assert.match(translateDiagnostic('3차원 시트 참조가 있어 합계 행 참조를 안전하게 이동할 수 없습니다.', 'en'), /3D sheet references/);
});
