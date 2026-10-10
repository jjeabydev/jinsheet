"use strict";
const assert = require("node:assert/strict");
const test = require("node:test");
const { languageOf, message, translateDiagnostic } = require("../src/i18n");

test("Korean VS Code uses Korean UI; other display languages use English", () => {
  assert.equal(languageOf("ko"), "ko");
  assert.equal(languageOf("ko-KR"), "ko");
  assert.equal(languageOf("en-US"), "en");
  assert.equal(languageOf("ja"), "en");
  assert.equal(message("insert-row", "en"), "+ Row");
  assert.equal(message("insert-row", "ko"), "+ 행");
  assert.equal(message("filter-hint", "en"), "Filter this column… (Enter to apply)");
  assert.equal(message("filter-hint", "ko"), "이 열의 필터… (Enter: 적용)");
  assert.equal(message("auto-fit-column", "en"), "Auto-fit column");
  assert.equal(message("auto-fit-column", "ko"), "열 너비 자동 맞춤");
});

test("common read-only and parser diagnostics are localized without leaking Korean into English", () => {
  const formula = "수식이 포함되어 있어 저장할 수 없습니다. 수식과 캐시값은 보기 전용입니다.";
  assert.equal(translateDiagnostic(formula, "ko"), formula);
  assert.match(translateDiagnostic(formula, "en"), /contains formulas/);
  assert.doesNotMatch(translateDiagnostic("알 수 없는 한국어 진단", "en"), /[가-힣]/);
  assert.match(translateDiagnostic("구분자를 확실히 판별할 수 없습니다. 쉼표 또는 탭 형식으로 저장된 파일인지 확인하세요.", "en"), /delimiter/);
  assert.match(translateDiagnostic("정렬 및 필터는 사용 행 100,000개 이하에서만 지원합니다.", "en"), /100,000 used rows/);
  assert.equal(translateDiagnostic("셀 값 편집", "en"), "Edit cell value");
  assert.equal(translateDiagnostic("열 너비 변경", "en"), "Resize column");
});
