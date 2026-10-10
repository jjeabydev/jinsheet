'use strict';

const messages = {
  en: {
    'insert-row': '+ Row', 'delete-row': '− Row', 'insert-column': '+ Column', 'delete-column': '− Column',
    'insert-row-title': 'Insert above selected row', 'delete-row-title': 'Delete selected row',
    'insert-column-title': 'Insert left of selected column', 'delete-column-title': 'Delete selected column',
    'cell-value': 'Cell value or formula', 'search-sheet': 'Find in sheet', 'find-hint': 'Find… (Ctrl/Cmd+F)',
    'filter-hint': 'Filter this column… (Enter to apply)', 'apply-filter': 'Filter', 'clear-view': 'Clear sort/filter', 'header-row': 'Header row',
    'sort-title': 'Click to sort · Shift-click to add a sort key · Ctrl/Cmd-click to remove', 'filter-active': 'Filter active', 'view-row-limit': 'Sorting and filtering are limited to sheets with at most 100,000 used rows.',
    'spreadsheet': 'Spreadsheet', 'document-preview': 'Document preview', 'modified': 'Modified', 'resize-column': 'Resize column ',
    'edit-shortcut': 'Edit cell: F2 · Clear selected cell: Delete/Backspace',
    'scroll-rows': 'Scroll to load more rows.', 'hidden': 'hidden', 'no-data': 'No data to display.',
    'cache-value': 'cached value', 'matches': 'matches', 'match-separator': ' ', 'no-matches': 'No matches',
    'render-error': 'Could not render document: ', 'pdf-loading': 'Loading PDF pages…',
    'pdf-load': 'Loading PDF… ', 'pdf-analyzing': 'Analyzing PDF…', 'pdf-layout': 'Preparing pages…',
    'pdf-compat': 'Retrying with PDF compatibility mode…', 'pdf-page-info': 'Loading page information…',
    'pdf-failed': 'Could not display PDF', 'zoom-out': 'Zoom out', 'zoom-in': 'Zoom in', 'fit-width': 'Fit to width',
    'pdf-worker-missing': 'The built-in PDF.js library could not be loaded.',
    'pdf-worker-timeout': 'The PDF worker did not respond within 5 seconds.',
    'pdf-fallback-timeout': 'PDF compatibility rendering did not respond within 30 seconds.',
    'pdf-page-limit': 'The PDF exceeds the 1,000-page limit.', 'ppt-worker-read': 'Could not read the PPTX worker file: HTTP ',
    'ppt-worker-invalid': 'The PPTX worker file is empty or exceeds the size limit.',
    'not-previewable': 'This document cannot be previewed.', 'aria-column': 'Column ', 'aria-row': 'Row ',
    'readonly-formula': 'This workbook contains formulas. Formula cells and cached values are view-only.',
    'readonly-unsupported': 'This workbook contains unsupported Excel features and is read-only to protect its contents.',
    'readonly-macro': 'Macro-enabled workbooks cannot be edited in JinSheet.',
    'readonly-office': 'Office documents are rendered by JinSheet and are read-only.',
    'readonly-pdf': 'PDF documents are rendered by the built-in viewer and are read-only.',
    'readonly-remote': 'Remote and virtual files are read-only because safe saving cannot be guaranteed.',
    'error-generic': 'JinSheet could not safely process this file. It has been opened read-only.',
    'error-csv-ambiguous': 'The delimiter could not be identified. Choose comma-separated or tab-separated text.',
    'error-csv-quote': 'The CSV contains malformed or unclosed quoted fields.',
    'error-pdf-header': 'A valid PDF file header could not be found.',
    'error-pdf-size': 'The PDF exceeds the 25 MiB size limit.',
    'error-unsupported': 'This file format is not supported for preview or editing.',
    'error-readonly': 'This document is read-only.',
    'error-external-change': 'The file changed outside JinSheet. Saving was stopped; reopen it or use Save As.',
    'error-cancelled': 'Saving was cancelled.', 'error-range': 'The requested cell range is invalid.',
    'error-formula': 'A range containing formula or error cells cannot be edited.',
    'error-structure': 'This document format does not support row or column editing.', 'page': 'Page ',
    'edit-cell': 'Edit cell value', 'paste-range': 'Paste cell range', 'insert-structure': 'Insert rows/columns', 'delete-structure': 'Delete rows/columns', 'resize-column-history': 'Resize column', 'auto-fit-column': 'Auto-fit column', 'auto-fit-column-title': 'Fit the selected column to its contents'
  },
  ko: {
    'insert-row': '+ 행', 'delete-row': '− 행', 'insert-column': '+ 열', 'delete-column': '− 열',
    'insert-row-title': '선택한 행 위에 삽입', 'delete-row-title': '선택한 행 삭제',
    'insert-column-title': '선택한 열 왼쪽에 삽입', 'delete-column-title': '선택한 열 삭제',
    'cell-value': '셀 값 또는 수식', 'search-sheet': '시트에서 찾기', 'find-hint': '찾기… (Ctrl/Cmd+F)',
    'filter-hint': '이 열의 필터… (Enter: 적용)', 'apply-filter': '필터', 'clear-view': '정렬/필터 해제', 'header-row': '첫 행은 머리글',
    'sort-title': '클릭: 정렬 · Shift+클릭: 보조 정렬 추가 · Ctrl/Cmd+클릭: 해당 정렬 해제', 'filter-active': '필터 적용 중', 'view-row-limit': '정렬과 필터는 사용 행 100,000개 이하인 시트에서만 지원합니다.',
    'spreadsheet': '스프레드시트', 'document-preview': '문서 미리보기', 'modified': '수정됨', 'resize-column': '열 너비 조정: ',
    'edit-shortcut': '셀 편집: F2 · 선택 셀 비우기: Delete/Backspace', 'scroll-rows': '행은 스크롤해 불러옵니다.',
    'hidden': '숨김', 'no-data': '표시할 데이터가 없습니다.', 'cache-value': '캐시값', 'matches': '개 일치', 'match-separator': '',
    'no-matches': '일치 항목 없음', 'render-error': '문서를 렌더링할 수 없습니다: ', 'pdf-loading': 'PDF 페이지를 불러오는 중…',
    'pdf-load': 'PDF 불러오는 중… ', 'pdf-analyzing': 'PDF 분석 중…', 'pdf-layout': '페이지 구성 불러오는 중…',
    'pdf-compat': 'PDF 호환 모드로 다시 여는 중…', 'pdf-page-info': '페이지 정보 불러오는 중…',
    'pdf-failed': 'PDF 표시 실패', 'zoom-out': '축소', 'zoom-in': '확대', 'fit-width': '너비에 맞춤',
    'pdf-worker-missing': '내장 PDF.js 라이브러리를 불러오지 못했습니다.',
    'pdf-worker-timeout': 'PDF Worker가 5초 안에 응답하지 않습니다.', 'pdf-fallback-timeout': 'PDF 호환 렌더링이 30초 안에 응답하지 않습니다.',
    'pdf-page-limit': 'PDF 페이지 수 제한(1,000)을 초과했습니다.', 'ppt-worker-read': 'PPTX Worker 파일을 읽지 못했습니다: HTTP ',
    'ppt-worker-invalid': 'PPTX Worker 파일이 비어 있거나 크기 제한을 초과했습니다.', 'not-previewable': '이 문서는 미리보기를 지원하지 않습니다.',
    'aria-column': '열 ', 'aria-row': '행 ', 'readonly-formula': '수식이 포함되어 있어 저장할 수 없습니다. 수식과 캐시값은 보기 전용입니다.',
    'readonly-unsupported': '지원하지 않는 Excel 기능이 포함되어 있어 내용을 보호하기 위해 읽기 전용으로 엽니다.',
    'readonly-macro': '매크로 통합문서는 JinSheet에서 수정할 수 없습니다.',
    'readonly-office': 'Office 문서는 JinSheet 내부 렌더러로 표시하며 읽기 전용입니다.',
    'readonly-pdf': 'PDF 문서는 내장 렌더러로 표시하며 읽기 전용입니다.',
    'readonly-remote': '원격·가상 파일 시스템에서는 안전한 저장을 보장할 수 없어 읽기 전용으로 엽니다.',
    'error-generic': '파일을 안전하게 처리하지 못해 읽기 전용으로 열었습니다.',
    'error-csv-ambiguous': '구분자를 확실히 판별할 수 없습니다. 쉼표 또는 탭 형식인지 확인하세요.',
    'error-csv-quote': 'CSV 인용 필드가 잘못되었거나 닫히지 않았습니다.', 'error-pdf-header': '유효한 PDF 파일 헤더를 찾을 수 없습니다.',
    'error-pdf-size': 'PDF 크기 제한(25 MiB)을 초과했습니다.', 'error-unsupported': '이 파일 형식은 미리보기 또는 편집을 지원하지 않습니다.',
    'error-readonly': '읽기 전용 문서입니다.', 'error-external-change': 'JinSheet 외부에서 파일이 변경되어 저장을 중단했습니다. 다시 열거나 다른 이름으로 저장하세요.',
    'error-cancelled': '저장이 취소되었습니다.', 'error-range': '요청한 셀 범위가 유효하지 않습니다.',
    'error-formula': '수식 또는 오류 셀이 포함된 범위는 수정할 수 없습니다.', 'error-structure': '이 문서 형식은 행·열 편집을 지원하지 않습니다.', 'page': '페이지 ',
    'edit-cell': '셀 값 편집', 'paste-range': '셀 범위 붙여넣기', 'insert-structure': '행·열 삽입', 'delete-structure': '행·열 삭제', 'resize-column-history': '열 너비 변경', 'auto-fit-column': '열 너비 자동 맞춤', 'auto-fit-column-title': '선택한 열의 내용에 맞춰 너비 조정'
  }
};

function languageOf(language) { return String(language || '').toLowerCase().startsWith('ko') ? 'ko' : 'en'; }
function message(key, language) { return messages[languageOf(language)][key] || messages.en[key] || key; }
function messagesFor(language) { return messages[languageOf(language)]; }
function translateDiagnostic(value, language) {
  const text = String(value || '');
  if (languageOf(language) === 'ko' || !/[가-힣]/.test(text)) return text;
  const exact = new Map([
    ['수식이 포함되어 있어 저장할 수 없습니다. 수식과 캐시값은 보기 전용입니다.', 'readonly-formula'],
    ['정렬 및 필터는 사용 행 100,000개 이하에서만 지원합니다.', 'view-row-limit'],
    ['매크로 통합문서는 JinSheet에서 수정할 수 없습니다.', 'readonly-macro'],
    ['원격·가상 파일 시스템에서는 안전한 저장을 보장할 수 없어 읽기 전용으로 엽니다.', 'readonly-remote'],
    ['Office 문서는 JinSheet 내부 렌더러로 표시하며 읽기 전용입니다.', 'readonly-office'],
    ['PDF 문서는 내장 렌더러로 표시하며 읽기 전용입니다.', 'readonly-pdf'],
    ['구분자를 확실히 판별할 수 없습니다. 쉼표 또는 탭 형식으로 저장된 파일인지 확인하세요.', 'error-csv-ambiguous'],
    ['PDF 파일이 비어 있거나 잘렸습니다.', 'error-pdf-header'], ['PDF 크기 제한(25 MiB)을 초과했습니다.', 'error-pdf-size'],
    ['유효한 PDF 파일 헤더를 찾을 수 없습니다.', 'error-pdf-header'], ['읽기 전용 파일입니다.', 'error-readonly'],
    ['저장이 취소되었습니다.', 'error-cancelled'],
    ['파일이 외부에서 변경되어 저장을 중단했습니다. 다시 열거나 다른 이름으로 저장하세요.', 'error-external-change'],
    ['조회 범위가 유효하지 않습니다.', 'error-range'], ['수식·오류 셀이 포함된 범위는 수정할 수 없습니다.', 'error-formula'],
    ['이 문서 형식은 행·열 구조를 편집할 수 없습니다.', 'error-structure'],
    ['셀 값 편집', 'edit-cell'], ['셀 범위 붙여넣기', 'paste-range'], ['행·열 삽입', 'insert-structure'], ['행·열 삭제', 'delete-structure'], ['열 너비 변경', 'resize-column-history']
  ]);
  if (exact.has(text)) return message(exact.get(text), language);
  if (/^문서 미리보기를 열 수 없습니다:/.test(text)) return `Could not open document preview. ${message('error-generic', language)}`;
  if (/^PDF 미리보기를 열 수 없습니다:/.test(text)) return `Could not open PDF preview. ${message('error-generic', language)}`;
  if (/수식/.test(text)) return message('readonly-formula', language);
  if (/외부에서 변경/.test(text)) return message('error-external-change', language);
  if (/지원하지 않는|저장할 수 없습니다|읽기 전용/.test(text)) return message('readonly-unsupported', language);
  if (/구분자/.test(text)) return message('error-csv-ambiguous', language);
  if (/인용|따옴표/.test(text)) return message('error-csv-quote', language);
  return message('error-generic', language);
}
module.exports = { languageOf, message, messagesFor, translateDiagnostic };
