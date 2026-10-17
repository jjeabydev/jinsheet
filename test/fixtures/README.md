# Test fixtures / 테스트 fixture

### English

These small sample files are used to check JinSheet's editor association and previews. The XLSX samples demonstrate editing and formatting; the DOCX, PPTX, and PDF samples are read-only previews. Open them in VS Code to check the result.

`LibreOffice-Charts-Roundtrip.xlsx` contains bar, line, area, pie and doughnut charts created by JinSheet and resaved by LibreOffice. The scatter chart is covered separately because it requires numeric X-axis data. The XLSX tests verify that JinSheet can reopen this file, edit the source cells, and refresh the chart data caches.

`LibreOffice-Criteria-Functions-Reference.xlsx` contains mixed text, numeric, Boolean, blank, empty-string formula and error cells with cached `COUNTIF`, `SUMIF`, `COUNTIFS` and `SUMIFS` results recalculated by LibreOffice Calc 24.2. The tests compare wildcard matching, a tilde-escaped literal wildcard, aligned multi-condition sums and a no-match sum. Calc's `COUNTIF(range,"*")` counts nonempty numeric, Boolean and error cells in this fixture, unlike JinSheet's Excel-oriented text wildcard rule; that result is intentionally not used as an Excel expectation.

`LibreOffice-Lookup-Duplicate-Error-Reference.xlsx` contains ascending and descending lookup vectors with duplicate keys and an error cell. LibreOffice Calc 24.2 cached results cover approximate `MATCH`, `VLOOKUP` and `LOOKUP`. These results are a repeatable Calc cross-check, not a complete Excel compatibility oracle.

`LibreOffice-Styled-Roundtrip.xlsx` is `JinSheet-Styled-Sample.xlsx` opened and resaved by LibreOffice Calc 24.2. The XLSX test edits a numeric cell and checks that the workbook can be reopened with its original formatting and remaining package parts intact. It is an interoperability sample, not a Microsoft Excel compatibility oracle.

`LibreOffice-Mixed-Features-Roundtrip.xlsx` combines a native table and filter, totals formulas, numeric validation, conditional formatting, merged cells, a formula, a chart and frozen panes. JinSheet created the supported features, then LibreOffice Calc 24.2 opened and resaved the workbook. Unit and VS Code Webview host tests edit, save and reopen it. The sample records the LibreOffice XML variants JinSheet accepts; it is not a Microsoft Excel compatibility oracle.

### 한국어

이 작은 샘플 파일은 JinSheet 편집기 연결과 미리보기를 확인하는 데 사용합니다. XLSX 샘플은 편집과 서식 표시를, DOCX·PPTX·PDF 샘플은 읽기 전용 미리보기를 확인합니다. VS Code에서 파일을 열어 결과를 확인하세요.

`LibreOffice-Charts-Roundtrip.xlsx`는 JinSheet에서 만든 막대형·꺾은선형·영역형·원형·도넛형 차트를 LibreOffice에서 다시 저장한 파일입니다. 숫자 X축 데이터가 필요한 분산형 차트는 별도 테스트에서 확인합니다. XLSX 테스트는 JinSheet가 이 파일을 다시 열고, 원본 셀 편집에 따라 차트 데이터 캐시를 갱신하는지 확인합니다.

`LibreOffice-Criteria-Functions-Reference.xlsx`는 텍스트·숫자·논리값·빈 셀·빈 문자열 수식·오류 셀과 LibreOffice Calc 24.2가 재계산한 `COUNTIF`, `SUMIF`, `COUNTIFS`, `SUMIFS` 결과를 포함합니다. 테스트는 wildcard, 물결표로 escape한 literal wildcard, 정렬된 다중 조건 합계와 일치 항목이 없는 합계를 비교합니다. 이 fixture의 Calc `COUNTIF(range,"*")`는 숫자·논리값·오류 등 비어 있지 않은 셀을 세며 JinSheet의 Excel 지향 텍스트 wildcard 규칙과 다릅니다. 이 값은 Excel 기대값으로 사용하지 않습니다.

`LibreOffice-Lookup-Duplicate-Error-Reference.xlsx`는 중복 키와 오류 셀이 든 오름차순·내림차순 조회 범위 및 LibreOffice Calc 24.2가 재계산한 근사 `MATCH`, `VLOOKUP`, `LOOKUP` 캐시를 포함합니다. 반복 가능한 Calc 교차 확인 자료이며 Excel 호환 전체의 기준 oracle은 아닙니다.

`LibreOffice-Styled-Roundtrip.xlsx`는 `JinSheet-Styled-Sample.xlsx`를 LibreOffice Calc 24.2에서 열고 다시 저장한 파일입니다. XLSX 검사는 숫자 셀을 수정한 뒤 다시 열 수 있는지와 기존 서식·나머지 패키지 구성요소를 보존하는지 확인합니다. Microsoft Excel 호환성 전체를 판정하는 기준은 아닙니다.

`LibreOffice-Mixed-Features-Roundtrip.xlsx`는 native 표·필터, 합계 수식, 숫자 유효성, 조건부 서식, 병합 셀, 수식, 차트와 고정 창을 함께 포함합니다. JinSheet가 지원 기능을 만든 뒤 LibreOffice Calc 24.2에서 열고 다시 저장했습니다. 단위 검사와 VS Code Webview 호스트 검사는 이 파일을 편집·저장·재열기합니다. 이 표본은 JinSheet가 허용하는 LibreOffice XML 변형을 기록하며 Microsoft Excel 호환성 전체의 기준 oracle은 아닙니다.
