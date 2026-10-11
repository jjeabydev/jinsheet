# Changelog / 변경 기록

## 1.0.4

### English

- Added a screenshot of JinSheet running in VS Code and on-screen sorting instructions.

### 한국어

- VS Code에서 실행 중인 JinSheet 스크린샷과 화면 내 정렬 안내 추가

## 1.0.3

### English

- Removed the unreliable filter controls and header-row toggle. Sorting keeps the first row as the header; choosing another header row is not supported yet.
- Preserved workbook cell colors and text contrast while selecting cells.
- Disabled **Clear sort** when no sort is active.

### 한국어

- 동작이 불안정했던 필터와 머리글 행 토글 제거. 정렬 시 첫 행을 머리글로 고정하며 다른 행 지정은 아직 지원하지 않음
- 셀 선택 중에도 통합 문서의 채우기 색과 글자 대비 유지
- 정렬 중이 아닐 때 **정렬 해제** 버튼 비활성화

## 1.0.2

### English

- Fixed direct cell entry for English keyboard input and Korean IME composition.
- Improved readability of the in-cell editor.
- Clarified supported features in the Marketplace descriptions.

### 한국어

- 셀을 바로 입력할 때 영문 키 입력과 한글 IME 입력이 되도록 수정
- 셀 편집 입력칸의 가독성 개선
- Marketplace 설명에 지원 기능을 명확히 안내

## 1.0.1

### English

- Simplified the README and added Korean guidance.
- Fixed a CI check that treated internal Korean translations as visible English UI.

### 한국어

- README를 간결하게 정리하고 한국어 안내 추가
- 내부 번역 문자열을 영어 UI로 잘못 검사하던 CI 수정

## 1.0.0

### English

- First public release on the VS Code Marketplace.
- Added editing tools for CSV, TSV, and supported XLSX workbooks, including range paste, sort, filter, search, and safer read-only handling.
- Added local PDF, Word, and PowerPoint previews with zoom and fit-to-width or fit-to-height controls.

### 한국어

- VS Code Marketplace 첫 공개
- CSV·TSV와 지원되는 XLSX의 셀 편집, 범위 붙여넣기, 정렬, 필터, 검색 및 안전한 읽기 전용 처리 추가
- PDF·Word·PowerPoint 로컬 미리보기와 확대·축소, 너비·높이 맞춤 추가
