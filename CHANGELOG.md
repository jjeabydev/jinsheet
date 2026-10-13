# Changelog / 변경 기록

## 1.0.8 (2026-10-14)

### English

- Shift+Tab now moves to the previous cell without extending the selection.
- Added a clear-search button beside Find, with Escape as a keyboard shortcut; search results clear immediately and focus returns to the search field.
- Separated and labeled the cell value and search controls for easier identification.

### 한국어

- Shift+Tab은 선택 범위를 늘리지 않고 이전 셀만 선택
- 찾기 옆에 지우기 버튼을 추가하고 Escape도 지원하며, 검색어·결과를 즉시 지우고 입력창에 포커스 유지
- 셀 값 입력과 검색 컨트롤을 분리하고 각 영역에 이름 표시

## 1.0.7 (2026-10-14)

### English

- Search runs only when Find or Enter is pressed, marks results without changing cell formatting, and advances through matches with Enter.
- Arrow-key navigation keeps the selected cell visible without unnecessary jumps; dragging past the grid edge scrolls and extends the selection.

### 한국어

- 찾기 버튼이나 Enter를 눌렀을 때만 검색하고, 셀 서식을 유지한 채 결과를 표시하며 Enter로 다음 결과로 이동
- 방향키 이동 시 불필요한 화면 점프를 막고, 그리드 바깥으로 드래그하면 화면 이동과 선택 범위 확장

## 1.0.6 (2026-10-14)

### English

- Updated the Marketplace and README descriptions to match current editing and preview features.
- Replaced the outdated screenshot that showed the removed sorting control.

### 한국어

- 현재 셀 편집·미리보기 기능에 맞게 마켓플레이스와 README 설명 갱신
- 제거된 정렬 버튼이 보이던 이전 스크린샷을 최신 화면으로 교체

## 1.0.5 (2026-10-14)

### English

- Removed column sorting to prevent rows from appearing out of order.
- Tab and Shift+Tab now commit cell edits and move to the next or previous cell.

### 한국어

- 행 데이터가 뒤섞여 보이는 문제를 막기 위해 열 정렬 기능 제거
- 셀 편집 중 Tab과 Shift+Tab으로 값을 저장하고 다음·이전 셀로 이동

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
