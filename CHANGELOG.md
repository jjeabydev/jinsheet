# Changelog / 변경 기록

## Unreleased
- Open and edit LibreOffice workbooks with standard custom document properties; preserve the custom properties part and worksheet row/column outline groups when saving.

## 1.2.0 / 2026-10-18
- Open and edit LibreOffice-resaved workbooks that use standard on/off boolean spellings, omit the optional table-parts count, and emit neutral conditional-format defaults; preserve table, validation, conditional-format, merge, formula, chart, and frozen-pane features through save and reopen.
- Navigate to multi-digit row ranges and sort explicitly selected CSV ranges up to the 20,000-cell edit limit; rebuild CSV state once per bulk edit so sorting and other range changes do not repeatedly reconstruct the full sheet.
- Apply large XLSX cell edits by assembling sorted XML patches in one pass, keeping full 100,000-cell sorts practical to save and undo.

## 1.1.1 / 2026-10-18
- Align `SUMIF` and `AVERAGEIF` sum ranges from their top-left cell when the sum range is a different size from the criteria range, including bounded and full-row/column references.

## 1.1.0 / 2026-10-18
- Edit each icon in an XLSX custom icon-set mapping using supported standard icons; preserve thresholds and display settings, and support save, undo, and redo.
- Display and filter native XLSX custom icon mappings that reference supported standard icon sets; preserve mappings on cell edits and allow mapping, threshold, visibility, and order edits.
- Allow boundary-triggered table expansion with a saved sort state and update the stored sort ranges as the table grows.
- Preserve saved table sort state while resizing a table; adjust row ranges and remove conditions that fall outside the resized table or reference removed columns.
- Update saved table sort ranges and conditions when inserting or deleting rows and columns; remove a sort condition if its key column is deleted.
- Move XLSX column width, hidden-state, and style-definition ranges with inserted or deleted columns instead of blocking structural edits.
- Update conditional-formatting ranges and supported rule formulas during XLSX row and column insertion or deletion; shrink or remove ranges when their cells are deleted, and preserve deleted formula references as `#REF!`.
- Exclude blank cells from numeric `cellIs` conditional-format comparisons, including equality to zero; continue applying numeric rules to number cells only.
- Add native XLSX cell-value conditional rules for numeric equality, inequality, comparisons, and inclusive/exclusive ranges; matching numeric cells receive a light-red differential fill, with save, undo, redo, and recalculation support.
- Edit existing XLSX top/bottom conditional-formatting rules from Manage rules, changing direction, count or percentage, and rank with undo/redo.
- Display and preserve native XLSX top/bottom conditional-formatting rules, including rank, percentage, and tied values.
- Highlight duplicate nonblank XLSX values in a selected range with a native conditional-formatting rule; matching is case-insensitive and supports save, undo, redo, and recalculation after edits.
- Navigate to cells and rectangular ranges on another worksheet from the cell reference box with sheet-qualified references, including quoted sheet names with spaces or apostrophes.
- Navigate to workbook- or worksheet-scoped XLSX defined names from the cell reference box; sheet-scoped names take precedence and navigation switches to the target worksheet.
- Add and manage workbook- or worksheet-scoped XLSX defined names for cell ranges up to 1,000 cells; edits save with undo/redo and recalculate dependent formulas.
- VS Code locales beginning with `ko` use Korean; all other locales use English.
- Display negative XLSX data bars from the automatic zero axis, extending negative values to the left and positive values to the right.
- Hide and show XLSX worksheets from the sheet bar while keeping at least one sheet visible.
- Set per-worksheet XLSX zoom from 10% to 400%; save and restore the selected zoom level with the workbook.
- Apply or remove borders on selected XLSX cells with edge and inside presets, line styles, and color; preserve other cell formatting through save and undo/redo.
- Follow supported XLSX hyperlinks to safe external addresses or internal cells and ranges.
- Add, edit, and remove XLSX hyperlinks from the toolbar; save changes and support undo/redo.
- Localize JinSheet interface and command titles for English and Korean; use English for other VS Code locales.
- Create, replace, and clear XLSX data-validation rules on selected ranges, with native save, undo, and redo support.
- Manage existing XLSX conditional-formatting rules by priority, type, condition, and applied range; edit cell-value/formula conditions, data-bar/color-scale colors, data-bar length and display settings, three-color scale middle thresholds, and differential font/fill colors; move rules up or down; and remove individual rules with undo and redo.
- Edit existing standard XLSX icon-set thresholds, comparison boundaries, reverse order, and cell-value visibility; preserve the selected icon style and support save, undo, and redo.
- Change the style of an existing standard XLSX icon set; keep thresholds when the old and new sets have the same number of icons, otherwise reset them to evenly spaced percentages; support save, undo, and redo.
- Apply any of the 17 supported standard XLSX icon sets to a selected range with percent-based thresholds, save, undo, and redo.
- Insert or delete rows and columns in supported chart workbooks; update chart references, cached values, and chart positions attached to worksheet cells.
- Rename worksheets in chart workbooks and update namespaced chart source formulas while preserving cached values and drawing placement.
- Reorder sheets in chart workbooks while preserving chart formulas, caches, and relationships.
- Calculate the classic Excel LOOKUP vector function for ascending numeric or text vectors, with an optional matching result vector.
- Insert or delete worksheet rows and columns outside merged cells, updating merge references and rejecting operations that cross a merged range.
- Insert or delete worksheet rows and columns with hyperlinks: move linked cells and internal targets with their references, remove links attached to deleted cells, and reject deleting an internal link target.

### English

- Highlight repeated nonblank values in a selected XLSX range with a light-red native conditional-formatting rule. Matching ignores case, skips blanks, updates after edits, and supports save, undo, and redo (100,000-cell range limit).

- Use sheet-qualified references such as `'Sales 2026'!A1:B5` in the cell reference box to select a range on another visible worksheet.

- Enter a defined name in the XLSX cell reference box to select its range, including ranges on another worksheet. Worksheet-scoped names take precedence over workbook-scoped names.

- Manage XLSX defined names from the **Names** control: add a named range, change its reference, or delete it; choose workbook or worksheet scope. Changes support save, undo, redo, and dependent formula recalculation.

- Apply native XLSX data bars to a selected range in blue, green, or orange. New rules take the highest priority and save, undo, and redo with the workbook.
- Create native XLSX color-scale rules for selected ranges with red–yellow–green, green–yellow–red, or white–blue palettes; preserve priority order and support save, undo, and redo.
- Clear conditional-formatting rules from selected XLSX cells while retaining unaffected parts of each rule; save, undo, and redo the change.
- Manage existing XLSX conditional-formatting rules by priority, type, condition, and range; edit cell-value/formula conditions, data-bar/color-scale colors and data-bar length/display settings, the middle threshold of three-color scales, and differential font/fill colors; move rules up or down in priority; and remove them with undo and redo support. Differential-style edits clone shared styles so other rules keep their formatting.
- Change an existing standard XLSX icon-set style from rule management. Thresholds are preserved when the icon count matches; a different count resets them to evenly spaced percentages. Save, undo, and redo are supported.
- Insert or delete rows and columns in supported chart workbooks. Update chart source references and caches, and move or resize cell-anchored charts with the worksheet structure.

- Add, edit, remove, and follow XLSX hyperlinks for external destinations and internal cells or ranges. Save link changes and support undo/redo.
- Row and column structure changes now move hyperlink anchors and internal destinations with their cells, remove hyperlinks attached to deleted cells, and protect live internal link targets from deletion.
- Follow the VS Code display language for English or Korean; fall back to English for other locales, including command titles.
- Create, replace, and clear native XLSX data-validation rules for lists, whole numbers, decimals, dates, times, and text lengths; preserve unaffected ranges and support save, undo, and redo.

- Read the workbook Normal-style font, measure its maximum digit width in the webview, and use it for XLSX column width conversion, grid display, resizing, printing, and chart geometry.

- Added row and column outline brackets to show expanded group boundaries and nested levels in the headers.

- CSV and TSV row and column groups now include +/− controls on summary headers, matching the XLSX group workflow. Group and hide state remains session-only for these formats.

- Added Find and Replace controls beside sheet search. Replace updates the current matching cell and Replace All updates the current sheet; both replace literal case-insensitive matches, preserve formatting, skip formula cells, and support one-step undo and redo.

- Added custom multi-key sorting for selected CSV, TSV, and XLSX ranges. Specify up to 64 unique columns in priority order, choose ascending or descending per level, and preserve header, expanded-range, hidden-row, formula, formatting, undo, and redo behavior.


- Added native XLSX row and column outline groups with up to seven levels, toolbar actions to group, ungroup, collapse and expand the selected range, and undo/redo. Preserve imported outline settings and keep group visibility separate from manual hiding through save and reopen.
- Added −/+ controls to group summary row and column headers. Each control collapses or expands every nested group sharing that summary position; toolbar actions remain available.
- Added row and column outline level buttons. Selecting a level updates all group states to that depth, and groups sharing one summary position change together.

- Row and column structure edits update supported workbook- and sheet-scoped defined names that refer to cell ranges. Deleting a named cell preserves the name as `#REF!` and clears formula caches so they recalculate.

- Row and column insertions/deletions now adjust supported data-validation ranges and references in validation formulas across sheets. Rules with no remaining targets are removed, and references to deleted cells become `#REF!`; Chromium coverage checks list suggestions after save and reopen.

- Expanded native XLSX shared formulas with relative references for calculation and preserved their groups during unrelated edits. Editing group cells individually or through a pasted range saves the changed cells and expands untouched members as independent formulas; row and column structure edits flatten remaining groups and update references. Malformed groups are read-only, and array formula cells remain protected.

- Added native icon-set conditional formatting and per-column icon filters. Standard icon IDs now drive the grid indicator, filter visibility, SUBTOTAL, and XLSX save/reopen; custom icon definitions remain unsupported.
- Corrected traffic-light icon colors so high, middle, and low values display green, yellow, and red in the standard order; added Chromium coverage that keeps the cell values visible.

- Create native XLSX tables from the selected range, using its first row as text headers. Generate unique names for blank or duplicate headers, migrate an exactly matching worksheet filter, preserve body values/formulas and cell formatting, and support undo/redo and save/reopen. Reject formula headers, overlapping tables/merged cells/dynamic-array results, partial filter overlaps, and ranges above 100,000 cells.


- Added native dynamic relative-date and annual month/quarter filters plus above/below-average conditions. Shared one calculation time across visibility and TODAY/SUBTOTAL, refreshed saved bounds, preserved filters through structure edits, and added deferred midnight/focus refresh without edit history. Clean relative-date documents refresh cached bounds on actual saves, Save As and backups; read-only copies remain preserved.

- Propagated the 1904 epoch through supported date creation/components, weeks, month/workday movement, day counts, date-based financial functions and date/time TEXT. Preserved sparse holiday/cash-flow ranges, validation and conditional-format formula contexts, and supported final-day times through date bounds. Verified date-filter, save/reopen and undo behavior.

- Added native year/month/day/hour/minute/second date-group filters, mixed individual values/blanks, explicit apply and restored settings, shared serial matching and SUBTOTAL visibility. Aligned numeric date display with 1900/1904 epochs and corrected leap-day time/day rollover; external date paste and date-entry shortcuts now honor the workbook epoch.

- Added top/bottom item and percentage filters with numeric-only populations, inclusive ties, rounded-up percentage counts, native `top10`/cutoff refresh, table-total exclusion and shared SUBTOTAL visibility. Added explicit filter controls with validation/retry, restore, clear, undo and structural preservation.

- Added table-aware Tab/Shift+Tab navigation and confirmed row append from the final data cell, including pending cell/formula edits, calculated columns, formatting, totals relocation and one undo. Preserved range cycling, hidden columns and input recovery on failure; guarded conflicting data, filtered-out new rows and stale responses.

- Linked table header edits and multi-header pastes to native column names and structured formulas across cells, calculated columns, totals and worksheets. Added blank/duplicate name handling, literal formula headers, atomic validation retries and undo; preserved column IDs, filters and styles. Fixed first-cell writes into newly created empty sheets.

- Added optional table expansion for contiguous boundary input and paste, with unique headers, data formatting, calculated-column filling and one undoable transaction. Preserved pasted values, validation retry atomicity and stale-edit protection; excluded totals-row append, chart workbooks and overlapping ranges. Fixed live numeric calculation in styled blank cells.

- Added native calculated columns with single-formula propagation, exception preservation, cell-only editing, explicit column apply/replace/detach, and inserted-row/table-growth filling. Reanchored templates after first-row edits, preserved data styles, validated the full operation, and fixed copy translation for cell-like sheet names.

- Added table totals on/off and explicit per-column aggregate controls with saved native metadata, custom formula restoration, filtering/manual-hide semantics, request acknowledgments and retry feedback. Supported escaped/Unicode structured references and fixed mixed table-column XML handling and structural updates of table formula metadata.

- Enabled vertical resizing of totals tables by moving totals cells and formatting, updating direct/range references, defined names and attached reference formulas, and preserving data outside the resized table. Added destination/feature conflict handling and save/reopen/undo coverage. Fixed explicit null-formula insertion and unquoted Unicode sheet references.

- Supported native table filters that exclude totals rows across filtering, clear, save/reopen, row/column structure changes and column resizing. Enabled implicit structured references in totals formulas, normalized earlier inclusive filter ranges, and requested full recalculation after saved filter visibility changes. Protected direct vertical resizing of totals tables with guidance to insert/delete data rows.

- Saved native XLSX filtered-row visibility and restored rows when clearing criteria after reopening. Preserved overlapping manual hiding with ignorable compatibility metadata, handled imported active filters, and retained fixed chart dimensions through visibility changes. Direct Excel metadata round-trip preservation remains unverified.

- Added XLSX numeric/text custom filters with two-condition AND/OR, wildcard escaping, apply/Enter confirmation, validation feedback and column-only clear. Shared matching now drives visibility and SUBTOTAL, including Boolean formula results. Preserved criteria through structural updates and fixed fast copy/paste clipboard ordering and grid/input focus during redraws.

- Corrected XLSX default dimensions and OOXML width conversion, preserving 1–1785px columns across grid, print, chart placement and save. Zero-pixel columns can be unhidden. Adapted cell padding/borders and vertical alignment for small cells, and kept narrow-cell editors visible. CSV sizing retains its existing defaults.

- Added three saved chart placement modes with UI selection and undo/redo. Fixed-size drawings now preserve their rectangle during dimension edits and AutoFit, including unsaved widths in print export. Shared drawings use per-worksheet copies when changed.

- Expanded printed charts to 32 series and 10,000 points per series (50,000 source values per chart), with full series legends and additional colors. Pie/doughnut charts draw every supported slice and explicitly summarize long category legends. Bounded chart markup generation and verified large-chart geometry in Chromium.

- Added basic chart printing at saved drawing coordinates, with print-area/page clipping, repeating titles, hidden dimensions and live source data. Added an export-only chart toggle and explicit messages for unsupported geometry or data limits. Fixed chart insertion into worksheets beginning with an XML declaration.

- Added manual row/column page breaks with header markers, reset, save/undo and structural updates. Print export honors breaks with titles and page fields, follows page order, and ignores manual breaks in fit-to-pages mode. Stored automatic break records remain preserved during ordinary edits.

- Added saved XLSX header/footer sections and first/odd/even page variants, dynamic fields and common text formatting, custom first page numbering, global area page totals, and browser margin fitting with overflow feedback. Unsupported header/footer contents remain preserved during ordinary edits.

- Added saved XLSX repeating print rows/columns with structural updates, undo and per-page layout. Titles outside the print area, hidden dimensions, merged-title boundaries and repeated-output limits are handled.

- Added saved XLSX print areas with set/clear controls, undo, structural reference updates and sheet rename/reorder support. Sheet printing honors up to 16 existing areas and whole row/column references; unreadable areas require explicit selection.

- Added undoable XLSX page setup and paginated HTML export for browser printing or PDF output, including selections, unsaved formulas, cell styles, merged cells and hidden rows/columns. Toolbar and chart panels now limit their height to keep the grid usable in smaller windows.

- Automatic row heights now follow content, formatting, formula, and column width changes without adding undo steps. Pending measurements are included in save, Save As, and backups; background layout updates preserve active cell editors.

- Added row header boundary dragging, multi-row resize previews, keyboard resizing, cancel support and double-click AutoFit.
- Added Wrap Text and measured row AutoFit with atomic undo and redo, manual-height preservation, and stale-result rejection. Simple workbooks without a styles part now gain one when formatting is saved.
- Added row height editing and reset in points, including efficient full-worksheet height changes and undo and redo.
- Added display of saved XLSX row heights and wrapped text with matching scrolling, selection, merged-cell and frozen-row geometry.
- Added undoable insertion of column, line, area, scatter, pie, and doughnut charts from selected XLSX data ranges, including worksheets that already contain supported charts.
- Added editing of chart titles and series names, with undo and redo support.
- Fixed chart previews to reflect unsaved source edits and refresh saved caches when chart formulas depend on edited cells outside the source range; empty charts keep their editing controls.
- Added undoable chart deletion with package cleanup and preservation of chart and drawing parts shared by other sheets.
- Added changing existing chart source ranges from the grid selection, preserving titles and series formatting and supporting undo and redo.

### 한국어

- XLSX 열을 삽입·삭제할 때 너비·숨김·스타일 정의 범위도 함께 이동·조정해 구조 변경을 허용.

- 선택 XLSX 범위에서 비어 있지 않은 중복 값을 연한 빨간색 native 조건부 서식으로 강조. 대소문자를 구분하지 않고 빈 셀을 제외하며 값 변경 후 다시 계산하고 저장·실행 취소·다시 실행을 지원 (100,000셀 범위 제한).

- 이름 상자에서 `'Sales 2026'!A1:B5` 같은 시트 지정 주소로 다른 표시 시트의 셀이나 범위를 선택.

- XLSX 이름 상자에 정의된 이름을 입력해 해당 범위를 선택하고 대상 시트로 이동. 시트 범위 이름은 통합 문서 범위 이름보다 우선.

- **이름 관리**에서 XLSX 셀 범위의 이름을 추가·수정·삭제하고 통합 문서 또는 시트 범위를 선택. 최대 1,000셀 참조를 지원하며 저장·실행 취소·다시 실행과 이름을 참조하는 수식 재계산을 지원.
- VS Code 언어 코드가 `ko`로 시작하면 한국어를 사용하고, 그 외 모든 언어에서는 영어를 사용.

- 선택 범위에 파랑·초록·주황 native XLSX 데이터 막대를 적용. 새 규칙을 가장 높은 우선 순위로 추가하고 저장·실행 취소·다시 실행을 지원.
- 선택 범위에 빨강–노랑–초록, 초록–노랑–빨강, 흰색–파랑 native XLSX 색조 규칙을 추가. 기존 규칙 우선 순서를 보존하고 저장·실행 취소·다시 실행을 지원.
- 조건부 서식 규칙에서 셀 값 비교·수식 조건, 데이터 막대·색조 색상, 데이터 막대 길이·셀 값 표시·그라데이션/단색 채우기 옵션, 세 색상 색조의 가운데 기준, 셀 값·수식 규칙의 차등 글꼴색·채우기색을 편집하고, 우선순위를 한 단계 올리거나 내리거나 개별 규칙을 제거. 공유 스타일은 복제해 다른 규칙의 서식을 보존하며 저장·실행 취소·다시 실행을 지원.
- 규칙 관리에서 기존 표준 아이콘 집합 스타일을 변경. 아이콘 개수가 같으면 기준을 유지하고, 다르면 균등한 백분율 기준으로 다시 설정하며 저장·실행 취소·다시 실행을 지원.
- 지원되는 차트 통합 문서에서 행·열을 삽입하거나 삭제할 때 차트 원본 참조·캐시와 셀에 연결된 차트 위치를 함께 갱신.

- XLSX에서 선택 셀의 하이퍼링크를 추가·편집·제거하고 외부 주소나 내부 셀·범위로 이동. 저장·실행 취소·다시 실행 지원.
- JinSheet 메시지와 명령 이름은 VS Code의 영어·한국어 언어를 따르고, 다른 언어에서는 영어를 사용.
- 선택 범위에 목록·정수·소수·날짜·시간·텍스트 길이 데이터 유효성 규칙을 만들거나 교체하고 지우는 기능 추가. 영향받지 않은 적용 범위를 보존하며 저장·실행 취소·다시 실행 지원.

- 통합 문서 Normal 스타일 글꼴의 최대 숫자 폭을 웹뷰에서 측정해 XLSX 열 너비 변환, 그리드 표시, 크기 조정, 인쇄와 차트 배치에 적용합니다.

- 펼쳐진 행·열 그룹의 범위와 중첩 수준을 머리글의 개요 괄호로 표시.

- CSV·TSV 행·열 그룹 요약 머리글에 +/− 접기·펼치기 버튼을 추가해 XLSX와 같은 조작 흐름을 제공. CSV·TSV의 그룹·숨김 상태는 현재 세션에만 유지.

- 시트 검색창 옆에 **바꾸기**와 **모두 바꾸기**를 추가. 현재 일치 셀 또는 현재 시트 전체에서 대소문자를 구분하지 않는 문자열을 바꾸고 서식을 보존한다. 수식 셀은 건너뛰며 실행 취소·다시 실행을 지원.

- 선택한 CSV·TSV·XLSX 범위를 최대 64개 열의 다중 기준으로 정렬. 기준 우선순위마다 오름차순·내림차순을 지정하고, 머리글·확장 범위·숨김 행·수식·서식과 실행 취소·다시 실행 동작을 보존.


- XLSX 행·열 선택 범위를 최대 7단계의 개요 그룹으로 만들고, 도구 모음에서 그룹화·해제·접기·펼치기를 실행하도록 추가. 기존 그룹 설정을 보존하고 그룹 숨김과 수동 숨김을 구분해 저장·재개방하며 실행 취소·다시 실행을 지원.
- 그룹 요약 행·열 머리글에 −/+ 버튼을 추가. 같은 요약 위치를 공유하는 중첩 그룹 전체를 함께 접거나 펼치며 도구 모음 조작을 유지.
- 행·열 개요 수준 숫자 버튼을 추가. 선택한 수준에 맞춰 그룹을 접거나 펼치며 같은 요약 위치를 공유하는 그룹은 함께 바뀜.

- 행·열 구조 변경 시 지원하는 통합 문서·시트 범위 이름의 참조를 갱신. 이름이 가리키는 단일 셀을 삭제하면 이름을 `#REF!`로 보존하고 수식 캐시를 비워 다시 계산.

- 행·열 삽입·삭제에서 데이터 유효성 적용 범위와 다른 시트까지 포함한 유효성 수식 참조를 갱신. 적용 대상이 남지 않는 규칙은 제거하고, 삭제된 셀 참조는 `#REF!`로 변환. 저장·재개방 후 목록 제안을 확인하는 Chromium 검사를 추가.

- XLSX 공유 수식 그룹을 셀별 상대 참조로 펼쳐 계산하고 다른 셀 수정 시 원본 그룹을 보존. 그룹 셀을 직접 수정하거나 범위 붙여넣기하면 바뀐 셀을 저장하고 미수정 구성원은 독립 수식으로 펼침. 행·열 구조 변경도 남은 공유 그룹을 먼저 풀어 참조를 갱신. 잘못된 그룹은 저장 가능한 파일로 취급하지 않으며, 배열 수식 셀은 편집·구조 변경에서 보호.

- 표준 아이콘 집합 조건부 서식 표시와 열별 아이콘 필터를 추가. 아이콘 번호를 그리드 표시·필터 숨김·SUBTOTAL·XLSX 저장/재개방에 공통 적용하고 사용자 지정 아이콘 정의는 미지원.
- 교통 신호등 집합의 색 순서를 바로잡아 높은 값은 초록, 중간은 노랑, 낮은 값은 빨강으로 표시하고 셀 값이 그대로 보이는 Chromium 회귀 검사를 추가.

- 선택 범위에서 첫 행을 텍스트 머리글로 쓰는 native XLSX 표 생성 추가. 빈·중복 머리글 이름 보정, 범위가 일치하는 워크시트 필터 이전, 데이터 행의 값·수식과 셀 서식 보존, Undo/Redo·저장/재개방을 지원. 수식 머리글, 기존 표·병합·동적 배열 결과 겹침, 필터 부분 겹침과 100,000셀 초과 범위를 거부.
- XLSX 표에 Light 1–21, Medium 1–28, Dark 1–11 기본 스타일 선택을 추가하고 행·열 줄무늬와 첫째·마지막 열 옵션을 native `tableStyleInfo`에 저장. 화면 미리보기를 연결하되 직접 셀 서식과 조건부 서식은 우선 유지. 새 표는 `TableStyleMedium2`로 생성.


- 표시된 머리글이 있는 표의 마지막 데이터 행 삭제를 지원. 빈 표·필터·총계·계산 열 정의와 Undo/Redo·저장·재입력을 연결하고, 계산 열 템플릿의 미래 데이터 행 상대 참조를 실제 셀 삭제 참조와 구분. 절대·외부·머리글·문자열 참조는 기존 규칙 유지. 머리글 없는 표의 마지막 행 보호 유지.

- 머리글만 있는 빈 표·단일 셀 ref·빈 표 필터와 구조화 집계를 지원. 첫 입력·머리글 Tab으로 계산 열을 채우고 빈 합계 켜기/끄기·첫 행 추가 시 총계 이동·Undo·재개방 연결. 머리글 서식의 데이터 행 복사와 빈 표의 총계 이동에 따른 계산 열 템플릿 이동을 방지. 필터 취소 후 Tab 재시도 검증에서 불필요한 셀 클릭을 제거하고 일반 편집 요청이 섞이지 않음을 검사.

- 표 마지막 셀의 Tab으로 추가할 행이 필터에 가려지면 대상 표의 필터 해제 후 추가 또는 필터 유지를 선택하도록 개선. 입력·필터·행·계산 열·총계 이동을 한 번의 Undo로 복원하고 취소 시 입력 유지. 수동 숨김·다른 필터 보존, 유효성 경고 연속 처리와 문서 변경 후 오래된 재시도 차단 검증.

- XLSX ISO 날짜 셀(t="d")의 날짜·시각·시간대·분수초를 1900/1904 숫자 날짜로 해석해 표시·수식·날짜 필터·합계에 연결. 잘못된 달력 날짜·시각을 거부하고 무관한 수정·동일 값 서식 변경·구조 이동에서 원래 ISO 표현 보존. 값 변경은 기존 서식을 유지한 숫자로 저장.

- 통합문서 테마·시스템 lastClr·HSL 밝기 조정과 표준/사용자 인덱스 팔레트를 표시·테두리·조건부 서식·색상 필터에 공통 적용. 기존 테마/팔레트 XML과 일치하는 DXF를 보존하며 저장·재개방·Undo·합계 연동 검증. 잘못된 색 정의·누락한 테마·시스템 인덱스 64/65는 추측하지 않고 거부.

- RGB 채우기색·글자색 필터와 채우기 없음·자동 글자색, 명시적 적용·설정 복원·열별 해제 추가. 조건부 서식·색상 척도의 표시색을 화면·SUBTOTAL·native 숨김에 공유하고 DXF 저장·Undo·Redo·재개방을 검증. 256색 목록·직접 선택을 제공하며 셀 서식은 유지. 동적 배열 결과의 기존 셀 서식 보존을 보정. 테마색·인덱스색·아이콘은 후속 범위.

- native 상대 날짜·매년 월/분기 필터와 평균 초과·미만 조건 추가. 표시·TODAY·SUBTOTAL의 계산 시각을 통일하고 저장 기준값·구조 변경 보존, 편집을 방해하지 않는 자정/포커스 복귀 갱신을 연결. 상대 날짜 문서는 실제 저장·다른 이름 저장·백업 때 수정 상태와 무관하게 기준값을 갱신하며 읽기 전용 복사 보존.

- 지원되는 날짜 생성·분해·주차·월/근무일 이동·날짜 차이·날짜 기반 금융 함수와 날짜/시간 TEXT에 1904 체계를 전달. 희소 휴일·현금 흐름 범위와 유효성·조건부 서식 수식 문맥을 유지하고 마지막 지원 날짜의 오후 경계값을 보정. 날짜 필터·저장/재개방·Undo 흐름 검증.

- native 연·월·일·시·분·초 날짜 그룹과 개별 값·빈 셀 혼합, 명시적 적용·설정 복원, 날짜 serial 공통 판정·SUBTOTAL 연동 추가. 숫자 날짜 표시의 1900/1904 체계·윤일 시간·일 경계 반올림을 보정하고 외부 날짜 붙여넣기·오늘 날짜 입력에 통합문서 체계를 반영.

- 상위·하위 개수/백분율 필터와 숫자 모집단·동점 포함·백분율 개수 올림, native `top10` 기준값 갱신·표 합계 제외·SUBTOTAL 공통 표시 판정 추가. 명시적 적용·오류 재시도·설정 복원·열별 해제·Undo·구조 변경 보존을 연결.

- 표 데이터의 Tab/Shift+Tab 행 이동과 마지막 데이터 셀의 성공 확인 후 새 행 추가 구현. 편집값·계산 열·서식·합계 이동을 단일 Undo로 묶고 범위 순환·숨긴 열·실패 입력 복구를 유지하며 아래 데이터 충돌·필터 숨김·오래된 응답 보호 추가.

- 표 머리글 입력·여러 머리글 붙여넣기를 native 열 이름과 셀·계산 열·합계·다른 시트의 구조화 참조에 연결. 빈 이름·중복 이름·수식 머리글 텍스트 처리와 유효성 경고 재시도·단일 Undo를 추가하고 열 ID·필터·서식 유지. 새 빈 시트의 첫 셀 저장 오류도 수정.

- 표 경계의 연속 입력·붙여넣기 자동 확장과 끄기 설정, 고유 머리글·데이터 서식·계산 열 채우기 및 단일 Undo 추가. 붙여넣은 값·유효성 경고 재시도의 원자성을 유지하고 오래된 계획의 덮어쓰기를 방지. 합계 아래 입력·차트 통합문서·범위 겹침은 제외하며, 서식이 있는 빈 셀의 숫자 입력 직후 계산도 보정.

- native 계산 열의 단일 수식 전파·예외 보존·한 셀만 수정·열 적용/전체 교체/해제, 새 행·표 확대 자동 채우기 추가. 첫 행 변경의 기준 수식·데이터 서식·전체 유효성을 보정하고 셀 주소처럼 보이는 시트 이름의 복사 참조 이동 수정.

- 표 합계 행 켜기/끄기와 열별 집계 적용 UI, native 메타데이터 저장, 사용자 수식 복구, 필터·수동 숨김 반영, 요청 응답·재시도 안내 추가. 특수문자·한글 구조 참조와 혼합 형태의 표 열 XML, 구조 변경 시 표 수식 메타데이터 갱신 보정.

- 총계 표의 세로 크기 조정에서 총계 셀·서식 이동, 직접/범위 참조·이름 정의·부속 참조 수식 갱신, 표 밖 데이터 보존 지원. 목적지·부속 기능 충돌 안내와 저장·재개방·Undo 검증 추가. 새 셀의 명시적 null 수식 처리와 따옴표 없는 한글 시트 참조 보정.

- 총계 행을 제외하는 native 표 필터의 적용·지우기·저장·재개방·행/열 구조 변경·열 크기 조정 지원. 총계 수식의 암시적 구조 참조 계산, 이전 총계 포함 필터 정규화, 저장된 필터 표시 변경 후 전체 재계산 요청 추가. 총계 표의 직접 세로 크기 조정은 데이터 행 삽입·삭제 안내와 함께 보호.

- XLSX 필터 제외 행의 native 숨김 상태 저장과 재개방 후 조건 해제 시 행 복구 추가. 호환성 메타데이터로 겹치는 수동 숨김을 보존하고 외부 활성 필터를 분류하며, 표시 변경 시 고정 차트 크기 유지. 실제 Excel 왕복 저장에서 메타데이터 보존은 직접 미검증.

- XLSX 숫자·텍스트 조건 필터와 두 조건 AND/OR, 와일드카드 이스케이프, 적용·Enter 확정, 실패 시 입력 유지, 열별 지우기 추가. 화면·SUBTOTAL·Boolean 수식에 같은 조건 판정 적용. 구조 변경에서 조건을 보존하고 빠른 복사·붙여넣기 경합 및 재렌더링 포커스 보정.

- XLSX 기본 셀 치수·OOXML 너비 변환을 보정하고 1~1785px 열을 그리드·인쇄·차트·저장에 일관되게 반영. 0픽셀 열의 숨김 해제 지원. 작은 셀의 여백·테두리·세로 정렬과 좁은 셀 편집창을 보정. CSV 치수 기본값 유지.

- 차트 배치 세 가지 선택 UI·저장·Undo 추가. 고정 차트의 셀 크기 변경·자동 맞춤 좌표를 갱신하고 미저장 열 너비도 인쇄에 반영. 공유 그리기 파일은 수정 대상 시트별로 분리.

- 차트 인쇄를 32계열·계열당 10,000개, 차트당 합계 50,000개로 확장. 전체 계열 범례·추가 색상을 표시하고 원형·도넛의 모든 조각을 그리며 긴 범례의 생략 수를 명시. 차트 마크업 생성 한도와 Chromium 대형 차트 배치 검증 추가.

- 저장된 차트 위치에 기본 모양으로 인쇄하고 영역·페이지 경계, 반복 제목, 숨김 행·열과 미저장 원본 값을 반영. 내보내기 전용 차트 포함 옵션과 미지원 위치·데이터 한도 안내 추가. XML 선언으로 시작하는 워크시트의 차트 추가 실패 수정.

- 수동 행·열 페이지 나누기와 머리글 표시·초기화·저장·Undo·구조 변경 참조 갱신 추가. 제목·쪽 번호와 함께 출력하고 인쇄 순서를 반영하며 페이지 수 맞춤에서는 수동 나누기를 무시. 일반 편집에서 저장된 자동 나누기 정보 보존

- XLSX 머리글·바닥글 구역과 첫 쪽·홀짝 구분, 동적 필드·기본 텍스트 서식, 시작 쪽 번호·여러 영역의 전체 쪽 수와 브라우저 여백 맞춤·넘침 안내 추가. 일반 편집에서 미지원 머리글·바닥글 내용 보존

- XLSX 반복 제목 행·열 저장과 구조 변경·Undo·페이지별 배치 추가. 인쇄 영역 밖 제목·숨김 상태·병합 경계와 반복 출력량 제한 처리

- XLSX 인쇄 영역 저장·교체·지우기·Undo와 행·열 구조 및 시트 이름·순서 변경에 따른 참조 갱신 추가. 기존 최대 16개 영역·전체 행·열 참조를 읽고 시트 인쇄에 적용하며 해석 불가 영역은 직접 선택하도록 안내

- XLSX 페이지 설정 저장·실행 취소와 브라우저 인쇄·PDF 저장용 페이지별 HTML 내보내기 추가. 선택 범위·미저장 수식·셀 서식·병합·숨김 행·열 반영. 작은 창에서 그리드 공간을 유지하도록 도구 모음·차트 영역 높이 제한

- 내용·서식·수식·열 너비 변경 후 자동 행 높이 지속 재계산 추가. 별도 Undo 단계 없이 저장·다른 이름 저장·백업에 반영하고 배경 높이 갱신 시 입력 중 셀과 포커스 보존

- 행 번호 경계 드래그·선택 행 미리보기·키보드 높이 조절·취소와 더블클릭 자동 맞춤 추가
- 텍스트 줄 바꿈 적용·해제와 실제 글꼴 기반 행 자동 맞춤, 서식·높이의 단일 실행 취소·다시 실행, 수동 높이 보존과 오래된 측정 결과 거부 추가
- 스타일 파일이 없는 단순 XLSX에 서식 저장 시 기본 스타일 파일 생성 지원
- 선택 행의 높이(pt) 적용·기본 높이 복원·실행 취소·다시 실행과 전체 행 기본 높이 변경 추가
- XLSX에 저장된 행 높이·줄 바꿈 표시와 스크롤·선택·병합·틀 고정 좌표 반영 추가
- 지원 차트가 이미 있는 워크시트에도 선택한 XLSX 데이터 범위로 세로 막대·선·영역·분산형·원형·도넛 차트를 삽입하고 실행 취소·다시 실행할 수 있도록 추가
- 차트 제목과 계열 이름을 편집하고 실행 취소·다시 실행할 수 있도록 추가
- 저장 전 원본 편집의 차트 미리보기 반영, 원본 범위 밖 수식 의존 셀 변경 시 저장 캐시 갱신과 빈 차트 편집 UI 유지 수정
- 차트 삭제·실행 취소·다시 실행과 패키지 정리를 추가하고 다른 시트가 공유하는 차트·그리기 파일 보존
- 선택한 셀 범위로 기존 차트 원본을 변경하고 제목·계열 서식을 보존하며 실행 취소·다시 실행할 수 있도록 추가

## 1.0.9 (2026-10-14)

### English

- Row and column structure edits update supported workbook- and sheet-scoped defined names that refer to cell ranges. Deleting a named cell preserves the name as `#REF!` and clears formula caches so they recalculate.

- Fixed mouse dragging in a cell editor so text can be selected without starting a grid selection or moving focus away from the input.

### 한국어

- 행·열 구조 변경 시 지원하는 통합 문서·시트 범위 이름의 참조를 갱신. 이름이 가리키는 단일 셀을 삭제하면 이름을 `#REF!`로 보존하고 수식 캐시를 비워 다시 계산.

- 셀 편집 입력란에서 마우스로 텍스트를 드래그해 선택할 수 있도록 수정하고, 그리드 범위 선택과 포커스 이동이 끼어들지 않게 개선

## 1.0.8 (2026-10-14)

### English

- Row and column structure edits update supported workbook- and sheet-scoped defined names that refer to cell ranges. Deleting a named cell preserves the name as `#REF!` and clears formula caches so they recalculate.

- Shift+Tab now moves to the previous cell without extending the selection.
- Added a clear-search button beside Find, with Escape as a keyboard shortcut; search results clear immediately and focus returns to the search field.
- Separated and labeled the cell value and search controls for easier identification.

### 한국어

- 행·열 구조 변경 시 지원하는 통합 문서·시트 범위 이름의 참조를 갱신. 이름이 가리키는 단일 셀을 삭제하면 이름을 `#REF!`로 보존하고 수식 캐시를 비워 다시 계산.

- Shift+Tab은 선택 범위를 늘리지 않고 이전 셀만 선택
- 찾기 옆에 지우기 버튼을 추가하고 Escape도 지원하며, 검색어·결과를 즉시 지우고 입력창에 포커스 유지
- 셀 값 입력과 검색 컨트롤을 분리하고 각 영역에 이름 표시

## 1.0.7 (2026-10-14)

### English

- Row and column structure edits update supported workbook- and sheet-scoped defined names that refer to cell ranges. Deleting a named cell preserves the name as `#REF!` and clears formula caches so they recalculate.

- Search runs only when Find or Enter is pressed, marks results without changing cell formatting, and advances through matches with Enter.
- Arrow-key navigation keeps the selected cell visible without unnecessary jumps; dragging past the grid edge scrolls and extends the selection.

### 한국어

- 행·열 구조 변경 시 지원하는 통합 문서·시트 범위 이름의 참조를 갱신. 이름이 가리키는 단일 셀을 삭제하면 이름을 `#REF!`로 보존하고 수식 캐시를 비워 다시 계산.

- 찾기 버튼이나 Enter를 눌렀을 때만 검색하고, 셀 서식을 유지한 채 결과를 표시하며 Enter로 다음 결과로 이동
- 방향키 이동 시 불필요한 화면 점프를 막고, 그리드 바깥으로 드래그하면 화면 이동과 선택 범위 확장

## 1.0.6 (2026-10-14)

### English

- Row and column structure edits update supported workbook- and sheet-scoped defined names that refer to cell ranges. Deleting a named cell preserves the name as `#REF!` and clears formula caches so they recalculate.

- Updated the Marketplace and README descriptions to match current editing and preview features.
- Replaced the outdated screenshot that showed the removed sorting control.

### 한국어

- 행·열 구조 변경 시 지원하는 통합 문서·시트 범위 이름의 참조를 갱신. 이름이 가리키는 단일 셀을 삭제하면 이름을 `#REF!`로 보존하고 수식 캐시를 비워 다시 계산.

- 현재 셀 편집·미리보기 기능에 맞게 마켓플레이스와 README 설명 갱신
- 제거된 정렬 버튼이 보이던 이전 스크린샷을 최신 화면으로 교체

## 1.0.5 (2026-10-14)

### English

- Row and column structure edits update supported workbook- and sheet-scoped defined names that refer to cell ranges. Deleting a named cell preserves the name as `#REF!` and clears formula caches so they recalculate.

- Removed column sorting to prevent rows from appearing out of order.
- Tab and Shift+Tab now commit cell edits and move to the next or previous cell.

### 한국어

- 행·열 구조 변경 시 지원하는 통합 문서·시트 범위 이름의 참조를 갱신. 이름이 가리키는 단일 셀을 삭제하면 이름을 `#REF!`로 보존하고 수식 캐시를 비워 다시 계산.

- 행 데이터가 뒤섞여 보이는 문제를 막기 위해 열 정렬 기능 제거
- 셀 편집 중 Tab과 Shift+Tab으로 값을 저장하고 다음·이전 셀로 이동

## 1.0.4

### English

- Row and column structure edits update supported workbook- and sheet-scoped defined names that refer to cell ranges. Deleting a named cell preserves the name as `#REF!` and clears formula caches so they recalculate.

- Added a screenshot of JinSheet running in VS Code and on-screen sorting instructions.

### 한국어

- 행·열 구조 변경 시 지원하는 통합 문서·시트 범위 이름의 참조를 갱신. 이름이 가리키는 단일 셀을 삭제하면 이름을 `#REF!`로 보존하고 수식 캐시를 비워 다시 계산.

- VS Code에서 실행 중인 JinSheet 스크린샷과 화면 내 정렬 안내 추가

## 1.0.3

### English

- Row and column structure edits update supported workbook- and sheet-scoped defined names that refer to cell ranges. Deleting a named cell preserves the name as `#REF!` and clears formula caches so they recalculate.

- Removed the unreliable filter controls and header-row toggle. Sorting keeps the first row as the header; choosing another header row is not supported yet.
- Preserved workbook cell colors and text contrast while selecting cells.
- Disabled **Clear sort** when no sort is active.

### 한국어

- 행·열 구조 변경 시 지원하는 통합 문서·시트 범위 이름의 참조를 갱신. 이름이 가리키는 단일 셀을 삭제하면 이름을 `#REF!`로 보존하고 수식 캐시를 비워 다시 계산.

- 동작이 불안정했던 필터와 머리글 행 토글 제거. 정렬 시 첫 행을 머리글로 고정하며 다른 행 지정은 아직 지원하지 않음
- 셀 선택 중에도 통합 문서의 채우기 색과 글자 대비 유지
- 정렬 중이 아닐 때 **정렬 해제** 버튼 비활성화

## 1.0.2

### English

- Row and column structure edits update supported workbook- and sheet-scoped defined names that refer to cell ranges. Deleting a named cell preserves the name as `#REF!` and clears formula caches so they recalculate.

- Fixed direct cell entry for English keyboard input and Korean IME composition.
- Improved readability of the in-cell editor.
- Clarified supported features in the Marketplace descriptions.

### 한국어

- 행·열 구조 변경 시 지원하는 통합 문서·시트 범위 이름의 참조를 갱신. 이름이 가리키는 단일 셀을 삭제하면 이름을 `#REF!`로 보존하고 수식 캐시를 비워 다시 계산.

- 셀을 바로 입력할 때 영문 키 입력과 한글 IME 입력이 되도록 수정
- 셀 편집 입력칸의 가독성 개선
- Marketplace 설명에 지원 기능을 명확히 안내

## 1.0.1

### English

- Row and column structure edits update supported workbook- and sheet-scoped defined names that refer to cell ranges. Deleting a named cell preserves the name as `#REF!` and clears formula caches so they recalculate.

- Simplified the README and added Korean guidance.
- Fixed a CI check that treated internal Korean translations as visible English UI.

### 한국어

- 행·열 구조 변경 시 지원하는 통합 문서·시트 범위 이름의 참조를 갱신. 이름이 가리키는 단일 셀을 삭제하면 이름을 `#REF!`로 보존하고 수식 캐시를 비워 다시 계산.

- README를 간결하게 정리하고 한국어 안내 추가
- 내부 번역 문자열을 영어 UI로 잘못 검사하던 CI 수정

## 1.0.0

### English

- Row and column structure edits update supported workbook- and sheet-scoped defined names that refer to cell ranges. Deleting a named cell preserves the name as `#REF!` and clears formula caches so they recalculate.

- First public release on the VS Code Marketplace.
- Added editing tools for CSV, TSV, and supported XLSX workbooks, including range paste, sort, filter, search, and safer read-only handling.
- Added local PDF, Word, and PowerPoint previews with zoom and fit-to-width or fit-to-height controls.

### 한국어

- 행·열 구조 변경 시 지원하는 통합 문서·시트 범위 이름의 참조를 갱신. 이름이 가리키는 단일 셀을 삭제하면 이름을 `#REF!`로 보존하고 수식 캐시를 비워 다시 계산.

- VS Code Marketplace 첫 공개
- CSV·TSV와 지원되는 XLSX의 셀 편집, 범위 붙여넣기, 정렬, 필터, 검색 및 안전한 읽기 전용 처리 추가
- PDF·Word·PowerPoint 로컬 미리보기와 확대·축소, 너비·높이 맞춤 추가
