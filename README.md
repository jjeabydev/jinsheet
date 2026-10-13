# JinSheet

[Install from the VS Code Marketplace](https://marketplace.visualstudio.com/items?itemName=jjeabydev.jinsheet) · [Report an issue](https://github.com/jjeabydev/jinsheet/issues)

JinSheet lets you edit CSV, TSV, and supported Excel workbooks, and preview Office and PDF files in VS Code. Files stay on your computer.

![JinSheet editing a spreadsheet in VS Code.](media/jinsheet-spreadsheet.png)

## Features

- Edit CSV, TSV, and supported XLSX workbooks. Search cells and paste cell ranges.
- Preview PDF pages, Word documents, and PowerPoint slides.
- Zoom pages and slides, or fit them to the viewer width or height.
- Move between cells with Tab and Shift+Tab while editing.
- Drag across cells to select a range; the grid scrolls when you drag past its edge.

## Cell editing and search

Enter a search term and press **Find** or **Enter** to search the current sheet. Matching cells are marked with an outline that preserves their fill and text formatting. Press **Enter** again to move to the next match. Typing alone does not run a search.

While editing a cell, **Tab** saves the value and moves to the next cell, and **Shift+Tab** saves it and moves to the previous cell. **Enter** saves and moves down; **Escape** cancels the edit. Sorting is not currently supported.

## File support

| Format | View | Edit and save |
| --- | --- | --- |
| CSV, TSV | Yes | Yes |
| XLSX | Yes | Supported workbooks |
| XLSM | Yes | No |
| PDF, DOCX, DOCM, PPTX, PPTM | Yes | No |
| XLS, ODS, DOC, PPT | No | No |

Unsupported XLSX workbooks open read-only. Formulas are shown but not recalculated. Word and PowerPoint previews may differ from Microsoft Office; complex animations, transitions, and media playback are not supported.

## 한국어

JinSheet는 VS Code에서 CSV·TSV와 지원되는 Excel 통합 문서를 편집하고 Office·PDF 파일을 미리 보는 확장입니다. 파일은 PC에 남습니다.

- CSV·TSV와 지원되는 XLSX 편집: 셀 검색, 범위 붙여넣기
- PDF 페이지, Word 문서, PowerPoint 슬라이드 미리보기
- 페이지·슬라이드 확대/축소와 뷰어 너비·높이에 맞춤

검색어를 입력하고 **찾기** 버튼이나 **Enter**를 눌러 현재 시트에서 찾습니다. 일치하는 셀은 채우기 색과 글꼴을 바꾸지 않고 테두리로 표시합니다. **Enter**를 다시 누르면 다음 결과로 이동합니다. 입력만으로는 검색하지 않습니다.

셀 편집 중 **Tab**을 누르면 값을 저장하고 다음 셀로, **Shift+Tab**을 누르면 이전 셀로 이동합니다. **Enter**는 저장 후 아래 셀로 이동하고 **Escape**는 편집을 취소합니다. 셀을 드래그해 범위를 선택할 때 그리드 가장자리를 지나면 화면이 해당 방향으로 스크롤됩니다. 정렬은 현재 지원하지 않습니다.

지원 범위를 벗어난 XLSX는 읽기 전용으로 열립니다. 수식은 표시하지만 다시 계산하지 않습니다. Word·PowerPoint 화면은 Microsoft Office와 다를 수 있으며 복잡한 애니메이션·전환·미디어 재생은 지원하지 않습니다.

## License / 라이선스

MIT. See [LICENSE](LICENSE) and [third-party notices / 타사 라이선스 고지](THIRD-PARTY-NOTICES.md).
