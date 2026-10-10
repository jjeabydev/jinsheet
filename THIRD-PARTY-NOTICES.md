# 타사 라이선스 고지

확장 내부 Webview 렌더러 번들에는 다음 타사 프로젝트가 포함됩니다. 원본 라이선스 전문은 해당 npm 패키지와 VSIX 번들에 포함합니다.

- `pdfjs-dist` 4.10.38 — Apache-2.0
- `docx-preview` 0.4.1 — Apache-2.0
- `@file-viewer/pptx` 3.1.2 — Apache-2.0
- `jszip` 3.10.2 — MIT (dual licensed with GPL-3.0-or-later)
- `dompurify` 3.4.16 — Apache-2.0 OR MPL-2.0
- `tinycolor2` 1.6.0 — MIT
- `utif` 3.1.0 — MIT
- `dingbat-to-unicode` 1.0.2 — BSD-2-Clause

렌더링은 외부 Office 프로그램이나 변환 프로세스를 실행하지 않습니다. 파일 파싱과 페이지·슬라이드 표시는 확장에 포함한 JavaScript/Webview 라이브러리로 로컬에서 수행합니다. PDF.js 표준 글꼴과 라이선스 고지도 VSIX 안에 포함합니다.

런타임 의존성은 다음과 같습니다.

- `fflate` 0.8.3 — MIT License
- `saxes` 6.0.0 — ISC License
- `xmlchars` 2.2.0 — MIT License (`saxes` dependency)

개발 의존성은 `@vscode/vsce` 4.0.0, `@vscode/test-electron` 3.1.0, `mocha` 12.0.3이며 각 패키지의 고지는 npm lockfile로 고정된 배포 패키지에서 확인할 수 있습니다. Marketplace 배포 전에는 전체 transitive dependency의 SPDX 및 고지 목록을 릴리스 아티팩트 기준으로 다시 검사하십시오.
