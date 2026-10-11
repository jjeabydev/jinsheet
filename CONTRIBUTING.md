# Contributing / 기여 안내

### English

Before opening a pull request, run `npm ci`, `npm test`, `npm run check`, and `npm run test:host`, then include the results.

For data-integrity or security changes, add a regression fixture. When expanding XLSX write support, verify unsupported content stays unchanged and saved files reopen correctly. Do not add real user documents, credentials, or other sensitive data to the repository.

### 한국어

Pull request를 열기 전에 `npm ci`, `npm test`, `npm run check`, `npm run test:host`를 실행하고 결과를 남겨주세요.

데이터 무결성이나 보안에 영향을 주는 변경에는 회귀 fixture를 추가하세요. XLSX 저장 지원을 넓힐 때는 미지원 콘텐츠가 그대로 보존되고 저장한 파일을 다시 열 수 있는지 확인하세요. 실제 사용자 문서, 인증 정보 등 민감한 데이터는 저장소에 넣지 마세요.
