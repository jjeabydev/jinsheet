'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { defaultHeaderFooter, parseHeaderFooterText, encodeHeaderFooterText, validateHeaderFooter, renderHeaderFooterText } = require('../src/header-footer');

test('머리글 구역과 리터럴 앰퍼샌드는 Excel 코드를 보존해 왕복한다', () => {
  const raw = '&LLeft && L&C&P / &N&R&A';
  const parsed = parseHeaderFooterText(raw);
  assert.deepEqual(parsed, { left: 'Left && L', center: '&P / &N', right: '&A' });
  assert.equal(encodeHeaderFooterText(parsed), raw);
  assert.deepEqual(parseHeaderFooterText('&&L literal'), { left: '', center: '&&L literal', right: '' });
  assert.equal(encodeHeaderFooterText({ left: '', center: 'x'.repeat(255), right: '' }).length, 255);
});

test('동적 머리글 필드는 쪽 번호·전체 수·파일·시트·경로·날짜·시간으로 안전하게 출력한다', () => {
  const context = { page: 7, pages: 12, fileName: '<img src=x>', sheetName: 'Sheet & One', path: '/tmp/', date: '2026-10-16', time: '17:00' };
  const html = renderHeaderFooterText('&P+2 / &N — &F &A &Z &D &T && &[Page]', context);
  assert.match(html, />9</); assert.match(html, />12</); assert.match(html, /&lt;img src=x&gt;/); assert.match(html, /Sheet &amp; One/); assert.match(html, /2026-10-16/); assert.doesNotMatch(html, /<img/);
  assert.match(html, />7</);
});

test('머리글의 기본 글꼴·색상·강조·위아래 첨자 코드를 렌더링한다', () => {
  const html = renderHeaderFooterText('&"Arial,Bold Italic"&14&K123456Title&U under&U&X2&X&Y3&Y', {});
  assert.match(html, /font-family:&quot;Arial&quot;/); assert.match(html, /font-weight:700/); assert.match(html, /font-style:italic/); assert.match(html, /color:#123456/); assert.match(html, /text-decoration:underline/); assert.match(html, /vertical-align:super/); assert.match(html, /vertical-align:sub/);
});

test('지원하지 않는 코드·과도한 길이·구역 삽입과 출력 폭주를 거부한다', () => {
  for (const raw of ['&G', '&Q', '&"Broken', '&K01+050', '&9999', 'x'.repeat(256)]) assert.throws(() => parseHeaderFooterText(raw));
  assert.throws(() => encodeHeaderFooterText({ left: '&Cwrong', center: '', right: '' }));
  assert.throws(() => renderHeaderFooterText('&F', { fileName: 'file' }, { characters: 999999 }), /출력/);
  const config = defaultHeaderFooter(); config.differentFirst = 'true'; assert.throws(() => validateHeaderFooter(config));
  const first = defaultHeaderFooter(), second = defaultHeaderFooter(); first.odd.header.left = 'changed'; assert.equal(second.odd.header.left, ''); assert.equal(first.even.header.left, '');
});
