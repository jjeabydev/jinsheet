'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { splitDelimitedText } = require('../src/text-to-columns');

test('텍스트 나누기는 따옴표 안 구분자와 이중 따옴표를 보존한다', () => {
  assert.deepEqual(splitDelimitedText('one,"two, too","say ""yes"""', ','), ['one', 'two, too', 'say "yes"']);
});

test('텍스트 나누기는 탭·여러 글자 구분자와 연속 구분자 옵션을 처리한다', () => {
  assert.deepEqual(splitDelimitedText('a\tb\tc', '\t'), ['a', 'b', 'c']);
  assert.deepEqual(splitDelimitedText('a<->b<->c', '<->'), ['a', 'b', 'c']);
  assert.deepEqual(splitDelimitedText('a,,b,,,c', ',', true), ['a', 'b', 'c']);
  assert.deepEqual(splitDelimitedText('a,,b', ',', false), ['a', '', 'b']);
});

test('텍스트 나누기는 잘못된 구분자와 닫히지 않은 인용 필드를 거부한다', () => {
  assert.throws(() => splitDelimitedText('a,b', ''), /구분자/);
  assert.throws(() => splitDelimitedText('a,b', '\n'), /구분자/);
  assert.throws(() => splitDelimitedText('"a,b', ','), /닫히지 않은/);
  assert.throws(() => splitDelimitedText('a,b', ',', 1), /옵션/);
});
