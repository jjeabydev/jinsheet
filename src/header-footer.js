'use strict';

const escape = value => String(value ?? '').replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&#39;');
const sides = ['left', 'center', 'right'];
const variants = ['odd', 'even', 'first'];
const fields = { P: 'page', N: 'pages', F: 'fileName', A: 'sheetName', Z: 'path', D: 'date', T: 'time' };
const aliases = { page: 'P', pages: 'N', file: 'F', tab: 'A', path: 'Z', date: 'D', time: 'T' };

function defaultHeaderFooter() {
  return { differentFirst: false, differentOddEven: false, scaleWithDoc: true, alignWithMargins: true, ...Object.fromEntries(variants.map(variant => [variant, { header: { left: '', center: '', right: '' }, footer: { left: '', center: '', right: '' } }])) };
}

function tokens(text) {
  if (typeof text !== 'string' || text.length > 32767) throw new RangeError('머리글·바닥글 문구가 유효하지 않습니다.');
  const result = []; let index = 0;
  while (index < text.length) {
    const start = index;
    if (text[index] !== '&') {
      while (index < text.length && text[index] !== '&') index++;
      result.push({ type: 'text', value: text.slice(start, index), raw: text.slice(start, index) }); continue;
    }
    let code = text[index + 1], length = 2;
    const alias = /^&\[([a-z]+)\]/i.exec(text.slice(index));
    if (alias) { code = aliases[alias[1].toLowerCase()]; length = alias[0].length; }
    if (code === '&') result.push({ type: 'text', value: '&', raw: '&&' });
    else if (['L', 'C', 'R'].includes(code)) result.push({ type: 'section', side: { L: 'left', C: 'center', R: 'right' }[code], raw: text.slice(index, index + length) });
    else if (fields[code]) {
      const offset = code === 'P' ? /^[+-]\d+/.exec(text.slice(index + length))?.[0] || '' : '';
      if (offset && Math.abs(Number(offset)) > 9999) throw new RangeError('지원하지 않는 머리글·바닥글 코드가 있습니다.');
      length += offset.length;
      result.push({ type: 'field', name: fields[code], offset: Number(offset) || 0, raw: text.slice(index, index + length) });
    } else if (['B', 'I', 'U', 'E', 'S', 'X', 'Y'].includes(code)) result.push({ type: 'toggle', code, raw: text.slice(index, index + length) });
    else if (code === '"') {
      const end = text.indexOf('"', index + 2);
      if (end < 0) throw new RangeError('지원하지 않는 머리글·바닥글 코드가 있습니다.');
      const font = text.slice(index + 2, end), comma = font.lastIndexOf(','), family = comma < 0 ? font : font.slice(0, comma), style = comma < 0 ? 'Regular' : font.slice(comma + 1);
      if (!/^[\p{L}\p{N} _.-]{1,100}$/u.test(family) || !['Regular', 'Bold', 'Italic', 'Bold Italic'].includes(style)) throw new RangeError('지원하지 않는 머리글·바닥글 코드가 있습니다.');
      length = end - index + 1; result.push({ type: 'font', family: family === '-' ? null : family, bold: style.includes('Bold'), italic: style.includes('Italic'), raw: text.slice(index, index + length) });
    } else if (code === 'K') {
      const color = /^&K([0-9a-f]{6})/i.exec(text.slice(index));
      if (!color) throw new RangeError('지원하지 않는 머리글·바닥글 코드가 있습니다.');
      length = color[0].length; result.push({ type: 'color', value: '#' + color[1], raw: color[0] });
    } else if (/\d/.test(code || '')) {
      const size = /^&(\d+(?:\.\d+)?)/.exec(text.slice(index));
      if (!size || Number(size[1]) < 1 || Number(size[1]) > 409) throw new RangeError('지원하지 않는 머리글·바닥글 코드가 있습니다.');
      length = size[0].length; result.push({ type: 'size', value: Number(size[1]), raw: size[0] });
    } else throw new RangeError('지원하지 않는 머리글·바닥글 코드가 있습니다.');
    index += length;
  }
  return result;
}

function parseHeaderFooterText(text) {
  const sections = { left: '', center: '', right: '' }; let side = 'center';
  if (text.length > 255) throw new RangeError('머리글·바닥글은 각각 제어 코드 포함 255자 이하로 입력하세요.');
  for (const token of tokens(text)) { if (token.type === 'section') side = token.side; else sections[side] += token.raw; }
  return sections;
}

function encodeHeaderFooterText(sections) {
  if (!sections || Array.isArray(sections) || !sides.every(side => typeof sections[side] === 'string')) throw new RangeError('머리글·바닥글 문구가 유효하지 않습니다.');
  let text = '';
  for (const side of sides) {
    if (tokens(sections[side]).some(token => token.type === 'section')) throw new RangeError('머리글·바닥글 구역 안에는 &L·&C·&R 대신 &&를 사용하세요.');
    if (sections[side]) text += (side === 'center' && !sections.left && !sections.right ? '' : '&' + { left: 'L', center: 'C', right: 'R' }[side]) + sections[side];
  }
  if (text.length > 255) throw new RangeError('머리글·바닥글은 각각 제어 코드 포함 255자 이하로 입력하세요.');
  return text;
}

function validateHeaderFooter(settings) {
  if (!settings || Array.isArray(settings) || !['differentFirst', 'differentOddEven', 'scaleWithDoc', 'alignWithMargins'].every(key => typeof settings[key] === 'boolean')) throw new RangeError('머리글·바닥글 설정이 유효하지 않습니다.');
  for (const variant of variants) for (const part of ['header', 'footer']) encodeHeaderFooterText(settings[variant]?.[part]);
  return settings;
}

function renderHeaderFooterText(text, context, budget = null) {
  const style = { bold: false, italic: false, underline: false, double: false, strike: false, sup: false, sub: false, size: 11, color: null, family: null }; let html = '';
  for (const token of tokens(text)) {
    if (token.type === 'section') throw new RangeError('머리글·바닥글 문구가 유효하지 않습니다.');
    if (token.type === 'toggle') {
      const key = { B: 'bold', I: 'italic', U: 'underline', E: 'double', S: 'strike', X: 'sup', Y: 'sub' }[token.code]; style[key] = !style[key];
      if (key === 'sup' && style.sup) style.sub = false; if (key === 'sub' && style.sub) style.sup = false;
      continue;
    }
    if (token.type === 'font') { style.family = token.family; style.bold = token.bold; style.italic = token.italic; continue; }
    if (token.type === 'color') { style.color = token.value; continue; }
    if (token.type === 'size') { style.size = token.value; continue; }
    const value = token.type === 'field' ? token.name === 'page' ? Number(context.page) + token.offset : context[token.name] ?? '' : token.value;
    if (String(value) === '') continue;
    if (budget) {
      budget.characters += String(value).length;
      if (budget.characters > 1000000) throw new RangeError('머리글·바닥글 출력이 너무 큽니다. 문구를 줄이세요.');
    }
    const css = ['font-size:' + (style.size / 11) + 'em'];
    if (style.bold) css.push('font-weight:700'); if (style.italic) css.push('font-style:italic');
    const decoration = [style.underline || style.double ? 'underline' : '', style.strike ? 'line-through' : ''].filter(Boolean).join(' ');
    if (decoration) css.push('text-decoration:' + decoration); if (style.double) css.push('text-decoration-style:double');
    if (style.sup || style.sub) css.push('vertical-align:' + (style.sup ? 'super' : 'sub'));
    if (style.color) css.push('color:' + style.color); if (style.family) css.push('font-family:"' + style.family + '"');
    html += '<span style="' + escape(css.join(';')) + '">' + escape(value) + '</span>';
  }
  return html;
}

module.exports = { defaultHeaderFooter, parseHeaderFooterText, encodeHeaderFooterText, validateHeaderFooter, renderHeaderFooterText };
