'use strict';

const MAX_CELL_CHARS = 32767;
const MAX_COLUMNS = 16384;

function splitDelimitedText(value, delimiter, mergeConsecutive = false) {
  if (typeof value !== 'string' || value.length > MAX_CELL_CHARS) throw new RangeError('텍스트 나누기 셀 값이 유효하지 않습니다.');
  if (typeof delimiter !== 'string' || delimiter.length < 1 || delimiter.length > 8 || /[\r\n"]/.test(delimiter)) throw new RangeError('구분자는 줄 바꿈·따옴표가 아닌 1~8자여야 합니다.');
  if (typeof mergeConsecutive !== 'boolean') throw new TypeError('연속 구분자 옵션이 유효하지 않습니다.');

  const fields = [];
  let field = '', quoted = false;
  for (let index = 0; index < value.length;) {
    const character = value[index];
    if (character === '"') {
      if (quoted && value[index + 1] === '"') { field += '"'; index += 2; continue; }
      if (quoted) quoted = false;
      else if (!field.length) quoted = true;
      else field += character;
      index++;
      continue;
    }
    if (!quoted && value.startsWith(delimiter, index)) {
      fields.push(field); field = ''; index += delimiter.length;
      if (mergeConsecutive) while (value.startsWith(delimiter, index)) index += delimiter.length;
      if (fields.length >= MAX_COLUMNS) throw new RangeError('나눈 결과가 워크시트 열 제한을 초과합니다.');
      continue;
    }
    field += character;
    index++;
  }
  if (quoted) throw new Error('닫히지 않은 텍스트 한정자가 있습니다. 따옴표를 확인하세요.');
  fields.push(field);
  if (fields.length > MAX_COLUMNS) throw new RangeError('나눈 결과가 워크시트 열 제한을 초과합니다.');
  return fields;
}

module.exports = { splitDelimitedText };
