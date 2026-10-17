'use strict';
function columnNumber(address) {
  let value = 0;
  for (const letter of /^([A-Z]+)/.exec(address)[1]) value = value * 26 + letter.charCodeAt(0) - 64;
  return value;
}
function columnLetters(number) {
  let text = '';
  for (let n = number; n > 0; n = Math.floor((n - 1) / 26)) text = String.fromCharCode(65 + (n - 1) % 26) + text;
  return text;
}
function translateFormulaReferences(formula, rowDelta, columnDelta) {
  if (typeof formula !== 'string' || formula.length > 100000 || !Number.isInteger(rowDelta) || !Number.isInteger(columnDelta)) throw new Error('수식 참조 이동 요청이 유효하지 않습니다.');
  let output = '', index = 0;
  const isName = character => Boolean(character && /[\p{L}\p{N}_.]/u.test(character));
  while (index < formula.length) {
    const quote = formula[index];
    if (quote === '"' || quote === "'") {
      const start = index++, escapeQuote = quote === '"' ? '""' : "''";
      while (index < formula.length) {
        if (formula.startsWith(escapeQuote, index)) { index += 2; continue; }
        if (formula[index++] === quote) break;
      }
      output += formula.slice(start, index);
      continue;
    }
    if (quote === '[') {
      const start = index++;
      let depth = 1;
      while (index < formula.length && depth) {
        if (formula[index] === "'" && /['\[\]#@]/.test(formula[index+1]||'')) { index+=2; continue; }
        if (formula[index] === '[') depth++;
        else if (formula[index] === ']') depth--;
        index++;
      }
      output += formula.slice(start, index);
      continue;
    }
    const previous = formula[index - 1] || '';
    const qualifier = /^([\p{L}_][\p{L}\p{N}_.]*(?::[\p{L}_][\p{L}\p{N}_.]*)?)!/u.exec(formula.slice(index));
    if (qualifier && !isName(previous)) { output += qualifier[0]; index += qualifier[0].length; continue; }
    const match = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)/.exec(formula.slice(index));
    if (!match || isName(previous)) { output += formula[index++]; continue; }
    const end = index + match[0].length, next = formula[end] || '';
    if (isName(next) || next === '(') { output += formula[index++]; continue; }
    const sourceColumn = columnNumber(match[2].toUpperCase() + '1'), sourceRow = Number(match[4]);
    const targetColumn = match[1] ? sourceColumn : sourceColumn + columnDelta;
    const targetRow = match[3] ? sourceRow : sourceRow + rowDelta;
    if (targetColumn < 1 || targetColumn > 16384 || targetRow < 1 || targetRow > 1048576) output += '#REF!';
    else {
      const letters = columnLetters(targetColumn), normalizedLetters = match[2] === match[2].toLowerCase() ? letters.toLowerCase() : letters;
      output += match[1] + normalizedLetters + match[3] + targetRow;
    }
    index = end;
  }
  return output;
}

module.exports={translateFormulaReferences};
