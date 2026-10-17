'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { validateCustomFilter, validateIconFilter, compileAutoFilterColumn } = require('../src/auto-filter');
const criterion = (operator, value) => ({ column: 0, custom: { and: false, conditions: [{ operator, value }] } });

test('조건 필터는 숫자 비교와 두 조건의 AND·OR를 평가한다', () => {
  const column = { column: 0, custom: { and: true, conditions: [{ operator: 'greaterThanOrEqual', value: '10' }, { operator: 'lessThan', value: '20' }] } };
  const matches = compileAutoFilterColumn(column);
  assert.equal(matches({ kind: 'number', value: '10' }), true); assert.equal(matches({ value: 15 }), true); assert.equal(matches({ value: 20 }), false); assert.equal(matches({ value: '15', kind: 'string' }), false); assert.equal(matches(null), false);
  column.custom.and = false; const either = compileAutoFilterColumn(column); assert.equal(either({ value: 5 }), true); assert.equal(either({ value: 30 }), true);
  for (const [operator, expected] of [['equal', true], ['notEqual', false], ['lessThan', false], ['lessThanOrEqual', true], ['greaterThan', false], ['greaterThanOrEqual', true]]) assert.equal(compileAutoFilterColumn(criterion(operator, '1e1'))({ kind: 'number', value: '10' }), expected);
});

test('텍스트 조건은 대소문자·와일드카드·이스케이프·빈 셀을 처리한다', () => {
  const contains = compileAutoFilterColumn(criterion('equal', '*한?*')); assert.equal(contains({ value: '앞한글뒤' }), true); assert.equal(contains({ value: '한' }), false);
  const literal = compileAutoFilterColumn(criterion('equal', 'a~*~?~~b')); assert.equal(literal({ value: 'A*?~B' }), true); assert.equal(literal({ value: 'AlphaB' }), false);
  assert.equal(compileAutoFilterColumn(criterion('equal', 'a?c'))({ value: 'A😀C' }), true);
  assert.equal(compileAutoFilterColumn(criterion('equal', 'abc'))({ value: 'ABC' }), true);
  assert.equal(compileAutoFilterColumn(criterion('equal', '*'))(null), false);
  assert.equal(compileAutoFilterColumn(criterion('equal', ''))(null), true);
  assert.equal(compileAutoFilterColumn(criterion('notEqual', ''))(null), false);
  assert.equal(compileAutoFilterColumn(criterion('notEqual', '*abc*'))(null), true);
  assert.equal(compileAutoFilterColumn(criterion('equal', '*10*'))({ kind: 'number', value: '10' }), false);
  assert.equal(compileAutoFilterColumn(criterion('equal', '*a*a*b'))({ value: 'a'.repeat(2000) + 'b' }), true);
});

test('조건 필터는 잘못된 연산자·조합·출력 한도를 거부한다', () => {
  for (const custom of [{ and: true, conditions: [] }, { and: 1, conditions: [{ operator: 'equal', value: 'x' }] }, { and: false, conditions: Array(3).fill({ operator: 'equal', value: 'x' }) }, { and: false, conditions: [{ operator: 'unknown', value: 'x' }] }, { and: false, conditions: [{ operator: 'equal', value: 10 }] }]) assert.throws(() => validateCustomFilter(custom));
  assert.throws(() => compileAutoFilterColumn(criterion('equal', '*a'.repeat(513))), /1,024/);
  assert.throws(() => compileAutoFilterColumn(criterion('equal', 'x'.repeat(32768))), /유효하지/);
  assert.equal(compileAutoFilterColumn(criterion('equal', 'x'.repeat(2000)))({ value: 'x'.repeat(2000) }), true);
});

test('와일드카드 자동자는 작은 패턴 전수와 긴 반복 텍스트에서 참조 판정과 일치한다', () => {
  const combinations = (alphabet, length) => length ? combinations(alphabet, length - 1).flatMap(prefix => alphabet.map(character => prefix + character)) : [''];
  const values = [1, 2, 3, 4].flatMap(length => combinations(['a', 'b'], length));
  const reference = (pattern, value) => {
    let previous = Array(value.length + 1).fill(false); previous[0] = true;
    for (const character of pattern) {
      const next = Array(value.length + 1).fill(false); next[0] = character === '*' && previous[0];
      for (let index = 1; index <= value.length; index++) next[index] = character === '*' ? previous[index] || next[index - 1] : (character === '?' || character === value[index - 1]) && previous[index - 1];
      previous = next;
    }
    return previous[value.length];
  };
  for (const length of [0, 1, 2, 3, 4, 5]) for (const pattern of combinations(['a', 'b', '*', '?'], length)) {
    const matches = compileAutoFilterColumn(criterion('equal', pattern));
    for (const value of values) assert.equal(matches({ value }), reference(pattern, value), pattern + ':' + value);
  }
  assert.equal(compileAutoFilterColumn(criterion('equal', '*' + 'a'.repeat(900) + '?b'))({ value: 'a'.repeat(20000) }), false);
});

test('상위·하위 개수와 백분율 필터는 숫자만 사용하고 동점과 올림을 처리한다', () => {
  const population=[10,20,20,30,40].map(value=>({kind:'number',value:String(value)})).concat([{value:'999',kind:'string'},{value:'TRUE',kind:'boolean'},{value:'',kind:'blank'},{value:'#DIV/0!',kind:'error'},null]);
  const rank=(top,percent,value)=>compileAutoFilterColumn({top10:{top,percent,value}},population);
  assert.deepEqual(population.filter(rank(true,false,2)).map(cell=>cell.value),['30','40']);assert.deepEqual(population.filter(rank(false,false,2)).map(cell=>cell.value),['10','20','20']);assert.deepEqual(population.filter(rank(true,true,21)).map(cell=>cell.value),['30','40']);assert.equal(population.filter(rank(false,true,100)).length,5);assert.equal(population.filter(rank(true,false,500)).length,5);assert.equal(rank(true,true,.01)({value:40}),true);
  assert.equal(compileAutoFilterColumn({top10:{top:true,percent:false,value:1}},[{value:'10',kind:'string'}])({value:10}),false);
});

test('순위 필터는 음수·수식 결과와 잘못된 값·과도한 모집단을 구분한다', () => {
  const {validateTopFilter,topFilterThreshold}=require('../src/auto-filter');const top10={top:true,percent:false,value:2};assert.equal(topFilterThreshold(top10,[{value:-5},{value:-1},{value:0},{value:Infinity},{value:false}]),-1);assert.equal(compileAutoFilterColumn({top10},[{value:-5},{value:-1},{value:0}])({value:-1}),true);
  for(const value of [0,-1,501,1.5,NaN,Infinity,'2'])assert.throws(()=>validateTopFilter({...top10,value}));for(const value of [0,-1,101,Infinity])assert.throws(()=>validateTopFilter({...top10,percent:true,value}));assert.throws(()=>validateTopFilter({...top10,top:1}));assert.throws(()=>validateTopFilter({...top10,unknown:true}));assert.throws(()=>topFilterThreshold(top10,(function*(){for(let i=0;i<=100000;i++)yield {value:i}})()),/100,000/);
});

test('날짜 그룹은 연·월·일·시간의 OR와 혼합 값·빈 셀을 숫자 날짜에 적용한다', () => {
  const {parseDateGroups}=require('../src/auto-filter');const dateGroups=parseDateGroups('2024-02\n2025\n2026-10-16T12:30\n2024-02');assert.equal(dateGroups.length,3);const matches=compileAutoFilterColumn({values:['literal'],blanks:true,dateGroups});
  assert.equal(matches({value:45351,kind:'number'}),true);assert.equal(matches({value:45352}),false);assert.equal(matches({value:45658}),true);assert.equal(matches({value:'45351',kind:'string'}),false);assert.equal(matches({value:'2024-02-29',kind:'string'}),false);assert.equal(matches({value:'literal',kind:'string'}),true);assert.equal(matches(null),true);assert.equal(matches({value:true}),false);
  const precise=compileAutoFilterColumn({values:[],blanks:false,dateGroups:parseDateGroups('1900-02-29T12:30:15')});assert.equal(precise({value:60+(12*3600+30*60+15)/86400}),true);assert.equal(precise({value:60+(12*3600+30*60+16)/86400}),false);
});

test('날짜 그룹은 윤년·1900 예외·1904 체계와 경계값을 검증한다', () => {
  const {parseDateGroups,validateDateGroups,serialDateParts}=require('../src/auto-filter');assert.equal(parseDateGroups('1900-02-29')[0].day,29);for(const text of ['2023-02-29','2024-04-31','2024-00','0000','2024-01-01T24','2024-01-01T12:60','2024-01-01T12:00:60','2024-1-1','2024/01/01'])assert.throws(()=>parseDateGroups(text),text);assert.throws(()=>validateDateGroups([{dateTimeGrouping:'month',year:2024,month:2,day:1}]));assert.throws(()=>validateDateGroups(Array(10001).fill({dateTimeGrouping:'year',year:2024})),/10,000/);
  assert.deepEqual(serialDateParts(60.5),{year:1900,month:2,day:29,weekday:3,hour:12,minute:0,second:0});assert.equal(serialDateParts(61).day,1);assert.equal(serialDateParts(0,true).year,1904);assert.equal(serialDateParts(0,true).day,1);assert.equal(serialDateParts(-1),null);assert.equal(serialDateParts(Infinity),null);assert.equal(serialDateParts(2958466),null);assert.equal(serialDateParts(2958465.5).hour,12);
  const matches=compileAutoFilterColumn({values:[],blanks:false,dateGroups:parseDateGroups('1904-01-01')},[],{date1904:true});assert.equal(matches({value:0}),true);assert.equal(matches({value:1}),false);
});

test('동적 날짜는 윤일·주·월·분기·연도 범위를 현재 날짜에서 다시 계산한다', () => {
  const {dynamicFilterCriteria}=require('../src/auto-filter'),{evaluateFormula}=require('../src/formula'),now=new Date(2024,1,29,18,15),serial=(y,m,d)=>evaluateFormula('DATE('+[y,m,d].join(',')+')');
  for(const [type,start,end] of [['today',[2024,2,29],[2024,3,1]],['yesterday',[2024,2,28],[2024,2,29]],['tomorrow',[2024,3,1],[2024,3,2]],['thisWeek',[2024,2,25],[2024,3,3]],['lastWeek',[2024,2,18],[2024,2,25]],['nextWeek',[2024,3,3],[2024,3,10]],['thisMonth',[2024,2,1],[2024,3,1]],['lastMonth',[2024,1,1],[2024,2,1]],['nextMonth',[2024,3,1],[2024,4,1]],['thisQuarter',[2024,1,1],[2024,4,1]],['lastQuarter',[2023,10,1],[2024,1,1]],['nextQuarter',[2024,4,1],[2024,7,1]],['thisYear',[2024,1,1],[2025,1,1]],['lastYear',[2023,1,1],[2024,1,1]],['nextYear',[2025,1,1],[2026,1,1]],['yearToDate',[2024,1,1],[2024,3,1]]]){
    const criteria=dynamicFilterCriteria({type,value:1,maxValue:2},[],{now});assert.equal(criteria.value,serial(...start),type);assert.equal(criteria.maxValue,serial(...end),type);const matches=compileAutoFilterColumn({dynamic:{type}},[],{now});assert.equal(matches({value:criteria.value}),true,type);assert.equal(matches({value:criteria.maxValue-.25}),true,type);assert.equal(matches({value:criteria.maxValue}),false,type);assert.equal(matches({value:String(criteria.value),kind:'string'}),false,type);
  }
  const criteria=dynamicFilterCriteria({type:'today'},[],{now,date1904:true});assert.equal(criteria.value,43889);assert.equal(criteria.maxValue,43890);
});

test('매년 월·분기는 모든 연도에 적용하고 동적 날짜 판정은 한 번의 시각을 유지한다', () => {
  const {evaluateFormula}=require('../src/formula'),now=new Date(2024,11,31,23,59),serial=(y,m,d)=>evaluateFormula('DATE('+[y,m,d].join(',')+')');const feb=compileAutoFilterColumn({dynamic:{type:'M2'}});assert.equal(feb({value:serial(2020,2,29)}),true);assert.equal(feb({value:serial(2024,2,29)}),true);assert.equal(feb({value:serial(2024,3,1)}),false);const first=compileAutoFilterColumn({dynamic:{type:'Q1'}});assert.equal(first({value:serial(2025,3,31)}),true);assert.equal(first({value:serial(2025,4,1)}),false);
  const current=compileAutoFilterColumn({dynamic:{type:'today'}},[],{now});now.setDate(now.getDate()+1);assert.equal(current({value:serial(2024,12,31)+.5}),true);assert.equal(current({value:serial(2025,1,1)}),false);assert.equal(compileAutoFilterColumn({dynamic:{type:'today'}},[],{now})({value:serial(2025,1,1)+.5}),true);assert.equal(compileAutoFilterColumn({dynamic:{type:'null'}})(null),true);
});

test('평균 초과·미만은 숫자 모집단과 동점을 구분하고 큰 수의 평균을 안정적으로 계산한다', () => {
  const {dynamicFilterCriteria}=require('../src/auto-filter'),population=[{value:10},{value:20},{value:30},{value:'999',kind:'string'},{value:true},null];const above=compileAutoFilterColumn({dynamic:{type:'aboveAverage'}},population),below=compileAutoFilterColumn({dynamic:{type:'belowAverage'}},population);assert.equal(above({value:30}),true);assert.equal(above({value:20}),false);assert.equal(below({value:10}),true);assert.equal(below({value:20}),false);assert.equal(above({value:'30',kind:'string'}),false);assert.equal(dynamicFilterCriteria({type:'aboveAverage'},[{value:-1e308},{value:1e308}]).value,0);assert.equal(compileAutoFilterColumn({dynamic:{type:'aboveAverage'}},[])({value:1}),false);
  const {evaluateFormula}=require('../src/formula');assert.equal(dynamicFilterCriteria({type:'aboveAverage'},[.1,.2,.3].map(value=>({value}))).value,evaluateFormula('AVERAGE(HSTACK(0.1,0.2,0.3))'));assert.equal(dynamicFilterCriteria({type:'aboveAverage'},[{value:1e308},{value:1e308}]).value,1e308);
});

test('동적 조건의 잘못된 종류·기준값·시각과 과도한 모집단을 거부한다', () => {
  const {validateDynamicFilter,dynamicFilterCriteria}=require('../src/auto-filter');for(const dynamic of [null,{type:'thisDecade'},{type:'today',value:'1'},{type:'today',maxValue:Infinity},{type:'today',unknown:true}])assert.throws(()=>validateDynamicFilter(dynamic));assert.throws(()=>dynamicFilterCriteria({type:'today'},[],{now:new Date(NaN)}),/기준 시각/);assert.throws(()=>dynamicFilterCriteria({type:'aboveAverage'},(function*(){for(let i=0;i<=100000;i++)yield {value:i}})()),/100,000/);
});

test('색상 필터는 채우기·글자색·자동 색을 값과 서식 변경 없이 판정한다', () => {
  const {validateColorFilter}=require('../src/auto-filter');const fill=compileAutoFilterColumn({colorFilter:{cellColor:true,color:'#ff0000'}}),font=compileAutoFilterColumn({colorFilter:{cellColor:false,color:'#00ff00'}}),none=compileAutoFilterColumn({colorFilter:{cellColor:true,color:null}});const cell={value:'unchanged',styleProperties:{fill:'#FF0000',color:'#00FF00',bold:true}};assert.equal(fill(cell),true);assert.equal(font(cell),true);assert.equal(none(cell),false);assert.equal(none(null),true);assert.equal(fill({value:10,styleProperties:{fill:'#0000FF'}}),false);assert.equal(cell.value,'unchanged');assert.equal(cell.styleProperties.bold,true);assert.equal(compileAutoFilterColumn({colorFilter:{cellColor:true,color:'#FF0000'}},[],{getStyle:()=>({fill:'#FF0000'})})({value:10}),true);
  for(const colorFilter of [{cellColor:1,color:null},{cellColor:true,color:'red'},{cellColor:true,color:'#12345'},{cellColor:true,color:null,dxfId:-1},{cellColor:true,color:null,unknown:true}])assert.throws(()=>validateColorFilter(colorFilter));
});

test('아이콘 필터는 표준 집합과 아이콘 번호를 검증하고 조건부 표시 아이콘만 비교한다', () => {
  const iconFilter = { iconSet: '3TrafficLights1', iconId: 1 };
  assert.equal(validateIconFilter(iconFilter), iconFilter);
  const matches = compileAutoFilterColumn({ column: 0, iconFilter }, [], { getStyle: cell => cell.style });
  assert.equal(matches({ value: 'unchanged', style: { icon: { set: '3TrafficLights1', id: 1 } } }), true);
  assert.equal(matches({ value: 'unchanged', style: { icon: { set: '3TrafficLights1', id: 0 } } }), false);
  assert.equal(matches({ value: 'unchanged', style: { icon: { set: '3Arrows', id: 1 } } }), false);
  for (const invalid of [null, { iconSet: 'unsupported', iconId: 0 }, { iconSet: '3Arrows', iconId: 3 }, { iconSet: '3Arrows', iconId: 0, extra: true }]) assert.throws(() => validateIconFilter(invalid));
});
