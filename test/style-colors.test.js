'use strict';
const test=require('node:test');
const assert=require('node:assert/strict');
const {createColorContext,resolveStyleColor}=require('../src/style-colors');
const node=(name,attributes={},children=[])=>({name,attributes,children});
const color=attributes=>node('color',attributes);

test('색상은 RGB·ARGB와 기본 인덱스 팔레트를 해석한다',()=>{
  for(const value of ['123abc','FF123ABC','00123ABC'])assert.equal(resolveStyleColor(color({rgb:value})),'#123ABC');
  for(const [index,value] of [[0,'#000000'],[1,'#FFFFFF'],[10,'#FF0000'],[22,'#C0C0C0'],[40,'#00CCFF'],[63,'#333333']])assert.equal(resolveStyleColor(color({indexed:String(index)})),value);
  for(const indexed of ['64','65','66','-1','1.5'])assert.equal(resolveStyleColor(color({indexed})),null);
});

test('테마의 명명된 슬롯은 XML 순서와 무관하며 시스템 lastClr를 사용한다',()=>{
  const theme=node('a:theme',{},[node('a:themeElements',{},[node('a:clrScheme',{},[
    node('a:accent1',{},[node('a:srgbClr',{val:'4F81BD'})]),
    node('a:dk1',{},[node('a:sysClr',{val:'windowText',lastClr:'123456'})]),
    node('a:lt1',{},[node('a:sysClr',{val:'window',lastClr:'FFFFFF'})])
  ])])]);
  const context=createColorContext(theme,null);
  assert.equal(resolveStyleColor(color({theme:'0'}),context),'#FFFFFF');
  assert.equal(resolveStyleColor(color({theme:'1'}),context),'#123456');
  assert.equal(resolveStyleColor(color({theme:'4'}),context),'#4F81BD');
  assert.equal(resolveStyleColor(color({theme:'5'}),context),null);
  assert.equal(resolveStyleColor(color({theme:'4'})),null);
});

test('밝기 조정은 HSL 명도와 양끝을 사용한다',()=>{
  for(const [tint,value] of [['0','#4F81BD'],['0.4','#95B3D7'],['-0.25','#376092'],['1','#FFFFFF'],['-1','#000000']])assert.equal(resolveStyleColor(color({rgb:'4F81BD',tint})),value,tint);
  assert.equal(resolveStyleColor(color({rgb:'000000',tint:'.5'})),'#808080');
  assert.equal(resolveStyleColor(color({indexed:'10',tint:'-0.5'})),'#800000');
});

test('사용자 팔레트는 기본 인덱스를 덮어쓰고 누락한 색을 추측하지 않는다',()=>{
  const styles=node('styleSheet',{},[node('colors',{},[node('indexedColors',{},[node('rgbColor',{rgb:'FF112233'}),node('rgbColor',{rgb:'FFABCDEF'})])])]);
  const context=createColorContext(null,styles);
  assert.equal(resolveStyleColor(color({indexed:'0'}),context),'#112233');
  assert.equal(resolveStyleColor(color({indexed:'1'}),context),'#ABCDEF');
  assert.equal(resolveStyleColor(color({indexed:'2'}),context),null);
  assert.equal(resolveStyleColor(color({indexed:'0'})),'#000000');
  const extended={indexed:Array(66).fill('FFFFFF')};
  assert.equal(resolveStyleColor(color({indexed:'64'}),extended),null);
  assert.equal(resolveStyleColor(color({indexed:'65'}),extended),null);
});

test('잘못된 색상·중복 원천·밝기·미지원 변환은 거부한다',()=>{
  for(const attributes of [{rgb:'12345'},{rgb:'GGGGGG'},{rgb:'112233',theme:'0'},{indexed:'0x10'},{rgb:'112233',tint:'2'},{rgb:'112233',tint:'NaN'},{rgb:'112233',tint:''},{rgb:'112233',tint:'0x1'},{rgb:'112233',other:'1'}])assert.equal(resolveStyleColor(color(attributes)),null,JSON.stringify(attributes));
  assert.equal(resolveStyleColor(node('color',{rgb:'112233'},[node('transform')])),null);
  assert.equal(resolveStyleColor(color({auto:'1'})),null);
});
