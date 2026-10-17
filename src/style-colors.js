'use strict';

// SpreadsheetML indexed colors 0–63; 64/65 are system foreground/background.
const INDEXED = ('000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF '+
  '000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF '+
  '800000 008000 000080 808000 800080 008080 C0C0C0 808080 '+
  '9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF '+
  '000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF '+
  '00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 '+
  '3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 '+
  '003366 339966 003300 333300 993300 993366 333399 333333').split(' ');
const THEME_ORDER = ['lt1','dk1','lt2','dk2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink'];
const local = node => node?.name?.split(':').at(-1);
const one = (node, name) => node?.children?.find(item => local(item) === name);
const rgb = value => typeof value === 'string' && /^(?:[\da-f]{2})?[\da-f]{6}$/i.test(value) ? value.slice(-6).toUpperCase() : null;

function tintColor(hex, tint) {
  if (!tint) return '#' + hex;
  const channels = [0,2,4].map(at => parseInt(hex.slice(at,at+2),16)/255);
  const max=Math.max(...channels),min=Math.min(...channels),delta=max-min;
  let light=(max+min)/2, saturation=delta ? delta/(1-Math.abs(2*light-1)) : 0, hue=0;
  if(delta) hue=(max===channels[0] ? (channels[1]-channels[2])/delta%6 : max===channels[1] ? (channels[2]-channels[0])/delta+2 : (channels[0]-channels[1])/delta+4)/6;
  if(hue<0)hue+=1;
  light=tint<0 ? light*(1+tint) : light*(1-tint)+tint;
  const chroma=(1-Math.abs(2*light-1))*saturation, h=hue*6, x=chroma*(1-Math.abs(h%2-1)), m=light-chroma/2;
  const values=h<1?[chroma,x,0]:h<2?[x,chroma,0]:h<3?[0,chroma,x]:h<4?[0,x,chroma]:h<5?[x,0,chroma]:[chroma,0,x];
  return '#'+values.map(value=>Math.round((value+m)*255).toString(16).padStart(2,'0')).join('').toUpperCase();
}

function createColorContext(themeRoot, stylesRoot) {
  const scheme=one(one(themeRoot,'themeElements'),'clrScheme');
  const theme=THEME_ORDER.map(name=>{
    const slot=one(scheme,name),node=slot?.children?.[0];
    if(slot?.children.length!==1 || node?.children?.length)return null;
    return local(node)==='srgbClr' ? rgb(node.attributes.val) : local(node)==='sysClr' ? rgb(node.attributes.lastClr) : null;
  });
  const palette=one(one(stylesRoot,'colors'),'indexedColors');
  const indexed=palette ? palette.children.map(node=>local(node)==='rgbColor'?rgb(node.attributes.rgb):null) : INDEXED;
  return {theme,indexed};
}

function resolveStyleColor(node, context) {
  if(!node || node.children?.length)return null;
  const attrs=node.attributes || {}, sources=['rgb','theme','indexed','auto'].filter(key=>attrs[key]!==undefined);
  if(sources.length!==1 || Object.keys(attrs).some(key=>!['rgb','theme','indexed','auto','tint'].includes(key)))return null;
  let color;
  if(sources[0]==='rgb')color=rgb(attrs.rgb);
  else if(sources[0]==='auto')return null;
  else {
    const key=sources[0],value=attrs[key];
    if(!/^(?:0|[1-9]\d*)$/.test(value) || Number(value)>=(key==='theme'?12:64))return null;
    color=(key==='theme'?context?.theme:context?.indexed || INDEXED)?.[Number(value)];
  }
  if(!color)return null;
  const raw=attrs.tint, tint=raw===undefined ? 0 : typeof raw==='string' && /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(raw) ? Number(raw) : NaN;
  return Number.isFinite(tint)&&tint>=-1&&tint<=1 ? tintColor(color,tint) : null;
}

module.exports={createColorContext,resolveStyleColor};
