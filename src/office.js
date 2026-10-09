'use strict';

const path = require('node:path').posix;
const { safeUnzip, parseXml, children } = require('./xlsx');

const MAX_PARAGRAPHS = 20000;
const MAX_TEXT_CHARS = 2_000_000;

function fail(message) { throw new Error(message); }
function child(node, name) { return node.children.find(item => item.name === name); }

function relationships(root, basePath) {
  const found = new Map();
  for (const item of children(root, 'Relationship')) {
    const { Id, Target, Type, TargetMode } = item.attributes;
    if (!Id || !Target || !Type || found.has(Id)) fail('Office 문서 관계가 유효하지 않습니다.');
    if (TargetMode === 'External') { found.set(Id, { external: true, Type }); continue; }
    const resolved = path.normalize(Target.startsWith('/') ? Target.slice(1) : path.join(basePath, Target));
    if (resolved === '..' || resolved.startsWith('../') || resolved.startsWith('/')) fail('Office 문서 관계가 패키지 외부 경로를 가리킵니다.');
    found.set(Id, { path: resolved, Type });
  }
  return found;
}

function orderedParagraphs(root) {
  const paragraphs = [];
  const paragraphText = node => {
    if (node.name.endsWith(':t')) return node.text;
    if (node.name.endsWith(':tab')) return '\t';
    if (node.name.endsWith(':br')) return '\n';
    return node.children.map(paragraphText).join('');
  };
  const visit = node => {
    if (node.name === 'p' || node.name.endsWith(':p')) {
      const value = paragraphText(node).replaceAll('\u00a0', ' ');
      if (value.trim()) paragraphs.push(value);
      if (paragraphs.length > MAX_PARAGRAPHS) fail('Office 문단 개수 제한을 초과했습니다.');
      return;
    }
    for (const item of node.children) visit(item);
  };
  visit(root);
  return paragraphs;
}

function checkedText(paragraphs) {
  const total = paragraphs.reduce((sum, value) => sum + value.length, 0);
  if (total > MAX_TEXT_CHARS) fail('Office 추출 텍스트 크기 제한을 초과했습니다.');
  return paragraphs;
}

function readDocx(files) {
  const documentPath = 'word/document.xml';
  const parsed = parseXml(files[documentPath], documentPath);
  if (parsed.root.name !== 'w:document' && !parsed.root.name.endsWith(':document')) fail('DOCX 본문 XML 형식이 아닙니다.');
  const body = child(parsed.root, 'w:body') || parsed.root.children.find(item => item.name.endsWith(':body'));
  if (!body) fail('DOCX 본문을 찾을 수 없습니다.');
  const paragraphs = orderedParagraphs(body);
  return { format: 'docx', title: 'Word 문서', sections: [{ title: '본문', paragraphs: checkedText(paragraphs) }] };
}

function readPptx(files) {
  const presentationPath = 'ppt/presentation.xml';
  const presentation = parseXml(files[presentationPath], presentationPath).root;
  const relsPath = 'ppt/_rels/presentation.xml.rels';
  const relsRoot = parseXml(files[relsPath], relsPath).root;
  const rels = relationships(relsRoot, 'ppt');
  const list = presentation.children.find(item => item.name.endsWith(':sldIdLst'));
  if (!list) fail('PPTX 슬라이드 목록을 찾을 수 없습니다.');
  const slideIds = list.children.filter(item => item.name.endsWith(':sldId'));
  if (!slideIds.length || slideIds.length > 2000) fail('PPTX 슬라이드 개수가 없거나 제한을 초과했습니다.');
  const sections = [];
  for (let i = 0; i < slideIds.length; i++) {
    const relId = slideIds[i].attributes['r:id'] || slideIds[i].attributes['relationships:id'];
    const rel = rels.get(relId);
    if (!rel?.path || rel.external || !rel.Type.endsWith('/slide') || !files[rel.path]) fail(`PPTX 슬라이드 관계가 유효하지 않습니다: ${i + 1}`);
    const slide = parseXml(files[rel.path], rel.path).root;
    const paragraphs = checkedText(orderedParagraphs(slide));
    sections.push({ title: `슬라이드 ${i + 1}`, paragraphs });
    if (sections.reduce((sum, section) => sum + section.paragraphs.length, 0) > MAX_PARAGRAPHS) fail('PPTX 문단 개수 제한을 초과했습니다.');
  }
  checkedText(sections.flatMap(section => section.paragraphs));
  return { format: 'pptx', title: 'PowerPoint 프레젠테이션', sections };
}

function analyzeOffice(input, extension) {
  const ext = String(extension).toLowerCase();
  if (!['.docx', '.docm', '.pptx', '.pptm'].includes(ext)) fail('구형 .doc/.ppt와 이 형식은 미리보기를 지원하지 않습니다. .docx 또는 .pptx로 저장해 주세요.');
  const { files } = safeUnzip(input);
  const types = parseXml(files['[Content_Types].xml'], '[Content_Types].xml').root;
  if (types.name !== 'Types') fail('Office 콘텐츠 유형 XML 형식이 아닙니다.');
  if (ext.startsWith('.doc')) {
    if (!files['word/document.xml']) fail('Word 본문을 찾을 수 없습니다.');
    return readDocx(files);
  }
  if (!files['ppt/presentation.xml']) fail('PowerPoint 본문을 찾을 수 없습니다.');
  return readPptx(files);
}

module.exports = { analyzeOffice, orderedParagraphs, relationships };
