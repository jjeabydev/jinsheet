'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
const { MAX_PDF_BYTES, MAX_PDF_CANVAS_PIXELS, MAX_PDF_CANVAS_DIMENSION, pdfCanvasOutputScale, validatePdfInput } = require('../src/pdf');


test('PDF magic header and size limits are validated', () => {
  const sample = fs.readFileSync(path.join(__dirname, 'fixtures/JinSheet-Sample.pdf'));
  assert.deepEqual(validatePdfInput(sample), { format: 'pdf', title: 'PDF 문서' });
  assert.throws(() => validatePdfInput(Buffer.from('not a PDF')), /PDF 파일 헤더/);
  assert.throws(() => validatePdfInput(Buffer.alloc(MAX_PDF_BYTES + 1, 0x25)), /25 MiB/);
});

test('bundled PDF.js parses and extracts text from a local fixture without Office software', async () => {
  const standardFontDataUrl = pathToFileURL(path.resolve(__dirname, '../media/pdf-standard-fonts') + path.sep).href;
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
  const sample = fs.readFileSync(path.join(__dirname, 'fixtures/JinSheet-Sample.pdf'));
  assert.deepEqual(validatePdfInput(sample), { format: 'pdf', title: 'PDF 문서' });
  const task = pdfjs.getDocument({ data: new Uint8Array(sample), isEvalSupported: false, disableFontFace: true, standardFontDataUrl, useSystemFonts: false });
  try {
    const document = await task.promise;
    assert.equal(document.numPages, 1);
    const page = await document.getPage(1);
    const content = await page.getTextContent();
    assert.ok(content.items.some(item => item.str === 'JinSheet PDF sample'));
    page.cleanup();
    await document.destroy();
  } finally { await task.destroy(); }
});

test('PDF canvas renders at device pixel resolution within safe canvas limits', () => {
  const width = 612 * 3, height = 792 * 3;
  const scale = pdfCanvasOutputScale(width, height, 2);
  assert.equal(scale, 2);
  assert.ok(width * height * scale * scale <= MAX_PDF_CANVAS_PIXELS);
  assert.ok(Math.max(width, height) * scale <= MAX_PDF_CANVAS_DIMENSION);
  assert.equal(pdfCanvasOutputScale(800, 1000, 1), 2);
  assert.equal(pdfCanvasOutputScale(800, 1000, 2), 2);
  assert.equal(pdfCanvasOutputScale(800, 1000, 3), 3);
  assert.throws(() => pdfCanvasOutputScale(0, 1000, 2), /PDF 페이지 크기/);
  const veryWideScale = pdfCanvasOutputScale(2_000_000, 100, 2);
  assert.ok(2_000_000 * veryWideScale <= MAX_PDF_CANVAS_DIMENSION);
});
