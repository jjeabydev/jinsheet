'use strict';

const MAX_PDF_BYTES = 25 * 1024 * 1024;
const MAX_PDF_CANVAS_PIXELS = 32 * 1024 * 1024;
const MAX_PDF_CANVAS_DIMENSION = 16384;

function pdfCanvasOutputScale(width, height, devicePixelRatio = 1) {
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) throw new RangeError('PDF 페이지 크기가 유효하지 않습니다.');
  const requestedScale = Number.isFinite(devicePixelRatio) && devicePixelRatio > 0 ? Math.min(4, Math.max(2, devicePixelRatio)) : 2;
  const areaScale = Math.sqrt(MAX_PDF_CANVAS_PIXELS / (width * height));
  const dimensionScale = MAX_PDF_CANVAS_DIMENSION / Math.max(width, height);
  return Math.min(requestedScale, areaScale, dimensionScale);
}

function validatePdfInput(input) {
  const bytes = Buffer.isBuffer(input) ? input : Buffer.from(input || []);
  if (bytes.length < 8) throw new Error('PDF 파일이 비어 있거나 잘렸습니다.');
  if (bytes.length > MAX_PDF_BYTES) throw new Error('PDF 크기 제한(25 MiB)을 초과했습니다.');
  const header = bytes.subarray(0, Math.min(bytes.length, 1024)).toString('latin1');
  if (!/%PDF-\d\.\d/.test(header)) throw new Error('유효한 PDF 파일 헤더를 찾을 수 없습니다.');
  return { format: 'pdf', title: 'PDF 문서' };
}

module.exports = { MAX_PDF_BYTES, MAX_PDF_CANVAS_PIXELS, MAX_PDF_CANVAS_DIMENSION, pdfCanvasOutputScale, validatePdfInput };
