'use strict';
const fs = require('node:fs');
const path = require('node:path');
const esbuild = require('esbuild');

const root = path.resolve(__dirname, '..');
const media = path.join(root, 'media');
fs.mkdirSync(media, { recursive: true });

async function main() {
  await esbuild.build({
    entryPoints: [path.join(root, 'src/webview/renderers.js')],
    bundle: true, minify: true, legalComments: 'eof', format: 'iife',
    platform: 'browser', target: ['chrome100'], outfile: path.join(media, 'renderers.js')
  });
  fs.cpSync(path.join(root, 'node_modules/pdfjs-dist/standard_fonts'), path.join(media, 'pdf-standard-fonts'), { recursive: true });
  const licenses = [
    ['node_modules/pdfjs-dist/LICENSE', 'LICENSE-pdfjs-dist.txt'],
    ['node_modules/docx-preview/LICENSE', 'LICENSE-docx-preview.txt'],
    ['node_modules/@file-viewer/pptx/LICENSE', 'LICENSE-file-viewer-pptx.txt'],
    ['node_modules/jszip/LICENSE.markdown', 'LICENSE-jszip.txt'],
    ['node_modules/dompurify/LICENSE', 'LICENSE-dompurify.txt'],
    ['node_modules/tinycolor2/LICENSE', 'LICENSE-tinycolor2.txt'],
    ['node_modules/utif/LICENSE', 'LICENSE-utif.txt'],
    ['node_modules/dingbat-to-unicode/LICENSE', 'LICENSE-dingbat-to-unicode.txt']
  ];
  for (const [source, target] of licenses) fs.copyFileSync(path.join(root, source), path.join(media, target));
  await esbuild.build({
    entryPoints: [path.join(root, 'node_modules/pdfjs-dist/build/pdf.mjs')],
    bundle: true, minify: true, legalComments: 'eof', format: 'iife', globalName: 'pdfjsLib',
    platform: 'browser', target: ['chrome100'], outfile: path.join(media, 'pdf.min.js')
  });
  const assets = [
    ['node_modules/pdfjs-dist/build/pdf.worker.min.mjs', 'pdf.worker.min.mjs'],
    ['node_modules/@file-viewer/pptx/dist/worker/pptx.worker.js', 'pptx.worker.js'],
    ['node_modules/@file-viewer/pptx/dist/styles/pptxjs.css', 'pptx.css']
  ];
  for (const [source, target] of assets) fs.copyFileSync(path.join(root, source), path.join(media, target));
}

main().catch(error => { console.error(error); process.exitCode = 1; });
