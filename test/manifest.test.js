const fs = require("node:fs");
"use strict";

const assert = require("node:assert/strict");
const test = require("node:test");
const manifest = require("../package.json");

test("Office files are associated with JinSheet as the default editor", () => {
  const editor = manifest.contributes.customEditors.find(item => item.viewType === "jinsheet.editor");
  assert.ok(editor, "JinSheet custom editor is contributed");
  assert.equal(editor.priority, "default");
  const patterns = new Set(editor.selector.map(selector => selector.filenamePattern));
  for (const extension of ["*.docx", "*.docm", "*.pptx", "*.pptm", "*.pdf"]) assert.ok(patterns.has(extension), `${extension} is associated with JinSheet`);
  for (const extension of ["*.doc", "*.ppt"]) assert.ok(!patterns.has(extension), `${extension} remains unsupported`);
});

test("author metadata keeps jelly separate from the publisher identifier", () => {
  assert.equal(manifest.author.name, "jelly");
  assert.equal(manifest.author.url, "https://jellybeanz.medium.com");
  assert.equal(manifest.publisher, "jjeabydev");
});


test("Office, PDF and spreadsheet rendering does not launch an external program", () => {
  const source = fs.readFileSync(require.resolve("../src/extension"), "utf8");
  assert.doesNotMatch(source, /child_process|execFile|spawn\(/);
  assert.deepEqual(Object.keys(manifest.dependencies).sort(), ["fflate", "saxes"]);
});

test("the PDF webview bundle is generated from the locked PDF.js version", () => {
  const extension = fs.readFileSync(require.resolve("../src/extension"), "utf8");
  const build = fs.readFileSync(require.resolve("../scripts/build-renderers.js"), "utf8");
  assert.match(extension, /asset\('pdf\.min\.js'\)/);
  assert.match(extension, /globalThis\.pdfjsLib/);
  assert.match(build, /pdf\.mjs/);
  assert.match(build, /globalName: 'pdfjsLib'/);
  assert.match(build, /pdf\.min\.js/);
});
