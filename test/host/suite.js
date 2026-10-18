'use strict';
const assert = require('assert');
const vscode = require('vscode');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const { performance } = require('node:perf_hooks');
const { zipSync } = require('../../vendor/node_modules/fflate');
const { analyze, patchColumnWidths, patchOutlineGroups, patchAddChart, patchAddTable, patchAutoFilters, patchPrintArea, patchPrintSettings, patchPrintTitles, patchHeaderFooter } = require('../../src/xlsx');
const { onePagePdf } = require('../helpers/pdf-fixture');
const { DEFAULT_PRINT_SETTINGS, renderPrintDocument } = require('../../src/printing');
const { defaultHeaderFooter } = require('../../src/header-footer');
const { chartPreviewHtml, JinSheetProvider } = require('../../src/extension');
const { chartPrintGeometry } = require('../../src/chart-print');

async function measurePrintPages(evaluate, html, fitHeaders = false) {
  return evaluate(`(async () => {
    const parsed=new DOMParser().parseFromString(${JSON.stringify(html)},'text/html'),script=parsed.querySelector('script')?.textContent;
    for(const node of parsed.querySelectorAll('script'))node.remove();
    const host=document.createElement('div');host.style.cssText='position:fixed;left:-20000px;top:0;visibility:hidden;font:13px Arial,sans-serif';document.body.append(host);
    const shadow=host.attachShadow({mode:'open'}),style=document.createElement('style');style.textContent=parsed.querySelector('style').textContent;shadow.append(style);
    const content=document.createElement('div');content.innerHTML=parsed.body.innerHTML;shadow.append(content);let printed=0;
    try {
      if(${fitHeaders}){const scopedDocument={querySelectorAll:selector=>shadow.querySelectorAll(selector),getElementById:id=>shadow.getElementById(id),fonts:document.fonts};new Function('document','window',script)(scopedDocument,{addEventListener(){},print(){printed++}});await document.fonts.ready;await new Promise(resolve=>requestAnimationFrame(resolve));shadow.getElementById('print').click()}
      return [...shadow.querySelectorAll('.page')].map(page=>{const computed=getComputedStyle(page),table=page.querySelector('table').getBoundingClientRect(),header=page.querySelector('.print-header'),footer=page.querySelector('.print-footer'),cells=[...page.querySelectorAll('.hf-cell')];return{
        cellWidths:[...page.querySelectorAll('tbody tr:first-child td')].map(cell=>cell.getBoundingClientRect().width/Number(page.dataset.printScale||1)),cellHeights:[...page.querySelectorAll('tbody tr:first-child td')].map(cell=>cell.getBoundingClientRect().height/Number(page.dataset.printScale||1)),width:table.width,height:table.height,availableWidth:page.clientWidth-parseFloat(computed.paddingLeft)-parseFloat(computed.paddingRight),availableHeight:page.clientHeight-parseFloat(computed.paddingTop)-parseFloat(computed.paddingBottom),
        headerText:header?.textContent||'',footerText:footer?.textContent||'',
        hfFits:cells.every(cell=>{const value=cell.querySelector('.hf-content');return value.scrollHeight<=cell.clientHeight+1&&value.scrollWidth<=cell.clientWidth+1}),
        clearOfBody:(!header||[...header.querySelectorAll('.hf-content')].every(value=>value.getBoundingClientRect().bottom<=table.top+1))&&(!footer||[...footer.querySelectorAll('.hf-content')].every(value=>value.getBoundingClientRect().top>=table.bottom-1)),
        charts:[...page.querySelectorAll('.print-chart-clip')].map(clip=>{const box=clip.getBoundingClientRect(),svg=clip.querySelector('svg')?.getBoundingClientRect();return{width:box.width,height:box.height,svgWidth:svg?.width||0,svgHeight:svg?.height||0,pointCount:clip.querySelectorAll('[data-chart-point]').length,legendCount:clip.querySelectorAll('[data-print-legend]').length,legendFits:[...clip.querySelectorAll('[data-print-legend]')].every(node=>{const value=node.getBBox(),view=node.ownerSVGElement.viewBox.baseVal;return value.x>=view.x&&value.y>=view.y&&value.x+value.width<=view.x+view.width+.1&&value.y+value.height<=view.y+view.height+.1}),inside:box.left>=table.left-1&&box.top>=table.top-1&&box.right<=table.right+1&&box.bottom<=table.bottom+1}}),
        fontSizes:cells.map(cell=>Number.parseFloat(cell.style.fontSize)),printDisabled:shadow.getElementById('print').disabled,warning:shadow.getElementById('printWarning')?.textContent||'',printed
      }})
    } finally {host.remove()}
  })()`);
}

function normalizePath(value) {
  const resolved = path.resolve(value);
  return process.platform === 'win32' ? resolved.toLowerCase() : resolved;
}

async function connectDevTools(target) {
  const socket = new WebSocket(target.webSocketDebuggerUrl), pending = new Map();
  let nextId = 0;
  socket.addEventListener('message', event => {
    const message = JSON.parse(event.data), callback = pending.get(message.id);
    if (!callback) return;
    pending.delete(message.id);
    clearTimeout(callback.timer);
    if (message.error) callback.reject(new Error(message.error.message));
    else callback.resolve(message.result);
  });
  await new Promise((resolve, reject) => {
    socket.addEventListener('open', resolve, { once: true });
    socket.addEventListener('error', reject, { once: true });
  });
  return {
    send(method, params = {}) {
      const id = ++nextId;
      return new Promise((resolve, reject) => {
        const timer = setTimeout(() => { pending.delete(id); reject(new Error('DevTools request timed out: ' + method)); }, 5000);
        pending.set(id, { resolve, reject, timer });
        socket.send(JSON.stringify({ id, method, params }));
      });
    },
    close() { socket.close(); }
  };
}

async function waitForSpreadsheetWebview(timeoutMs = 15000, expectedCellValue = undefined) {
  const end = Date.now() + timeoutMs;
  let lastError, lastTargets = [], lastEvaluations = [], lastFrames = [];
  while (Date.now() < end) {
    try {
      const response = await fetch('http://127.0.0.1:9333/json/list');
      lastTargets = await response.json();
      for (const target of lastTargets) {
        if (!['page', 'iframe'].includes(target.type) || !target.webSocketDebuggerUrl) continue;
        let devtools;
        try {
          devtools = await connectDevTools(target);
          await devtools.send('Runtime.enable');
          await devtools.send('Page.enable');
          const frameTree = await devtools.send('Page.getFrameTree');
          const frames = [], collect = tree => { frames.push(tree.frame); for (const child of tree.childFrames || []) collect(child); };
          collect(frameTree.frameTree);
          lastFrames = frames.map(frame => frame.url);
          for (const frame of frames.filter(item => item.url.startsWith('vscode-webview://'))) {
            const world = await devtools.send('Page.createIsolatedWorld', { frameId: frame.id, worldName: 'jinsheet-ui-test', grantUniveralAccess: true });
            const result = await devtools.send('Runtime.evaluate', {
              expression: 'Boolean(document.getElementById("grid"))', contextId: world.executionContextId, returnByValue: true
            });
            lastEvaluations.push({ target: frame.url, hasGrid: result.result?.value });
            if (!result.result?.value) continue;
            if (expectedCellValue !== undefined) {
              const valueDeadline = Date.now() + 5000;
              let cellValue;
              do {
                const cell = await devtools.send('Runtime.evaluate', {
                  expression: 'Array.from(document.querySelectorAll(".cell")).some(cell => cell.textContent?.trim() === ' + JSON.stringify(expectedCellValue) + ')', contextId: world.executionContextId, returnByValue: true
                });
                cellValue = cell.result?.value;
                if (cellValue) break;
                await new Promise(resolve => setTimeout(resolve, 25));
              } while (Date.now() < valueDeadline);
              lastEvaluations.push({ target: frame.url, hasGrid: true, cellValue });
              if (!cellValue) continue;
            }
            await devtools.send('DOM.enable');
            const owner = await devtools.send('DOM.getFrameOwner', { frameId: frame.id });
            const box = await devtools.send('DOM.getBoxModel', { backendNodeId: owner.backendNodeId });
            const quad = box.model.content, xs = quad.filter((_, index) => index % 2 === 0), ys = quad.filter((_, index) => index % 2 === 1);
            devtools.contextId = world.executionContextId;
            devtools.frameOffset = { x: Math.min(...xs), y: Math.min(...ys) };
            const matched = devtools;
            devtools = null;
            return matched;
          }
        } catch (error) { lastError = error; }
        finally { if (devtools) devtools.close(); }
      }
    } catch (error) { lastError = error; }
    await new Promise(resolve => setTimeout(resolve, 100));
  }
  const targetSummary = lastTargets.map(target => `${target.type}:${target.url}`).join(' | ');
  throw new Error('JinSheet Webview was not exposed over DevTools' + (lastError ? ': ' + lastError.message : '') + (targetSummary ? '; targets=' + targetSummary : '; no targets were listed') + '; frames=' + JSON.stringify(lastFrames) + '; evaluations=' + JSON.stringify(lastEvaluations.slice(-4)));
}

suite('JinSheet Extension Host', () => {
  suiteSetup(async () => {
    const extension = vscode.extensions.getExtension('jjeabydev.jinsheet');
    assert.ok(extension, 'JinSheet extension is registered');
    // Provider registration must finish before testing default editor resolution.
    // Continue to use vscode.open, so these tests still exercise the association.
    await extension.activate();
  });

  test('opens PPTX with JinSheet from the default editor association', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-pptx-default-'));
    const file = path.join(directory, 'default-smoke.pptx');
    const bytes = zipSync({
      '[Content_Types].xml': new TextEncoder().encode('<Types/>'),
      'ppt/presentation.xml': new TextEncoder().encode('<p:presentation xmlns:p="urn:p" xmlns:r="urn:r"><p:sldIdLst><p:sldId r:id="r1"/></p:sldIdLst></p:presentation>'),
      'ppt/_rels/presentation.xml.rels': new TextEncoder().encode('<Relationships><Relationship Id="r1" Type="urn/slide" Target="slides/slide1.xml"/></Relationships>'),
      'ppt/slides/slide1.xml': new TextEncoder().encode('<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><a:p><a:r><a:t>Preview</a:t></a:r></a:p></p:sld>')
    });
    await fs.writeFile(file, bytes);
    try {
      await vscode.commands.executeCommand('vscode.open', vscode.Uri.file(file));
      const deadline = Date.now() + 10000;
      const findTab = () => vscode.window.tabGroups.all.flatMap(group => group.tabs).find(tab => tab.input.uri && normalizePath(tab.input.uri.fsPath) === normalizePath(file));
      let tab = findTab();
      while (Date.now() < deadline && tab?.input.viewType !== 'jinsheet.editor') {
        await new Promise(resolve => setTimeout(resolve, 100));
        tab = findTab();
      }
      assert.ok(tab, 'PPTX editor tab opened');
      assert.equal(tab.input.viewType, 'jinsheet.editor');
      assert.equal(normalizePath(tab.input.uri.fsPath), normalizePath(file));
    } finally {
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });


  test('opens PDF with JinSheet from the default editor association', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-pdf-default-'));
    const file = path.join(directory, 'default-smoke.pdf');
    await fs.writeFile(file, onePagePdf());
    try {
      await vscode.commands.executeCommand('vscode.open', vscode.Uri.file(file));
      await new Promise(resolve => setTimeout(resolve, 500));
      const tab = vscode.window.tabGroups.activeTabGroup.activeTab;
      assert.ok(tab, 'PDF editor tab opened');
      assert.equal(tab.input.viewType, 'jinsheet.editor');
      assert.equal(normalizePath(tab.input.uri.fsPath), normalizePath(file));
    } finally {
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('opens XLSX fixtures in JinSheet through the default editor association', async () => {
    const fixtures = ['JinSheet-Sample.xlsx', 'JinSheet-Styled-Sample.xlsx'];
    try {
      for (const name of fixtures) {
        const file = path.join(__dirname, '..', 'fixtures', name);
        await vscode.commands.executeCommand('vscode.open', vscode.Uri.file(file));
        await new Promise(resolve => setTimeout(resolve, 500));
        const tab = vscode.window.tabGroups.activeTabGroup.activeTab;
        assert.ok(tab, `${name} custom editor tab opened`);
        assert.equal(tab.input.viewType, 'jinsheet.editor');
        assert.equal(normalizePath(tab.input.uri.fsPath), normalizePath(file));
        await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      }
    } finally {
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
    }
  });

  test('Chromium Webview displays INDEX row and column arrays through edit, save and reopen', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-index-array-host-'));
    const file = path.join(directory, 'index-arrays.xlsx');
    const parts = {
      '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
      '_rels/.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
      'xl/workbook.xml': '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="s1"/></sheets></workbook>',
      'xl/_rels/workbook.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="s1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
      'xl/worksheets/sheet1.xml': '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
        + '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c><c r="D1"><f>INDEX(A1:C3,2,0)</f><v>4</v></c><c r="G1"><f>SUM(D1#)</f><v>15</v></c><c r="I1"><f>INDEX(A1:C3,0,2)</f><v>2</v></c><c r="K1"><f>SUM(I1#)</f><v>15</v></c><c r="L1"><f>TRIM(A1:A3)</f><v>1</v></c><c r="N1"><f>CHOOSE(IF(A1:A3&lt;5,1,2),"low","high")</f><v>low</v></c><c r="O1"><f>IFS(A1:A3&lt;2,"small",A1:A3&lt;5,"medium",TRUE,"large")</f><v>small</v></c><c r="P1"><f>SWITCH(A1:A3,1,"one",4,"four","other")</f><v>one</v></c></row>'
        + '<row r="2"><c r="A2"><v>4</v></c><c r="B2"><v>5</v></c><c r="C2"><v>6</v></c></row>'
        + '<row r="3"><c r="A3"><v>7</v></c><c r="B3"><v>8</v></c><c r="C3"><v>9</v></c></row>'
        + '<row r="4"><c r="B4"><f>D4:D5&amp;"!"</f><v>Excel!</v></c><c r="D4" t="inlineStr"><is><t>Excel</t></is></c><c r="F4"><f>FIND("e",D4:D5)</f><v>4</v></c><c r="H4"><f>SEARCH("e",D4:D5,2)</f><v>4</v></c><c r="J4"><f>SUBSTITUTE(D4:D5,"e","x")</f><v>Excxl</v></c><c r="L4"><f>REPLACE(D4:D5,2,1,"Z")</f><v>EZcel</v></c></row>'
        + '<row r="5"><c r="D5" t="inlineStr"><is><t>😀Excel</t></is></c></row>'
        + '<row r="9"><c r="A9" t="b"><v>1</v></c><c r="C9"><f>IF(A9:A11,"yes","no")</f><v>yes</v></c><c r="E9"><f>IFERROR(A9:A11,"recovered")</f><v>1</v></c><c r="G9"><f>IFNA(A9:A11,"missing")</f><v>1</v></c><c r="I9"><f>NOT(A9:A11)</f><v>0</v></c></row>'
        + '<row r="10"><c r="A10" t="b"><v>0</v></c></row>'
        + '<row r="11"><c r="A11" t="e"><v>#N/A</v></c></row>'
        + '<row r="13"><c r="D13"><f>A1:A3+B1:B3</f><v>3</v></c><c r="F13"><f>A1:A3*2</f><v>2</v></c><c r="H13"><f>A1:A3&gt;=4</f><v>0</v></c><c r="J13"><f>-A1:A3</f><v>-1</v></c><c r="L13"><f>SUM(D13#)</f><v>27</v></c></row>'
        + '<row r="7"><c r="D7" t="inlineStr"><is><t>12.5</t></is></c><c r="F7"><f>VALUE(D7:D8)</f><v>12.5</v></c><c r="H7"><f>TEXT(F7#,"0.00")</f><v>12.50</v></c><c r="J7"><f>ROUND(F7#,0)</f><v>13</v></c><c r="L7"><f>MOD(F7#,10)</f><v>2.5</v></c></row>'
        + '<row r="8"><c r="D8" t="inlineStr"><is><t>1,234.5</t></is></c></row>'
        + '</sheetData></worksheet>'
    };
    await fs.writeFile(file, zipSync(Object.fromEntries(Object.entries(parts).map(([name, value]) => [name, new TextEncoder().encode(value)]))));
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.exception?.description || response.exceptionDetails.text);
        return response.result?.value;
      };
      const waitForCell = async (row, column, value, message) => {
        const selector = `#cell-${row}-${column}`;
        const deadline = Date.now() + 5000;
        while (Date.now() < deadline && await evaluate(`document.querySelector(${JSON.stringify(selector)})?.textContent.trim()`) !== String(value)) await new Promise(resolve => setTimeout(resolve, 25));
        assert.equal(await evaluate(`document.querySelector(${JSON.stringify(selector)})?.textContent.trim()`), String(value), message);
      };
      await waitForCell(0, 3, 4, 'INDEX row result renders in its anchor cell');
      await waitForCell(0, 4, 5, 'INDEX row result spills to the next column');
      await waitForCell(0, 5, 6, 'INDEX row result spills across the selected source row');
      await waitForCell(0, 6, 15, 'SUM reads the spilled INDEX row');
      await waitForCell(0, 8, 2, 'INDEX column result renders in its anchor cell');
      await waitForCell(1, 8, 5, 'INDEX column result spills to the next row');
      await waitForCell(2, 8, 8, 'INDEX column result spills through the selected source column');
      await waitForCell(0, 10, 15, 'SUM reads the spilled INDEX column');
      await waitForCell(0, 11, '1', 'TRIM converts the first range value and renders its anchor');
      await waitForCell(1, 11, '4', 'TRIM spills the second range value');
      await waitForCell(2, 11, '7', 'TRIM spills through the end of the referenced range');
      await waitForCell(0, 13, 'low', 'CHOOSE renders the first array branch');
      await waitForCell(1, 13, 'low', 'CHOOSE broadcasts its scalar branch over a range');
      await waitForCell(2, 13, 'high', 'CHOOSE spills the alternate array branch');
      await waitForCell(0, 14, 'small', 'IFS renders the first matching range condition');
      await waitForCell(1, 14, 'medium', 'IFS evaluates the next condition for each array item');
      await waitForCell(2, 14, 'large', 'IFS spills the fallback condition');
      await waitForCell(0, 15, 'one', 'SWITCH renders the first matching expression');
      await waitForCell(1, 15, 'four', 'SWITCH selects a different branch per array item');
      await waitForCell(2, 15, 'other', 'SWITCH broadcasts its default branch');
      await waitForCell(3, 5, 4, 'FIND returns the first range result');
      await waitForCell(4, 5, 5, 'FIND spills the Unicode-aware second range result');
      await waitForCell(3, 7, 4, 'SEARCH applies the scalar start position to the first range value');
      await waitForCell(4, 7, 2, 'SEARCH spills the scalar-start second range result');
      await waitForCell(3, 9, 'Excxl', 'SUBSTITUTE maps text edits over the referenced range');
      await waitForCell(4, 9, '😀Excxl', 'SUBSTITUTE spills the Unicode source result');
      await waitForCell(3, 11, 'EZcel', 'REPLACE maps character positions over the referenced range');
      await waitForCell(4, 11, '😀Zxcel', 'REPLACE counts Unicode characters before spilling');
      await waitForCell(6, 5, 12.5, 'VALUE spills the numeric value of the first text cell');
      await waitForCell(7, 5, 1234.5, 'VALUE spills the grouped numeric text conversion');
      await waitForCell(6, 7, '12.50', 'TEXT formats the first VALUE spill');
      await waitForCell(7, 7, '1234.50', 'TEXT maps its format across a spill reference');
      await waitForCell(6, 9, 13, 'ROUND maps over the VALUE spill reference');
      await waitForCell(7, 9, 1235, 'ROUND spills the second calculated value');
      await waitForCell(6, 11, 2.5, 'MOD maps over the VALUE spill reference');
      await waitForCell(7, 11, 4.5, 'MOD spills the second remainder');
      await waitForCell(8, 2, 'yes', 'IF selects from the first range condition');
      await waitForCell(9, 2, 'no', 'IF spills the false branch for the next range condition');
      await waitForCell(10, 2, '#N/A', 'IF preserves a condition error in its result array');
      await waitForCell(8, 4, 'TRUE', 'IFERROR spills the first non-error value');
      await waitForCell(10, 4, 'recovered', 'IFERROR handles an error in one array position');
      await waitForCell(10, 6, 'missing', 'IFNA handles #N/A in one array position');
      await waitForCell(8, 8, 'FALSE', 'NOT spills the inverse of the first Boolean');
      await waitForCell(9, 8, 'TRUE', 'NOT spills the inverse of the second Boolean');
      await waitForCell(3, 1, 'Excel!', 'string concatenation spills its first range result');
      await waitForCell(4, 1, '😀Excel!', 'string concatenation spills its Unicode range result');
      await waitForCell(12, 3, 3, 'range addition renders its anchor');
      await waitForCell(13, 3, 9, 'range addition spills its second row');
      await waitForCell(14, 3, 15, 'range addition spills its final row');
      await waitForCell(12, 5, 2, 'range multiplication broadcasts a scalar');
      await waitForCell(14, 5, 14, 'range multiplication spills the final row');
      await waitForCell(12, 7, 'FALSE', 'range comparison spills a Boolean result');
      await waitForCell(14, 7, 'TRUE', 'range comparison spills later Boolean results');
      await waitForCell(12, 9, -1, 'unary negation spills from a range reference');
      await waitForCell(14, 9, -7, 'unary negation spills the final range result');
      await waitForCell(12, 11, 27, 'SUM reads the range arithmetic spill reference');

      const sourceCell = await evaluate(`(()=>{const rect=document.querySelector('#cell-1-0').getBoundingClientRect();return{x:rect.left+rect.width/2,y:rect.top+rect.height/2}})()`);
      await devtools.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: sourceCell.x + devtools.frameOffset.x, y: sourceCell.y + devtools.frameOffset.y, button: 'left', buttons: 1, clickCount: 1 });
      await devtools.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: sourceCell.x + devtools.frameOffset.x, y: sourceCell.y + devtools.frameOffset.y, button: 'left', buttons: 0, clickCount: 1 });
      await evaluate(`(()=>{const input=document.querySelector('#formula');input.value='40';input.dispatchEvent(new Event('input',{bubbles:true}));input.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}))})()`);
      await waitForCell(0, 3, 40, 'editing the source recalculates the INDEX row');
      await waitForCell(0, 6, 51, 'the formula using the row spill recalculates');
      await waitForCell(1, 11, '40', 'editing the source recalculates the TRIM spill');
      await waitForCell(1, 13, 'high', 'editing the source recalculates CHOOSE array branches');
      await waitForCell(1, 14, 'large', 'editing the source recalculates IFS array conditions');
      await waitForCell(1, 15, 'other', 'editing the source recalculates SWITCH array matching');
      await vscode.commands.executeCommand('workbench.action.files.save');
      const saved = analyze(await fs.readFile(file));
      assert.equal(saved.sheets[0].cells.get('A2').value, '40');
      assert.equal(saved.sheets[0].cells.get('D1').formula, 'INDEX(A1:C3,2,0)');
      assert.match(new TextDecoder().decode(saved.files['xl/workbook.xml']), /forceFullCalc="1"/);

      await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools.close();
      devtools = await waitForSpreadsheetWebview();
      await waitForCell(0, 3, 40, 'the saved INDEX array reopens with its recalculated anchor');
      await waitForCell(0, 6, 51, 'the saved spill reference reopens with the recalculated total');
      await waitForCell(2, 8, 8, 'the column INDEX array remains available after reopening');
      await waitForCell(1, 11, '40', 'the saved TRIM spill reopens with the recalculated source');
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium edits and reopens a LibreOffice workbook with mixed table features', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-libreoffice-mixed-host-'));
    const file = path.join(directory, 'mixed-features.xlsx');
    await fs.copyFile(path.join(__dirname, '..', 'fixtures', 'LibreOffice-Mixed-Features-Roundtrip.xlsx'), file);
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.exception?.description || response.exceptionDetails.text);
        return response.result?.value;
      };
      const waitFor = async (expression, message) => {
        const deadline = Date.now() + 8000;
        while (Date.now() < deadline && !await evaluate(expression)) await new Promise(resolve => setTimeout(resolve, 25));
        if (!await evaluate(expression)) assert.fail(message + ': ' + JSON.stringify(await evaluate('({cell:document.querySelector("#cell-1-2")?.textContent,rendered:[...document.querySelectorAll(".cell")].slice(0,12).map(cell=>[cell.id,cell.textContent]),notice:document.querySelector("#notice")?.textContent,readOnly:document.querySelector("#grid")?.getAttribute("aria-readonly")})')));
      };
      await waitFor('document.querySelector("#cell-1-2")?.textContent === "1500"', 'LibreOffice table data renders in the editable grid');
      const point = await evaluate('(()=>{const r=document.querySelector("#cell-1-2").getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+r.height/2}})()');
      await devtools.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: point.x + devtools.frameOffset.x, y: point.y + devtools.frameOffset.y, button: 'left', buttons: 1, clickCount: 1 });
      await devtools.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: point.x + devtools.frameOffset.x, y: point.y + devtools.frameOffset.y, button: 'left', buttons: 0, clickCount: 1 });
      await evaluate(`(()=>{const input=document.querySelector('#formula');input.value='1600';input.dispatchEvent(new Event('input',{bubbles:true}));input.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}))})()`);
      await waitFor('document.querySelector("#cell-1-2")?.textContent === "1600"', 'the table cell edit is applied');
      await vscode.commands.executeCommand('workbench.action.files.save');
      let saved = analyze(await fs.readFile(file));
      assert.equal(saved.sheets[0].cells.get('C2').value, '1600');
      assert.equal(saved.sheets[0].tables[0].name, 'Inventory');
      assert.equal(saved.sheets[0].dataValidations.length, 1);
      assert.equal(saved.sheets[0].conditionalFormats.length, 1);
      assert.equal(saved.sheets[0].charts.length, 1);

      await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools.close();
      devtools = await waitForSpreadsheetWebview();
      await waitFor('document.querySelector("#cell-1-2")?.textContent === "1600"', 'the saved LibreOffice workbook reopens with the edited value');
      saved = analyze(await fs.readFile(file));
      assert.equal(saved.writable, true, saved.readOnlyReason);
      assert.equal(saved.sheets[0].charts.length, 1);
      assert.deepEqual(saved.sheets[0].freezePane, { rows: 1, columns: 1 });
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('activate command and open CSV in the registered Custom Editor', async () => {
    const extension = vscode.extensions.getExtension('jjeabydev.jinsheet');
    assert.ok(extension, 'JinSheet extension is registered');
    await extension.activate();
    const commands = await vscode.commands.getCommands(true);
    assert.ok(commands.includes('jinsheet.open'));
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-host-'));
    const file = path.join(directory, 'host-smoke.csv');
    await fs.writeFile(file, 'name,value\nalpha,1\n', 'utf8');
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      await new Promise(resolve => setTimeout(resolve, 800));
      const tab = vscode.window.tabGroups.activeTabGroup.activeTab;
      assert.ok(tab, 'custom editor tab opened');
      assert.equal(tab.input.viewType, 'jinsheet.editor');
      assert.equal(normalizePath(tab.input.uri.fsPath), normalizePath(file));
    } finally {
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview can collapse the command menu while keeping the formula and find controls visible', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-toolbar-collapse-'));
    const file = path.join(directory, 'toolbar.csv');
    await fs.writeFile(file, 'Name,Value\nAlpha,1\n', 'utf8');
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.text);
        return response.result?.value;
      };
      assert.deepEqual(await evaluate(`(() => ({commandsHidden:document.querySelector('#toolbarCommands').hidden,expanded:document.querySelector('#toggleToolbar').getAttribute('aria-expanded'),label:document.querySelector('#toggleToolbar').textContent.trim()}))()`), {
        commandsHidden: true, expanded: 'false', label: 'Show menu'
      }, 'the first view starts with the sheet taking priority over the command menu');

      await evaluate('document.querySelector("#toggleToolbar").click()');
      assert.deepEqual(await evaluate(`(() => ({commandsHidden:document.querySelector('#toolbarCommands').hidden,expanded:document.querySelector('#toggleToolbar').getAttribute('aria-expanded'),label:document.querySelector('#toggleToolbar').textContent.trim(),formulaVisible:document.querySelector('#formula').getClientRects().length>0,findVisible:document.querySelector('#searchGroup').getClientRects().length>0}))()`), {
        commandsHidden: false, expanded: 'true', label: 'Hide menu', formulaVisible: true, findVisible: true
      }, 'expanding the menu keeps the formula bar and find controls available');
      assert.equal(await evaluate('getComputedStyle(document.querySelector("#toolbarCommands")).display'), 'flex', 'the menu keeps its wrapping layout while expanded');

      await evaluate('document.querySelector("#toggleToolbar").click()');
      assert.deepEqual(await evaluate(`(() => ({commandsHidden:document.querySelector('#toolbarCommands').hidden,expanded:document.querySelector('#toggleToolbar').getAttribute('aria-expanded'),label:document.querySelector('#toggleToolbar').textContent.trim(),formulaVisible:document.querySelector('#formula').getClientRects().length>0,findVisible:document.querySelector('#searchGroup').getClientRects().length>0}))()`), {
        commandsHidden: true, expanded: 'false', label: 'Show menu', formulaVisible: true, findVisible: true
      }, 'collapsing the menu leaves the formula bar and find controls available');

      await evaluate('document.querySelector("#toggleToolbar").click()');
      assert.equal(await evaluate('!document.querySelector("#toolbarCommands").hidden && document.querySelector("#toggleToolbar").getAttribute("aria-expanded")==="true"'), true, 'the menu can be expanded again');
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview opens safe external XLSX hyperlinks and navigates internal links', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-hyperlink-host-')), file = path.join(directory, 'links.xlsx');
    const encode = text => new TextEncoder().encode(text);
    const bytes = zipSync({
      '[Content_Types].xml': encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'),
      '_rels/.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
      'xl/workbook.xml': encode('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="r1"/></sheets></workbook>'),
      'xl/_rels/workbook.xml.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'),
      'xl/worksheets/sheet1.xml': encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:C2"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Source</t></is></c><c r="B1" t="inlineStr"><is><t>Jump</t></is></c><c r="C1" t="inlineStr"><is><t>Unsafe</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Target</t></is></c></row></sheetData><hyperlinks><hyperlink ref="A1" r:id="rId1" display="Friendly label" tooltip="Open source"/><hyperlink ref="B1" location="$A$2"/><hyperlink ref="C1" r:id="rId2"/></hyperlinks></worksheet>'),
      'xl/worksheets/_rels/sheet1.xml.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/source" TargetMode="External"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="javascript:alert(1)" TargetMode="External"/></Relationships>')
    });
    await fs.writeFile(file, bytes);
    const originalOpen = vscode.env.openExternal, opened = [];
    let devtools;
    vscode.env.openExternal = async target => { opened.push(target.toString()); return true; };
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.text);
        return response.result?.value;
      };
      const wait = async expression => {
        const deadline = Date.now() + 5000;
        while (Date.now() < deadline) { if (await evaluate(expression)) return; await new Promise(resolve => setTimeout(resolve, 25)); }
        assert.fail('Hyperlink condition failed: ' + expression);
      };
      await wait('document.querySelector("#cell-0-0")?.textContent==="Friendly label"&&document.querySelector("#cell-0-0")?.classList.contains("hyperlink")');
      assert.equal(await evaluate('getComputedStyle(document.querySelector("#cell-0-0")).textDecorationLine'), 'underline');
      assert.match(await evaluate('document.querySelector("#cell-0-0")?.title'), /Open source/);
      await evaluate('document.querySelector("#cell-0-0").dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true,ctrlKey:true}))');
      const externalDeadline = Date.now() + 5000;
      while (Date.now() < externalDeadline && opened.length === 0) await new Promise(resolve => setTimeout(resolve, 25));
      assert.deepEqual(opened, ['https://example.test/source']);
      await evaluate('document.querySelector("#cell-0-1").dispatchEvent(new MouseEvent("click",{bubbles:true,ctrlKey:true}))');
      await wait('document.querySelector("#address")?.value==="A2"');
      devtools.close();
      devtools = await waitForSpreadsheetWebview(15000, 'Unsafe');
      await wait('Boolean(document.querySelector("#cell-0-2"))');
      await evaluate('document.querySelector("#cell-0-2").dispatchEvent(new MouseEvent("click",{bubbles:true,ctrlKey:true}))');
      await wait('/hyperlink|하이퍼링크/i.test(document.querySelector("#notice")?.textContent||"")');
      assert.deepEqual(opened, ['https://example.test/source'], 'unsafe schemes are rejected before calling VS Code');
    } finally {
      vscode.env.openExternal = originalOpen;
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('JinSheet keeps undo and redo histories separate for simultaneously open documents', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-multi-document-history-'));
    const firstFile = path.join(directory, 'first.csv'), secondFile = path.join(directory, 'second.csv');
    await fs.writeFile(firstFile, 'Name,Value\nAlpha,1\n', 'utf8');
    await fs.writeFile(secondFile, 'Name,Value\nBeta,2\n', 'utf8');
    let devtools;
    const connectToCell = async value => {
      devtools = await waitForSpreadsheetWebview(15000, value);
      return async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.exception?.description || response.exceptionDetails.text);
        return response.result?.value;
      };
    };
    const editFirstCell = async (evaluate, value) => {
      const point = await evaluate(`(() => {const r=document.querySelector('#cell-1-0').getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+r.height/2}})()`);
      const sendMouse = params => devtools.send('Input.dispatchMouseEvent', { ...params, x: params.x + devtools.frameOffset.x, y: params.y + devtools.frameOffset.y });
      await sendMouse({ type: 'mousePressed', ...point, button: 'left', buttons: 1, clickCount: 2 });
      await sendMouse({ type: 'mouseReleased', ...point, button: 'left', buttons: 0, clickCount: 2 });
      const deadline = Date.now() + 3000;
      while (Date.now() < deadline && !await evaluate('Boolean(document.querySelector("#cell-1-0 textarea.cell-editor"))')) await new Promise(resolve => setTimeout(resolve, 25));
      const editorState = await evaluate('(()=>{const input=document.querySelector("#cell-1-0 textarea.cell-editor");return{value:input?.value,selectionStart:input?.selectionStart,selectionEnd:input?.selectionEnd,active:document.activeElement===input}})()');
      assert.equal(editorState.active, true);
      await evaluate('document.querySelector("#cell-1-0 textarea.cell-editor").setSelectionRange(0,document.querySelector("#cell-1-0 textarea.cell-editor").value.length)');
      await devtools.send('Input.insertText', { text: value });
      await devtools.send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13 });
      await devtools.send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13 });
      const committed = Date.now() + 3000;
      while (Date.now() < committed && await evaluate('document.querySelector("#cell-1-0")?.textContent') !== value) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('document.querySelector("#cell-1-0")?.textContent'), value);
      devtools.close();
      devtools = null;
    };
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(firstFile), 'jinsheet.editor');
      await new Promise(resolve => setTimeout(resolve, 400));
      await editFirstCell(await connectToCell('Alpha'), 'Alpha edited');

      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(secondFile), 'jinsheet.editor');
      await new Promise(resolve => setTimeout(resolve, 400));
      await editFirstCell(await connectToCell('Beta'), 'Beta edited');
      assert.equal(normalizePath(vscode.window.tabGroups.activeTabGroup.activeTab.input.uri.fsPath), normalizePath(secondFile));
      await vscode.commands.executeCommand('jinsheet.undo');
      devtools = await waitForSpreadsheetWebview(15000, 'Beta');
      let evaluate = expression => devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true }).then(response => response.result?.value);
      assert.equal(await evaluate('document.querySelector("#cell-1-0")?.textContent'), 'Beta');
      await vscode.commands.executeCommand('jinsheet.redo');
      await new Promise(resolve => setTimeout(resolve, 100));
      devtools.close(); devtools = null;

      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(firstFile), 'jinsheet.editor');
      await new Promise(resolve => setTimeout(resolve, 400));
      devtools = await waitForSpreadsheetWebview(15000, 'Alpha edited');
      evaluate = expression => devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true }).then(response => response.result?.value);
      assert.equal(normalizePath(vscode.window.tabGroups.activeTabGroup.activeTab.input.uri.fsPath), normalizePath(firstFile));
      await vscode.commands.executeCommand('jinsheet.undo');
      await new Promise(resolve => setTimeout(resolve, 100));
      assert.equal(await evaluate('document.querySelector("#cell-1-0")?.textContent'), 'Alpha');
      await vscode.commands.executeCommand('jinsheet.redo');
      const redoDeadline = Date.now() + 3000;
      while (Date.now() < redoDeadline && await evaluate('document.querySelector("#cell-1-0")?.textContent') !== 'Alpha edited') await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('document.querySelector("#cell-1-0")?.textContent'), 'Alpha edited');
      devtools.close(); devtools = null;

      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(secondFile), 'jinsheet.editor');
      await new Promise(resolve => setTimeout(resolve, 400));
      devtools = await waitForSpreadsheetWebview(15000, 'Beta edited');
      evaluate = expression => devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true }).then(response => response.result?.value);
      assert.equal(await evaluate('document.querySelector("#cell-1-0")?.textContent'), 'Beta edited', 'undo in the first document leaves the second document history untouched');
      assert.equal(await fs.readFile(firstFile, 'utf8'), 'Name,Value\nAlpha,1\n');
      assert.equal(await fs.readFile(secondFile, 'utf8'), 'Name,Value\nBeta,2\n');
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview inserts static date and time with Excel keyboard shortcuts', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-static-datetime-'));
    const file = path.join(directory, 'datetime.csv');
    await fs.writeFile(file, 'Date,Time\n,\n', 'utf8');
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.exception?.description || response.exceptionDetails.text);
        return response.result?.value;
      };
      const waitFor = async (expression, message) => {
        const deadline = Date.now() + 3000;
        while (Date.now() < deadline && !await evaluate(expression)) await new Promise(resolve => setTimeout(resolve, 25));
        assert.ok(await evaluate(expression), message);
      };
      const press = async (key, code, windowsVirtualKeyCode, modifiers = 0) => {
        await devtools.send('Input.dispatchKeyEvent', { type: 'keyDown', key, code, windowsVirtualKeyCode, modifiers });
        await devtools.send('Input.dispatchKeyEvent', { type: 'keyUp', key, code, windowsVirtualKeyCode, modifiers });
      };
      await waitFor('Boolean(document.querySelector("#cell-0-0"))', 'the date cell is rendered');
      await press(';', 'Semicolon', 186, 2);
      await waitFor('/^\\d{4}-\\d{2}-\\d{2}$/.test(document.querySelector("#cell-0-0")?.textContent)', 'Ctrl+; inserts the current static date');
      await press('ArrowRight', 'ArrowRight', 39);
      await press(':', 'Semicolon', 186, 10);
      await waitFor('/^\\d{2}:\\d{2}$/.test(document.querySelector("#cell-0-1")?.textContent)', 'Ctrl+Shift+; inserts the current static time');
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview moves the active cell with arrows, Tab and Enter', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-keyboard-navigation-'));
    const file = path.join(directory, 'keyboard.csv');
    await fs.writeFile(file, 'A,B,C\n1,2,3\n4,5,6\n', 'utf8');
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.text);
        return response.result?.value;
      };
      const waitForAddress = async address => {
        const deadline = Date.now() + 3000;
        while (Date.now() < deadline && await evaluate('document.querySelector("#address")?.value') !== address) await new Promise(resolve => setTimeout(resolve, 25));
        assert.equal(await evaluate('document.querySelector("#address")?.value'), address);
      };
      const press = async (key, code, windowsVirtualKeyCode, modifiers = 0) => {
        await devtools.send('Input.dispatchKeyEvent', { type: 'keyDown', key, code, windowsVirtualKeyCode, modifiers });
        await devtools.send('Input.dispatchKeyEvent', { type: 'keyUp', key, code, windowsVirtualKeyCode, modifiers });
      };
      const renderDeadline = Date.now() + 5000;
      while (Date.now() < renderDeadline && !await evaluate('Boolean(document.querySelector("#cell-0-0"))')) await new Promise(resolve => setTimeout(resolve, 25));
      assert.ok(await evaluate('Boolean(document.querySelector("#cell-0-0"))'), 'initial grid cells render before measuring the selected cell');
      const origin = await evaluate(`(() => {const r=document.querySelector('#cell-0-0').getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+r.height/2}})()`);
      await devtools.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: origin.x + devtools.frameOffset.x, y: origin.y + devtools.frameOffset.y, button: 'left', buttons: 1, clickCount: 1 });
      await devtools.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: origin.x + devtools.frameOffset.x, y: origin.y + devtools.frameOffset.y, button: 'left', buttons: 0, clickCount: 1 });
      await waitForAddress('A1');

      await press('ArrowRight', 'ArrowRight', 39);
      await waitForAddress('B1');
      await press('ArrowDown', 'ArrowDown', 40);
      await waitForAddress('B2');
      await press('ArrowLeft', 'ArrowLeft', 37, 2);
      await waitForAddress('A2');
      await press('ArrowRight', 'ArrowRight', 39, 2);
      await waitForAddress('C2');

      await press('ArrowUp', 'ArrowUp', 38, 8);
      await waitForAddress('C1');
      assert.equal(await evaluate('document.querySelectorAll(".cell[aria-selected=true]").length'), 2, 'Shift+ArrowUp extends the selection by one cell');
      await press('ArrowLeft', 'ArrowLeft', 37);
      await waitForAddress('B1');
      assert.equal(await evaluate('document.querySelectorAll(".cell[aria-selected=true]").length'), 1, 'an unmodified arrow collapses the extended range to the active cell');

      await press('Tab', 'Tab', 9);
      await waitForAddress('C1');
      await press('Tab', 'Tab', 9, 8);
      await waitForAddress('B1');
      await press('Enter', 'Enter', 13);
      await waitForAddress('B2');
      await press('Enter', 'Enter', 13, 8);
      await waitForAddress('B1');

      await press('F2', 'F2', 113);
      const editDeadline = Date.now() + 3000;
      while (Date.now() < editDeadline && !await evaluate('Boolean(document.querySelector("#cell-0-1 textarea.cell-editor"))')) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('document.querySelector("#cell-0-1 textarea.cell-editor")?.value'), 'B', 'F2 edits the active cell in place');
      await press('a', 'KeyA', 65, 2);
      await devtools.send('Input.insertText', { text: 'discard' });
      await press('Escape', 'Escape', 27);
      const cancelDeadline = Date.now() + 3000;
      while (Date.now() < cancelDeadline && await evaluate('Boolean(document.querySelector("#cell-0-1 textarea.cell-editor"))')) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('document.querySelector("#cell-0-1")?.textContent'), 'B', 'Escape cancels the uncommitted edit');

      await press('z', 'KeyZ', 90);
      assert.equal(await evaluate('document.querySelector("#cell-0-1 textarea.cell-editor")?.value'), 'z', 'typing replaces the active cell and starts editing');
      await press('Enter', 'Enter', 13);
      await waitForAddress('B2');
      const commitDeadline = Date.now() + 3000;
      while (Date.now() < commitDeadline && await evaluate('document.querySelector("#cell-0-1")?.textContent') !== 'z') await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('document.querySelector("#cell-0-1")?.textContent'), 'z', 'Enter commits the typed replacement');
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview wraps Tab and Shift+Tab inside a multi-cell selection', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-range-tab-navigation-'));
    const file = path.join(directory, 'range-tab.csv');
    await fs.writeFile(file, 'A,B,C\n1,2,3\n4,5,6\n', 'utf8');
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.text);
        return response.result?.value;
      };
      const waitForAddress = async address => {
        const deadline = Date.now() + 3000;
        while (Date.now() < deadline && await evaluate('document.querySelector("#address")?.value') !== address) await new Promise(resolve => setTimeout(resolve, 25));
        assert.equal(await evaluate('document.querySelector("#address")?.value'), address);
      };
      const press = async (key, code, windowsVirtualKeyCode, modifiers = 0) => {
        await devtools.send('Input.dispatchKeyEvent', { type: 'keyDown', key, code, windowsVirtualKeyCode, modifiers });
        await devtools.send('Input.dispatchKeyEvent', { type: 'keyUp', key, code, windowsVirtualKeyCode, modifiers });
      };
      const origin = await evaluate(`(() => {const r=document.querySelector('#cell-0-1').getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+r.height/2}})()`);
      await devtools.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: origin.x + devtools.frameOffset.x, y: origin.y + devtools.frameOffset.y, button: 'left', buttons: 1, clickCount: 1 });
      await devtools.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: origin.x + devtools.frameOffset.x, y: origin.y + devtools.frameOffset.y, button: 'left', buttons: 0, clickCount: 1 });
      await waitForAddress('B1');
      await press('ArrowRight', 'ArrowRight', 39, 8);
      await waitForAddress('C1');
      await press('ArrowDown', 'ArrowDown', 40, 8);
      await waitForAddress('C2');
      assert.equal(await evaluate('document.querySelectorAll(".cell[aria-selected=true]").length'), 4, 'Shift+Arrow selects B1:C2');
      await press('Tab', 'Tab', 9);
      await waitForAddress('B1');
      assert.equal(await evaluate('document.querySelectorAll(".cell[aria-selected=true]").length'), 4, 'Tab keeps the selected range while wrapping to its first cell');
      await press('Tab', 'Tab', 9);
      await waitForAddress('C1');
      await press('Tab', 'Tab', 9, 8);
      await waitForAddress('B1');
      assert.equal(await evaluate('document.querySelectorAll(".cell[aria-selected=true]").length'), 4, 'Shift+Tab moves backward within the same range');
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview clears a selected range as one saved undoable operation', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-range-clear-history-'));
    const file = path.join(directory, 'range-clear.csv');
    const original = 'A,B\n1,2\n';
    await fs.writeFile(file, original, 'utf8');
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.text);
        return response.result?.value;
      };
      const waitFor = async (expression, message) => {
        const deadline = Date.now() + 3000;
        while (Date.now() < deadline && !await evaluate(expression)) await new Promise(resolve => setTimeout(resolve, 25));
        assert.ok(await evaluate(expression), message);
      };
      const pointFor = cellId => evaluate(`(() => {const r=document.querySelector('#${cellId}').getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+r.height/2}})()`);
      const dispatchMouse = params => devtools.send('Input.dispatchMouseEvent', { ...params, x: params.x + devtools.frameOffset.x, y: params.y + devtools.frameOffset.y });
      const press = async (key, code, windowsVirtualKeyCode) => {
        await devtools.send('Input.dispatchKeyEvent', { type: 'keyDown', key, code, windowsVirtualKeyCode });
        await devtools.send('Input.dispatchKeyEvent', { type: 'keyUp', key, code, windowsVirtualKeyCode });
      };
      await waitFor('Boolean(document.querySelector("#cell-1-0")) && Boolean(document.querySelector("#cell-1-1"))', 'the source range is rendered');

      const singleCell = await pointFor('cell-1-0');
      await dispatchMouse({ type: 'mousePressed', ...singleCell, button: 'left', buttons: 1, clickCount: 1 });
      await dispatchMouse({ type: 'mouseReleased', ...singleCell, button: 'left', buttons: 0, clickCount: 1 });
      await press('Backspace', 'Backspace', 8);
      await waitFor('Boolean(document.querySelector("#cell-1-0 textarea.cell-editor")) && document.querySelector("#cell-1-0 textarea.cell-editor").value === ""', 'Backspace enters an empty inline edit for one cell');
      await press('Escape', 'Escape', 27);
      await waitFor('!document.querySelector("#cell-1-0 textarea.cell-editor") && document.querySelector("#cell-1-0")?.textContent === "1"', 'Escape cancels Backspace edit without deleting the value');

      const start = await pointFor('cell-1-0'), end = await pointFor('cell-1-1');
      await dispatchMouse({ type: 'mousePressed', ...start, button: 'left', buttons: 1, clickCount: 1 });
      await dispatchMouse({ type: 'mouseMoved', ...end, button: 'left', buttons: 1 });
      await dispatchMouse({ type: 'mouseReleased', ...end, button: 'left', buttons: 0, clickCount: 1 });
      assert.equal(await evaluate('document.querySelectorAll(".cell[aria-selected=true]").length'), 2, 'drag selects both cells');
      await press('Delete', 'Delete', 46);
      await waitFor('document.querySelector("#cell-1-0")?.textContent === "" && document.querySelector("#cell-1-1")?.textContent === ""', 'Delete clears both cells in the selected range');
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(await fs.readFile(file, 'utf8'), 'A,B\n,\n', 'one range clear is saved to CSV');

      await vscode.commands.executeCommand('jinsheet.undo');
      await waitFor('document.querySelector("#cell-1-0")?.textContent === "1" && document.querySelector("#cell-1-1")?.textContent === "2"', 'one undo restores the entire cleared range');
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(await fs.readFile(file, 'utf8'), original, 'undo saves both original values');

      await vscode.commands.executeCommand('jinsheet.redo');
      await waitFor('document.querySelector("#cell-1-0")?.textContent === "" && document.querySelector("#cell-1-1")?.textContent === ""', 'one redo clears the entire range again');
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(await fs.readFile(file, 'utf8'), 'A,B\n,\n', 'redo saves both cleared values');
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview moves a cut range as one saved undoable operation', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-range-move-history-'));
    const file = path.join(directory, 'range-move.csv');
    const original = 'A,B,C,D\na1,b1,,\na2,b2,,\n';
    const moved = 'A,B,C,D\n,,a1,b1\n,,a2,b2\n';
    await fs.writeFile(file, original, 'utf8');
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.text);
        return response.result?.value;
      };
      const waitFor = async (expression, message) => {
        const deadline = Date.now() + 4000;
        while (Date.now() < deadline && !await evaluate(expression)) await new Promise(resolve => setTimeout(resolve, 25));
        assert.ok(await evaluate(expression), message);
      };
      const pointFor = cellId => evaluate(`(() => {const r=document.querySelector('#${cellId}').getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+r.height/2}})()`);
      const dispatchMouse = params => devtools.send('Input.dispatchMouseEvent', { ...params, x: params.x + devtools.frameOffset.x, y: params.y + devtools.frameOffset.y });
      const press = async (key, code, windowsVirtualKeyCode, modifiers = 0) => {
        await devtools.send('Input.dispatchKeyEvent', { type: 'keyDown', key, code, windowsVirtualKeyCode, modifiers });
        await devtools.send('Input.dispatchKeyEvent', { type: 'keyUp', key, code, windowsVirtualKeyCode, modifiers });
      };
      await waitFor('Boolean(document.querySelector("#cell-1-0")) && Boolean(document.querySelector("#cell-2-1"))', 'the move source range is rendered');

      const start = await pointFor('cell-1-0'), end = await pointFor('cell-2-1');
      await dispatchMouse({ type: 'mousePressed', ...start, button: 'left', buttons: 1, clickCount: 1 });
      await dispatchMouse({ type: 'mouseMoved', ...end, button: 'left', buttons: 1 });
      await dispatchMouse({ type: 'mouseReleased', ...end, button: 'left', buttons: 0, clickCount: 1 });
      assert.equal(await evaluate('document.querySelectorAll(".cell[aria-selected=true]").length'), 4, 'drag selects the complete 2-by-2 source range');
      await press('x', 'KeyX', 88, 2);
      await waitFor('document.querySelectorAll(".cell.cut-source").length === 4', 'Ctrl+X marks the selected source range as pending cut');

      const destination = await pointFor('cell-1-2');
      await dispatchMouse({ type: 'mousePressed', ...destination, button: 'left', buttons: 1, clickCount: 1 });
      await dispatchMouse({ type: 'mouseReleased', ...destination, button: 'left', buttons: 0, clickCount: 1 });
      await press('v', 'KeyV', 86, 2);
      await waitFor('document.querySelector("#cell-1-0")?.textContent === "" && document.querySelector("#cell-1-1")?.textContent === "" && document.querySelector("#cell-1-2")?.textContent === "a1" && document.querySelector("#cell-1-3")?.textContent === "b1"', 'Ctrl+V moves the first source row to the destination');
      await waitFor('document.querySelector("#cell-2-0")?.textContent === "" && document.querySelector("#cell-2-1")?.textContent === "" && document.querySelector("#cell-2-2")?.textContent === "a2" && document.querySelector("#cell-2-3")?.textContent === "b2"', 'Ctrl+V moves the entire second source row');
      assert.equal(await evaluate('document.querySelectorAll(".cell.cut-source").length'), 0, 'successful move clears the pending-cut outline');
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(await fs.readFile(file, 'utf8'), moved, 'the full cut range move is saved');

      await vscode.commands.executeCommand('jinsheet.undo');
      await waitFor('document.querySelector("#cell-1-0")?.textContent === "a1" && document.querySelector("#cell-2-1")?.textContent === "b2" && document.querySelector("#cell-1-2")?.textContent === ""', 'one undo restores the source and clears the destination');
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(await fs.readFile(file, 'utf8'), original, 'undo saves the original range layout');

      await vscode.commands.executeCommand('jinsheet.redo');
      await waitFor('document.querySelector("#cell-1-0")?.textContent === "" && document.querySelector("#cell-2-1")?.textContent === "" && document.querySelector("#cell-1-2")?.textContent === "a1"', 'one redo reapplies the range move');
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(await fs.readFile(file, 'utf8'), moved, 'redo saves the moved range');
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview copies XLSX cell values and formatting through save, undo and redo', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-formatted-paste-history-'));
    const file = path.join(directory, 'formatted-paste.xlsx');
    await fs.copyFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx'), file);
    let devtools; const originalWrite = JinSheetProvider.prototype.writeClipboardText;
    JinSheetProvider.prototype.writeClipboardText = function(text) { return originalWrite.call(this, text, async value => { await new Promise(resolve => setTimeout(resolve, 400)); return vscode.env.clipboard.writeText(value); }); };
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.text);
        return response.result?.value;
      };
      const waitFor = async (expression, message) => {
        const deadline = Date.now() + 5000;
        while (Date.now() < deadline && !await evaluate(expression)) await new Promise(resolve => setTimeout(resolve, 25));
        if (!await evaluate(expression)) {
          const state = await evaluate(`(() => {const cell=document.querySelector('#cell-1-4'),data=cell?.__jinSheetData;return{source:document.querySelector('#cell-1-2')?.textContent,target:cell?.textContent,targetDisplayValue:data?.displayValue??'missing',targetValue:data?.value??'missing',targetFormatCode:data?.formatCode??'missing',targetStyle:cell?.getAttribute('style'),selected:document.querySelector('.cell.selected')?.id,address:document.querySelector('#address')?.value,notice:document.querySelector('#notice')?.textContent}})()`);
          assert.fail(message + ': ' + JSON.stringify(state));
        }
      };
      const pointFor = cellId => evaluate(`(() => {const r=document.querySelector('#${cellId}').getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+r.height/2}})()`);
      const dispatchMouse = params => devtools.send('Input.dispatchMouseEvent', { ...params, x: params.x + devtools.frameOffset.x, y: params.y + devtools.frameOffset.y });
      const press = async (key, code, windowsVirtualKeyCode, modifiers = 0) => {
        await devtools.send('Input.dispatchKeyEvent', { type: 'keyDown', key, code, windowsVirtualKeyCode, modifiers });
        await devtools.send('Input.dispatchKeyEvent', { type: 'keyUp', key, code, windowsVirtualKeyCode, modifiers });
      };
      await waitFor('Boolean(document.querySelector("#cell-1-2")) && document.querySelector("#cell-1-2")?.textContent.length > 0 && Boolean(document.querySelector("#cell-1-4"))', 'the styled source and destination cells are rendered');
      const sourceBefore = await evaluate(`(() => {const cell=document.querySelector('#cell-1-2'),style=getComputedStyle(cell);return{text:cell.textContent,formatting:{color:style.color,backgroundColor:style.backgroundColor,fontSize:style.fontSize,fontWeight:style.fontWeight,fontStyle:style.fontStyle,textDecorationLine:style.textDecorationLine,textAlign:style.textAlign,borderTop:style.borderTop,borderRight:style.borderRight,borderBottom:style.borderBottom,borderLeft:style.borderLeft}}})()`);
      assert.ok(sourceBefore.formatting, 'the source cell has visible formatting');

      const source = await pointFor('cell-1-2');
      await dispatchMouse({ type: 'mousePressed', ...source, button: 'left', buttons: 1, clickCount: 1 });
      await dispatchMouse({ type: 'mouseReleased', ...source, button: 'left', buttons: 0, clickCount: 1 });
      await press('c', 'KeyC', 67, 2);
      const destination = await pointFor('cell-1-4');
      await dispatchMouse({ type: 'mousePressed', ...destination, button: 'left', buttons: 1, clickCount: 1 });
      await dispatchMouse({ type: 'mouseReleased', ...destination, button: 'left', buttons: 0, clickCount: 1 });
      await press('v', 'KeyV', 86, 2);
      await waitFor(`document.querySelector('#cell-1-4')?.textContent === ${JSON.stringify(sourceBefore.text)}`, 'Ctrl+V copies the source cell value');
      const pastedFormatting = await evaluate(`(() => {const style=getComputedStyle(document.querySelector('#cell-1-4'));return{color:style.color,backgroundColor:style.backgroundColor,fontSize:style.fontSize,fontWeight:style.fontWeight,fontStyle:style.fontStyle,textDecorationLine:style.textDecorationLine,textAlign:style.textAlign,borderTop:style.borderTop,borderRight:style.borderRight,borderBottom:style.borderBottom,borderLeft:style.borderLeft}})()`);
      assert.deepEqual(pastedFormatting, sourceBefore.formatting, 'Ctrl+V applies the same cell formatting in the Webview');

      await vscode.commands.executeCommand('workbench.action.files.save');
      const saved = analyze(await fs.readFile(file)).sheets[0].cells;
      const sourceSaved = saved.get('C2'), targetSaved = saved.get('E2');
      assert.equal(targetSaved.value, sourceSaved.value, 'saved target cell contains the copied value');
      assert.equal(targetSaved.style, sourceSaved.style, 'saved target cell has the same style index');
      assert.deepEqual(targetSaved.styleProperties, sourceSaved.styleProperties, 'saved target cell preserves formatting properties');

      await vscode.commands.executeCommand('jinsheet.undo');
      await waitFor('document.querySelector("#cell-1-4")?.textContent === ""', 'one undo removes the copied cell and its format');
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(analyze(await fs.readFile(file)).sheets[0].cells.has('E2'), false, 'undo removes the pasted cell from the saved workbook');

      await vscode.commands.executeCommand('jinsheet.redo');
      await waitFor(`document.querySelector('#cell-1-4')?.textContent === ${JSON.stringify(sourceBefore.text)}`, 'one redo restores the copied value');
      await vscode.commands.executeCommand('workbench.action.files.save');
      const redone = analyze(await fs.readFile(file)).sheets[0].cells;
      assert.equal(redone.get('E2')?.value, sourceSaved.value);
      assert.deepEqual(redone.get('E2')?.styleProperties, sourceSaved.styleProperties, 'redo restores the formatting with the value');
    } finally {
      JinSheetProvider.prototype.writeClipboardText = originalWrite;
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview applies a number format to a selected XLSX cell and saves it through undo and redo', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-number-format-host-'));
    const file = path.join(directory, 'number-format.xlsx');
    await fs.copyFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx'), file);
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.text);
        return response.result?.value;
      };
      const cellPoint = await evaluate(`(() => {const r=document.querySelector('#cell-1-2').getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+r.height/2}})()`);
      await devtools.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: cellPoint.x + devtools.frameOffset.x, y: cellPoint.y + devtools.frameOffset.y, button: 'left', buttons: 1, clickCount: 1 });
      await devtools.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: cellPoint.x + devtools.frameOffset.x, y: cellPoint.y + devtools.frameOffset.y, button: 'left', buttons: 0, clickCount: 1 });
      const originalDisplay = await evaluate('document.querySelector("#cell-1-2")?.textContent');
      const original = analyze(await fs.readFile(file)).sheets[0].cells.get('C2');
      assert.equal(original.kind, 'number');
      const saveUntil = async predicate => {
        const deadline = Date.now() + 3000;
        do {
          await vscode.commands.executeCommand('workbench.action.files.save');
          const cell = analyze(await fs.readFile(file)).sheets[0].cells.get('C2');
          if (predicate(cell)) return cell;
          await new Promise(resolve => setTimeout(resolve, 25));
        } while (Date.now() < deadline);
        assert.fail('Number-format save did not reach the expected state');
      };
      await evaluate('document.querySelector("#formatPercent").click()');
      const formattedDisplay = '150000%';
      const deadline = Date.now() + 3000;
      while (Date.now() < deadline && await evaluate('document.querySelector("#cell-1-2")?.textContent') !== formattedDisplay) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('document.querySelector("#cell-1-2")?.textContent'), formattedDisplay);
      await evaluate('document.querySelector("#printPanel").open=true');
      let saved = await saveUntil(cell => cell.numFmtId === 9);
      assert.equal(saved.numFmtId, 9);
      assert.deepEqual(saved.styleProperties, original.styleProperties, 'applying a number format keeps the cell font, fill, border and alignment');
      await vscode.commands.executeCommand('jinsheet.undo');
      const undoDeadline = Date.now() + 3000;
      while (Date.now() < undoDeadline && await evaluate('document.querySelector("#cell-1-2")?.textContent') !== originalDisplay) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('document.querySelector("#cell-1-2")?.textContent'), originalDisplay);
      saved = await saveUntil(cell => cell.numFmtId === original.numFmtId);
      assert.equal(saved.numFmtId, original.numFmtId, 'undo restores the original number format in the file');
      await vscode.commands.executeCommand('jinsheet.redo');
      const redoDeadline = Date.now() + 3000;
      while (Date.now() < redoDeadline && await evaluate('document.querySelector("#cell-1-2")?.textContent') !== formattedDisplay) await new Promise(resolve => setTimeout(resolve, 25));
      saved = await saveUntil(cell => cell.numFmtId === 9);
      assert.equal(saved.numFmtId, 9, 'redo reapplies the selected number format in the file');
      await evaluate('document.querySelector("#formatCurrency").click()');
      const currencyDeadline = Date.now() + 3000;
      while (Date.now() < currencyDeadline && await evaluate('document.querySelector("#cell-1-2")?.textContent') !== '$1,500.00') await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('document.querySelector("#cell-1-2")?.textContent'), '$1,500.00');
      saved = await saveUntil(cell => cell.formatCode === '"$"#,##0.00');
      assert.ok(saved.numFmtId >= 164);
      assert.equal(saved.formatCode, '"$"#,##0.00');
      await vscode.commands.executeCommand('jinsheet.undo');
      const currencyUndoDeadline = Date.now() + 3000;
      while (Date.now() < currencyUndoDeadline && await evaluate('document.querySelector("#cell-1-2")?.textContent') !== formattedDisplay) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('document.querySelector("#cell-1-2")?.textContent'), formattedDisplay);
      await saveUntil(cell => cell.numFmtId === 9);
      await vscode.commands.executeCommand('jinsheet.redo');
      const currencyRedoDeadline = Date.now() + 3000;
      while (Date.now() < currencyRedoDeadline && await evaluate('document.querySelector("#cell-1-2")?.textContent') !== '$1,500.00') await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('document.querySelector("#cell-1-2")?.textContent'), '$1,500.00');
      await saveUntil(cell => cell.formatCode === '"$"#,##0.00');
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview toggles bold on a selected XLSX cell and saves it through undo and redo', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-font-format-host-'));
    const file = path.join(directory, 'font-format.xlsx');
    await fs.copyFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx'), file);
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.text);
        return response.result?.value;
      };
      const cellDeadline = Date.now() + 5000;
      while (Date.now() < cellDeadline && !await evaluate('Boolean(document.querySelector("#cell-1-2"))')) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('Boolean(document.querySelector("#cell-1-2"))'), true);
      const original = analyze(await fs.readFile(file)).sheets[0].cells.get('C2');
      const cellPoint = await evaluate(`(() => {const r=document.querySelector('#cell-1-2').getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+r.height/2}})()`);
      await devtools.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: cellPoint.x + devtools.frameOffset.x, y: cellPoint.y + devtools.frameOffset.y, button: 'left', buttons: 1, clickCount: 1 });
      await devtools.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: cellPoint.x + devtools.frameOffset.x, y: cellPoint.y + devtools.frameOffset.y, button: 'left', buttons: 0, clickCount: 1 });
      assert.equal(await evaluate('document.querySelector("#address")?.value'), 'C2');
      await evaluate('document.querySelector("#formatBold").click()');
      const modifiedDeadline = Date.now() + 3000;
      while (Date.now() < modifiedDeadline && await evaluate('document.querySelector("#notice")?.textContent.includes("Modified")') !== true) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('document.querySelector("#notice")?.textContent.includes("Modified")'), true, 'bold toggle marks the selected workbook modified');
      await vscode.commands.executeCommand('workbench.action.files.save');
      let saved = analyze(await fs.readFile(file)).sheets[0].cells.get('C2');
      assert.equal(saved.styleProperties.bold, true);
      assert.equal(saved.formatCode, original.formatCode, 'bold formatting preserves the number format');
      assert.deepEqual(saved.styleProperties.borders, original.styleProperties.borders, 'bold formatting preserves borders');
      await vscode.commands.executeCommand('jinsheet.undo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file)).sheets[0].cells.get('C2');
      assert.equal(saved.styleProperties.bold, undefined, 'undo restores the original font weight');
      assert.equal(saved.formatCode, original.formatCode);
      await vscode.commands.executeCommand('jinsheet.redo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file)).sheets[0].cells.get('C2');
      assert.equal(saved.styleProperties.bold, true, 'redo reapplies the font weight');
      await evaluate('document.querySelector("#formatUnderline").click()');
      const underlineDeadline = Date.now() + 3000;
      while (Date.now() < underlineDeadline && await evaluate('document.querySelector("#notice")?.textContent.includes("Modified")') !== true) await new Promise(resolve => setTimeout(resolve, 25));
      const underlineExpression = 'String(getComputedStyle(document.querySelector("#cell-1-2")).textDecorationLine).includes("underline")';
      while (Date.now() < underlineDeadline && !await evaluate('Boolean(document.querySelector("#cell-1-2"))&&' + underlineExpression)) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('Boolean(document.querySelector("#cell-1-2"))&&' + underlineExpression), true, 'wait for the underline edit to reach the rendered cell before saving');
      await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file)).sheets[0].cells.get('C2');
      assert.equal(saved.styleProperties.underline, true);
      assert.equal(saved.styleProperties.bold, true);
      await evaluate('document.querySelector("#formatStrike").click()');
      const strikeDeadline = Date.now() + 3000;
      while (Date.now() < strikeDeadline && await evaluate('document.querySelector("#notice")?.textContent.includes("Modified")') !== true) await new Promise(resolve => setTimeout(resolve, 25));
      while (Date.now() < strikeDeadline && !String(await evaluate('(()=>{const cell=document.querySelector("#cell-1-2");return cell?getComputedStyle(cell).textDecorationLine:""})()')).includes('line-through')) await new Promise(resolve => setTimeout(resolve, 25));
      const strikeSaveDeadline = Date.now() + 5000;
      do {
        await vscode.commands.executeCommand('workbench.action.files.save');
        saved = analyze(await fs.readFile(file)).sheets[0].cells.get('C2');
        if (saved.styleProperties.strike === true) break;
        await new Promise(resolve => setTimeout(resolve, 25));
      } while (Date.now() < strikeSaveDeadline);
      assert.equal(saved.styleProperties.strike, true);
      await evaluate('document.querySelector("#formatStrike").click()');
      const clearStrikeDeadline = Date.now() + 3000;
      while (Date.now() < clearStrikeDeadline && vscode.window.tabGroups.activeTabGroup.activeTab?.isDirty !== true) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(vscode.window.tabGroups.activeTabGroup.activeTab?.isDirty, true, 'clearing strikethrough marks the workbook modified before saving');
      while (Date.now() < clearStrikeDeadline && String(await evaluate('(()=>{const cell=document.querySelector("#cell-1-2");return cell?getComputedStyle(cell).textDecorationLine:""})()')).includes('line-through')) await new Promise(resolve => setTimeout(resolve, 25));
      await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file)).sheets[0].cells.get('C2');
      assert.equal(saved.styleProperties.strike, undefined);
      await evaluate('document.querySelector("#fontSize").value="18";document.querySelector("#applyFontSize").click()');
      const sizeDeadline = Date.now() + 3000;
      while (Date.now() < sizeDeadline && await evaluate('(()=>{const cell=document.querySelector("#cell-1-2");return cell?getComputedStyle(cell).fontSize:"missing"})()') !== '24px') await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('(()=>{const cell=document.querySelector("#cell-1-2");return cell?getComputedStyle(cell).fontSize:"missing"})()'), '24px');
      while (Date.now() < sizeDeadline && vscode.window.tabGroups.activeTabGroup.activeTab?.isDirty !== true) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(vscode.window.tabGroups.activeTabGroup.activeTab?.isDirty, true, 'applying a font size marks the workbook modified before saving');
      await vscode.commands.executeCommand('workbench.action.files.save');
      const fontSizeSaveDeadline = Date.now() + 5000;
      while (Date.now() < fontSizeSaveDeadline && analyze(await fs.readFile(file)).sheets[0].cells.get('C2')?.styleProperties.fontSize !== 18) await new Promise(resolve => setTimeout(resolve, 25));
      saved = analyze(await fs.readFile(file)).sheets[0].cells.get('C2');
      assert.equal(saved.styleProperties.fontSize, 18);
      assert.equal(saved.styleProperties.underline, true);
      await vscode.commands.executeCommand('jinsheet.undo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file)).sheets[0].cells.get('C2');
      assert.equal(saved.styleProperties.fontSize, original.styleProperties.fontSize);
      assert.equal(saved.styleProperties.underline, true);
      await vscode.commands.executeCommand('jinsheet.redo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file)).sheets[0].cells.get('C2');
      assert.equal(saved.styleProperties.fontSize, 18);
      await evaluate('document.querySelector("#fontFamily").value="Aptos";document.querySelector("#applyFontFamily").click()');
      const familyDeadline = Date.now() + 3000;
      while (Date.now() < familyDeadline && await evaluate('document.querySelector("#notice")?.textContent.includes("Modified")') !== true) await new Promise(resolve => setTimeout(resolve, 25));
      await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file)).sheets[0].cells.get('C2');
      assert.equal(saved.styleProperties.fontFamily, 'Aptos');
      await evaluate('document.querySelector("#clearFontFamily").click()');
      const defaultFamilyDeadline = Date.now() + 3000;
      while (Date.now() < defaultFamilyDeadline && await evaluate('document.querySelector("#notice")?.textContent.includes("Modified")') !== true) await new Promise(resolve => setTimeout(resolve, 25));
      const defaultFamilySaveDeadline = Date.now() + 5000;
      do {
        await vscode.commands.executeCommand('workbench.action.files.save');
        saved = analyze(await fs.readFile(file)).sheets[0].cells.get('C2');
        if (saved.styleProperties.fontFamily === undefined) break;
        await new Promise(resolve => setTimeout(resolve, 25));
      } while (Date.now() < defaultFamilySaveDeadline);
      assert.equal(saved.styleProperties.fontFamily, undefined);
      await evaluate('document.querySelector("#clearFontSize").click()');
      const clearSizeDeadline = Date.now() + 3000;
      while (Date.now() < clearSizeDeadline && await evaluate('document.querySelector("#notice")?.textContent.includes("Modified")') !== true) await new Promise(resolve => setTimeout(resolve, 25));
      const defaultSizeSaveDeadline = Date.now() + 5000;
      do {
        await vscode.commands.executeCommand('workbench.action.files.save');
        saved = analyze(await fs.readFile(file)).sheets[0].cells.get('C2');
        if (saved.styleProperties.fontSize === undefined) break;
        await new Promise(resolve => setTimeout(resolve, 25));
      } while (Date.now() < defaultSizeSaveDeadline);
      assert.equal(saved.styleProperties.fontSize, undefined);
      assert.equal(saved.styleProperties.underline, true);
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview calculates, edits and structurally moves XLSX shared formulas', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-shared-formula-host-'));
    const file = path.join(directory, 'shared-formula.xlsx');
    const encode = text => new TextEncoder().encode(text);
    const bytes = zipSync({
      '[Content_Types].xml': encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'),
      '_rels/.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
      'xl/workbook.xml': encode('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="r1"/></sheets></workbook>'),
      'xl/_rels/workbook.xml.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'),
      'xl/worksheets/sheet1.xml': encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C3"/><sheetData><row r="1"><c r="A1"><v>10</v></c><c r="C1"><v>5</v></c></row><row r="2"><c r="A2"><f t="shared" ref="A2:A3" si="0">A1*2</f><v>20</v></c></row><row r="3"><c r="A3"><f t="shared" si="0"/><v>40</v></c></row></sheetData></worksheet>')
    });
    await fs.writeFile(file, bytes);
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.text);
        return response.result?.value;
      };
      const renderedDeadline = Date.now() + 5000;
      while (Date.now() < renderedDeadline && (await evaluate('document.querySelector("#cell-1-0")?.textContent') !== '20' || await evaluate('document.querySelector("#cell-2-0")?.textContent') !== '40')) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('document.querySelector("#cell-1-0")?.textContent'), '20');
      assert.equal(await evaluate('document.querySelector("#cell-2-0")?.textContent'), '40');
      await evaluate('document.querySelector("#cell-0-2").click();const input=document.querySelector("#formula");input.value="8";input.dispatchEvent(new Event("input",{bubbles:true}));input.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))');
      const editedDeadline = Date.now() + 5000;
      while (Date.now() < editedDeadline && await evaluate('document.querySelector("#notice")?.textContent.includes("Modified")') !== true) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('document.querySelector("#notice")?.textContent.includes("Modified")'), true);
      await vscode.commands.executeCommand('workbench.action.files.save');
      const saved = analyze(await fs.readFile(file));
      assert.equal(saved.writable, true, saved.readOnlyReason);
      assert.equal(saved.sheets[0].cells.get('A2').formula, 'A1*2');
      assert.equal(saved.sheets[0].cells.get('A3').formula, 'A2*2');
      assert.match(saved.sheets[0].xml, /<f t="shared" ref="A2:A3" si="0">A1\*2<\/f>/);
      assert.match(saved.sheets[0].xml, /<f t="shared" si="0"\/>/);

      devtools.close();
      devtools = null;
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview(15000, '20');
      assert.equal(await evaluate('document.querySelector("#cell-1-0")?.textContent'), '20');
      assert.equal(await evaluate('document.querySelector("#cell-2-0")?.textContent'), '40');

      assert.equal(await evaluate('Boolean(document.querySelector("#cell-2-0"))'), true, 'the shared formula cell remains rendered after the unrelated save');
      await evaluate('document.querySelector("#cell-2-0").click()');
      await evaluate('(()=>{const formulaEditorForShared=document.querySelector("#formula");formulaEditorForShared.value="=A2+10";formulaEditorForShared.dispatchEvent(new Event("input",{bubbles:true}));return formulaEditorForShared.value})()');
      await evaluate('document.querySelector("#formula").dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))');
      const detachedDeadline = Date.now() + 5000;
      while (Date.now() < detachedDeadline && (await evaluate('document.querySelector("#cell-2-0")?.textContent') !== '30' || await evaluate('document.querySelector("#notice")?.textContent.includes("Modified")') !== true)) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('document.querySelector("#cell-2-0")?.textContent'), '30');
      await vscode.commands.executeCommand('workbench.action.files.save');
      const detached = analyze(await fs.readFile(file));
      assert.equal(detached.writable, true, detached.readOnlyReason);
      assert.equal(detached.sheets[0].cells.get('A2').formula, 'A1*2');
      assert.equal(detached.sheets[0].cells.get('A3').formula, 'A2+10');
      assert.equal(detached.sheets[0].cells.get('A2').formulaType, null);
      assert.equal(detached.sheets[0].cells.get('A3').formulaType, null);
      assert.doesNotMatch(detached.sheets[0].xml, /t="shared"/);

      devtools.close();
      devtools = null;
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const reopenedEvaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.text);
        return response.result?.value;
      };
      const reopenedDeadline = Date.now() + 5000;
      while (Date.now() < reopenedDeadline && (await reopenedEvaluate('document.querySelector("#cell-1-0")?.textContent') !== '20' || await reopenedEvaluate('document.querySelector("#cell-2-0")?.textContent') !== '30')) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await reopenedEvaluate('document.querySelector("#cell-1-0")?.textContent'), '20');
      assert.equal(await reopenedEvaluate('document.querySelector("#cell-2-0")?.textContent'), '30');

      const structureFile = path.join(directory, 'shared-formula-structure.xlsx');
      await fs.writeFile(structureFile, bytes);
      devtools.close();
      devtools = null;
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(structureFile), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview(15000, '20');
      await evaluate('(()=>{document.querySelector("#cell-0-0").click();document.querySelector("#insertRow").click()})()');
      const structureDeadline = Date.now() + 5000;
      while (Date.now() < structureDeadline && (await evaluate('document.querySelector("#cell-2-0")?.textContent') !== '20' || await evaluate('document.querySelector("#cell-3-0")?.textContent') !== '40')) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('document.querySelector("#cell-2-0")?.textContent'), '20');
      assert.equal(await evaluate('document.querySelector("#cell-3-0")?.textContent'), '40');
      await vscode.commands.executeCommand('workbench.action.files.save');
      const structurallySaved = analyze(await fs.readFile(structureFile));
      assert.equal(structurallySaved.writable, true, structurallySaved.readOnlyReason);
      assert.equal(structurallySaved.sheets[0].cells.get('A3').formula, 'A2*2');
      assert.equal(structurallySaved.sheets[0].cells.get('A4').formula, 'A3*2');
      assert.equal(structurallySaved.sheets[0].sharedFormulaGroups.length, 0);
      assert.doesNotMatch(structurallySaved.sheets[0].xml, /t="shared"/);
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview keeps defined-name formulas valid through row insertion, save and reopen', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-defined-name-host-'));
    const file = path.join(directory, 'defined-name.xlsx');
    const encode = text => new TextEncoder().encode(text);
    const bytes = zipSync({
      '[Content_Types].xml': encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'),
      '_rels/.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
      'xl/workbook.xml': encode('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="r1"/></sheets><definedNames><definedName name="DataRange">Data!$A$1:$A$2</definedName></definedNames></workbook>'),
      'xl/_rels/workbook.xml.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'),
      'xl/worksheets/sheet1.xml': encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B2"/><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>SUM(DataRange)</f><v>3</v></c></row><row r="2"><c r="A2"><v>2</v></c></row></sheetData></worksheet>')
    });
    await fs.writeFile(file, bytes);
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview(15000, '1');
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.text);
        return response.result?.value;
      };
      assert.equal(await evaluate('document.querySelector("#cell-0-1")?.textContent'), '3');
      assert.equal(await evaluate('document.querySelector("#insertRow").disabled'), false);
      await evaluate('document.querySelector("#cell-0-0").click();document.querySelector("#insertRow").click()');
      const deadline = Date.now() + 5000;
      while (Date.now() < deadline && (await evaluate('document.querySelector("#cell-1-0")?.textContent') !== '1' || await evaluate('document.querySelector("#cell-1-1")?.textContent') !== '3')) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('document.querySelector("#cell-1-0")?.textContent'), '1');
      assert.equal(await evaluate('document.querySelector("#cell-1-1")?.textContent'), '3');
      await vscode.commands.executeCommand('workbench.action.files.save');
      const saved = analyze(await fs.readFile(file));
      assert.equal(saved.writable, true, saved.readOnlyReason);
      assert.equal(saved.definedNames[0].source.top, 1);
      assert.equal(saved.definedNames[0].source.bottom, 2);
      assert.equal(saved.sheets[0].cells.get('B2').formula, 'SUM(DataRange)');

      devtools.close();
      devtools = null;
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview(15000, '1');
      const reopened = await devtools.send('Runtime.evaluate', { expression: 'document.querySelector("#cell-1-1")?.textContent', contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
      assert.equal(reopened.result?.value, '3');
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview adds a defined name, recalculates formulas, and saves through undo and redo', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-defined-name-edit-host-'));
    const file = path.join(directory, 'defined-name-edit.xlsx'), encode = text => new TextEncoder().encode(text);
    const bytes = zipSync({
      '[Content_Types].xml': encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'),
      '_rels/.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
      'xl/workbook.xml': encode('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="r1"/></sheets></workbook>'),
      'xl/_rels/workbook.xml.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'),
      'xl/worksheets/sheet1.xml': encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C2"/><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>SUM(AddedRange)</f><v>0</v></c></row><row r="2"><c r="A2"><v>2</v></c></row></sheetData></worksheet>')
    });
    await fs.writeFile(file, bytes);
    const originalQuickPick = vscode.window.showQuickPick, originalInputBox = vscode.window.showInputBox;
    let pickCount = 0, inputCount = 0;
    vscode.window.showQuickPick = async items => pickCount++ === 0 ? items.find(item => item.action === 'add') : items.find(item => item.localSheetId === null);
    vscode.window.showInputBox = async () => inputCount++ === 0 ? 'AddedRange' : "'Data'!$A$1:$A$2";
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview(15000, '1');
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.text);
        return response.result?.value;
      };
      assert.equal(await evaluate('document.querySelector("#cell-0-1")?.textContent'), '#NAME?');
      await evaluate('document.querySelector("#manageDefinedNames").click()');
      const deadline = Date.now() + 5000;
      while (Date.now() < deadline && (await evaluate('document.querySelector("#cell-0-1")?.textContent') !== '3' || pickCount < 2 || inputCount < 2)) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(pickCount, 2, 'name and workbook scope prompts were shown');
      assert.equal(inputCount, 2, 'name and range prompts were shown');
      assert.equal(await evaluate('document.querySelector("#cell-0-1")?.textContent'), '3', 'the formula recalculates when its defined name is added');
      await vscode.commands.executeCommand('workbench.action.files.save');
      let saved = analyze(await fs.readFile(file));
      assert.equal(saved.definedNames[0].name, 'AddedRange');
      assert.equal(saved.definedNames[0].source.top, 0);
      assert.equal(saved.definedNames[0].source.bottom, 1);
      assert.equal(saved.sheets[0].cells.get('B1').formula, 'SUM(AddedRange)');

      await vscode.commands.executeCommand('jinsheet.undo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file));
      assert.deepEqual(saved.definedNames, [], 'undo removes the defined name');

      await vscode.commands.executeCommand('jinsheet.redo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file));
      assert.equal(saved.definedNames[0].name, 'AddedRange', 'redo restores the defined name');

      devtools.close();
      devtools = null;
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview(15000, '3');
      assert.equal(await evaluate('document.querySelector("#cell-0-1")?.textContent'), '3', 'the named formula still calculates after reopening');
    } finally {
      vscode.window.showQuickPick = originalQuickPick;
      vscode.window.showInputBox = originalInputBox;
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview navigates to defined names from the Name Box with sheet scope taking priority', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-defined-name-navigation-host-'));
    const file = path.join(directory, 'defined-name-navigation.xlsx'), encode = text => new TextEncoder().encode(text);
    const bytes = zipSync({
      '[Content_Types].xml': encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'),
      '_rels/.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
      'xl/workbook.xml': encode('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Overview Q1" sheetId="1" r:id="r1"/><sheet name="Data" sheetId="2" r:id="r2"/></sheets><definedNames><definedName name="Target">Data!$A$1</definedName><definedName name="Target" localSheetId="0">Data!$B$2:$B$3</definedName></definedNames></workbook>'),
      'xl/_rels/workbook.xml.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="r2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>'),
      'xl/worksheets/sheet1.xml': encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Overview</t></is></c></row></sheetData></worksheet>'),
      'xl/worksheets/sheet2.xml': encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B2"/><sheetData><row r="1"><c r="A1"><v>11</v></c></row><row r="2"><c r="B2"><v>22</v></c></row></sheetData></worksheet>')
    });
    await fs.writeFile(file, bytes);
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview(15000, 'Overview');
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.text);
        return response.result?.value;
      };
      await evaluate('(()=>{const input=document.querySelector("#address");input.value="Target";input.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))})()');
      const deadline = Date.now() + 5000;
      while (Date.now() < deadline && await evaluate('document.querySelector("[role=tab][aria-selected=true]")?.textContent.trim()==="Data"&&document.querySelector(".cell.selected")?.id==="cell-1-1"') !== true) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('document.querySelector("[role=tab][aria-selected=true]")?.textContent.trim()'), 'Data', 'navigation changes to the named range worksheet');
      assert.equal(await evaluate('document.querySelector(".cell.selected")?.id'), 'cell-1-1', 'the active cell follows the worksheet-scoped name');
      assert.equal(await evaluate('document.querySelectorAll(".cell.in-range").length'), 2, 'the worksheet-scoped named range selection is applied');
      assert.equal(await evaluate('document.querySelector("#cell-1-1")?.textContent'), '22', 'the local name shadows the workbook name');

      await evaluate('(()=>{const input=document.querySelector("#address");input.value="\'Overview Q1\'!$A$1:$B$2";input.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))})()');
      const qualifiedDeadline = Date.now() + 5000;
      while (Date.now() < qualifiedDeadline && await evaluate('document.querySelector("[role=tab][aria-selected=true]")?.textContent.trim()==="Overview Q1"&&document.querySelector(".cell.selected")?.id==="cell-0-0"&&document.querySelectorAll(".cell.in-range").length===4') !== true) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('document.querySelector("[role=tab][aria-selected=true]")?.textContent.trim()'), 'Overview Q1', 'a quoted sheet name with spaces is activated');
      assert.equal(await evaluate('document.querySelector(".cell.selected")?.id'), 'cell-0-0', 'the sheet-qualified cell reference is selected');
      assert.equal(await evaluate('document.querySelectorAll(".cell.in-range").length'), 4, 'the sheet-qualified range is selected');
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview inserts a chart-source row and saves, undoes and redoes chart references and caches', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-chart-structure-host-'));
    const file = path.join(directory, 'chart-structure.xlsx'), encode = text => new TextEncoder().encode(text);
    const base = zipSync({
      '[Content_Types].xml': encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'),
      '_rels/.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
      'xl/workbook.xml': encode('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="r1"/></sheets></workbook>'),
      'xl/_rels/workbook.xml.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'),
      'xl/worksheets/sheet1.xml': encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C3"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Month</t></is></c><c r="B1" t="inlineStr"><is><t>Sales</t></is></c><c r="C1" t="inlineStr"><is><t>Profit</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Jan</t></is></c><c r="B2"><v>4</v></c><c r="C2"><v>2</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>Feb</t></is></c><c r="B3"><v>6</v></c><c r="C3"><v>3</v></c></row></sheetData></worksheet>')
    });
    const initial = patchAddChart(analyze(base), 0, { top: 0, bottom: 2, left: 0, right: 1 });
    assert.equal(initial.model.writable, true, initial.model.readOnlyReason);
    await fs.writeFile(file, initial.bytes);
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.exception?.description || response.exceptionDetails.text);
        return response.result?.value;
      };
      await evaluate('document.querySelector("#cell-2-0").click();document.querySelector("#insertRow").click()');
      const deadline = Date.now() + 5000;
      while (Date.now() < deadline && await evaluate('document.querySelector("#cell-3-0")?.textContent') !== 'Feb') await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('document.querySelector("#cell-2-0")?.textContent'), '');
      assert.equal(await evaluate('document.querySelector("#cell-3-0")?.textContent'), 'Feb');
      await vscode.commands.executeCommand('workbench.action.files.save');
      const inserted = analyze(await fs.readFile(file));
      assert.equal(inserted.writable, true, inserted.readOnlyReason);
      assert.match(new TextDecoder().decode(inserted.files['xl/charts/chart1.xml']), /&apos;Data&apos;!\$B\$2:\$B\$4/);
      assert.deepEqual(inserted.sheets[0].charts[0].series[0].values, [4, null, 6]);
      await vscode.commands.executeCommand('jinsheet.undo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      const undone = analyze(await fs.readFile(file));
      assert.equal(undone.sheets[0].cells.get('A3').value, 'Feb');
      assert.deepEqual(undone.sheets[0].charts[0].series[0].values, [4, 6]);
      await vscode.commands.executeCommand('jinsheet.redo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      const redone = analyze(await fs.readFile(file));
      assert.deepEqual(redone.sheets[0].charts[0].series[0].values, [4, null, 6]);
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview expands a chart-backed table and updates series ranges and caches through undo and redo', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-table-chart-expand-host-'));
    const file = path.join(directory, 'table-chart.xlsx'), encode = text => new TextEncoder().encode(text);
    const base = zipSync({
      '[Content_Types].xml': encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'),
      '_rels/.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
      'xl/workbook.xml': encode('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="r1"/></sheets></workbook>'),
      'xl/_rels/workbook.xml.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'),
      'xl/worksheets/sheet1.xml': encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C3"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Month</t></is></c><c r="B1" t="inlineStr"><is><t>Sales</t></is></c><c r="C1" t="inlineStr"><is><t>Profit</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Jan</t></is></c><c r="B2"><v>4</v></c><c r="C2"><v>2</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>Feb</t></is></c><c r="B3"><v>6</v></c><c r="C3"><v>3</v></c></row></sheetData></worksheet>')
    });
    let model = patchAddChart(analyze(base), 0, { top: 0, bottom: 2, left: 0, right: 1 }).model;
    model = patchAddTable(model, { sheetIndex: 0, range: { top: 0, bottom: 2, left: 0, right: 1 }, name: 'MonthlySales' }).model;
    await fs.writeFile(file, model.bytes);
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.exception?.description || response.exceptionDetails.text);
        return response.result?.value;
      };
      const wait = async expression => {
        const deadline = Date.now() + 6000;
        while (Date.now() < deadline) { if (await evaluate(expression)) return; await new Promise(resolve => setTimeout(resolve, 25)); }
        assert.fail('Chart-backed table resize timed out: ' + expression + ' ' + JSON.stringify(await evaluate('({notice:document.querySelector("#notice")?.textContent,series:document.querySelectorAll("[data-chart-series-name]").length,tableRange:document.querySelector("#tableTotalsColumn")?.options?.length,selected:document.querySelector("#address")?.value})')));
      };
      const edit = async (row, column, value) => evaluate(`(()=>{document.querySelector("#cell-${row}-${column}").click();const input=document.querySelector("#formula");input.value=${JSON.stringify(value)};input.dispatchEvent(new Event("input",{bubbles:true}));input.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))})()`);
      await wait('Boolean(document.querySelector("#cell-3-0"))');
      await edit(3, 0, 'Mar');
      const deadline = Date.now() + 5000;
      while (Date.now() < deadline && await evaluate('document.querySelector("#cell-3-0")?.textContent') !== 'Mar') await new Promise(resolve => setTimeout(resolve, 25));
      await edit(3, 1, '8');
      await vscode.commands.executeCommand('workbench.action.files.save');
      let saved = analyze(await fs.readFile(file));
      assert.equal(saved.sheets[0].tables[0].range.bottom, 3);
      assert.deepEqual(saved.sheets[0].charts[0].series[0].categories, ['Jan', 'Feb', 'Mar']);
      assert.deepEqual(saved.sheets[0].charts[0].series[0].values, [4, 6, 8]);
      assert.match(new TextDecoder().decode(saved.files['xl/charts/chart1.xml']), /(?:Data|&apos;Data&apos;)!\$B\$2:\$B\$4/);
      await vscode.commands.executeCommand('jinsheet.undo');
      await vscode.commands.executeCommand('jinsheet.undo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file));
      assert.equal(saved.sheets[0].tables[0].range.bottom, 2);
      assert.deepEqual(saved.sheets[0].charts[0].series[0].values, [4, 6]);
      assert.match(new TextDecoder().decode(saved.files['xl/charts/chart1.xml']), /(?:Data|&apos;Data&apos;)!\$B\$2:\$B\$3/);
      await vscode.commands.executeCommand('jinsheet.redo');
      await vscode.commands.executeCommand('jinsheet.redo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file));
      assert.equal(saved.sheets[0].tables[0].range.bottom, 3);
      assert.deepEqual(saved.sheets[0].charts[0].series[0].values, [4, 6, 8]);
      await evaluate('(()=>{const address=document.querySelector("#address");address.value="A1:C4";address.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))})()');
      await wait('document.querySelectorAll(".cell[aria-selected=true]").length===12');
      await wait('!document.querySelector("#resizeTable").disabled'); await evaluate('document.querySelector("#resizeTable").click()');
      await wait('document.querySelectorAll("[data-chart-series-name]").length===2');
      await vscode.commands.executeCommand('workbench.action.files.save'); saved = analyze(await fs.readFile(file));
      assert.deepEqual(saved.sheets[0].tables[0].columns, ['Month', 'Sales', 'Profit']);
      assert.equal(saved.sheets[0].charts[0].series.length, 2);
      assert.deepEqual(saved.sheets[0].charts[0].series[1].values, [2, 3, null]);
      await vscode.commands.executeCommand('jinsheet.undo'); await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file)); assert.deepEqual(saved.sheets[0].tables[0].columns, ['Month', 'Sales']); assert.equal(saved.sheets[0].charts[0].series.length, 1);
      await vscode.commands.executeCommand('jinsheet.redo'); await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file)); assert.deepEqual(saved.sheets[0].tables[0].columns, ['Month', 'Sales', 'Profit']); assert.equal(saved.sheets[0].charts[0].series.length, 2);
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview maps icon-set values from low red through middle yellow to high green', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-icon-set-host-'));
    const file = path.join(directory, 'icon-set.xlsx');
    const encode = text => new TextEncoder().encode(text);
    const bytes = zipSync({
      '[Content_Types].xml': encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'),
      '_rels/.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
      'xl/workbook.xml': encode('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="r1"/></sheets></workbook>'),
      'xl/_rels/workbook.xml.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'),
      'xl/worksheets/sheet1.xml': encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A4"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Score</t></is></c></row><row r="2"><c r="A2"><v>1</v></c></row><row r="3"><c r="A3"><v>5</v></c></row><row r="4"><c r="A4"><v>9</v></c></row></sheetData><conditionalFormatting sqref="A2:A4"><cfRule type="iconSet" priority="1"><iconSet iconSet="3TrafficLights1"><cfvo type="min"/><cfvo type="num" val="4"/><cfvo type="max"/></iconSet></cfRule></conditionalFormatting></worksheet>')
    });
    await fs.writeFile(file, bytes);
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.text);
        return response.result?.value;
      };
      const iconColor = row => evaluate(`document.querySelector("#cell-${row}-0 svg circle:nth-child(2)")?.getAttribute("fill")`);
      const deadline = Date.now() + 5000;
      while (Date.now() < deadline && (await iconColor(1) !== '#F8696B' || await iconColor(2) !== '#FFEB84' || await iconColor(3) !== '#63BE7B')) await new Promise(resolve => setTimeout(resolve, 25));
      assert.deepEqual([await iconColor(1), await iconColor(2), await iconColor(3)], ['#F8696B', '#FFEB84', '#63BE7B']);
      assert.deepEqual(await evaluate('[1,2,3].map(row=>document.querySelector("#cell-"+row+"-0")?.textContent)'), ['1', '5', '9']);
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview displays supported XLSX conditional formatting and preserves it after edits', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-conditional-format-host-'));
    const file = path.join(directory, 'conditional-format.xlsx');
    const encode = text => new TextEncoder().encode(text);
    const bytes = zipSync({
      '[Content_Types].xml': encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'),
      '_rels/.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
      'xl/workbook.xml': encode('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="r1"/></sheets></workbook>'),
      'xl/_rels/workbook.xml.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="r2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'),
      'xl/styles.xml': encode('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs><dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs></styleSheet>'),
      'xl/worksheets/sheet1.xml': encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A2"/><sheetData><row r="1"><c r="A1"><v>42</v></c></row><row r="2"><c r="A2"><v>1</v></c></row></sheetData><conditionalFormatting sqref="A1:A2"><cfRule type="expression" dxfId="0" priority="1"><formula>A1&gt;10</formula></cfRule></conditionalFormatting></worksheet>')
    });
    await fs.writeFile(file, bytes);
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.text);
        return response.result?.value;
      };
      const redDeadline = Date.now() + 5000;
      while (Date.now() < redDeadline && await evaluate('(()=>{const cell=document.querySelector("#cell-0-0");return cell?getComputedStyle(cell).backgroundColor:""})()') !== 'rgb(255, 0, 0)') await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('getComputedStyle(document.querySelector("#cell-0-0")).backgroundColor'), 'rgb(255, 0, 0)');
      await evaluate('document.querySelector("#cell-1-0").click();const input=document.querySelector("#formula");input.value="15";input.dispatchEvent(new Event("input",{bubbles:true}));input.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))');
      const editedDeadline = Date.now() + 5000;
      while (Date.now() < editedDeadline && await evaluate('(()=>{const cell=document.querySelector("#cell-1-0");return cell?getComputedStyle(cell).backgroundColor:""})()') !== 'rgb(255, 0, 0)') await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('getComputedStyle(document.querySelector("#cell-1-0")).backgroundColor'), 'rgb(255, 0, 0)');
      await vscode.commands.executeCommand('workbench.action.files.save');
      const saved = analyze(await fs.readFile(file));
      assert.equal(saved.writable, true, saved.readOnlyReason);
      assert.equal(saved.sheets[0].cells.get('A2').value, '15');
      assert.equal(saved.sheets[0].conditionalFormats.length, 1);
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview highlights duplicate nonblank values through save, undo and redo', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-duplicate-values-host-'));
    const file = path.join(directory, 'duplicate-values.xlsx'), encode = text => new TextEncoder().encode(text);
    const bytes = zipSync({
      '[Content_Types].xml': encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'),
      '_rels/.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
      'xl/workbook.xml': encode('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="r1"/></sheets></workbook>'),
      'xl/_rels/workbook.xml.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'),
      'xl/worksheets/sheet1.xml': encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A3"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Apple</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>apple</t></is></c></row><row r="3"><c r="A3"/></row></sheetData></worksheet>')
    });
    await fs.writeFile(file, bytes);
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview(15000, 'Apple');
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.text);
        return response.result?.value;
      };
      const duplicateControlDeadline = Date.now() + 5000;
      while (Date.now() < duplicateControlDeadline && !await evaluate('Boolean(document.querySelector("#duplicateValues"))')) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('Boolean(document.querySelector("#duplicateValues"))'), true, 'duplicate-value control is rendered');
      await evaluate('(()=>{const input=document.querySelector("#address");input.value="A1:A3";input.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}));document.querySelector("#duplicateValues").click()})()');
      const highlightDeadline = Date.now() + 5000;
      while (Date.now() < highlightDeadline && await evaluate('(()=>{const cell=document.querySelector("#cell-0-0");return cell&&getComputedStyle(cell).backgroundColor==="rgb(244, 204, 204)"})()') !== true) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('getComputedStyle(document.querySelector("#cell-0-0")).backgroundColor'), 'rgb(244, 204, 204)');
      assert.equal(await evaluate('getComputedStyle(document.querySelector("#cell-1-0")).backgroundColor'), 'rgb(244, 204, 204)');
      assert.notEqual(await evaluate('getComputedStyle(document.querySelector("#cell-2-0")).backgroundColor'), 'rgb(244, 204, 204)', 'blank cells are not highlighted');
      await vscode.commands.executeCommand('workbench.action.files.save');
      let saved = analyze(await fs.readFile(file));
      assert.equal(saved.writable, true, saved.readOnlyReason);
      assert.equal(saved.sheets[0].conditionalFormats[0].formula, 'AND(A1<>"",COUNTIF($A$1:$A$3,A1)>1)');

      await vscode.commands.executeCommand('jinsheet.undo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file));
      assert.equal(saved.sheets[0].conditionalFormats.length, 0, 'undo removes the new conditional-formatting rule');

      await vscode.commands.executeCommand('jinsheet.redo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file));
      assert.equal(saved.sheets[0].conditionalFormats.length, 1, 'redo restores the rule');
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview creates numeric cell rules with localized selection and undo/redo', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-cell-rule-host-'));
    const file = path.join(directory, 'cell-rule.xlsx'), encode = text => new TextEncoder().encode(text);
    const bytes = zipSync({
      '[Content_Types].xml': encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'),
      '_rels/.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
      'xl/workbook.xml': encode('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="r1"/></sheets></workbook>'),
      'xl/_rels/workbook.xml.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'),
      'xl/worksheets/sheet1.xml': encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A3"/><sheetData><row r="1"><c r="A1"><v>5</v></c></row><row r="2"><c r="A2"><v>15</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>25</t></is></c></row></sheetData></worksheet>')
    });
    await fs.writeFile(file, bytes);
    let devtools; const originalQuickPick = vscode.window.showQuickPick, originalInputBox = vscode.window.showInputBox;
    try {
      vscode.window.showQuickPick = async items => items.find(item => item.value === 'greaterThan');
      vscode.window.showInputBox = async () => '10';
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview(15000, '15');
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.text);
        return response.result?.value;
      };
      assert.equal(await evaluate("document.querySelector('#cellValueRule').textContent"), 'Cell rules');
      await evaluate('(()=>{const input=document.querySelector("#address");input.value="A1:A3";input.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}));document.querySelector("#cellValueRule").click()})()');
      const highlightDeadline = Date.now() + 5000;
      while (Date.now() < highlightDeadline && !await evaluate('(()=>{const cell=document.querySelector("#cell-1-0");return cell&&getComputedStyle(cell).backgroundColor==="rgb(244, 204, 204)"})()')) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('(()=>{const cell=document.querySelector("#cell-1-0");return cell&&getComputedStyle(cell).backgroundColor})()'), 'rgb(244, 204, 204)');
      assert.notEqual(await evaluate('getComputedStyle(document.querySelector("#cell-0-0")).backgroundColor'), 'rgb(244, 204, 204)');
      assert.notEqual(await evaluate('getComputedStyle(document.querySelector("#cell-2-0")).backgroundColor'), 'rgb(244, 204, 204)');
      await vscode.commands.executeCommand('workbench.action.files.save');
      let saved = analyze(await fs.readFile(file));
      assert.equal(saved.sheets[0].conditionalFormats[0].operator, 'greaterThan');
      assert.deepEqual(saved.sheets[0].conditionalFormats[0].values, [10]);
      await vscode.commands.executeCommand('jinsheet.undo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file));
      assert.equal(saved.sheets[0].conditionalFormats.length, 0);
      await vscode.commands.executeCommand('jinsheet.redo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file));
      assert.equal(saved.sheets[0].conditionalFormats.length, 1);
    } finally {
      vscode.window.showQuickPick = originalQuickPick; vscode.window.showInputBox = originalInputBox;
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview positions negative data bars around the zero axis', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-negative-data-bar-host-'));
    const file = path.join(directory, 'negative-data-bar.xlsx'), encode = text => new TextEncoder().encode(text);
    const bytes = zipSync({
      '[Content_Types].xml': encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'),
      '_rels/.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
      'xl/workbook.xml': encode('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="r1"/></sheets></workbook>'),
      'xl/_rels/workbook.xml.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'),
      'xl/worksheets/sheet1.xml': encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C1"/><sheetData><row r="1"><c r="A1"><v>-10</v></c><c r="B1"><v>0</v></c><c r="C1"><v>10</v></c></row></sheetData><conditionalFormatting sqref="A1:C1"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule></conditionalFormatting></worksheet>')
    });
    await fs.writeFile(file, bytes);
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.text);
        return response.result?.value;
      };
      const deadline = Date.now() + 5000;
      let bars;
      do {
        bars = await evaluate('[0,1,2].map(column=>document.querySelector("#cell-0-"+column)?.style.backgroundImage)');
        if (bars.every(Boolean)) break;
        await new Promise(resolve => setTimeout(resolve, 25));
      } while (Date.now() < deadline);
      assert.match(bars[0], /transparent 0%.*rgb\(99, 142, 198\) 0%.*50%/);
      assert.equal(bars[1], 'none');
      assert.match(bars[2], /transparent 50%.*rgb\(99, 142, 198\) 50%.*100%/);
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview moves a conditional-formatting rule priority and saves undo/redo', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-conditional-format-priority-host-'));
    const file = path.join(directory, 'conditional-format-priority.xlsx'), encode = text => new TextEncoder().encode(text);
    const dataBar = '<cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF4472C4"/></dataBar></cfRule>';
    const colorScale = '<cfRule type="colorScale" priority="2"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FF00FF00"/></colorScale></cfRule>';
    const bytes = zipSync({
      '[Content_Types].xml': encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'),
      '_rels/.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
      'xl/workbook.xml': encode('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="r1"/></sheets></workbook>'),
      'xl/_rels/workbook.xml.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'),
      'xl/worksheets/sheet1.xml': encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B1"/><sheetData><row r="1"><c r="A1"><v>42</v></c><c r="B1"><v>5</v></c></row></sheetData><conditionalFormatting sqref="A1">' + dataBar + '</conditionalFormatting><conditionalFormatting sqref="B1">' + colorScale + '</conditionalFormatting></worksheet>')
    });
    await fs.writeFile(file, bytes);
    const originalQuickPick = vscode.window.showQuickPick;
    let pickCount = 0, actionItems = [];
    vscode.window.showQuickPick = async items => {
      if (pickCount++ === 0) return items.find(item => item.priority === 1);
      actionItems = items;
      return items.find(item => item.value === 'down');
    };
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.exception?.description || response.exceptionDetails.text);
        return response.result?.value;
      };
      const buttonDeadline = Date.now() + 5000;
      while (Date.now() < buttonDeadline && await evaluate('document.querySelector("#manageConditionalFormats")?.disabled') !== false) await new Promise(resolve => setTimeout(resolve, 25));
      await evaluate('document.querySelector("#manageConditionalFormats").click()');
      const pickerDeadline = Date.now() + 5000;
      while (Date.now() < pickerDeadline && pickCount < 2) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(pickCount, 2);
      assert.deepEqual(actionItems.map(item => item.value), ['colors', 'data-bar-display', 'down', 'remove'], 'the first rule cannot be moved above the top priority');
      await vscode.commands.executeCommand('workbench.action.files.save');
      const moved = analyze(await fs.readFile(file));
      assert.equal(moved.writable, true, moved.readOnlyReason);
      assert.equal(moved.sheets[0].conditionalFormats.find(rule => rule.type === 'colorScale').priority, 1);
      assert.equal(moved.sheets[0].conditionalFormats.find(rule => rule.type === 'dataBar').priority, 2);
      await vscode.commands.executeCommand('jinsheet.undo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      const undone = analyze(await fs.readFile(file));
      assert.equal(undone.sheets[0].conditionalFormats.find(rule => rule.type === 'dataBar').priority, 1);
      assert.equal(undone.sheets[0].conditionalFormats.find(rule => rule.type === 'colorScale').priority, 2);
      await vscode.commands.executeCommand('jinsheet.redo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      const redone = analyze(await fs.readFile(file));
      assert.equal(redone.sheets[0].conditionalFormats.find(rule => rule.type === 'colorScale').priority, 1);
      assert.equal(redone.sheets[0].conditionalFormats.find(rule => rule.type === 'dataBar').priority, 2);
    } finally {
      vscode.window.showQuickPick = originalQuickPick;
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview edits existing conditional-format conditions and colors from the rules manager', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-conditional-format-edit-host-'));
    const file = path.join(directory, 'conditional-format-edit.xlsx'), encode = text => new TextEncoder().encode(text);
    const comparison = '<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>10</formula></cfRule>';
    const colorScale = '<cfRule type="colorScale" priority="2"><colorScale><cfvo type="min"/><cfvo type="num" val="50"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FFFFFF00"/><color rgb="FF00FF00"/></colorScale></cfRule>';
    const dataBar = '<cfRule type="dataBar" priority="3"><dataBar minLength="5" maxLength="90"><cfvo type="min"/><cfvo type="max"/><color rgb="FF4472C4"/></dataBar></cfRule>';
    const bytes = zipSync({
      '[Content_Types].xml': encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'),
      '_rels/.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
      'xl/workbook.xml': encode('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="r1"/></sheets></workbook>'),
      'xl/_rels/workbook.xml.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="r2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'),
      'xl/styles.xml': encode('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs><dxfs count="1"><dxf><font><b/><color rgb="FF008000"/></font><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs></styleSheet>'),
      'xl/worksheets/sheet1.xml': encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B1"/><sheetData><row r="1"><c r="A1"><v>42</v></c><c r="B1"><v>5</v></c></row></sheetData><conditionalFormatting sqref="A1">' + comparison + '</conditionalFormatting><conditionalFormatting sqref="B1">' + colorScale + '</conditionalFormatting><conditionalFormatting sqref="C1">' + dataBar + '</conditionalFormatting></worksheet>')
    });
    await fs.writeFile(file, bytes);
    const originalQuickPick = vscode.window.showQuickPick, originalInputBox = vscode.window.showInputBox;
    const choose = [
      items => items.find(item => item.priority === 1),
      items => items.find(item => item.value === 'condition'),
      items => items.find(item => item.value === 'between'),
      items => items.find(item => item.priority === 2),
      items => items.find(item => item.value === 'colors'),
      items => items.find(item => item.priority === 2),
      items => items.find(item => item.value === 'threshold'),
      items => items.find(item => item.value === 'percent'),
      items => items.find(item => item.priority === 1),
      items => items.find(item => item.value === 'font-color'),
      items => items.find(item => item.priority === 1),
      items => items.find(item => item.value === 'fill-color'),
      items => items.find(item => item.priority === 3),
      items => items.find(item => item.value === 'data-bar-display'),
      items => items.find(item => item.value === 'minLength'),
      items => items.find(item => item.priority === 3),
      items => items.find(item => item.value === 'data-bar-display'),
      items => items.find(item => item.value === 'maxLength'),
      items => items.find(item => item.priority === 3),
      items => items.find(item => item.value === 'data-bar-display'),
      items => items.find(item => item.value === 'showValue'),
      items => items.find(item => item.value === false),
      items => items.find(item => item.priority === 3),
      items => items.find(item => item.value === 'data-bar-display'),
      items => items.find(item => item.value === 'gradient'),
      items => items.find(item => item.value === false)
    ];
    let pickCount = 0, inputCount = 0;
    vscode.window.showQuickPick = async items => choose[pickCount++]?.(items);
    vscode.window.showInputBox = async () => ['40', '45', '#123456', '#654321', '#ABCDEF', '60', '#123ABC', '#654321', '75', '95'][inputCount++];
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.exception?.description || response.exceptionDetails.text);
        return response.result?.value;
      };
      const buttonDeadline = Date.now() + 5000;
      while (Date.now() < buttonDeadline && await evaluate('document.querySelector("#manageConditionalFormats")?.disabled') !== false) await new Promise(resolve => setTimeout(resolve, 25));
      await evaluate('document.querySelector("#manageConditionalFormats").click()');
      const conditionDeadline = Date.now() + 5000;
      while (Date.now() < conditionDeadline && (pickCount < 3 || inputCount < 2)) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(pickCount, 3); assert.equal(inputCount, 2);
      await vscode.commands.executeCommand('workbench.action.files.save');
      const conditionSaved = analyze(await fs.readFile(file));
      assert.deepEqual(conditionSaved.sheets[0].conditionalFormats[0].values, [40, 45]);
      assert.equal(conditionSaved.sheets[0].conditionalFormats[0].operator, 'between');
      await evaluate('document.querySelector("#manageConditionalFormats").click()');
      const colorDeadline = Date.now() + 5000;
      while (Date.now() < colorDeadline && (pickCount < 5 || inputCount < 5)) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(pickCount, 5); assert.equal(inputCount, 5);
      const colorSaveDeadline = Date.now() + 5000;
      let colorSaved;
      do { await vscode.commands.executeCommand('workbench.action.files.save'); colorSaved = analyze(await fs.readFile(file)); if (colorSaved.sheets[0].conditionalFormats.find(rule => rule.type === 'colorScale')?.colorScale.colors?.[0] === '#123456') break; await new Promise(resolve => setTimeout(resolve, 25)); } while (Date.now() < colorSaveDeadline);
      assert.deepEqual(colorSaved.sheets[0].conditionalFormats.find(rule => rule.type === 'colorScale').colorScale.colors, ['#123456', '#654321', '#ABCDEF']);
      await evaluate('document.querySelector("#manageConditionalFormats").click()');
      const thresholdDeadline = Date.now() + 5000;
      while (Date.now() < thresholdDeadline && (pickCount < 8 || inputCount < 6)) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(pickCount, 8); assert.equal(inputCount, 6);
      const thresholdSaveDeadline = Date.now() + 5000;
      let thresholdSaved;
      do { await vscode.commands.executeCommand('workbench.action.files.save'); thresholdSaved = analyze(await fs.readFile(file)); if (thresholdSaved.sheets[0].conditionalFormats.find(rule => rule.type === 'colorScale')?.colorScale.stops?.[1]?.value === 60) break; await new Promise(resolve => setTimeout(resolve, 25)); } while (Date.now() < thresholdSaveDeadline);
      assert.deepEqual(thresholdSaved.sheets[0].conditionalFormats.find(rule => rule.type === 'colorScale').colorScale.stops[1], { type: 'percent', value: 60 });
      await evaluate('document.querySelector("#manageConditionalFormats").click()');
      const fontColorDeadline = Date.now() + 5000;
      while (Date.now() < fontColorDeadline && (pickCount < 10 || inputCount < 7)) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(pickCount, 10); assert.equal(inputCount, 7);
      const fontColorSaveDeadline = Date.now() + 5000;
      let fontColorSaved;
      do { await vscode.commands.executeCommand('workbench.action.files.save'); fontColorSaved = analyze(await fs.readFile(file)); if (fontColorSaved.sheets[0].conditionalFormats.find(rule => rule.type === 'cellIs')?.style?.color === '#123ABC') break; await new Promise(resolve => setTimeout(resolve, 25)); } while (Date.now() < fontColorSaveDeadline);
      assert.deepEqual(fontColorSaved.sheets[0].conditionalFormats.find(rule => rule.type === 'cellIs').style, { bold: true, color: '#123ABC', fill: '#FF0000' });
      await evaluate('document.querySelector("#manageConditionalFormats").click()');
      const fillColorDeadline = Date.now() + 5000;
      while (Date.now() < fillColorDeadline && (pickCount < 12 || inputCount < 8)) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(pickCount, 12); assert.equal(inputCount, 8);
      await vscode.commands.executeCommand('workbench.action.files.save');
      const fillColorSaved = analyze(await fs.readFile(file));
      assert.deepEqual(fillColorSaved.sheets[0].conditionalFormats.find(rule => rule.type === 'cellIs').style, { bold: true, color: '#123ABC', fill: '#654321' });
      await evaluate('document.querySelector("#manageConditionalFormats").click()');
      const minLengthDeadline = Date.now() + 5000;
      while (Date.now() < minLengthDeadline && (pickCount < 15 || inputCount < 9)) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(pickCount, 15); assert.equal(inputCount, 9);
      await evaluate('document.querySelector("#manageConditionalFormats").click()');
      const maxLengthDeadline = Date.now() + 5000;
      while (Date.now() < maxLengthDeadline && (pickCount < 18 || inputCount < 10)) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(pickCount, 18); assert.equal(inputCount, 10);
      await evaluate('document.querySelector("#manageConditionalFormats").click()');
      const valueDisplayDeadline = Date.now() + 5000;
      while (Date.now() < valueDisplayDeadline && pickCount < 22) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(pickCount, 22);
      await evaluate('document.querySelector("#manageConditionalFormats").click()');
      const fillDisplayDeadline = Date.now() + 5000;
      while (Date.now() < fillDisplayDeadline && pickCount < 26) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(pickCount, 26);
      const displaySaveDeadline = Date.now() + 5000;
      let displaySaved;
      do {
        await vscode.commands.executeCommand('workbench.action.files.save');
        displaySaved = analyze(await fs.readFile(file));
        const dataBar = displaySaved.sheets[0].conditionalFormats.find(rule => rule.type === 'dataBar')?.dataBar;
        if (dataBar?.minLength === 75 && dataBar.maxLength === 95 && dataBar.showValue === false && dataBar.gradient === false) break;
        await new Promise(resolve => setTimeout(resolve, 25));
      } while (Date.now() < displaySaveDeadline);
      assert.deepEqual(displaySaved.sheets[0].conditionalFormats.find(rule => rule.type === 'dataBar').dataBar, { color: '#4472C4', minLength: 75, maxLength: 95, showValue: false, gradient: false });
    } finally {
      vscode.window.showQuickPick = originalQuickPick; vscode.window.showInputBox = originalInputBox;
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview edits an XLSX top/bottom conditional-format rank with undo and redo', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-top-bottom-rule-edit-host-'));
    const file = path.join(directory, 'top-bottom-rule.xlsx'), encode = text => new TextEncoder().encode(text);
    const bytes = zipSync({
      '[Content_Types].xml': encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'),
      '_rels/.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
      'xl/workbook.xml': encode('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="r1"/></sheets></workbook>'),
      'xl/_rels/workbook.xml.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="r2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'),
      'xl/styles.xml': encode('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs><dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs></styleSheet>'),
      'xl/worksheets/sheet1.xml': encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A3"/><sheetData><row r="1"><c r="A1"><v>10</v></c></row><row r="2"><c r="A2"><v>20</v></c></row><row r="3"><c r="A3"><v>30</v></c></row></sheetData><conditionalFormatting sqref="A1:A3"><cfRule type="top10" dxfId="0" priority="1" rank="500"/></conditionalFormatting></worksheet>')
    });
    await fs.writeFile(file, bytes);
    const originalQuickPick = vscode.window.showQuickPick, originalInputBox = vscode.window.showInputBox;
    let pickCount = 0, inputCount = 0, inputDefault, devtools;
    vscode.window.showQuickPick = async items => {
      const step = pickCount++;
      if (step === 0) return items.find(item => item.priority === 1);
      if (step === 1) return items.find(item => item.value === 'condition');
      if (step === 2) return items.find(item => item.value === true);
      if (step === 3) return items.find(item => item.value === true);
      return undefined;
    };
    vscode.window.showInputBox = async options => { inputCount++; inputDefault = options.value; return '50'; };
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.text);
        return response.result?.value;
      };
      await evaluate('document.querySelector("#manageConditionalFormats").click()');
      const deadline = Date.now() + 5000;
      while (Date.now() < deadline && (pickCount < 4 || inputCount < 1)) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(pickCount, 4); assert.equal(inputCount, 1);
      assert.equal(inputDefault, '100', 'switching from a large count to a percentage clamps the default into range');
      await vscode.commands.executeCommand('workbench.action.files.save');
      let saved = analyze(await fs.readFile(file));
      assert.deepEqual(saved.sheets[0].conditionalFormats[0].top10, { rank: 50, percent: true, bottom: true });
      await vscode.commands.executeCommand('jinsheet.undo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file));
      assert.deepEqual(saved.sheets[0].conditionalFormats[0].top10, { rank: 500, percent: false, bottom: false });
      await vscode.commands.executeCommand('jinsheet.redo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file));
      assert.deepEqual(saved.sheets[0].conditionalFormats[0].top10, { rank: 50, percent: true, bottom: true });
    } finally {
      vscode.window.showQuickPick = originalQuickPick; vscode.window.showInputBox = originalInputBox;
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview creates a standard icon set through the formatting toolbar and saves undo/redo', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-icon-set-create-host-'));
    const file = path.join(directory, 'icon-set-create.xlsx'), encode = text => new TextEncoder().encode(text);
    const bytes = zipSync({
      '[Content_Types].xml': encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'),
      '_rels/.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
      'xl/workbook.xml': encode('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="r1"/></sheets></workbook>'),
      'xl/_rels/workbook.xml.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'),
      'xl/worksheets/sheet1.xml': encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A3"/><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>5</v></c></row><row r="3"><c r="A3"><v>9</v></c></row></sheetData></worksheet>')
    });
    await fs.writeFile(file, bytes);
    const originalQuickPick = vscode.window.showQuickPick;
    let pickCount = 0, choiceCount = 0;
    vscode.window.showQuickPick = async items => { pickCount++; choiceCount = items.length; return items.find(item => item.value === '5Quarters'); };
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.exception?.description || response.exceptionDetails.text);
        return response.result?.value;
      };
      const buttonDeadline = Date.now() + 5000;
      while (Date.now() < buttonDeadline && await evaluate('document.querySelector("#iconSet")?.disabled') !== false) await new Promise(resolve => setTimeout(resolve, 25));
      await evaluate('document.querySelector("#iconSet").click()');
      const pickerDeadline = Date.now() + 5000;
      while (Date.now() < pickerDeadline && pickCount < 1) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(pickCount, 1); assert.equal(choiceCount, 17);
      await new Promise(resolve => setTimeout(resolve, 50));
      await vscode.commands.executeCommand('workbench.action.files.save');
      const saved = analyze(await fs.readFile(file));
      assert.equal(saved.writable, true, saved.readOnlyReason);
      assert.equal(saved.sheets[0].conditionalFormats[0].type, 'iconSet');
      assert.equal(saved.sheets[0].conditionalFormats[0].iconSet.name, '5Quarters');
      assert.equal(saved.sheets[0].conditionalFormats[0].iconSet.thresholds.length, 5);
      await vscode.commands.executeCommand('jinsheet.undo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(analyze(await fs.readFile(file)).sheets[0].conditionalFormats.length, 0);
      await vscode.commands.executeCommand('jinsheet.redo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(analyze(await fs.readFile(file)).sheets[0].conditionalFormats[0].iconSet.name, '5Quarters');
    } finally {
      vscode.window.showQuickPick = originalQuickPick;
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview edits icon-set thresholds and display options from the rules manager', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-icon-set-edit-host-'));
    const file = path.join(directory, 'icon-set-edit.xlsx'), encode = text => new TextEncoder().encode(text);
    const bytes = zipSync({
      '[Content_Types].xml': encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'),
      '_rels/.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
      'xl/workbook.xml': encode('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="r1"/></sheets></workbook>'),
      'xl/_rels/workbook.xml.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'),
      'xl/worksheets/sheet1.xml': encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A3"/><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>5</v></c></row><row r="3"><c r="A3"><v>9</v></c></row></sheetData><conditionalFormatting sqref="A1:A3"><cfRule type="iconSet" priority="1"><iconSet iconSet="3TrafficLights1" percent="1" showValue="1"><cfvo type="min"/><cfvo type="num" val="4" gte="0"/><cfvo type="max"/></iconSet></cfRule></conditionalFormatting></worksheet>')
    });
    await fs.writeFile(file, bytes);
    const originalQuickPick = vscode.window.showQuickPick, originalInputBox = vscode.window.showInputBox;
    const choose = [
      items => items.find(item => item.priority === 1), items => items.find(item => item.value === 'icon-set'),
      items => items.find(item => item.value === 'threshold:1'), items => items.find(item => item.value === 'percent'), items => items.find(item => item.value === true),
      items => items.find(item => item.priority === 1), items => items.find(item => item.value === 'icon-set'),
      items => items.find(item => item.value === 'showValue'), items => items.find(item => item.value === false),
      items => items.find(item => item.priority === 1), items => items.find(item => item.value === 'icon-set'),
      items => items.find(item => item.value === 'reverse'), items => items.find(item => item.value === true),
      items => items.find(item => item.priority === 1), items => items.find(item => item.value === 'icon-set'),
      items => items.find(item => item.value === 'style'),
      items => { const choice = items.find(item => item.value === '5Quarters'); assert.match(choice.description, /threshold|기준/i); return choice; },
      items => items.find(item => item.priority === 1), items => items.find(item => item.value === 'icon-set'), items => items.find(item => item.value === 'mapping'),
      ...Array.from({ length: 5 }, () => [items => items.find(item => item.value === '3ArrowsGray'), items => items.find(item => item.value === 1)]).flat(),
      items => items.find(item => item.priority === 1), items => items.find(item => item.value === 'icon-set'), items => items.find(item => item.value === 'threshold:3'), items => items.find(item => item.value === 'num'), items => items.find(item => item.value === true)
    ];
    let pickCount = 0, inputCount = 0;
    vscode.window.showQuickPick = async items => choose[pickCount++]?.(items);
    vscode.window.showInputBox = async () => { inputCount++; return '60'; };
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.exception?.description || response.exceptionDetails.text);
        return response.result?.value;
      };
      const buttonDeadline = Date.now() + 5000;
      while (Date.now() < buttonDeadline && await evaluate('document.querySelector("#manageConditionalFormats")?.disabled') !== false) await new Promise(resolve => setTimeout(resolve, 25));
      await evaluate('document.querySelector("#manageConditionalFormats").click()');
      const thresholdDeadline = Date.now() + 5000;
      while (Date.now() < thresholdDeadline && (pickCount < 5 || inputCount < 1)) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(pickCount, 5); assert.equal(inputCount, 1);
      await evaluate('document.querySelector("#manageConditionalFormats").click()');
      const valueDeadline = Date.now() + 5000;
      while (Date.now() < valueDeadline && pickCount < 9) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(pickCount, 9);
      await evaluate('document.querySelector("#manageConditionalFormats").click()');
      const reverseDeadline = Date.now() + 5000;
      while (Date.now() < reverseDeadline && pickCount < 13) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(pickCount, 13);
      await evaluate('document.querySelector("#manageConditionalFormats").click()');
      const styleDeadline = Date.now() + 5000;
      while (Date.now() < styleDeadline && pickCount < 17) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(pickCount, 17);
      await evaluate('document.querySelector("#manageConditionalFormats").click()');
      const mappingDeadline = Date.now() + 10000;
      while (Date.now() < mappingDeadline && pickCount < 30) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(pickCount, 30);
      await evaluate('document.querySelector("#manageConditionalFormats").click()');
      const customThresholdDeadline = Date.now() + 5000;
      while (Date.now() < customThresholdDeadline && (pickCount < 35 || inputCount < 2)) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(pickCount, 35); assert.equal(inputCount, 2);
      await new Promise(resolve => setTimeout(resolve, 75));
      await vscode.commands.executeCommand('workbench.action.files.save');
      const saved = analyze(await fs.readFile(file));
      assert.equal(saved.writable, true, saved.readOnlyReason);
      assert.deepEqual(saved.sheets[0].conditionalFormats[0].iconSet, { name: '5Quarters', showValue: false, reverse: true, thresholds: [
        { type: 'percent', value: 0, gte: true }, { type: 'percent', value: 20, gte: true }, { type: 'percent', value: 40, gte: true },
        { type: 'num', value: 60, gte: true }, { type: 'percent', value: 80, gte: true }
      ], customIcons: Array.from({ length: 5 }, () => ({ iconSet: '3ArrowsGray', iconId: 1 })) });
      await vscode.commands.executeCommand('jinsheet.undo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.deepEqual(analyze(await fs.readFile(file)).sheets[0].conditionalFormats[0].iconSet.thresholds[3], { type: 'percent', value: 60, gte: true });
      await vscode.commands.executeCommand('jinsheet.undo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(analyze(await fs.readFile(file)).sheets[0].conditionalFormats[0].iconSet.customIcons, undefined);
      await vscode.commands.executeCommand('jinsheet.undo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      const restoredStyle = analyze(await fs.readFile(file)).sheets[0].conditionalFormats[0].iconSet;
      assert.deepEqual(restoredStyle, { name: '3TrafficLights1', showValue: false, reverse: true, thresholds: [{ type: 'min', value: null, gte: true }, { type: 'percent', value: 60, gte: true }, { type: 'max', value: null, gte: true }] });
      await vscode.commands.executeCommand('jinsheet.undo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(analyze(await fs.readFile(file)).sheets[0].conditionalFormats[0].iconSet.reverse, false);
      await vscode.commands.executeCommand('jinsheet.redo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(analyze(await fs.readFile(file)).sheets[0].conditionalFormats[0].iconSet.reverse, true);
      await vscode.commands.executeCommand('jinsheet.redo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      await vscode.commands.executeCommand('jinsheet.redo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      await vscode.commands.executeCommand('jinsheet.redo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.deepEqual(analyze(await fs.readFile(file)).sheets[0].conditionalFormats[0].iconSet, saved.sheets[0].conditionalFormats[0].iconSet);
    } finally {
      vscode.window.showQuickPick = originalQuickPick; vscode.window.showInputBox = originalInputBox;
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview applies list and numeric validation in cells and preserves rules after saving', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-data-validation-host-'));
    const file = path.join(directory, 'data-validation.xlsx');
    const encode = text => new TextEncoder().encode(text);
    const listValidation = '<dataValidation type="list" allowBlank="1" sqref="A1"><formula1>=$B$1:$B$3</formula1></dataValidation>';
    const numberValidation = '<dataValidation type="whole" operator="between" errorTitle="Invalid value" error="Enter a whole number from 1 to 10" promptTitle="Allowed range" prompt="Enter 1 through 10" sqref="C1"><formula1>D1</formula1><formula2>E1</formula2></dataValidation>';
    const validation = '<dataValidations count="2">' + listValidation + numberValidation + '</dataValidations>';
    const bytes = zipSync({
      '[Content_Types].xml': encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'),
      '_rels/.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
      'xl/workbook.xml': encode('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="r1"/></sheets></workbook>'),
      'xl/_rels/workbook.xml.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'),
      'xl/worksheets/sheet1.xml': encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:E3"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Open</t></is></c><c r="B1" t="inlineStr"><is><t>Open</t></is></c><c r="C1"><v>5</v></c><c r="D1"><v>1</v></c><c r="E1"><v>10</v></c></row><row r="2"><c r="B2" t="inlineStr"><is><t>In Progress</t></is></c></row><row r="3"><c r="B3" t="inlineStr"><is><t>Closed</t></is></c></row></sheetData>' + validation + '</worksheet>')
    });
    await fs.writeFile(file, bytes);
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.exception?.description || response.exceptionDetails.text);
        return response.result?.value;
      };
      const optionsDeadline = Date.now() + 5000;
      while (Date.now() < optionsDeadline && await evaluate('document.querySelectorAll("#cellValidationOptions option").length') !== 3) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('Array.from(document.querySelectorAll("#cellValidationOptions option"),option=>option.value).join(",")'), 'Open,In Progress,Closed');
      await evaluate('(()=>{const input=document.querySelector("#formula");input.value="Pending";input.dispatchEvent(new Event("input",{bubbles:true}));input.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))})()');
      const rejectedDeadline = Date.now() + 3000;
      while (Date.now() < rejectedDeadline && !/allowed list|허용 목록/.test(String(await evaluate('document.querySelector("#notice").textContent')))) await new Promise(resolve => setTimeout(resolve, 25));
      assert.match(await evaluate('document.querySelector("#notice").textContent'), /allowed list|허용 목록/);
      assert.equal(analyze(await fs.readFile(file)).sheets[0].cells.get('A1').value, 'Open');
      await evaluate('document.querySelector("#cell-0-0").click()');
      await evaluate('(()=>{const input=document.querySelector("#formula");input.focus();input.value="Closed";input.dispatchEvent(new Event("input",{bubbles:true}));input.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))})()');
      const modifiedDeadline = Date.now() + 3000;
      while (Date.now() < modifiedDeadline && await evaluate('document.querySelector("#notice").textContent.includes("Modified")') !== true) await new Promise(resolve => setTimeout(resolve, 25));
      assert.match(await evaluate('document.querySelector("#notice").textContent'), /Modified/, await evaluate('JSON.stringify({notice:document.querySelector("#notice")?.textContent,selected:document.querySelector(".cell.selected")?.id,formula:document.querySelector("#formula")?.value,cell:document.querySelector("#cell-0-0")?.__jinSheetData})'));
      await vscode.commands.executeCommand('workbench.action.files.save');
      const saved = analyze(await fs.readFile(file));
      assert.equal(saved.sheets[0].cells.get('A1').value, 'Closed');
      assert.ok(saved.sheets[0].xml.includes(validation));
      await evaluate('(()=>{const grid=document.querySelector("#grid");grid.scrollLeft=200;grid.dispatchEvent(new Event("scroll"));return true})()');
      const numericCellDeadline = Date.now() + 3000;
      while (Date.now() < numericCellDeadline && !await evaluate('Boolean(document.querySelector("#cell-0-2"))')) await new Promise(resolve => setTimeout(resolve, 25));
      assert.ok(await evaluate('Boolean(document.querySelector("#cell-0-2"))'), await evaluate('JSON.stringify({cells:[...document.querySelectorAll(".cell")].map(cell=>cell.id).slice(0,12),scrollLeft:document.querySelector("#grid")?.scrollLeft,clientWidth:document.querySelector("#grid")?.clientWidth})'));
      assert.match(await evaluate('document.querySelector("#cell-0-2").title'), /Allowed range\nEnter 1 through 10/);
      await evaluate('document.querySelector("#cell-0-2").click()');
      await evaluate('(()=>{const input=document.querySelector("#formula");input.value="12";input.dispatchEvent(new Event("input",{bubbles:true}));input.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))})()');
      const numericRejectedDeadline = Date.now() + 3000;
      while (Date.now() < numericRejectedDeadline && !/Enter a whole number from 1 to 10/.test(String(await evaluate('document.querySelector("#notice").textContent')))) await new Promise(resolve => setTimeout(resolve, 25));
      assert.match(await evaluate('document.querySelector("#notice").textContent'), /Enter a whole number from 1 to 10/);
      assert.equal(analyze(await fs.readFile(file)).sheets[0].cells.get('C1').value, '5');
      await evaluate('document.querySelector("#cell-0-2").click()');
      await evaluate('(()=>{const input=document.querySelector("#formula");input.focus();input.value="9";input.dispatchEvent(new Event("input",{bubbles:true}));input.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))})()');
      const secondModifiedDeadline = Date.now() + 3000;
      while (Date.now() < secondModifiedDeadline && await evaluate('document.querySelector("#notice").textContent.includes("Modified")') !== true) await new Promise(resolve => setTimeout(resolve, 25));
      await vscode.commands.executeCommand('workbench.action.files.save');
      const finalSaved = analyze(await fs.readFile(file));
      assert.equal(finalSaved.sheets[0].cells.get('C1').value, '9');
      assert.ok(finalSaved.sheets[0].xml.includes(validation));
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview creates a list validation for the selected XLSX cells', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-validation-create-host-'));
    const file = path.join(directory, 'validation-create.xlsx'), encode = text => new TextEncoder().encode(text);
    const bytes = zipSync({
      '[Content_Types].xml': encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'),
      '_rels/.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
      'xl/workbook.xml': encode('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="r1"/></sheets></workbook>'),
      'xl/_rels/workbook.xml.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'),
      'xl/worksheets/sheet1.xml': encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>old</t></is></c></row></sheetData></worksheet>')
    });
    await fs.writeFile(file, bytes);
    const originalQuickPick = vscode.window.showQuickPick, originalInputBox = vscode.window.showInputBox, prompts = [];
    vscode.window.showQuickPick = async items => { prompts.push('type'); return items.find(item => item.type === 'list'); };
    vscode.window.showInputBox = async options => { prompts.push(options.prompt); return 'Approved,Pending'; };
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.exception?.description || response.exceptionDetails.text);
        return response.result?.value;
      };
      const buttonDeadline = Date.now() + 5000;
      while (Date.now() < buttonDeadline && await evaluate('document.querySelector("#dataValidation")?.disabled') !== false) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('document.querySelector("#dataValidation")?.textContent'), 'Data validation');
      await evaluate('document.querySelector("#dataValidation").click()');
      const validationDeadline = Date.now() + 5000;
      while (Date.now() < validationDeadline && await evaluate('document.querySelectorAll("#cellValidationOptions option").length') !== 2) await new Promise(resolve => setTimeout(resolve, 25));
      assert.deepEqual(prompts, ['type', 'Comma-separated values, a cell range (for example =$A$1:$A$5), or a defined name']);
      assert.equal(await evaluate('Array.from(document.querySelectorAll("#cellValidationOptions option"),option=>option.value).join(",")'), 'Approved,Pending');
      await vscode.commands.executeCommand('workbench.action.files.save');
      const saved = analyze(await fs.readFile(file));
      assert.equal(saved.writable, true, saved.readOnlyReason);
      assert.deepEqual(saved.sheets[0].dataValidations[0].values, ['Approved', 'Pending']);
      await evaluate('(()=>{const input=document.querySelector("#formula");input.focus();input.value="Other";input.dispatchEvent(new Event("input",{bubbles:true}));input.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))})()');
      const rejectedDeadline = Date.now() + 3000;
      while (Date.now() < rejectedDeadline && !/allowed list|허용 목록/.test(String(await evaluate('document.querySelector("#notice").textContent'))) ) await new Promise(resolve => setTimeout(resolve, 25));
      assert.match(await evaluate('document.querySelector("#notice").textContent'), /allowed list|허용 목록/);
    } finally {
      vscode.window.showQuickPick = originalQuickPick; vscode.window.showInputBox = originalInputBox;
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview updates list validation after row insertion, save and reopen', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-validation-structure-host-'));
    const file = path.join(directory, 'validation-structure.xlsx'), encode = text => new TextEncoder().encode(text);
    const validation = '<dataValidations count="1"><dataValidation type="list" allowBlank="1" sqref="A1"><formula1>=$B$1:$B$3</formula1></dataValidation></dataValidations>';
    await fs.writeFile(file, zipSync({
      '[Content_Types].xml': encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'),
      '_rels/.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
      'xl/workbook.xml': encode('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="r1"/></sheets></workbook>'),
      'xl/_rels/workbook.xml.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'),
      'xl/worksheets/sheet1.xml': encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B3"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Open</t></is></c><c r="B1" t="inlineStr"><is><t>Open</t></is></c></row><row r="2"><c r="B2" t="inlineStr"><is><t>In Progress</t></is></c></row><row r="3"><c r="B3" t="inlineStr"><is><t>Closed</t></is></c></row></sheetData>' + validation + '</worksheet>')
    }));
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview(15000, 'Open');
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.exception?.description || response.exceptionDetails.text);
        return response.result?.value;
      };
      assert.equal(await evaluate('document.querySelector("#insertRow").disabled'), false);
      await evaluate('document.querySelector("#cell-0-0").click();document.querySelector("#insertRow").click()');
      const insertedDeadline = Date.now() + 5000;
      while (Date.now() < insertedDeadline && await evaluate('document.querySelector("#cell-1-0")?.textContent') !== 'Open') await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('document.querySelector("#cell-1-0")?.textContent'), 'Open');
      await evaluate('document.querySelector("#cell-1-0").click()');
      const options = 'Array.from(document.querySelectorAll("#cellValidationOptions option"),option=>option.value).join(",")';
      const optionsDeadline = Date.now() + 5000;
      while (Date.now() < optionsDeadline && await evaluate(`(${options})`) !== 'Open,In Progress,Closed') await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate(`(${options})`), 'Open,In Progress,Closed');
      await vscode.commands.executeCommand('workbench.action.files.save');
      const saved = analyze(await fs.readFile(file));
      assert.equal(saved.writable, true, saved.readOnlyReason);
      assert.deepEqual(saved.sheets[0].dataValidations[0].ranges, [{ top: 1, bottom: 1, left: 0, right: 0 }]);
      assert.match(saved.sheets[0].xml, /sqref="A2"><formula1>=\$B\$2:\$B\$4<\/formula1>/);

      devtools.close();
      devtools = null;
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview(15000, 'Open');
      const reopenedEvaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.exception?.description || response.exceptionDetails.text);
        return response.result?.value;
      };
      await reopenedEvaluate('document.querySelector("#cell-1-0").click()');
      const reopenedOptionsDeadline = Date.now() + 5000;
      while (Date.now() < reopenedOptionsDeadline && await reopenedEvaluate(`(${options})`) !== 'Open,In Progress,Closed') await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await reopenedEvaluate(`(${options})`), 'Open,In Progress,Closed');
      const reopened = analyze(await fs.readFile(file));
      assert.equal(reopened.sheets[0].dataValidations[0].ranges[0].top, 1);
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview asks before warning or information data-validation entries and saves accepted values', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-validation-alerts-host-'));
    const file = path.join(directory, 'validation-alerts.xlsx'), encode = text => new TextEncoder().encode(text);
    const validation = '<dataValidations count="2"><dataValidation type="list" errorStyle="warning" errorTitle="Check" error="Value is outside the list" sqref="A1"><formula1>"Yes,No"</formula1></dataValidation><dataValidation type="list" errorStyle="information" errorTitle="Notice" error="Value is outside the list" sqref="B1"><formula1>"Yes,No"</formula1></dataValidation></dataValidations>';
    await fs.writeFile(file, zipSync({
      '[Content_Types].xml': encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'),
      '_rels/.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
      'xl/workbook.xml': encode('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="r1"/></sheets></workbook>'),
      'xl/_rels/workbook.xml.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'),
      'xl/worksheets/sheet1.xml': encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Yes</t></is></c><c r="B1" t="inlineStr"><is><t>No</t></is></c></row></sheetData>' + validation + '</worksheet>')
    }));
    const originalWarning = vscode.window.showWarningMessage, originalInformation = vscode.window.showInformationMessage, prompts = [];
    vscode.window.showWarningMessage = async (message, ...actions) => { prompts.push({ type: 'warning', message }); return actions[0]; };
    vscode.window.showInformationMessage = async (message, ...actions) => { prompts.push({ type: 'information', message }); return actions[0]; };
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.exception?.description || response.exceptionDetails.text);
        return response.result?.value;
      };
      await evaluate('(()=>{const input=document.querySelector("#formula");input.value="Maybe";input.dispatchEvent(new Event("input",{bubbles:true}));input.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))})()');
      const warningDeadline = Date.now() + 5000;
      while (Date.now() < warningDeadline && await evaluate('document.querySelector("#cell-0-0")?.textContent') !== 'Maybe') await new Promise(resolve => setTimeout(resolve, 25));
      await evaluate('document.querySelector("#cell-0-1").click();(()=>{const input=document.querySelector("#formula");input.value="Perhaps";input.dispatchEvent(new Event("input",{bubbles:true}));input.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))})()');
      const informationDeadline = Date.now() + 5000;
      while (Date.now() < informationDeadline && await evaluate('document.querySelector("#cell-0-1")?.textContent') !== 'Perhaps') await new Promise(resolve => setTimeout(resolve, 25));
      assert.deepEqual(prompts.map(prompt => prompt.type), ['warning', 'information']);
      assert.match(prompts[0].message, /Check: Value is outside the list/);
      assert.match(prompts[1].message, /Notice: Value is outside the list/);
      await vscode.commands.executeCommand('workbench.action.files.save');
      const saved = analyze(await fs.readFile(file));
      assert.equal(saved.sheets[0].cells.get('A1').value, 'Maybe');
      assert.equal(saved.sheets[0].cells.get('B1').value, 'Perhaps');
      assert.ok(saved.sheets[0].xml.includes(validation));
    } finally {
      vscode.window.showWarningMessage = originalWarning;
      vscode.window.showInformationMessage = originalInformation;
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview resizes row boundaries with mouse and keyboard and double-clicks AutoFit', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-row-boundary-host-')), file = path.join(directory, 'row-boundary.xlsx');
    await fs.copyFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Sample.xlsx'), file);
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const result = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
        return result.result?.value;
      };
      const waitHeight = async height => {
        const deadline = Date.now() + 5000;
        while (Date.now() < deadline) {
          if (await evaluate('document.querySelector("#cell-0-0")?.getBoundingClientRect().height') === height) return;
          await new Promise(resolve => setTimeout(resolve, 25));
        }
        assert.fail('Row boundary did not set height ' + height);
      };
      const boundaryPoint = async () => {
        const deadline = Date.now() + 5000;
        while (Date.now() < deadline) {
          const point = await evaluate(`(() => {const handle=document.querySelector('.rowhead[data-view-row="0"] .row-resize-handle');if(!handle)return null;const r=handle.getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+r.height/2}})()`);
          if (point) return point;
          await new Promise(resolve => setTimeout(resolve, 25));
        }
        assert.fail('Row boundary was not rendered after model refresh');
      };
      const mouse = (type, point, count = 1) => devtools.send('Input.dispatchMouseEvent', { type, x: point.x + devtools.frameOffset.x, y: point.y + devtools.frameOffset.y, button: type === 'mouseMoved' ? 'none' : 'left', buttons: type === 'mouseReleased' ? 0 : 1, clickCount: count });
      await waitHeight(20);
      const point = await boundaryPoint();
      await mouse('mousePressed', point); await mouse('mouseMoved', { ...point, y: point.y + 32 }); await mouse('mouseReleased', { ...point, y: point.y + 32 });
      await waitHeight(52);
      const dirtyDeadline = Date.now() + 5000;
      while (Date.now() < dirtyDeadline && !vscode.window.tabGroups.activeTabGroup.activeTab?.isDirty) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(vscode.window.tabGroups.activeTabGroup.activeTab?.isDirty, true);
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(analyze(await fs.readFile(file)).sheets[0].rowHeights[0], 52);
      await vscode.commands.executeCommand('jinsheet.undo'); await waitHeight(20);
      await vscode.commands.executeCommand('workbench.action.files.save'); assert.equal(analyze(await fs.readFile(file)).sheets[0].rowHeights[0], undefined);
      await vscode.commands.executeCommand('jinsheet.redo'); await waitHeight(52);
      await vscode.commands.executeCommand('workbench.action.files.save');
      const next = await boundaryPoint();
      for (const count of [1, 2]) { await mouse('mousePressed', next, count); await mouse('mouseReleased', next, count); }
      const fitDeadline = Date.now() + 5000;
      let fittedHeight = 52;
      while (Date.now() < fitDeadline) {
        fittedHeight = await evaluate('document.querySelector("#cell-0-0")?.getBoundingClientRect().height||52');
        if (fittedHeight < 52) break;
        await new Promise(resolve => setTimeout(resolve, 25));
      }
      assert.ok(fittedHeight < 52, 'double-click auto-fits to the actual font height: ' + JSON.stringify(await evaluate('({height:document.querySelector("#cell-0-0")?.style.height,notice:document.querySelector("#notice").textContent})')));
      await vscode.commands.executeCommand('workbench.action.files.save');
      const saved = analyze(await fs.readFile(file)); assert.ok(Math.abs(saved.sheets[0].rowHeights[0] - fittedHeight) < .05); assert.equal(saved.sheets[0].customRowHeights.includes(0), false);
      const handleFocusDeadline = Date.now() + 5000; let handleFocused = false;
      while (Date.now() < handleFocusDeadline) {
        handleFocused = await evaluate('(()=>{const handle=document.querySelector(\'.rowhead[data-view-row="0"] .row-resize-handle\');if(!handle)return false;handle.focus();return document.activeElement===handle})()');
        if (handleFocused) break;
        await new Promise(resolve => setTimeout(resolve, 25));
      }
      assert.equal(handleFocused, true, 'focus a currently rendered row boundary after save');
      for (let index = 0; index < 2; index++) {
        const previous = await evaluate('(()=>{window.__rowHandleBeforeKey=document.activeElement;return Number(document.activeElement.getAttribute("aria-valuenow"))})()');
        await devtools.send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'ArrowDown', code: 'ArrowDown', windowsVirtualKeyCode: 40 });
        await devtools.send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'ArrowDown', code: 'ArrowDown', windowsVirtualKeyCode: 40 });
        const focusDeadline = Date.now() + 5000;
        const restored = 'document.activeElement!==window.__rowHandleBeforeKey&&document.activeElement.classList.contains("row-resize-handle")';
        while (Date.now() < focusDeadline && !await evaluate(restored)) await new Promise(resolve => setTimeout(resolve, 25));
        assert.equal(await evaluate(restored), true, 'a model refresh keeps the keyboard on the row boundary');
        assert.equal(await evaluate('Number(document.activeElement.getAttribute("aria-valuenow"))'), previous + 1);
      }
      await vscode.commands.executeCommand('workbench.action.files.save');
    } finally {
      devtools?.close(); await vscode.commands.executeCommand('workbench.action.closeAllEditors'); await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview groups rows through save, collapse, expand and undo', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-outline-host-')), file = path.join(directory, 'outline.xlsx');
    await fs.copyFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx'), file);
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const result = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
        return result.result?.value;
      };
      const wait = async condition => {
        const deadline = Date.now() + 6000;
        while (Date.now() < deadline) { if (await evaluate(condition)) return; await new Promise(resolve => setTimeout(resolve, 25)); }
        assert.fail('Outline group condition failed: ' + condition);
      };
      await wait('Boolean(document.querySelector("#cell-2-0"))');
      await evaluate('document.querySelector("#cell-1-0").click();const target=document.querySelector("#cell-2-0");target.dispatchEvent(new MouseEvent("mousedown",{bubbles:true,shiftKey:true}));target.dispatchEvent(new MouseEvent("click",{bubbles:true,shiftKey:true}));document.dispatchEvent(new MouseEvent("mouseup",{bubbles:true}));document.querySelector("#groupRows").click()');
      await wait('document.querySelector("#notice").textContent.includes("Modified")');
      await wait('document.querySelector("#collapseRows").disabled===false');
      await wait('Boolean(document.querySelector(".rowhead[data-row=\\"1\\"] .outline-row-bracket.outline-bracket-start"))');
      await wait('Boolean(document.querySelector(".rowhead[data-row=\\"2\\"] .outline-row-bracket.outline-bracket-end"))');
      await wait('document.querySelector(".rowhead[data-row=\\"3\\"] .outline-toggle")?.textContent==="−"');
      await wait('(()=>{const toggle=document.querySelector(".rowhead[data-row=\\"3\\"] .outline-toggle");if(!toggle)return false;toggle.click();return true})()');
      await wait('document.querySelector("#expandRows").disabled===false');
      await vscode.commands.executeCommand('workbench.action.files.save');
      let saved = analyze(await fs.readFile(file));
      assert.deepEqual(saved.sheets[0].rowGroups, [{ start: 1, end: 2, level: 1, summary: 3, collapsed: true }]);
      await vscode.commands.executeCommand('jinsheet.undo');
      await wait('document.querySelector("#collapseRows").disabled===false');
      await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file));
      assert.deepEqual(saved.sheets[0].rowGroups, [{ start: 1, end: 2, level: 1, summary: 3, collapsed: false }]);
      await vscode.commands.executeCommand('jinsheet.redo');
      await wait('document.querySelector("#expandRows").disabled===false');
      await wait('document.querySelector(".rowhead[data-row=\\"3\\"] .outline-toggle")?.textContent==="+"');
      await wait('(()=>{const toggle=document.querySelector(".rowhead[data-row=\\"3\\"] .outline-toggle");if(!toggle)return false;toggle.click();return true})()');
      await wait('document.querySelector("#expandRows").disabled===true');
      await evaluate('document.querySelector("#ungroupRows").click()');
      await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file));
      assert.deepEqual(saved.sheets[0].rowGroups, []);
      assert.deepEqual(saved.sheets[0].outlineHiddenRows, []);
    } finally {
      devtools?.close(); await vscode.commands.executeCommand('workbench.action.closeAllEditors'); await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview CSV summary controls collapse and expand session groups without changing the file', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-csv-outline-host-'));
    const file = path.join(directory, 'outline.csv'), original = Buffer.from('Name,Amount\r\nAlpha,10\r\nBeta,20\r\nTotal,30\r\n');
    await fs.writeFile(file, original);
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const result = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
        return result.result?.value;
      };
      const wait = async condition => {
        const deadline = Date.now() + 6000;
        while (Date.now() < deadline) { if (await evaluate(condition)) return; await new Promise(resolve => setTimeout(resolve, 25)); }
        assert.fail('CSV outline condition failed: ' + condition);
      };
      await wait('Boolean(document.querySelector("#cell-1-0"))&&Boolean(document.querySelector("#cell-2-0"))');
      assert.equal(await evaluate('document.querySelector("#csvOutlineSessionNote").hidden'), false);
      await evaluate('document.querySelector("#cell-1-0").click();const target=document.querySelector("#cell-2-0");target.dispatchEvent(new MouseEvent("mousedown",{bubbles:true,shiftKey:true}));target.dispatchEvent(new MouseEvent("click",{bubbles:true,shiftKey:true}));document.dispatchEvent(new MouseEvent("mouseup",{bubbles:true}));document.querySelector("#groupRows").click()');
      await wait('document.querySelector(".rowhead[data-row=\\"3\\"] .outline-toggle")?.textContent==="−"');
      await evaluate('document.querySelector(".rowhead[data-row=\\"3\\"] .outline-toggle").click()');
      await wait('!document.querySelector("#cell-1-0")&&!document.querySelector("#cell-2-0")&&document.querySelector(".rowhead[data-row=\\"3\\"] .outline-toggle")?.textContent==="+"');
      await evaluate('document.querySelector(".rowhead[data-row=\\"3\\"] .outline-toggle").click()');
      await wait('Boolean(document.querySelector("#cell-1-0"))&&Boolean(document.querySelector("#cell-2-0"))&&document.querySelector(".rowhead[data-row=\\"3\\"] .outline-toggle")?.textContent==="−"');
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.deepEqual(await fs.readFile(file), original, 'session-only outline changes do not alter the CSV data');
    } finally {
      devtools?.close(); await vscode.commands.executeCommand('workbench.action.closeAllEditors'); await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview summary toggle controls nested groups sharing one row', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-outline-summary-host-')), file = path.join(directory, 'outline-summary.xlsx');
    let grouped = analyze(await fs.readFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx')));
    grouped = patchOutlineGroups(grouped, [{ sheetIndex: 0, axis: 'row', action: 'group', start: 1, end: 2 }]).model;
    grouped = patchOutlineGroups(grouped, [{ sheetIndex: 0, axis: 'row', action: 'group', start: 2, end: 2 }]).model;
    assert.deepEqual(grouped.sheets[0].rowGroups.map(({ start, end, level, summary }) => ({ start, end, level, summary })), [
      { start: 1, end: 2, level: 1, summary: 3 },
      { start: 2, end: 2, level: 2, summary: 3 }
    ]);
    await fs.writeFile(file, grouped.bytes);
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const result = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
        return result.result?.value;
      };
      const wait = async condition => {
        const deadline = Date.now() + 6000;
        while (Date.now() < deadline) { if (await evaluate(condition)) return; await new Promise(resolve => setTimeout(resolve, 25)); }
        assert.fail('Outline summary group condition failed: ' + condition);
      };
      await wait('Boolean(document.querySelector("#cell-2-0"))');
      await wait('Boolean(document.querySelector(".rowhead[data-row=\\"3\\"] .outline-toggle"))');
      assert.equal(await evaluate('document.querySelectorAll(".rowhead[data-row=\\"3\\"] .outline-toggle").length'), 1);
      await wait('(()=>{const toggle=document.querySelector(".rowhead[data-row=\\"3\\"] .outline-toggle");if(!toggle)return false;toggle.click();return true})()');
      await wait('document.querySelector(".rowhead[data-row=\\"3\\"] .outline-toggle")?.textContent==="+"');
      await wait('!document.querySelector("#cell-1-0")&&!document.querySelector("#cell-2-0")');
      await vscode.commands.executeCommand('workbench.action.files.save');
      let saved = analyze(await fs.readFile(file));
      assert.deepEqual(saved.sheets[0].rowGroups.map(group => group.collapsed), [true, true]);
      assert.deepEqual(saved.sheets[0].outlineHiddenRows, [1, 2]);

      await wait('Boolean(document.querySelector(".rowhead[data-row=\\"3\\"] .outline-toggle"))');
      await wait('(()=>{const toggle=document.querySelector(".rowhead[data-row=\\"3\\"] .outline-toggle");if(!toggle)return false;toggle.click();return true})()');
      await wait('document.querySelector(".rowhead[data-row=\\"3\\"] .outline-toggle")?.textContent==="−"');
      await wait('Boolean(document.querySelector("#cell-1-0"))&&Boolean(document.querySelector("#cell-2-0"))');
      await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file));
      assert.deepEqual(saved.sheets[0].rowGroups.map(group => group.collapsed), [false, false]);
      assert.deepEqual(saved.sheets[0].outlineHiddenRows, []);
    } finally {
      devtools?.close(); await vscode.commands.executeCommand('workbench.action.closeAllEditors'); await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview outline level buttons save the selected row depth', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-outline-level-host-')), file = path.join(directory, 'outline-level.xlsx');
    let grouped = analyze(await fs.readFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx')));
    grouped = patchOutlineGroups(grouped, [{ sheetIndex: 0, axis: 'row', action: 'group', start: 1, end: 2 }]).model;
    await fs.writeFile(file, grouped.bytes);
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const result = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
        return result.result?.value;
      };
      const wait = async condition => {
        const deadline = Date.now() + 6000;
        while (Date.now() < deadline) { if (await evaluate(condition)) return; await new Promise(resolve => setTimeout(resolve, 25)); }
        assert.fail('Outline level condition failed: ' + condition);
      };
      await wait('Boolean(document.querySelector("#cell-2-0"))');
      await wait('document.querySelectorAll("#rowOutlineLevels button").length===2');
      assert.equal(await evaluate('document.querySelector(".rowhead[data-row=\\"3\\"] .outline-toggle")?.textContent'), '−');
      await evaluate('document.querySelector("#rowOutlineLevels button:nth-child(1)").click()');
      await wait('document.querySelector(".rowhead[data-row=\\"3\\"] .outline-toggle")?.textContent==="+"');
      await wait('!document.querySelector("#cell-1-0")&&!document.querySelector("#cell-2-0")');
      await vscode.commands.executeCommand('workbench.action.files.save');
      let saved = analyze(await fs.readFile(file));
      assert.deepEqual(saved.sheets[0].rowGroups.map(group => group.collapsed), [true]);
      await wait('Boolean(document.querySelector(".rowhead[data-row=\\"3\\"] .outline-toggle"))');
      await evaluate('document.querySelector("#rowOutlineLevels button:nth-child(2)").click()');
      await wait('document.querySelector(".rowhead[data-row=\\"3\\"] .outline-toggle")?.textContent==="−"');
      await wait('Boolean(document.querySelector("#cell-1-0"))&&Boolean(document.querySelector("#cell-2-0"))');
      await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file));
      assert.deepEqual(saved.sheets[0].rowGroups.map(group => group.collapsed), [false]);
      assert.deepEqual(saved.sheets[0].outlineHiddenRows, []);
    } finally {
      devtools?.close(); await vscode.commands.executeCommand('workbench.action.closeAllEditors'); await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview edits row heights and auto-fits wrap changes through save, undo and redo', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-row-heights-host-')), file = path.join(directory, 'heights.xlsx');
    const original = analyze(await fs.readFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx'))), files = { ...original.files };
    const sheetPath = original.sheets[0].path;
    files[sheetPath] = new TextEncoder().encode(original.sheets[0].xml.replace('<sheetData>', '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultColWidth="20.5703125" defaultRowHeight="18"/><sheetData>').replace('<row r="1">', '<row r="1" ht="60" customHeight="1">').replace('<row r="2">', '<row r="2" ht="30" customHeight="1">').replace('Product', 'First line\nSecond line'));
    files[original.stylesPath] = new TextEncoder().encode(new TextDecoder().decode(files[original.stylesPath]).replace('<alignment horizontal="center" vertical="center"/>', '<alignment horizontal="center" vertical="center" wrapText="1"/>'));
    await fs.writeFile(file, zipSync(files));
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const result = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
        return result.result?.value;
      };
      const deadline = Date.now() + 5000;
      while (Date.now() < deadline && !await evaluate('Boolean(document.querySelector("#cell-2-0"))')) await new Promise(resolve => setTimeout(resolve, 25));
      const before = await evaluate(`(() => {const first=document.querySelector('#cell-0-0'),second=document.querySelector('#cell-1-0'),third=document.querySelector('#cell-2-0');return{height:first.getBoundingClientRect().height,second:second.getBoundingClientRect().height,offset:third.getBoundingClientRect().top-first.getBoundingClientRect().top,wrap:getComputedStyle(first).whiteSpace,top:first.getBoundingClientRect().top-document.querySelector('#grid').getBoundingClientRect().top,text:first.textContent}})()`);
      assert.equal(before.height, 80); assert.equal(before.second, 40); assert.equal(before.offset, 120); assert.equal(before.wrap, 'pre-wrap'); assert.equal(before.text, 'First line\nSecond line');
      await evaluate(`document.querySelector('#grid').scrollTop=100`);
      await evaluate('new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)))');
      assert.equal(await evaluate('document.querySelector("#cell-0-0").getBoundingClientRect().top-document.querySelector("#grid").getBoundingClientRect().top'), before.top);
      await vscode.commands.executeCommand('workbench.action.files.save');
      const saved = analyze(await fs.readFile(file));
      assert.deepEqual(saved.sheets[0].rowHeights, { 0: 80, 1: 40 }); assert.equal(saved.sheets[0].defaultRowHeight, 24);
      const waitForHeight = async height => {
        const end = Date.now() + 5000;
        while (Date.now() < end) {
          if (await evaluate('document.querySelector("#cell-0-0")?.getBoundingClientRect().height') === height) return;
          await new Promise(resolve => setTimeout(resolve, 25));
        }
        assert.fail('Row height did not become ' + height);
      };
      await evaluate(`(() => {document.querySelector('#rowHeight').value='45';document.querySelector('#applyRowHeight').click()})()`);
      await waitForHeight(60); await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(analyze(await fs.readFile(file)).sheets[0].rowHeights[0], 60);
      await vscode.commands.executeCommand('jinsheet.undo'); await waitForHeight(80);
      await vscode.commands.executeCommand('workbench.action.files.save'); assert.equal(analyze(await fs.readFile(file)).sheets[0].rowHeights[0], 80);
      await vscode.commands.executeCommand('jinsheet.redo'); await waitForHeight(60);
      await evaluate('document.querySelector("#resetRowHeight").click()');
      const resetDeadline = Date.now() + 5000;
      while (Date.now() < resetDeadline && !await evaluate('document.querySelector("#cell-0-0")?.getBoundingClientRect().height>24&&document.querySelector("#cell-0-0")?.getBoundingClientRect().height<60')) await new Promise(resolve => setTimeout(resolve, 25));
      await vscode.commands.executeCommand('workbench.action.files.save');
      const reset = analyze(await fs.readFile(file));
      assert.equal(reset.sheets[0].rowHeights[1], 40); assert.ok(reset.sheets[0].rowHeights[0] > 24 && reset.sheets[0].rowHeights[0] < 60, 'reset returns a wrapped row to automatic sizing'); assert.equal(reset.sheets[0].cells.get('A1').styleProperties.wrapText, true);
      await evaluate('document.querySelector("#autoFitRow").click()');
      const fitDeadline = Date.now() + 5000;
      let fittedHeight = 0;
      while (Date.now() < fitDeadline) {
        fittedHeight = await evaluate('document.querySelector("#cell-0-0")?.getBoundingClientRect().height||0');
        if (fittedHeight > 24) break;
        await new Promise(resolve => setTimeout(resolve, 25));
      }
      assert.ok(fittedHeight > 24, 'auto-fit grows a two-line row');
      assert.equal(await evaluate('(()=>{const cell=document.querySelector("#cell-0-0");return cell.scrollHeight<=cell.clientHeight+1})()'), true, 'wrapped text fits without vertical clipping');
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(analyze(await fs.readFile(file)).sheets[0].customRowHeights.includes(0), false);
      await evaluate('document.querySelector("#wrapText").click()');
      const wrapDeadline = Date.now() + 5000;
      while (Date.now() < wrapDeadline && await evaluate('document.querySelector("#cell-0-0")?getComputedStyle(document.querySelector("#cell-0-0")).whiteSpace:"missing"') !== 'nowrap') await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('document.querySelector("#cell-0-0").getBoundingClientRect().height<' + fittedHeight), true, 'disabling wrap also shrinks an automatic row');
      await vscode.commands.executeCommand('jinsheet.undo'); await waitForHeight(fittedHeight);
      assert.equal(await evaluate('getComputedStyle(document.querySelector("#cell-0-0")).whiteSpace'), 'pre-wrap', 'one undo restores both wrap and row height');
      await vscode.commands.executeCommand('jinsheet.redo');
      const redoDeadline = Date.now() + 5000;
      while (Date.now() < redoDeadline && await evaluate('document.querySelector("#cell-0-0")?getComputedStyle(document.querySelector("#cell-0-0")).whiteSpace:"missing"') !== 'nowrap') await new Promise(resolve => setTimeout(resolve, 25));
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(Boolean(analyze(await fs.readFile(file)).sheets[0].cells.get('A1').styleProperties.wrapText), false);
    } finally {
      devtools?.close(); await vscode.commands.executeCommand('workbench.action.closeAllEditors'); await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview recalculates automatic rows after text and width changes without extra undo', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-auto-rows-host-')), file = path.join(directory, 'automatic.xlsx');
    await fs.copyFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx'), file);
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const result = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
        return result.result?.value;
      };
      const wait = async expression => {
        const deadline = Date.now() + 6000;
        while (Date.now() < deadline) { if (await evaluate(expression)) return; await new Promise(resolve => setTimeout(resolve, 25)); }
        assert.fail('Automatic row condition failed: ' + expression + ' ' + JSON.stringify(await evaluate('({height:document.querySelector("#cell-0-0")?.style.height,width:document.querySelector("#cell-0-0")?.style.width,notice:document.querySelector("#notice")?.textContent})')));
      };
      const height = 'document.querySelector("#cell-0-0")?.getBoundingClientRect().height';
      await wait('Boolean(document.querySelector("#cell-0-0"))');
      await evaluate('document.querySelector("#cell-0-0").click();document.querySelector("#wrapText").click()');
      await wait('document.querySelector("#cell-0-0")&&getComputedStyle(document.querySelector("#cell-0-0")).whiteSpace==="pre-wrap"');
      const initialHeight = await evaluate(height), text = 'Automatic row sizing follows content and column width changes without losing the text being edited. '.repeat(1);
      await evaluate(`(() => {const input=document.querySelector('#formula');input.value=${JSON.stringify(text)};input.dispatchEvent(new Event('input',{bubbles:true}));input.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}))})()`);
      await wait(height + '>' + initialHeight);
      const contentHeight = await evaluate(height);
      await vscode.commands.executeCommand('jinsheet.undo');
      await wait('document.querySelector("#cell-0-0")?.textContent==="Product"&&' + height + '===' + initialHeight);
      await vscode.commands.executeCommand('jinsheet.redo'); await wait(height + '===' + contentHeight);
      const previousWidth = await evaluate('document.querySelector("#cell-0-0").getBoundingClientRect().width');
      await evaluate(`for(let i=0;i<6;i++)document.querySelector('.colhead[data-column="0"] .col-resize-handle').dispatchEvent(new KeyboardEvent('keydown',{key:'ArrowLeft',bubbles:true}))`);
      await wait('document.querySelector("#cell-0-0")?.getBoundingClientRect().width<' + previousWidth);
      await wait(height + '>' + contentHeight);
      for (let i = 0; i < 6; i++) await vscode.commands.executeCommand('jinsheet.undo');
      await wait(height + '===' + contentHeight);
      await vscode.commands.executeCommand('workbench.action.files.save');
      const saved = analyze(await fs.readFile(file));
      assert.equal(saved.sheets[0].cells.get('A1').value, text); assert.equal(saved.sheets[0].rowHeights[0], contentHeight); assert.deepEqual(saved.sheets[0].customRowHeights, []);
    } finally {
      devtools?.close(); await vscode.commands.executeCommand('workbench.action.closeAllEditors'); await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview saves page setup and exports a selected printable sheet without changing the source', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-print-host-')), file = path.join(directory, 'print.xlsx'), output = path.join(directory, 'print.html');
    await fs.copyFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx'), file);
    const originalDialog = vscode.window.showSaveDialog, originalOpen = vscode.env.openExternal, opened = [];
    vscode.window.showSaveDialog = async () => vscode.Uri.file(output);
    vscode.env.openExternal = async target => { opened.push(target.fsPath); return true; };
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const result = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
        return result.result?.value;
      };
      const wait = async expression => {
        const deadline = Date.now() + 5000;
        while (Date.now() < deadline) { if (await evaluate(expression)) return; await new Promise(resolve => setTimeout(resolve, 25)); }
        assert.fail('Page setup condition failed: ' + expression);
      };
      await wait('Boolean(document.querySelector("#cell-0-0"))');
      await evaluate(`document.querySelector('#printPanel').open=true;document.querySelector('#printPaper').value='1';document.querySelector('#printOrientation').value='landscape';document.querySelector('#printScale').value='75';document.querySelector('#printGridlines').checked=true;document.querySelector('#printMargin-left').value='0.5';document.querySelector('#applyPageSettings').click()`);
      await wait('document.querySelector("#notice").textContent.includes("Modified")');
      await vscode.commands.executeCommand('workbench.action.files.save');
      const saved = analyze(await fs.readFile(file));
      assert.equal(saved.sheets[0].printSettings.orientation, 'landscape'); assert.equal(saved.sheets[0].printSettings.paperSize, 1); assert.equal(saved.sheets[0].printSettings.scale, 75); assert.equal(saved.sheets[0].printSettings.margins.left, .5);
      await vscode.commands.executeCommand('jinsheet.undo'); await wait('document.querySelector("#printOrientation").value==="portrait"');
      await vscode.commands.executeCommand('jinsheet.redo'); await wait('document.querySelector("#printOrientation").value==="landscape"');
      await vscode.commands.executeCommand('workbench.action.files.save');
      const source = await fs.readFile(file);
      await evaluate('document.querySelector("#printSelection").click()');
      const deadline = Date.now() + 5000;
      while (Date.now() < deadline && !opened.length) await new Promise(resolve => setTimeout(resolve, 25));
      assert.deepEqual(opened.map(normalizePath), [normalizePath(output)]);
      const html = await fs.readFile(output, 'utf8');
      assert.match(html, />Product</); assert.doesNotMatch(html, />Pen</); assert.match(html, /size:11in 8\.5in/); assert.match(html, /transform:scale\(0\.75\)/); assert.match(html, /window\.print\(\)/);
      const printRows = Array.from({ length: 80 }, (_, index) => index), printColumns = Array.from({ length: 6 }, (_, index) => index);
      const pagination = renderPrintDocument({ title: 'Pagination', rows: printRows, columns: printColumns, widths: Object.fromEntries(printColumns.map(index => [index, 144])), heights: Object.fromEntries(printRows.map(index => [index, 52])), cells: new Map(printRows.flatMap(row => printColumns.map(column => [row + ':' + column, { displayValue: 'Wrapped print content '.repeat(10), style: { wrapText: true, fontSize: 14 } }]))), settings: { ...DEFAULT_PRINT_SETTINGS, headings: true } });
      const layout = await measurePrintPages(evaluate, pagination);
      assert.ok(layout.length > 1);
      for (const page of layout) { assert.ok(page.width <= page.availableWidth + 1, JSON.stringify(page)); assert.ok(page.height <= page.availableHeight + 1, JSON.stringify(page)); }
      assert.deepEqual(await fs.readFile(file), source); assert.equal(vscode.window.tabGroups.activeTabGroup.activeTab?.isDirty, false);
      vscode.window.showSaveDialog = async () => undefined;
      await evaluate('document.querySelector("#printSheet").click()');
      await evaluate('new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)))');
      assert.equal(opened.length, 1, 'canceling export does not open another browser page');
      await evaluate('document.querySelector("#printScale").value="1";document.querySelector("#applyPageSettings").click()');
      await wait('document.querySelector("#notice").textContent.includes("Choose valid page settings")');
      assert.equal(vscode.window.tabGroups.activeTabGroup.activeTab?.isDirty, false);
    } finally {
      vscode.window.showSaveDialog = originalDialog; vscode.env.openExternal = originalOpen;
      devtools?.close(); await vscode.commands.executeCommand('workbench.action.closeAllEditors'); await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview saves and clears print areas and follows row insertion through undo and export', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-print-area-host-')), file = path.join(directory, 'area.xlsx'), output = path.join(directory, 'area.html');
    await fs.copyFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx'), file);
    const originalDialog = vscode.window.showSaveDialog, originalOpen = vscode.env.openExternal, opened = [];
    vscode.window.showSaveDialog = async () => vscode.Uri.file(output);
    vscode.env.openExternal = async target => { opened.push(target.fsPath); return true; };
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const result = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
        return result.result?.value;
      };
      const wait = async expression => {
        const deadline = Date.now() + 5000;
        while (Date.now() < deadline) { if (await evaluate(expression)) return; await new Promise(resolve => setTimeout(resolve, 25)); }
        assert.fail('Print area condition failed: ' + expression);
      };
      await wait('Boolean(document.querySelector("#cell-0-0"))');
      await evaluate('document.querySelector("#cell-0-0").click();document.querySelector("#printPanel").open=true;document.querySelector("#setPrintArea").click()');
      await wait('document.querySelector("#printAreaStatus").textContent==="Print area: A1:A1"');
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.deepEqual(analyze(await fs.readFile(file)).sheets[0].printAreas[0], { sheetName: 'Inventory', top: 0, bottom: 0, left: 0, right: 0 });
      await evaluate('document.querySelector("#printSheet").click()');
      const deadline = Date.now() + 5000;
      while (Date.now() < deadline && !opened.length) await new Promise(resolve => setTimeout(resolve, 25));
      assert.deepEqual(opened.map(normalizePath), [normalizePath(output)]); const html = await fs.readFile(output, 'utf8');
      assert.match(html, />Product</); assert.doesNotMatch(html, />Pen</); assert.equal((html.match(/<section class="page"/g) || []).length, 1);
      await evaluate('document.querySelector("#clearPrintArea").click()');
      await wait('document.querySelector("#clearPrintArea").disabled');
      const clearedAreaDeadline = Date.now() + 5000;
      let clearedAreas;
      do { await vscode.commands.executeCommand('workbench.action.files.save'); clearedAreas = analyze(await fs.readFile(file)).sheets[0].printAreas; if (!clearedAreas.length) break; await new Promise(resolve => setTimeout(resolve, 25)); } while (Date.now() < clearedAreaDeadline);
      assert.deepEqual(clearedAreas, []);
      await vscode.commands.executeCommand('jinsheet.undo'); await wait('document.querySelector("#printAreaStatus").textContent==="Print area: A1:A1"');
      await vscode.commands.executeCommand('jinsheet.redo'); await wait('document.querySelector("#clearPrintArea").disabled');
      await vscode.commands.executeCommand('jinsheet.undo'); await wait('document.querySelector("#printAreaStatus").textContent==="Print area: A1:A1"');
      await evaluate('document.querySelector("#insertRow").click()');
      await wait('document.querySelector("#printAreaStatus").textContent==="Print area: A2:A2"');
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(analyze(await fs.readFile(file)).sheets[0].printAreas[0].top, 1);
      await vscode.commands.executeCommand('jinsheet.undo'); await wait('document.querySelector("#printAreaStatus").textContent==="Print area: A1:A1"');
      await vscode.commands.executeCommand('workbench.action.files.save'); assert.equal(analyze(await fs.readFile(file)).sheets[0].printAreas[0].top, 0);
    } finally {
      vscode.window.showSaveDialog = originalDialog; vscode.env.openExternal = originalOpen;
      devtools?.close(); await vscode.commands.executeCommand('workbench.action.closeAllEditors'); await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview repeats saved print titles outside the area and keeps every page within its paper', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-print-titles-host-')), file = path.join(directory, 'titles.xlsx'), output = path.join(directory, 'titles.html');
    const original = analyze(await fs.readFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx'))), files = { ...original.files };
    const rows = Array.from({ length: 70 }, (_, row) => '<row r="' + (row + 1) + '">' + Array.from({ length: 8 }, (_, column) => '<c r="' + String.fromCharCode(65 + column) + (row + 1) + '" t="inlineStr"><is><t>' + (row === 0 && column === 0 ? 'Corner title' : row === 0 ? 'Header-' + column : column === 0 ? 'Row-' + row : 'Value-' + row + '-' + column) + '</t></is></c>').join('') + '</row>').join('');
    files[original.sheets[0].path] = new TextEncoder().encode(original.sheets[0].xml.replace(/<sheetData>[\s\S]*?<\/sheetData>/, '<sheetData>' + rows + '</sheetData>'));
    const configured = patchPrintArea(analyze(zipSync(files)), 0, { top: 1, bottom: 69, left: 1, right: 7 }).model;
    await fs.writeFile(file, patchPrintSettings(configured, 0, { ...configured.sheets[0].printSettings, headings: true }).bytes);
    const originalDialog = vscode.window.showSaveDialog, originalOpen = vscode.env.openExternal, opened = [];
    vscode.window.showSaveDialog = async () => vscode.Uri.file(output);
    vscode.env.openExternal = async target => { opened.push(target.fsPath); return true; };
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const result = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
        return result.result?.value;
      };
      const wait = async expression => {
        const deadline = Date.now() + 5000;
        while (Date.now() < deadline) { if (await evaluate(expression)) return; await new Promise(resolve => setTimeout(resolve, 25)); }
        assert.fail('Print titles condition failed: ' + expression);
      };
      await wait('Boolean(document.querySelector("#cell-0-0"))');
      await evaluate('document.querySelector("#printPanel").open=true;document.querySelector("#printTitleRows").value="1";document.querySelector("#printTitleColumns").value="A";document.querySelector("#applyPrintTitles").click()');
      await wait('document.querySelector("#printTitleRows").value==="1:1"&&document.querySelector("#printTitleColumns").value==="A:A"');
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.deepEqual(analyze(await fs.readFile(file)).sheets[0].printTitles, { rows: { start: 0, end: 0 }, columns: { start: 0, end: 0 } });
      await evaluate('document.querySelector("#printSheet").click()');
      const deadline = Date.now() + 7000;
      while (Date.now() < deadline && !opened.length) await new Promise(resolve => setTimeout(resolve, 25));
      assert.deepEqual(opened.map(normalizePath), [normalizePath(output)]);
      const html = await fs.readFile(output, 'utf8'), pages = html.match(/<section class="page"[\s\S]*?<\/section>/g);
      assert.ok(pages.length > 1);
      for (const page of pages) { assert.equal(page.split('>Corner title<').length - 1, 1); assert.match(page, /data-repeat-row="true"/); assert.match(page, /data-repeat-column="true"/); }
      assert.equal(html.split('>Value-69-7<').length - 1, 1);
      const layout = await measurePrintPages(evaluate, html);
      for (const page of layout) { assert.ok(page.width <= page.availableWidth + 1, JSON.stringify(page)); assert.ok(page.height <= page.availableHeight + 1, JSON.stringify(page)); }
      await vscode.commands.executeCommand('jinsheet.undo'); await wait('document.querySelector("#printTitleRows").value===""');
      await vscode.commands.executeCommand('jinsheet.redo'); await wait('document.querySelector("#printTitleRows").value==="1:1"');
      await evaluate('document.querySelector("#clearPrintTitles").click()');
      await wait('document.querySelector("#notice").textContent.includes("Modified")');
      await vscode.commands.executeCommand('workbench.action.files.save'); assert.deepEqual(analyze(await fs.readFile(file)).sheets[0].printTitles, { rows: null, columns: null });
      await vscode.commands.executeCommand('jinsheet.undo'); await wait('document.querySelector("#printTitleRows").value==="1:1"');
      await evaluate('document.querySelector("#insertRow").click()'); await wait('document.querySelector("#printTitleRows").value==="2:2"');
      await vscode.commands.executeCommand('workbench.action.files.save'); assert.equal(analyze(await fs.readFile(file)).sheets[0].printTitles.rows.start, 1);
    } finally {
      vscode.window.showSaveDialog = originalDialog; vscode.env.openExternal = originalOpen;
      devtools?.close(); await vscode.commands.executeCommand('workbench.action.closeAllEditors'); await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview saves header/footer variants and fits page fields inside print margins', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-header-footer-host-')), file = path.join(directory, 'headers.xlsx'), output = path.join(directory, 'headers.html');
    const original = analyze(await fs.readFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx'))), files = { ...original.files };
    const rows = Array.from({ length: 110 }, (_, row) => '<row r="' + (row + 1) + '"><c r="A' + (row + 1) + '" t="inlineStr"><is><t>Row ' + (row + 1) + '</t></is></c><c r="B' + (row + 1) + '"><v>' + row + '</v></c></row>').join('');
    files[original.sheets[0].path] = new TextEncoder().encode(original.sheets[0].xml.replace(/<sheetData>[\s\S]*?<\/sheetData>/, '<sheetData>' + rows + '</sheetData>'));
    await fs.writeFile(file, zipSync(files));
    const originalDialog = vscode.window.showSaveDialog, originalOpen = vscode.env.openExternal, opened = [];
    vscode.window.showSaveDialog = async () => vscode.Uri.file(output);
    vscode.env.openExternal = async target => { opened.push(target.fsPath); return true; };
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor'); devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const result = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
        return result.result?.value;
      };
      const wait = async expression => {
        const deadline = Date.now() + 5000;
        while (Date.now() < deadline) { if (await evaluate(expression)) return; await new Promise(resolve => setTimeout(resolve, 25)); }
        assert.fail('Header/footer condition failed: ' + expression);
      };
      await wait('Boolean(document.querySelector("#cell-0-0"))');
      await evaluate('document.querySelector("#printPanel").open=true;document.querySelector("#printUseFirstPage").checked=true;document.querySelector("#printFirstPage").value="10";document.querySelector("#applyPageSettings").click()');
      await wait('document.querySelector("#notice").textContent.includes("Modified")'); await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(analyze(await fs.readFile(file)).sheets[0].printSettings.firstPageNumber, 10);
      await evaluate(`(() => {for(const [variant,title] of [['odd','Odd'],['even','Even'],['first','First']]){const selector=document.querySelector('#hfVariant');selector.value=variant;selector.dispatchEvent(new Event('change',{bubbles:true}));document.querySelector('#hf-header-left').value='&BReport &F&B';document.querySelector('#hf-header-center').value=title;document.querySelector('#hf-header-right').value='&A';document.querySelector('#hf-footer-center').value='Page &P / &N'}document.querySelector('#hfDifferentFirst').checked=true;document.querySelector('#hfDifferentEven').checked=true;document.querySelector('#hfScale').checked=false;document.querySelector('#hfApply').click()})()`);
      await wait('document.querySelector("#hfVariant").value==="odd"&&document.querySelector("#hfDifferentFirst").checked'); await vscode.commands.executeCommand('workbench.action.files.save');
      const saved = analyze(await fs.readFile(file)).sheets[0]; assert.equal(saved.headerFooter.first.header.center, 'First'); assert.equal(saved.headerFooter.even.header.center, 'Even'); assert.equal(saved.headerFooter.odd.footer.center, 'Page &P / &N');
      await evaluate('document.querySelector("#printSheet").click()');
      const deadline = Date.now() + 7000; while (Date.now() < deadline && !opened.length) await new Promise(resolve => setTimeout(resolve, 25));
      assert.deepEqual(opened.map(normalizePath), [normalizePath(output)]); const html = await fs.readFile(output, 'utf8'), layout = await measurePrintPages(evaluate, html, true);
      assert.ok(layout.length >= 3);
      for (const [index, page] of layout.entries()) {
        assert.match(page.headerText, /Report headers.xlsx/); assert.match(page.headerText, /Inventory/);
        assert.ok(page.headerText.includes(index === 0 ? 'First' : (index + 10) % 2 === 0 ? 'Even' : 'Odd'));
        assert.equal(page.footerText, 'Page ' + (index + 10) + ' / ' + layout.length);
        assert.equal(page.hfFits, true, JSON.stringify(page)); assert.equal(page.clearOfBody, true, JSON.stringify(page)); assert.equal(page.printDisabled, false); assert.equal(page.printed, 1);
      }
      await vscode.commands.executeCommand('jinsheet.undo'); await wait('document.querySelector("#hf-header-left").value===""');
      await vscode.commands.executeCommand('jinsheet.redo'); await wait('document.querySelector("#hf-header-left").value.includes("Report")');
      await evaluate('document.querySelector("#hfClear").click()'); await wait('document.querySelector("#notice").textContent.includes("Modified")'); await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(analyze(await fs.readFile(file)).sheets[0].headerFooter.odd.header.left, '');
      const config = defaultHeaderFooter(); config.odd.header.center = 'Long text '.repeat(20);
      const projection = { title: 'Font fit', rows: [0], columns: [0], widths: { 0: 144 }, heights: { 0: 26 }, cells: new Map([['0:0', { displayValue: 'Body' }]]), headerFooter: config };
      const fitted = await measurePrintPages(evaluate, renderPrintDocument(projection), true); assert.equal(fitted[0].hfFits, true); assert.ok(fitted[0].fontSizes.some(size => size < 11));
      config.odd.header.center = '&409Too big';
      const invalid = await measurePrintPages(evaluate, renderPrintDocument(projection), true); assert.equal(invalid[0].printDisabled, true); assert.equal(invalid[0].printed, 0); assert.match(invalid[0].warning, /does not fit/);
    } finally {
      vscode.window.showSaveDialog = originalDialog; vscode.env.openExternal = originalOpen;
      devtools?.close(); await vscode.commands.executeCommand('workbench.action.closeAllEditors'); await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview edits manual page breaks through save, undo, structure and print fitting', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-page-break-host-')), file = path.join(directory, 'breaks.xlsx'), output = path.join(directory, 'breaks.html');
    const original = analyze(await fs.readFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx'))), files = { ...original.files };
    const rows = Array.from({ length: 4 }, (_, row) => '<row r="' + (row + 1) + '">' + Array.from({ length: 4 }, (_, column) => '<c r="' + String.fromCharCode(65 + column) + (row + 1) + '" t="inlineStr"><is><t>v' + row + '-' + column + '</t></is></c>').join('') + '</row>').join('');
    files[original.sheets[0].path] = new TextEncoder().encode(original.sheets[0].xml.replace(/<sheetData>[\s\S]*?<\/sheetData>/, '<sheetData>' + rows + '</sheetData>'));
    let model = patchPrintArea(analyze(zipSync(files)), 0, { top: 1, bottom: 3, left: 1, right: 3 }).model;
    model = patchPrintTitles(model, 0, { rows: { start: 0, end: 0 }, columns: { start: 0, end: 0 } }).model;
    const headers = defaultHeaderFooter(); headers.odd.footer.center = '&P / &N'; model = patchHeaderFooter(model, 0, headers).model; await fs.writeFile(file, model.bytes);
    const originalDialog = vscode.window.showSaveDialog, originalOpen = vscode.env.openExternal, opened = [];
    vscode.window.showSaveDialog = async () => vscode.Uri.file(output); vscode.env.openExternal = async target => { opened.push(target.fsPath); return true; };
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor'); devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const result = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
        return result.result?.value;
      };
      const wait = async expression => {
        const deadline = Date.now() + 5000;
        while (Date.now() < deadline) { if (await evaluate(expression)) return; await new Promise(resolve => setTimeout(resolve, 25)); }
        assert.fail('Page break condition failed: ' + expression + ' ' + JSON.stringify(await evaluate('({status:document.querySelector("#pageBreakStatus")?.textContent,notice:document.querySelector("#notice")?.textContent,address:document.querySelector("#address")?.value,rowMarkers:document.querySelectorAll(".manual-row-break").length,columnMarkers:document.querySelectorAll(".manual-column-break").length})')));
      };
      await wait('Boolean(document.querySelector("#cell-2-2"))');
      await evaluate('document.querySelector("#cell-2-2").click();document.querySelector("#printPanel").open=true;document.querySelector("#breakAddRow").click()');
      await wait('Boolean(document.querySelector(".rowhead.manual-row-break"))');
      await evaluate('document.querySelector("#breakAddColumn").click()'); await wait('Boolean(document.querySelector(".colhead.manual-column-break"))');
      await vscode.commands.executeCommand('workbench.action.files.save');
      let saved = analyze(await fs.readFile(file)).sheets[0]; assert.deepEqual(saved.pageBreaks.rows, [2]); assert.deepEqual(saved.pageBreaks.columns, [2]);
      await evaluate('document.querySelector("#breakRemoveRow").click()'); await wait('document.querySelector("#pageBreakStatus").textContent.includes("Before rows: None")');
      await vscode.commands.executeCommand('jinsheet.undo'); await wait('Boolean(document.querySelector(".rowhead.manual-row-break"))');
      await vscode.commands.executeCommand('jinsheet.redo'); await wait('document.querySelector("#pageBreakStatus").textContent.includes("Before rows: None")');
      await vscode.commands.executeCommand('jinsheet.undo'); await wait('Boolean(document.querySelector(".rowhead.manual-row-break"))');
      await evaluate('document.querySelector("#insertRow").click()'); await wait('document.querySelector("#pageBreakStatus").textContent.includes("Before rows: 4")');
      await vscode.commands.executeCommand('workbench.action.files.save'); assert.deepEqual(analyze(await fs.readFile(file)).sheets[0].pageBreaks.rows, [3]);
      await vscode.commands.executeCommand('jinsheet.undo'); await wait('document.querySelector("#pageBreakStatus").textContent.includes("Before rows: 3")'); await vscode.commands.executeCommand('workbench.action.files.save');
      const exportPrint = async count => {
        await evaluate('document.querySelector("#printSheet").click()'); const deadline = Date.now() + 7000;
        while (Date.now() < deadline && opened.length < count) await new Promise(resolve => setTimeout(resolve, 25));
        assert.equal(opened.length, count); return fs.readFile(output, 'utf8');
      };
      const html = await exportPrint(1), pages = html.match(/<section class="page"[\s\S]*?<\/section>/g); assert.equal(pages.length, 4);
      assert.match(pages[0], />v1-1</); assert.match(pages[1], />v2-1</); assert.match(pages[2], />v1-2</);
      for (const page of pages) assert.equal(page.split('>v0-0<').length - 1, 1, 'titles repeat once');
      const layout = await measurePrintPages(evaluate, html, true);
      for (const [index, page] of layout.entries()) { assert.equal(page.footerText, (index + 1) + ' / 4'); assert.ok(page.width <= page.availableWidth + 1); assert.ok(page.height <= page.availableHeight + 1); assert.equal(page.clearOfBody, true); }
      await evaluate('document.querySelector("#printFit").checked=true;document.querySelector("#printFitWidth").value="1";document.querySelector("#printFitHeight").value="1";document.querySelector("#applyPageSettings").click()');
      await wait('document.querySelector("#pageBreakStatus").textContent.includes("ignores manual")'); await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(((await exportPrint(2)).match(/<section class="page"/g) || []).length, 1);
      saved = analyze(await fs.readFile(file)).sheets[0]; assert.deepEqual(saved.pageBreaks.rows, [2]); assert.deepEqual(saved.pageBreaks.columns, [2]);
      await evaluate('document.querySelector("#breakReset").click()');
      await wait('document.querySelector("#breakReset").disabled&&!document.querySelector(".manual-row-break")&&!document.querySelector(".manual-column-break")');
      await vscode.commands.executeCommand('workbench.action.files.save');
      let resetDeadline = Date.now() + 3000, resetSheet = analyze(await fs.readFile(file)).sheets[0];
      while (resetSheet.pageBreaks.hasAny && Date.now() < resetDeadline) {
        await new Promise(resolve => setTimeout(resolve, 50));
        await vscode.commands.executeCommand('workbench.action.files.save');
        resetSheet = analyze(await fs.readFile(file)).sheets[0];
      }
      assert.equal(resetSheet.pageBreaks.hasAny, false);
      await vscode.commands.executeCommand('jinsheet.undo'); await wait('!document.querySelector("#breakReset").disabled'); await vscode.commands.executeCommand('workbench.action.files.save');
      assert.deepEqual(analyze(await fs.readFile(file)).sheets[0].pageBreaks.rows, [2]);
    } finally {
      vscode.window.showSaveDialog = originalDialog; vscode.env.openExternal = originalOpen;
      devtools?.close(); await vscode.commands.executeCommand('workbench.action.closeAllEditors'); await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium exports charts at drawing positions and supports table-only printing without changing the workbook', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-chart-print-host-')), file = path.join(directory, 'chart.xlsx'), output = path.join(directory, 'chart.html');
    const source = analyze(await fs.readFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx')));
    let model = patchAddChart(source, 0, { top: 0, bottom: 3, left: 0, right: 1 }).model;
    model = patchPrintSettings(model, 0, { ...model.sheets[0].printSettings, fitToPage: true, fitToWidth: 1, fitToHeight: 1 }).model;
    await fs.writeFile(file, model.bytes); const before = await fs.readFile(file);
    const originalDialog = vscode.window.showSaveDialog, originalOpen = vscode.env.openExternal, opened = [];
    vscode.window.showSaveDialog = async () => vscode.Uri.file(output); vscode.env.openExternal = async target => { opened.push(target.fsPath); return true; };
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor'); devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const result = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
        return result.result?.value;
      };
      const deadline = Date.now() + 5000;
      while (Date.now() < deadline && !await evaluate('Boolean(document.querySelector("#chartPreview svg"))')) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('document.querySelector("#printIncludeCharts").checked'), true);
      const exportPrint = async count => {
        await evaluate('document.querySelector("#printSheet").click()'); const limit = Date.now() + 7000;
        while (Date.now() < limit && opened.length < count) await new Promise(resolve => setTimeout(resolve, 25));
        assert.equal(opened.length, count); return fs.readFile(output, 'utf8');
      };
      const html = await exportPrint(1), layout = await measurePrintPages(evaluate, html, true);
      assert.equal(layout.length, 1); assert.ok(layout[0].charts.length > 0);
      for (const chart of layout[0].charts) { assert.ok(chart.width > 0 && chart.height > 0 && chart.svgWidth > 0 && chart.svgHeight > 0); assert.equal(chart.inside, true); }
      assert.equal(layout[0].printDisabled, false); assert.equal(layout[0].printed, 1);
      const indexes = Array.from({ length: 12 }, (_, index) => index), columns = [0, 1, 2, 3];
      const expandedChart = { type: 'bar', title: 'All series', series: Array.from({ length: 32 }, (_, index) => ({ name: 'Series-' + index, categories: [], values: Array.from({ length: 100 }, (_, point) => point + 1) })) };
      const projection = { title: 'Expanded chart', rows: indexes, columns, widths: Object.fromEntries(columns.map(index => [index, 144])), heights: Object.fromEntries(indexes.map(index => [index, 26])), rowPositions: Object.fromEntries(indexes.map(index => [index, index * 26])), columnPositions: Object.fromEntries(columns.map(index => [index, index * 144])), cells: new Map(), charts: [{ x: 0, y: 0, width: 576, height: 300, html: chartPreviewHtml([expandedChart], 'en', false, true) }] };
      const expanded = await measurePrintPages(evaluate, renderPrintDocument(projection), true);
      assert.equal(expanded.length, 1); assert.equal(expanded[0].charts[0].pointCount, 3200); assert.equal(expanded[0].charts[0].legendCount, 32); assert.equal(expanded[0].charts[0].legendFits, true);
      expandedChart.type = 'pie'; expandedChart.series = [{ name: 'Categories', categories: Array.from({ length: 80 }, (_, index) => 'Category-' + index), values: Array(80).fill(1) }];
      projection.charts[0].html = chartPreviewHtml([expandedChart], 'en', false, true);
      const pie = await measurePrintPages(evaluate, renderPrintDocument(projection), true);
      assert.equal(pie[0].charts[0].pointCount, 80); assert.equal(pie[0].charts[0].legendCount, 12); assert.equal(pie[0].charts[0].legendFits, true);
      const settingsDeadline = Date.now() + 10000;
      while (Date.now() < settingsDeadline && !await evaluate('Boolean(document.querySelector("#applyPageSettings")&&document.querySelector("#cell-4-4"))')) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('Boolean(document.querySelector("#applyPageSettings")&&document.querySelector("#cell-4-4"))'), true, 'page settings and the selected worksheet cells are rendered');
      await evaluate('document.querySelector("#printPanel").open=true');
      const panelDeadline = Date.now() + 5000;
      while (Date.now() < panelDeadline && !await evaluate('Boolean(document.querySelector("#applyPageSettings")&&!document.querySelector("#applyPageSettings").disabled)')) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('Boolean(document.querySelector("#applyPageSettings")&&!document.querySelector("#applyPageSettings").disabled)'), true, 'page settings are enabled after opening the print panel');
      await evaluate('document.querySelector("#printFit").checked=false;document.querySelector("#printScale").value="100";document.querySelector("#applyPageSettings").click()');
      const updateDeadline = Date.now() + 10000;
      while (Date.now() < updateDeadline && !await evaluate('document.querySelector("#notice").textContent.includes("Modified")')) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('document.querySelector("#notice").textContent.includes("Modified")'), true, 'page settings are applied before checking the worksheet viewport');
      const rowCellDeadline=Date.now()+10000;while(Date.now()<rowCellDeadline&&await evaluate('Boolean(document.querySelector("#applyPageSettings")&&!document.querySelector("#cell-4-4"))'))await new Promise(resolve=>setTimeout(resolve,25));
      assert.equal(await evaluate('Boolean(document.querySelector("#cell-4-4"))'),true,'worksheet cells are rendered after applying page settings');
      await evaluate('document.querySelector("#cell-4-4").click()');
      const rowSelectionDeadline=Date.now()+5000;while(Date.now()<rowSelectionDeadline&&await evaluate('document.querySelector("#address").value!=="E5"'))await new Promise(resolve=>setTimeout(resolve,25));
      assert.equal(await evaluate('document.querySelector("#address").value'), 'E5', 'the target row is selected before adding a page break');
      await evaluate('document.querySelector("#breakAddRow").click()');
      const rowLimit = Date.now() + 5000;
      while (Date.now() < rowLimit && !await evaluate('Boolean(document.querySelector(".manual-row-break"))')) await new Promise(resolve => setTimeout(resolve, 25));
      await evaluate('document.querySelector("#breakAddColumn").click()');
      const columnLimit = Date.now() + 5000;
      while (Date.now() < columnLimit && !await evaluate('Boolean(document.querySelector(".manual-column-break"))')) await new Promise(resolve => setTimeout(resolve, 25));
      const split = await exportPrint(2), splitLayout = await measurePrintPages(evaluate, split, true);
      assert.ok(splitLayout.length >= 4); assert.ok(splitLayout.filter(page => page.charts.length).length >= 4);
      for (const page of splitLayout) for (const chart of page.charts) assert.equal(chart.inside, true);
      await vscode.commands.executeCommand('jinsheet.undo'); await vscode.commands.executeCommand('jinsheet.undo'); await vscode.commands.executeCommand('jinsheet.undo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      await evaluate('document.querySelector("#printIncludeCharts").checked=false');
      const tableOnly = await exportPrint(3); assert.doesNotMatch(tableOnly, /data-print-chart=/); assert.match(tableOnly, /charts are not included/);
      assert.deepEqual(await fs.readFile(file), before); assert.equal(vscode.window.tabGroups.all.flatMap(group => group.tabs).some(tab => tab.isDirty), false);
    } finally {
      vscode.window.showSaveDialog = originalDialog; vscode.env.openExternal = originalOpen;
      devtools?.close(); await vscode.commands.executeCommand('workbench.action.closeAllEditors'); await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium changes chart placement and preserves fixed size through dimension edits, save and undo', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-chart-placement-host-')), file = path.join(directory, 'placement.xlsx'), output = path.join(directory, 'placement.html');
    let model = patchAddChart(analyze(await fs.readFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx'))), 0, { top: 0, bottom: 3, left: 0, right: 1 }).model;
    model = patchPrintSettings(model, 0, { ...model.sheets[0].printSettings, fitToPage: true, fitToWidth: 1, fitToHeight: 1 }).model;
    const originalBox = chartPrintGeometry(model.sheets[0], model.sheets[0].chartAnchors).boxes[0]; await fs.writeFile(file, model.bytes); const originalBytes = await fs.readFile(file);
    const originalDialog = vscode.window.showSaveDialog, originalOpen = vscode.env.openExternal, opened = [];
    vscode.window.showSaveDialog = async () => vscode.Uri.file(output); vscode.env.openExternal = async uri => { opened.push(uri.fsPath); return true; };
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor'); devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const result = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
        return result.result?.value;
      };
      const wait = async expression => { const deadline = Date.now() + 5000; while (Date.now() < deadline) { if (await evaluate(expression)) return; await new Promise(resolve => setTimeout(resolve, 25)); } assert.fail('Chart placement condition failed: ' + expression); };
      await wait('Boolean(document.querySelector("[data-chart-placement]"))');
      assert.equal(await evaluate('document.querySelector("[data-chart-placement]").value'), 'twoCell');
      await evaluate('window.__placementBefore=document.querySelector("[data-chart-placement]");window.__placementBefore.value="oneCell";document.querySelector("[data-chart-placement-apply]").click()');
      await wait('document.querySelector("[data-chart-placement]")!==window.__placementBefore&&document.querySelector("[data-chart-placement]")?.value==="oneCell"');
      await wait('Boolean(document.querySelector(\'.colhead[data-column="3"] .col-resize-handle\'))');
      await evaluate('window.__placementBefore=document.querySelector("[data-chart-placement]");document.querySelector(\'.colhead[data-column="3"] .col-resize-handle\').dispatchEvent(new KeyboardEvent("keydown",{key:"ArrowRight",bubbles:true}))');
      await wait('document.querySelector("[data-chart-placement]")!==window.__placementBefore');
      await wait('Boolean(document.querySelector("#cell-3-0"))');
      await evaluate('window.__placementBefore=document.querySelector("[data-chart-placement]");document.querySelector("#cell-3-0").click();document.querySelector("#rowHeight").value="39";document.querySelector("#applyRowHeight").click()');
      await wait('document.querySelector("[data-chart-placement]")!==window.__placementBefore');
      await vscode.commands.executeCommand('workbench.action.files.save');
      let saved = analyze(await fs.readFile(file)), box = chartPrintGeometry(saved.sheets[0], saved.sheets[0].chartAnchors).boxes[0];
      assert.equal(saved.sheets[0].chartPlacements[0], 'oneCell'); assert.notEqual(saved.sheets[0].columnWidths[3], model.sheets[0].columnWidths[3]); assert.equal(saved.sheets[0].rowHeights[3], 52);
      assert.ok(Math.abs(box.width - originalBox.width) < .001); assert.ok(Math.abs(box.height - originalBox.height) < .001);
      await evaluate('document.querySelector("#printSheet").click()'); const printDeadline = Date.now() + 7000;
      while (Date.now() < printDeadline && !opened.length) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(opened.length, 1); const pages = await measurePrintPages(evaluate, await fs.readFile(output, 'utf8'), true); assert.ok(pages[0].charts.length > 0); assert.equal(pages[0].charts[0].inside, true);
      for (let step = 0; step < 3; step++) { await vscode.commands.executeCommand('jinsheet.undo'); await vscode.commands.executeCommand('workbench.action.files.save'); }
      await wait('document.querySelector("[data-chart-placement]")?.value==="twoCell"');
      const restored = analyze(await fs.readFile(file)), original = analyze(originalBytes);
      assert.deepEqual(Object.keys(restored.files).sort(), Object.keys(original.files).sort());
      for (const [name, bytes] of Object.entries(original.files)) assert.deepEqual(restored.files[name], bytes, name);
      for (let step = 0; step < 3; step++) { await vscode.commands.executeCommand('jinsheet.redo'); await vscode.commands.executeCommand('workbench.action.files.save'); }
      saved = analyze(await fs.readFile(file)); assert.equal(saved.sheets[0].chartPlacements[0], 'oneCell'); box = chartPrintGeometry(saved.sheets[0], saved.sheets[0].chartAnchors).boxes[0];
      assert.ok(Math.abs(box.width - originalBox.width) < .001); assert.ok(Math.abs(box.height - originalBox.height) < .001);
      await wait('document.querySelector("[data-chart-placement]")?.value==="oneCell"');
      await evaluate('window.__placementBefore=document.querySelector("[data-chart-placement]");window.__placementBefore.value="absolute";document.querySelector("[data-chart-placement-apply]").click()');
      await wait('document.querySelector("[data-chart-placement]")!==window.__placementBefore&&document.querySelector("[data-chart-placement]")?.value==="absolute"');
      await wait('Boolean(document.querySelector(\'.colhead[data-column="0"] .col-resize-handle\'))');
      await evaluate('window.__placementBefore=document.querySelector("[data-chart-placement]");document.querySelector(\'.colhead[data-column="0"] .col-resize-handle\').dispatchEvent(new KeyboardEvent("keydown",{key:"ArrowRight",bubbles:true}))');
      await wait('document.querySelector("[data-chart-placement]")!==window.__placementBefore'); await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file)); assert.equal(saved.sheets[0].chartPlacements[0], 'absolute');
      const fixed = chartPrintGeometry(saved.sheets[0], saved.sheets[0].chartAnchors).boxes[0];
      for (const key of ['x', 'y', 'width', 'height']) assert.ok(Math.abs(fixed[key] - box[key]) < .001, key);
    } finally {
      vscode.window.showSaveDialog = originalDialog; vscode.env.openExternal = originalOpen;
      devtools?.close(); await vscode.commands.executeCommand('workbench.action.closeAllEditors'); await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium applies numeric and text filter conditions with save, clear, cancel and undo', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-custom-filter-host-')), file = path.join(directory, 'conditions.xlsx');
    const original = analyze(await fs.readFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx'))), files = { ...original.files };
    const rows = '<row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Score</t></is></c></row>' + [['alpha', 5], ['beta', 10], ['ALPHA', 20], ['gamma', 15], ['literal*?', 25]].map(([name, score], index) => `<row r="${index + 2}"><c r="A${index + 2}" t="inlineStr"><is><t>${name}</t></is></c><c r="B${index + 2}"><v>${score}</v></c></row>`).join('') + '<row r="7"><c r="B7"><f>D7*2</f><v>12</v></c><c r="D7"><v>6</v></c></row><row r="9"><c r="B9"><f>SUBTOTAL(9,B2:B7)</f><v>87</v></c></row>';
    files[original.sheets[0].path] = new TextEncoder().encode(original.sheets[0].xml.replace(/<sheetData>[\s\S]*?<\/sheetData>/, '<sheetData>' + rows + '</sheetData>')); await fs.writeFile(file, zipSync(files));
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor'); devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const result = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
        return result.result?.value;
      };
      const wait = async expression => { const deadline = Date.now() + 6000; while (Date.now() < deadline) { if (await evaluate(expression)) return; await new Promise(resolve => setTimeout(resolve, 25)); } assert.fail('Condition filter failed: ' + expression + ' ' + await evaluate('document.querySelector("#notice")?.textContent')); };
      const open = async column => { await wait('(()=>{const cell=document.querySelector("#cell-0-' + column + '");if(!cell)return false;cell.click();document.querySelector("#filterValuesButton").click();return true})()'); await wait('!document.querySelector("#filterPanel").hidden'); };
      await open(1);
      await evaluate('document.querySelector("#filterMode").value="conditions";document.querySelector("#filterMode").dispatchEvent(new Event("change"));document.querySelector("#filterOperator1").value="greaterThanOrEqual";document.querySelector("#filterConditionValue1").value="10";document.querySelector("#filterUseSecond").checked=true;document.querySelector("#filterUseSecond").dispatchEvent(new Event("change"));document.querySelector("#filterLogic").value="and";document.querySelector("#filterOperator2").value="lessThan";document.querySelector("#filterConditionValue2").value="20"');
      assert.equal(await evaluate('Boolean(document.querySelector("#cell-1-1"))&&Boolean(document.querySelector("#cell-3-1"))'), true, 'typing criteria does not filter or edit cells');
      await evaluate('document.querySelector("#filterConditionValue2").dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))');
      await wait('document.querySelector("#filterPanel").hidden&&!document.querySelector("#cell-1-1")&&!document.querySelector("#cell-3-1")&&document.querySelector("#cell-8-1")?.textContent==="37"');
      await vscode.commands.executeCommand('workbench.action.files.save');
      let saved = analyze(await fs.readFile(file)); assert.equal(saved.writable, true); assert.equal(saved.sheets[0].autoFilter.columns[0].custom.and, true); assert.equal(saved.sheets[0].autoFilter.columns[0].custom.conditions.length, 2); assert.deepEqual(saved.sheets[0].nativeHiddenRows, [1, 3, 5]); assert.deepEqual(saved.sheets[0].hiddenRows, []);
      await open(1); assert.equal(await evaluate('document.querySelector("#filterMode").value'), 'conditions'); assert.equal(await evaluate('document.querySelector("#filterConditionValue1").value'), '10');
      await evaluate('document.querySelector("#filterConditionValue1").value="99";document.querySelector("#filterCancel").click()'); assert.equal((await fs.readFile(file)).equals(Buffer.from(saved.bytes)), true);
      await open(0); await evaluate('document.querySelector("#filterMode").value="conditions";document.querySelector("#filterMode").dispatchEvent(new Event("change"));document.querySelector("#filterOperator1").value="contains";document.querySelector("#filterConditionValue1").value="BETA";document.querySelector("#filterApply").click()');
      await wait('document.querySelector("#filterPanel").hidden&&document.querySelector("#cell-8-1")?.textContent==="10"'); await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file)); assert.equal(saved.sheets[0].autoFilter.columns.length, 2); assert.equal(saved.sheets[0].autoFilter.columns.find(item => item.column === 0).custom.conditions[0].value, '*BETA*');
      await open(0); assert.equal(await evaluate('document.querySelector("#filterOperator1").value'), 'contains'); await evaluate('document.querySelector("#filterClear").click()');
      await wait('document.querySelector("#filterPanel").hidden&&document.querySelector("#cell-8-1")?.textContent==="37"'); await vscode.commands.executeCommand('workbench.action.files.save'); assert.equal(analyze(await fs.readFile(file)).sheets[0].autoFilter.columns.length, 1);
      await vscode.commands.executeCommand('jinsheet.undo'); await wait('document.querySelector("#cell-8-1")?.textContent==="10"'); await vscode.commands.executeCommand('workbench.action.files.save');
      await vscode.commands.executeCommand('jinsheet.redo'); await wait('document.querySelector("#cell-8-1")?.textContent==="37"'); await vscode.commands.executeCommand('workbench.action.files.save');
      devtools.close(); devtools = null; await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor'); devtools = await waitForSpreadsheetWebview();
      await wait('document.querySelector("#cell-8-1")?.textContent==="37"&&!document.querySelector("#cell-1-1")');
      await open(1); await evaluate('document.querySelector("#filterClear").click()');
      await wait('document.querySelector("#filterPanel").hidden&&document.querySelector("#cell-8-1")?.textContent==="87"&&Boolean(document.querySelector("#cell-1-1"))&&Boolean(document.querySelector("#cell-3-1"))');
      await vscode.commands.executeCommand('workbench.action.files.save'); saved = analyze(await fs.readFile(file));
      assert.deepEqual(saved.sheets[0].nativeHiddenRows, []); assert.deepEqual(saved.sheets[0].hiddenRows, []);

    } finally { devtools?.close(); await vscode.commands.executeCommand('workbench.action.closeAllEditors'); await fs.rm(directory, { recursive: true, force: true }); }
  });

  test('Chromium expands table boundaries atomically and preserves the opt-out through input, paste, save and reopen', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-expansion-host-')), file = path.join(directory, 'expansion.xlsx');
    const sample = analyze(await fs.readFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx'))), files = { ...sample.files }, sheet = sample.sheets[0];
    const rows = '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c><c r="C1" t="inlineStr"><is><t>Result</t></is></c></row>' + [10,20,30].map((value,index)=>'<row r="'+(index+2)+'"><c r="B'+(index+2)+'"><v>'+value+'</v></c></row>').join('') + '</sheetData>';
    files[sheet.path] = new TextEncoder().encode(sheet.xml.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ').replace(/<sheetData>[\s\S]*?<\/sheetData>/, rows).replace('</worksheet>','<tableParts count="1"><tablePart r:id="t1"/></tableParts></worksheet>'));
    files[path.posix.join(path.posix.dirname(sheet.path),'_rels',path.posix.basename(sheet.path)+'.rels')] = new TextEncoder().encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="t1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>');
    files['xl/tables/table1.xml'] = new TextEncoder().encode('<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:C4"><autoFilter ref="A1:C4"/><tableColumns count="3"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Amount"/><tableColumn id="3" name="Result"/></tableColumns></table>');
    files['[Content_Types].xml'] = new TextEncoder().encode(new TextDecoder().decode(files['[Content_Types].xml']).replace('</Types>','<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>')); await fs.writeFile(file,zipSync(files)); let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith',vscode.Uri.file(file),'jinsheet.editor'); devtools=await waitForSpreadsheetWebview();
      const evaluate=async expression=>{const result=await devtools.send('Runtime.evaluate',{expression,contextId:devtools.contextId,returnByValue:true,awaitPromise:true});if(result.exceptionDetails)throw new Error(result.exceptionDetails.exception?.description||result.exceptionDetails.text);return result.result?.value;};
      const wait=async expression=>{const deadline=Date.now()+6000;while(Date.now()<deadline){if(await evaluate(expression))return;await new Promise(resolve=>setTimeout(resolve,25));}assert.fail('Table expansion failed: '+expression+' '+await evaluate('document.querySelector("#notice").textContent'));};
      const act=async(row,column,script='')=>wait('(()=>{const cell=document.querySelector('+JSON.stringify('#cell-'+row+'-'+column)+');if(!cell)return false;const rect=cell.getBoundingClientRect(),event={bubbles:true,button:0,clientX:rect.left+4,clientY:rect.top+4};cell.dispatchEvent(new MouseEvent("mousedown",event));cell.dispatchEvent(new MouseEvent("mouseup",event));cell.dispatchEvent(new MouseEvent("click",event));if(!cell.classList.contains("selected")||document.querySelector("#address").value!=='+JSON.stringify(String.fromCharCode(65+column)+(row+1))+')return false;'+script+';return true})()');
      const enter=async(row,column,value)=>act(row,column,'const input=document.querySelector("#formula");input.value='+JSON.stringify(value)+';input.dispatchEvent(new Event("input",{bubbles:true}));input.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))');
      const save=async()=>{await vscode.commands.executeCommand('workbench.action.files.save');return analyze(await fs.readFile(file));};
      const saveUntil=async predicate=>{const end=Date.now()+6000;do{const workbook=await save();if(predicate(workbook))return workbook;await new Promise(resolve=>setTimeout(resolve,25));}while(Date.now()<end);assert.fail('Table expansion save did not reach the expected state');};
      await wait('document.querySelector("#autoExpandTables")?.checked&&!document.querySelector("#autoExpandTables").disabled');
      await enter(1,2,'=B2*2');await wait('document.querySelector("#cell-3-2")?.textContent==="60"');await save();
      await enter(4,1,'7');await wait('document.querySelector("#cell-4-2")?.textContent==="14"');
      let saved=await save();assert.equal(saved.sheets[0].tables[0].range.bottom,4);assert.equal(saved.sheets[0].cells.get('C5').formula,'B5*2');
      await vscode.commands.executeCommand('jinsheet.undo');await wait('document.querySelector("#cell-4-1")?.textContent===""&&document.querySelector("#cell-4-2")?.textContent===""');saved=await save();assert.equal(saved.sheets[0].tables[0].range.bottom,3);
      await vscode.commands.executeCommand('jinsheet.redo');await wait('document.querySelector("#cell-4-2")?.textContent==="14"');
      await evaluate('document.querySelector("#autoExpandTables").checked=false');await enter(5,1,'8');await wait('document.querySelector("#cell-5-1")?.textContent==="8"&&document.querySelector("#cell-5-2")?.textContent===""');saved=await save();assert.equal(saved.sheets[0].tables[0].range.bottom,4);
      await evaluate('document.querySelector("#autoExpandTables").checked=true');await enter(5,1,'9');await wait('document.querySelector("#cell-5-2")?.textContent==="18"');saved=await save();const expansionDeadline=Date.now()+3000;while(saved.sheets[0].tables[0].range.bottom!==5&&Date.now()<expansionDeadline){await new Promise(resolve=>setTimeout(resolve,50));saved=await save()}assert.equal(saved.sheets[0].tables[0].range.bottom,5);
      await enter(0,3,'Amount');await wait('document.querySelector("#cell-0-3")?.textContent==="Amount2"');saved=await save();assert.equal(saved.sheets[0].tables[0].range.right,3);
      await vscode.commands.executeCommand('jinsheet.undo');await wait('document.querySelector("#cell-0-3")?.textContent===""');saved=await save();assert.equal(saved.sheets[0].tables[0].range.right,2);
      await vscode.commands.executeCommand('jinsheet.redo');await wait('document.querySelector("#cell-0-3")?.textContent==="Amount2"');
      await enter(1,3,'=B2+1');await wait('document.querySelector("#cell-5-3")?.textContent==="10"');
      await vscode.env.clipboard.writeText('11\n12');await act(6,1,'document.querySelector("#grid").dispatchEvent(new KeyboardEvent("keydown",{key:"v",ctrlKey:true,bubbles:true}))');
      await wait('document.querySelector("#cell-6-2")?.textContent==="22"&&document.querySelector("#cell-7-2")?.textContent==="24"&&document.querySelector("#cell-7-3")?.textContent==="13"');
      saved=await save();assert.equal(saved.sheets[0].tables[0].range.bottom,7);assert.equal(saved.sheets[0].tables[0].autoFilter.bottom,7);assert.equal(saved.sheets[0].cells.get('B5').value,'7');assert.equal(saved.sheets[0].cells.get('B6').value,'9');assert.equal(saved.sheets[0].cells.get('C8').formula,'B8*2');assert.equal(saved.sheets[0].tables[0].calculatedColumns[3],'B2+1');
      await vscode.commands.executeCommand('jinsheet.undo');await wait('document.querySelector("#cell-6-1")?.textContent===""&&document.querySelector("#cell-7-2")?.textContent===""');saved=await save();assert.equal(saved.sheets[0].tables[0].range.bottom,5);
      await vscode.commands.executeCommand('jinsheet.redo');await wait('document.querySelector("#cell-7-3")?.textContent==="13"');await saveUntil(workbook=>workbook.sheets[0].tables[0].range.bottom===7&&workbook.sheets[0].tables[0].range.right===3&&workbook.sheets[0].cells.get('C8')?.formula==='B8*2');
      devtools.close();devtools=null;await vscode.commands.executeCommand('workbench.action.closeAllEditors');await vscode.commands.executeCommand('vscode.openWith',vscode.Uri.file(file),'jinsheet.editor');devtools=await waitForSpreadsheetWebview();await wait('document.querySelector("#cell-7-2")?.textContent==="24"&&document.querySelector("#cell-0-3")?.textContent==="Amount2"');
    } finally {devtools?.close();await vscode.commands.executeCommand('workbench.action.closeAllEditors');await fs.rm(directory,{recursive:true,force:true});}
  });

  test('Chromium renames table headers and structured formulas through save, reopen and one undo', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-header-host-')), file = path.join(directory, 'headers.xlsx');
    const sample = analyze(await fs.readFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx'))), files = { ...sample.files }, sheet = sample.sheets[0];
    const rows = '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c><c r="C1" t="inlineStr"><is><t>Result</t></is></c></row>' + [10,20,30].map((value,index)=>'<row r="'+(index+2)+'"><c r="B'+(index+2)+'"><v>'+value+'</v></c></row>').join('') + '<row r="6"><c r="D6"><f>SUBTOTAL(9,Sales[Result])</f><v>0</v></c></row></sheetData>';
    files[sheet.path] = new TextEncoder().encode(sheet.xml.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ').replace(/<sheetData>[\s\S]*?<\/sheetData>/, rows).replace('</worksheet>','<tableParts count="1"><tablePart r:id="t1"/></tableParts></worksheet>'));
    files[path.posix.join(path.posix.dirname(sheet.path),'_rels',path.posix.basename(sheet.path)+'.rels')] = new TextEncoder().encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="t1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>');
    files['xl/tables/table1.xml'] = new TextEncoder().encode('<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:C4"><autoFilter ref="A1:C4"/><tableColumns count="3"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Amount"/><tableColumn id="3" name="Result"/></tableColumns></table>');
    files['[Content_Types].xml'] = new TextEncoder().encode(new TextDecoder().decode(files['[Content_Types].xml']).replace('</Types>','<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>')); await fs.writeFile(file,zipSync(files)); let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith',vscode.Uri.file(file),'jinsheet.editor');devtools=await waitForSpreadsheetWebview();
      const evaluate=async expression=>{const result=await devtools.send('Runtime.evaluate',{expression,contextId:devtools.contextId,returnByValue:true,awaitPromise:true});if(result.exceptionDetails)throw new Error(result.exceptionDetails.exception?.description||result.exceptionDetails.text);return result.result?.value;};
      const wait=async expression=>{const deadline=Date.now()+6000;while(Date.now()<deadline){if(await evaluate(expression))return;await new Promise(resolve=>setTimeout(resolve,25));}assert.fail('Table header condition failed: '+expression+' '+await evaluate('document.querySelector("#notice").textContent'));};
      const enter=async(row,column,value)=>wait('(()=>{const cell=document.querySelector('+JSON.stringify('#cell-'+row+'-'+column)+');if(!cell)return false;cell.click();const input=document.querySelector("#formula");input.value='+JSON.stringify(value)+';input.dispatchEvent(new Event("input",{bubbles:true}));input.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}));return true})()');
      const save=async()=>{await vscode.commands.executeCommand('workbench.action.files.save');return analyze(await fs.readFile(file));};
      await enter(1,2,'=[@Amount]*2');await wait('document.querySelector("#cell-3-2")?.textContent==="60"');await save();
      await enter(0,1,'가격');await wait('document.querySelector("#cell-0-1")?.textContent==="가격"&&document.querySelector("#cell-3-2")?.textContent==="60"');
      let saved=await save();assert.equal(saved.sheets[0].tables[0].columns[1],'가격');assert.equal(saved.sheets[0].cells.get('C3').formula,'[@가격]*2');assert.equal(saved.sheets[0].tables[0].calculatedColumns[2],'[@가격]*2');
      await vscode.commands.executeCommand('jinsheet.undo');await wait('document.querySelector("#cell-0-1")?.textContent==="Amount"');saved=await save();assert.equal(saved.sheets[0].cells.get('C3').formula,'[@Amount]*2');
      await vscode.commands.executeCommand('jinsheet.redo');await wait('document.querySelector("#cell-0-1")?.textContent==="가격"');await enter(0,2,'Total');await wait('document.querySelector("#cell-0-2")?.textContent==="Total"&&document.querySelector("#cell-5-3")?.textContent==="120"');saved=await save();assert.equal(saved.sheets[0].cells.get('D6').formula,'SUBTOTAL(9,Sales[Total])');
      await enter(0,0,'가격');await wait('document.querySelector("#cell-0-0")?.textContent==="가격2"');const headerSaveDeadline=Date.now()+8000;do{saved=await save();if(JSON.stringify(saved.sheets[0].tables[0].columns)===JSON.stringify(['가격2','가격','Total']))break;await new Promise(resolve=>setTimeout(resolve,25))}while(Date.now()<headerSaveDeadline);assert.deepEqual(saved.sheets[0].tables[0].columns,['가격2','가격','Total']);
      devtools.close();devtools=null;await vscode.commands.executeCommand('workbench.action.closeAllEditors');await vscode.commands.executeCommand('vscode.openWith',vscode.Uri.file(file),'jinsheet.editor');devtools=await waitForSpreadsheetWebview();await wait('document.querySelector("#cell-0-0")?.textContent==="가격2"&&document.querySelector("#cell-3-2")?.textContent==="60"&&document.querySelector("#cell-5-3")?.textContent==="120"');
    } finally {devtools?.close();await vscode.commands.executeCommand('workbench.action.closeAllEditors');await fs.rm(directory,{recursive:true,force:true});}
  });

  test('Chromium creates a named table from a selection with retry, undo, filters and reopen',async()=>{
    const directory=await fs.mkdtemp(path.join(os.tmpdir(),'jinsheet-create-table-host-')),file=path.join(directory,'new-table.xlsx'),encode=text=>new TextEncoder().encode(text);
    const sample=analyze(await fs.readFile(path.join(__dirname,'..','fixtures','JinSheet-Styled-Sample.xlsx'))),files={...sample.files},sheet=sample.sheets[0];
    const rows='<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c><c r="C1" t="inlineStr"><is><t>Result</t></is></c></row><row r="2"><c r="B2"><v>7</v></c></row><row r="3"><c r="B3"><v>9</v></c></row><row r="6"><c r="D6"><f>SUBTOTAL(9,C2:C3)</f><v>0</v></c></row></sheetData>';
    files[sheet.path]=encode(sheet.xml.replace(/<sheetData>[\s\S]*?<\/sheetData>/,rows));await fs.writeFile(file,zipSync(files));let devtools;
    try{
      await vscode.commands.executeCommand('vscode.openWith',vscode.Uri.file(file),'jinsheet.editor');devtools=await waitForSpreadsheetWebview();
      const evaluate=async expression=>{const result=await devtools.send('Runtime.evaluate',{expression,contextId:devtools.contextId,returnByValue:true,awaitPromise:true});if(result.exceptionDetails)throw Error(result.exceptionDetails.exception?.description||result.exceptionDetails.text);return result.result?.value};
      const wait=async expression=>{const deadline=Date.now()+6000;while(Date.now()<deadline){if(await evaluate(expression))return;await new Promise(resolve=>setTimeout(resolve,25))}assert.fail('Create table: '+expression+' '+JSON.stringify(await evaluate('({notice:document.querySelector("#notice").textContent,address:document.querySelector("#address").value,formula:document.querySelector("#formula").value,cells:[1,2,5].map(row=>[1,2,3].map(column=>{const cell=document.querySelector("#cell-"+row+"-"+column);return cell?{text:cell.textContent,data:cell.__jinSheetData}:null})),selected:document.querySelector(".cell.selected")?.id})')))};
      const act=async(row,column,script='')=>wait('(()=>{const cell=document.querySelector('+JSON.stringify('#cell-'+row+'-'+column)+');if(!cell)return false;cell.click();'+script+';return true})()');
      const save=async()=>{await vscode.commands.executeCommand('workbench.action.files.save');return analyze(await fs.readFile(file))};
      await wait('Boolean(document.querySelector("#cell-2-2"))');await evaluate('document.querySelector("#cell-0-0").dispatchEvent(new MouseEvent("mousedown",{bubbles:true,button:0}));document.querySelector("#cell-2-2").dispatchEvent(new MouseEvent("mouseenter"));document.dispatchEvent(new MouseEvent("mouseup",{bubbles:true}));document.querySelector("#newTableName").value="A1";document.querySelector("#createTable").click()');await wait('!document.querySelector("#createTable").disabled&&document.querySelector("#notice").textContent.includes("name")');assert.equal(await evaluate('document.querySelector("#newTableName").value'),'A1');assert.equal((await save()).sheets[0].tables.length,0);
      await evaluate('document.querySelector("#newTableName").value="Sales";document.querySelector("#createTable").click()');await wait('document.querySelector("#newTableName").value===""&&document.querySelector("#notice").textContent.includes("Sales")');let saved=await save();assert.equal(saved.sheets[0].tables[0].name,'Sales');assert.deepEqual(saved.sheets[0].tables[0].range,{top:0,bottom:2,left:0,right:2});assert.deepEqual(saved.sheets[0].tables[0].columns,['Name','Amount','Result']);
      await vscode.commands.executeCommand('jinsheet.undo');await wait('!document.querySelector("#createTable").disabled');assert.equal((await save()).sheets[0].tables.length,0);await vscode.commands.executeCommand('jinsheet.redo');await wait('document.querySelector("#createTable").disabled');assert.equal((await save()).sheets[0].tables[0].name,'Sales');
      await wait('document.querySelector("#applyTableStyle")&&!document.querySelector("#applyTableStyle").disabled');const previousStyleNotice=await evaluate('document.querySelector("#notice").textContent');await evaluate('(()=>{const select=document.querySelector("#tableStyleName");select.value="TableStyleDark3";select.dispatchEvent(new Event("change",{bubbles:true}));const rows=document.querySelector("#tableStyleRows");rows.checked=false;rows.dispatchEvent(new Event("change",{bubbles:true}));const columns=document.querySelector("#tableStyleColumns");columns.checked=true;columns.dispatchEvent(new Event("change",{bubbles:true}));document.querySelector("#applyTableStyle").click()})()');await wait('document.querySelector("#notice").textContent!=='+JSON.stringify(previousStyleNotice)+'&&document.querySelector("#tableStyleName").value==="TableStyleDark3"&&!document.querySelector("#tableStyleRows").checked&&document.querySelector("#tableStyleColumns").checked&&!document.querySelector("#applyTableStyle").disabled');saved=await save();const styleDeadline=Date.now()+8000;while(saved.sheets[0].tables[0].styleInfo?.name!=='TableStyleDark3'&&Date.now()<styleDeadline){await new Promise(resolve=>setTimeout(resolve,50));await vscode.commands.executeCommand('workbench.action.files.save');saved=analyze(await fs.readFile(file));}assert.deepEqual(saved.sheets[0].tables[0].styleInfo,{name:'TableStyleDark3',showFirstColumn:false,showLastColumn:false,showRowStripes:false,showColumnStripes:true});
      await wait('document.querySelector("#address").value==="C3"');await evaluate('(()=>{const input=document.querySelector("#formula");input.value="=B3*2";input.dispatchEvent(new Event("input",{bubbles:true}));input.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))})()');await wait('document.querySelector("#cell-1-2")?.textContent==="14"&&document.querySelector("#cell-2-2")?.textContent==="18"&&document.querySelector("#cell-5-3")?.textContent==="32"');
      await act(0,1,'document.querySelector("#filterValuesButton").click()');await wait('!document.querySelector("#filterPanel").hidden');await evaluate('document.querySelector("#filterMode").value="top10";document.querySelector("#filterMode").dispatchEvent(new Event("change"));document.querySelector("#filterTopValue").value="1";document.querySelector("#filterApply").click()');await wait('document.querySelector("#filterPanel").hidden&&document.querySelector("#cell-5-3")?.textContent==="18"&&!document.querySelector("#cell-1-1")');saved=await save();assert.equal(saved.sheets[0].tables[0].calculatedColumns[2],'B2*2');assert.equal(saved.sheets[0].tables[0].autoFilter.columns[0].top10.value,1);
      devtools.close();devtools=null;await vscode.commands.executeCommand('workbench.action.closeAllEditors');await vscode.commands.executeCommand('vscode.openWith',vscode.Uri.file(file),'jinsheet.editor');devtools=await waitForSpreadsheetWebview();await wait('document.querySelector("#cell-5-3")?.textContent==="18"&&document.querySelector("#cell-0-0")&&getComputedStyle(document.querySelector("#cell-0-0")).backgroundColor==="rgb(64, 64, 64)"');const reopened=analyze(await fs.readFile(file));assert.equal(reopened.sheets[0].tables[0].name,'Sales');assert.deepEqual(reopened.sheets[0].tables[0].styleInfo,{name:'TableStyleDark3',showFirstColumn:false,showLastColumn:false,showRowStripes:false,showColumnStripes:true});
    }finally{devtools?.close();await vscode.commands.executeCommand('workbench.action.closeAllEditors');await fs.rm(directory,{recursive:true,force:true})}
  });

  test('Chromium opens header-only tables and deletes and restores their last data row with totals, undo and reopen', async()=>{
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-empty-table-host-')), file = path.join(directory, 'tab.xlsx');
    const sample = analyze(await fs.readFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx'))), files = { ...sample.files }, sheet = sample.sheets[0];
    const rows = '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c><c r="C1" t="inlineStr"><is><t>Result</t></is></c></row>' + [].map((value,index)=>'<row r="'+(index+2)+'"><c r="B'+(index+2)+'"><v>'+value+'</v></c></row>').join('') + '<row r="6"><c r="D6"><f>SUBTOTAL(9,Sales[Result])</f><v>0</v></c></row></sheetData>';
    files[sheet.path] = new TextEncoder().encode(sheet.xml.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ').replace(/<sheetData>[\s\S]*?<\/sheetData>/, rows).replace('</worksheet>','<tableParts count="1"><tablePart r:id="t1"/></tableParts></worksheet>'));
    files[path.posix.join(path.posix.dirname(sheet.path),'_rels',path.posix.basename(sheet.path)+'.rels')] = new TextEncoder().encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="t1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>');
    files['xl/tables/table1.xml'] = new TextEncoder().encode('<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:C1"><autoFilter ref="A1:C1"/><tableColumns count="3"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Amount"/><tableColumn id="3" name="Result"><calculatedColumnFormula>B2*2</calculatedColumnFormula></tableColumn></tableColumns></table>');
    files['[Content_Types].xml']=new TextEncoder().encode(new TextDecoder().decode(files['[Content_Types].xml']).replace('</Types>','<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>'));await fs.writeFile(file,zipSync(files));let devtools;
    try{
      await vscode.commands.executeCommand('vscode.openWith',vscode.Uri.file(file),'jinsheet.editor');devtools=await waitForSpreadsheetWebview();
      const evaluate=async expression=>{const result=await devtools.send('Runtime.evaluate',{expression,contextId:devtools.contextId,returnByValue:true,awaitPromise:true});if(result.exceptionDetails)throw Error(result.exceptionDetails.exception?.description||result.exceptionDetails.text);return result.result?.value};
      const wait=async expression=>{const end=Date.now()+6000;while(Date.now()<end){if(await evaluate(expression))return;await new Promise(resolve=>setTimeout(resolve,25))}assert.fail('Empty table: '+expression+' '+await evaluate('document.querySelector("#notice").textContent'))};
      const act=async(row,column,script='')=>wait('(()=>{const cell=document.querySelector('+JSON.stringify('#cell-'+row+'-'+column)+');if(!cell)return false;cell.click();'+script+';return true})()');
      const enter=async(row,column,value)=>act(row,column,'const input=document.querySelector("#formula");input.value='+JSON.stringify(value)+';input.dispatchEvent(new Event("input",{bubbles:true}));input.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))');
      const save=async()=>{await vscode.commands.executeCommand('workbench.action.files.save');return analyze(await fs.readFile(file))};
      const saveUntil=async predicate=>{const end=Date.now()+6000;do{const workbook=await save();if(predicate(workbook))return workbook;await new Promise(resolve=>setTimeout(resolve,25))}while(Date.now()<end);assert.fail('Empty table save did not reach the expected state')};
      await wait('document.querySelector("#cell-5-3")?.textContent==="0"');await act(0,1,'document.querySelector("#filterValuesButton").click()');await wait('!document.querySelector("#filterPanel").hidden');await evaluate('document.querySelector("#filterCancel").click()');
      await enter(1,1,'7');await wait('document.querySelector("#cell-1-2")?.textContent==="14"&&document.querySelector("#cell-5-3")?.textContent==="14"');let saved=await save();assert.equal(saved.sheets[0].tables[0].range.bottom,1);assert.equal(saved.sheets[0].cells.get('C2').formula,'B2*2');await vscode.commands.executeCommand('jinsheet.undo');await wait('document.querySelector("#cell-5-3")?.textContent==="0"&&document.querySelector("#cell-1-2")?.textContent===""');saved=await save();assert.equal(saved.sheets[0].tables[0].range.bottom,0);
      await act(0,2,'document.querySelector("#tableTotals").click()');await wait('document.querySelector("#cell-1-2")?.textContent==="0"');await act(0,2,'document.querySelector("#grid").dispatchEvent(new KeyboardEvent("keydown",{key:"Tab",bubbles:true}))');await wait('document.querySelector("#address").value==="A2"&&document.querySelector("#cell-2-2")?.textContent==="0"');await enter(1,1,'9');await wait('document.querySelector("#cell-1-2")?.textContent==="18"&&document.querySelector("#cell-2-2")?.textContent==="18"');saved=await save();assert.equal(saved.sheets[0].tables[0].calculatedColumns[2],'B2*2');assert.equal(saved.sheets[0].tables[0].totalsRowCount,1);assert.equal(saved.sheets[0].tables[0].range.bottom,2);
      await act(1,1,'document.querySelector("#deleteRow").click()');await wait('document.querySelector("#cell-1-2")?.textContent==="0"');saved=await saveUntil(workbook=>workbook.sheets[0].tables[0].range.bottom===1&&workbook.sheets[0].tables[0].autoFilter.bottom===0);assert.equal(saved.sheets[0].tables[0].calculatedColumns[2],'B2*2');await vscode.commands.executeCommand('jinsheet.undo');await wait('document.querySelector("#cell-1-2")?.textContent==="18"&&document.querySelector("#cell-2-2")?.textContent==="18"');saved=await saveUntil(workbook=>workbook.sheets[0].tables[0].range.bottom===2);await vscode.commands.executeCommand('jinsheet.redo');await wait('document.querySelector("#cell-1-2")?.textContent==="0"');saved=await saveUntil(workbook=>workbook.sheets[0].tables[0].range.bottom===1);await act(0,2,'document.querySelector("#grid").dispatchEvent(new KeyboardEvent("keydown",{key:"Tab",bubbles:true}))');await wait('document.querySelector("#address").value==="A2"&&document.querySelector("#cell-2-2")?.textContent==="0"');await enter(1,1,'9');await wait('document.querySelector("#cell-1-2")?.textContent==="18"&&document.querySelector("#cell-2-2")?.textContent==="18"');saved=await saveUntil(workbook=>workbook.sheets[0].tables[0].range.bottom===2&&workbook.sheets[0].cells.get('C2')?.formula==='B2*2');
      devtools.close();devtools=null;await vscode.commands.executeCommand('workbench.action.closeAllEditors');await vscode.commands.executeCommand('vscode.openWith',vscode.Uri.file(file),'jinsheet.editor');devtools=await waitForSpreadsheetWebview();await wait('document.querySelector("#cell-1-2")?.textContent==="18"&&document.querySelector("#cell-2-2")?.textContent==="18"');
    }finally{devtools?.close();await vscode.commands.executeCommand('workbench.action.closeAllEditors');await fs.rm(directory,{recursive:true,force:true})}
  });

  for(const filtered of [false,true]) test('Chromium appends table rows with Tab '+(filtered?'after filter choice and cancellation':'after editing')+' and moves totals through save and undo', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-table-tab-host-')), file = path.join(directory, 'tab.xlsx');
    const sample = analyze(await fs.readFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx'))), files = { ...sample.files }, sheet = sample.sheets[0];
    const rows = '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c><c r="C1" t="inlineStr"><is><t>Result</t></is></c></row>' + [10,20,30].map((value,index)=>'<row r="'+(index+2)+'"><c r="B'+(index+2)+'"><v>'+value+'</v></c></row>').join('') + '<row r="6"><c r="D6"><f>SUBTOTAL(9,Sales[Result])</f><v>0</v></c></row></sheetData>';
    files[sheet.path] = new TextEncoder().encode(sheet.xml.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ').replace(/<sheetData>[\s\S]*?<\/sheetData>/, rows).replace('</worksheet>','<tableParts count="1"><tablePart r:id="t1"/></tableParts></worksheet>'));
    files[path.posix.join(path.posix.dirname(sheet.path),'_rels',path.posix.basename(sheet.path)+'.rels')] = new TextEncoder().encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="t1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>');
    files['xl/tables/table1.xml'] = new TextEncoder().encode('<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:C4"><autoFilter ref="A1:C4"/><tableColumns count="3"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Amount"/><tableColumn id="3" name="Result"/></tableColumns></table>');
    files['[Content_Types].xml'] = new TextEncoder().encode(new TextDecoder().decode(files['[Content_Types].xml']).replace('</Types>','<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>')); if(filtered)files['xl/tables/table1.xml']=new TextEncoder().encode(new TextDecoder().decode(files['xl/tables/table1.xml']).replace('<autoFilter ref="A1:C4"/>','<autoFilter ref="A1:C4"><filterColumn colId="1"><customFilters><customFilter operator="greaterThan" val="0"/></customFilters></filterColumn></autoFilter>'));
    await fs.writeFile(file,zipSync(files)); let devtools;
    const originalWarning=vscode.window.showWarningMessage,prompts=[];let acceptFilterChoice=false;
    if(filtered)vscode.window.showWarningMessage=async(message,...actions)=>{prompts.push({message,actions});return acceptFilterChoice?actions[0]:actions[1]};
    try {
      await vscode.commands.executeCommand('vscode.openWith',vscode.Uri.file(file),'jinsheet.editor');devtools=await waitForSpreadsheetWebview();
      const evaluate=async expression=>{const result=await devtools.send('Runtime.evaluate',{expression,contextId:devtools.contextId,returnByValue:true,awaitPromise:true});if(result.exceptionDetails)throw new Error(result.exceptionDetails.exception?.description||result.exceptionDetails.text);return result.result?.value;};
      const wait=async expression=>{const deadline=Date.now()+6000;while(Date.now()<deadline){if(await evaluate(expression))return;await new Promise(resolve=>setTimeout(resolve,25));}assert.fail('Table Tab failed: '+expression+' '+await evaluate('document.querySelector("#notice").textContent'));};
      const act=async(row,column,script='')=>wait('(()=>{const cell=document.querySelector('+JSON.stringify('#cell-'+row+'-'+column)+');if(!cell)return false;cell.click();'+script+';return true})()');
      const enter=async(row,column,value,key='Enter')=>act(row,column,'const input=document.querySelector("#formula");input.value='+JSON.stringify(value)+';input.dispatchEvent(new Event("input",{bubbles:true}));input.dispatchEvent(new KeyboardEvent("keydown",{key:'+JSON.stringify(key)+',bubbles:true}))');
      const tab=async(row,column,backwards=false)=>act(row,column,'document.querySelector("#grid").dispatchEvent(new KeyboardEvent("keydown",{key:"Tab",shiftKey:'+backwards+',bubbles:true}))');
      const save=async()=>{await vscode.commands.executeCommand('workbench.action.files.save');return analyze(await fs.readFile(file));};
      await enter(1,2,'=B2*2');await wait('document.querySelector("#cell-3-2")?.textContent==="60"');await save();
      await tab(1,2);await wait('document.querySelector("#address").value==="A3"');await tab(2,0,true);await wait('document.querySelector("#address").value==="C2"');
      if(filtered){
        await enter(3,2,'=B4*3','Tab');await wait('document.querySelector("#notice").textContent.includes("cancelled")');assert.equal(prompts.length,1);assert.equal(await evaluate('document.querySelector("#formula").value'),'=B4*3');assert.equal(await evaluate('document.querySelector("#address").value'),'C4');const cancelled=await save();assert.equal(cancelled.sheets[0].tables[0].range.bottom,3);assert.equal(cancelled.sheets[0].cells.get('C4').formula,'B4*2');acceptFilterChoice=true;
      }
      if(filtered)await evaluate('document.querySelector("#formula").dispatchEvent(new KeyboardEvent("keydown",{key:"Tab",bubbles:true}))');else await enter(3,2,'=B4*3','Tab');await wait('document.querySelector("#address").value==="A5"&&document.querySelector("#cell-4-2")?.textContent==="0"&&document.querySelector("#cell-3-2")?.textContent==="90"');let saved=await save();assert.equal(saved.sheets[0].tables[0].range.bottom,4);assert.equal(saved.sheets[0].cells.get('C5').formula,'B5*3');if(filtered){assert.equal(prompts.length,2);assert.match(prompts[1].message,/Sales/);assert.match(prompts[1].actions[0],/Clear table filters/);assert.equal(saved.sheets[0].tables[0].autoFilter.columns.length,0);}
      await vscode.commands.executeCommand('jinsheet.undo');await wait('document.querySelector("#cell-3-2")?.textContent==="60"&&document.querySelector("#cell-4-2")?.textContent===""');saved=await save();assert.equal(saved.sheets[0].tables[0].range.bottom,3);if(filtered)assert.equal(saved.sheets[0].tables[0].autoFilter.columns.length,1);
      await vscode.commands.executeCommand('jinsheet.redo');await wait('document.querySelector("#cell-4-2")?.textContent==="0"');await enter(4,1,'7');await wait('document.querySelector("#cell-4-2")?.textContent==="21"');
      await act(4,2,'document.querySelector("#tableTotals").click()');await wait('document.querySelector("#cell-5-2")?.textContent==="201"');await save();await tab(4,2);await wait('document.querySelector("#address").value==="A6"&&document.querySelector("#cell-5-2")?.textContent==="0"&&document.querySelector("#cell-6-2")?.textContent==="201"');
      const expandedTotalsDeadline=Date.now()+5000;do{saved=await save();if(saved.sheets[0].tables[0].range.bottom===6&&saved.sheets[0].tables[0].autoFilter.bottom===5&&saved.sheets[0].cells.get('C7')?.formula==='SUBTOTAL(109,[[#Data],[Result]])')break;await new Promise(resolve=>setTimeout(resolve,25))}while(Date.now()<expandedTotalsDeadline);
      assert.equal(saved.sheets[0].tables[0].range.bottom,6);assert.equal(saved.sheets[0].tables[0].autoFilter.bottom,5);assert.equal(saved.sheets[0].cells.get('C7').formula,'SUBTOTAL(109,[[#Data],[Result]])');
      await vscode.commands.executeCommand('jinsheet.undo');await wait('document.querySelector("#cell-5-2")?.textContent==="201"');await vscode.commands.executeCommand('jinsheet.redo');await wait('document.querySelector("#cell-6-2")?.textContent==="201"');await save();
      devtools.close();devtools=null;await vscode.commands.executeCommand('workbench.action.closeAllEditors');await vscode.commands.executeCommand('vscode.openWith',vscode.Uri.file(file),'jinsheet.editor');devtools=await waitForSpreadsheetWebview();await wait('document.querySelector("#cell-6-2")?.textContent==="201"&&document.querySelector("#cell-5-2")?.textContent==="0"');
    } finally {vscode.window.showWarningMessage=originalWarning;devtools?.close();await vscode.commands.executeCommand('workbench.action.closeAllEditors');await fs.rm(directory,{recursive:true,force:true});}
  });

  test('Chromium applies top and bottom percent filters with ties, validation, save, reopen and undo', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-rank-filter-host-')), file = path.join(directory, 'rank.xlsx');
    const sample = analyze(await fs.readFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx'))), files = { ...sample.files }, sheet = sample.sheets[0];
    const rows = '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c><c r="C1" t="inlineStr"><is><t>Result</t></is></c></row>' + [10,20,30].map((value,index)=>'<row r="'+(index+2)+'"><c r="B'+(index+2)+'"><v>'+value+'</v></c></row>').join('') + '<row r="6"><c r="D6"><f>SUBTOTAL(9,Sales[Result])</f><v>0</v></c></row></sheetData>';
    files[sheet.path] = new TextEncoder().encode(sheet.xml.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ').replace(/<sheetData>[\s\S]*?<\/sheetData>/, rows).replace('</worksheet>','<tableParts count="1"><tablePart r:id="t1"/></tableParts></worksheet>'));
    files[path.posix.join(path.posix.dirname(sheet.path),'_rels',path.posix.basename(sheet.path)+'.rels')] = new TextEncoder().encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="t1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>');
    files['xl/tables/table1.xml'] = new TextEncoder().encode('<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:C4"><autoFilter ref="A1:C4"/><tableColumns count="3"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Amount"/><tableColumn id="3" name="Result"/></tableColumns></table>');
    files['[Content_Types].xml'] = new TextEncoder().encode(new TextDecoder().decode(files['[Content_Types].xml']).replace('</Types>','<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>')); await fs.writeFile(file,zipSync(files)); let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith',vscode.Uri.file(file),'jinsheet.editor');devtools=await waitForSpreadsheetWebview();
      const evaluate=async expression=>{const result=await devtools.send('Runtime.evaluate',{expression,contextId:devtools.contextId,returnByValue:true,awaitPromise:true});if(result.exceptionDetails)throw new Error(result.exceptionDetails.exception?.description||result.exceptionDetails.text);return result.result?.value;};
      const wait=async expression=>{const deadline=Date.now()+6000;while(Date.now()<deadline){if(await evaluate(expression))return;await new Promise(resolve=>setTimeout(resolve,25));}assert.fail('Rank filter failed: '+expression+' '+await evaluate('document.querySelector("#notice").textContent'));};
      const act=async(row,column,script='')=>wait('(()=>{const cell=document.querySelector('+JSON.stringify('#cell-'+row+'-'+column)+');if(!cell)return false;cell.click();'+script+';return true})()');
      const enter=async(row,column,value)=>act(row,column,'const input=document.querySelector("#formula");input.value='+JSON.stringify(value)+';input.dispatchEvent(new Event("input",{bubbles:true}));input.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))');
      const open=async()=>{await act(0,1,'document.querySelector("#filterValuesButton").click()');await wait('!document.querySelector("#filterPanel").hidden')};const save=async()=>{await vscode.commands.executeCommand('workbench.action.files.save');return analyze(await fs.readFile(file));};
      await enter(3,1,'20');await wait('document.querySelector("#cell-3-1")?.textContent==="20"');await enter(1,2,'=B2*2');await wait('document.querySelector("#cell-3-2")?.textContent==="40"');await save();
      await open();await evaluate('document.querySelector("#filterMode").value="top10";document.querySelector("#filterMode").dispatchEvent(new Event("change"));document.querySelector("#filterTopValue").value="0";document.querySelector("#filterApply").click()');await wait('!document.querySelector("#filterEditor").disabled&&document.querySelector("#notice").textContent.includes("1–500")');assert.equal(await evaluate('document.querySelector("#filterTopValue").value'),'0');
      await evaluate('const input=document.querySelector("#filterTopValue");input.value="1";input.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))');await wait('document.querySelector("#filterPanel").hidden&&!document.querySelector("#cell-1-1")&&document.querySelector("#cell-2-1")?.textContent==="20"&&document.querySelector("#cell-3-1")?.textContent==="20"&&document.querySelector("#cell-5-3")?.textContent==="80"');let saved=await save();assert.equal(saved.sheets[0].tables[0].autoFilter.columns[0].top10.filterValue,20);assert.deepEqual(saved.sheets[0].nativeHiddenRows,[1]);
      await vscode.commands.executeCommand('jinsheet.undo');await wait('document.querySelector("#cell-1-1")?.textContent==="10"&&document.querySelector("#cell-5-3")?.textContent==="100"');await vscode.commands.executeCommand('jinsheet.redo');await wait('document.querySelector("#cell-5-3")?.textContent==="80"');await save();
      devtools.close();devtools=null;await vscode.commands.executeCommand('workbench.action.closeAllEditors');await vscode.commands.executeCommand('vscode.openWith',vscode.Uri.file(file),'jinsheet.editor');devtools=await waitForSpreadsheetWebview();await wait('document.querySelector("#cell-5-3")?.textContent==="80"');await open();assert.equal(await evaluate('document.querySelector("#filterMode").value'),'top10');assert.equal(await evaluate('document.querySelector("#filterTopValue").value'),'1');
      await evaluate('document.querySelector("#filterTopDirection").value="bottom";document.querySelector("#filterTopUnit").value="percent";document.querySelector("#filterTopUnit").dispatchEvent(new Event("change"));document.querySelector("#filterTopValue").value="33";document.querySelector("#filterApply").click()');await wait('document.querySelector("#filterPanel").hidden&&document.querySelector("#cell-1-1")?.textContent==="10"&&!document.querySelector("#cell-2-1")&&!document.querySelector("#cell-3-1")&&document.querySelector("#cell-5-3")?.textContent==="20"');saved=await save();assert.deepEqual(saved.sheets[0].tables[0].autoFilter.columns[0].top10,{top:false,percent:true,value:33,filterValue:10});
      await open();assert.equal(await evaluate('document.querySelector("#filterTopDirection").value'),'bottom');assert.equal(await evaluate('document.querySelector("#filterTopUnit").value'),'percent');await evaluate('document.querySelector("#filterClear").click()');await wait('document.querySelector("#filterPanel").hidden&&document.querySelector("#cell-5-3")?.textContent==="100"');saved=await save();assert.equal(saved.sheets[0].tables[0].autoFilter.columns.length,0);assert.deepEqual(saved.sheets[0].nativeHiddenRows,[]);
    } finally {devtools?.close();await vscode.commands.executeCommand('workbench.action.closeAllEditors');await fs.rm(directory,{recursive:true,force:true});}
  });

  test('Chromium applies native month and time groups with validation, save, reopen and undo', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-date-filter-host-')), file = path.join(directory, 'dates.xlsx');
    const sample = analyze(await fs.readFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx'))), files = { ...sample.files }, sheet = sample.sheets[0];
    const rows = '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c><c r="C1" t="inlineStr"><is><t>Result</t></is></c></row>' + [10,20,30].map((value,index)=>'<row r="'+(index+2)+'"><c r="B'+(index+2)+'"><v>'+value+'</v></c></row>').join('') + '<row r="6"><c r="D6"><f>SUBTOTAL(9,Sales[Result])</f><v>0</v></c></row></sheetData>';
    files[sheet.path] = new TextEncoder().encode(sheet.xml.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ').replace(/<sheetData>[\s\S]*?<\/sheetData>/, rows).replace('</worksheet>','<tableParts count="1"><tablePart r:id="t1"/></tableParts></worksheet>'));
    files[path.posix.join(path.posix.dirname(sheet.path),'_rels',path.posix.basename(sheet.path)+'.rels')] = new TextEncoder().encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="t1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>');
    files['xl/tables/table1.xml'] = new TextEncoder().encode('<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:C4"><autoFilter ref="A1:C4"/><tableColumns count="3"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Amount"/><tableColumn id="3" name="Result"/></tableColumns></table>');
    files['[Content_Types].xml'] = new TextEncoder().encode(new TextDecoder().decode(files['[Content_Types].xml']).replace('</Types>','<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>')); await fs.writeFile(file,zipSync(files)); let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith',vscode.Uri.file(file),'jinsheet.editor');devtools=await waitForSpreadsheetWebview();
      const evaluate=async expression=>{const result=await devtools.send('Runtime.evaluate',{expression,contextId:devtools.contextId,returnByValue:true,awaitPromise:true});if(result.exceptionDetails)throw new Error(result.exceptionDetails.exception?.description||result.exceptionDetails.text);return result.result?.value;};
      const wait=async expression=>{const deadline=Date.now()+6000;while(Date.now()<deadline){if(await evaluate(expression))return;await new Promise(resolve=>setTimeout(resolve,25));}assert.fail('Date filter failed: '+expression+' '+await evaluate('document.querySelector("#notice").textContent'));};
      const act=async(row,column,script='')=>wait('(()=>{const cell=document.querySelector('+JSON.stringify('#cell-'+row+'-'+column)+');if(!cell)return false;cell.click();'+script+';return true})()');
      const enter=async(row,column,value)=>act(row,column,'const input=document.querySelector("#formula");input.value='+JSON.stringify(value)+';input.dispatchEvent(new Event("input",{bubbles:true}));input.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))');
      const open=async()=>{await act(0,1,'document.querySelector("#filterValuesButton").click()');await wait('!document.querySelector("#filterPanel").hidden')};const save=async()=>{await vscode.commands.executeCommand('workbench.action.files.save');return analyze(await fs.readFile(file));};
      for(const [row,value] of [[1,'45351.5'],[2,'45352'],[3,'=DATE(2024,2,29)+0.75']]){await enter(row,1,value);await act(row,1,'document.querySelector("#formatDate").click()');await wait('document.querySelector("#cell-'+row+'-1")?.textContent.includes("/24")')}
      for(const [row,value] of [[1,'10'],[2,'20'],[3,'30']]){await enter(row,2,value);await wait('document.querySelector("#cell-'+row+'-2")?.textContent==="'+value+'"')}await wait('document.querySelector("#cell-5-3")?.textContent==="60"');await save();
      await open();await evaluate('document.querySelector("#filterMode").value="dates";document.querySelector("#filterMode").dispatchEvent(new Event("change"));document.querySelector("#filterDateGroups").value="2023-02-29";document.querySelector("#filterApply").click()');await wait('!document.querySelector("#filterEditor").disabled&&document.querySelector("#notice").textContent.includes("invalid")');assert.equal(await evaluate('document.querySelector("#filterDateGroups").value'),'2023-02-29');
      await evaluate('const input=document.querySelector("#filterDateGroups");input.value="2024-02";input.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))');await wait('document.querySelector("#filterPanel").hidden&&!document.querySelector("#cell-2-1")&&document.querySelector("#cell-1-1")?.textContent==="2/29/24"&&document.querySelector("#cell-3-1")?.textContent==="2/29/24"&&document.querySelector("#cell-5-3")?.textContent==="40"');let saved=await save();assert.deepEqual(saved.sheets[0].tables[0].autoFilter.columns[0].dateGroups,[{dateTimeGrouping:'month',year:2024,month:2}]);assert.deepEqual(saved.sheets[0].nativeHiddenRows,[2]);
      await vscode.commands.executeCommand('jinsheet.undo');await wait('document.querySelector("#cell-5-3")?.textContent==="60"');await vscode.commands.executeCommand('jinsheet.redo');await wait('document.querySelector("#cell-5-3")?.textContent==="40"');await save();devtools.close();devtools=null;await vscode.commands.executeCommand('workbench.action.closeAllEditors');await vscode.commands.executeCommand('vscode.openWith',vscode.Uri.file(file),'jinsheet.editor');devtools=await waitForSpreadsheetWebview();await wait('document.querySelector("#cell-5-3")?.textContent==="40"');await open();assert.equal(await evaluate('document.querySelector("#filterMode").value'),'dates');assert.equal(await evaluate('document.querySelector("#filterDateGroups").value'),'2024-02');
      await evaluate('document.querySelector("#filterDateGroups").value="2024-02-29T12";document.querySelector("#filterApply").click()');await wait('document.querySelector("#filterPanel").hidden&&document.querySelector("#cell-5-3")?.textContent==="10"&&!document.querySelector("#cell-3-1")');saved=await save();assert.equal(saved.sheets[0].tables[0].autoFilter.columns[0].dateGroups[0].hour,12);
      await open();await evaluate('document.querySelector("#filterDateGroups").value="2024-02-29T12"+String.fromCharCode(10)+"2024-03";document.querySelector("#filterApply").click()');await wait('document.querySelector("#filterPanel").hidden&&document.querySelector("#cell-5-3")?.textContent==="30"');saved=await save();assert.equal(saved.sheets[0].tables[0].autoFilter.columns[0].dateGroups.length,2);await open();await evaluate('document.querySelector("#filterClear").click()');await wait('document.querySelector("#filterPanel").hidden&&document.querySelector("#cell-5-3")?.textContent==="60"');saved=await save();assert.deepEqual(saved.sheets[0].nativeHiddenRows,[]);
    } finally {devtools?.close();await vscode.commands.executeCommand('workbench.action.closeAllEditors');await fs.rm(directory,{recursive:true,force:true});}
  });

  test('Chromium calculates 1904 date formulas through date filters, save, reopen and undo', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-epoch-formulas-host-')), file = path.join(directory, 'epoch.xlsx');
    const sample = analyze(await fs.readFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx'))), files = { ...sample.files }, sheet = sample.sheets[0];
    const rows = '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c><c r="C1" t="inlineStr"><is><t>Result</t></is></c></row>' + [10,20,30].map((value,index)=>'<row r="'+(index+2)+'"><c r="B'+(index+2)+'"><v>'+value+'</v></c></row>').join('') + '<row r="6"><c r="D6"><f>SUBTOTAL(9,Sales[Result])</f><v>0</v></c></row></sheetData>';
    files[sheet.path] = new TextEncoder().encode(sheet.xml.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ').replace(/<sheetData>[\s\S]*?<\/sheetData>/, rows).replace('</worksheet>','<tableParts count="1"><tablePart r:id="t1"/></tableParts></worksheet>'));
    files[path.posix.join(path.posix.dirname(sheet.path),'_rels',path.posix.basename(sheet.path)+'.rels')] = new TextEncoder().encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="t1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>');
    files['xl/tables/table1.xml'] = new TextEncoder().encode('<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:C4"><autoFilter ref="A1:C4"/><tableColumns count="3"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Amount"/><tableColumn id="3" name="Result"/></tableColumns></table>');
    files['[Content_Types].xml'] = new TextEncoder().encode(new TextDecoder().decode(files['[Content_Types].xml']).replace('</Types>','<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>')); const workbookXml=new TextDecoder().decode(files['xl/workbook.xml']);files['xl/workbook.xml']=new TextEncoder().encode(/<workbookPr\b/.test(workbookXml)?workbookXml.replace(/<workbookPr\b[^>]*\/>/,'<workbookPr date1904="1"/>'):workbookXml.replace('<sheets>','<workbookPr date1904="1"/><sheets>'));await fs.writeFile(file,zipSync(files)); let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith',vscode.Uri.file(file),'jinsheet.editor');devtools=await waitForSpreadsheetWebview();
      const evaluate=async expression=>{const result=await devtools.send('Runtime.evaluate',{expression,contextId:devtools.contextId,returnByValue:true,awaitPromise:true});if(result.exceptionDetails)throw new Error(result.exceptionDetails.exception?.description||result.exceptionDetails.text);return result.result?.value;};
      const wait=async expression=>{const deadline=Date.now()+6000;while(Date.now()<deadline){if(await evaluate(expression))return;await new Promise(resolve=>setTimeout(resolve,25));}assert.fail('1904 formulas failed: '+expression+' '+await evaluate('document.querySelector("#notice").textContent'));};
      const act=async(row,column,script='')=>wait('(()=>{const cell=document.querySelector('+JSON.stringify('#cell-'+row+'-'+column)+');if(!cell)return false;cell.click();'+script+';return true})()');
      const enter=async(row,column,value)=>act(row,column,'const input=document.querySelector("#formula");input.value='+JSON.stringify(value)+';input.dispatchEvent(new Event("input",{bubbles:true}));input.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))');const save=async()=>{await vscode.commands.executeCommand('workbench.action.files.save');return analyze(await fs.readFile(file));};
      await wait('Boolean(document.querySelector("#autoCalculatedColumns"))');await evaluate('document.querySelector("#autoCalculatedColumns").checked=false');
      for(const [row,value,display] of [[1,'=DATE(2024,2,29)','2/29/24'],[2,'=EDATE(B2,1)','3/29/24'],[3,'=WORKDAY(B2,1)','3/1/24']]){await enter(row,1,value);await act(row,1,'document.querySelector("#formatDate").click()');await wait('document.querySelector("#cell-'+row+'-1")?.textContent==="'+display+'"')}
      for(const [row,value] of [[1,'10'],[2,'20'],[3,'30']]){await enter(row,2,value);await wait('document.querySelector("#cell-'+row+'-2")?.textContent==="'+value+'"')}
      await enter(1,5,'=YEAR(B2)');await wait('document.querySelector("#cell-1-5")?.textContent==="2024"');await enter(2,5,'=TEXT(B2,"yyyy-mm-dd")');await wait('document.querySelector("#cell-2-5")?.textContent==="2024-02-29"');await enter(3,5,'=WEEKDAY(0)');await wait('document.querySelector("#cell-3-5")?.textContent==="6"');await enter(2,8,'=TEXT(B2+0.5,"yyyy-mm-dd ""at"" h:mm A/P")');await wait('document.querySelector("#cell-2-8")?.textContent==="2024-02-29 at 12:00 P"');let saved=await save();assert.equal(saved.date1904,true);assert.equal(saved.sheets[0].cells.get('B2').formula,'DATE(2024,2,29)');assert.equal(saved.sheets[0].cells.get('I3').formula,'TEXT(B2+0.5,"yyyy-mm-dd ""at"" h:mm A/P")');
      await enter(1,1,'=DATE(2024,3,1)');await wait('document.querySelector("#cell-1-1")?.textContent==="3/1/24"&&document.querySelector("#cell-2-1")?.textContent==="4/1/24"&&document.querySelector("#cell-3-1")?.textContent==="3/4/24"');await vscode.commands.executeCommand('jinsheet.undo');await wait('document.querySelector("#cell-1-1")?.textContent==="2/29/24"');await vscode.commands.executeCommand('jinsheet.redo');await wait('document.querySelector("#cell-1-1")?.textContent==="3/1/24"');await save();
      await act(0,1,'document.querySelector("#filterValuesButton").click()');await wait('!document.querySelector("#filterPanel").hidden');await evaluate('document.querySelector("#filterMode").value="dates";document.querySelector("#filterMode").dispatchEvent(new Event("change"));document.querySelector("#filterDateGroups").value="2024-03";document.querySelector("#filterApply").click()');await wait('document.querySelector("#filterPanel").hidden&&!document.querySelector("#cell-2-1")&&document.querySelector("#cell-5-3")?.textContent==="40"');saved=await save();assert.deepEqual(saved.sheets[0].nativeHiddenRows,[2]);
      devtools.close();devtools=null;await vscode.commands.executeCommand('workbench.action.closeAllEditors');await vscode.commands.executeCommand('vscode.openWith',vscode.Uri.file(file),'jinsheet.editor');devtools=await waitForSpreadsheetWebview();await wait('document.querySelector("#cell-1-1")?.textContent==="3/1/24"&&document.querySelector("#cell-3-1")?.textContent==="3/4/24"&&document.querySelector("#cell-5-3")?.textContent==="40"');
    } finally {devtools?.close();await vscode.commands.executeCommand('workbench.action.closeAllEditors');await fs.rm(directory,{recursive:true,force:true});}
  });

  test('Chromium applies relative-date and average filters through explicit apply, save, reopen and undo', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-dynamic-filter-host-')), file = path.join(directory, 'relative.xlsx');
    const sample = analyze(await fs.readFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx'))), files = { ...sample.files }, sheet = sample.sheets[0];
    const rows = '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c><c r="C1" t="inlineStr"><is><t>Result</t></is></c></row>' + [10,20,30].map((value,index)=>'<row r="'+(index+2)+'"><c r="B'+(index+2)+'"><v>'+value+'</v></c></row>').join('') + '<row r="6"><c r="D6"><f>SUBTOTAL(9,Sales[Result])</f><v>0</v></c></row></sheetData>';
    files[sheet.path] = new TextEncoder().encode(sheet.xml.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ').replace(/<sheetData>[\s\S]*?<\/sheetData>/, rows).replace('</worksheet>','<tableParts count="1"><tablePart r:id="t1"/></tableParts></worksheet>'));
    files[path.posix.join(path.posix.dirname(sheet.path),'_rels',path.posix.basename(sheet.path)+'.rels')] = new TextEncoder().encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="t1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>');
    files['xl/tables/table1.xml'] = new TextEncoder().encode('<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:C4"><autoFilter ref="A1:C4"/><tableColumns count="3"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Amount"/><tableColumn id="3" name="Result"/></tableColumns></table>');
    files['[Content_Types].xml'] = new TextEncoder().encode(new TextDecoder().decode(files['[Content_Types].xml']).replace('</Types>','<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>')); await fs.writeFile(file,zipSync(files)); let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith',vscode.Uri.file(file),'jinsheet.editor');devtools=await waitForSpreadsheetWebview();
      const evaluate=async expression=>{const result=await devtools.send('Runtime.evaluate',{expression,contextId:devtools.contextId,returnByValue:true,awaitPromise:true});if(result.exceptionDetails)throw new Error(result.exceptionDetails.exception?.description||result.exceptionDetails.text);return result.result?.value;};
      const wait=async expression=>{const deadline=Date.now()+6000;while(Date.now()<deadline){if(await evaluate(expression))return;await new Promise(resolve=>setTimeout(resolve,25));}assert.fail('Dynamic filter failed: '+expression+' '+await evaluate('document.querySelector("#notice").textContent'));};
      const act=async(row,column,script='')=>wait('(()=>{const cell=document.querySelector('+JSON.stringify('#cell-'+row+'-'+column)+');if(!cell)return false;cell.click();'+script+';return true})()');
      const enter=async(row,column,value)=>act(row,column,'const input=document.querySelector("#formula");input.value='+JSON.stringify(value)+';input.dispatchEvent(new Event("input",{bubbles:true}));input.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))');const save=async()=>{await vscode.commands.executeCommand('workbench.action.files.save');return analyze(await fs.readFile(file));};const open=async column=>{await act(0,column,'document.querySelector("#filterValuesButton").click()');await wait('!document.querySelector("#filterPanel").hidden')};
      await wait('Boolean(document.querySelector("#autoCalculatedColumns"))');await evaluate('document.querySelector("#autoCalculatedColumns").checked=false');for(const [row,value] of [[1,'=TODAY()+0.5'],[2,'=TODAY()-1'],[3,'=TODAY()+1.75']]){await enter(row,1,value);await act(row,1,'document.querySelector("#formatDate").click()');await wait('document.querySelector("#cell-'+row+'-1")?.textContent.includes("/")')}
      for(const [row,value] of [[1,'10'],[2,'20'],[3,'30']]){await enter(row,2,value);await wait('document.querySelector("#cell-'+row+'-2")?.textContent==="'+value+'"')}await wait('document.querySelector("#cell-5-3")?.textContent==="60"');await save();
      await open(1);await evaluate('document.querySelector("#filterMode").value="dynamic";document.querySelector("#filterMode").dispatchEvent(new Event("change"));document.querySelector("#filterDynamicType").value="today"');assert.equal(await evaluate('document.querySelector("#cell-5-3").textContent'),'60');await evaluate('document.querySelector("#filterApply").click()');await wait('document.querySelector("#filterPanel").hidden&&document.querySelector("#cell-5-3")?.textContent==="10"&&!document.querySelector("#cell-2-1")&&!document.querySelector("#cell-3-1")');let saved=await save();const today=saved.sheets[0].tables[0].autoFilter.columns[0].dynamic;assert.equal(today.type,'today');assert.equal(today.maxValue-today.value,1);assert.deepEqual(saved.sheets[0].nativeHiddenRows,[2,3]);
      devtools.close();devtools=null;await vscode.commands.executeCommand('workbench.action.closeAllEditors');await vscode.commands.executeCommand('vscode.openWith',vscode.Uri.file(file),'jinsheet.editor');devtools=await waitForSpreadsheetWebview();await wait('document.querySelector("#cell-5-3")?.textContent==="10"');await open(1);assert.equal(await evaluate('document.querySelector("#filterMode").value'),'dynamic');assert.equal(await evaluate('document.querySelector("#filterDynamicType").value'),'today');await evaluate('const input=document.querySelector("#filterDynamicType");input.value="yesterday";input.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))');await wait('document.querySelector("#filterPanel").hidden&&document.querySelector("#cell-5-3")?.textContent==="20"');await save();
      await open(1);await evaluate('document.querySelector("#filterDynamicType").value="null";document.querySelector("#filterApply").click()');await wait('document.querySelector("#filterPanel").hidden&&document.querySelector("#cell-5-3")?.textContent==="60"');await open(2);await evaluate('document.querySelector("#filterMode").value="dynamic";document.querySelector("#filterMode").dispatchEvent(new Event("change"));document.querySelector("#filterDynamicType").value="aboveAverage";document.querySelector("#filterApply").click()');await wait('document.querySelector("#filterPanel").hidden&&document.querySelector("#cell-5-3")?.textContent==="30"');saved=await save();assert.equal(saved.sheets[0].tables[0].autoFilter.columns.find(item=>item.column===2).dynamic.value,20);
      await vscode.commands.executeCommand('jinsheet.undo');await wait('document.querySelector("#cell-5-3")?.textContent==="60"');await vscode.commands.executeCommand('jinsheet.redo');await wait('document.querySelector("#cell-5-3")?.textContent==="30"');await open(2);await evaluate('document.querySelector("#filterClear").click()');await wait('document.querySelector("#filterPanel").hidden&&document.querySelector("#cell-5-3")?.textContent==="60"');await open(1);await evaluate('document.querySelector("#filterClear").click()');await wait('document.querySelector("#filterPanel").hidden&&document.querySelector("#cell-5-3")?.textContent==="60"');saved=await save();assert.equal(saved.sheets[0].tables[0].autoFilter.columns.length,0);assert.deepEqual(saved.sheets[0].nativeHiddenRows,[]);
    } finally {devtools?.close();await vscode.commands.executeCommand('workbench.action.closeAllEditors');await fs.rm(directory,{recursive:true,force:true});}
  });

  test('Chromium displays ISO date cells and preserves dates through filter, edit, save and undo', async () => {
    const directory=await fs.mkdtemp(path.join(os.tmpdir(),'jinsheet-iso-date-host-')),file=path.join(directory,'dates.xlsx');
    const sample=analyze(await fs.readFile(path.join(__dirname,'..','fixtures','JinSheet-Styled-Sample.xlsx'))),files={...sample.files},sheet=sample.sheets[0],encode=text=>new TextEncoder().encode(text),decode=bytes=>new TextDecoder().decode(bytes);
    const data='<sheetData><row r="1"><c r="B1" t="inlineStr"><is><t>Date</t></is></c></row><row r="2"><c r="B2"><v>45351.5</v></c></row><row r="3"><c r="B3"><v>45352</v></c></row><row r="6"><c r="C6"><f>SUBTOTAL(9,B2:B3)</f><v>0</v></c></row></sheetData>';
    files[sheet.path]=encode(sheet.xml.replace(/<sheetData>[\s\S]*?<\/sheetData>/,data).replace('</worksheet>','<autoFilter ref="B1:B3"/></worksheet>'));
    const {patchCells}=require('../../src/xlsx'),styled=patchCells(analyze(zipSync(files)),[{sheetIndex:0,address:'B2',value:'45351.5',kind:'number',numberFormatId:14,forceNumberFormat:true},{sheetIndex:0,address:'B3',value:'45352',kind:'number',numberFormatId:14,forceNumberFormat:true}]).model,prepared={...styled.files};
    let xml=decode(prepared[sheet.path]);for(const [address,date]of [['B2','2024-02-29T12:00:00Z'],['B3','2024-03-01']])xml=xml.replace(new RegExp('(<c\\b[^>]*\\br="'+address+'"[^>]*)(><v>)[^<]*(</v></c>)'),'$1 t="d"$2'+date+'$3');prepared[sheet.path]=encode(xml);await fs.writeFile(file,zipSync(prepared));let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith',vscode.Uri.file(file),'jinsheet.editor');devtools=await waitForSpreadsheetWebview();
      const evaluate=async expression=>{const result=await devtools.send('Runtime.evaluate',{expression,contextId:devtools.contextId,returnByValue:true,awaitPromise:true});if(result.exceptionDetails)throw Error(result.exceptionDetails.exception?.description||result.exceptionDetails.text);return result.result?.value;};
      const wait=async expression=>{const deadline=Date.now()+6000;while(Date.now()<deadline){if(await evaluate(expression))return;await new Promise(resolve=>setTimeout(resolve,25));}assert.fail('ISO date UI: '+expression);};
      const act=async(row,column,script='')=>wait('(()=>{const cell=document.querySelector('+JSON.stringify('#cell-'+row+'-'+column)+');if(!cell)return false;cell.click();'+script+';return true})()');
      const open=async()=>{await act(0,1,'document.querySelector("#filterValuesButton").click()');await wait('!document.querySelector("#filterPanel").hidden')};
      await wait('document.querySelector("#cell-5-2")?.textContent==="90703.5"');assert.match(await evaluate('document.querySelector("#cell-1-1").textContent'),/2024|24/);
      await open();await evaluate('document.querySelector("#filterMode").value="dynamic";document.querySelector("#filterMode").dispatchEvent(new Event("change"));document.querySelector("#filterDynamicType").value="M2";document.querySelector("#filterApply").click()');await wait('document.querySelector("#filterPanel").hidden&&document.querySelector("#cell-5-2")?.textContent==="45351.5"&&!document.querySelector("#cell-2-1")');
      await vscode.commands.executeCommand('workbench.action.files.save');let saved=analyze(await fs.readFile(file));assert.equal(saved.sheets[0].cells.get('B2').isoDate,'2024-02-29T12:00:00Z');assert.deepEqual(saved.sheets[0].nativeHiddenRows,[2]);
      await act(1,1,'const input=document.querySelector("#formula");input.value="45352";input.dispatchEvent(new Event("input",{bubbles:true}));input.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))');await wait('document.querySelector("#cell-5-2")?.textContent==="0"');await vscode.commands.executeCommand('jinsheet.undo');await wait('document.querySelector("#cell-5-2")?.textContent==="45351.5"');await vscode.commands.executeCommand('jinsheet.redo');await wait('document.querySelector("#cell-5-2")?.textContent==="0"');await vscode.commands.executeCommand('jinsheet.undo');await wait('document.querySelector("#cell-5-2")?.textContent==="45351.5"');await vscode.commands.executeCommand('workbench.action.files.save');
      devtools.close();devtools=null;await vscode.commands.executeCommand('workbench.action.closeAllEditors');await vscode.commands.executeCommand('vscode.openWith',vscode.Uri.file(file),'jinsheet.editor');devtools=await waitForSpreadsheetWebview();await wait('document.querySelector("#cell-5-2")?.textContent==="45351.5"');await open();assert.equal(await evaluate('document.querySelector("#filterDynamicType").value'),'M2');await evaluate('document.querySelector("#filterClear").click()');await wait('document.querySelector("#cell-5-2")?.textContent==="90703.5"');await vscode.commands.executeCommand('workbench.action.files.save');saved=analyze(await fs.readFile(file));assert.equal(saved.sheets[0].cells.get('B2').isoDate,'2024-02-29T12:00:00Z');assert.deepEqual(saved.sheets[0].nativeHiddenRows,[]);
    } finally {devtools?.close();await vscode.commands.executeCommand('workbench.action.closeAllEditors');await fs.rm(directory,{recursive:true,force:true});}
  });

  for (const themed of [false, true]) test('Chromium filters '+(themed?'theme, tint and custom indexed':'RGB')+' fill and font colors including conditional formatting through save, reopen and undo', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-color-filter-host-')), file = path.join(directory, 'colors.xlsx');
    const sample = analyze(await fs.readFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx'))), files = { ...sample.files }, sheet = sample.sheets[0];
    const rows = '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c><c r="C1" t="inlineStr"><is><t>Result</t></is></c></row>' + [10,20,30].map((value,index)=>'<row r="'+(index+2)+'"><c r="B'+(index+2)+'"><v>'+value+'</v></c></row>').join('') + '<row r="6"><c r="D6"><f>SUBTOTAL(9,Sales[Result])</f><v>0</v></c></row></sheetData>';
    files[sheet.path] = new TextEncoder().encode(sheet.xml.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ').replace(/<sheetData>[\s\S]*?<\/sheetData>/, rows).replace('</worksheet>','<tableParts count="1"><tablePart r:id="t1"/></tableParts></worksheet>'));
    files[path.posix.join(path.posix.dirname(sheet.path),'_rels',path.posix.basename(sheet.path)+'.rels')] = new TextEncoder().encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="t1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>');
    files['xl/tables/table1.xml'] = new TextEncoder().encode('<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:C4"><autoFilter ref="A1:C4"/><tableColumns count="3"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Amount"/><tableColumn id="3" name="Result"/></tableColumns></table>');
    files['[Content_Types].xml'] = new TextEncoder().encode(new TextDecoder().decode(files['[Content_Types].xml']).replace('</Types>','<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>')); const {patchCells}=require('../../src/xlsx');const colored=patchCells(analyze(zipSync(files)),[...[10,20,30].flatMap((value,index)=>[{sheetIndex:0,address:'B'+(index+2),value:String(value),fillColor:index===1?'#0000FF':'#FF0000',font:{bold:index===0,italic:false,color:index===2?'#0000FF':'#00FF00'}},{sheetIndex:0,address:'C'+(index+2),value:String(value*2),formula:'B'+(index+2)+'*2'}])]);const prepared={...colored.model.files},stylesPath=colored.model.stylesPath,styles=new TextDecoder().decode(prepared[stylesPath]),dxf='<dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs>';prepared[stylesPath]=new TextEncoder().encode(/<dxfs\b/.test(styles)?styles.replace(/<dxfs\b[^>]*(?:\/>|>[\s\S]*?<\/dxfs>)/,dxf):styles.replace(/(?=<tableStyles\b|<colors\b|<extLst\b|<\/styleSheet>)/,dxf));prepared[sheet.path]=new TextEncoder().encode(new TextDecoder().decode(prepared[sheet.path]).replace('</worksheet>','<conditionalFormatting sqref="C2:C4"><cfRule type="cellIs" operator="greaterThan" dxfId="0" priority="1"><formula>30</formula></cfRule></conditionalFormatting></worksheet>'));if(themed){
      const encode=text=>new TextEncoder().encode(text),decode=bytes=>new TextDecoder().decode(bytes);
      prepared['xl/_rels/workbook.xml.rels']=encode(decode(prepared['xl/_rels/workbook.xml.rels']).replace(/<Relationship\b[^>]*Type="[^"]*\/theme"[^>]*\/>/g,'').replace('</Relationships>','<Relationship Id="colorTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>'));
      prepared['xl/theme/theme1.xml']=encode('<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Colors"><a:themeElements><a:clrScheme name="Colors"><a:accent1><a:srgbClr val="800000"/></a:accent1></a:clrScheme></a:themeElements></a:theme>');
      let xml=decode(prepared[stylesPath]).replaceAll('rgb="FFFF0000"','theme="4" tint="0.3333333333333333"').replaceAll('rgb="FF00FF00"','indexed="0"').replaceAll('rgb="FF0000FF"','indexed="1"');
      const colors='<colors><indexedColors><rgbColor rgb="FF00FF00"/><rgbColor rgb="FF0000FF"/></indexedColors></colors>';
      xml=/<colors\b/.test(xml)?xml.replace(/<colors\b[^>]*(?:\/>|>[\s\S]*?<\/colors>)/,colors):xml.replace('</styleSheet>',colors+'</styleSheet>');prepared[stylesPath]=encode(xml);
      const types=decode(prepared['[Content_Types].xml']);if(!types.includes('PartName="/xl/theme/theme1.xml"'))prepared['[Content_Types].xml']=encode(types.replace('</Types>','<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>'));
    }
    await fs.writeFile(file,zipSync(prepared));let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith',vscode.Uri.file(file),'jinsheet.editor');devtools=await waitForSpreadsheetWebview();
      const evaluate=async expression=>{const result=await devtools.send('Runtime.evaluate',{expression,contextId:devtools.contextId,returnByValue:true,awaitPromise:true});if(result.exceptionDetails)throw new Error(result.exceptionDetails.exception?.description||result.exceptionDetails.text);return result.result?.value;};
      const wait=async expression=>{const deadline=Date.now()+6000;while(Date.now()<deadline){if(await evaluate(expression))return;await new Promise(resolve=>setTimeout(resolve,25));}assert.fail('Color filter failed: '+expression+' '+await evaluate('document.querySelector("#notice").textContent'));};const act=async(row,column,script='')=>wait('(()=>{const cell=document.querySelector('+JSON.stringify('#cell-'+row+'-'+column)+');if(!cell)return false;cell.click();'+script+';return true})()');const open=async column=>{await act(0,column,'document.querySelector("#filterValuesButton").click()');await wait('!document.querySelector("#filterPanel").hidden')};const save=async()=>{await vscode.commands.executeCommand('workbench.action.files.save');return analyze(await fs.readFile(file));};
      await wait('document.querySelector("#cell-5-3")?.textContent==="120"');const initial=analyze(await fs.readFile(file)),originalStyles=['B2','B3','B4'].map(address=>initial.sheets[0].cells.get(address).style);await open(1);await evaluate('document.querySelector("#filterMode").value="color";document.querySelector("#filterMode").dispatchEvent(new Event("change"));document.querySelector("#filterColorValue").value="#FF0000";document.querySelector("#filterColorValue").dispatchEvent(new Event("change"))');assert.equal(await evaluate('document.querySelector("#cell-5-3").textContent'),'120');await evaluate('document.querySelector("#filterApply").click()');await wait('document.querySelector("#filterPanel").hidden&&document.querySelector("#cell-5-3")?.textContent==="80"&&!document.querySelector("#cell-2-1")');let saved=await save();assert.equal(saved.sheets[0].tables[0].autoFilter.columns[0].colorFilter.color,'#FF0000');assert.deepEqual(['B2','B3','B4'].map(address=>saved.sheets[0].cells.get(address).style),originalStyles);
      await open(1);await evaluate('document.querySelector("#filterColorTarget").value="font";document.querySelector("#filterColorTarget").dispatchEvent(new Event("change"));document.querySelector("#filterColorValue").value="#00FF00";document.querySelector("#filterApply").click()');await wait('document.querySelector("#filterPanel").hidden&&document.querySelector("#cell-5-3")?.textContent==="60"&&!document.querySelector("#cell-3-1")');saved=await save();assert.equal(saved.sheets[0].tables[0].autoFilter.columns[0].colorFilter.cellColor,false);await vscode.commands.executeCommand('jinsheet.undo');await wait('document.querySelector("#cell-5-3")?.textContent==="80"');await vscode.commands.executeCommand('jinsheet.redo');await wait('document.querySelector("#cell-5-3")?.textContent==="60"');await save();devtools.close();devtools=null;await vscode.commands.executeCommand('workbench.action.closeAllEditors');await vscode.commands.executeCommand('vscode.openWith',vscode.Uri.file(file),'jinsheet.editor');devtools=await waitForSpreadsheetWebview();await wait('document.querySelector("#cell-5-3")?.textContent==="60"');await open(1);assert.equal(await evaluate('document.querySelector("#filterMode").value'),'color');assert.equal(await evaluate('document.querySelector("#filterColorTarget").value'),'font');assert.equal(await evaluate('document.querySelector("#filterColorValue").value'),'#00FF00');await evaluate('document.querySelector("#filterClear").click()');await wait('document.querySelector("#filterPanel").hidden&&document.querySelector("#cell-5-3")?.textContent==="120"');
      await open(2);await evaluate('document.querySelector("#filterMode").value="color";document.querySelector("#filterMode").dispatchEvent(new Event("change"));document.querySelector("#filterColorValue").value="#FF0000";document.querySelector("#filterApply").click()');await wait('document.querySelector("#filterPanel").hidden&&document.querySelector("#cell-5-3")?.textContent==="100"&&!document.querySelector("#cell-1-2")');await act(2,1,'const input=document.querySelector("#formula");input.value="5";input.dispatchEvent(new Event("input",{bubbles:true}));input.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))');await wait('document.querySelector("#cell-5-3")?.textContent==="60"&&!document.querySelector("#cell-2-2")');saved=await save();assert.deepEqual(saved.sheets[0].nativeHiddenRows,[1,2]);assert.equal(saved.sheets[0].cells.get('C3').formula,'B3*2');assert.equal(saved.sheets[0].cells.get('B3').style,originalStyles[1]);await vscode.commands.executeCommand('jinsheet.undo');await wait('document.querySelector("#cell-5-3")?.textContent==="100"');await vscode.commands.executeCommand('jinsheet.redo');await wait('document.querySelector("#cell-5-3")?.textContent==="60"');await open(2);await evaluate('document.querySelector("#filterClear").click()');await wait('document.querySelector("#filterPanel").hidden&&document.querySelector("#cell-5-3")?.textContent==="90"');saved=await save();assert.deepEqual(saved.sheets[0].nativeHiddenRows,[]);
    } finally {devtools?.close();await vscode.commands.executeCommand('workbench.action.closeAllEditors');await fs.rm(directory,{recursive:true,force:true});}
  });

  test('Chromium auto-fills calculated columns and preserves exceptions through explicit apply, structure, save and undo', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-calculated-host-')), file = path.join(directory, 'calculated.xlsx');
    const sample = analyze(await fs.readFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx'))), files = { ...sample.files }, sheet = sample.sheets[0];
    const rows = '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c><c r="C1" t="inlineStr"><is><t>Result</t></is></c></row>' + [10,20,30].map((value,index)=>'<row r="'+(index+2)+'"><c r="B'+(index+2)+'"><v>'+value+'</v></c></row>').join('') + '<row r="6"><c r="D6"><f>SUBTOTAL(9,Sales[Result])</f><v>0</v></c></row></sheetData>';
    files[sheet.path] = new TextEncoder().encode(sheet.xml.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ').replace(/<sheetData>[\s\S]*?<\/sheetData>/, rows).replace('</worksheet>','<tableParts count="1"><tablePart r:id="t1"/></tableParts></worksheet>'));
    files[path.posix.join(path.posix.dirname(sheet.path),'_rels',path.posix.basename(sheet.path)+'.rels')] = new TextEncoder().encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="t1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>');
    files['xl/tables/table1.xml'] = new TextEncoder().encode('<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:C4"><autoFilter ref="A1:C4"/><tableColumns count="3"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Amount"/><tableColumn id="3" name="Result"/></tableColumns></table>');
    files['[Content_Types].xml'] = new TextEncoder().encode(new TextDecoder().decode(files['[Content_Types].xml']).replace('</Types>','<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>')); await fs.writeFile(file,zipSync(files)); let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith',vscode.Uri.file(file),'jinsheet.editor'); devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {const result = await devtools.send('Runtime.evaluate',{expression,contextId:devtools.contextId,returnByValue:true,awaitPromise:true});if(result.exceptionDetails)throw new Error(result.exceptionDetails.exception?.description||result.exceptionDetails.text);return result.result?.value;};
      const wait = async expression => {const deadline=Date.now()+6000;while(Date.now()<deadline){if(await evaluate(expression))return;await new Promise(resolve=>setTimeout(resolve,25));}assert.fail('Calculated column failed: '+expression+' '+await evaluate('document.querySelector("#notice").textContent'));};
      const act = async (row,column,script='') => wait('(()=>{const cell=document.querySelector('+JSON.stringify('#cell-'+row+'-'+column)+');if(!cell)return false;cell.click();'+script+';return true})()');
      const enter = async (row,formula) => act(row,2,'const input=document.querySelector("#formula");input.value='+JSON.stringify(formula)+';input.dispatchEvent(new Event("input",{bubbles:true}));input.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))');
      await enter(2,'=B3*2'); await wait('document.querySelector("#cell-1-2")?.textContent==="20"&&document.querySelector("#cell-2-2")?.textContent==="40"&&document.querySelector("#cell-3-2")?.textContent==="60"'); await vscode.commands.executeCommand('workbench.action.files.save');
      let saved=analyze(await fs.readFile(file));assert.equal(saved.sheets[0].tables[0].calculatedColumns[2],'B2*2');assert.equal(saved.sheets[0].cells.get('C4').formula,'B4*2');
      await vscode.commands.executeCommand('jinsheet.undo');await wait('document.querySelector("#cell-1-2")?.textContent===""&&document.querySelector("#cell-2-2")?.textContent===""');await vscode.commands.executeCommand('workbench.action.files.save');assert.equal(analyze(await fs.readFile(file)).sheets[0].tables[0].calculatedColumns[2],null);
      await vscode.commands.executeCommand('jinsheet.redo');await wait('document.querySelector("#cell-3-2")?.textContent==="60"');
      await evaluate('document.querySelector("#autoCalculatedColumns").checked=false');await enter(2,'=B3*5');await wait('document.querySelector("#cell-1-2")?.textContent==="20"&&document.querySelector("#cell-2-2")?.textContent==="100"&&document.querySelector("#cell-3-2")?.textContent==="60"');
      await act(2,2,'document.querySelector("#grid").dispatchEvent(new KeyboardEvent("keydown",{key:"c",ctrlKey:true,bubbles:true}))');await act(1,2,'document.querySelector("#grid").dispatchEvent(new KeyboardEvent("keydown",{key:"v",ctrlKey:true,bubbles:true}))');await wait('document.querySelector("#cell-1-2")?.textContent==="50"&&document.querySelector("#cell-2-2")?.textContent==="100"&&document.querySelector("#cell-3-2")?.textContent==="60"');
      await evaluate('document.querySelector("#autoCalculatedColumns").checked=true');await enter(1,'=B2*3');await wait('document.querySelector("#cell-1-2")?.textContent==="30"&&document.querySelector("#cell-2-2")?.textContent==="100"&&document.querySelector("#cell-3-2")?.textContent==="90"');
      await act(1,2);await wait('!document.querySelector("#applyCalculatedColumn").disabled');await evaluate('document.querySelector("#calculatedColumnScope").value="preserve";document.querySelector("#applyCalculatedColumn").click()');await wait('!document.querySelector("#applyCalculatedColumn").disabled&&document.querySelector("#notice").textContent.includes("Exceptions kept: 1")&&document.querySelector("#cell-2-2")?.textContent==="100"');
      await act(1,2);await wait('!document.querySelector("#applyCalculatedColumn").disabled');await evaluate('document.querySelector("#calculatedColumnScope").value="all";document.querySelector("#applyCalculatedColumn").click()');await wait('document.querySelector("#cell-2-2")?.textContent==="60"&&!document.querySelector("#applyCalculatedColumn").disabled');await vscode.commands.executeCommand('workbench.action.files.save');
      saved=analyze(await fs.readFile(file));assert.equal(saved.sheets[0].cells.get('C3').formula,'B3*3');
      await act(1,2,'document.querySelector("#insertRow").click()');await wait('document.querySelector("#cell-1-2")?.textContent==="0"&&document.querySelector("#cell-2-2")?.textContent==="30"&&document.querySelector("#cell-4-2")?.textContent==="90"');await vscode.commands.executeCommand('workbench.action.files.save');
      saved=analyze(await fs.readFile(file));assert.equal(saved.sheets[0].tables[0].calculatedColumns[2],'B2*3');assert.equal(saved.sheets[0].cells.get('C2').formula,'B2*3');
      await act(1,2,'document.querySelector("#deleteRow").click()');await wait('document.querySelector("#cell-1-2")?.textContent==="30"&&document.querySelector("#cell-2-2")?.textContent==="60"&&document.querySelector("#cell-3-2")?.textContent==="90"');
      const calculatedSaveDeadline=Date.now()+5000;do{await vscode.commands.executeCommand('workbench.action.files.save');saved=analyze(await fs.readFile(file));if(saved.sheets[0].tables[0].range.bottom===3&&saved.sheets[0].cells.get('C2')?.formula==='B2*3'&&saved.sheets[0].cells.get('C3')?.formula==='B3*3'&&saved.sheets[0].cells.get('C4')?.formula==='B4*3')break;await new Promise(resolve=>setTimeout(resolve,25))}while(Date.now()<calculatedSaveDeadline);
      assert.equal(saved.sheets[0].tables[0].calculatedColumns[2],'B2*3');assert.equal(saved.sheets[0].cells.get('C4')?.formula,'B4*3');
      devtools.close();devtools=null;await vscode.commands.executeCommand('workbench.action.closeAllEditors');await vscode.commands.executeCommand('vscode.openWith',vscode.Uri.file(file),'jinsheet.editor');devtools=await waitForSpreadsheetWebview();await wait('document.querySelector("#cell-3-2")?.textContent==="90"');
      await act(1,2,'document.querySelector("#clearCalculatedColumn").click()');await wait('document.querySelector("#clearCalculatedColumn").disabled&&document.querySelector("#cell-1-2")?.textContent==="30"');await act(2,2,'document.querySelector("#insertRow").click()');await wait('document.querySelector("#cell-2-2")?.textContent===""&&document.querySelector("#cell-3-2")?.textContent==="60"');await vscode.commands.executeCommand('workbench.action.files.save');
      saved=analyze(await fs.readFile(file));assert.equal(saved.sheets[0].tables[0].calculatedColumns[2],null);assert.equal(saved.sheets[0].cells.get('C3')?.formula??null,null);
      await vscode.commands.executeCommand('jinsheet.undo');await wait('document.querySelector("#cell-2-2")?.textContent==="60"');await vscode.commands.executeCommand('jinsheet.undo');await wait('!document.querySelector("#clearCalculatedColumn").disabled');await vscode.commands.executeCommand('workbench.action.files.save');assert.equal(analyze(await fs.readFile(file)).sheets[0].tables[0].calculatedColumns[2],'B2*3');
    } finally {devtools?.close();await vscode.commands.executeCommand('workbench.action.closeAllEditors');await fs.rm(directory,{recursive:true,force:true});}
  });

  test('Chromium preserves native table totals while filtering, inserting rows and resizing through save and undo', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-native-table-totals-')), file = path.join(directory, 'totals.xlsx');
    const sample = analyze(await fs.readFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx'))), files = { ...sample.files }, sheet = sample.sheets[0];
    const rows = '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Alpha</t></is></c><c r="B2"><v>10</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>Beta</t></is></c><c r="B3"><v>20</v></c></row><row r="4"><c r="A4" t="inlineStr"><is><t>Total</t></is></c><c r="B4"><f>SUBTOTAL(109,[Amount])</f><v>30</v></c><c r="C4"><f>B4</f><v>30</v></c></row><row r="5"><c r="B5"><v>7</v></c></row></sheetData>';
    files[sheet.path] = new TextEncoder().encode(sheet.xml.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ').replace(/<sheetData>[\s\S]*?<\/sheetData>/, rows).replace('</worksheet>', '<tableParts count="1"><tablePart r:id="t1"/></tableParts></worksheet>'));
    files[path.posix.join(path.posix.dirname(sheet.path), '_rels', path.posix.basename(sheet.path) + '.rels')] = new TextEncoder().encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="t1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>');
    files['xl/tables/table1.xml'] = new TextEncoder().encode('<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:B4" totalsRowCount="1"><autoFilter ref="A1:B3"/><tableColumns count="2"><tableColumn id="1" name="Name" totalsRowLabel="Total"/><tableColumn id="2" name="Amount" totalsRowFunction="sum"/></tableColumns></table>');
    files['[Content_Types].xml'] = new TextEncoder().encode(new TextDecoder().decode(files['[Content_Types].xml']).replace('</Types>', '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>'));
    const bytes = zipSync(files), model = analyze(bytes); assert.equal(model.writable, true, model.readOnlyReason); await fs.writeFile(file, bytes); let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor'); devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const result = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
        return result.result?.value;
      };
      const wait = async expression => { const deadline = Date.now() + 6000; while (Date.now() < deadline) { if (await evaluate(expression)) return; await new Promise(resolve => setTimeout(resolve, 25)); } assert.fail('Table totals condition failed: ' + expression); };
      const actOnCell = async (row, column, script = '') => wait('(()=>{const cell=document.querySelector('+JSON.stringify('#cell-'+row+'-'+column)+');if(!cell)return false;cell.click();'+script+';return true})()');
      await wait('document.querySelector("#cell-3-1")?.textContent==="30"');
      await actOnCell(3,1,"document.querySelector(\"#filterValuesButton\").click()"); await wait('!document.querySelector("#filterPanel").hidden');
      await evaluate('document.querySelector("#filterMode").value="conditions";document.querySelector("#filterMode").dispatchEvent(new Event("change"));document.querySelector("#filterOperator1").value="greaterThan";document.querySelector("#filterConditionValue1").value="15";document.querySelector("#filterApply").click()');
      await wait('document.querySelector("#filterPanel").hidden&&!document.querySelector("#cell-1-1")&&document.querySelector("#cell-3-1")?.textContent==="20"&&document.querySelector("#cell-3-0")?.textContent==="Total"');
      await vscode.commands.executeCommand('workbench.action.files.save'); let saved = analyze(await fs.readFile(file)); assert.deepEqual(saved.sheets[0].nativeHiddenRows, [1]); assert.equal(saved.sheets[0].autoFilter, null);
      devtools.close(); devtools = null; await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor'); devtools = await waitForSpreadsheetWebview(); await wait('document.querySelector("#cell-3-1")?.textContent==="20"');
      await actOnCell(3,1,"document.querySelector(\"#filterValuesButton\").click()"); await wait('!document.querySelector("#filterPanel").hidden'); await evaluate('document.querySelector("#filterClear").click()');
      await wait('document.querySelector("#filterPanel").hidden&&document.querySelector("#cell-3-1")?.textContent==="30"&&Boolean(document.querySelector("#cell-1-1"))');
      await actOnCell(3,1,"document.querySelector(\"#insertRow\").click()");
      await wait('document.querySelector("#cell-4-1")?.textContent==="30"&&document.querySelector("#cell-4-0")?.textContent==="Total"'); await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file)); assert.equal(saved.sheets[0].tables[0].range.bottom, 4); assert.equal(saved.sheets[0].tables[0].autoFilter.bottom, 3); assert.equal(saved.sheets[0].cells.get('B5').formula, 'SUBTOTAL(109,[Amount])'); assert.deepEqual(saved.sheets[0].nativeHiddenRows, []);
      await vscode.commands.executeCommand('jinsheet.undo'); await wait('document.querySelector("#cell-3-1")?.textContent==="30"&&document.querySelector("#cell-3-0")?.textContent==="Total"'); await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file)); assert.equal(saved.sheets[0].tables[0].range.bottom, 3); assert.equal(saved.sheets[0].tables[0].autoFilter.bottom, 2);
      const selectTable = async row => {
        await wait('Boolean(document.querySelector("#cell-' + row + '-1"))');
        await wait('(()=>{const first=document.querySelector("#cell-0-0"),last=document.querySelector("#cell-' + row + '-1");if(!first||!last)return false;const box=first.getBoundingClientRect();first.dispatchEvent(new MouseEvent("mousedown",{button:0,bubbles:true,clientX:box.left+2,clientY:box.top+2}));last.dispatchEvent(new MouseEvent("mouseenter",{bubbles:false}));document.dispatchEvent(new MouseEvent("mouseup",{bubbles:true}));return true})()');
        await wait('!document.querySelector("#resizeTable").disabled'); await evaluate('document.querySelector("#resizeTable").click()');
      };
      await selectTable(5); await wait('document.querySelector("#cell-5-1")?.textContent==="37"&&document.querySelector("#cell-5-0")?.textContent==="Total"&&document.querySelector("#cell-3-1")?.textContent===""&&document.querySelector("#cell-3-2")?.textContent==="37"');
      await vscode.commands.executeCommand('workbench.action.files.save'); saved = analyze(await fs.readFile(file)); assert.equal(saved.sheets[0].tables[0].autoFilter.bottom, 4); assert.equal(saved.sheets[0].cells.get('C4').formula, 'B6');
      await vscode.commands.executeCommand('jinsheet.undo'); await wait('document.querySelector("#cell-3-1")?.textContent==="30"&&document.querySelector("#cell-3-2")?.textContent==="30"'); await vscode.commands.executeCommand('workbench.action.files.save');
      await vscode.commands.executeCommand('jinsheet.redo'); await wait('document.querySelector("#cell-5-1")?.textContent==="37"'); await vscode.commands.executeCommand('workbench.action.files.save');
      devtools.close(); devtools = null; await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor'); devtools = await waitForSpreadsheetWebview(); await wait('document.querySelector("#cell-5-1")?.textContent==="37"');
      await selectTable(3); await wait('document.querySelector("#cell-3-1")?.textContent==="30"&&document.querySelector("#cell-3-2")?.textContent==="30"&&document.querySelector("#cell-4-1")?.textContent==="7"'); await vscode.commands.executeCommand('workbench.action.files.save');
      const shrunkBytes = await fs.readFile(file); saved = analyze(shrunkBytes); assert.equal(saved.sheets[0].tables[0].range.bottom, 3); assert.equal(saved.sheets[0].cells.get('C4').formula, 'B4');
      await selectTable(2); await wait('document.querySelector("#notice").textContent.includes("contains data")'); assert.equal(await evaluate('document.querySelector("#cell-3-1").textContent'), '30'); await vscode.commands.executeCommand('workbench.action.files.save'); assert.equal((await fs.readFile(file)).equals(shrunkBytes), true);
      await actOnCell(3,1,""); await wait('document.querySelector("#tableTotalsFunction").value==="sum"&&!document.querySelector("#tableTotals").disabled');
      await evaluate('document.querySelector("#tableTotalsFunction").value="average";document.querySelector("#tableTotalsFunction").dispatchEvent(new Event("change"))'); assert.equal(await evaluate('document.querySelector("#cell-3-1").textContent'), '30');
      await evaluate('document.querySelector("#applyTableTotalFunction").click()'); await wait('document.querySelector("#cell-3-1")?.textContent==="15"&&!document.querySelector("#tableTotals").disabled'); await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file)); assert.equal(saved.sheets[0].tables[0].totalColumns[1].function, 'average');
      await vscode.commands.executeCommand('jinsheet.undo'); await wait('document.querySelector("#cell-3-1")?.textContent==="30"'); await vscode.commands.executeCommand('workbench.action.files.save');
      await vscode.commands.executeCommand('jinsheet.redo'); await wait('document.querySelector("#cell-3-1")?.textContent==="15"'); await vscode.commands.executeCommand('workbench.action.files.save');
      await evaluate('document.querySelector("#tableTotals").click()'); await wait('document.querySelector("#tableTotals").getAttribute("aria-pressed")==="false"&&!document.querySelector("#tableTotals").disabled&&document.querySelector("#cell-3-1")?.textContent===""');
      const totalsSaveDeadline = Date.now() + 5000;
      let totalsSaved;
      do {
        await vscode.commands.executeCommand('workbench.action.files.save');
        totalsSaved = analyze(await fs.readFile(file));
        if (totalsSaved.sheets[0].tables[0].totalsRowCount === 0) break;
        await new Promise(resolve => setTimeout(resolve, 25));
      } while (Date.now() < totalsSaveDeadline);
      assert.equal(totalsSaved.sheets[0].tables[0].totalsRowCount, 0);
      await evaluate('document.querySelector("#tableTotals").click()'); await wait('document.querySelector("#cell-3-1")?.textContent==="15"&&document.querySelector("#tableTotalsFunction").value==="average"');
      const averageSaveDeadline = Date.now() + 8000; let averageSaved;
      do { await vscode.commands.executeCommand('workbench.action.files.save'); averageSaved = analyze(await fs.readFile(file)); if (averageSaved.sheets[0].tables[0].totalColumns[1].function === 'average') break; await new Promise(resolve => setTimeout(resolve, 25)); } while (Date.now() < averageSaveDeadline);
      assert.equal(averageSaved.sheets[0].tables[0].totalColumns[1].function, 'average');
      devtools.close(); devtools = null; await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor'); devtools = await waitForSpreadsheetWebview(); await wait('document.querySelector("#cell-3-1")?.textContent==="15"');
      await actOnCell(3,1,"document.querySelector(\"#tableTotalsFunction\").value=\"count\";document.querySelector(\"#tableTotalsFunction\").dispatchEvent(new Event(\"change\"));document.querySelector(\"#applyTableTotalFunction\").click()");
      await wait('document.querySelector("#cell-3-1")?.textContent==="2"&&!document.querySelector("#tableTotals").disabled'); await vscode.commands.executeCommand('workbench.action.files.save'); assert.equal(analyze(await fs.readFile(file)).sheets[0].tables[0].totalColumns[1].function, 'count');
      await evaluate('document.querySelector("#tableTotalsFunction").value="none";document.querySelector("#tableTotalsFunction").dispatchEvent(new Event("change"));document.querySelector("#applyTableTotalFunction").click()'); await wait('document.querySelector("#cell-3-1")?.textContent===""&&!document.querySelector("#tableTotals").disabled'); await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(analyze(await fs.readFile(file)).sheets[0].tables[0].totalColumns[1].function, 'none');
      await actOnCell(3,0,'document.querySelector("#tableTotalsFunction").value="none";document.querySelector("#tableTotalsFunction").dispatchEvent(new Event("change"));document.querySelector("#applyTableTotalFunction").click()'); await wait('document.querySelector("#cell-3-0")?.textContent===""&&!document.querySelector("#tableTotals").disabled');
      await evaluate('document.querySelector("#tableTotals").click()'); await wait('document.querySelector("#tableTotals").getAttribute("aria-pressed")==="false"&&!document.querySelector("#tableTotals").disabled');
      await evaluate('document.querySelector("#tableTotals").click()'); await wait('document.querySelector("#tableTotals").getAttribute("aria-pressed")==="true"&&!document.querySelector("#tableTotals").disabled&&document.querySelector("#cell-3-0")?.textContent===""&&document.querySelector("#cell-3-1")?.textContent===""');

      await evaluate('document.querySelector("#tableTotals").click()'); await wait('document.querySelector("#tableTotals").getAttribute("aria-pressed")==="false"&&!document.querySelector("#tableTotals").disabled'); await vscode.commands.executeCommand('workbench.action.files.save'); const beforeConflict = await fs.readFile(file);
      // Keep this outside-table collision scenario independent of automatic expansion.
      await evaluate('document.querySelector("#autoExpandTables").checked=false');
      await actOnCell(3,1,"const input=document.querySelector(\"#formula\");input.value=\"42\";input.dispatchEvent(new Event(\"input\",{bubbles:true}));input.dispatchEvent(new KeyboardEvent(\"keydown\",{key:\"Enter\",bubbles:true}))"); await wait('document.querySelector("#cell-3-1")?.textContent==="42"');
      await actOnCell(2,1,"document.querySelector(\"#tableTotals\").click()"); await wait('document.querySelector("#notice").textContent.includes("contains data")&&!document.querySelector("#tableTotals").disabled&&document.querySelector("#tableTotals").getAttribute("aria-pressed")==="false"');
      assert.equal((await fs.readFile(file)).equals(beforeConflict), true); await vscode.commands.executeCommand('workbench.action.files.save'); saved = analyze(await fs.readFile(file)); assert.equal(saved.sheets[0].tables[0].totalsRowCount, 0); assert.equal(saved.sheets[0].cells.get('B4').value, '42');



    } finally { devtools?.close(); await vscode.commands.executeCommand('workbench.action.closeAllEditors'); await fs.rm(directory, { recursive: true, force: true }); }
  });

  test('Chromium uses native XLSX default dimensions and preserves 1px and 1785px columns through save and undo', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-native-dimensions-host-')), file = path.join(directory, 'dimensions.xlsx');
    const sample = analyze(await fs.readFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx'))), files = { ...sample.files }, sheet = sample.sheets[0];
    const rows = '<sheetData><row r="1">' + ['A', 'B', 'C', 'D'].map(column => '<c r="' + column + '1" t="inlineStr"><is><t>x</t></is></c>').join('') + '</row></sheetData>';
    files[sheet.path] = new TextEncoder().encode(sheet.xml.replace(/<sheetFormatPr\b[^>]*\/>/, '').replace(/<cols>[\s\S]*?<\/cols>/, '').replace(/<sheetData>[\s\S]*?<\/sheetData>/, rows));
    const model = patchColumnWidths(analyze(zipSync(files)), [{ sheetIndex: 0, column: 1, width: .14 }, { sheetIndex: 0, column: 2, width: 255 }, { sheetIndex: 0, column: 3, width: 0 }]).model;
    await fs.writeFile(file, model.bytes); let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor'); devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const result = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
        return result.result?.value;
      };
      const wait = async expression => { const deadline = Date.now() + 5000; while (Date.now() < deadline) { if (await evaluate(expression)) return; await new Promise(resolve => setTimeout(resolve, 25)); } assert.fail('Native dimension condition failed: ' + expression); };
      await wait('Boolean(document.querySelector("#cell-0-2"))');
      await wait('document.querySelector("#cell-0-2")?.getBoundingClientRect().width===1785');
      const sizes = await evaluate('(()=>{const cells=[0,1,2].map(column=>document.querySelector("#cell-0-"+column).getBoundingClientRect());return{widths:cells.map(box=>box.width),height:cells[0].height,positions:cells.map(box=>[box.left,box.right]),adjacent:Math.abs(cells[1].left-cells[0].right)<.1&&Math.abs(cells[2].left-cells[1].right)<.1,hidden:!document.querySelector("#cell-0-3")||document.querySelector("#cell-0-3").getBoundingClientRect().width===0,dpr:devicePixelRatio,maximumDigitWidth:window.__jinsheetModel?.maximumDigitWidth,modelDefaultWidth:window.__jinsheetModel?.defaultColumnWidth,modelZoom:window.__jinsheetModel?.zoomScale,canvasZoom:getComputedStyle(document.querySelector("#canvas")).zoom}})()');
      assert.deepEqual(sizes.widths, [64, 1, 1785], JSON.stringify(sizes)); assert.equal(sizes.height, 20); assert.equal(sizes.adjacent, true, JSON.stringify(sizes)); assert.equal(sizes.hidden, true);
      assert.equal(await evaluate('(()=>{const cell=document.querySelector("#cell-0-0");return cell.scrollHeight<=cell.clientHeight+1})()'), true);
      await evaluate('document.querySelector("#cell-0-1").click();document.querySelector("#cell-0-1").dispatchEvent(new KeyboardEvent("keydown",{key:"F2",bubbles:true}))');
      await wait('Boolean(document.querySelector(".cell-editor"))');
      const editor = await evaluate('(()=>{const input=document.querySelector(".cell-editor"),box=input.getBoundingClientRect();return{width:box.width,visible:document.elementFromPoint(box.left+10,box.top+4)===input}})()');
      assert.ok(editor.width >= 24); assert.equal(editor.visible, true);
      await evaluate('document.querySelector(".cell-editor").dispatchEvent(new KeyboardEvent("keydown",{key:"Escape",bubbles:true}))');
      await wait('Boolean(document.querySelector("#cell-0-2"))&&!document.querySelector(".cell-editor")');
      const projection = { title: 'Native cell widths', rows: [0], columns: [0, 1, 2], widths: { 0: 64, 1: 1, 2: 1785 }, heights: { 0: 20 }, cells: new Map([0, 1, 2].map(column => ['0:' + column, { displayValue: 'x' }])), settings: { ...DEFAULT_PRINT_SETTINGS, fitToPage: true, fitToWidth: 1, fitToHeight: 1 } };
      const print = await measurePrintPages(evaluate, renderPrintDocument(projection), true);
      assert.equal(print.length, 1); for (const [index, width] of [64, 1, 1785].entries()) assert.ok(Math.abs(print[0].cellWidths[index] - width) < .25, JSON.stringify(print[0].cellWidths));
      assert.ok(print[0].cellHeights.every(height => height <= 21));
      await evaluate('document.querySelector(\'.colhead[data-column="2"] .col-resize-handle\').dispatchEvent(new KeyboardEvent("keydown",{key:"ArrowLeft",bubbles:true}))');
      await wait('document.querySelector("#cell-0-2")?.style.width==="1775px"'); await vscode.commands.executeCommand('workbench.action.files.save');
      const saved = analyze(await fs.readFile(file)); assert.equal(saved.sheets[0].columnWidths[2], 253.57); assert.equal(saved.sheets[0].defaultColumnWidth, 64); assert.equal(saved.sheets[0].defaultRowHeight, 20); assert.deepEqual(saved.sheets[0].hiddenColumns, [3]);
      await vscode.commands.executeCommand('jinsheet.undo'); await vscode.commands.executeCommand('workbench.action.files.save');
      const restored = analyze(await fs.readFile(file)); assert.deepEqual(Object.keys(restored.files).sort(), Object.keys(model.files).sort());
      for (const [name, bytes] of Object.entries(model.files)) assert.deepEqual(restored.files[name], bytes, name);
    } finally { devtools?.close(); await vscode.commands.executeCommand('workbench.action.closeAllEditors'); await fs.rm(directory, { recursive: true, force: true }); }
  });

  test('Chromium Webview displays merged cells as one spanning cell and preserves them after editing', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-merged-cells-host-'));
    const file = path.join(directory, 'merged-cells.xlsx'), encode = text => new TextEncoder().encode(text);
    const mergeXml = '<mergeCells count="1"><mergeCell ref="A1:B2"/></mergeCells>';
    await fs.writeFile(file, zipSync({
      '[Content_Types].xml': encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'),
      '_rels/.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
      'xl/workbook.xml': encode('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="r1"/></sheets></workbook>'),
      'xl/_rels/workbook.xml.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'),
      'xl/worksheets/sheet1.xml': encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Heading</t></is></c></row></sheetData>' + mergeXml + '</worksheet>')
    }));
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.exception?.description || response.exceptionDetails.text);
        return response.result?.value;
      };
      assert.equal(await evaluate('document.querySelector("#cell-0-0")?.textContent'), 'Heading');
      assert.equal(await evaluate('document.querySelector("#cell-0-1")'), null, 'covered cells are not separate grid cells');
      assert.equal(await evaluate('document.querySelector("#cell-1-0")'), null, 'the merged region covers both rows');
      assert.equal(await evaluate('getComputedStyle(document.querySelector("#cell-0-0")).width'), '128px');
      assert.equal(await evaluate('getComputedStyle(document.querySelector("#cell-0-0")).height'), '40px');
      await evaluate('document.querySelector("#cell-0-0").dispatchEvent(new KeyboardEvent("keydown",{key:"ArrowRight",bubbles:true}))');
      const navigationDeadline = Date.now() + 5000;
      while (Date.now() < navigationDeadline && !await evaluate('document.querySelector("#cell-0-2")?.classList.contains("selected")')) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('document.querySelector("#cell-0-2")?.classList.contains("selected")'), true, 'arrow navigation skips the rest of the merged area');
      await evaluate('document.querySelector("#cell-0-2").dispatchEvent(new KeyboardEvent("keydown",{key:"ArrowDown",bubbles:true}))');
      const belowMergeDeadline = Date.now() + 5000;
      while (Date.now() < belowMergeDeadline && !await evaluate('document.querySelector("#cell-1-2")?.classList.contains("selected")')) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('document.querySelector("#cell-1-2")?.classList.contains("selected")'), true, 'down arrow reaches the next row outside the merged area');
      await evaluate('document.querySelector("#cell-1-2").dispatchEvent(new KeyboardEvent("keydown",{key:"ArrowLeft",bubbles:true}))');
      const backToMergeDeadline = Date.now() + 5000;
      while (Date.now() < backToMergeDeadline && !await evaluate('document.querySelector("#cell-0-0")?.classList.contains("selected")')) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('document.querySelector("#cell-0-0")?.classList.contains("selected")'), true, 'left arrow selects the merged cell anchor instead of an invisible covered cell');
      await evaluate('document.querySelector("#cell-0-0").click();(()=>{const input=document.querySelector("#formula");input.value="Updated heading";input.dispatchEvent(new Event("input",{bubbles:true}));input.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))})()');
      const editDeadline = Date.now() + 5000;
      while (Date.now() < editDeadline && await evaluate('document.querySelector("#cell-0-0")?.textContent') !== 'Updated heading') await new Promise(resolve => setTimeout(resolve, 25));
      await vscode.commands.executeCommand('workbench.action.files.save');
      const saved = analyze(await fs.readFile(file));
      assert.equal(saved.writable, true, saved.readOnlyReason);
      assert.equal(saved.sheets[0].cells.get('A1').value, 'Updated heading');
      assert.ok(saved.sheets[0].xml.includes(mergeXml));
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview inserts a scatter chart and updates unsaved source edits through undo and redo', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-scatter-host-'));
    const file = path.join(directory, 'scatter.xlsx');
    await fs.copyFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Sample.xlsx'), file);
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const result = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
        return result.result?.value;
      };
      const waitFor = async expression => {
        const deadline = Date.now() + 5000;
        while (Date.now() < deadline) {
          if (await evaluate(expression)) return;
          await new Promise(resolve => setTimeout(resolve, 25));
        }
        assert.fail('Scatter UI did not reach expected state: ' + expression);
      };
      await waitFor('Boolean(document.querySelector("#cell-0-1")&&document.querySelector("#cell-2-2"))');
      const points = await evaluate(`['#cell-0-1','#cell-2-2'].map(selector=>{const r=document.querySelector(selector).getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+r.height/2}})`);
      await devtools.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: points[0].x + devtools.frameOffset.x, y: points[0].y + devtools.frameOffset.y, button: 'left', buttons: 1, clickCount: 1 });
      await devtools.send('Input.dispatchMouseEvent', { type: 'mouseMoved', x: points[1].x + devtools.frameOffset.x, y: points[1].y + devtools.frameOffset.y, buttons: 1 });
      await devtools.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: points[1].x + devtools.frameOffset.x, y: points[1].y + devtools.frameOffset.y, button: 'left', buttons: 0, clickCount: 1 });
      await evaluate(`(() => {document.querySelector('#chartType').value='scatter';document.querySelector('#insertChart').click()})()`);
      await waitFor('Boolean(document.querySelector(".chart-card svg circle"))');
      await vscode.commands.executeCommand('workbench.action.files.save');
      let saved = analyze(await fs.readFile(file)).sheets[0].charts[0];
      assert.equal(saved.type, 'scatter'); assert.deepEqual(saved.series[0].xValues, [12, 7]); assert.deepEqual(saved.series[0].values, [1500, 2200]);
      await vscode.commands.executeCommand('jinsheet.undo');
      await waitFor('document.querySelectorAll(".chart-card").length===0');
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(analyze(await fs.readFile(file)).sheets[0].charts.length, 0);
      await vscode.commands.executeCommand('jinsheet.redo');
      await waitFor('Boolean(document.querySelector(".chart-card svg circle"))');
      await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file)).sheets[0].charts[0];
      assert.equal(saved.type, 'scatter'); assert.deepEqual(saved.series[0].xValues, [12, 7]);
      await waitFor('(()=>{const cell=document.querySelector("#cell-1-1");if(!cell)return false;cell.scrollIntoView({block:"nearest",inline:"nearest"});return true})()');
      await evaluate('new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)))');
      let sourcePoint; const sourceDeadline=Date.now()+5000;
      while(Date.now()<sourceDeadline){sourcePoint=await evaluate('(()=>{const cell=document.querySelector("#cell-1-1"),circle=document.querySelector(".chart-card svg circle");if(!cell||!circle)return null;const r=cell.getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+r.height/2,previousX:Number(circle.getAttribute("cx"))}})()');if(sourcePoint)break;await new Promise(resolve=>setTimeout(resolve,25));}
      assert.ok(sourcePoint,'wait for current source cell and chart after saving');const previousX=sourcePoint.previousX;
      await devtools.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: sourcePoint.x + devtools.frameOffset.x, y: sourcePoint.y + devtools.frameOffset.y, button: 'left', buttons: 1, clickCount: 1 });
      await devtools.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: sourcePoint.x + devtools.frameOffset.x, y: sourcePoint.y + devtools.frameOffset.y, button: 'left', buttons: 0, clickCount: 1 });
      await waitFor('document.querySelector("#address").value==="B2"');
      await evaluate(`(() => {const input=document.querySelector('#formula');input.value='5';input.dispatchEvent(new Event('input',{bubbles:true}));input.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}))})()`);
      await waitFor('Number(document.querySelector(".chart-card svg circle")?.getAttribute("cx"))<' + previousX);
      assert.deepEqual(analyze(await fs.readFile(file)).sheets[0].charts[0].series[0].xValues, [12, 7], 'chart preview updates before saving');
      await vscode.commands.executeCommand('jinsheet.undo');
      await waitFor('Number(document.querySelector(".chart-card svg circle")?.getAttribute("cx"))===' + previousX);
      await vscode.commands.executeCommand('jinsheet.redo');
      await waitFor('Number(document.querySelector(".chart-card svg circle")?.getAttribute("cx"))<' + previousX);
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.deepEqual(analyze(await fs.readFile(file)).sheets[0].charts[0].series[0].xValues, [5, 7]);
    } finally {
      devtools?.close(); await vscode.commands.executeCommand('workbench.action.closeAllEditors'); await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview omits filtered rows from the chart preview', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-filtered-chart-host-'));
    const file = path.join(directory, 'filtered-chart.xlsx');
    let model = analyze(await fs.readFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Sample.xlsx')));
    model = patchAddChart(model, 0, { top: 0, bottom: 2, left: 0, right: 1 }).model;
    model = patchAutoFilters(model, [{ sheetIndex: 0, filter: { top: 0, bottom: 2, left: 0, right: 2, columns: [{ column: 0, values: ['사과'], blanks: false }] } }]).model;
    assert.deepEqual(model.sheets[0].nativeHiddenRows, [2]);
    await fs.writeFile(file, model.bytes);
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const result = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
        return result.result?.value;
      };
      const blueBarCount = 'document.querySelectorAll(' + JSON.stringify('.chart-card svg rect[fill="#4f81bd"]') + ').length';
      const deadline = Date.now() + 6000;
      while (Date.now() < deadline && !await evaluate('Boolean(document.querySelector(".chart-card svg"))')) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('Boolean(document.querySelector(".chart-card svg"))'), true, 'the chart preview is visible');
      assert.equal(await evaluate(blueBarCount), 2, 'one visible bar and its legend marker are rendered; the filtered point is omitted');
      await evaluate('(()=>{const input=document.querySelector("#formula");document.querySelector("#cell-1-1").click();input.value="15";input.dispatchEvent(new Event("input",{bubbles:true}));input.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))})()');
      const editDeadline = Date.now() + 6000;
      while (Date.now() < editDeadline && await evaluate(blueBarCount) !== 2) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate(blueBarCount), 2, 'an unsaved visible point edit does not bring the filtered point back');
      assert.deepEqual(analyze(await fs.readFile(file)).sheets[0].nativeHiddenRows, [2], 'the preview update does not require saving filter state');
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview edits and deletes charts through save, undo and redo', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-chart-details-host-'));
    const file = path.join(directory, 'chart-details.xlsx');
    const source = analyze(await fs.readFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Sample.xlsx')));
    const original = patchAddChart(source, 0, { top: 0, bottom: 2, left: 0, right: 1 });
    await fs.writeFile(file, original.bytes);
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const result = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (result.exceptionDetails) throw new Error(result.exceptionDetails.text);
        return result.result?.value;
      };
      const waitForTitle = async title => {
        const deadline = Date.now() + 5000;
        while (Date.now() < deadline) {
          if (await evaluate('document.querySelector(".chart-card figcaption")?.textContent') === title) return;
          await new Promise(resolve => setTimeout(resolve, 25));
        }
        assert.fail('Chart title did not update to ' + title);
      };
      await waitForTitle('Chart');
      await evaluate(`(() => {document.querySelector('[data-chart-title]').value='재고 <현황> & 수량';document.querySelector('[data-chart-series-name]').value='재고 수량';document.querySelector('[data-chart-edit]').click()})()`);
      await waitForTitle('재고 <현황> & 수량');
      await vscode.commands.executeCommand('workbench.action.files.save');
      const saved = analyze(await fs.readFile(file));
      assert.equal(saved.sheets[0].charts[0].title, '재고 <현황> & 수량');
      assert.equal(saved.sheets[0].charts[0].series[0].name, '재고 수량');
      assert.equal(saved.sheets[0].cells.get('B1').value, '수량');
      await vscode.commands.executeCommand('jinsheet.undo');
      await waitForTitle('Chart');
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(analyze(await fs.readFile(file)).sheets[0].charts[0].series[0].name, '수량');
      await vscode.commands.executeCommand('jinsheet.redo');
      await waitForTitle('재고 <현황> & 수량');
      const redoSaveDeadline = Date.now() + 5000;
      let redoneChart;
      do {
        await vscode.commands.executeCommand('workbench.action.files.save');
        redoneChart = analyze(await fs.readFile(file));
        if (redoneChart.sheets[0].charts[0].series[0].name === '재고 수량') break;
        await new Promise(resolve => setTimeout(resolve, 25));
      } while (Date.now() < redoSaveDeadline);
      assert.equal(redoneChart.sheets[0].charts[0].series[0].name, '재고 수량');
      const cellDeadline = Date.now() + 5000; let points;
      while (Date.now() < cellDeadline) {
        points = await evaluate('(()=>{const cells=[document.querySelector("#cell-0-0"),document.querySelector("#cell-2-2")];if(cells.some(cell=>!cell))return null;return cells.map(cell=>{const r=cell.getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+r.height/2}})})()');
        if (points) break;
        await new Promise(resolve => setTimeout(resolve, 25));
      }
      assert.equal(points?.length, 2, 'wait for current grid range geometry after saving');
      await devtools.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: points[0].x + devtools.frameOffset.x, y: points[0].y + devtools.frameOffset.y, button: 'left', buttons: 1, clickCount: 1 });
      await devtools.send('Input.dispatchMouseEvent', { type: 'mouseMoved', x: points[1].x + devtools.frameOffset.x, y: points[1].y + devtools.frameOffset.y, buttons: 1 });
      await devtools.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: points[1].x + devtools.frameOffset.x, y: points[1].y + devtools.frameOffset.y, button: 'left', buttons: 0, clickCount: 1 });
      await evaluate('const address=document.querySelector("#address");address.value="A1:C3";address.dispatchEvent(new Event("input",{bubbles:true}));address.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))');
      const sourceEnableDeadline=Date.now()+5000;while(Date.now()<sourceEnableDeadline&&await evaluate('document.querySelectorAll(".cell.in-range").length<9'))await new Promise(resolve=>setTimeout(resolve,25));
      assert.equal(await evaluate('document.querySelectorAll(".cell.in-range").length>=9'),true);
      assert.equal(await evaluate('document.querySelector("[data-chart-source]").disabled'),false);
      await evaluate('document.querySelector("[data-chart-source]").click()');
      const waitForSeries = async count => {
        const deadline = Date.now() + 5000;
        while (Date.now() < deadline) {
          if (await evaluate('document.querySelectorAll("[data-chart-series-name]").length') === count) return;
          await new Promise(resolve => setTimeout(resolve, 25));
        }
        assert.fail('Chart series did not update to ' + count);
      };
      const saveUntil = async predicate => {
        const deadline = Date.now() + 8000;
        let saved;
        do {
          await vscode.commands.executeCommand('workbench.action.files.save');
          saved = analyze(await fs.readFile(file));
          if (predicate(saved)) return saved;
          await new Promise(resolve => setTimeout(resolve, 25));
        } while (Date.now() < deadline);
        assert.fail('Chart save did not reach the expected state');
      };
      await waitForSeries(2);
      await saveUntil(saved => JSON.stringify(saved.sheets[0].charts[0].series.map(item => item.values)) === JSON.stringify([[12, 7], [1500, 2200]]));
      assert.deepEqual(analyze(await fs.readFile(file)).sheets[0].charts[0].series.map(item => item.values), [[12, 7], [1500, 2200]]);
      await vscode.commands.executeCommand('jinsheet.undo');
      await waitForSeries(1);
      await saveUntil(saved => saved.sheets[0].charts[0].series[0].name === '재고 수량');
      assert.equal(analyze(await fs.readFile(file)).sheets[0].charts[0].series[0].name, '재고 수량');
      await vscode.commands.executeCommand('jinsheet.redo');
      await waitForSeries(2);
      await saveUntil(saved => saved.sheets[0].charts[0].series[1].name === '단가');
      assert.equal(analyze(await fs.readFile(file)).sheets[0].charts[0].series[1].name, '단가');
      await evaluate('document.querySelector("[data-chart-delete]").click()');
      await waitForSeries(0);
      await saveUntil(saved => !saved.hasCharts);
      assert.equal(analyze(await fs.readFile(file)).hasCharts, false);
      await vscode.commands.executeCommand('jinsheet.undo');
      await waitForSeries(2);
      await saveUntil(saved => saved.sheets[0].charts[0].series[1].name === '단가');
      assert.equal(analyze(await fs.readFile(file)).sheets[0].charts[0].series[1].name, '단가');
      await vscode.commands.executeCommand('jinsheet.redo');
      await waitForSeries(0);
      await saveUntil(saved => !saved.hasCharts);
      assert.equal(analyze(await fs.readFile(file)).hasCharts, false);
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview applies and clears an XLSX fill color while preserving other styles', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-fill-format-host-'));
    const file = path.join(directory, 'fill-format.xlsx');
    await fs.copyFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx'), file);
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.text);
        return response.result?.value;
      };
      const original = analyze(await fs.readFile(file)).sheets[0].cells.get('C2');
      const saveUntil = async predicate => {
        const deadline = Date.now() + 8000;
        let cell;
        do {
          await vscode.commands.executeCommand('workbench.action.files.save');
          cell = analyze(await fs.readFile(file)).sheets[0].cells.get('C2');
          if (predicate(cell)) return cell;
          await new Promise(resolve => setTimeout(resolve, 25));
        } while (Date.now() < deadline);
        assert.fail('Fill save did not reach the expected state');
      };
      const cellPoint = await evaluate(`(() => {const r=document.querySelector('#cell-1-2').getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+r.height/2}})()`);
      await devtools.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: cellPoint.x + devtools.frameOffset.x, y: cellPoint.y + devtools.frameOffset.y, button: 'left', buttons: 1, clickCount: 1 });
      await devtools.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: cellPoint.x + devtools.frameOffset.x, y: cellPoint.y + devtools.frameOffset.y, button: 'left', buttons: 0, clickCount: 1 });
      assert.equal(await evaluate('document.querySelector("#address")?.value'), 'C2');
      await evaluate(`(() => {const input=document.querySelector('#formatFillColor');input.value='#FFCC00';input.dispatchEvent(new Event('change',{bubbles:true}))})()`);
      const modifiedDeadline = Date.now() + 3000;
      while (Date.now() < modifiedDeadline && await evaluate('document.querySelector("#notice")?.textContent.includes("Modified")') !== true) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('document.querySelector("#notice")?.textContent.includes("Modified")'), true);
      let saved = await saveUntil(cell => cell.styleProperties.fill === '#FFCC00');
      assert.equal(saved.styleProperties.fill, '#FFCC00');
      assert.equal(saved.formatCode, original.formatCode);
      assert.equal(saved.styleProperties.bold, original.styleProperties.bold);
      assert.deepEqual(saved.styleProperties.borders, original.styleProperties.borders);
      await evaluate('document.querySelector("#clearFillColor").click()');
      const clearDeadline = Date.now() + 3000;
      while (Date.now() < clearDeadline && await evaluate('document.querySelector("#notice")?.textContent.includes("Modified")') !== true) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('document.querySelector("#notice")?.textContent.includes("Modified")'), true, 'clearing fill marks the workbook modified');
      saved = await saveUntil(cell => !cell.styleProperties.fill);
      assert.equal(saved.styleProperties.fill, undefined);
      assert.equal(saved.formatCode, original.formatCode);
      assert.deepEqual(saved.styleProperties.borders, original.styleProperties.borders);
      await vscode.commands.executeCommand('jinsheet.undo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file)).sheets[0].cells.get('C2');
      assert.equal(saved.styleProperties.fill, '#FFCC00', 'undo restores the selected fill color');
      await vscode.commands.executeCommand('jinsheet.undo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file)).sheets[0].cells.get('C2');
      assert.equal(saved.styleProperties.fill, original.styleProperties.fill, 'undo restores the original fill');
      await vscode.commands.executeCommand('jinsheet.redo');
      await vscode.commands.executeCommand('jinsheet.redo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file)).sheets[0].cells.get('C2');
      assert.equal(saved.styleProperties.fill, undefined, 'redo reapplies fill and clear operations');
      assert.deepEqual(saved.styleProperties.borders, original.styleProperties.borders);
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview applies and clears a font color while preserving other styles', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-font-color-host-'));
    const file = path.join(directory, 'font-color.xlsx');
    await fs.copyFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx'), file);
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.text);
        return response.result?.value;
      };
      const original = analyze(await fs.readFile(file)).sheets[0].cells.get('C2');
      const saveUntil = async predicate => {
        const deadline = Date.now() + 3000;
        do {
          await vscode.commands.executeCommand('workbench.action.files.save');
          const cell = analyze(await fs.readFile(file)).sheets[0].cells.get('C2');
          if (predicate(cell)) return cell;
          await new Promise(resolve => setTimeout(resolve, 25));
        } while (Date.now() < deadline);
        assert.fail('Font-color save did not reach the expected state');
      };
      const point = await evaluate(`(() => {const r=document.querySelector('#cell-1-2').getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+r.height/2}})()`);
      await devtools.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: point.x + devtools.frameOffset.x, y: point.y + devtools.frameOffset.y, button: 'left', buttons: 1, clickCount: 1 });
      await devtools.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: point.x + devtools.frameOffset.x, y: point.y + devtools.frameOffset.y, button: 'left', buttons: 0, clickCount: 1 });
      await evaluate(`(() => {const input=document.querySelector('#fontColor');input.value='#0055CC';input.dispatchEvent(new Event('change',{bubbles:true}))})()`);
      const deadline = Date.now() + 3000;
      while (Date.now() < deadline && await evaluate('(()=>{const cell=document.querySelector("#cell-1-2");return cell?getComputedStyle(cell).color:"missing"})()') !== 'rgb(0, 85, 204)') await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('(()=>{const cell=document.querySelector("#cell-1-2");return cell?getComputedStyle(cell).color:"missing"})()'), 'rgb(0, 85, 204)');
      let saved = await saveUntil(cell => cell.styleProperties.color === '#0055CC');
      assert.equal(saved.styleProperties.color, '#0055CC');
      assert.equal(saved.styleProperties.fill, original.styleProperties.fill);
      assert.deepEqual(saved.styleProperties.borders, original.styleProperties.borders);
      await evaluate('document.querySelector("#clearFontColor").click()');
      const clearDeadline = Date.now() + 3000;
      while (Date.now() < clearDeadline && await evaluate('document.querySelector("#notice")?.textContent.includes("Modified")') !== true) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('document.querySelector("#notice")?.textContent.includes("Modified")'), true, 'clearing font color marks the workbook modified');
      saved = await saveUntil(cell => cell.styleProperties.color === undefined);
      assert.equal(saved.styleProperties.color, undefined);
      assert.equal(saved.styleProperties.fill, original.styleProperties.fill);
      await vscode.commands.executeCommand('jinsheet.undo');
      saved = await saveUntil(cell => cell.styleProperties.color === '#0055CC');
      assert.equal(saved.styleProperties.color, '#0055CC');
      await vscode.commands.executeCommand('jinsheet.redo');
      saved = await saveUntil(cell => cell.styleProperties.color === undefined);
      assert.equal(saved.styleProperties.color, undefined);
      assert.deepEqual(saved.styleProperties.borders, original.styleProperties.borders);
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview applies and removes XLSX cell borders while preserving other styles', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-border-host-'));
    const file = path.join(directory, 'border.xlsx');
    await fs.copyFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx'), file);
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.text);
        return response.result?.value;
      };
      const cellDeadline = Date.now() + 3000;
      while (Date.now() < cellDeadline && !await evaluate('Boolean(document.querySelector("#cell-1-2"))')) await new Promise(resolve => setTimeout(resolve, 25));
      const original = analyze(await fs.readFile(file)).sheets[0].cells.get('C2');
      const point = await evaluate('(() => {const r=document.querySelector("#cell-1-2").getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+r.height/2}})()');
      await devtools.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: point.x + devtools.frameOffset.x, y: point.y + devtools.frameOffset.y, button: 'left', buttons: 1, clickCount: 1 });
      await devtools.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: point.x + devtools.frameOffset.x, y: point.y + devtools.frameOffset.y, button: 'left', buttons: 0, clickCount: 1 });
      await evaluate('document.querySelector("#borderPreset").value="all";document.querySelector("#borderLine").value="double";document.querySelector("#borderColor").value="#336699";document.querySelector("#applyBorderPreset").click()');
      const borderDeadline = Date.now() + 3000;
      while (Date.now() < borderDeadline && !(await evaluate('document.querySelector("#notice").textContent.includes("Modified") && Boolean(document.querySelector("#cell-1-2")) && getComputedStyle(document.querySelector("#cell-1-2")).borderTopColor === "rgb(51, 102, 153)"'))) await new Promise(resolve => setTimeout(resolve, 25));
      const saveUntil = async predicate => {
        const deadline = Date.now() + 8000;
        let cell;
        do {
          await vscode.commands.executeCommand('workbench.action.files.save');
          cell = analyze(await fs.readFile(file)).sheets[0].cells.get('C2');
          if (predicate(cell)) return cell;
          await new Promise(resolve => setTimeout(resolve, 25));
        } while (Date.now() < deadline);
        assert.fail('Border save did not reach the expected state');
      };
      let saved = await saveUntil(cell => Boolean(cell.styleProperties.borders?.left && cell.styleProperties.borders?.right && cell.styleProperties.borders?.top && cell.styleProperties.borders?.bottom));
      assert.deepEqual(saved.styleProperties.borders, { left: { line: 'double', color: '#336699' }, right: { line: 'double', color: '#336699' }, top: { line: 'double', color: '#336699' }, bottom: { line: 'double', color: '#336699' } });
      assert.equal(saved.formatCode, original.formatCode);
      assert.equal(saved.styleProperties.bold, original.styleProperties.bold);
      await evaluate('document.querySelector("#borderPreset").value="none";document.querySelector("#applyBorderPreset").click()');
      const clearDeadline = Date.now() + 3000;
      while (Date.now() < clearDeadline && !(await evaluate('document.querySelector("#notice").textContent.includes("Modified") && Boolean(document.querySelector("#cell-1-2")) && getComputedStyle(document.querySelector("#cell-1-2")).borderTopColor !== "rgb(51, 102, 153)"'))) await new Promise(resolve => setTimeout(resolve, 25));
      saved = await saveUntil(cell => !cell.styleProperties.borders);
      assert.equal(saved.styleProperties.borders, undefined);
      assert.equal(saved.styleProperties.bold, original.styleProperties.bold);
    } finally {
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview applies horizontal and vertical alignment while keeping cell styles', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-alignment-host-'));
    const file = path.join(directory, 'alignment.xlsx');
    await fs.copyFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx'), file);
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.text);
        return response.result?.value;
      };
      const cellDeadline = Date.now() + 3000;
      while (Date.now() < cellDeadline && !await evaluate('Boolean(document.querySelector("#cell-1-2"))')) await new Promise(resolve => setTimeout(resolve, 25));
      const original = analyze(await fs.readFile(file)).sheets[0].cells.get('C2');
      const saveUntil = async predicate => {
        const deadline = Date.now() + 3000;
        do {
          await vscode.commands.executeCommand('workbench.action.files.save');
          const cell = analyze(await fs.readFile(file)).sheets[0].cells.get('C2');
          if (predicate(cell)) return cell;
          await new Promise(resolve => setTimeout(resolve, 25));
        } while (Date.now() < deadline);
        assert.fail('Alignment save did not reach the expected state');
      };
      const point = await evaluate(`(() => {const r=document.querySelector('#cell-1-2').getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+r.height/2}})()`);
      await devtools.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: point.x + devtools.frameOffset.x, y: point.y + devtools.frameOffset.y, button: 'left', buttons: 1, clickCount: 1 });
      await devtools.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: point.x + devtools.frameOffset.x, y: point.y + devtools.frameOffset.y, button: 'left', buttons: 0, clickCount: 1 });
      await evaluate('document.querySelector("#alignCenter").click()');
      const centerDeadline = Date.now() + 3000;
      let centered = false;
      while (Date.now() < centerDeadline) { centered = await evaluate('Boolean(document.querySelector("#cell-1-2")) && getComputedStyle(document.querySelector("#cell-1-2")).textAlign === "center"'); if (centered) break; await new Promise(resolve => setTimeout(resolve, 25)); }
      assert.equal(centered, true);
      await evaluate('window.__alignmentBefore=document.querySelector("#cell-1-2");document.querySelector("#alignBottom").click()');
      const bottomDeadline = Date.now() + 3000;
      let bottom = false;
      while (Date.now() < bottomDeadline) { bottom = await evaluate('Boolean(document.querySelector("#cell-1-2")) && document.querySelector("#cell-1-2")!==window.__alignmentBefore && getComputedStyle(document.querySelector("#cell-1-2")).justifyContent==="flex-end"'); if (bottom) break; await new Promise(resolve => setTimeout(resolve, 25)); }
      assert.equal(bottom, true, JSON.stringify(await evaluate('(()=>{const cell=document.querySelector("#cell-1-2");return{justify:cell&&getComputedStyle(cell).justifyContent,display:cell&&getComputedStyle(cell).display,notice:document.querySelector("#notice")?.textContent,address:document.querySelector("#address")?.value}})()')));
      await vscode.commands.executeCommand('workbench.action.files.save');
      let saved = analyze(await fs.readFile(file)).sheets[0].cells.get('C2');
      assert.equal(saved.styleProperties.horizontal, 'center');
      assert.equal(saved.styleProperties.vertical, 'bottom');
      assert.equal(saved.formatCode, original.formatCode);
      assert.deepEqual(saved.styleProperties.borders, original.styleProperties.borders);
      await vscode.commands.executeCommand('jinsheet.undo');
      saved = await saveUntil(cell => cell.styleProperties.horizontal === 'center' && cell.styleProperties.vertical === original.styleProperties.vertical);
      assert.equal(saved.styleProperties.vertical, original.styleProperties.vertical);
      assert.equal(saved.styleProperties.horizontal, 'center');
      await vscode.commands.executeCommand('jinsheet.undo');
      saved = await saveUntil(cell => cell.styleProperties.horizontal === original.styleProperties.horizontal);
      assert.equal(saved.styleProperties.horizontal, original.styleProperties.horizontal);
      await vscode.commands.executeCommand('jinsheet.redo');
      saved = await saveUntil(cell => cell.styleProperties.horizontal === 'center' && cell.styleProperties.vertical === original.styleProperties.vertical);
      assert.equal(saved.styleProperties.horizontal, 'center');
      await vscode.commands.executeCommand('jinsheet.redo');
      saved = await saveUntil(cell => cell.styleProperties.horizontal === 'center' && cell.styleProperties.vertical === 'bottom');
      assert.equal(saved.styleProperties.vertical, 'bottom');
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview applies a custom number format and preserves it through undo and redo', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-custom-number-format-host-'));
    const file = path.join(directory, 'custom-number-format.xlsx');
    await fs.copyFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx'), file);
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.text);
        return response.result?.value;
      };
      const original = analyze(await fs.readFile(file)).sheets[0].cells.get('C2');
      const point = await evaluate(`(() => {const r=document.querySelector('#cell-1-2').getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+r.height/2}})()`);
      await devtools.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: point.x + devtools.frameOffset.x, y: point.y + devtools.frameOffset.y, button: 'left', buttons: 1, clickCount: 1 });
      await devtools.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: point.x + devtools.frameOffset.x, y: point.y + devtools.frameOffset.y, button: 'left', buttons: 0, clickCount: 1 });
      await evaluate(`(() => {document.querySelector('#customNumberFormat').value='#,##0.000';document.querySelector('#applyCustomNumberFormat').click()})()`);
      const deadline = Date.now() + 3000;
      while (Date.now() < deadline && await evaluate('document.querySelector("#cell-1-2")?.textContent') !== '1,500.000') await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(await evaluate('document.querySelector("#cell-1-2")?.textContent'), '1,500.000');
      await vscode.commands.executeCommand('workbench.action.files.save');
      let saved = analyze(await fs.readFile(file)).sheets[0].cells.get('C2');
      assert.equal(saved.formatCode, '#,##0.000');
      assert.ok(saved.numFmtId >= 164);
      assert.deepEqual(saved.styleProperties.borders, original.styleProperties.borders);
      await vscode.commands.executeCommand('jinsheet.undo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file)).sheets[0].cells.get('C2');
      assert.equal(saved.formatCode, original.formatCode);
      await vscode.commands.executeCommand('jinsheet.redo');
      await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file)).sheets[0].cells.get('C2');
      assert.equal(saved.formatCode, '#,##0.000');
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview imports external TSV values, dates and times into XLSX with undo and redo', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-external-typed-paste-'));
    const file = path.join(directory, 'external-typed-paste.xlsx');
    await fs.copyFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx'), file);
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.text);
        return response.result?.value;
      };
      const waitFor = async (expression, message) => {
        const deadline = Date.now() + 5000;
        while (Date.now() < deadline && !await evaluate(expression)) await new Promise(resolve => setTimeout(resolve, 25));
        if (!await evaluate(expression)) assert.fail(message + ': ' + JSON.stringify(await evaluate('(() => {const cells=[0,1,2,3,4].map(column=>{const cell=document.querySelector("#cell-1-"+column),data=cell?.__jinSheetData;return{text:cell?.textContent,value:data?.value,kind:data?.kind,numFmtId:data?.numFmtId,formatCode:data?.formatCode}});return{cells,notice:document.querySelector("#notice")?.textContent,selected:document.querySelector(".cell.selected")?.id,address:document.querySelector("#address")?.value}})()')));
      };
      const pointFor = cellId => evaluate(`(() => {const r=document.querySelector('#${cellId}').getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+r.height/2}})()`);
      const dispatchMouse = params => devtools.send('Input.dispatchMouseEvent', { ...params, x: params.x + devtools.frameOffset.x, y: params.y + devtools.frameOffset.y });
      const press = async (key, code, windowsVirtualKeyCode, modifiers = 0) => {
        await devtools.send('Input.dispatchKeyEvent', { type: 'keyDown', key, code, windowsVirtualKeyCode, modifiers });
        await devtools.send('Input.dispatchKeyEvent', { type: 'keyUp', key, code, windowsVirtualKeyCode, modifiers });
      };
      await waitFor('Boolean(document.querySelector("#cell-1-0")) && Boolean(document.querySelector("#cell-1-7"))', 'the existing XLSX cells and date, percentage, datetime and time destinations are rendered');
      await vscode.env.clipboard.writeText('42\tTRUE\t\t\t2024. 2. 29.\t12.5%\t2024-02-29 12:30\t13:05');
      const target = await pointFor('cell-1-0');
      await dispatchMouse({ type: 'mousePressed', ...target, button: 'left', buttons: 1, clickCount: 1 });
      await dispatchMouse({ type: 'mouseReleased', ...target, button: 'left', buttons: 0, clickCount: 1 });
      await press('v', 'KeyV', 86, 2);
      await waitFor('document.querySelector("#cell-1-0")?.textContent === "42" && document.querySelector("#cell-1-1")?.textContent === "TRUE" && document.querySelector("#cell-1-2")?.textContent === "" && document.querySelector("#cell-1-4")?.textContent === "2/29/24" && document.querySelector("#cell-1-5")?.textContent === "12.50%" && document.querySelector("#cell-1-6")?.textContent === "2/29/24 12:30" && document.querySelector("#cell-1-7")?.textContent === "13:05"', 'external TSV values infer numbers, booleans, blanks, dates, percentages and times');
      await vscode.commands.executeCommand('workbench.action.files.save');
      let cells = analyze(await fs.readFile(file)).sheets[0].cells;
      assert.deepEqual([cells.get('A2')?.value, cells.get('A2')?.kind, cells.get('B2')?.value, cells.get('B2')?.kind, cells.get('C2')?.value, cells.get('C2')?.kind, cells.get('E2')?.value, cells.get('E2')?.kind, cells.get('E2')?.numFmtId, cells.get('F2')?.value, cells.get('F2')?.numFmtId, cells.get('G2')?.numFmtId, cells.get('H2')?.value, cells.get('H2')?.numFmtId], ['42', 'number', 'TRUE', 'boolean', '', 'blank', '45351', 'number', 14, '0.125', 10, 22, String(785 / 1440), 20]);

      await vscode.commands.executeCommand('jinsheet.undo');
      await waitFor('document.querySelector("#cell-1-0")?.textContent === "Apples" && document.querySelector("#cell-1-1")?.textContent === "12"', 'undo restores the original values');
      await vscode.commands.executeCommand('workbench.action.files.save');
      cells = analyze(await fs.readFile(file)).sheets[0].cells;
      assert.deepEqual([cells.get('A2')?.value, cells.get('A2')?.kind, cells.get('B2')?.value, cells.get('B2')?.kind, cells.get('C2')?.value, cells.get('C2')?.kind], ['Apples', 'string', '12', 'number', '1500', 'number']);

      await vscode.commands.executeCommand('jinsheet.redo');
      await waitFor('document.querySelector("#cell-1-0")?.textContent === "42" && document.querySelector("#cell-1-1")?.textContent === "TRUE" && document.querySelector("#cell-1-2")?.textContent === "" && document.querySelector("#cell-1-4")?.textContent === "2/29/24" && document.querySelector("#cell-1-5")?.textContent === "12.50%" && document.querySelector("#cell-1-6")?.textContent === "2/29/24 12:30" && document.querySelector("#cell-1-7")?.textContent === "13:05"', 'redo reapplies inferred values, blanks, dates, percentages and time display');
      await vscode.commands.executeCommand('workbench.action.files.save');
      cells = analyze(await fs.readFile(file)).sheets[0].cells;
      assert.deepEqual([cells.get('A2')?.kind, cells.get('B2')?.kind, cells.get('C2')?.kind, cells.get('E2')?.value, cells.get('E2')?.numFmtId, cells.get('F2')?.value, cells.get('F2')?.numFmtId, cells.get('G2')?.numFmtId, cells.get('H2')?.value, cells.get('H2')?.numFmtId], ['number', 'boolean', 'blank', '45351', 14, '0.125', 10, 22, String(785 / 1440), 20]);
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview sorts selected CSV data by multiple custom levels and undoes the sort', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-multi-sort-host-'));
    const file = path.join(directory, 'multi-sort.csv');
    const original = 'Department,Priority,Item\nSales,2,zeta\nSales,1,alpha\nIT,3,delta\nIT,1,charlie\n';
    await fs.writeFile(file, original, 'utf8');
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.text);
        return response.result?.value;
      };
      const waitFor = async (expression, message) => {
        const deadline = Date.now() + 5000;
        while (Date.now() < deadline && !await evaluate(expression)) await new Promise(resolve => setTimeout(resolve, 40));
        if (!await evaluate(expression)) assert.fail(message + ': ' + JSON.stringify(await evaluate('({notice:document.querySelector("#notice")?.textContent,panelHidden:document.querySelector("#sortPanel")?.hidden,selection:document.querySelectorAll(".cell.in-range").length})')));
      };
      const pointFor = async id => evaluate(`(() => {const rect=document.querySelector('#${id}').getBoundingClientRect();return{x:rect.left+rect.width/2,y:rect.top+rect.height/2}})()`);
      const dispatchMouse = params => devtools.send('Input.dispatchMouseEvent', { ...params, x: params.x + devtools.frameOffset.x, y: params.y + devtools.frameOffset.y });
      await waitFor('Boolean(document.querySelector("#cell-4-2"))', 'all CSV values are rendered');
      const point = await pointFor('cell-1-0');
      await dispatchMouse({ type: 'mousePressed', ...point, button: 'left', buttons: 1, clickCount: 1 });
      await dispatchMouse({ type: 'mouseReleased', ...point, button: 'left', buttons: 0, clickCount: 1 });
      await devtools.send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'a', code: 'KeyA', windowsVirtualKeyCode: 65, modifiers: 2 });
      await devtools.send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'a', code: 'KeyA', windowsVirtualKeyCode: 65, modifiers: 2 });
      await waitFor('document.querySelectorAll(".cell.in-range").length === 15', 'Ctrl+A selects the contiguous header and data range');
      await evaluate('document.getElementById("customSort").click()');
      await waitFor('document.getElementById("sortPanel")?.hidden === false && document.querySelectorAll("#sortLevels .sort-level").length === 1', 'custom sort dialog opens with its first level');
      await evaluate('document.getElementById("sortAddLevel").click()');
      await waitFor('document.querySelectorAll("#sortLevels .sort-level").length === 2', 'a secondary sort level can be added');
      await evaluate(`(() => {
        const levels=[...document.querySelectorAll('#sortLevels .sort-level')];
        levels[0].querySelectorAll('select')[0].value='0';levels[0].querySelectorAll('select')[1].value='ascending';
        levels[1].querySelectorAll('select')[0].value='1';levels[1].querySelectorAll('select')[1].value='descending';
        document.getElementById('sortDialogHasHeader').checked=true;
        document.getElementById('sortApply').click();return true;
      })()`);
      await waitFor('document.querySelector("#cell-1-0")?.textContent === "IT" && document.querySelector("#cell-1-1")?.textContent === "3" && document.querySelector("#cell-2-0")?.textContent === "IT" && document.querySelector("#cell-2-1")?.textContent === "1" && document.querySelector("#cell-3-0")?.textContent === "Sales" && document.querySelector("#cell-3-1")?.textContent === "2"', 'priority sorts within each department in descending order');
      await vscode.commands.executeCommand('workbench.action.files.save');
      const sorted = 'Department,Priority,Item\nIT,3,delta\nIT,1,charlie\nSales,2,zeta\nSales,1,alpha\n';
      assert.equal(await fs.readFile(file, 'utf8'), sorted, 'the ordered rows are saved together');
      await vscode.commands.executeCommand('jinsheet.undo');
      await waitFor('document.querySelector("#cell-1-0")?.textContent === "Sales" && document.querySelector("#cell-2-0")?.textContent === "Sales" && document.querySelector("#cell-3-0")?.textContent === "IT"', 'one undo restores the original row order');
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(await fs.readFile(file, 'utf8'), original, 'undo persists the original row order');
      await vscode.commands.executeCommand('jinsheet.redo');
      await waitFor('document.querySelector("#cell-1-0")?.textContent === "IT"', 'redo reapplies the multi-key sort');
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview sorts and saves a manually selected 20,000-cell CSV range', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-large-sort-host-'));
    const file = path.join(directory, 'large-sort.csv');
    const rowCount = 2000, columnCount = 10;
    const original = Array.from({ length: rowCount }, (_, row) => Array.from({ length: columnCount }, (_, column) => column === 0 ? String(rowCount - row) : `row-${row}-column-${column}`).join(',')).join('\n') + '\n';
    await fs.writeFile(file, original, 'utf8');
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.result?.exceptionDetails?.exception?.description || response.exceptionDetails.text);
        return response.result?.value;
      };
      const waitFor = async (expression, message) => {
        const deadline = Date.now() + 10000;
        while (Date.now() < deadline && !await evaluate(expression)) await new Promise(resolve => setTimeout(resolve, 40));
        const result = await evaluate(expression);
        if (!result) assert.fail(message + ': ' + JSON.stringify(await evaluate('({notice:document.querySelector("#notice")?.textContent,panelHidden:document.querySelector("#sortPanel")?.hidden,selectionCount:document.querySelectorAll(".cell.in-range").length,address:document.querySelector("#address")?.value,addressInvalid:document.querySelector("#address")?.getAttribute("aria-invalid"),goToStatus:document.querySelector("#goToStatus")?.textContent,first:document.querySelector("#cell-0-0")?.textContent})')));
      };
      await waitFor('Boolean(document.querySelector("#cell-1-0"))', 'the CSV grid is ready');
      await evaluate(`(() => {const input=document.getElementById('address');input.focus();input.value='A1:J2000';return true})()`);
      await devtools.send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13 });
      await devtools.send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13 });
      await waitFor('document.querySelector("#address")?.value === "A1" && document.querySelectorAll(".cell.in-range").length > 1', 'the name box selects the full 20,000-cell range');
      await evaluate('document.getElementById("customSort").click()');
      await waitFor('document.getElementById("sortPanel")?.hidden === false && document.querySelectorAll("#sortLevels .sort-level").length === 1', 'the custom sort dialog opens for the selected range');
      await evaluate(`(() => {
        const level=document.querySelector('#sortLevels .sort-level');
        level.querySelectorAll('select')[0].value='0';
        level.querySelectorAll('select')[1].value='ascending';
        document.getElementById('sortDialogHasHeader').checked=false;
        document.getElementById('sortDialogExpandSelection').checked=false;
        document.getElementById('sortApply').click();
        return true;
      })()`);
      await waitFor('document.querySelector("#cell-0-0")?.textContent === "1"', 'the selected 20,000-cell range is sorted in the Webview');
      await vscode.commands.executeCommand('workbench.action.files.save');
      const sorted = (await fs.readFile(file, 'utf8')).split('\n');
      assert.equal(sorted[0], ['1', ...Array.from({ length: columnCount - 1 }, (_, column) => `row-${rowCount - 1}-column-${column + 1}`)].join(','), 'the first sorted record keeps all adjacent values together');
      assert.equal(sorted[rowCount - 1], [String(rowCount), ...Array.from({ length: columnCount - 1 }, (_, column) => `row-0-column-${column + 1}`)].join(','), 'the last sorted record is persisted');
      await vscode.commands.executeCommand('jinsheet.undo');
      await waitFor('document.querySelector("#cell-0-0")?.textContent === "2000"', 'one undo restores the original first record');
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(await fs.readFile(file, 'utf8'), original, 'undo restores the complete original CSV');
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview replaces one matching cell and the current sheet with undo and redo', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-replace-ui-'));
    const file = path.join(directory, 'replace.csv'), original = 'name,note\nalpha alpha,alphabet\nAlpha,plain\n';
    await fs.writeFile(file, original, 'utf8');
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      await new Promise(resolve => setTimeout(resolve, 400));
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.text);
        return response.result?.value;
      };
      const waitFor = async (expression, message, timeoutMs = 5000) => {
        const end = Date.now() + timeoutMs;
        while (Date.now() < end) {
          if (await evaluate(expression)) return;
          await new Promise(resolve => setTimeout(resolve, 50));
        }
        assert.fail(message + ': ' + JSON.stringify(await evaluate(`(() => ({status:document.querySelector('#replaceStatus')?.textContent, first:document.querySelector('#cell-1-0')?.textContent, second:document.querySelector('#cell-1-1')?.textContent, third:document.querySelector('#cell-2-0')?.textContent}))()`)));
      };
      await waitFor('Boolean(document.querySelector("#cell-1-0"))', 'CSV cells rendered');
      await evaluate(`(() => { document.querySelector('#search').value=' '; document.querySelector('#findButton').click(); return true; })()`);
      await waitFor('document.querySelector("#searchCount")?.textContent.includes("1")', 'Find preserves a space-only query and reports its match');
      await evaluate(`document.querySelector('#clearSearch').click()`);
      await evaluate(`(() => { document.querySelector('#search').value='alpha'; document.querySelector('#replaceValue').value='Omega'; document.querySelector('#replaceCurrentButton').click(); return true; })()`);
      await waitFor('document.querySelector("#cell-1-0")?.textContent === "Omega Omega" && document.querySelector("#replaceStatus")?.textContent.includes("2 occurrence")', 'Replace without a prior Find updates only the first matching cell');
      await evaluate(`document.querySelector('#replaceAllButton').click()`);
      await waitFor('document.querySelector("#cell-1-1")?.textContent === "Omegabet" && document.querySelector("#cell-2-0")?.textContent === "Omega"', 'Replace All updates remaining matches in the current sheet');
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(await fs.readFile(file, 'utf8'), 'name,note\nOmega Omega,Omegabet\nOmega,plain\n');
      await vscode.commands.executeCommand('jinsheet.undo');
      await waitFor('document.querySelector("#cell-1-1")?.textContent === "alphabet" && document.querySelector("#cell-2-0")?.textContent === "Alpha"', 'one undo restores the Replace All operation');
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(await fs.readFile(file, 'utf8'), 'name,note\nOmega Omega,alphabet\nAlpha,plain\n');
      await vscode.commands.executeCommand('jinsheet.undo');
      await waitFor('document.querySelector("#cell-1-0")?.textContent === "alpha alpha"', 'a second undo restores Replace');
      await vscode.commands.executeCommand('jinsheet.redo');
      await waitFor('document.querySelector("#cell-1-0")?.textContent === "Omega Omega"', 'redo reapplies Replace');
      await vscode.commands.executeCommand('jinsheet.redo');
      await waitFor('document.querySelector("#cell-1-1")?.textContent === "Omegabet"', 'redo reapplies Replace All');
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(await fs.readFile(file, 'utf8'), 'name,note\nOmega Omega,Omegabet\nOmega,plain\n');
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview paginates and wraps search results in a 100,000-cell CSV', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-find-pages-ui-'));
    const file = path.join(directory, 'matches.csv');
    await fs.writeFile(file, Array.from({ length: 10000 }, (_, index) => Array.from({ length: 10 }, (_, column) => column === 0 && index <= 9800 && index % 49 === 0 ? `needle-${index}` : `other-${index}-${column}`).join(',')).join('\n'), 'utf8');
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.text);
        return response.result?.value;
      };
      const waitFor = async (expression, message, timeoutMs = 15000) => {
        const end = Date.now() + timeoutMs;
        while (Date.now() < end) { if (await evaluate(expression)) return; await new Promise(resolve => setTimeout(resolve, 30)); }
        assert.fail(message + ': ' + JSON.stringify(await evaluate(`(()=>({address:document.querySelector('#address')?.value,count:document.querySelector('#searchCount')?.textContent,selected:document.querySelector('.cell.selected')?.id}))()`)));
      };
      await waitFor('document.querySelector("#cell-0-0")?.textContent.trim() === "needle-0"', 'the first CSV record is rendered before searching');
      await waitFor('Boolean(document.querySelector("#search"))', 'Find control is ready');
      await evaluate(`(()=>{const search=document.querySelector('#search');search.value='needle';search.dispatchEvent(new Event('input',{bubbles:true}));document.querySelector('#findButton').click()})()`);
      await waitFor('document.querySelector("#searchCount")?.textContent.includes("201") && document.querySelector("#address")?.value === "A1"', 'first result page starts at the first match in the 100,000-cell file');
      let started = performance.now();
      await evaluate(`(()=>{const search=document.querySelector('#search');for(let index=0;index<100;index++)search.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true,cancelable:true}))})()`);
      await waitFor('document.querySelector("#address")?.value === "A4901" && document.querySelector(".cell.selected")?.id === "cell-4900-0"', 'Enter moves past the first 100 results to the next matching row');
      const secondPageMs = performance.now() - started;
      started = performance.now();
      await evaluate(`(()=>{const search=document.querySelector('#search');for(let index=0;index<100;index++)search.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true,cancelable:true}))})()`);
      await waitFor('document.querySelector("#address")?.value === "A9801" && document.querySelector(".cell.selected")?.id === "cell-9800-0"', 'the last result page remains row-ordered in the large file');
      const thirdPageMs = performance.now() - started;
      started = performance.now();
      await evaluate(`document.querySelector('#search').dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true,cancelable:true}))`);
      await waitFor('document.querySelector("#address")?.value === "A1" && document.querySelector(".cell.selected")?.id === "cell-0-0"', 'Enter wraps from the last result to the first');
      console.log('[JinSheet Search 100k]', JSON.stringify({ cells: 100000, rows: 10000, columns: 10, matches: 201, secondPageMs: Math.round(secondPageMs), thirdPageMs: Math.round(thirdPageMs), wrapMs: Math.round(performance.now() - started) }));
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview opens, searches, edits and saves a 100,000-cell CSV', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-large-csv-host-'));
    const file = path.join(directory, 'large.csv');
    const rows = Array.from({ length: 10000 }, (_, row) => Array.from({ length: 10 }, (_, column) =>
      row === 0 && column === 0 ? 'first-cell' : row === 9999 && column === 9 ? 'needle-token' : `r${String(row).padStart(5, '0')}-c${String(column).padStart(2, '0')}-` + 'x'.repeat(68)
    ).join(','));
    await fs.writeFile(file, rows.join('\r\n'), 'utf8');
    const sourceBytes = (await fs.stat(file)).size;
    rows.length = 0;
    await new Promise(resolve => setImmediate(resolve));
    let devtools;
    try {
      let started = performance.now();
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview(30000, 'first-cell');
      const openedMs = performance.now() - started;
      const memoryAfterOpen = process.memoryUsage();
      await devtools.send('Performance.enable');
      const rendererMetrics = await devtools.send('Performance.getMetrics');
      const rendererJsHeapBytes = rendererMetrics.metrics?.find(metric => metric.name === 'JSHeapUsedSize')?.value ?? null;
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.exception?.description || response.exceptionDetails.text);
        return response.result?.value;
      };
      const waitFor = async (expression, message, timeoutMs = 30000) => {
        const deadline = Date.now() + timeoutMs;
        while (Date.now() < deadline) {
          if (await evaluate(expression)) return;
          await new Promise(resolve => setTimeout(resolve, 25));
        }
        assert.fail(message + ': ' + await evaluate('document.querySelector("#notice")?.textContent'));
      };

      await waitFor('Boolean(document.querySelector("#search"))&&Boolean(document.querySelector("#findButton"))', 'large CSV search controls render');
      started = performance.now();
      await evaluate('(()=>{const search=document.querySelector("#search");search.value="needle-token";search.dispatchEvent(new Event("input",{bubbles:true}));document.querySelector("#findButton").click()})()');
      await waitFor('document.querySelector("#searchCount")?.textContent.includes("1")&&document.querySelector("#address")?.value==="J10000"&&document.querySelector(".cell.selected")?.id==="cell-9999-9"', 'search reaches the unique last cell');
      const searchMs = performance.now() - started;

      started = performance.now();
      await evaluate('(()=>{const input=document.querySelector("#formula");input.value="edited-value";input.dispatchEvent(new Event("input",{bubbles:true}));input.dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}))})()');
      await waitFor('document.querySelector("#cell-9999-9")?.textContent==="edited-value"', 'the last CSV cell accepts an edit');
      const editToRenderMs = performance.now() - started;

      started = performance.now();
      const saveDeadline = Date.now() + 30000;
      let savedText = '';
      do {
        await vscode.commands.executeCommand('workbench.action.files.save');
        savedText = await fs.readFile(file, 'utf8');
        if (savedText.endsWith('edited-value')) break;
        await new Promise(resolve => setTimeout(resolve, 25));
      } while (Date.now() < saveDeadline);
      const saveToDiskMs = performance.now() - started;
      assert.equal(savedText.endsWith('edited-value'), true, 'the edited last cell reaches the CSV file');
      assert.equal(savedText.includes('needle-token'), false, 'the original last-cell value is replaced');
      console.log('[JinSheet CSV 100k]', JSON.stringify({ sourceBytes, rows: 10000, columns: 10, cells: 100000, openedMs: Math.round(openedMs), searchMs: Math.round(searchMs), editToRenderMs: Math.round(editToRenderMs), saveToDiskMs: Math.round(saveToDiskMs), extensionHostHeapMiB: Number((memoryAfterOpen.heapUsed / 1048576).toFixed(1)), extensionHostExternalMiB: Number((memoryAfterOpen.external / 1048576).toFixed(1)), extensionHostArrayBuffersMiB: Number((memoryAfterOpen.arrayBuffers / 1048576).toFixed(1)), extensionHostRssMiB: Number((memoryAfterOpen.rss / 1048576).toFixed(1)), chromiumRendererJsHeapMiB: rendererJsHeapBytes === null ? null : Number((rendererJsHeapBytes / 1048576).toFixed(1)) }));
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview sorts, saves and undoes a 100,000-cell XLSX range', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-large-xlsx-sort-host-'));
    const file = path.join(directory, 'large-sort.xlsx');
    const rowCount = 20000, columnCount = 5;
    const base = analyze(await fs.readFile(path.join(__dirname, '..', 'fixtures', 'JinSheet-Styled-Sample.xlsx')));
    const workbookFiles = { ...base.files };
    const columnName = column => String.fromCharCode(64 + column);
    const rows = [];
    for (let row = 1; row <= rowCount; row++) {
      const cells = [`<c r="A${row}"><v>${rowCount - row + 1}</v></c>`];
      for (let column = 2; column <= columnCount; column++) {
        cells.push(`<c r="${columnName(column)}${row}" t="inlineStr"><is><t>row-${row}-column-${column}</t></is></c>`);
      }
      rows.push(`<row r="${row}">${cells.join('')}</row>`);
    }
    workbookFiles[base.sheets[0].path] = new TextEncoder().encode(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:E${rowCount}"/><sheetData>${rows.join('')}</sheetData></worksheet>`);
    await fs.writeFile(file, zipSync(workbookFiles, { level: 0 }));
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview(30000, '20000');
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.exception?.description || response.exceptionDetails.text);
        return response.result?.value;
      };
      const waitFor = async (expression, message, timeoutMs = 30000) => {
        const deadline = Date.now() + timeoutMs;
        while (Date.now() < deadline) {
          if (await evaluate(expression)) return;
          await new Promise(resolve => setTimeout(resolve, 40));
        }
        assert.fail(message + ': ' + JSON.stringify(await evaluate('({notice:document.querySelector("#notice")?.textContent,address:document.querySelector("#address")?.value,goToStatus:document.querySelector("#goToStatus")?.textContent,sortPanelHidden:document.querySelector("#sortPanel")?.hidden,firstCell:document.querySelector("#cell-0-0")?.textContent})')));
      };
      await waitFor('document.querySelector("#cell-0-0")?.textContent === "20000"', 'the 100,000-cell XLSX grid renders its original first value');
      await evaluate(`(()=>{const input=document.querySelector('#address');input.focus();input.value='A1:E20000'})()`);
      await devtools.send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13 });
      await devtools.send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13 });
      await waitFor('document.querySelector("#address")?.value === "A1" && document.querySelectorAll(".cell.in-range").length > 1', 'the name box selects all 100,000 cells');
      await evaluate('document.querySelector("#customSort").click()');
      await waitFor('document.querySelector("#sortPanel")?.hidden === false && document.querySelectorAll("#sortLevels .sort-level").length === 1', 'the custom sort dialog opens for the full XLSX range');
      await evaluate(`(()=>{const level=document.querySelector('#sortLevels .sort-level');level.querySelectorAll('select')[0].value='0';level.querySelectorAll('select')[1].value='ascending';document.querySelector('#sortDialogHasHeader').checked=false;document.querySelector('#sortDialogExpandSelection').checked=false;document.querySelector('#sortApply').click()})()`);
      const sortStarted = performance.now();
      await waitFor('document.querySelector("#cell-0-0")?.textContent === "1"', 'the complete XLSX range sorts in the Webview');
      const sortToRenderMs = performance.now() - sortStarted;
      await vscode.commands.executeCommand('workbench.action.files.save');
      let saved = analyze(await fs.readFile(file));
      assert.equal(saved.writable, true, saved.readOnlyReason);
      assert.equal(saved.sheets[0].cells.size, 100000);
      assert.equal(saved.sheets[0].cells.get('A1').value, '1');
      assert.equal(saved.sheets[0].cells.get('B1').value, 'row-20000-column-2');
      assert.equal(saved.sheets[0].cells.get('E1').value, 'row-20000-column-5');
      assert.equal(saved.sheets[0].cells.get('A20000').value, '20000');

      await vscode.commands.executeCommand('jinsheet.undo');
      await waitFor('document.querySelector("#cell-0-0")?.textContent === "20000"', 'one undo restores the original first row');
      await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file));
      assert.equal(saved.sheets[0].cells.size, 100000);
      assert.equal(saved.sheets[0].cells.get('A1').value, '20000');
      assert.equal(saved.sheets[0].cells.get('B1').value, 'row-1-column-2');
      assert.equal(saved.sheets[0].cells.get('A20000').value, '1');
      await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
      devtools.close(); devtools = null;
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview(30000, '20000');
      await waitFor('document.querySelector("#cell-0-0")?.textContent === "20000" && document.querySelector("#cell-0-1")?.textContent === "row-1-column-2"', 'the saved 100,000-cell workbook reopens in its original row order');
      console.log('[JinSheet XLSX sort 100k]', JSON.stringify({ cells: 100000, sortToRenderMs: Math.round(sortToRenderMs) }));
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview previews duplicate rows, preserves data on cancel and removes them through save, undo and reopen', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-remove-duplicates-ui-'));
    const file = path.join(directory, 'duplicates.csv');
    const original = 'Name,Region\nAda,West\nada,WEST\nLin,East\nLin,East\n';
    await fs.writeFile(file, original, 'utf8');
    const originalWarning = vscode.window.showWarningMessage;
    const prompts = [];
    let acceptWarning = false;
    vscode.window.showWarningMessage = async (text, _options, ...actions) => {
      prompts.push({ text, actions });
      return acceptWarning ? actions[0] : undefined;
    };
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.exception?.description || response.exceptionDetails.text);
        return response.result?.value;
      };
      const waitFor = async (expression, message) => {
        const deadline = Date.now() + 5000;
        while (Date.now() < deadline && !await evaluate(expression)) await new Promise(resolve => setTimeout(resolve, 25));
        if (!await evaluate(expression)) assert.fail(message + ': ' + JSON.stringify(await evaluate('({status:document.querySelector("#removeDuplicatesStatus")?.textContent,notice:document.querySelector("#notice")?.textContent,panelHidden:document.querySelector("#removeDuplicatesPanel")?.hidden,selected:document.querySelector(".cell.selected")?.id,range:[...document.querySelectorAll(".cell.in-range")].map(cell=>cell.id),address:document.querySelector("#address")?.value,columns:[...document.querySelectorAll("#removeDuplicatesColumns input")].map(input=>({value:input.value,checked:input.checked})),warnings:'+JSON.stringify(prompts)+'})')));
      };
      await waitFor('Boolean(document.querySelector("#cell-4-1"))', 'the selected data and duplicate records are rendered');
      const first = await evaluate('(()=>{const r=document.querySelector("#cell-0-0").getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+r.height/2}})()');
      const last = await evaluate('(()=>{const r=document.querySelector("#cell-4-1").getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+r.height/2}})()');
      const mouse = (type, point) => devtools.send('Input.dispatchMouseEvent', { type, x: point.x + devtools.frameOffset.x, y: point.y + devtools.frameOffset.y, button: type === 'mouseMoved' ? 'none' : 'left', buttons: type === 'mouseReleased' ? 0 : 1, clickCount: 1 });
      await mouse('mousePressed', first); await mouse('mouseMoved', last); await mouse('mouseReleased', last);
      await waitFor('document.querySelectorAll(".cell.in-range").length === 10', 'the entire range including its header is selected');
      await evaluate('document.querySelector("#removeDuplicates").click()');
      await waitFor('document.querySelectorAll("#removeDuplicatesColumns input").length === 2', 'selected columns are available as comparison keys');
      await waitFor('Boolean(document.querySelector("#removeDuplicatesPanel").hidden === false)', 'the duplicate preview panel opens');
      await new Promise(resolve => setTimeout(resolve, 250));
      const isKorean = await evaluate('document.documentElement.lang === "ko"');
      if (process.env.JINSHEET_HOST_LOCALE?.toLowerCase().startsWith('ko')) assert.equal(isKorean, true, 'the requested Korean locale selects the Korean JinSheet UI');
      const previewStatus = await evaluate('document.querySelector("#removeDuplicatesStatus")?.textContent');
      assert.match(previewStatus, isKorean ? /중복 행 2개/ : /2 duplicate rows/, 'the preview reports two duplicate rows');
      assert.equal(await evaluate('document.querySelector("#removeDuplicatesApply").disabled'), false);
      await evaluate('document.querySelector("#removeDuplicatesApply").click()');
      const cancelDeadline = Date.now() + 5000;
      while (Date.now() < cancelDeadline && prompts.length < 1) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(prompts.length, 1, 'applying requires an explicit confirmation');
      assert.match(prompts[0].text, /2/);
      await waitFor(`document.querySelector("#removeDuplicatesStatus")?.textContent.includes(${JSON.stringify(isKorean ? '취소했습니다' : 'cancelled')})`, 'cancelling the confirmation keeps the tool open with feedback');
      assert.equal(await evaluate('document.querySelector("#cell-4-0")?.textContent'), 'Lin', 'cancel leaves all source rows intact');

      acceptWarning = true;
      await evaluate('document.querySelector("#removeDuplicatesApply").click()');
      const applyDeadline = Date.now() + 5000;
      while (Date.now() < applyDeadline && prompts.length < 2) await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(prompts.length, 2, 'the destructive action asks again before deleting rows');
      await waitFor('document.querySelector("#removeDuplicatesPanel").hidden && document.querySelector("#cell-2-0")?.textContent === "Lin" && document.querySelector("#cell-3-0")?.textContent === ""', 'confirming removes both duplicate rows and closes the tool');
      await vscode.commands.executeCommand('workbench.action.files.save');
      const deduplicated = 'Name,Region\nAda,West\nLin,East\n';
      assert.equal(await fs.readFile(file, 'utf8'), deduplicated, 'save writes only the first row from each duplicate group');

      await vscode.commands.executeCommand('jinsheet.undo');
      await waitFor('document.querySelector("#cell-4-0")?.textContent === "Lin"', 'one undo restores both removed rows');
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(await fs.readFile(file, 'utf8'), original, 'undo restores the complete original CSV');
      await vscode.commands.executeCommand('jinsheet.redo');
      await waitFor('document.querySelector("#cell-2-0")?.textContent === "Lin" && document.querySelector("#cell-3-0")?.textContent === ""', 'one redo reapplies the duplicate removal');
      await vscode.commands.executeCommand('workbench.action.files.save');
      devtools.close(); devtools = null;
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      await waitFor('document.querySelector("#cell-2-0")?.textContent === "Lin" && document.querySelector("#cell-3-0")?.textContent === ""', 'the saved result survives reopening');
      assert.equal(await fs.readFile(file, 'utf8'), deduplicated);
    } finally {
      vscode.window.showWarningMessage = originalWarning;
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview supports editor selection, clipboard and range drag auto-scroll', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-webview-ui-'));
    const file = path.join(directory, 'webview-ui.csv');
    await fs.writeFile(file, 'name,value\nalpha,1\n', 'utf8');
    let devtools;
    const originalClipboardRead = JinSheetProvider.prototype.readClipboardText;
    const originalClipboardWrite = JinSheetProvider.prototype.writeClipboardText;
    let testClipboard = '';
    try {
      JinSheetProvider.prototype.readClipboardText = function(reader) { return originalClipboardRead.call(this, reader || (async () => testClipboard)); };
      JinSheetProvider.prototype.writeClipboardText = function(text) { return originalClipboardWrite.call(this, text, async value => { testClipboard = value; }); };
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      await new Promise(resolve => setTimeout(resolve, 400));
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.text);
        return response.result?.value;
      };
      const waitFor = async (expression, message, timeoutMs = 5000) => {
        const end = Date.now() + timeoutMs;
        while (Date.now() < end) {
          if (await evaluate(expression)) return;
          await new Promise(resolve => setTimeout(resolve, 50));
        }
        assert.fail(message + ': ' + JSON.stringify(await evaluate(`(() => { const grid=document.querySelector('#grid'), rect=grid?.getBoundingClientRect(); return {target:document.querySelector('#cell-1-1')?.textContent, selected:document.querySelector('.cell.selected')?.id, active:document.activeElement?.id, address:document.querySelector('#address')?.value, formula:document.querySelector('#formula')?.value, notice:document.querySelector('#notice')?.textContent, grid:rect&&{scrollLeft:grid.scrollLeft,clientWidth:grid.clientWidth,scrollWidth:grid.scrollWidth,innerWidth,rect:{left:rect.left,right:rect.right}}}; })()`)));
      };
      const dispatchMouse = params => devtools.send('Input.dispatchMouseEvent', {
        ...params, x: params.x + devtools.frameOffset.x, y: params.y + devtools.frameOffset.y
      });
      await waitFor('Boolean(document.querySelector("#cell-1-0"))', 'CSV cells rendered in the Chromium Webview');

      const cellPoint = await evaluate(`(() => { const r = document.querySelector('#cell-1-0').getBoundingClientRect(); return {x:r.left + r.width / 2, y:r.top + r.height / 2}; })()`);
      for (let clickCount = 1; clickCount <= 2; clickCount++) {
        await dispatchMouse({ type: 'mousePressed', ...cellPoint, button: 'left', buttons: 1, clickCount });
        await dispatchMouse({ type: 'mouseReleased', ...cellPoint, button: 'left', buttons: 0, clickCount });
      }
      await waitFor('Boolean(document.querySelector("#cell-1-0 textarea.cell-editor"))', 'double-click opened the inline cell editor');
      const editorGeometry = await evaluate(`(() => {
        const editor = document.querySelector('#cell-1-0 textarea.cell-editor'), rect = editor.getBoundingClientRect();
        const context = document.createElement('canvas').getContext('2d'); context.font = getComputedStyle(editor).font;
        const style = getComputedStyle(editor), contentLeft = rect.left + parseFloat(style.paddingLeft);
        return {x:contentLeft + context.measureText('a').width * 1.5, endX:contentLeft + context.measureText('alph').width - context.measureText('a').width * .5, y:rect.top + rect.height / 2, value:editor.value};
      })()`);
      assert.equal(editorGeometry.value, 'alpha');
      await dispatchMouse({ type: 'mousePressed', x: editorGeometry.x, y: editorGeometry.y, button: 'left', buttons: 1, clickCount: 1 });
      await dispatchMouse({ type: 'mouseMoved', x: editorGeometry.endX, y: editorGeometry.y, button: 'left', buttons: 1 });
      await dispatchMouse({ type: 'mouseReleased', x: editorGeometry.endX, y: editorGeometry.y, button: 'left', buttons: 0, clickCount: 1 });
      const selection = await evaluate(`(() => { const editor = document.querySelector('#cell-1-0 textarea.cell-editor'); return {text:editor.value.slice(editor.selectionStart, editor.selectionEnd), value:editor.value, active:document.activeElement === editor, connected:editor.isConnected}; })()`);
      assert.ok(selection.text.length > 0 && selection.value === 'alpha' && selection.active && selection.connected, 'mouse dragging selects editor text without replacing or blurring its textarea');
      await devtools.send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'Escape', code: 'Escape', windowsVirtualKeyCode: 27 });
      await devtools.send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'Escape', code: 'Escape', windowsVirtualKeyCode: 27 });

      const sourcePoint = await evaluate(`(() => { const r = document.querySelector('#cell-0-0').getBoundingClientRect(); return {x:r.left + r.width / 2, y:r.top + r.height / 2}; })()`);
      await dispatchMouse({ type: 'mousePressed', ...sourcePoint, button: 'left', buttons: 1, clickCount: 1 });
      await dispatchMouse({ type: 'mouseReleased', ...sourcePoint, button: 'left', buttons: 0, clickCount: 1 });
      await devtools.send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'c', code: 'KeyC', windowsVirtualKeyCode: 67, modifiers: 2 });
      await devtools.send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'c', code: 'KeyC', windowsVirtualKeyCode: 67, modifiers: 2 });
      const clipboardDeadline = Date.now() + 5000;
      while (Date.now() < clipboardDeadline && testClipboard !== 'name') await new Promise(resolve => setTimeout(resolve, 25));
      assert.equal(testClipboard, 'name', 'Ctrl+C writes the selected cell to the clipboard');

      const targetPoint = await evaluate(`(() => { const r = document.querySelector('#cell-1-1').getBoundingClientRect(); return {x:r.left + r.width / 2, y:r.top + r.height / 2}; })()`);
      await dispatchMouse({ type: 'mousePressed', ...targetPoint, button: 'left', buttons: 1, clickCount: 1 });
      await dispatchMouse({ type: 'mouseReleased', ...targetPoint, button: 'left', buttons: 0, clickCount: 1 });
      await devtools.send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'v', code: 'KeyV', windowsVirtualKeyCode: 86, modifiers: 2 });
      await devtools.send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'v', code: 'KeyV', windowsVirtualKeyCode: 86, modifiers: 2 });
      await waitFor('document.querySelector("#cell-1-1")?.textContent === "name"', 'Ctrl+V pastes the copied cell into a different selected cell');
      await vscode.commands.executeCommand('workbench.action.files.save');
      const savedDeadline = Date.now() + 5000;
      while (Date.now() < savedDeadline && (await fs.readFile(file, 'utf8')) === 'name,value\nalpha,1\n') await new Promise(resolve => setTimeout(resolve, 50));
      assert.equal(await fs.readFile(file, 'utf8'), 'name,value\nalpha,name\n', 'pasted value persists after the custom editor saves');
      await vscode.commands.executeCommand('jinsheet.undo');
      await waitFor('document.querySelector("#cell-1-1")?.textContent === "1"', 'undo removes the pasted cell value');
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(await fs.readFile(file, 'utf8'), 'name,value\nalpha,1\n', 'undo of a clipboard paste is saved to the CSV');
      await vscode.commands.executeCommand('jinsheet.redo');
      await waitFor('document.querySelector("#cell-1-1")?.textContent === "name"', 'redo restores the clipboard paste');
      await vscode.commands.executeCommand('workbench.action.files.save');
      const redoSaveDeadline = Date.now() + 5000;
      while (Date.now() < redoSaveDeadline && await fs.readFile(file, 'utf8') !== 'name,value\nalpha,name\n') await new Promise(resolve => setTimeout(resolve, 50));
      assert.equal(await fs.readFile(file, 'utf8'), 'name,value\nalpha,name\n', 'redo of a clipboard paste is saved to the CSV');

      devtools.close();
      devtools = await waitForSpreadsheetWebview(15000, 'name');
      await waitFor('Boolean(document.querySelector("#cell-1-0"))', 'source row is rendered again after saving');
      const externalTarget = await evaluate(`(() => { const r = document.querySelector('#cell-1-0').getBoundingClientRect(); return {x:r.left + r.width / 2, y:r.top + r.height / 2}; })()`);
      await dispatchMouse({ type: 'mousePressed', ...externalTarget, button: 'left', buttons: 1, clickCount: 1 });
      await dispatchMouse({ type: 'mouseReleased', ...externalTarget, button: 'left', buttons: 0, clickCount: 1 });
      testClipboard = 'external';
      await devtools.send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'v', code: 'KeyV', windowsVirtualKeyCode: 86, modifiers: 2 });
      await devtools.send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'v', code: 'KeyV', windowsVirtualKeyCode: 86, modifiers: 2 });
      await waitFor('document.querySelector("#cell-1-0")?.textContent === "external"', 'Ctrl+V pastes text from an external clipboard source');
      await vscode.commands.executeCommand('workbench.action.files.save');
      const externalSaveDeadline = Date.now() + 5000;
      while (Date.now() < externalSaveDeadline && (await fs.readFile(file, 'utf8')) === 'name,value\nalpha,name\n') await new Promise(resolve => setTimeout(resolve, 50));
      assert.equal(await fs.readFile(file, 'utf8'), 'name,value\nexternal,name\n', 'external clipboard content persists after the custom editor saves');
      await vscode.commands.executeCommand('jinsheet.undo');
      await waitFor('document.querySelector("#cell-1-0")?.textContent === "alpha"', 'undo restores the original value after an external clipboard paste');
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(await fs.readFile(file, 'utf8'), 'name,value\nalpha,name\n', 'undo of an external clipboard paste is saved');
      await vscode.commands.executeCommand('jinsheet.redo');
      await waitFor('document.querySelector("#cell-1-0")?.textContent === "external"', 'redo restores the external clipboard paste');
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(await fs.readFile(file, 'utf8'), 'name,value\nexternal,name\n', 'redo of an external clipboard paste is saved');

      await waitFor('Boolean(document.querySelector("#cell-1-0"))', 'starting cell is rendered for drag auto-scroll');
      const dragStartExpression = `(() => { const cell = document.querySelector('#cell-1-0'), grid = document.querySelector('#grid'); if (!cell || !grid) return null; const cellRect=cell.getBoundingClientRect(), gridRect=grid.getBoundingClientRect(); return {x:cellRect.left + cellRect.width / 2, y:cellRect.top + cellRect.height / 2, edgeX:cellRect.left + cellRect.width / 2, edgeY:gridRect.bottom - 2, startScroll:grid.scrollTop}; })()`;
      let dragStart = await evaluate(dragStartExpression);
      const dragStartDeadline = Date.now() + 5000;
      while (!dragStart && Date.now() < dragStartDeadline) { await new Promise(resolve => setTimeout(resolve, 25)); dragStart = await evaluate(dragStartExpression); }
      assert.ok(dragStart, 'the starting cell and grid remain rendered before drag auto-scroll');
      await dispatchMouse({ type: 'mousePressed', x: dragStart.x, y: dragStart.y, button: 'left', buttons: 1, clickCount: 1 });
      await dispatchMouse({ type: 'mouseMoved', x: dragStart.edgeX, y: dragStart.edgeY, button: 'left', buttons: 1 });
      await waitFor(`document.querySelector('#grid').scrollTop > ${dragStart.startScroll + 20}`, 'dragging at the grid edge scrolls the worksheet');
      await waitFor(`document.querySelector('#address').value !== 'A2'`, 'drag auto-scroll extends the active cell selection');
      await dispatchMouse({ type: 'mouseReleased', x: dragStart.edgeX, y: dragStart.edgeY, button: 'left', buttons: 0, clickCount: 1 });

      await evaluate(`(() => { const grid=document.querySelector('#grid'); grid.scrollTop=0; grid.scrollLeft=0; return true; })()`);
      await waitFor('Boolean(document.querySelector("#cell-1-0"))', 'top-left cells are rendered before testing horizontal drag');
      const dragOrigin = await evaluate(`(() => { const r=document.querySelector('#cell-1-0').getBoundingClientRect(); return {x:r.left+r.width/2,y:r.top+r.height/2}; })()`);
      await dispatchMouse({ type: 'mousePressed', ...dragOrigin, button: 'left', buttons: 1, clickCount: 1 });
      await dispatchMouse({ type: 'mouseReleased', ...dragOrigin, button: 'left', buttons: 0, clickCount: 1 });
      const dragRight = await evaluate(`(() => { const cell = document.querySelector('.cell.selected').getBoundingClientRect(), grid = document.querySelector('#grid').getBoundingClientRect(); return {x:cell.left + cell.width / 2, y:cell.top + cell.height / 2, edgeX:grid.right - 2, scroll:document.querySelector('#grid').scrollLeft, address:document.querySelector('#address').value}; })()`);
      await dispatchMouse({ type: 'mousePressed', x: dragRight.x, y: dragRight.y, button: 'left', buttons: 1, clickCount: 1 });
      await dispatchMouse({ type: 'mouseMoved', x: dragRight.edgeX, y: dragRight.y, button: 'left', buttons: 1 });
      await waitFor(`document.querySelector('#grid').scrollLeft > ${dragRight.scroll + 20}`, 'dragging at the right edge scrolls horizontally');
      await waitFor(`document.querySelector('#address').value !== '${dragRight.address}'`, 'horizontal auto-scroll extends the selection');
      await dispatchMouse({ type: 'mouseReleased', x: dragRight.edgeX, y: dragRight.y, button: 'left', buttons: 0, clickCount: 1 });

      await evaluate(`(() => { const grid=document.querySelector('#grid'); grid.scrollLeft=200; grid.scrollTop=0; return true; })()`);
      await waitFor('Boolean(document.querySelector("#cell-1-1"))', 'column B is rendered before testing leftward drag');
      const leftOrigin = await evaluate(`(() => { const r=document.querySelector('#cell-1-1').getBoundingClientRect(); return {x:r.left+r.width/2,y:r.top+r.height/2}; })()`);
      await dispatchMouse({ type: 'mousePressed', ...leftOrigin, button: 'left', buttons: 1, clickCount: 1 });
      await dispatchMouse({ type: 'mouseReleased', ...leftOrigin, button: 'left', buttons: 0, clickCount: 1 });
      await waitFor('document.hasFocus()&&document.activeElement?.id==="cell-1-1"&&document.querySelector("#address").value==="B2"', 'the leftward drag source owns focus before pressing the mouse again');
      const dragLeft = await evaluate(`(() => { const cell = document.querySelector('.cell.selected').getBoundingClientRect(), grid = document.querySelector('#grid').getBoundingClientRect(); return {x:cell.left + cell.width / 2, y:cell.top + cell.height / 2, edgeX:grid.left + 2, scroll:document.querySelector('#grid').scrollLeft, address:document.querySelector('#address').value}; })()`);
      await dispatchMouse({ type: 'mousePressed', x: dragLeft.x, y: dragLeft.y, button: 'left', buttons: 1, clickCount: 1 });
      await dispatchMouse({ type: 'mouseMoved', x: dragLeft.edgeX, y: dragLeft.y, button: 'left', buttons: 1 });
      await waitFor(`document.querySelector('#grid').scrollLeft < ${dragLeft.scroll - 20}`, 'dragging at the left edge scrolls back');
      await waitFor(`document.querySelector('#address').value !== '${dragLeft.address}'`, 'leftward auto-scroll updates the active selection');
      await dispatchMouse({ type: 'mouseReleased', x: dragLeft.edgeX, y: dragLeft.y, button: 'left', buttons: 0, clickCount: 1 });

      await evaluate(`(() => { const grid=document.querySelector('#grid'); grid.scrollTop=400; grid.scrollLeft=0; return true; })()`);
      await waitFor('Boolean(document.querySelector("#cell-20-0"))', 'row 21 is rendered before testing upward drag');
      const upOrigin = await evaluate(`(() => { const r=document.querySelector('#cell-20-0').getBoundingClientRect(); return {x:r.left+r.width/2,y:r.top+r.height/2}; })()`);
      await dispatchMouse({ type: 'mousePressed', ...upOrigin, button: 'left', buttons: 1, clickCount: 1 });
      await dispatchMouse({ type: 'mouseReleased', ...upOrigin, button: 'left', buttons: 0, clickCount: 1 });
      const dragUp = await evaluate(`(() => { const cell = document.querySelector('.cell.selected').getBoundingClientRect(), grid = document.querySelector('#grid').getBoundingClientRect(); return {x:cell.left + cell.width / 2, y:cell.top + cell.height / 2, edgeY:grid.top + 2, scroll:document.querySelector('#grid').scrollTop, address:document.querySelector('#address').value}; })()`);
      await dispatchMouse({ type: 'mousePressed', x: dragUp.x, y: dragUp.y, button: 'left', buttons: 1, clickCount: 1 });
      await dispatchMouse({ type: 'mouseMoved', x: dragUp.x, y: dragUp.edgeY, button: 'left', buttons: 1 });
      await waitFor(`document.querySelector('#grid').scrollTop < ${dragUp.scroll - 20}`, 'dragging at the top edge scrolls upward');
      await waitFor(`document.querySelector('#address').value !== '${dragUp.address}'`, 'upward auto-scroll updates the active selection');
      await dispatchMouse({ type: 'mouseReleased', x: dragUp.x, y: dragUp.edgeY, button: 'left', buttons: 0, clickCount: 1 });
    } finally {
      JinSheetProvider.prototype.readClipboardText = originalClipboardRead;
      JinSheetProvider.prototype.writeClipboardText = originalClipboardWrite;
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('Chromium Webview maps drag auto-scroll across filtered rows and frozen panes', async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-filter-freeze-drag-'));
    const file = path.join(directory, 'filtered-frozen.xlsx');
    const inlineCell = (address, value) => `<c r="${address}" t="inlineStr"><is><t>${value}</t></is></c>`;
    const rows = Array.from({ length: 300 }, (_, index) => {
      const row = index + 1, type = row === 1 ? 'Type' : row % 2 === 0 ? 'keep' : 'hide';
      return `<row r="${row}">${inlineCell(`A${row}`, type)}${inlineCell(`B${row}`, `v${row}`)}</row>`;
    }).join('');
    const sheet = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B300"/><sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/></sheetView></sheetViews><sheetData>${rows}</sheetData><autoFilter ref="A1:B300"><filterColumn colId="0"><filters><filter val="keep"/></filters></filterColumn></autoFilter></worksheet>`;
    const bytes = zipSync({
      '[Content_Types].xml': new TextEncoder().encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'),
      '_rels/.rels': new TextEncoder().encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
      'xl/workbook.xml': new TextEncoder().encode('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Filter" sheetId="1" r:id="s1"/></sheets></workbook>'),
      'xl/_rels/workbook.xml.rels': new TextEncoder().encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="s1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'),
      'xl/worksheets/sheet1.xml': new TextEncoder().encode(sheet)
    });
    await fs.writeFile(file, bytes);
    let devtools;
    try {
      await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
      devtools = await waitForSpreadsheetWebview();
      const evaluate = async expression => {
        const response = await devtools.send('Runtime.evaluate', { expression, contextId: devtools.contextId, returnByValue: true, awaitPromise: true });
        if (response.exceptionDetails) throw new Error(response.exceptionDetails.text);
        return response.result?.value;
      };
      const waitFor = async (expression, message, timeoutMs = 7000) => {
        const end = Date.now() + timeoutMs;
        while (Date.now() < end) {
          if (await evaluate(expression)) return;
          await new Promise(resolve => setTimeout(resolve, 50));
        }
        assert.fail(message + ': ' + JSON.stringify(await evaluate(`(() => ({address:document.querySelector('#address')?.value,selected:document.querySelector('.cell.selected')?.id,selectedCell:document.querySelector('#cell-1-1')?.outerHTML,selectedCount:document.querySelectorAll('.cell.selected').length,notice:document.querySelector('#notice')?.textContent,grid:{scrollTop:document.querySelector('#grid')?.scrollTop,scrollLeft:document.querySelector('#grid')?.scrollLeft}}))()`)));
      };
      const dispatchMouse = params => devtools.send('Input.dispatchMouseEvent', {
        ...params, x: params.x + devtools.frameOffset.x, y: params.y + devtools.frameOffset.y
      });
      await waitFor('Boolean(document.querySelector("#cell-1-0")) && !document.querySelector("#cell-2-0")', 'filter removes hidden rows while preserving the frozen header and first data row');
      await evaluate(`(() => {const grid=document.querySelector('#grid');grid.scrollTop=0;grid.scrollLeft=0;return true})()`);
      await evaluate('new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)))');
      const frozenBefore = await evaluate(`(() => {const g=document.querySelector('#grid').getBoundingClientRect();return [0,1].map(row=>{const r=document.querySelector('#cell-'+row+'-0').getBoundingClientRect();return {top:r.top-g.top,left:r.left-g.left}})})()`);
      await evaluate(`(() => {document.querySelector('#grid').scrollTop=260;document.querySelector('#grid').scrollLeft=220;return true})()`);
      await waitFor('document.querySelector("#grid").scrollTop === 260 && document.querySelector("#grid").scrollLeft === 220', 'worksheet scrolls behind the frozen panes');
      await evaluate('new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)))');
      const frozenAfter = await evaluate(`(() => {const g=document.querySelector('#grid').getBoundingClientRect();return [0,1].map(row=>{const r=document.querySelector('#cell-'+row+'-0').getBoundingClientRect();return {top:r.top-g.top,left:r.left-g.left}})})()`);
      assert.deepEqual(frozenAfter, frozenBefore, 'frozen rows and column keep their screen positions while the grid scrolls');

      await evaluate(`(() => {const grid=document.querySelector('#grid');grid.scrollTop=0;grid.scrollLeft=0;return true})()`);
      await evaluate('new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)))');
      await waitFor('Boolean(document.querySelector("#cell-1-0"))', 'frozen starting cell is rendered');
      const origin = await evaluate(`(() => {const r=document.querySelector('#cell-1-0').getBoundingClientRect(),g=document.querySelector('#grid').getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+r.height/2,edgeX:r.left+r.width/2,edgeY:g.bottom-2,start:document.querySelector('#grid').scrollTop}})()`);
      await dispatchMouse({ type: 'mousePressed', x: origin.x, y: origin.y, button: 'left', buttons: 1, clickCount: 1 });
      await dispatchMouse({ type: 'mouseMoved', x: origin.edgeX, y: origin.edgeY, button: 'left', buttons: 1 });
      await waitFor(`document.querySelector('#grid').scrollTop > ${origin.start + 20}`, 'dragging below frozen rows autoscrolls the filtered worksheet');
      const selected = await evaluate(`(() => {const id=document.querySelector('.cell.selected')?.id||'',match=/^cell-(\\d+)-(\\d+)$/.exec(id),row=match?Number(match[1]):-1;return{row,value:document.querySelector('#cell-'+row+'-0')?.textContent,address:document.querySelector('#address')?.value}})()`);
      assert.ok(selected.row > 5 && selected.value === 'keep', `drag endpoint must map to a visible source row, got ${JSON.stringify(selected)}`);
      await dispatchMouse({ type: 'mouseReleased', x: origin.edgeX, y: origin.edgeY, button: 'left', buttons: 0, clickCount: 1 });

      await evaluate(`(() => {const grid=document.querySelector('#grid');grid.scrollTop=0;grid.scrollLeft=0;return true})()`);
      await evaluate('new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)))');
      await waitFor('Boolean(document.querySelector("#cell-1-1"))', 'frozen cell B2 is rendered before horizontal drag');
      const horizontalOrigin = await evaluate(`(() => {const r=document.querySelector('#cell-1-1').getBoundingClientRect(),g=document.querySelector('#grid').getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+r.height/2,edgeX:g.right-2,scroll:document.querySelector('#grid').scrollLeft,address:document.querySelector('#address').value}})()`);
      await dispatchMouse({ type: 'mousePressed', x: horizontalOrigin.x, y: horizontalOrigin.y, button: 'left', buttons: 1, clickCount: 1 });
      await dispatchMouse({ type: 'mouseMoved', x: horizontalOrigin.edgeX, y: horizontalOrigin.y, button: 'left', buttons: 1 });
      await waitFor(`document.querySelector('#grid').scrollLeft > ${horizontalOrigin.scroll + 20}`, 'dragging in a filtered frozen pane scrolls horizontally');
      await waitFor(`document.querySelector('#address').value !== '${horizontalOrigin.address}'`, 'horizontal drag extends the selection in the filtered frozen pane');
      await dispatchMouse({ type: 'mouseReleased', x: horizontalOrigin.edgeX, y: horizontalOrigin.y, button: 'left', buttons: 0, clickCount: 1 });

      await evaluate(`(() => {const grid=document.querySelector('#grid');grid.scrollTop=0;grid.scrollLeft=0;return true})()`);
      await evaluate('new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)))');
      const fillOrigin = await evaluate(`(() => {const r=document.querySelector('#cell-1-1').getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+r.height/2}})()`);
      await dispatchMouse({ type: 'mousePressed', ...fillOrigin, button: 'left', buttons: 1, clickCount: 1 });
      await dispatchMouse({ type: 'mouseReleased', ...fillOrigin, button: 'left', buttons: 0, clickCount: 1 });
      await waitFor('Boolean(document.querySelector("#cell-1-1 .fill-handle"))', 'single-cell selection exposes the fill handle');
      await waitFor('Boolean(document.querySelector("#cell-5-3"))', 'fill destination is rendered before drag');
      const fillGeometry = await evaluate(`(() => {const h=document.querySelector('#cell-1-1 .fill-handle').getBoundingClientRect(),t=document.querySelector('#cell-5-3').getBoundingClientRect();return{start:{x:h.left+h.width/2,y:h.top+h.height/2},end:{x:t.left+t.width/2,y:t.top+t.height/2}}})()`);
      await dispatchMouse({ type: 'mousePressed', ...fillGeometry.start, button: 'left', buttons: 1, clickCount: 1 });
      await dispatchMouse({ type: 'mouseMoved', ...fillGeometry.end, button: 'left', buttons: 1 });
      await dispatchMouse({ type: 'mouseReleased', ...fillGeometry.end, button: 'left', buttons: 0, clickCount: 1 });
      let saved = analyze(await fs.readFile(file), '.xlsx').sheets[0].cells;
      assert.equal(saved.get('B4')?.value, 'v4', 'fill leaves the source file unchanged until saved');
      await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file), '.xlsx').sheets[0].cells;
      assert.equal(saved.get('B4')?.value, 'v2', 'saving commits the filled range');
      await vscode.commands.executeCommand('jinsheet.undo');
      await waitFor('document.querySelector("#cell-3-1")?.textContent === "v4"', 'Ctrl+Z restores the filled cells in the Webview');
      assert.equal(vscode.window.tabGroups.activeTabGroup.activeTab?.isDirty, true, 'undo after a save marks the restored workbook as modified');
      await vscode.commands.executeCommand('workbench.action.files.save');
      saved = analyze(await fs.readFile(file), '.xlsx').sheets[0].cells;
      assert.equal(saved.get('B4')?.value, 'v4', 'one undo restores the full pre-fill range');
      assert.equal(saved.has('D6'), false, 'undo removes cells introduced by the fill');
      assert.equal(saved.get('B3')?.value, 'v3', 'undo preserves the filtered hidden row');
      await vscode.commands.executeCommand('jinsheet.redo');
      await waitFor('document.querySelector("#cell-3-1")?.textContent === "v2"', 'Ctrl+Y reapplies the diagonal fill in the Webview');
      await vscode.commands.executeCommand('workbench.action.files.save');
      assert.equal(vscode.window.tabGroups.activeTabGroup.activeTab?.isDirty, false, 'saving after redo clears the modified marker');
      saved = analyze(await fs.readFile(file), '.xlsx').sheets[0].cells;
      assert.equal(saved.get('B4')?.value, 'v2', 'redo reapplies the fill range');
      assert.equal(saved.get('D6')?.value, 'v2', 'redo restores the full diagonal range');
      assert.equal(saved.get('B5')?.value, 'v5', 'redo still preserves filtered hidden rows');
    } finally {
      devtools?.close();
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });

  test('open DOCX, DOCM, PPTX and PPTM as JinSheet read-only custom documents', async () => {
    const extension = vscode.extensions.getExtension('jjeabydev.jinsheet');
    await extension.activate();
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'jinsheet-office-host-'));
    const docxBytes = zipSync({
      '[Content_Types].xml': new TextEncoder().encode('<Types/>'),
      'word/document.xml': new TextEncoder().encode('<w:document xmlns:w="urn:w"><w:body><w:p><w:r><w:t>Preview</w:t></w:r></w:p></w:body></w:document>')
    });
    const pptxBytes = zipSync({
      '[Content_Types].xml': new TextEncoder().encode('<Types/>'),
      'ppt/presentation.xml': new TextEncoder().encode('<p:presentation xmlns:p="urn:p" xmlns:r="urn:r"><p:sldIdLst><p:sldId r:id="r1"/></p:sldIdLst></p:presentation>'),
      'ppt/_rels/presentation.xml.rels': new TextEncoder().encode('<Relationships><Relationship Id="r1" Type="urn/slide" Target="slides/slide1.xml"/></Relationships>'),
      'ppt/slides/slide1.xml': new TextEncoder().encode('<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><a:p><a:r><a:t>Preview</a:t></a:r></a:p></p:sld>')
    });
    try {
      for (const [extensionName, bytes] of [['docx', docxBytes], ['docm', docxBytes], ['pptx', pptxBytes], ['pptm', pptxBytes]]) {
        const file = path.join(directory, `host-smoke.${extensionName}`);
        await fs.writeFile(file, bytes);
        await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(file), 'jinsheet.editor');
        await new Promise(resolve => setTimeout(resolve, 250));
        const tab = vscode.window.tabGroups.activeTabGroup.activeTab;
        assert.ok(tab, `${extensionName} custom editor tab opened`);
        assert.equal(tab.input.viewType, 'jinsheet.editor');
        assert.equal(normalizePath(tab.input.uri.fsPath), normalizePath(file));
        await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
      }
    } finally {
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });
});
