'use strict';
const assert = require('assert');
const vscode = require('vscode');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const { zipSync } = require('../../vendor/node_modules/fflate');

suite('JinSheet Extension Host', () => {
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
      assert.equal(tab.input.uri.fsPath, file);
    } finally {
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
        assert.equal(tab.input.uri.fsPath, file);
        await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
      }
    } finally {
      await vscode.commands.executeCommand('workbench.action.closeAllEditors');
      await fs.rm(directory, { recursive: true, force: true });
    }
  });
});
