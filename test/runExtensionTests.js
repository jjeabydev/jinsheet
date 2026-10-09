'use strict';
const path = require('node:path');
const { downloadAndUnzipVSCode, runTests } = require('@vscode/test-electron');

async function main() {
  const vscodeExecutablePath = process.env.VSCODE_EXECUTABLE || await downloadAndUnzipVSCode('stable');
  const launchArgs = ['--disable-extensions', '--disable-gpu', `--user-data-dir=${path.join(require('node:os').tmpdir(), 'jinsheet-host-user-data')}`];
  if (process.platform === 'linux') launchArgs.push('--no-sandbox');
  const code = await runTests({
    vscodeExecutablePath,
    extensionDevelopmentPath: path.resolve(__dirname, '..'),
    extensionTestsPath: path.resolve(__dirname, 'host'),
    launchArgs
  });
  if (code !== 0) process.exitCode = code || 1;
}

main().catch(error => { console.error(error); process.exitCode = 1; });
