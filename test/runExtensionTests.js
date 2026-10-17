'use strict';
const path = require('node:path');
const fs = require('node:fs');
const { spawnSync } = require('node:child_process');
const { downloadAndUnzipVSCode, runTests } = require('@vscode/test-electron');

async function main() {
  const vscodeExecutablePath = process.env.VSCODE_EXECUTABLE || await downloadAndUnzipVSCode('stable');
  const locale = process.env.JINSHEET_HOST_LOCALE;
  if (locale && !locale.toLowerCase().startsWith('ko')) throw new Error('JINSHEET_HOST_LOCALE currently supports the Korean locale only.');
  const localeSuffix = locale ? '-' + locale.toLowerCase().replace(/[^a-z0-9-]/g, '-') : '';
  const userDataDir = path.join(require('node:os').tmpdir(), 'jinsheet-host-user-data' + localeSuffix);
  const extensionsDir = path.join(require('node:os').tmpdir(), 'jinsheet-host-extensions' + localeSuffix);
  if (locale) {
    fs.mkdirSync(userDataDir, { recursive: true });
    const resourceCandidates = [
      path.join(path.dirname(vscodeExecutablePath), 'resources', 'app', 'out', 'nls.messages.json'),
      path.resolve(vscodeExecutablePath, '..', '..', 'Resources', 'app', 'out', 'nls.messages.json'),
      path.resolve(vscodeExecutablePath, '..', '..', '..', 'resources', 'app', 'out', 'nls.messages.json')
    ];
    const defaultMessagesFile = resourceCandidates.find(candidate => fs.existsSync(candidate));
    if (!defaultMessagesFile) throw new Error('Could not locate the VS Code default NLS messages for the requested host locale.');
    const cliPath = path.join(path.dirname(defaultMessagesFile), 'cli.js');
    if (!fs.existsSync(cliPath)) throw new Error('Could not locate the VS Code CLI for installing the requested language pack.');
    fs.mkdirSync(path.join(userDataDir, 'User'), { recursive: true });
    fs.writeFileSync(path.join(userDataDir, 'User', 'locale.json'), JSON.stringify({ locale: locale.split('-')[0] }));
    const install = spawnSync(vscodeExecutablePath, [cliPath, `--user-data-dir=${userDataDir}`, `--extensions-dir=${extensionsDir}`, '--install-extension=MS-CEINTL.vscode-language-pack-ko', '--force'], {
      encoding: 'utf8',
      timeout: 180000,
      env: { ...process.env, ELECTRON_RUN_AS_NODE: '1' }
    });
    if (install.error || install.status !== 0) throw new Error(`Could not install the VS Code Korean language pack: ${install.error?.message || install.stderr || install.stdout || `exit ${install.status}`}`);
  }
  const launchArgs = [...(locale ? [] : ['--disable-extensions']), '--disable-gpu', '--remote-debugging-port=9333', '--remote-allow-origins=*', `--user-data-dir=${userDataDir}`];
  if (locale) launchArgs.push(`--locale=${locale.split('-')[0]}`, `--extensions-dir=${extensionsDir}`);
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
