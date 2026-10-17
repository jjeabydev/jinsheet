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
    const resourceCandidates = [];
    let directory = path.dirname(vscodeExecutablePath);
    for (let depth = 0; depth < 6; depth++) {
      resourceCandidates.push(
        path.join(directory, 'resources', 'app', 'out', 'nls.messages.json'),
        path.join(directory, 'resources', 'app', 'out', 'vs', 'nls.messages.json'),
        path.join(directory, 'Contents', 'Resources', 'app', 'out', 'nls.messages.json'),
        path.join(directory, 'Contents', 'Resources', 'app', 'out', 'vs', 'nls.messages.json'),
        path.join(directory, 'resources', 'app', 'out', 'vs', 'nls.messages.json')
      );
      if (process.platform === 'win32') {
        resourceCandidates.push(
          path.join(directory, 'resources', 'app', 'out', 'vs', 'nls.messages.json'),
          path.join(directory, 'resources', 'app', 'out', 'nls.messages.json')
        );
      }
      const parent = path.dirname(directory);
      if (parent === directory) break;
      directory = parent;
    }
    const defaultMessagesFile = resourceCandidates.find(candidate => fs.existsSync(candidate));
    if (!defaultMessagesFile) throw new Error('Could not locate the VS Code default NLS messages for the requested host locale.');
    const cliPath = path.join(path.dirname(defaultMessagesFile), 'cli.js');
    if (!fs.existsSync(cliPath)) throw new Error('Could not locate the VS Code CLI for installing the requested language pack.');
    const install = spawnSync(vscodeExecutablePath, [cliPath, `--user-data-dir=${userDataDir}`, `--extensions-dir=${extensionsDir}`, '--install-extension=MS-CEINTL.vscode-language-pack-ko', '--force'], {
      encoding: 'utf8',
      timeout: 180000,
      env: { ...process.env, ELECTRON_RUN_AS_NODE: '1' }
    });
    if (install.error || install.status !== 0) throw new Error(`Could not install the VS Code Korean language pack: ${install.error?.message || install.stderr || install.stdout || `exit ${install.status}`}`);
    const languagePackDirectory = fs.readdirSync(extensionsDir, { withFileTypes: true })
      .filter(entry => entry.isDirectory() && entry.name.toLowerCase().startsWith('ms-ceintl.vscode-language-pack-ko-'))
      .map(entry => path.join(extensionsDir, entry.name))
      .find(candidate => fs.existsSync(path.join(candidate, 'package.json')));
    if (!languagePackDirectory) throw new Error('The VS Code Korean language pack was not installed in the isolated test extension directory.');
    const languagePack = JSON.parse(fs.readFileSync(path.join(languagePackDirectory, 'package.json'), 'utf8'));
    const localization = languagePack.contributes?.localizations?.find(item => item.languageId === 'ko');
    const translationsConfigFile = localization?.translations?.find(item => item.id === 'vscode')?.path;
    if (!translationsConfigFile) throw new Error('The VS Code Korean language pack does not declare its core translations.');
    const nlsConfig = {
      userLocale: 'ko',
      osLocale: Intl.DateTimeFormat().resolvedOptions().locale || 'en',
      resolvedLanguage: 'ko',
      defaultMessagesFile,
      languagePack: { translationsConfigFile: path.join(languagePackDirectory, translationsConfigFile) }
    };
    process.env.VSCODE_NLS_CONFIG = JSON.stringify(nlsConfig);
    process.env.VSCODE_PORTABLE = userDataDir;
    fs.writeFileSync(path.join(userDataDir, 'argv.json'), JSON.stringify({ locale: 'ko' }));
  }
  const launchArgs = [...(locale ? [] : ['--disable-extensions']), '--disable-gpu', '--remote-debugging-port=9333', '--remote-allow-origins=*', `--user-data-dir=${userDataDir}`];
  if (locale) launchArgs.push(`--locale=${locale.split('-')[0]}`, `--extensions-dir=${extensionsDir}`);
  if (process.platform === 'linux') launchArgs.push('--no-sandbox');
  if (locale) Object.assign(process.env, { LANG: 'ko_KR.UTF-8', LANGUAGE: 'ko_KR:ko', LC_ALL: 'ko_KR.UTF-8' });
  const code = await runTests({
    vscodeExecutablePath,
    extensionDevelopmentPath: path.resolve(__dirname, '..'),
    extensionTestsPath: path.resolve(__dirname, 'host'),
    launchArgs
  });
  if (code !== 0) process.exitCode = code || 1;
}

main().catch(error => { console.error(error); process.exitCode = 1; });
