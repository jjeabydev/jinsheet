'use strict';
const Mocha = require('mocha');
const path = require('node:path');

async function run() {
  const mocha = new Mocha({ ui: 'tdd', color: true, timeout: 30000 });
  mocha.addFile(path.resolve(__dirname, 'suite.js'));
  return new Promise((resolve, reject) => mocha.run(failures => failures ? reject(new Error(`${failures} Extension Host test(s) failed`)) : resolve()));
}
module.exports = { run };
