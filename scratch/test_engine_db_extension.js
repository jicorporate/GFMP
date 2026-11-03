const fs = require('fs');
const path = require('path');

global.window = global;
global.localStorage = {
  _data: {},
  getItem(k) { return this._data[k] || null; },
  setItem(k, v) { this._data[k] = String(v); },
  removeItem(k) { delete this._data[k]; }
};

const baselineCode = fs.readFileSync(path.join(__dirname, '../mobile-web/baseline_data.js'), 'utf8');
eval(baselineCode);

const engineCode = fs.readFileSync(path.join(__dirname, '../mobile-web/analytics-engine.js'), 'utf8');
eval(engineCode);

const engine = window.gfmpEngine;
engine.init();

console.log('Engine initialized with lang:', engine.lang);
console.log('Account 1:', engine.data.comptes[0]);
console.log('Clean asset name:', engine.cleanAssetName(engine.data.comptes[0].nom));
