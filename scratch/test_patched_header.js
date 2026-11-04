// Test Header Filters Interactivity with Patched Engine
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

const engineCode = fs.readFileSync(path.join(__dirname, '../scratch/patched_engine.js'), 'utf8');
eval(engineCode);

const engine = window.gfmpEngine;
const sampleFixture = JSON.parse(fs.readFileSync(path.join(__dirname, 'sample_test_dataset.json'), 'utf8'));
window.GFMP_BASELINE = sampleFixture;
engine.init(sampleFixture);

console.log('=== TEST 1: Initial State ===');
console.log('Year:', engine.year, 'Month:', engine.month, 'Currency:', engine.currency, 'Lang:', engine.lang);
if (engine.year !== '2026' || engine.currency !== 'MUR' || engine.lang !== 'FR') {
  throw new Error('Default state mismatch');
}

console.log('\n=== TEST 2: Year Stepper Interactivity ===');
let notifyCount = 0;
engine.subscribe(() => { notifyCount++; });
engine.nextYear();
console.log('After nextYear -> Year:', engine.year, 'NotifyCount:', notifyCount);
if (engine.year !== '2027' || notifyCount !== 1) throw new Error('nextYear failed');

engine.prevYear();
console.log('After prevYear -> Year:', engine.year, 'NotifyCount:', notifyCount);
if (engine.year !== '2026' || notifyCount !== 2) throw new Error('prevYear failed');

console.log('\n=== TEST 3: Month Strip Interactivity & Annual Aggregation (00) ===');
engine.setMonth('05');
console.log('Month 05 label:', engine.tr(engine.getMonthLabel('05')));
const cfMay = engine.computeCashflow();
console.log('May Cashflow tx count:', cfMay.transactions.length);

engine.setMonth('00');
console.log('Month 00 label:', engine.tr(engine.getMonthLabel('00')));
const cfYear = engine.computeCashflow();
console.log('Full Year Cashflow tx count:', cfYear.transactions.length, 'Total Rev:', cfYear.totalRevenus, 'Total Dep:', cfYear.totalDepenses);
if (cfYear.transactions.length === 0) throw new Error('Month 00 should aggregate all 2026 transactions');

const bgYear = engine.computeBudgetZBB();
console.log('Full Year Budget total alloue:', bgYear.totalAlloue, 'total consomme:', bgYear.totalConsomme, 'mois label:', bgYear.mois);
if (!bgYear.mois.includes('ANNÉE ENTIÈRE')) throw new Error('Budget ZBB label should reflect full year');

console.log('\n=== TEST 4: Multi-Currency O(1) Switching ===');
const currencies = ['MUR', 'EUR', 'USD', 'GBP', 'ZAR', 'XOF'];
currencies.forEach(curr => {
  engine.setCurrency(curr);
  const sym = engine.getCurrencySymbol();
  const formatted = engine.formatMoney(1000);
  console.log(`Currency: ${curr} -> Symbol: ${sym} -> Formatted 1000: "${formatted}"`);
  if (!formatted.includes(sym)) throw new Error(`Formatted money missing symbol for ${curr}`);
});

console.log('\n=== TEST 5: Multilingual Translations (FR, EN, ES, PT) ===');
const langs = ['FR', 'EN', 'ES', 'PT'];
langs.forEach(l => {
  engine.setLanguage(l);
  console.log(`\nTesting translations for language [${l}]:`);
  console.log(`  NAV_HOME: "${engine.tr('NAV_HOME')}", REPORT_TITLE: "${engine.tr('REPORT_TITLE')}", Month 02: "${engine.tr('M02')}"`);
  console.log(`  DB Account 1: "${engine.cleanAssetName('Compte Courant Principal')}"`);
  console.log(`  DB Cat 5: "${engine.trEntity('Alimentation & Supermarché')}"`);
  console.log(`  DB Note: "${engine.trEntity('Premier loyer résidence')}"`);
  if (!engine.tr('NAV_HOME') || !engine.tr('REPORT_TITLE')) throw new Error(`Missing translations for ${l}`);
});

console.log('\n>>> ALL TESTS PASSED WITH 0 ERRORS! <<<');
