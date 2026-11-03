// Test Header Filters Interactivity: Year, Month, Currency, Language
const fs = require('fs');
const path = require('path');

// Mock browser globals
global.window = global;
global.localStorage = {
  _data: {},
  getItem(k) { return this._data[k] || null; },
  setItem(k, v) { this._data[k] = String(v); },
  removeItem(k) { delete this._data[k]; }
};

// Load baseline data and engine
const baselineCode = fs.readFileSync(path.join(__dirname, '../mobile-web/baseline_data.js'), 'utf8');
eval(baselineCode);

const engineCode = fs.readFileSync(path.join(__dirname, '../mobile-web/analytics-engine.js'), 'utf8');
eval(engineCode);

const engine = window.gfmpEngine;
engine.init();

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
console.log('Month 05 label:', engine.getMonthLabel());
if (engine.month !== '05' || !engine.getMonthLabel().includes('MAI')) throw new Error('setMonth(05) failed');

const cfMay = engine.computeCashflow();
console.log('May Cashflow tx count:', cfMay.transactions.length);

engine.setMonth('00'); // Full Year
console.log('Month 00 label:', engine.getMonthLabel());
const cfYear = engine.computeCashflow();
const bgYear = engine.computeBudgetZBB();
console.log('Full Year Cashflow tx count:', cfYear.transactions.length, 'Total Rev:', cfYear.totalRevenus, 'Total Dep:', cfYear.totalDepenses);
console.log('Full Year Budget total alloue:', bgYear.totalAlloue, 'total consomme:', bgYear.totalConsomme, 'mois label:', bgYear.mois);

if (cfYear.transactions.length < cfMay.transactions.length) {
  throw new Error('Annual aggregation should include at least as many transactions as May');
}
if (!bgYear.mois.includes('2026') || !bgYear.mois.includes('ANNÉE')) {
  throw new Error('Budget annual month label incorrect: ' + bgYear.mois);
}

console.log('\n=== TEST 4: Multi-Currency O(1) Switching ===');
const currencies = ['MUR', 'EUR', 'USD', 'GBP', 'ZAR', 'XOF'];
const expectedSymbols = { MUR: 'Rs', EUR: '€', USD: '$', GBP: '£', ZAR: 'R', XOF: 'CFA' };

currencies.forEach(curr => {
  engine.setCurrency(curr);
  const sym = engine.getCurrencySymbol();
  const formatted = engine.formatMoney(1000);
  console.log(`Currency: ${curr} -> Symbol: ${sym} -> Formatted 1000: "${formatted}"`);
  if (sym !== expectedSymbols[curr]) throw new Error(`Symbol mismatch for ${curr}: got ${sym}, expected ${expectedSymbols[curr]}`);
  if (!formatted.includes(expectedSymbols[curr])) throw new Error(`Formatted money missing symbol ${expectedSymbols[curr]}`);
  
  // Verify Net Worth and Cashflow in this currency
  const nw = engine.computeNetWorth();
  const cf = engine.computeCashflow();
  if (typeof nw.netWorth !== 'number' || typeof cf.netCashflow !== 'number') {
    throw new Error(`Calculation failed in currency ${curr}`);
  }
});

// Restore MUR and Feb
engine.setCurrency('MUR');
engine.setMonth('02');

console.log('\n=== TEST 5: Multilingual Translations (FR, EN, ES, PT) ===');
const languages = ['FR', 'EN', 'ES', 'PT'];
const testKeys = [
  'NAV_HOME', 'NAV_SAISIE', 'NAV_CASHFLOW', 'NAV_NETWORTH', 'NAV_BUDGET', 'NAV_SETTINGS',
  'REPORT_TITLE', 'REPORT_TOT_REV', 'REPORT_TOT_DEP', 'REPORT_NET_CF', 'REPORT_SAVINGS_RATE',
  'TREND_CHART_TITLE', 'BUDGET_TITLE', 'NETWORTH_TITLE', 'TH_DATE', 'TH_MONTANT', 'STATUS_OVER', 'STATUS_OK'
];

languages.forEach(lang => {
  engine.setLanguage(lang);
  console.log(`\nTesting translations for language [${lang}]:`);
  testKeys.forEach(k => {
    const translation = engine.tr(k);
    if (!translation || translation === k) {
      throw new Error(`Untranslated or missing key ${k} in language ${lang}`);
    }
  });
  console.log(`  NAV_HOME: "${engine.tr('NAV_HOME')}", REPORT_TITLE: "${engine.tr('REPORT_TITLE')}", Month 02: "${engine.getMonthLabel()}"`);
});

console.log('\n=== TEST 6: Executive Financial Report Reactivity ===');
engine.setLanguage('EN');
engine.setCurrency('USD');
engine.setMonth('02');
engine.setYear('2026');
const reportEn = engine.generateFinancialReport();
console.log('Report EN in USD for Feb 2026:');
console.log('  Period:', reportEn.period);
console.log('  Currency:', reportEn.currency);
console.log('  Revenues:', reportEn.revenus);
console.log('  Expenses:', reportEn.depenses);
console.log('  Net Worth:', reportEn.netWorth);

if (!reportEn.period.includes('FEBRUARY') || reportEn.currency !== 'USD' || !reportEn.revenus.includes('$')) {
  throw new Error('Report did not reflect active filters');
}

console.log('\n>>> ALL 6 INTERACTIVITY TESTS PASSED WITH 0 ERRORS! <<<');
