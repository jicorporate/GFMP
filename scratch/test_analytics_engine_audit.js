const fs = require('fs');
global.window = {};
require('../mobile-web/baseline_data.js');

// Load analytics-engine.js and test
let engineCode = fs.readFileSync('mobile-web/analytics-engine.js', 'utf8');
const GFMPEngine = eval(engineCode + '; GFMPEngine;');
const eng = new GFMPEngine();
eng.init();

console.log('--- AUDITING CURRENT ENGINE CALCULATIONS ---');
eng.year = '2026';
eng.month = '02';
const b2 = eng.computeBudgetZBB();
console.log('Feb 2026 Budget:', { alloue: b2.totalAlloue, consomme: b2.totalConsomme, ecart: b2.totalEcart, rate: b2.consumptionRate, count: b2.categories.length });

const nw = eng.computeNetWorth();
console.log('Net Worth Feb 2026:', { assets: nw.totalAssets, debts: nw.totalLiabilities, net: nw.netWorth });
nw.accounts.forEach(a => console.log('  ' + a.nom + ': ' + a.soldeConverti + ' ' + a.deviseNative));

const cf = eng.computeCashflow();
console.log('Cashflow Feb 2026:', { rev: cf.totalRevenus, dep: cf.totalDepenses, net: cf.netCashflow, rate: cf.savingsRate });
