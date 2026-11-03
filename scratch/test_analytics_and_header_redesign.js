/**
 * Comprehensive Automated Test Suite:
 * 1. Analytics Engine Audit & Mathematical Precision (ZBB, Net Worth Wealth Accumulation, Transfers, Rounding, Solvency)
 * 2. Header Redesign & Settings Button Removal Verification
 */

const fs = require('fs');
const path = require('path');
const assert = require('assert');

console.log('=== STARTING TEST SUITE: ANALYTICS ENGINE & HEADER REDESIGN ===\n');

// 1. Verify HTML Structure (Header Redesign & Settings Button Removal)
const html = fs.readFileSync(path.join(__dirname, '../mobile-web/index.html'), 'utf8');

assert(!html.includes('id="btn-header-settings"'), 'FAIL: id="btn-header-settings" should be completely removed from index.html');
assert(!html.includes('btn-header-settings'), 'FAIL: class="btn-header-settings" should not exist in index.html');
console.log('✔ PASS: Header settings button successfully and cleanly removed from index.html');

assert(html.includes('class="lang-selector" role="group" aria-label="Sélecteur de langue"'), 'FAIL: lang-selector should have accessible role="group" and aria-label');
console.log('✔ PASS: Header language selector has accessible group role and aria-label');

assert(html.includes('styles.css?v=5.5.0'), 'FAIL: styles.css should be cache-busted with ?v=5.5.0');
assert(html.includes('analytics-engine.js?v=5.5.0'), 'FAIL: analytics-engine.js should be cache-busted with ?v=5.5.0');
assert(html.includes('app.js?v=5.5.0'), 'FAIL: app.js should be cache-busted with ?v=5.5.0');
console.log('✔ PASS: Cache busters in index.html verified at ?v=5.5.0');

// 2. Verify CSS Styles
const css = fs.readFileSync(path.join(__dirname, '../mobile-web/styles.css'), 'utf8');
assert(!css.includes('.btn-header-settings'), 'FAIL: .btn-header-settings rules should be removed from styles.css');
assert(css.includes('mask-image: linear-gradient(to right, transparent 0px, black 8px, black calc(100% - 8px), transparent 100%)'), 'FAIL: horizontal scroll containers should have edge mask in styles.css');
console.log('✔ PASS: styles.css clean of obsolete settings styles and has modern edge fading masks');

// 3. Verify Service Worker
const sw = fs.readFileSync(path.join(__dirname, '../mobile-web/sw.js'), 'utf8');
assert(sw.includes("const CACHE_NAME = 'gfmp-mobile-v5.5.0';"), 'FAIL: sw.js CACHE_NAME should be gfmp-mobile-v5.5.0');
console.log('✔ PASS: Service worker cache version bumped to gfmp-mobile-v5.5.0');

// 4. Test Analytics Engine
const store = {};
global.window = {
  localStorage: {
    getItem: (k) => store[k] || null,
    setItem: (k, v) => { store[k] = String(v); },
    removeItem: (k) => { delete store[k]; }
  }
};
require('../mobile-web/baseline_data.js');
const engineCode = fs.readFileSync(path.join(__dirname, '../mobile-web/analytics-engine.js'), 'utf8');
eval(engineCode);

const engine = window.gfmpEngine;
engine.init(global.window.GFMP_BASELINE_DATA);

// Test 4.1: Floating point rounding helper
assert.strictEqual(engine.roundMoney(0.1 + 0.2), 0.3, 'roundMoney(0.1 + 0.2) must be exactly 0.3');
assert.strictEqual(engine.roundMoney(10.005), 10.01, 'roundMoney(10.005) must round up correctly to 10.01');
assert.strictEqual(engine.roundMoney(-0.0001), 0, 'roundMoney(-0.0001) must round to 0 without negative sign');
console.log('✔ PASS: roundMoney IEEE 754 precision helper verified');

// Test 4.2: formatMoney no negative zero
const formattedZero = engine.formatMoney(-0.00001);
assert(!formattedZero.includes('-0'), `formatMoney must not output -0, got: ${formattedZero}`);
console.log('✔ PASS: formatMoney eliminates negative zero (-0,00)');

// Test 4.3: Cashflow computations
const cfFeb = engine.computeCashflow('2026-02');
console.log('Cashflow Feb 2026:', cfFeb);
assert(cfFeb.totalRevenus > 0, 'Revenues should be positive');
assert(cfFeb.totalDepenses > 0, 'Expenses should be positive');
assert.strictEqual(engine.roundMoney(cfFeb.totalRevenus - cfFeb.totalDepenses), cfFeb.netCashflow, 'netCashflow must strictly equal totalRevenus - totalDepenses');
console.log('✔ PASS: Cashflow mathematical consistency verified');

// Test 4.4: Net Worth & Transfers
const nwFeb = engine.computeNetWorth('2026-02');
console.log('Net Worth Feb 2026:', {
  totalAssets: nwFeb.totalAssets,
  totalLiabilities: nwFeb.totalLiabilities,
  netWorth: nwFeb.netWorth,
  accountsCount: nwFeb.accounts.length
});

assert.strictEqual(engine.roundMoney(nwFeb.totalAssets - nwFeb.totalLiabilities), nwFeb.netWorth, 'netWorth must strictly equal totalAssets - totalLiabilities');

// Checking account (Compte Courant id 1) should reflect the transfer outflow (77,252 instead of inflated 177,252)
const cptCourant = nwFeb.accounts.find(c => c.id === '1');
assert(cptCourant, 'Compte Courant must exist');
console.log('Compte Courant balance:', cptCourant.soldeConverti);
assert.strictEqual(cptCourant.soldeConverti, 77252, `Compte Courant should be 77,252 MUR, got: ${cptCourant.soldeConverti}`);

// Livret d'Epargne (id 2) should reflect the 100,000 MUR transfer inflow
const cptEpargne = nwFeb.accounts.find(c => c.id === '2');
assert(cptEpargne, 'Livret d Epargne must exist');
console.log('Livret d Epargne balance:', cptEpargne.soldeConverti);
assert.strictEqual(cptEpargne.soldeConverti, 100000, `Livret d Epargne should have received 100,000 MUR transfer, got: ${cptEpargne.soldeConverti}`);

// Crypto Wallet (id 5) should reflect the 250 USD transfer inflow
const cptCrypto = nwFeb.accounts.find(c => c.id === '5');
assert(cptCrypto, 'Crypto Portfolio must exist');
console.log('Crypto Portfolio balance (MUR converted):', cptCrypto.soldeConverti);
assert(cptCrypto.soldeConverti > 0, `Crypto Portfolio should have received transfer, got: ${cptCrypto.soldeConverti}`);
console.log('✔ PASS: Net Worth wealth accumulation & transfer directionality verified');

// Test 4.5: Budget ZBB computations
const zbbFeb = engine.computeBudgetZBB('2026-02');
console.log('Budget ZBB Feb 2026:', {
  totalAlloue: zbbFeb.totalAlloue,
  totalConsomme: zbbFeb.totalConsomme,
  variance: zbbFeb.totalEcart,
  consumptionRate: zbbFeb.consumptionRate,
  categoriesCount: zbbFeb.categories.length
});

assert.strictEqual(zbbFeb.totalAlloue, 55000, `Expected total allocated to be 55,000 MUR from baseline, got: ${zbbFeb.totalAlloue}`);
assert(zbbFeb.totalConsomme > 0, 'Total consumed should be positive');
assert.strictEqual(engine.roundMoney(zbbFeb.totalAlloue - zbbFeb.totalConsomme), zbbFeb.totalEcart, 'Budget variance must equal totalAlloue - totalConsomme');
console.log('✔ PASS: Budget ZBB loads all baseline data and computes exact variances');

// Test 4.6: Consolidated Hub Data & Solvency
const hubData = engine.getConsolidatedHubData('2026-02');
console.log('Hub Consolidated Data:', {
  healthScore: hubData.healthScore,
  debtRatio: hubData.debtRatio,
  totalAssets: hubData.totalAssets,
  totalLiabilities: hubData.totalLiabilities
});

assert(hubData.healthScore >= 10 && hubData.healthScore <= 100, `Health score must be between 10 and 100, got: ${hubData.healthScore}`);
console.log('✔ PASS: Consolidated Hub Data & Financial Health Score verified');

console.log('\n=== ALL TESTS PASSED SUCCESSFULLY! ZERO REGRESSION! ===');
