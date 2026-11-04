/**
 * GFMP (SFP v1.0) — Automated Verification Suite:
 * 1. User Choice to Use or Not Use Test Financial Data (DEMO vs EMPTY / Portefeuille Vierge)
 *    - Step 5 in Startup Configuration Modal (#modal-initialization -> #init-datamode-section / #init-datamode-grid)
 *    - 1-Click Switcher in Settings (#view-settings -> #settings-datamode-selector) & Hub Banner (#hub-datamode-banner)
 *    - Persistence of EMPTY mode across reloads in localStorage (gfmp_dataset_v3 & gfmp_data_mode) without overwriting with GFMP_BASELINE
 * 2. Dynamic, Robust, Resilient & Optimized Tables, Visualisations, Reports & Pages When Zero Financial Data Exists
 *    - Zero NaN / Infinity / -0 / undefined across Cashflow, Trend, Donut, Net Worth, ZBB Budget, Consolidated Hub, and Executive Report
 *    - Neutral Health Score (70/100) when netCashflow === 0 (no false -15 penalty)
 *    - Empty topAssets ([]) when all accounts have 0.00 balance
 *    - Empty ZBB Budget categories ([]) when alloue === 0 && consomme === 0 (while keeping 6 categories in 2026-02 Demo Mode)
 *    - Non-collapsing Zero-State Donut SVG (.empty-state-donut-wrap) & Trend Callout (.trend-empty-callout)
 *    - Rich Interactive Empty States (.empty-state-rich) with delegated [data-empty-action] CTAs
 */

const fs = require('fs');
const path = require('path');
const assert = require('assert');

const ROOT = path.resolve(__dirname, '..');
const MOBILE_WEB = path.join(ROOT, 'mobile-web');

const html = fs.readFileSync(path.join(MOBILE_WEB, 'index.html'), 'utf8');
const css = fs.readFileSync(path.join(MOBILE_WEB, 'styles.css'), 'utf8');
const appJs = fs.readFileSync(path.join(MOBILE_WEB, 'app.js'), 'utf8');
const engineCode = fs.readFileSync(path.join(MOBILE_WEB, 'analytics-engine.js'), 'utf8');
const baselineCode = fs.readFileSync(path.join(MOBILE_WEB, 'baseline_data.js'), 'utf8');

const store = {};
global.window = {
  localStorage: {
    getItem: (k) => (k in store ? store[k] : null),
    setItem: (k, v) => { store[k] = String(v); },
    removeItem: (k) => { delete store[k]; },
    clear: () => { for (const k of Object.keys(store)) delete store[k]; }
  }
};
global.localStorage = global.window.localStorage;

eval(baselineCode);
eval(engineCode);

const engine = window.gfmpEngine;

console.log('================================================================================');
console.log('🧪 TEST SUITE: DATA MODE CHOICE (DEMO VS EMPTY) & ZERO-DATA UI/ENGINE RESILIENCE');
console.log('================================================================================\n');

// 1. Verify Step 5 Data Mode Choice in #modal-initialization & #view-settings & #hub-datamode-banner
console.log('--- 1. Onboarding Step 5, Settings Switcher & Hub Banner DOM/CSS/i18n ---');
assert(html.includes('id="init-datamode-section"') && html.includes('id="init-datamode-grid"'), 'index.html must include Step 5 #init-datamode-section and #init-datamode-grid');
assert(html.includes('data-init-datamode="demo"') && html.includes('data-init-datamode="empty"'), 'index.html must include demo and empty choice cards in #init-datamode-grid');
assert(html.includes('id="hub-datamode-banner"'), 'index.html #view-home must include #hub-datamode-banner');
assert(html.includes('id="settings-datamode-card"') && html.includes('id="btn-datamode-demo"') && html.includes('id="btn-datamode-empty"'), 'index.html #view-settings must include #settings-datamode-card switcher');
assert(css.includes('.empty-state-rich') && css.includes('.empty-state-donut-wrap') && css.includes('.hub-datamode-banner') && css.includes(':has(.empty-state-rich)'), 'styles.css must define .empty-state-rich, .empty-state-donut-wrap, .hub-datamode-banner, and :has(.empty-state-rich)');

const requiredKeys = [
  'INIT_STEP5_BADGE', 'INIT_STEP5_TITLE', 'INIT_STEP5_DESC',
  'DATA_MODE_DEMO_BADGE', 'DATA_MODE_DEMO_TITLE', 'DATA_MODE_DEMO_SUB',
  'DATA_MODE_EMPTY_BADGE', 'DATA_MODE_EMPTY_TITLE', 'DATA_MODE_EMPTY_SUB',
  'SOLVENCY_EMPTY', 'TREND_EMPTY_TITLE', 'TREND_EMPTY_SUB', 'DONUT_EMPTY_HINT',
  'EMPTY_PATRIMOINE_SUB', 'EMPTY_BUDGET_SUB', 'EMPTY_TX_SUB',
  'EMPTY_CTA_ADD_TX', 'EMPTY_CTA_ADD_EXP', 'EMPTY_CTA_ALLOC_BG', 'EMPTY_CTA_ADD_ACC',
  'EMPTY_CTA_RESET_FILTER', 'EMPTY_CTA_FULL_YEAR', 'EMPTY_CTA_LOAD_DEMO', 'EMPTY_CTA_SWITCH_EMPTY',
  'BANNER_DEMO_ACTIVE', 'BANNER_EMPTY_ACTIVE',
  'SETTINGS_DATAMODE_TITLE', 'SETTINGS_DATAMODE_DESC',
  'REPORT_EMPTY_DIAG_TITLE', 'REPORT_EMPTY_DIAG_DESC'
];
['FR', 'EN', 'ES', 'PT'].forEach(lang => {
  requiredKeys.forEach(k => {
    const val = engine.tr(k, {}, lang);
    assert(val && val !== k, `Missing ${lang} translation for key ${k}`);
  });
});
console.log('✔ PASS: Onboarding Step 5, Settings Data Mode Switcher, Hub Banner & Quadrilingual keys verified.');

// 2. Verify initializeSystem with useDemoData = false ('EMPTY') vs useDemoData = true ('DEMO')
console.log('\n--- 2. Engine Data Mode Initialization & Persistence Across Reloads ---');
global.localStorage.clear();
const initEmptyRes = engine.initializeSystem('EUR', 'FR', '2026-02-27', 'emerald', '2026', '02', false);
assert.strictEqual(initEmptyRes.dataMode, 'EMPTY', 'initializeSystem(..., false) must return dataMode === EMPTY');
assert.strictEqual(engine.getDataMode(), 'EMPTY', 'engine.getDataMode() must be EMPTY');
assert.strictEqual(engine.data.transactions.length, 0, 'Empty portfolio must have 0 transactions');
assert.strictEqual(engine.data.budgets.length, 0, 'Empty portfolio must have 0 budgets');
assert.strictEqual(engine.data.comptes.length, 9, 'Empty portfolio must preserve the 9 reference accounts ready for use');
assert.strictEqual(engine.data.categories.length, 11, 'Empty portfolio must preserve the 11 reference categories ready for use');
assert.strictEqual(engine.data.tiers.length, 7, 'Empty portfolio must preserve the 7 reference counterparties ready for use');
assert.strictEqual(engine.hasFinancialData(), false, 'engine.hasFinancialData() must be false when portfolio is empty');

// Simulate page reload (engine.init() with no arguments reading from localStorage)
engine.data = null;
engine.init();
assert.strictEqual(engine.getDataMode(), 'EMPTY', 'After reload, engine.getDataMode() must remain EMPTY (not overwritten by GFMP_BASELINE)');
assert.strictEqual(engine.data.transactions.length, 0, 'After reload, transactions.length must remain 0');
assert.strictEqual(engine.data.budgets.length, 0, 'After reload, budgets.length must remain 0');
console.log('✔ PASS: Empty portfolio persists across reloads in localStorage without being overwritten by baseline demo data.');

// 3. Verify Zero-Financial-Data Analytical Resilience (0 NaN, 0 Infinity, 0 -0, clean empty arrays)
console.log('\n--- 3. Zero-Financial-Data Analytical Resilience Across All Modules ---');
const cfEmpty = engine.computeCashflow('2026-02');
assert.strictEqual(cfEmpty.totalRevenus, 0, 'Empty cashflow totalRevenus === 0');
assert.strictEqual(cfEmpty.totalDepenses, 0, 'Empty cashflow totalDepenses === 0');
assert.strictEqual(cfEmpty.netCashflow, 0, 'Empty cashflow netCashflow === 0');
assert.strictEqual(cfEmpty.savingsRate, 0, 'Empty cashflow savingsRate === 0');
assert.strictEqual(cfEmpty.transactions.length, 0, 'Empty cashflow transactions === []');
assert.strictEqual(cfEmpty.breakdown.length, 0, 'Empty cashflow breakdown === []');

const trendEmpty = engine.computeCashflowTrend(6, '2026-02');
assert.strictEqual(trendEmpty.length, 6, 'Empty 6-month trend still returns 6 month slots');
assert(trendEmpty.every(t => t.revenus === 0 && t.depenses === 0 && t.net === 0), 'All 6 trend months have 0 revenus/depenses/net');

const nwEmpty = engine.computeNetWorth('2026-02');
assert.strictEqual(nwEmpty.totalAssets, 0, 'Empty netWorth totalAssets === 0');
assert.strictEqual(nwEmpty.totalLiabilities, 0, 'Empty netWorth totalLiabilities === 0');
assert.strictEqual(nwEmpty.netWorth, 0, 'Empty netWorth === 0');
assert.strictEqual(nwEmpty.accounts.length, 8, 'NetWorth inventory lists the 8 active configured accounts at 0.00');

const bgEmpty = engine.computeBudgetZBB('2026-02');
assert.strictEqual(bgEmpty.totalAlloue, 0, 'Empty ZBB totalAlloue === 0');
assert.strictEqual(bgEmpty.totalConsomme, 0, 'Empty ZBB totalConsomme === 0');
assert.strictEqual(bgEmpty.totalEcart, 0, 'Empty ZBB totalEcart === 0');
assert.strictEqual(bgEmpty.consumptionRate, 0, 'Empty ZBB consumptionRate === 0');
assert.strictEqual(bgEmpty.categories.length, 0, 'Empty ZBB categories must be [] (no dummy 0/0 rows)');

const hubEmpty = engine.getConsolidatedHubData('2026-02');
assert.strictEqual(hubEmpty.hasData, false, 'hubEmpty.hasData === false');
assert.strictEqual(hubEmpty.hasCashflowData, false, 'hubEmpty.hasCashflowData === false');
assert.strictEqual(hubEmpty.hasPatrimoineData, false, 'hubEmpty.hasPatrimoineData === false');
assert.strictEqual(hubEmpty.hasBudgetData, false, 'hubEmpty.hasBudgetData === false');
assert.strictEqual(hubEmpty.topAssets.length, 0, 'hubEmpty.topAssets must be [] when all accounts have 0.00 balance');
assert.strictEqual(hubEmpty.topBudgetCategories.length, 0, 'hubEmpty.topBudgetCategories must be [] when no budgets exist');
assert.strictEqual(hubEmpty.recentTransactions.length, 0, 'hubEmpty.recentTransactions must be [] when no transactions exist');
assert(hubEmpty.healthScore >= 60, `Neutral healthScore for empty debt-free portfolio should not be penalized (got ${hubEmpty.healthScore})`);

const repEmpty = engine.generateFinancialReport('2026-02');
assert.strictEqual(repEmpty.hasData, false, 'repEmpty.hasData === false');
assert.strictEqual(repEmpty.topExpenses.length, 0, 'repEmpty.topExpenses === []');
console.log('✔ PASS: All analytical modules return clean, resilient zero-states without dummy rows or false penalties.');

// 4. Verify Dynamic Switching Between DEMO and EMPTY Modes + Purge Persistence
console.log('\n--- 4. Live Switching (setDataMode DEMO <-> EMPTY) & purgeDatabase() ---');
const switchDemo = engine.setDataMode('DEMO');
assert.strictEqual(switchDemo.dataMode, 'DEMO', 'setDataMode("DEMO") activates DEMO mode');
assert.strictEqual(switchDemo.txCount, 60, 'setDataMode("DEMO") loads 60 sample transactions');
assert.strictEqual(switchDemo.budgetCount, 17, 'setDataMode("DEMO") loads 17 sample budgets');
assert.strictEqual(engine.hasFinancialData('2026-02'), true, 'hasFinancialData("2026-02") is true in DEMO mode');
assert.strictEqual(engine.computeBudgetZBB('2026-02').categories.length, 6, 'computeBudgetZBB("2026-02") returns all 6 active categories in DEMO mode');

engine.purgeDatabase();
assert.strictEqual(engine.getDataMode(), 'EMPTY', 'purgeDatabase() sets dataMode to EMPTY');
assert.strictEqual(engine.data.transactions.length, 0, 'purgeDatabase() clears transactions');
assert.strictEqual(engine.data.budgets.length, 0, 'purgeDatabase() clears budgets');

// Restore default DEMO mode for any subsequent state
engine.setDataMode('DEMO');
assert.strictEqual(engine.getDataMode(), 'DEMO', 'Restored DEMO mode cleanly');

// 5. Verify app.js Rich Empty State & Delegated Action Wiring
console.log('\n--- 5. App Controller Rich Empty States & Action Delegation ---');
assert(appJs.includes('empty-state-donut-wrap') && appJs.includes('trend-empty-callout'), 'app.js renders .empty-state-donut-wrap and .trend-empty-callout');
assert(appJs.includes('data-empty-action="add-tx"') && appJs.includes('data-empty-action="open-alloc"') && appJs.includes('data-empty-action="reset-cf-filters"'), 'app.js renders contextual [data-empty-action] CTA buttons');
assert(appJs.includes("e.target.closest('[data-empty-action]')"), 'app.js binds delegated click handler for all [data-empty-action] buttons');
console.log('✔ PASS: App controller rich empty states and delegated action handlers verified.');

console.log('\n================================================================================');
console.log('🏆 ALL DATA MODE CHOICE & EMPTY-STATE RESILIENCE CHECKS PASSED! (0 ERRORS)');
console.log('================================================================================');
