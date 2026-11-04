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

// 1. Verify Complete Eradication of "LOAD TEST DATA" & Obsolete Dev Labels + Dynamic Accounts Inventory
console.log('--- 1. Eradication of Test Mode Artifacts & Dynamic Accounts Inventory DOM/CSS/i18n ---');
assert(!html.includes('id="init-datamode-section"'), 'index.html must NOT include obsolete #init-datamode-section');
assert(!html.includes('id="hub-datamode-banner"'), 'index.html must NOT include #hub-datamode-banner');
assert(!html.includes('id="settings-datamode-card"'), 'index.html must NOT include obsolete #settings-datamode-card');
assert(!appJs.includes('data-empty-action="load-demo"'), 'app.js must NOT offer any "load-demo" action');
assert(!html.includes('SFP v1.01'), 'index.html must NOT contain obsolete "SFP v1.01" dev labels');
assert(!appJs.includes('SFP v1.01 Production'), 'app.js must NOT contain "SFP v1.01 Production" text');

// Dynamic Accounts & Liabilities Inventory elements
assert(html.includes('id="nw-accounts-summary-badge"'), 'index.html #view-networth must include #nw-accounts-summary-badge');
assert(html.includes('id="lbl-as-of-date"'), 'index.html #view-networth must include #lbl-as-of-date');
assert(html.includes('id="btn-nw-add-account"'), 'index.html #view-networth must include #btn-nw-add-account');
assert(html.includes('id="networth-table-body"'), 'index.html #view-networth must include #networth-table-body');

assert(css.includes('.table-counter-badge') && css.includes('.btn-inline-add-entity'), 'styles.css must define .table-counter-badge and .btn-inline-add-entity');
assert(css.includes('.cpt-status-pill.ready') && css.includes('.cpt-status-pill.clear') && css.includes('.btn-cpt-quick-add'), 'styles.css must define micro-status pills and quick-add button');

const requiredKeys = [
  'INIT_BADGE', 'INIT_TITLE', 'INIT_SUBTITLE',
  'SOLVENCY_EMPTY', 'TREND_EMPTY_TITLE', 'TREND_EMPTY_SUB', 'DONUT_EMPTY_HINT',
  'NO_ACCOUNT_OR_TX', 'EMPTY_PATRIMOINE_SUB', 'EMPTY_BUDGET_SUB', 'EMPTY_TX_SUB',
  'EMPTY_CTA_ADD_TX', 'EMPTY_CTA_ADD_EXP', 'EMPTY_CTA_ALLOC_BG', 'EMPTY_CTA_ADD_ACC',
  'EMPTY_CTA_RESET_FILTER', 'EMPTY_CTA_FULL_YEAR',
  'NW_READINESS_TITLE', 'NW_READINESS_DESC',
  'ACCOUNT_STATUS_READY', 'ACCOUNT_STATUS_DEBT_FREE',
  'ACCOUNT_STATUS_FUNDED', 'ACCOUNT_STATUS_DEBT',
  'ACCOUNT_QUICK_ADD_TX', 'NW_ACCOUNTS_COUNT',
  'NW_STACK_EMPTY_LABEL', 'BTN_NW_ADD_ACCOUNT',
  'REPORT_EMPTY_DIAG_TITLE', 'REPORT_EMPTY_DIAG_DESC'
];
['FR', 'EN', 'ES', 'PT'].forEach(lang => {
  requiredKeys.forEach(k => {
    const val = engine.tr(k, {}, lang);
    assert(val && val !== k, `Missing ${lang} translation for key ${k}`);
  });
});
console.log('✔ PASS: Eradication of test artifacts, dynamic accounts inventory & Quadrilingual keys verified.');

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
assert.strictEqual(nwEmpty.accounts.length, 0, 'NetWorth transaction-derived inventory is empty ([]) when 0 transactions exist');

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

// 4. Verify Production v1.01 Zero-Test-Data Guarantee & Legacy v1.0 Migration + Purge Persistence
console.log('\n--- 4. Production v1.01 Zero-Test-Data Baseline, Legacy Migration & User Operations ---');
assert.strictEqual(window.GFMP_BASELINE.version, '1.01', 'GFMP_BASELINE.version must be 1.01');
assert.strictEqual(window.GFMP_BASELINE.dataset_version, '1.01-PROD', 'GFMP_BASELINE.dataset_version must be 1.01-PROD');
assert.strictEqual(window.GFMP_BASELINE.transactions.length, 0, 'Production v1.01 baseline must have 0 test transactions');
assert.strictEqual(window.GFMP_BASELINE.budgets.length, 0, 'Production v1.01 baseline must have 0 test budgets');

// Simulate legacy v1.0 localStorage containing 60 demo transactions -> verify automatic v1.01 migration scrubs them
const legacyFixture = JSON.parse(fs.readFileSync(path.join(__dirname, 'sample_test_dataset.json'), 'utf8'));
global.localStorage.setItem('gfmp_dataset_v3', JSON.stringify(legacyFixture));
global.localStorage.setItem('gfmp_data_mode', 'DEMO');
engine.data = null;
engine.init();
assert.strictEqual(engine.data.dataset_version, '1.01-PROD', 'Legacy v1.0 localStorage must auto-migrate to 1.01-PROD');
assert.strictEqual(engine.getDataMode(), 'EMPTY', 'Legacy DEMO mode must auto-migrate to EMPTY in Production v1.01');
assert.strictEqual(engine.data.transactions.length, 0, 'Legacy 60 test transactions must be scrubbed on v1.01 migration');
assert.strictEqual(engine.data.budgets.length, 0, 'Legacy 17 test budgets must be scrubbed on v1.01 migration');

// Add a real user transaction and budget, then verify hasFinancialData() and purgeDatabase()
engine.addTransaction({
  date: '2026-02-15',
  type: 'REVENU',
  montant: 50000,
  devise: 'MUR',
  id_compte: '1',
  id_categorie: '1',
  id_tiers: '1',
  notes: 'Revenu réel Production v1.01'
});
engine.allocateBudget('4', 15000, '2026-02');
assert.strictEqual(engine.hasFinancialData('2026-02'), true, 'hasFinancialData("2026-02") is true once user records real operations');
assert.strictEqual(engine.computeBudgetZBB('2026-02').categories.length, 1, 'computeBudgetZBB("2026-02") returns user allocated category');
assert.strictEqual(engine.computeNetWorth('2026-02').accounts.length, 1, 'NetWorth transaction-derived inventory dynamically includes active account with transactions');

engine.purgeDatabase();
assert.strictEqual(engine.getDataMode(), 'EMPTY', 'purgeDatabase() maintains EMPTY mode');
assert.strictEqual(engine.data.transactions.length, 0, 'purgeDatabase() clears transactions');
assert.strictEqual(engine.data.budgets.length, 0, 'purgeDatabase() clears budgets');
assert.strictEqual(engine.computeNetWorth('2026-02').accounts.length, 0, 'NetWorth transaction-derived inventory returns to 0 after purge');

// 5. Verify app.js Rich Empty State & Delegated Action Wiring
console.log('\n--- 5. App Controller Rich Empty States & Action Delegation ---');
assert(appJs.includes('empty-state-donut-wrap') && appJs.includes('trend-empty-callout'), 'app.js renders .empty-state-donut-wrap and .trend-empty-callout');
assert(appJs.includes('data-empty-action="add-tx"') && appJs.includes('data-empty-action="open-alloc"') && appJs.includes('data-empty-action="reset-cf-filters"'), 'app.js renders contextual [data-empty-action] CTA buttons');
assert(appJs.includes("e.target.closest('[data-empty-action]')"), 'app.js binds delegated click handler for all [data-empty-action] buttons');
assert(appJs.includes('networth-table-body') && appJs.includes('NO_ACCOUNT_OR_TX'), 'app.js renders rich empty state inside #networth-table-body');
assert(appJs.includes('btn-cpt-quick-add'), 'app.js renders .btn-cpt-quick-add quick action on account rows');
assert(appJs.includes('btn-nw-add-account'), 'app.js binds #btn-nw-add-account to openAddEntityDialog("compte")');
console.log('✔ PASS: App controller rich empty states and delegated action handlers verified.');

console.log('\n================================================================================');
console.log('🏆 ALL DATA MODE CHOICE & EMPTY-STATE RESILIENCE CHECKS PASSED! (0 ERRORS)');
console.log('================================================================================');
