const fs = require('fs');
const path = require('path');
const assert = require('assert');

console.log('================================================================================');
console.log('🧪 VERIFICATION SUITE: FOREX CONVERSION, ZERO-FOUC & 3-KPI CARDS (GOLD STANDARD)');
console.log('================================================================================\n');

const html = fs.readFileSync(path.join(__dirname, '../mobile-web/index.html'), 'utf8');
const css = fs.readFileSync(path.join(__dirname, '../mobile-web/styles.css'), 'utf8');
const appJs = fs.readFileSync(path.join(__dirname, '../mobile-web/app.js'), 'utf8');
const engineCode = fs.readFileSync(path.join(__dirname, '../mobile-web/analytics-engine.js'), 'utf8');
const baselineCode = fs.readFileSync(path.join(__dirname, '../mobile-web/baseline_data.js'), 'utf8');

// --- TEST 1: ZERO-FOUC PROTECTION ---
console.log('--- 1. Verification of Zero-FOUC Startup Onboarding Protection ---');
assert(html.includes("document.documentElement.classList.add('gfmp-onboarding-active')"), 'index.html <head> must contain synchronous zero-FOUC script');
assert(css.includes('html.gfmp-onboarding-active body'), 'styles.css must style html.gfmp-onboarding-active body');
assert(css.includes('html.gfmp-onboarding-active #modal-init-system'), 'styles.css must force #modal-init-system visible during onboarding');
assert(appJs.includes("document.documentElement.classList.remove('gfmp-onboarding-active')"), 'app.js must remove gfmp-onboarding-active when onboarding is complete');
console.log('✔ PASS: Zero-FOUC Startup Onboarding Protection verified.');

// --- TEST 2: ALLOCATE BUDGET CURRENCY SELECT ---
console.log('\n--- 2. Verification of Allocate Budget Currency Select ---');
assert(html.includes('id="alloc-devise"'), 'index.html must include select #alloc-devise for dynamic budget currency selection');
assert(appJs.includes("const dev = document.getElementById('alloc-devise')?.value || engine.currency;"), 'app.js must read chosen currency from #alloc-devise on submit');
assert(appJs.includes("engine.allocateBudget(cat, amt, mois, dev);"), 'app.js must pass currency to allocateBudget');
console.log('✔ PASS: Allocate budget currency selector verified.');

// --- TEST 3: 3 KPI CARDS ON CASHFLOW AND NET WORTH ---
console.log('\n--- 3. Verification of 3 KPI Cards on Cashflow and Net Worth ---');
assert(html.includes('class="kpi-row kpi-cashflow-row"'), 'index.html must contain .kpi-row.kpi-cashflow-row on view-cashflow');
assert(html.includes('id="card-cf-rev"') && html.includes('id="card-cf-dep"') && html.includes('id="card-kpi-net"'), 'Cashflow must contain 3 dedicated KPI cards');
assert(html.includes('id="lbl-cf-sub-rev"') && html.includes('id="lbl-cf-sub-dep"') && html.includes('id="lbl-cf-sub-net"'), 'Cashflow KPI cards must have subtitle slots');

assert(html.includes('class="kpi-row kpi-networth-row" id="networth-hero-card"'), 'index.html must contain .kpi-row.kpi-networth-row on view-networth');
assert(html.includes('id="card-nw-networth"') && html.includes('id="card-nw-assets"') && html.includes('id="card-nw-debts"'), 'Net Worth must contain 3 dedicated KPI cards');
assert(html.includes('id="nw-allocation-card"'), 'Net Worth must have dedicated Asset Allocation & Wealth Structure card');
assert(css.includes('.kpi-card') && css.includes('.kpi-card::before'), 'styles.css must style .kpi-card with accent border');
console.log('✔ PASS: 3 Dedicated KPI Cards on Cashflow and Net Worth verified.');

// --- TEST 4: STRICT BAN ON UNWANTED LABELS AND MANUAL OVERRIDE WIDGETS ---
console.log('\n--- 4. Verification of Ban on Unwanted Labels, Manual Overrides & Badges ---');
assert(!html.includes('id="forex-manual-table"'), 'Must NOT have manual forex table');
assert(!html.includes('id="forex-calculator"'), 'Must NOT have forex calculator');
assert(!html.includes('SFP v1.01 Production Sovereign Personal Portfolio ready for your entries'), 'Must NOT have unwanted production label');
assert(!html.includes('LOAD TEST DATA'), 'Must NOT have LOAD TEST DATA label');
console.log('✔ PASS: Zero unwanted labels, zero decorative badges.');

// --- TEST 5: MATHEMATICAL AND FOREX AUTOMATIC ENGINE VERIFICATION ---
console.log('\n--- 5. Forex Automatic Conversion & Mathematical Precision Verification ---');
const store = {};
global.window = {
  localStorage: {
    getItem: (k) => store[k] || null,
    setItem: (k, v) => { store[k] = String(v); },
    removeItem: (k) => { delete store[k]; }
  }
};
global.localStorage = global.window.localStorage;
eval(baselineCode);
eval(engineCode);

const engine = window.gfmpEngine;
engine.init();

// Test baseline conversion rates
const murToXof = engine.convert(90000, 'MUR', 'XOF');
console.log(`90 000 MUR in XOF: ${murToXof} FCFA`);
assert(murToXof > 1000000, `90 000 MUR in XOF must represent true purchasing power (~1 100 187 FCFA), got ${murToXof}`);
assert.strictEqual(murToXof, 1100187, `Expected exactly 1 100 187 FCFA, got ${murToXof}`);

const xofToXof = engine.convert(90000, 'XOF', 'XOF');
console.log(`90 000 XOF in XOF: ${xofToXof} FCFA`);
assert.strictEqual(xofToXof, 90000, `90 000 XOF in XOF must be exactly 90 000, got ${xofToXof}`);

// Initialize user system in XOF
engine.initializeSystem('XOF', 'FR', '2026-02-25', 'blue', '2026', '02', false);
assert.strictEqual(engine.currency, 'XOF', 'System currency must be XOF');
assert.strictEqual(engine.data.comptes[0].devise, 'XOF', 'Checking account must be aligned to user base currency XOF');
assert.strictEqual(engine.data.comptes[1].devise, 'XOF', 'Savings account must be aligned to user base currency XOF');

// Add transaction of 90 000 MUR
engine.addTransaction({
  date: '2026-02-15',
  id_categorie: '1',
  id_tiers: '1',
  id_compte: '1',
  montant: 90000,
  notes: 'Salaire Maurice',
  devise: 'MUR'
});

const cfXof = engine.computeCashflow('2026-02');
console.log(`Total revenus in system currency (XOF): ${cfXof.totalRevenus} FCFA`);
assert.strictEqual(cfXof.totalRevenus, 1100187, `Total revenus in XOF must be 1 100 187, got ${cfXof.totalRevenus}`);

const nwXof = engine.computeNetWorth('2026-02');
console.log(`Net worth in system currency (XOF): ${nwXof.netWorth} FCFA`);
assert.strictEqual(nwXof.netWorth, 1100187, `Net worth in XOF must be 1 100 187, got ${nwXof.netWorth}`);

// Switch view dynamically to MUR
engine.setCurrency('MUR');
const cfMur = engine.computeCashflow('2026-02');
console.log(`Total revenus dynamically viewed in MUR: ${cfMur.totalRevenus} Rs`);
assert.strictEqual(cfMur.totalRevenus, 90000, `Total revenus in MUR must be exactly 90 000, got ${cfMur.totalRevenus}`);

const nwMur = engine.computeNetWorth('2026-02');
console.log(`Net worth dynamically viewed in MUR: ${nwMur.netWorth} Rs`);
assert.strictEqual(nwMur.netWorth, 90000, `Net worth in MUR must be exactly 90 000, got ${nwMur.netWorth}`);

// Test budget allocation in custom currency
engine.allocateBudget('4', 1500, '2026-02', 'EUR');
const savedBudget = engine.data.budgets.find(b => b.mois === '2026-02' && b.id_categorie === '4');
assert(savedBudget, 'Budget must be saved');
assert.strictEqual(savedBudget.devise, 'EUR', `Budget currency must be EUR, got ${savedBudget.devise}`);
assert.strictEqual(savedBudget.montant, 1500, 'Budget amount must be 1500');

// --- TEST 6: DYNAMIC FORM CURRENCY ADAPTATION (NO HARDCODED MUR) ---
console.log('\n--- 6. Dynamic Form Currencies Adaptation (No Hardcoded MUR) ---');
assert(appJs.includes('let txDeviseUserModified = false;'), 'app.js must track txDeviseUserModified');
assert(appJs.includes('let expTxDeviseUserModified = false;'), 'app.js must track expTxDeviseUserModified');
assert(appJs.includes('let allocDeviseUserModified = false;'), 'app.js must track allocDeviseUserModified');

// Verify that fillSelects dynamically selects the account native currency or user base currency
assert(appJs.includes("const defaultDev = (selectedCpt && selectedCpt.devise) ? selectedCpt.devise : (engine.currency || 'MUR');"), 'fillSelects must select account currency or engine currency');
assert(appJs.includes('devSel.value = activeDev;'), 'fillSelects must set devSel.value to activeDev');

// Verify that onboarding confirmation sets all form currencies to selectedInitCurr
assert(appJs.includes('if (txDev) txDev.value = selectedInitCurr;'), 'btn-confirm-init must set txDev.value to selectedInitCurr');
assert(appJs.includes('if (expDev) expDev.value = selectedInitCurr;'), 'btn-confirm-init must set expDev.value to selectedInitCurr');
assert(appJs.includes('if (allocDev) allocDev.value = selectedInitCurr;'), 'btn-confirm-init must set allocDev.value to selectedInitCurr');
assert(appJs.includes('if (addEntityDev) addEntityDev.value = selectedInitCurr;'), 'btn-confirm-init must set addEntityDev.value to selectedInitCurr');

// Verify that renderCurrentView dynamically synchronizes form currencies when not user-modified
assert(appJs.includes('txDevise.value = (activeCpt && activeCpt.devise) ? activeCpt.devise : engine.currency;'), 'renderCurrentView must sync txDevise when not user modified');
assert(appJs.includes('expTxDevise.value = (activeExpCpt && activeExpCpt.devise) ? activeExpCpt.devise : engine.currency;'), 'renderCurrentView must sync expTxDevise when not user modified');

console.log('✔ PASS: Dynamic Form Currencies Adaptation verified (Forms never stuck on MUR).');

console.log('\n================================================================================');
console.log('🏆 ALL FOREX CONVERSION, ZERO-FOUC & DYNAMIC FORM CHECKS PASSED WITH 0 ERRORS!');
console.log('================================================================================');
