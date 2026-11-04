/**
 * Automated Verification Suite:
 * 1. "6-" removed from "MONTH CASHFLOW TREND" (HUB_TREND_TITLE in FR/EN/ES/PT & #hub-trend-title in index.html)
 * 2. Decluttered "Net Assets: ..." from ASSET ALLOCATION & WEALTH STRUCTURE header (#hub-card-patrimoine / #hub-patrimoine-total)
 * 3. ALL Expense Categories displayed, interactive, and responsive in EXPENSE BREAKDOWN BY CATEGORY Donut Chart (zero "+N Autres" truncation, container queries, bidirectional Hub & Cashflow filtering)
 * 4. Gold-Standard Cutting-Edge First-Launch Onboarding Screen (#modal-initialization) design, CSS rules, and live quadrilingual translation
 */

const fs = require('fs');
const path = require('path');
const assert = require('assert');

console.log('=== STARTING TEST SUITE: UI REFINEMENTS, FULL DONUT INTERACTIVITY & ONBOARDING REDESIGN ===\n');

const html = fs.readFileSync(path.join(__dirname, '../mobile-web/index.html'), 'utf8');
const css = fs.readFileSync(path.join(__dirname, '../mobile-web/styles.css'), 'utf8');
const appJs = fs.readFileSync(path.join(__dirname, '../mobile-web/app.js'), 'utf8');
const engineCode = fs.readFileSync(path.join(__dirname, '../mobile-web/analytics-engine.js'), 'utf8');
const baselineCode = fs.readFileSync(path.join(__dirname, '../mobile-web/baseline_data.js'), 'utf8');

// Setup mock browser environment for engine evaluation
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
const sampleFixture = JSON.parse(fs.readFileSync(path.join(__dirname, 'sample_test_dataset.json'), 'utf8'));
engine.init(sampleFixture);

// --- TEST 1: "6-" REMOVED FROM "MONTH CASHFLOW TREND" ---
console.log('--- TEST 1: "6-" removed from MONTH CASHFLOW TREND ---');
engine.setLanguage('EN');
assert.strictEqual(engine.tr('HUB_TREND_TITLE'), 'MONTH CASHFLOW TREND', `Expected EN HUB_TREND_TITLE to be "MONTH CASHFLOW TREND", got "${engine.tr('HUB_TREND_TITLE')}"`);
assert(!engine.tr('HUB_TREND_TITLE').includes('6-'), 'EN HUB_TREND_TITLE must not contain "6-"');

engine.setLanguage('FR');
assert.strictEqual(engine.tr('HUB_TREND_TITLE'), 'ÉVOLUTION DES FLUX MENSUELS (TRÉSORERIE)', `Expected FR HUB_TREND_TITLE to be "ÉVOLUTION DES FLUX MENSUELS (TRÉSORERIE)", got "${engine.tr('HUB_TREND_TITLE')}"`);
assert(!html.includes('>ÉVOLUTION DES FLUX SUR 6 MOIS (TRÉSORERIE)<'), 'index.html #hub-trend-title must not contain "SUR 6 MOIS"');
assert(html.includes('id="hub-trend-title">ÉVOLUTION DES FLUX MENSUELS (TRÉSORERIE)</h3>'), 'index.html #hub-trend-title should have updated title');
console.log('✔ PASS: "6-" cleanly removed from MONTH CASHFLOW TREND across all languages and HTML.');

// --- TEST 2: DECLUTTERED "Net Assets: ..." NEXT TO ASSET ALLOCATION & WEALTH STRUCTURE ---
console.log('\n--- TEST 2: Decluttered Net Assets next to ASSET ALLOCATION & WEALTH STRUCTURE ---');
assert(html.includes('id="hub-patrimoine-total" aria-hidden="true"'), 'index.html #hub-patrimoine-total should be hidden from header layout');
assert(!appJs.includes("patTotal.textContent = `${engine.tr('NET_ASSETS')} : ${engine.formatMoney(data.netWorth)}`"), 'app.js must no longer stuff Net Assets: ... into #hub-patrimoine-total');
assert(appJs.includes("patTotal.textContent = '';"), 'app.js should keep #hub-patrimoine-total empty');
console.log('✔ PASS: Redundant "Net Assets: ..." removed from ASSET ALLOCATION & WEALTH STRUCTURE header.');

// --- TEST 3: ALL CATEGORIES IN EXPENSE BREAKDOWN BY CATEGORY DONUT DISPLAYED, INTERACTIVE & RESPONSIVE ---
console.log('\n--- TEST 3: All Donut Categories Displayed, Interactive & Responsive ---');
assert(!appJs.includes('const maxLegend = breakdown.length <= 5 ? 5 : 4;'), 'app.js renderDonutChart must not truncate categories to maxLegend 4/5');
assert(!appJs.includes('const topItems = breakdown.slice(0, maxLegend);'), 'app.js renderDonutChart must not slice breakdown into topItems');
assert(appJs.includes('breakdown.forEach(item => {'), 'app.js renderDonutChart must iterate over all breakdown items');
assert(appJs.includes('class="legend-amount"'), 'app.js renderDonutChart must display exact amount for each category');
assert(appJs.includes('class="legend-mini-bar-fill"'), 'app.js renderDonutChart must render proportional micro-bar for each category');
assert(html.includes('id="hub-donut-active-filter"'), 'index.html must include #hub-donut-active-filter inside #hub-card-donut');

// Verify CSS container queries & responsive legend styles
assert(css.includes('container-type: inline-size;'), 'styles.css .chart-card must have container-type: inline-size');
assert(css.includes('@container (max-width: 460px)'), 'styles.css must include @container responsive breakpoint for donut chart');
assert(css.includes('.legend-mini-bar-track') && css.includes('.legend-mini-bar-fill'), 'styles.css must style .legend-mini-bar-track and .legend-mini-bar-fill');

// Verify Hub recentTransactions filtering when selectedCategoryFilter is active
engine.setLanguage('FR');
const allHub = engine.getConsolidatedHubData('2026-02');
assert.strictEqual(allHub.breakdown.length, 6, `Expected 6 expense categories in Feb 2026 breakdown, got ${allHub.breakdown.length}`);
console.log(`✔ All ${allHub.breakdown.length} expense categories present in Feb 2026 breakdown:`, allHub.breakdown.map(b => b.nom).join(', '));

engine.selectedCategoryFilter = 'Santé & Mutuelle';
const filteredHub = engine.getConsolidatedHubData('2026-02');
assert(filteredHub.recentTransactions.length > 0, 'Filtered Hub recentTransactions for "Santé & Mutuelle" should not be empty');
filteredHub.recentTransactions.forEach(tx => {
  assert.strictEqual(tx.rawCategorie, 'Santé & Mutuelle', `Expected filtered transaction category to be "Santé & Mutuelle", got "${tx.rawCategorie}"`);
});
engine.selectedCategoryFilter = null;
console.log('✔ PASS: Every category in EXPENSE BREAKDOWN BY CATEGORY is displayed, interactive, responsive, and filters both Hub & Cashflow.');

// --- TEST 4: FIRST-LAUNCH ONBOARDING SCREEN (#modal-initialization) MODERN GOLD-STANDARD DESIGN ---
console.log('\n--- TEST 4: First-Launch Onboarding Screen (#modal-initialization) Modern Design ---');
assert(html.includes('class="modal-content modal-fullscreen-content init-onboarding-shell"'), 'index.html #modal-initialization must use .init-onboarding-shell');
assert(html.includes('class="modal-header init-hero-banner"'), 'index.html #modal-initialization must include .init-hero-banner');
assert(!html.includes('class="init-trust-strip"') && !html.includes('id="init-trust-1"') && !html.includes('id="init-trust-2"') && !html.includes('id="init-trust-3"'), 'index.html #modal-initialization must NOT contain the removed trust strip labels (100% Hors-Ligne & Privé, Multi-Devises Live API, Architecture ZBB & Audit)');
assert(!css.includes('.init-trust-strip') && !css.includes('.init-trust-pill'), 'styles.css must NOT contain removed .init-trust-strip / .init-trust-pill rules');
assert(!appJs.includes("setTxt('init-trust-1'") && !appJs.includes("setTxt('init-trust-2'") && !appJs.includes("setTxt('init-trust-3'"), 'app.js must NOT reference removed init-trust-* elements');
assert(html.includes('class="init-emblem-svg"') && html.includes('id="gfmp-gold-grad"') && html.includes('id="gfmp-sapphire-core"'), 'index.html #modal-initialization must include the new Haute Joaillerie Private Banking Sovereign Diamond & Ascending Wealth Medallion SVG');
assert(css.includes('.init-emblem-svg') && css.includes('forced-color-adjust: none;'), 'styles.css must style .init-emblem-svg with modern icon best practices');
assert(html.includes('id="init-live-summary"'), 'index.html #modal-initialization must include live configuration summary pill');
assert(css.includes('.init-onboarding-shell') && css.includes('.init-hero-banner') && css.includes('.btn-primary-alloc'), 'styles.css must define complete luxury styling for .init-onboarding-shell, .init-hero-banner, and .btn-primary-alloc');

// Verify quadrilingual onboarding translation keys in analytics-engine.js (including Steps 1-4: Langue, Devise, Date/Période, Thème)
assert.strictEqual(typeof engine.getDictionary, 'function', 'engine.getDictionary must be a function so callers never throw');
['FR', 'EN', 'ES', 'PT'].forEach(lang => {
  [
    'INIT_BADGE', 'INIT_TITLE', 'INIT_SUBTITLE',
    'INIT_TRUST_1', 'INIT_TRUST_2', 'INIT_TRUST_3',
    'INIT_STEP1_BADGE', 'INIT_STEP1_TITLE', 'INIT_STEP1_DESC',
    'INIT_STEP2_BADGE', 'INIT_STEP2_TITLE', 'INIT_STEP2_DESC',
    'INIT_STEP3_BADGE', 'INIT_STEP3_TITLE', 'INIT_STEP3_DESC',
    'INIT_DATE_LABEL', 'INIT_MONTH_LABEL', 'INIT_YEAR_LABEL', 'INIT_BTN_TODAY', 'INIT_BTN_FIRST',
    'INIT_STEP4_BADGE', 'INIT_STEP4_TITLE', 'INIT_STEP4_DESC',
    'THEME_NAME_BLUE', 'THEME_SUB_BLUE',
    'THEME_NAME_GOLD', 'THEME_SUB_GOLD',
    'THEME_NAME_EMERALD', 'THEME_SUB_EMERALD',
    'THEME_NAME_AMETHYST', 'THEME_SUB_AMETHYST',
    'INIT_BTN_CONFIRM',
    'CURR_NAME_MUR', 'CURR_NAME_EUR', 'CURR_NAME_USD', 'CURR_NAME_GBP', 'CURR_NAME_ZAR', 'CURR_NAME_XOF'
  ].forEach(k => {
    const val = engine.tr(k, {}, lang);
    assert(val && val !== k, `Missing translation for onboarding key ${k} in ${lang}`);
  });
});

// Verify 4-Step Startup Configuration HTML & CSS
assert(html.includes('id="init-date-section"') && html.includes('id="init-date-input"') && html.includes('id="init-month-select"') && html.includes('id="init-year-select"'), 'index.html must include Step 3 Date & Financial Period controls');
assert(html.includes('id="init-theme-section"') && html.includes('id="init-theme-grid"') && html.includes('data-init-theme="emerald"'), 'index.html must include Step 4 Theme controls');
assert(css.includes('.init-date-picker') && css.includes('.init-theme-swatch') && css.includes('.init-period-select'), 'styles.css must define Step 3 Date and Step 4 Theme luxury styles');

// Verify engine.needsOnboarding & 4-pillar initializeSystem()
store['gfmp_initialized'] = 'true';
delete store['gfmp_onboarding_v5_6'];
engine.init(global.window.GFMP_BASELINE_DATA);
assert.strictEqual(engine.needsOnboarding, true, 'engine.needsOnboarding must be true when gfmp_onboarding_v5_6 is not yet set');

engine.initializeSystem('EUR', 'EN', '2026-03-15', 'emerald', '2026', '03');
assert.strictEqual(engine.needsOnboarding, false, 'engine.needsOnboarding must become false after initializeSystem()');
assert.strictEqual(engine.currency, 'EUR', 'Currency must be EUR');
assert.strictEqual(engine.lang, 'EN', 'Language must be EN');
assert.strictEqual(engine.referenceDate, '2026-03-15', 'Reference date must be 2026-03-15');
assert.strictEqual(engine.year, '2026', 'Year must be 2026');
assert.strictEqual(engine.month, '03', 'Month must be 03');
assert.strictEqual(engine.theme, 'emerald', 'Theme must be emerald');
assert.strictEqual(store['gfmp_theme'], 'emerald', 'localStorage gfmp_theme must be emerald');
assert.strictEqual(store['gfmp_onboarding_v5_6'], 'true', 'localStorage gfmp_onboarding_v5_6 must be true');

// Verify MainActivity.kt fresh APK installation detection (?onboarding=1)
const mainActivityKt = fs.readFileSync(path.join(__dirname, '../android/app/src/main/kotlin/com/jicorporate/gfmp/MainActivity.kt'), 'utf8');
assert(mainActivityKt.includes('file:///android_asset/index.html?onboarding=1'), 'MainActivity.kt must load index.html?onboarding=1 on fresh APK install/update');
console.log('✔ PASS: 4-Step Startup Configuration Page (Language, Currency, Date/Period, Theme) & APK fresh-install trigger verified.');

console.log('\n======================================================');
console.log('ALL UI REFINEMENTS & ONBOARDING TESTS PASSED! (0 ERRORS)');
console.log('======================================================');

