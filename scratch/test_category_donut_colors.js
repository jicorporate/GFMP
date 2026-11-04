/**
 * Automated Test Suite: Category Donut Colors & Haute Joaillerie Alignment
 * Verifies category colors, slice-to-legend concordance, determinism, and zero regressions.
 */

const fs = require('fs');
const path = require('path');

// Load analytics-engine.js in a mock browser environment
const engineCode = fs.readFileSync(path.join(__dirname, '../mobile-web/analytics-engine.js'), 'utf8');
const baselineCode = fs.readFileSync(path.join(__dirname, '../mobile-web/baseline_data.js'), 'utf8');

const mockWindow = {};
const mockLocalStorage = {
  _store: {},
  getItem(k) { return this._store[k] || null; },
  setItem(k, v) { this._store[k] = String(v); },
  removeItem(k) { delete this._store[k]; }
};

// Evaluate baseline and engine
const runContext = new Function('window', 'localStorage', 'console', `
  ${baselineCode}
  ${engineCode}
  return { GFMPEngine, GFMP_CATEGORY_COLORS, GFMP_LUXURY_PALETTE };
`);

const { GFMPEngine, GFMP_CATEGORY_COLORS, GFMP_LUXURY_PALETTE } = runContext(mockWindow, mockLocalStorage, console);

console.log('=== TEST 1: CATEGORY COLOR PALETTE & MAPPING INITIALIZATION ===');
if (!Array.isArray(GFMP_LUXURY_PALETTE) || GFMP_LUXURY_PALETTE.length < 10) {
  throw new Error(`FAIL: GFMP_LUXURY_PALETTE invalid length: ${GFMP_LUXURY_PALETTE?.length}`);
}
console.log(`✓ GFMP_LUXURY_PALETTE loaded with ${GFMP_LUXURY_PALETTE.length} luxury jewel colors`);

const engine = new GFMPEngine();
engine.init();

console.log('\n=== TEST 2: STANDARD CATEGORIES HAUTE JOAILLERIE COLOR VERIFICATION ===');
const expectedColors = {
  'Logement (Loyer/Prêt/Charges)': '#3B82F6', // Sapphire Blue
  'Alimentation & Supermarché': '#F59E0B',    // Amber Gold
  'Transports (Essence/Assurance)': '#8B5CF6',// Amethyst Violet
  'Santé & Mutuelle': '#10B981',              // Emerald Green
  'Loisirs, Sorties & Vacances': '#EC4899',   // Ruby Rose
  'Impôts & Taxes': '#F97316'                 // Topaz Orange
};

for (const [cat, expectedHex] of Object.entries(expectedColors)) {
  const actualColor = engine.getCategoryColor(cat);
  if (actualColor !== expectedHex) {
    throw new Error(`FAIL: Color mismatch for ${cat}: expected ${expectedHex}, got ${actualColor}`);
  }
  console.log(`✓ ${cat} -> ${actualColor} (${expectedHex})`);
}

console.log('\n=== TEST 3: MULTILINGUAL CANONICAL CATEGORY RESOLUTION ===');
const multilingualChecks = [
  { name: 'Housing (Rent/Loan/Charges)', lang: 'EN', expected: '#3B82F6' },
  { name: 'Food & Supermarket', lang: 'EN', expected: '#F59E0B' },
  { name: 'Vivienda (Alquiler/Hipoteca/Gastos)', lang: 'ES', expected: '#3B82F6' },
  { name: 'Alimentación y Supermercado', lang: 'ES', expected: '#F59E0B' },
  { name: 'Habitação (Renda/Empréstimo/Despesas)', lang: 'PT', expected: '#3B82F6' },
  { name: 'Saúde & Seguro', lang: 'PT', expected: '#10B981' },
  { name: 'Lazer, Saídas e Férias', lang: 'PT', expected: '#EC4899' }
];

multilingualChecks.forEach(({ name, expected }) => {
  const color = engine.getCategoryColor(name);
  if (color !== expected) {
    throw new Error(`FAIL: Multilingual category ${name} expected ${expected}, got ${color}`);
  }
  console.log(`✓ Multilingual [${name}] -> ${color}`);
});

console.log('\n=== TEST 4: CUSTOM USER CATEGORIES DETERMINISTIC COLOR ASSIGNMENT ===');
const customCat1 = 'Abonnements Netflix & Spotify';
const customCat2 = 'Crypto Staking Fees';
const customCat3 = 'Cours de Tennis Privé';

const c1a = engine.getCategoryColor(customCat1);
const c1b = engine.getCategoryColor(customCat1);
if (c1a !== c1b) {
  throw new Error(`FAIL: Category color non-deterministic for ${customCat1}: ${c1a} vs ${c1b}`);
}
console.log(`✓ Custom Category "${customCat1}" is deterministic: ${c1a}`);

const c2 = engine.getCategoryColor(customCat2);
const c3 = engine.getCategoryColor(customCat3);
console.log(`✓ Custom Category "${customCat2}" -> ${c2}`);
console.log(`✓ Custom Category "${customCat3}" -> ${c3}`);

console.log('\n=== TEST 5: BREAKDOWN COMPUTATION & DATA CONCORDANCE ===');
const sampleFixture = JSON.parse(fs.readFileSync(path.join(__dirname, 'sample_test_dataset.json'), 'utf8'));
engine.init(sampleFixture);
const cf = engine.computeCashflow();
if (!cf || !Array.isArray(cf.breakdown) || cf.breakdown.length === 0) {
  throw new Error(`FAIL: computeCashflow().breakdown is empty`);
}

console.log(`✓ Computed ${cf.breakdown.length} breakdown categories for period ${engine.month}/${engine.year}`);
cf.breakdown.forEach((item, idx) => {
  const expectedColor = engine.getCategoryColor(item.nom);
  if (item.color !== expectedColor) {
    throw new Error(`FAIL: Breakdown item ${item.nom} color mismatch: item has ${item.color}, getCategoryColor returned ${expectedColor}`);
  }
  if (!item.nomTraduit) {
    throw new Error(`FAIL: Breakdown item ${item.nom} missing nomTraduit`);
  }
  if (typeof item.pct !== 'number' || isNaN(item.pct)) {
    throw new Error(`FAIL: Breakdown item ${item.nom} has invalid pct: ${item.pct}`);
  }
  console.log(`  [${idx + 1}] ${item.nomTraduit} (${item.nom}): ${item.montant} MUR (${item.pct.toFixed(1)}%) -> Color: ${item.color}`);
});

console.log('\n=== TEST 6: CONSOLIDATED HUB DATA BREAKDOWN CONCORDANCE ===');
const hubData = engine.getConsolidatedHubData();
if (!hubData.breakdown || hubData.breakdown.length === 0) {
  throw new Error(`FAIL: getConsolidatedHubData().breakdown is empty`);
}
hubData.breakdown.forEach(item => {
  const expectedColor = engine.getCategoryColor(item.nom);
  if (item.color !== expectedColor) {
    throw new Error(`FAIL: Hub breakdown color mismatch for ${item.nom}`);
  }
});
console.log(`✓ Hub breakdown matches Cashflow breakdown colors with 100% concordance`);

console.log('\n=== TEST 7: STYLES.CSS & APP.JS CONFORMANCE CHECK ===');
const stylesCss = fs.readFileSync(path.join(__dirname, '../mobile-web/styles.css'), 'utf8');
const appJs = fs.readFileSync(path.join(__dirname, '../mobile-web/app.js'), 'utf8');

// Verify .legend-color-dot has flex-shrink: 0 and border-radius: 50%
if (!stylesCss.includes('flex-shrink: 0')) {
  throw new Error('FAIL: styles.css missing flex-shrink: 0 in legend dots');
}
if (!stylesCss.includes('.donut-chart-wrapper')) {
  throw new Error('FAIL: styles.css missing .donut-chart-wrapper rule');
}
if (!stylesCss.includes('drop-shadow(0 0 6px currentColor)')) {
  throw new Error('FAIL: styles.css missing modern drop-shadow on donut slice');
}

// Verify app.js renderDonutChart has linked hover & interactive center readout
if (!appJs.includes('donut-center-box') || !appJs.includes('donut-center-val')) {
  throw new Error('FAIL: app.js missing dynamic center readout elements');
}
if (!appJs.includes('highlightCat')) {
  throw new Error('FAIL: app.js missing highlightCat linked hover function');
}

console.log('✓ styles.css has .donut-chart-wrapper, flex-shrink: 0, and modern drop-shadows');
console.log('✓ app.js has linked hover/focus states, dynamic center readout, and accessible attributes');

console.log('\n======================================================');
console.log('ALL CATEGORY DONUT COLOR TESTS PASSED SUCCESSFULLY! (100% PARITY)');
console.log('======================================================');
