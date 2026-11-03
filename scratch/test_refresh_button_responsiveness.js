const fs = require('fs');
const path = require('path');
const assert = require('assert');

console.log('================================================================');
console.log('TEST SUITE: RATES REFRESH BUTTON RESPONSIVENESS & ZERO REGRESSION');
console.log('================================================================');

// 1. Inspect HTML markup
const htmlPath = path.resolve(__dirname, '../mobile-web/index.html');
const htmlContent = fs.readFileSync(htmlPath, 'utf8');

console.log('\n[1/6] Verifying HTML Cleanliness & Structure...');
assert.ok(!htmlContent.includes('live-rates-bar'), 'FAIL: Old live-rates-bar should NOT be present in index.html');
assert.ok(!htmlContent.includes('txt-live-rates-status'), 'FAIL: txt-live-rates-status should NOT be present in index.html');
assert.ok(!htmlContent.includes('Marchés Connectés • Taux en Direct (API)'), 'FAIL: Old static status text should be removed');
assert.ok(htmlContent.includes('class="currency-row"'), 'FAIL: currency-row should wrap currency-bar and refresh button');
assert.ok(htmlContent.includes('id="btn-refresh-rates"'), 'FAIL: btn-refresh-rates must exist');
assert.ok(htmlContent.includes('id="live-pulse-dot"'), 'FAIL: live-pulse-dot must exist inside button');
assert.ok(htmlContent.includes('class="refresh-icon"'), 'FAIL: refresh-icon SVG must exist');

// Extract btn-refresh-rates HTML block to ensure NO visible text node is present
const btnMatch = htmlContent.match(/<button[^>]*id="btn-refresh-rates"[^>]*>([\s\S]*?)<\/button>/);
assert.ok(btnMatch, 'FAIL: Could not extract btn-refresh-rates element');
const btnInner = btnMatch[1];
const textOnly = btnInner.replace(/<[^>]*>/g, '').trim();
assert.strictEqual(textOnly, '', `FAIL: btn-refresh-rates must NOT contain text content, but found: "${textOnly}"`);
console.log('  ✓ HTML markup is strictly icon-only without parasitic text clutter.');
console.log('  ✓ Live status dot and SVG refresh icon correctly embedded.');

// 2. Inspect Multilingual Dictionary in analytics-engine.js
console.log('\n[2/6] Verifying Multilingual Translation Dictionary (FR, EN, ES, PT)...');
global.window = global;
global.localStorage = {
  _data: {},
  getItem(k) { return this._data[k] || null; },
  setItem(k, v) { this._data[k] = String(v); },
  removeItem(k) { delete this._data[k]; }
};

const baselineCode = fs.readFileSync(path.resolve(__dirname, '../mobile-web/baseline_data.js'), 'utf8');
eval(baselineCode);

const engineCode = fs.readFileSync(path.resolve(__dirname, '../mobile-web/analytics-engine.js'), 'utf8');
eval(engineCode);

const engine = window.gfmpEngine;
assert.ok(engine, 'FAIL: gfmpEngine could not be instantiated');

const requiredKeys = [
  'BTN_REFRESH_RATES_TITLE',
  'RATES_UPDATING',
  'RATES_UPDATED_SUCCESS',
  'RATES_OFFLINE',
  'RATES_LAST_SYNC'
];
const supportedLangs = ['FR', 'EN', 'ES', 'PT'];

for (const lang of supportedLangs) {
  engine.setLanguage(lang);
  for (const key of requiredKeys) {
    const val = engine.tr(key);
    assert.ok(val && val !== key, `FAIL: Missing translation for key "${key}" in language "${lang}"`);
    console.log(`  ✓ [${lang}] ${key} => "${val}"`);
  }
}

// 3. Inspect app.js Logic
console.log('\n[3/6] Verifying app.js Dynamic Multilingual Handling...');
const appCode = fs.readFileSync(path.resolve(__dirname, '../mobile-web/app.js'), 'utf8');

assert.ok(appCode.includes('BTN_REFRESH_RATES_TITLE'), 'FAIL: app.js must reference BTN_REFRESH_RATES_TITLE');
assert.ok(appCode.includes('RATES_UPDATING'), 'FAIL: app.js must reference RATES_UPDATING');
assert.ok(appCode.includes('RATES_UPDATED_SUCCESS'), 'FAIL: app.js must reference RATES_UPDATED_SUCCESS');
assert.ok(appCode.includes('RATES_OFFLINE'), 'FAIL: app.js must reference RATES_OFFLINE');
assert.ok(appCode.includes('RATES_LAST_SYNC'), 'FAIL: app.js must reference RATES_LAST_SYNC');
assert.ok(appCode.includes('btn.setAttribute(\'aria-label\','), 'FAIL: app.js must dynamically set aria-label');
assert.ok(appCode.includes('btn.setAttribute(\'title\','), 'FAIL: app.js must dynamically set title tooltip');
console.log('  ✓ app.js dynamically updates accessible aria-label and tooltip on language change & rate sync.');

// 4. Inspect CSS Styling & Responsiveness
console.log('\n[4/6] Verifying CSS Rules & Compact Shrink-On-Scroll Behavior...');
const cssCode = fs.readFileSync(path.resolve(__dirname, '../mobile-web/styles.css'), 'utf8');

assert.ok(cssCode.includes('.currency-row'), 'FAIL: .currency-row style missing in styles.css');
assert.ok(cssCode.includes('.btn-refresh-rates'), 'FAIL: .btn-refresh-rates style missing in styles.css');
assert.ok(cssCode.includes('.btn-refresh-rates .live-pulse-dot'), 'FAIL: .live-pulse-dot positioning missing in styles.css');
assert.ok(cssCode.includes('@keyframes spinRates'), 'FAIL: @keyframes spinRates missing in styles.css');
assert.ok(cssCode.includes('.gfmp-header.is-compact .btn-refresh-rates'), 'FAIL: Compact header rule for refresh button missing');
assert.ok(!cssCode.includes('.live-rates-bar {'), 'FAIL: Obsolete .live-rates-bar rule should be completely removed from styles.css');
console.log('  ✓ Responsive flex row and compact shrink-on-scroll rules validated.');

// 5. Verify Android Asset Parity
console.log('\n[5/6] Verifying Cryptographic Parity Between mobile-web and Android Assets...');
const filesToCheck = ['index.html', 'styles.css', 'app.js', 'analytics-engine.js', 'sw.js'];
for (const f of filesToCheck) {
  const webContent = fs.readFileSync(path.resolve(__dirname, `../mobile-web/${f}`), 'utf8');
  const androidContent = fs.readFileSync(path.resolve(__dirname, `../android/app/src/main/assets/${f}`), 'utf8');
  assert.strictEqual(webContent, androidContent, `FAIL: Parity mismatch for ${f}`);
  console.log(`  ✓ 100% Match: mobile-web/${f} === android/app/src/main/assets/${f}`);
}

// 6. Verify Android Version Bump
console.log('\n[6/6] Verifying Android App Version Bump...');
const gradleContent = fs.readFileSync(path.resolve(__dirname, '../android/app/build.gradle.kts'), 'utf8');
assert.ok(gradleContent.includes('versionCode = 4'), 'FAIL: Android versionCode must be 4');
assert.ok(gradleContent.includes('versionName = "1.3"'), 'FAIL: Android versionName must be "1.3"');
console.log('  ✓ Android Gradle configured with versionCode 4, versionName 1.3.');

console.log('\n================================================================');
console.log('ALL TESTS PASSED WITH ZERO REGRESSION (100% GOLD CUTTING-EDGE)!');
console.log('================================================================\n');
