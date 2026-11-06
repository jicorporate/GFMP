/**
 * Verification Test Suite: Cutting-Edge Icon Theme Integration
 */

const fs = require('fs');
const path = require('path');
const assert = require('assert');
const crypto = require('crypto');

console.log('================================================================================');
console.log('🧪 VERIFICATION SUITE: CUTTING-EDGE ICON THEME (GOLD STANDARD)');
console.log('================================================================================\n');

const ROOT = path.resolve(__dirname, '..');
const css = fs.readFileSync(path.join(ROOT, 'mobile-web', 'styles.css'), 'utf8');
const html = fs.readFileSync(path.join(ROOT, 'mobile-web', 'index.html'), 'utf8');
const engineJs = fs.readFileSync(path.join(ROOT, 'mobile-web', 'analytics-engine.js'), 'utf8');

// 1. Verify CSS styles for Cutting-Edge Theme
assert(css.includes('body.theme-cutting-edge'), 'body.theme-cutting-edge must be defined in styles.css');
assert(css.includes('#F308FD') && css.includes('#8D60F0') && css.includes('#29B8E4'), 
  'styles.css cutting-edge theme must use the exact icon palette (Magenta #F308FD, Violet #8D60F0, Cyan #29B8E4)');
console.log('✔ PASS 1: CSS definitions for Cutting-Edge Theme match exact icon color palette.');

// 2. Verify Theme selector in Settings & Onboarding
assert(html.includes('data-theme="cutting-edge"'), 'Settings theme selector must contain cutting-edge button');
assert(html.includes('data-init-theme="cutting-edge"'), 'Onboarding theme grid must contain cutting-edge card');
assert(/class="init-choice-card\s+active"[^>]*data-init-theme="cutting-edge"/.test(html) || /data-init-theme="cutting-edge"[^>]*class="init-choice-card\s+active"/.test(html), 'data-init-theme="cutting-edge" must be active by default in onboarding HTML');
assert(!/class="init-choice-card\s+active"[^>]*data-init-theme="blue"/.test(html) && !/data-init-theme="blue"[^>]*class="init-choice-card\s+active"/.test(html), 'data-init-theme="blue" must NOT be active by default in onboarding HTML');
assert(html.includes('<body class="theme-cutting-edge">'), 'index.html body must have class theme-cutting-edge by default');
console.log('✔ PASS 2: Cutting-Edge Theme controls exist in Settings and Startup Onboarding.');

// 3. Verify Engine translations
assert(engineJs.includes("'THEME_NAME_CUTTING_EDGE'") && engineJs.includes("'THEME_NAME_CUTTING-EDGE'"), 
  'analytics-engine.js must contain THEME_NAME_CUTTING_EDGE translations');
assert(engineJs.includes("'THEME_SUB_CUTTING_EDGE'") && engineJs.includes("'THEME_SUB_CUTTING-EDGE'"), 
  'analytics-engine.js must contain THEME_SUB_CUTTING_EDGE translations');

const store = {};
global.window = {
  localStorage: {
    getItem: (k) => (k in store ? store[k] : null),
    setItem: (k, v) => { store[k] = String(v); },
    removeItem: (k) => { delete store[k]; },
    clear: () => { for (const k in store) delete store[k]; }
  }
};
global.localStorage = global.window.localStorage;

const baselineCode = fs.readFileSync(path.join(ROOT, 'mobile-web', 'baseline_data.js'), 'utf8');
eval(baselineCode);
eval(engineJs);
const engine = window.gfmpEngine;
engine.init();

for (const lang of ['FR', 'EN', 'ES', 'PT']) {
  engine.setLanguage(lang);
  const name = engine.tr('THEME_NAME_CUTTING_EDGE');
  const sub = engine.tr('THEME_SUB_CUTTING_EDGE');
  assert(name && sub, `Missing cutting-edge theme translation in ${lang}`);
}
console.log('✔ PASS 3: Quadrilingual translations for Cutting-Edge Theme verified across FR, EN, ES, PT.');

// 4. Verify initializeSystem supports cutting-edge
assert.strictEqual(engine.theme, 'cutting-edge', 'GFMPEngine default theme must be cutting-edge');
engine.initializeSystem('EUR', 'FR', '2026-02-25');
assert.strictEqual(engine.theme, 'cutting-edge', 'initializeSystem must default theme to cutting-edge when unspecified');

engine.initializeSystem('EUR', 'FR', '2026-02-25', 'cutting-edge');
assert.strictEqual(engine.theme, 'cutting-edge', 'initializeSystem must set theme to cutting-edge');

engine.initializeSystem('USD', 'EN', '2026-02-25', 'neon');
assert.strictEqual(engine.theme, 'cutting-edge', 'initializeSystem must normalize neon alias to cutting-edge');
console.log('✔ PASS 4: initializeSystem successfully handles and normalizes cutting-edge theme.');

// 5. Verify reference APK sanctity
const sealedApk = path.join(ROOT, 'GFMP-v1.0.apk');
assert.strictEqual(fs.statSync(sealedApk).size, 3083929, 'GFMP-v1.0.apk size altered!');
const h = crypto.createHash('sha256').update(fs.readFileSync(sealedApk)).digest('hex').toUpperCase();
assert.strictEqual(h, '4E29C86CDC7D18B2585849F17474F721D8185F7CC37A22F5FE00D1E27CC80DB9', 'GFMP-v1.0.apk hash altered!');
console.log('✔ PASS 5: Reference APK GFMP-v1.0.apk strictly intact.');

console.log('\n================================================================================');
console.log('🏆 ALL CUTTING-EDGE THEME CHECKS PASSED WITH 0 ERRORS!');
console.log('================================================================================');
