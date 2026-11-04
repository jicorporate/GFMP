/**
 * Verification Test Suite: Saisie Express Button & Event Handling Verification
 */

const fs = require('fs');
const path = require('path');
const assert = require('assert');

console.log('================================================================================');
console.log('🧪 VERIFICATION SUITE: SAISIE EXPRESS BUTTON & INTERACTION (GOLD STANDARD)');
console.log('================================================================================\n');

// 1. Verify index.html button existence and accessibility attributes
const htmlPath = path.resolve(__dirname, '../mobile-web/index.html');
const htmlContent = fs.readFileSync(htmlPath, 'utf8');

assert(htmlContent.includes('id="btn-hub-quick-saisie"'), 'Missing id="btn-hub-quick-saisie" in index.html');
assert(htmlContent.includes('class="btn-allouer" id="btn-hub-quick-saisie"'), 'Missing class="btn-allouer" on #btn-hub-quick-saisie');
assert(htmlContent.includes('id="modal-bottom-sheet-tx"'), 'Missing #modal-bottom-sheet-tx modal in index.html');
console.log('✔ PASS 1: #btn-hub-quick-saisie and #modal-bottom-sheet-tx exist in index.html with Gold Standard luxury styling.');

// 2. Verify app.js event binding directly calls openExpressModal
const appJsPath = path.resolve(__dirname, '../mobile-web/app.js');
const appJsContent = fs.readFileSync(appJsPath, 'utf8');

assert(!appJsContent.includes("document.getElementById('btn-open-saisie-express')?.click()"), 
  'Found obsolete synthetic click on non-existent element btn-open-saisie-express!');

assert(appJsContent.includes("document.getElementById('btn-hub-quick-saisie')?.addEventListener('click', (e) => {\n    e.preventDefault();\n    e.stopPropagation();\n    openExpressModal();\n  });"),
  '#btn-hub-quick-saisie not properly hooked to openExpressModal() with preventDefault/stopPropagation');

assert(appJsContent.includes("validAccountId = (typeof defaultAccountId === 'string' || typeof defaultAccountId === 'number')"),
  'openExpressModal() lacks MouseEvent guard for defaultAccountId');

console.log('✔ PASS 2: #btn-hub-quick-saisie is cleanly and directly hooked to openExpressModal() with zero synthetic dependencies.');
console.log('✔ PASS 3: openExpressModal() guards against MouseEvent argument pollution of account select.');

// 3. Verify Android assets synchronization
const androidHtml = fs.readFileSync(path.resolve(__dirname, '../android/app/src/main/assets/index.html'), 'utf8');
const androidAppJs = fs.readFileSync(path.resolve(__dirname, '../android/app/src/main/assets/app.js'), 'utf8');

assert.strictEqual(htmlContent, androidHtml, 'Android index.html asset is out of sync with mobile-web/index.html');
assert.strictEqual(appJsContent, androidAppJs, 'Android app.js asset is out of sync with mobile-web/app.js');
console.log('✔ PASS 4: Android assets are 100% byte-for-byte in sync with mobile-web.');

console.log('\n================================================================================');
console.log('🏆 ALL SAISIE EXPRESS BUTTON CHECKS PASSED WITH 0 ERRORS!');
console.log('================================================================================');

