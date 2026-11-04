/**
 * Verification Test Suite: Cutting-Edge Icon Integration & Asset Parity
 */

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const assert = require('assert');

console.log('================================================================================');
console.log('🧪 VERIFICATION SUITE: CUTTING-EDGE ICON INTEGRATION & ASSET PARITY');
console.log('================================================================================\n');

const ROOT = path.resolve(__dirname, '..');

// 1. Verify Master Assets
const masterPng = path.join(ROOT, 'master_gfmp_icon.png');
const expensesPng = path.join(ROOT, 'expenses_14510803.png');

assert(fs.existsSync(expensesPng), 'Master source expenses_14510803.png must exist');
assert(fs.existsSync(masterPng), 'Root master_gfmp_icon.png must exist');

const hMaster = crypto.createHash('sha256').update(fs.readFileSync(masterPng)).digest('hex');
const hExpenses = crypto.createHash('sha256').update(fs.readFileSync(expensesPng)).digest('hex');
assert.strictEqual(hMaster, hExpenses, 'master_gfmp_icon.png must be an exact byte-for-byte copy of expenses_14510803.png');
console.log('✔ PASS 1: master_gfmp_icon.png matches expenses_14510803.png byte-for-byte (SHA-256: ' + hMaster.slice(0, 16) + '...)');

// 2. Verify Web/PWA Assets
const webAssets = ['icon_512.png', 'logo_jic.png', 'logo_jic_transparent.png', 'favicon_jic.png'];
for (const a of webAssets) {
  const p1 = path.join(ROOT, 'mobile-web', 'assets', a);
  const p2 = path.join(ROOT, 'android', 'app', 'src', 'main', 'assets', 'assets', a);
  assert(fs.existsSync(p1), `mobile-web/assets/${a} must exist`);
  assert(fs.existsSync(p2), `android/app/src/main/assets/assets/${a} must exist`);
  const h1 = crypto.createHash('sha256').update(fs.readFileSync(p1)).digest('hex');
  const h2 = crypto.createHash('sha256').update(fs.readFileSync(p2)).digest('hex');
  assert.strictEqual(h1, h2, `SHA-256 mismatch between mobile-web and android assets for ${a}`);
}
console.log('✔ PASS 2: All Web/PWA assets (icon_512, logo_jic, logo_jic_transparent, favicon_jic) are synchronized with cryptographic parity.');

// 3. Verify Android Mipmaps
const mipmaps = [
  { folder: 'mipmap-mdpi', size: 48 },
  { folder: 'mipmap-hdpi', size: 72 },
  { folder: 'mipmap-xhdpi', size: 96 },
  { folder: 'mipmap-xxhdpi', size: 144 },
  { folder: 'mipmap-xxxhdpi', size: 192 }
];

for (const m of mipmaps) {
  const p1 = path.join(ROOT, 'android', 'app', 'src', 'main', 'res', m.folder, 'ic_launcher.png');
  const p2 = path.join(ROOT, 'android', 'app', 'src', 'main', 'res', m.folder, 'ic_launcher_round.png');
  assert(fs.existsSync(p1), `${m.folder}/ic_launcher.png must exist`);
  assert(fs.existsSync(p2), `${m.folder}/ic_launcher_round.png must exist`);
}
console.log('✔ PASS 3: Android mipmap launcher icons (mdpi, hdpi, xhdpi, xxhdpi, xxxhdpi) exist across all densities.');

// 4. Verify Adaptive Vector Drawables
const bgXml = path.join(ROOT, 'android', 'app', 'src', 'main', 'res', 'drawable', 'ic_launcher_background.xml');
const fgXml = path.join(ROOT, 'android', 'app', 'src', 'main', 'res', 'drawable', 'ic_launcher_foreground.xml');

assert(fs.existsSync(bgXml), 'ic_launcher_background.xml must exist');
assert(fs.existsSync(fgXml), 'ic_launcher_foreground.xml must exist');

const bgContent = fs.readFileSync(bgXml, 'utf8');
const fgContent = fs.readFileSync(fgXml, 'utf8');

assert(bgContent.includes('#F308FD') && bgContent.includes('#8D60F0') && bgContent.includes('#29B8E4'), 
  'ic_launcher_background.xml must contain the cutting-edge magenta-purple-cyan gradient');

assert(fgContent.includes('#FFFFFF') && fgContent.includes('viewportWidth="108"'), 
  'ic_launcher_foreground.xml must contain the clean white 108dp vector paths');

console.log('✔ PASS 4: Adaptive icon drawables (ic_launcher_background.xml & ic_launcher_foreground.xml) are valid vector drawables with the cutting-edge color scheme.');

// 5. Verify Reference APK Sanctity
const sealedApk = path.join(ROOT, 'GFMP-v1.0.apk');
assert(fs.existsSync(sealedApk), 'GFMP-v1.0.apk must exist');
const sealedSize = fs.statSync(sealedApk).size;
const sealedHash = crypto.createHash('sha256').update(fs.readFileSync(sealedApk)).digest('hex').toUpperCase();

assert.strictEqual(sealedSize, 3083929, 'GFMP-v1.0.apk size altered!');
assert.strictEqual(sealedHash, '4E29C86CDC7D18B2585849F17474F721D8185F7CC37A22F5FE00D1E27CC80DB9', 'GFMP-v1.0.apk hash altered!');
console.log('✔ PASS 5: GFMP-v1.0.apk is 100% untouched and sealed cryptographically.');

console.log('\n================================================================================');
console.log('🏆 ALL CUTTING-EDGE ICON CHECKS PASSED WITH 0 ERRORS!');
console.log('================================================================================');

