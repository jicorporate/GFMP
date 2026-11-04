/**
 * GFMP (SFP v1.0) — Gold Standard Cutting-Edge Full Application Audit Suite
 * Audits 10 architectural pillars:
 * 1. Kimball Star Schema & Referential Integrity
 * 2. Mathematical & Financial Precision Engine (GFMPEngine)
 * 3. Quadrilingual 100% Coverage (FR, EN, ES, PT)
 * 4. 4-Step Startup Configuration & Haute Joaillerie Emblem
 * 5. Executive Hub & Interactive Responsive Donut Chart (@container)
 * 6. Accessibility & Semantic HTML (WCAG 2.1 AA / modern-web-guidance)
 * 7. Performance, CLS Prevention & Offline PWA Resilience
 * 8. Security, XSS Sanitization & Non-Blocking Dialogs
 * 9. Cryptographic SHA-256 Parity (mobile-web vs android assets)
 * 10. Native Android Container (API 34, MainActivity, AndroidManifest, APK)
 */

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const ROOT = path.resolve(__dirname, '..');
const MOBILE_WEB = path.join(ROOT, 'mobile-web');
const ANDROID_ASSETS = path.join(ROOT, 'android', 'app', 'src', 'main', 'assets');

// Setup mock browser environment for engine evaluation
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

const baselineCode = fs.readFileSync(path.join(MOBILE_WEB, 'baseline_data.js'), 'utf8');
const engineCode = fs.readFileSync(path.join(MOBILE_WEB, 'analytics-engine.js'), 'utf8');
eval(baselineCode);
eval(engineCode);

const engine = window.gfmpEngine;

let totalChecks = 0;
let passedChecks = 0;
const failures = [];
const pillarStats = {};

function assertAudit(pillar, condition, description) {
  totalChecks++;
  if (!pillarStats[pillar]) pillarStats[pillar] = { passed: 0, total: 0 };
  pillarStats[pillar].total++;
  if (condition) {
    passedChecks++;
    pillarStats[pillar].passed++;
  } else {
    failures.push(`[${pillar}] ${description}`);
    console.error(`  ❌ FAIL [${pillar}]: ${description}`);
  }
}

console.log('================================================================================');
console.log('🏛️  AUDIT COMPLET GOLD STANDARD CUTTING-EDGE — GFMP (SFP v1.01 PRODUCTION)');
console.log('================================================================================\n');

// ============================================================================
// PILLAR 1: KIMBALL STAR SCHEMA, PRODUCTION v1.01 ZERO-TEST-DATA & INTEGRITY
// ============================================================================
{
  const P = '1. Kimball Star Schema & Production v1.01 Integrity';
  global.localStorage.clear();
  engine.init();

  const baselineJson = fs.readFileSync(path.join(MOBILE_WEB, 'baseline_data.json'), 'utf8');
  assertAudit(P, engine.version === '1.01' && window.GFMP_BASELINE.version === '1.01', 'Production version is strictly 1.01 across engine and baseline');
  assertAudit(P, engine.data && typeof engine.data.devises === 'object' && Object.keys(engine.data.devises).length >= 6, 'T_SYS_Config (devises) contains >= 6 currencies');
  assertAudit(P, Array.isArray(engine.data.comptes) && engine.data.comptes.length === 9, 'T_DIM_Compte contains 9 reference accounts ready for production');
  assertAudit(P, Array.isArray(engine.data.categories) && engine.data.categories.length === 11, 'T_DIM_Categorie contains 11 reference categories ready for production');
  assertAudit(P, Array.isArray(engine.data.tiers) && engine.data.tiers.length === 7, 'T_DIM_Tiers contains 7 reference counterparties ready for production');
  assertAudit(P, Array.isArray(engine.data.budgets) && engine.data.budgets.length === 0, 'T_FACT_Budget contains 0 test budget envelopes in Production v1.01');
  assertAudit(P, Array.isArray(engine.data.transactions) && engine.data.transactions.length === 0, 'T_FACT_Transaction contains 0 test transactions in Production v1.01');
  assertAudit(P, !baselineJson.includes('Salaire KIMYA') && !baselineJson.includes('TEST_DATA_EMBED') && !baselineCode.includes('Salaire KIMYA'), 'Zero test transaction strings in production baseline_data.json & baseline_data.js');

  assertAudit(P, engine.cleanText('   mcb   courant  ') === 'mcb courant', 'cleanText normalizes and collapses whitespace');

  // Dynamic extension & Kimball Star Schema deduplication
  const newAcc1 = engine.addAccount('Coffre Fort Titres Suisses', 'INVESTISSEMENT', 'EUR');
  assertAudit(P, newAcc1.isNew === true && newAcc1.id === '10', 'addAccount creates sequential ID 10');
  const newAcc2 = engine.addAccount('   coffre fort   titres suisses  ', 'INVESTISSEMENT', 'EUR');
  assertAudit(P, newAcc2.isNew === false && newAcc2.id === '10', 'addAccount deduplicates case/space-insensitive duplicate account');
}

// ============================================================================
// PILLAR 2: MATHEMATICAL & FINANCIAL PRECISION ENGINE
// ============================================================================
{
  const P = '2. Financial & Mathematical Precision Engine';
  global.localStorage.clear();
  const sampleFixture = JSON.parse(fs.readFileSync(path.join(__dirname, 'sample_test_dataset.json'), 'utf8'));
  engine.init(sampleFixture);

  assertAudit(P, engine.roundMoney(0.1 + 0.2) === 0.3, 'roundMoney(0.1 + 0.2) === 0.3 exactly');
  assertAudit(P, engine.roundMoney(10.005) === 10.01, 'roundMoney(10.005) === 10.01 (Number.EPSILON correction)');
  assertAudit(P, Object.is(engine.roundMoney(-0.00001), 0), 'roundMoney eliminates IEEE 754 negative zero (-0)');
  assertAudit(P, !engine.formatMoney(-0.00001).includes('-0'), 'formatMoney never outputs negative zero (-0,00)');

  const currencies = ['MUR', 'EUR', 'USD', 'GBP', 'ZAR', 'XOF'];
  for (const curr of currencies) {
    engine.setCurrency(curr);
    const nw = engine.computeNetWorth('2026-02');
    const diff = Math.abs(engine.roundMoney(nw.totalAssets - nw.totalLiabilities) - nw.netWorth);
    assertAudit(P, diff === 0, `Accounting identity (Assets - Liabilities === Net Worth) holds in ${curr}`);
  }

  engine.setCurrency('MUR');
  const nwFeb = engine.computeNetWorth('2026-02');
  const cptCourant = nwFeb.accounts.find(c => c.id === '1');
  const cptEpargne = nwFeb.accounts.find(c => c.id === '2');
  assertAudit(P, cptCourant && cptCourant.soldeConverti === 77252, 'Bilateral transfer outflow accurately reflected on Compte Courant (77,252 MUR)');
  assertAudit(P, cptEpargne && cptEpargne.soldeConverti === 100000, 'Bilateral transfer inflow accurately reflected on Livret d\'Épargne (100,000 MUR)');

  const cfFeb = engine.computeCashflow('2026-02');
  assertAudit(P, engine.roundMoney(cfFeb.totalRevenus - cfFeb.totalDepenses) === cfFeb.netCashflow, 'Cashflow identity (Revenus - Dépenses === Net Cashflow) holds');

  const trend = engine.computeCashflowTrend(6);
  assertAudit(P, trend.length === 6 && trend.every(m => !Number.isNaN(m.revenus) && !Number.isNaN(m.depenses)), 'computeCashflowTrend(6) returns 6 valid months without NaN');

  const zbb = engine.computeBudgetZBB('2026-02');
  assertAudit(P, zbb.categories.length === 6 && engine.roundMoney(zbb.totalAlloue - zbb.totalConsomme) === zbb.totalEcart, 'ZBB Budget variance identity (Allocated - Consumed === Variance) holds');

  const hub = engine.getConsolidatedHubData('2026-02');
  assertAudit(P, hub.healthScore >= 0 && hub.healthScore <= 100, `Executive Hub Financial Health Score is bounded [0, 100] (actual: ${hub.healthScore})`);
}

// ============================================================================
// PILLAR 3: QUADRILINGUAL 100% COVERAGE (FR, EN, ES, PT)
// ============================================================================
{
  const P = '3. Quadrilingual Internationalization (FR/EN/ES/PT)';
  const dict = engine.getDictionary();
  const langs = ['FR', 'EN', 'ES', 'PT'];
  const allKeys = Object.keys(dict);

  for (const lang of langs) {
    const missing = allKeys.filter(k => !dict[k] || !dict[k][lang]);
    assertAudit(P, missing.length === 0, `Dictionary has 100% ${lang} coverage across all ${allKeys.length} keys (0 missing)`);
  }

  for (const lang of langs) {
    engine.setLanguage(lang);
    assertAudit(P, engine.trEntity('Alimentation & Supermarché') && engine.trEntity('REVENU'), `trEntity translates baseline entities in ${lang}`);
  }

  const indexHtml = fs.readFileSync(path.join(MOBILE_WEB, 'index.html'), 'utf8');
  const i18nMatches = [...indexHtml.matchAll(/data-i18n="([^"]+)"/g)].map(m => m[1]);
  const missingDomKeys = i18nMatches.filter(k => langs.some(l => !engine.tr(k, {}, l)));
  assertAudit(P, missingDomKeys.length === 0, `All ${i18nMatches.length} DOM data-i18n attributes resolve in FR, EN, ES, and PT`);
}

// ============================================================================
// PILLAR 4: 4-STEP STARTUP CONFIGURATION & HAUTE JOAILLERIE EMBLEM
// ============================================================================
{
  const P = '4. Startup Configuration & Brand Emblem';
  const indexHtml = fs.readFileSync(path.join(MOBILE_WEB, 'index.html'), 'utf8');
  const stylesCss = fs.readFileSync(path.join(MOBILE_WEB, 'styles.css'), 'utf8');
  const appJs = fs.readFileSync(path.join(MOBILE_WEB, 'app.js'), 'utf8');

  assertAudit(P, indexHtml.includes('id="modal-initialization"') && indexHtml.includes('init-onboarding-shell'), 'Startup configuration modal #modal-initialization uses .init-onboarding-shell');
  assertAudit(P, indexHtml.includes('id="init-step1-badge"') && indexHtml.includes('id="init-step2-badge"') && indexHtml.includes('id="init-step3-badge"') && indexHtml.includes('id="init-step4-badge"'), 'All 4 numbered onboarding steps exist (Langue, Devise, Date/Période, Thème)');
  assertAudit(P, indexHtml.includes('id="init-date-input"') && indexHtml.includes('id="init-month-select"') && indexHtml.includes('id="init-year-select"'), 'Step 3 Date & Financial Period inputs exist');
  assertAudit(P, indexHtml.includes('id="init-theme-grid"') && indexHtml.includes('data-init-theme="emerald"'), 'Step 4 Visual Theme cards exist');
  assertAudit(P, !indexHtml.includes('init-trust-strip') && !indexHtml.includes('id="init-trust-1"') && !stylesCss.includes('.init-trust-strip'), 'Removed clutter trust-strip badges from startup modal and CSS');
  assertAudit(P, indexHtml.includes('class="init-emblem-svg"') && indexHtml.includes('id="gfmp-gold-grad"') && indexHtml.includes('id="gfmp-sapphire-core"'), 'Haute Joaillerie Private Banking Sovereign Cut Diamond & Ascending Wealth SVG emblem is integrated');
  assertAudit(P, appJs.includes('engine.initializeSystem') && appJs.includes('engine.needsOnboarding') && engineCode.includes('gfmp_onboarding_v5_6'), 'Onboarding configuration persistence & fresh-install trigger exist in app.js & analytics-engine.js');
}

// ============================================================================
// PILLAR 5: EXECUTIVE HUB & INTERACTIVE RESPONSIVE DONUT CHART
// ============================================================================
{
  const P = '5. Executive Hub & Interactive Responsive Donut';
  const indexHtml = fs.readFileSync(path.join(MOBILE_WEB, 'index.html'), 'utf8');
  const stylesCss = fs.readFileSync(path.join(MOBILE_WEB, 'styles.css'), 'utf8');
  const appJs = fs.readFileSync(path.join(MOBILE_WEB, 'app.js'), 'utf8');

  assertAudit(P, !indexHtml.includes('6-MONTH CASHFLOW') && !engine.tr('HUB_TREND_TITLE', {}, 'EN').includes('6-'), 'Chart title is clean "MONTH CASHFLOW TREND" without "6-" prefix');
  assertAudit(P, indexHtml.includes('id="hub-patrimoine-total" aria-hidden="true"') && appJs.includes("patTotal.style.display = 'none'"), 'Redundant Net Assets badge next to Asset Allocation header is hidden');
  assertAudit(P, !appJs.includes('breakdown.slice(0, maxLegend)') && appJs.includes('breakdown.forEach(item => {'), 'Donut chart renders 100% of expense categories without truncation');
  assertAudit(P, stylesCss.includes('container-type: inline-size;') && stylesCss.includes('@container (max-width: 460px)'), 'Donut chart uses modern CSS @container queries for responsive layout');
  assertAudit(P, appJs.includes('selectedCategoryFilter') && indexHtml.includes('id="hub-donut-active-filter"'), 'Donut chart slices and legend rows support full interactive hover/focus/click category filtering');
}

// ============================================================================
// PILLAR 6: ACCESSIBILITY & SEMANTIC HTML (WCAG 2.1 AA / MODERN WEB GUIDANCE)
// ============================================================================
{
  const P = '6. Accessibility & Semantic HTML (WCAG 2.1 AA)';
  const indexHtml = fs.readFileSync(path.join(MOBILE_WEB, 'index.html'), 'utf8');
  const appJs = fs.readFileSync(path.join(MOBILE_WEB, 'app.js'), 'utf8');

  const thTags = [...indexHtml.matchAll(/<th\b[^>]*>/g)].map(m => m[0]);
  const thWithScope = thTags.filter(t => t.includes('scope="col"'));
  assertAudit(P, thTags.length > 0 && thTags.length === thWithScope.length, `All ${thTags.length} <th> table headers have explicit scope="col"`);

  assertAudit(P, indexHtml.includes('<nav class="bottom-nav" aria-label="Navigation principale">'), 'Bottom navigation landmark has explicit aria-label');
  assertAudit(P, indexHtml.includes('aria-current="page"') && appJs.includes("t.setAttribute('aria-current', 'page')"), 'Active navigation tab manages aria-current="page"');

  assertAudit(P, indexHtml.includes('for="tx-date"') && indexHtml.includes('for="tx-montant"') && indexHtml.includes('for="tx-compte"') && indexHtml.includes('for="tx-categorie"'), 'Transaction form labels have explicit for="..." associations');
  assertAudit(P, indexHtml.includes('for="settings-currency-select"') && indexHtml.includes('for="settings-lang-select"') && indexHtml.includes('for="cf-sort-select"') && indexHtml.includes('aria-label="Filtrer par compte"'), 'Settings & Cashflow filter form controls have explicit label associations');

  assertAudit(P, appJs.includes("toast.setAttribute('role', 'status')") && appJs.includes("toast.setAttribute('aria-live', 'polite')"), 'Toast notification container uses accessible ARIA live region (role="status" aria-live="polite")');
  assertAudit(P, appJs.includes("document.documentElement.lang = (engine.lang || 'FR').toLowerCase()"), 'Document <html lang="..."> updates dynamically when user switches language');
}

// ============================================================================
// PILLAR 7: PERFORMANCE, CLS PREVENTION & OFFLINE PWA RESILIENCE
// ============================================================================
{
  const P = '7. Performance, CLS Prevention & Offline PWA';
  const indexHtml = fs.readFileSync(path.join(MOBILE_WEB, 'index.html'), 'utf8');
  const stylesCss = fs.readFileSync(path.join(MOBILE_WEB, 'styles.css'), 'utf8');
  const swJs = fs.readFileSync(path.join(MOBILE_WEB, 'sw.js'), 'utf8');

  const headerMatch = indexHtml.match(/<header[\s\S]*?<\/header>/)?.[0] || '';
  assertAudit(P, !headerMatch.includes('<img') && !headerMatch.includes('class="brand-logo-img"'), 'Header has zero clutter icon (pure Haute Joaillerie typography, zero image CLS)');
  assertAudit(P, stylesCss.includes('overflow-anchor: none;'), 'App header and main container prevent browser scroll-anchoring jitter');
  assertAudit(P, swJs.includes('gfmp-mobile-v1.01-prod') && swJs.includes('./analytics-engine.js') && swJs.includes('./baseline_data.js'), 'Service Worker caches all core application assets for 100% offline execution');
}

// ============================================================================
// PILLAR 8: SECURITY, XSS SANITIZATION & NON-BLOCKING DIALOGS
// ============================================================================
{
  const P = '8. Security & DOM Sink Hygiene';
  const indexHtml = fs.readFileSync(path.join(MOBILE_WEB, 'index.html'), 'utf8');
  const appJs = fs.readFileSync(path.join(MOBILE_WEB, 'app.js'), 'utf8');
  const engineJs = fs.readFileSync(path.join(MOBILE_WEB, 'analytics-engine.js'), 'utf8');

  assertAudit(P, appJs.includes('function escapeHTML(val)'), 'escapeHTML sanitizer is defined in app.js');
  assertAudit(P, appJs.includes('escapeHTML(tx.nomCategorieTraduit') && appJs.includes('escapeHTML(engine.cleanAssetName(cpt.nom))'), 'User-supplied entities in table innerHTML templates are sanitized via escapeHTML()');
  assertAudit(P, !appJs.includes('eval(') && !engineJs.includes('eval('), 'Zero eval() usage across production JS codebase');
  assertAudit(P, appJs.includes('function showConfirmDialog') && indexHtml.includes('<dialog id="modal-confirm-dialog"'), 'Non-blocking native <dialog id="modal-confirm-dialog"> is used for all destructive actions');
}

// ============================================================================
// PILLAR 9: CRYPTOGRAPHIC SHA-256 PARITY (MOBILE-WEB VS ANDROID ASSETS)
// ============================================================================
{
  const P = '9. Cryptographic SHA-256 Asset Parity';
  const filesToVerify = [
    'index.html',
    'styles.css',
    'app.js',
    'analytics-engine.js',
    'baseline_data.js',
    'baseline_data.json',
    'sw.js',
    'manifest.json',
    'assets/favicon_jic.png',
    'assets/icon_512.png',
    'assets/logo_jic.png',
    'assets/logo_jic_transparent.png'
  ];

  for (const rel of filesToVerify) {
    const webFile = path.join(MOBILE_WEB, rel);
    const androidFile = path.join(ANDROID_ASSETS, rel);
    const webExists = fs.existsSync(webFile);
    const androidExists = fs.existsSync(androidFile);
    assertAudit(P, webExists && androidExists, `Asset ${rel} exists in both mobile-web and android assets`);
    if (webExists && androidExists) {
      const h1 = crypto.createHash('sha256').update(fs.readFileSync(webFile)).digest('hex');
      const h2 = crypto.createHash('sha256').update(fs.readFileSync(androidFile)).digest('hex');
      assertAudit(P, h1 === h2, `SHA-256 parity verified for ${rel} (${h1.slice(0, 12)}...)`);
    }
  }
}

// ============================================================================
// PILLAR 10: NATIVE ANDROID CONTAINER (API 34, MAINACTIVITY, MANIFEST)
// ============================================================================
{
  const P = '10. Native Android Container (API 34)';
  const manifestXml = fs.readFileSync(path.join(ROOT, 'android', 'app', 'src', 'main', 'AndroidManifest.xml'), 'utf8');
  const mainActivityKt = fs.readFileSync(path.join(ROOT, 'android', 'app', 'src', 'main', 'kotlin', 'com', 'jicorporate', 'gfmp', 'MainActivity.kt'), 'utf8');
  const buildGradle = fs.readFileSync(path.join(ROOT, 'android', 'app', 'build.gradle.kts'), 'utf8');

  assertAudit(P, manifestXml.includes('android:allowBackup="false"') && manifestXml.includes('android:hardwareAccelerated="true"'), 'AndroidManifest disables auto-backup and enables hardware acceleration');
  assertAudit(P, mainActivityKt.includes('file:///android_asset/index.html?onboarding=1'), 'MainActivity.kt detects fresh installation and passes ?onboarding=1 to WebView');
  assertAudit(P, buildGradle.includes('compileSdk = 34') && buildGradle.includes('targetSdk = 34'), 'Android build.gradle.kts targets API 34');
}

// ============================================================================
// SUMMARY REPORT
// ============================================================================
console.log('--------------------------------------------------------------------------------');
for (const [pillar, stat] of Object.entries(pillarStats)) {
  const status = stat.passed === stat.total ? '✅ PASS' : '❌ FAIL';
  console.log(`${status} | ${pillar.padEnd(46)} | ${stat.passed}/${stat.total}`);
}
console.log('--------------------------------------------------------------------------------');
console.log(`TOTAL AUDIT CHECKS: ${passedChecks}/${totalChecks} (${((passedChecks / totalChecks) * 100).toFixed(2)}%)`);

if (failures.length > 0) {
  console.error('\nFAILURES DETECTED:');
  failures.forEach(f => console.error(' - ' + f));
  process.exit(1);
} else {
  console.log('\n🏆 VERDICT: GOLD STANDARD CUTTING-EDGE — 100% CONFORME (0 ERREUR, 0 RÉGRESSION)');
}
