// Integration Test for Database Multilingual Translation & Dynamic Extension
const fs = require('fs');
const path = require('path');

global.window = global;
global.localStorage = {
  _data: {},
  getItem(k) { return this._data[k] || null; },
  setItem(k, v) { this._data[k] = String(v); },
  removeItem(k) { delete this._data[k]; }
};

const baselineCode = fs.readFileSync(path.join(__dirname, '../mobile-web/baseline_data.js'), 'utf8');
eval(baselineCode);

const engineCode = fs.readFileSync(path.join(__dirname, '../mobile-web/analytics-engine.js'), 'utf8');
eval(engineCode);

const engine = window.gfmpEngine;
engine.init();

console.log('=== TEST 1: Database Quadri-Lingual Translation ===');
const sampleEntities = [
  { raw: 'Compte Courant Principal', type: 'compte' },
  { raw: 'Livret d\'Épargne', type: 'compte' },
  { raw: 'Assurance Vie', type: 'compte' },
  { raw: 'Carte de Crédit (Différé)', type: 'compte' },
  { raw: 'Salaire / Revenus Pro', type: 'categorie' },
  { raw: 'Alimentation & Supermarché', type: 'categorie' },
  { raw: 'Transports (Essence/Assurance)', type: 'categorie' },
  { raw: 'Employeur principal', type: 'tiers' },
  { raw: 'Supermarché / Alimentaire', type: 'tiers' },
  { raw: 'LIQUIDITE', type: 'classe' },
  { raw: 'INVESTISSEMENT', type: 'classe' },
  { raw: 'DETTE', type: 'classe' },
  { raw: 'Premier loyer résidence', type: 'note' },
  { raw: 'Super U Grand Baie - Alimentation', type: 'note' }
];

['FR', 'EN', 'ES', 'PT'].forEach(lang => {
  engine.setLanguage(lang);
  console.log(`\n--- Active Language: [${lang}] ---`);
  sampleEntities.forEach(item => {
    const tr = engine.trEntity(item.raw);
    console.log(`  [${item.type}] "${item.raw}" -> "${tr}"`);
    if (!tr) throw new Error(`Empty translation for ${item.raw} in ${lang}`);
  });
});

console.log('\n=== TEST 2: Views Data Payloads Contain Translations ===');
engine.setLanguage('EN');
const cf = engine.computeCashflow();
const firstTx = cf.transactions[0];
console.log('First Cashflow Tx Translated in EN:');
console.log('  Category:', firstTx.nomCategorieTraduit, '(Original:', firstTx.nomCategorie, ')');
console.log('  Account:', firstTx.nomCompteTraduit, '(Original:', firstTx.nomCompte, ')');
console.log('  Tiers:', firstTx.nomTiersTraduit, '(Original:', firstTx.nomTiers, ')');
console.log('  Notes:', firstTx.notesTraduites, '(Original:', firstTx.notes, ')');
if (!firstTx.nomCategorieTraduit || !firstTx.nomCompteTraduit) {
  throw new Error('Cashflow transaction missing translated entity fields');
}

const nw = engine.computeNetWorth();
const firstAcc = nw.accounts[0];
console.log('First Net Worth Account Translated in EN:');
console.log('  Account Name:', firstAcc.nomTraduit, '(Original:', firstAcc.nom, ')');
console.log('  Asset Class:', firstAcc.classeTraduit, '(Original:', firstAcc.classe, ')');
if (!firstAcc.nomTraduit || !firstAcc.classeTraduit) {
  throw new Error('Net worth account missing translated entity fields');
}

const bg = engine.computeBudgetZBB();
const firstEnv = bg.categories[0];
console.log('First Budget Envelope Translated in EN:');
console.log('  Envelope Name:', firstEnv.nomTraduit, '(Original:', firstEnv.nom, ')');
if (!firstEnv.nomTraduit) {
  throw new Error('Budget envelope missing translated entity fields');
}

console.log('\n=== TEST 3: Dynamic Entity Extension & Kimball Star Schema Deduplication ===');
engine.setLanguage('FR');
let notifyCount = 0;
engine.subscribe(() => { notifyCount++; });

// 1. Add Account
const newAcc1 = engine.addAccount('Coffre Fort Titres Suisses', 'INVESTISSEMENT', 'EUR');
console.log('Added new account:', newAcc1);
if (!newAcc1.isNew || newAcc1.id !== '10' || newAcc1.classe !== 'INVESTISSEMENT') {
  throw new Error('Failed to create new account');
}

// 2. Add Same Account (Deduplication Check)
const newAcc2 = engine.addAccount('coffre fort titres suisses');
console.log('Deduplication existing account returned:', newAcc2);
if (newAcc2.isNew || newAcc2.id !== '10') {
  throw new Error('Deduplication failed for account');
}

// 3. Add Category
const newCat1 = engine.addCategory('Formations & Livres Pro', 'DEPENSE');
console.log('Added new category:', newCat1);
if (!newCat1.isNew || newCat1.id !== '12') {
  throw new Error('Failed to create new category');
}

// 4. Add Same Category (Deduplication Check)
const newCat2 = engine.addCategory('formations & livres pro');
console.log('Deduplication existing category returned:', newCat2);
if (newCat2.isNew || newCat2.id !== '12') {
  throw new Error('Deduplication failed for category');
}

// 5. Add Tiers
const newTiers1 = engine.addTiers('Université Paris-Dauphine', 'INSTITUTION');
console.log('Added new tiers:', newTiers1);
if (!newTiers1.isNew || newTiers1.id !== '8') {
  throw new Error('Failed to create new tiers');
}

// 6. Add Same Tiers (Deduplication Check)
const newTiers2 = engine.addTiers('université paris-dauphine');
console.log('Deduplication existing tiers returned:', newTiers2);
if (newTiers2.isNew || newTiers2.id !== '8') {
  throw new Error('Deduplication failed for tiers');
}

// 7. Verify Audit Log
const history = engine.data.audit_history;
console.log('Audit history count:', history.length);
const recentActions = history.slice(0, 3).map(h => `${h.action}: ${h.description}`);
console.log('Recent audit actions:', recentActions);
if (history.length < 3) throw new Error('Audit history entries missing');

console.log('Notification count triggered:', notifyCount);
if (notifyCount < 3) throw new Error('Subscribers not notified upon entity creation');

console.log('\n>>> ALL DATABASE TRANSLATION & DYNAMIC EXTENSION TESTS PASSED (0 ERRORS) <<<');
