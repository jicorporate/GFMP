const fs = require('fs');

global.window = global;
global.localStorage = {
  _data: {},
  getItem(k) { return this._data[k] || null; },
  setItem(k, v) { this._data[k] = String(v); },
  removeItem(k) { delete this._data[k]; }
};

const baselineCode = fs.readFileSync('mobile-web/baseline_data.js', 'utf8');
eval(baselineCode);

const patchedCode = fs.readFileSync('scratch/patched_engine.js', 'utf8');
eval(patchedCode);

const engine = window.gfmpEngine;
engine.init();

console.log('=== TEST 1: Initialized ===');
console.log('Lang:', engine.lang, 'Currency:', engine.currency);

console.log('=== TEST 2: Account Clean & Translate ===');
engine.setLanguage('EN');
console.log('EN Account 1:', engine.cleanAssetName('Compte Courant Principal'));
console.log('EN Account 2:', engine.cleanAssetName('Livret d\'Épargne'));
console.log('EN Account 4:', engine.cleanAssetName('Assurance Vie'));
console.log('EN Category 5:', engine.trEntity('Alimentation & Supermarché'));
console.log('EN Tiers 4:', engine.trEntity('Supermarché / Alimentaire'));
console.log('EN Note:', engine.trEntity('Premier loyer résidence'));
console.log('EN Class:', engine.trEntity('LIQUIDITE'));

engine.setLanguage('ES');
console.log('ES Category 5:', engine.trEntity('Alimentation & Supermarché'));
console.log('ES Account 1:', engine.cleanAssetName('Compte Courant Principal'));

engine.setLanguage('PT');
console.log('PT Category 5:', engine.trEntity('Alimentation & Supermarché'));
console.log('PT Account 1:', engine.cleanAssetName('Compte Courant Principal'));

console.log('=== TEST 3: Kimball Star Schema Auto-Learning & Deduplication ===');
engine.setLanguage('FR');
const beforeCount = engine.data.comptes.length;
console.log('Comptes count before:', beforeCount);

// Adding new account
const added1 = engine.addAccount('Livret Développement Durable', 'LIQUIDITE', 'EUR');
console.log('Added 1:', added1);
if (!added1.isNew || added1.id !== '10') throw new Error('Failed to add new account');

// Adding same account again (deduplication check)
const added2 = engine.addAccount('livret développement durable');
console.log('Added 2 (dedup):', added2);
if (added2.isNew || added2.id !== '10') throw new Error('Deduplication failed for account');

// Adding new category
const cat1 = engine.addCategory('Abonnements SaaS & Tech', 'DEPENSE');
console.log('Cat 1:', cat1);
if (!cat1.isNew) throw new Error('Failed to add new category');

// Adding existing category check
const cat2 = engine.addCategory('Abonnements SaaS & Tech');
console.log('Cat 2 (dedup):', cat2);
if (cat2.isNew) throw new Error('Deduplication failed for category');

// Adding new tiers
const tiers1 = engine.addTiers('GitHub / Microsoft', 'ENTREPRISE');
console.log('Tiers 1:', tiers1);
if (!tiers1.isNew) throw new Error('Failed to add new tiers');

const tiers2 = engine.addTiers('github / microsoft');
console.log('Tiers 2 (dedup):', tiers2);
if (tiers2.isNew) throw new Error('Deduplication failed for tiers');

console.log('>>> PATCHED ENGINE VALIDATION PASSED COMPLETELY! <<<');
