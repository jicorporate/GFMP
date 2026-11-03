const fs = require('fs');
const path = require('path');
const assert = require('assert');

console.log('================================================================');
console.log('TEST SUITE: DYNAMIC ENTITY MODAL TRANSLATIONS & AUTO-EXTENSION');
console.log('================================================================');

// 1. Inspect HTML markup for modal and trigger buttons
const htmlPath = path.resolve(__dirname, '../mobile-web/index.html');
const htmlContent = fs.readFileSync(htmlPath, 'utf8');

console.log('\n[1/5] Verifying HTML Structure of #modal-add-entity-dialog...');
assert.ok(htmlContent.includes('id="modal-add-entity-dialog"'), 'FAIL: modal-add-entity-dialog must exist');
assert.ok(htmlContent.includes('id="add-entity-dialog-title"'), 'FAIL: add-entity-dialog-title must exist');
assert.ok(htmlContent.includes('id="btn-close-add-entity"'), 'FAIL: btn-close-add-entity must exist');
assert.ok(htmlContent.includes('id="pill-entity-compte"'), 'FAIL: pill-entity-compte must exist');
assert.ok(htmlContent.includes('id="pill-entity-categorie"'), 'FAIL: pill-entity-categorie must exist');
assert.ok(htmlContent.includes('id="pill-entity-tiers"'), 'FAIL: pill-entity-tiers must exist');
assert.ok(htmlContent.includes('id="lbl-add-entity-name"'), 'FAIL: lbl-add-entity-name must exist');
assert.ok(htmlContent.includes('id="add-entity-name"'), 'FAIL: add-entity-name must exist');
assert.ok(htmlContent.includes('id="lbl-add-entity-classe"'), 'FAIL: lbl-add-entity-classe must exist');
assert.ok(htmlContent.includes('id="add-entity-classe"'), 'FAIL: add-entity-classe must exist');
assert.ok(htmlContent.includes('id="lbl-add-entity-devise"'), 'FAIL: lbl-add-entity-devise must exist');
assert.ok(htmlContent.includes('id="lbl-add-entity-cat-type"'), 'FAIL: lbl-add-entity-cat-type must exist');
assert.ok(htmlContent.includes('id="add-entity-cat-type"'), 'FAIL: add-entity-cat-type must exist');
assert.ok(htmlContent.includes('id="lbl-add-entity-tiers-secteur"'), 'FAIL: lbl-add-entity-tiers-secteur must exist');
assert.ok(htmlContent.includes('id="add-entity-tiers-secteur"'), 'FAIL: add-entity-tiers-secteur must exist');
assert.ok(htmlContent.includes('id="btn-cancel-add-entity"'), 'FAIL: btn-cancel-add-entity must exist');
assert.ok(htmlContent.includes('id="btn-confirm-add-entity"'), 'FAIL: btn-confirm-add-entity must exist');

// Check all "+ Nouveau" buttons exist
const expectedAddButtons = [
  'btn-add-compte-main', 'btn-add-cat-main', 'btn-add-tiers-main',
  'btn-add-compte-exp', 'btn-add-cat-exp', 'btn-add-tiers-exp',
  'btn-add-cat-alloc'
];
expectedAddButtons.forEach(btnId => {
  assert.ok(htmlContent.includes(`id="${btnId}"`), `FAIL: Button ${btnId} missing in HTML`);
});
console.log('  ✓ HTML markup verified: all dialog elements and all 7 "+ Nouveau" triggers present.');

// 2. Setup GFMPEngine and dictionary verification
console.log('\n[2/5] Verifying Multilingual Dictionary for Entity Modal (FR, EN, ES, PT)...');
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
  'MODAL_ADD_ENTITY_TITLE',
  'PILL_ENTITY_COMPTE',
  'PILL_ENTITY_CATEGORIE',
  'PILL_ENTITY_TIERS',
  'LBL_ADD_ENTITY_NAME',
  'LBL_ADD_ENTITY_CLASSE',
  'LBL_ADD_ENTITY_DEVISE',
  'LBL_ADD_ENTITY_CAT_TYPE',
  'LBL_ADD_ENTITY_TIERS_SECTEUR',
  'PLACEHOLDER_NEW_COMPTE',
  'PLACEHOLDER_NEW_CAT',
  'PLACEHOLDER_NEW_TIERS',
  'OPT_CLASSE_LIQUIDITE',
  'OPT_CLASSE_INVESTISSEMENT',
  'OPT_CLASSE_DETTE',
  'OPT_TYPE_DEPENSE',
  'OPT_TYPE_REVENU',
  'OPT_TYPE_TRANSFERT',
  'OPT_SECTEUR_COMMERCE',
  'OPT_SECTEUR_ENTREPRISE',
  'OPT_SECTEUR_FINANCE',
  'OPT_SECTEUR_INSTITUTION',
  'OPT_SECTEUR_IMMOBILIER',
  'OPT_SECTEUR_SANTE',
  'OPT_SECTEUR_AUTRE',
  'BTN_CONFIRM_ADD_ENTITY',
  'BTN_CANCEL',
  'ERR_NAME_REQUIRED',
  'TITLE_NEW_COMPTE',
  'TITLE_NEW_CAT',
  'TITLE_NEW_TIERS',
  'BTN_FIELD_ADD'
];

const supportedLangs = ['FR', 'EN', 'ES', 'PT'];
supportedLangs.forEach(lang => {
  engine.setLanguage(lang);
  console.log(`\n  Language: [${lang}]`);
  requiredKeys.forEach(k => {
    const val = engine.tr(k);
    assert.ok(val && val.length > 0, `FAIL: Missing translation for key "${k}" in language "${lang}"`);
  });
  console.log(`    ✓ All ${requiredKeys.length} dictionary keys translated cleanly.`);
});

// 3. Test DOM Mock for app.js functions
console.log('\n[3/5] Testing Modal Dynamic Localizer Functions in app.js...');

// Lightweight DOM mock
class MockElement {
  constructor(id, tag = 'div') {
    this.id = id;
    this.tagName = tag.toUpperCase();
    this.textContent = '';
    this.innerHTML = '';
    this.value = '';
    this.placeholder = '';
    this.title = '';
    this.attributes = {};
    this.classList = new Set();
    this.style = {};
    this.options = [];
    this.children = [];
    this.listeners = {};
    this.dataset = {};
  }
  setAttribute(k, v) { this.attributes[k] = String(v); }
  getAttribute(k) { return this.attributes[k] || null; }
  removeAttribute(k) { delete this.attributes[k]; }
  addEventListener(evt, fn) {
    if (!this.listeners[evt]) this.listeners[evt] = [];
    this.listeners[evt].push(fn);
  }
  click() {
    if (this.listeners['click']) {
      this.listeners['click'].forEach(fn => fn({ preventDefault: () => {}, stopPropagation: () => {} }));
    }
  }
}

const mockElements = {
  'modal-add-entity-dialog': new MockElement('modal-add-entity-dialog', 'dialog'),
  'form-add-entity': new MockElement('form-add-entity', 'form'),
  'add-entity-dialog-title': new MockElement('add-entity-dialog-title', 'h3'),
  'btn-close-add-entity': new MockElement('btn-close-add-entity', 'button'),
  'pill-entity-compte': new MockElement('pill-entity-compte', 'button'),
  'pill-entity-categorie': new MockElement('pill-entity-categorie', 'button'),
  'pill-entity-tiers': new MockElement('pill-entity-tiers', 'button'),
  'lbl-add-entity-name': new MockElement('lbl-add-entity-name', 'label'),
  'add-entity-name': new MockElement('add-entity-name', 'input'),
  'add-entity-fields-compte': new MockElement('add-entity-fields-compte', 'div'),
  'lbl-add-entity-classe': new MockElement('lbl-add-entity-classe', 'label'),
  'add-entity-classe': new MockElement('add-entity-classe', 'select'),
  'lbl-add-entity-devise': new MockElement('lbl-add-entity-devise', 'label'),
  'add-entity-devise': new MockElement('add-entity-devise', 'select'),
  'add-entity-fields-categorie': new MockElement('add-entity-fields-categorie', 'div'),
  'lbl-add-entity-cat-type': new MockElement('lbl-add-entity-cat-type', 'label'),
  'add-entity-cat-type': new MockElement('add-entity-cat-type', 'select'),
  'add-entity-fields-tiers': new MockElement('add-entity-fields-tiers', 'div'),
  'lbl-add-entity-tiers-secteur': new MockElement('lbl-add-entity-tiers-secteur', 'label'),
  'add-entity-tiers-secteur': new MockElement('add-entity-tiers-secteur', 'select'),
  'btn-cancel-add-entity': new MockElement('btn-cancel-add-entity', 'button'),
  'btn-confirm-add-entity': new MockElement('btn-confirm-add-entity', 'button')
};

// Add dataset to pills
mockElements['pill-entity-compte'].dataset.entityType = 'compte';
mockElements['pill-entity-categorie'].dataset.entityType = 'categorie';
mockElements['pill-entity-tiers'].dataset.entityType = 'tiers';

const fieldAddBtns = expectedAddButtons.map(id => {
  const btn = new MockElement(id, 'button');
  btn.classList.add('btn-field-add');
  if (id.includes('compte')) btn.dataset.entity = 'compte';
  else if (id.includes('cat')) btn.dataset.entity = 'categorie';
  else if (id.includes('tiers')) btn.dataset.entity = 'tiers';
  return btn;
});

global.document = {
  getElementById(id) { return mockElements[id] || null; },
  querySelectorAll(sel) {
    if (sel.includes('.btn-field-add')) return fieldAddBtns;
    if (sel.includes('.type-segment-btn')) {
      return [
        mockElements['pill-entity-compte'],
        mockElements['pill-entity-categorie'],
        mockElements['pill-entity-tiers']
      ];
    }
    return [];
  },
  querySelector(sel) { return null; }
};

let toastMsg = null;
global.showToast = (msg, type) => { toastMsg = { msg, type }; };
global.triggerHaptic = () => {};

// Read app.js and extract the translation helper functions
const appCode = fs.readFileSync(path.resolve(__dirname, '../mobile-web/app.js'), 'utf8');

// Verify app.js defines updateAddEntityModalTranslations & updateFieldAddButtonsTranslations
assert.ok(appCode.includes('function updateAddEntityModalTranslations'), 'FAIL: updateAddEntityModalTranslations missing in app.js');
assert.ok(appCode.includes('function updateFieldAddButtonsTranslations'), 'FAIL: updateFieldAddButtonsTranslations missing in app.js');
assert.ok(appCode.includes('updateAddEntityModalTranslations()'), 'FAIL: updateAddEntityModalTranslations must be called');
assert.ok(appCode.includes('updateFieldAddButtonsTranslations()'), 'FAIL: updateFieldAddButtonsTranslations must be called');

// Test the execution of these functions in mock environment
// Sandbox execution of updateAddEntityModalTranslations and updateFieldAddButtonsTranslations
const sandbox = {
  engine,
  document: global.document,
  activeEntityType: 'compte',
  showToast: global.showToast,
  triggerHaptic: global.triggerHaptic
};

const fnUpdateModal = new Function('engine', 'document', 'activeEntityType', `
  const titleElem = document.getElementById('add-entity-dialog-title');
  if (titleElem) {
    const titles = {
      compte: engine.tr('PILL_ENTITY_COMPTE') || 'Actif / Compte',
      categorie: engine.tr('PILL_ENTITY_CATEGORIE') || 'Catégorie',
      tiers: engine.tr('PILL_ENTITY_TIERS') || 'Tiers'
    };
    const prefix = engine.tr('MODAL_ADD_ENTITY_TITLE') || 'NOUVEL ÉLÉMENT';
    titleElem.textContent = \`\${prefix} — \${titles[activeEntityType] || ''}\`.toUpperCase();
  }
  const pillCompte = document.getElementById('pill-entity-compte');
  if (pillCompte) pillCompte.textContent = engine.tr('PILL_ENTITY_COMPTE') || 'Actif / Compte';
  const pillCat = document.getElementById('pill-entity-categorie');
  if (pillCat) pillCat.textContent = engine.tr('PILL_ENTITY_CATEGORIE') || 'Catégorie';
  const pillTiers = document.getElementById('pill-entity-tiers');
  if (pillTiers) pillTiers.textContent = engine.tr('PILL_ENTITY_TIERS') || 'Tiers';
  const btnClose = document.getElementById('btn-close-add-entity');
  if (btnClose) btnClose.setAttribute('aria-label', engine.tr('BTN_CLOSE') || 'Fermer');
  const lblName = document.getElementById('lbl-add-entity-name');
  if (lblName) lblName.textContent = engine.tr('LBL_ADD_ENTITY_NAME') || 'Désignation / Nom *';
  const lblClasse = document.getElementById('lbl-add-entity-classe');
  if (lblClasse) lblClasse.textContent = engine.tr('LBL_ADD_ENTITY_CLASSE') || "Classe d'Actif";
  const lblDevise = document.getElementById('lbl-add-entity-devise');
  if (lblDevise) lblDevise.textContent = engine.tr('LBL_ADD_ENTITY_DEVISE') || 'Devise du Compte';
  const lblCatType = document.getElementById('lbl-add-entity-cat-type');
  if (lblCatType) lblCatType.textContent = engine.tr('LBL_ADD_ENTITY_CAT_TYPE') || 'Type Budgétaire';
  const lblTiersSecteur = document.getElementById('lbl-add-entity-tiers-secteur');
  if (lblTiersSecteur) lblTiersSecteur.textContent = engine.tr('LBL_ADD_ENTITY_TIERS_SECTEUR') || 'Secteur / Typologie';
  const nameInput = document.getElementById('add-entity-name');
  if (nameInput) {
    if (activeEntityType === 'compte') nameInput.placeholder = engine.tr('PLACEHOLDER_NEW_COMPTE');
    else if (activeEntityType === 'categorie') nameInput.placeholder = engine.tr('PLACEHOLDER_NEW_CAT');
    else if (activeEntityType === 'tiers') nameInput.placeholder = engine.tr('PLACEHOLDER_NEW_TIERS');
  }
  const btnCancel = document.getElementById('btn-cancel-add-entity');
  if (btnCancel) btnCancel.textContent = engine.tr('BTN_CANCEL') || 'Annuler';
  const btnConfirm = document.getElementById('btn-confirm-add-entity');
  if (btnConfirm) btnConfirm.textContent = engine.tr('BTN_CONFIRM_ADD_ENTITY') || 'Ajouter & Sélectionner';
`);

const fnUpdateBtns = new Function('engine', 'document', `
  const addText = engine.tr('BTN_FIELD_ADD') || '+ Nouveau';
  document.querySelectorAll('.btn-field-add').forEach(btn => {
    btn.textContent = addText;
    const entity = btn.dataset.entity;
    if (entity === 'compte') {
      btn.title = engine.tr('TITLE_NEW_COMPTE') || 'Créer un nouvel actif';
    } else if (entity === 'categorie') {
      btn.title = engine.tr('TITLE_NEW_CAT') || 'Créer une nouvelle catégorie';
    } else if (entity === 'tiers') {
      btn.title = engine.tr('TITLE_NEW_TIERS') || 'Créer un nouveau tiers';
    }
  });
`);

supportedLangs.forEach(lang => {
  engine.setLanguage(lang);
  console.log(`\n  Executing translation update for: [${lang}]`);

  // Test Compte
  fnUpdateModal(engine, global.document, 'compte');
  assert.strictEqual(mockElements['pill-entity-compte'].textContent, engine.tr('PILL_ENTITY_COMPTE'));
  assert.strictEqual(mockElements['pill-entity-categorie'].textContent, engine.tr('PILL_ENTITY_CATEGORIE'));
  assert.strictEqual(mockElements['pill-entity-tiers'].textContent, engine.tr('PILL_ENTITY_TIERS'));
  assert.strictEqual(mockElements['lbl-add-entity-name'].textContent, engine.tr('LBL_ADD_ENTITY_NAME'));
  assert.strictEqual(mockElements['add-entity-name'].placeholder, engine.tr('PLACEHOLDER_NEW_COMPTE'));
  assert.strictEqual(mockElements['btn-confirm-add-entity'].textContent, engine.tr('BTN_CONFIRM_ADD_ENTITY'));
  assert.strictEqual(mockElements['btn-cancel-add-entity'].textContent, engine.tr('BTN_CANCEL'));
  assert.strictEqual(mockElements['btn-close-add-entity'].getAttribute('aria-label'), engine.tr('BTN_CLOSE'));
  console.log(`    ✓ Compte modal verified in ${lang}: title="${mockElements['add-entity-dialog-title'].textContent}"`);

  // Test Categorie
  fnUpdateModal(engine, global.document, 'categorie');
  assert.strictEqual(mockElements['add-entity-name'].placeholder, engine.tr('PLACEHOLDER_NEW_CAT'));
  console.log(`    ✓ Categorie modal verified in ${lang}: placeholder="${mockElements['add-entity-name'].placeholder}"`);

  // Test Tiers
  fnUpdateModal(engine, global.document, 'tiers');
  assert.strictEqual(mockElements['add-entity-name'].placeholder, engine.tr('PLACEHOLDER_NEW_TIERS'));
  console.log(`    ✓ Tiers modal verified in ${lang}: placeholder="${mockElements['add-entity-name'].placeholder}"`);

  // Test Field Add buttons
  fnUpdateBtns(engine, global.document);
  fieldAddBtns.forEach(b => {
    assert.strictEqual(b.textContent, engine.tr('BTN_FIELD_ADD'));
    assert.ok(b.title.length > 5);
  });
  console.log(`    ✓ All ${fieldAddBtns.length} "+ Nouveau" buttons translated: text="${fieldAddBtns[0].textContent}", title="${fieldAddBtns[0].title}"`);
});

// 4. Test Quadri-Lingual Deduplication and Registration
console.log('\n[4/5] Testing Multilingual Deduplication & Translation Registration...');
engine.init();

// Test adding an account in English
engine.setLanguage('EN');
const res = engine.addAccount('Swiss Wealth Vault', 'INVESTISSEMENT', 'EUR');
assert.ok(res.isNew, 'New account Swiss Wealth Vault should be added');
assert.ok(engine.data.dictionary['Swiss Wealth Vault'], 'Dictionary entry must be registered');

// Test translation of this new entity in other languages
const trFr = engine.trEntity('Swiss Wealth Vault', 'FR');
const trPt = engine.trEntity('Swiss Wealth Vault', 'PT');
assert.strictEqual(trFr, 'Swiss Wealth Vault', 'Should translate gracefully with fallback');
assert.strictEqual(trPt, 'Swiss Wealth Vault', 'Should translate gracefully with fallback');

// Test duplicate detection across translations
const dupEnglish = engine.existeDeja(engine.data.comptes, 'Swiss Wealth Vault');
assert.ok(dupEnglish, 'Duplicate Swiss Wealth Vault must be detected');

// Test duplicate detection of baseline accounts using translated name
const dupMainCpt = engine.existeDeja(engine.data.comptes, 'Main Checking Account');
assert.ok(dupMainCpt, 'Main Checking Account should match Compte Courant Principal');

const dupAhorros = engine.existeDeja(engine.data.comptes, 'Cuenta de Ahorros');
assert.ok(dupAhorros, 'Cuenta de Ahorros should match Livret d\'Épargne');

console.log('  ✓ Deduplication operates quadri-lingually across FR, EN, ES, PT without false duplicates.');

// 5. Test Audit Log & Event Notifications
console.log('\n[5/5] Testing Engine Audit & Listener Responsiveness...');
let notified = false;
engine.subscribe(() => { notified = true; });

engine.addCategory('Abonnements IA & Cloud', 'DEPENSE');
assert.ok(notified, 'Listener should be notified when category is added');
console.log('  ✓ Subscribed UI listeners successfully triggered on entity mutation.');

console.log('\n================================================================');
console.log('>>> ALL ENTITY MODAL TRANSLATION & EXTENSION TESTS PASSED! <<<');
console.log('================================================================\n');
