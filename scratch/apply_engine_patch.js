const fs = require('fs');

const dictCode = fs.readFileSync('scratch/build_dictionary.js', 'utf8');
const dictMatch = dictCode.match(/const dbEntityDict = (\{[\s\S]*?\n\});/);
if (!dictMatch) throw new Error('Dict not found');
const dictObjStr = dictMatch[1];

let engineSrc = fs.readFileSync('mobile-web/analytics-engine.js', 'utf8');

// 1. Add GFMP_DB_TRANSLATIONS before class GFMPEngine
const topAddition = `const GFMP_DB_TRANSLATIONS = ${dictObjStr};\n\nclass GFMPEngine {`;
engineSrc = engineSrc.replace('class GFMPEngine {', topAddition);

// 2. Add methods cleanText, trEntity, cleanAssetName after getCurrencySymbol
const helperMethods = `  cleanText(str) {
    if (!str) return '';
    return String(str).replace(/\\s+/g, ' ').trim();
  }

  trEntity(name, lang = this.lang) {
    if (!name) return '';
    const clean = this.cleanText(name);
    const targetLang = (lang && ['FR', 'EN', 'ES', 'PT'].includes(lang)) ? lang : 'FR';
    
    // Direct lookup
    if (GFMP_DB_TRANSLATIONS[clean] && GFMP_DB_TRANSLATIONS[clean][targetLang]) {
      return GFMP_DB_TRANSLATIONS[clean][targetLang];
    }
    
    // Stripped lookup without parenthesized suffixes
    const stripped = clean.replace(/\\s*\\([A-Z]{3,4}\\)/g, '')
                          .replace(/\\s*\\(Préciser\\.\\.\\.\\)/i, '')
                          .replace(/\\s*\\(Différé\\)/i, '')
                          .trim();
    if (GFMP_DB_TRANSLATIONS[stripped] && GFMP_DB_TRANSLATIONS[stripped][targetLang]) {
      return GFMP_DB_TRANSLATIONS[stripped][targetLang];
    }
    
    return clean;
  }

  cleanAssetName(name) {
    if (!name) return '';
    const cleaned = String(name).replace(/\\s*\\([A-Z]{3,4}\\)/g, '').replace(/\\s*\\(Préciser\\.\\.\\.\\)/i, '').trim();
    return this.trEntity(cleaned);
  }

  genererNouvelId(collection) {
    if (!Array.isArray(collection) || collection.length === 0) return '1';
    const maxId = collection.reduce((max, item) => {
      const num = parseInt(item.id, 10);
      return (!isNaN(num) && num > max) ? num : max;
    }, 0);
    return String(maxId + 1);
  }

  existeDeja(collection, nom) {
    if (!Array.isArray(collection) || !nom) return null;
    const target = this.cleanText(nom).toLowerCase();
    return collection.find(item => {
      const rawNom = this.cleanText(item.nom).toLowerCase();
      const trNom = this.cleanText(this.trEntity(item.nom)).toLowerCase();
      return rawNom === target || trNom === target;
    }) || null;
  }`;

// Replace original cleanAssetName
engineSrc = engineSrc.replace(/cleanAssetName\(name\) \{[\s\S]*?return String\(name\)\.replace\(\/\\s\*\\\(\[A-Z\]\{3,4\}\\\)\/g, ''\)\.trim\(\);\s*\}/, helperMethods);

// 3. Add modal keys into fallbackDict
const extraDictKeys = `      'MODAL_ADD_ENTITY_TITLE': { FR: 'NOUVEL ÉLÉMENT', EN: 'NEW ELEMENT', ES: 'NUEVO ELEMENTO', PT: 'NOVO ELEMENTO' },
      'PILL_ENTITY_COMPTE': { FR: 'Actif / Compte', EN: 'Asset / Account', ES: 'Activo / Cuenta', PT: 'Ativo / Conta' },
      'PILL_ENTITY_CATEGORIE': { FR: 'Catégorie', EN: 'Category', ES: 'Categoría', PT: 'Categoria' },
      'PILL_ENTITY_TIERS': { FR: 'Tiers / Bénéficiaire', EN: 'Third-Party / Payee', ES: 'Tercero / Beneficiario', PT: 'Terceiro / Beneficiário' },
      'LBL_ADD_ENTITY_NAME': { FR: 'Désignation / Nom *', EN: 'Name / Designation *', ES: 'Nombre / Designación *', PT: 'Nome / Designação *' },
      'LBL_ADD_ENTITY_CLASSE': { FR: 'Classe d\\'Actif', EN: 'Asset Class', ES: 'Clase de Activo', PT: 'Classe de Ativo' },
      'LBL_ADD_ENTITY_DEVISE': { FR: 'Devise du Compte', EN: 'Account Currency', ES: 'Moneda de la Cuenta', PT: 'Moeda da Conta' },
      'LBL_ADD_ENTITY_CAT_TYPE': { FR: 'Type Budgétaire', EN: 'Budgetary Type', ES: 'Tipo Presupuestario', PT: 'Tipo Orçamentário' },
      'LBL_ADD_ENTITY_TIERS_SECTEUR': { FR: 'Secteur d\\'Activité', EN: 'Industry Sector', ES: 'Sector de Actividad', PT: 'Setor de Atividade' },
      'BTN_CONFIRM_ADD_ENTITY': { FR: 'Ajouter & Sélectionner', EN: 'Add & Select', ES: 'Agregar y Seleccionar', PT: 'Adicionar e Selecionar' },
      'BTN_CANCEL': { FR: 'Annuler', EN: 'Cancel', ES: 'Cancelar', PT: 'Cancelar' },
      'AJOUTER_NOUVEAU': { FR: 'Ajouter nouveau', EN: 'Add new', ES: 'Agregar nuevo', PT: 'Adicionar novo' },
      'BTN_FIELD_ADD': { FR: '+ Nouveau', EN: '+ New', ES: '+ Nuevo', PT: '+ Novo' },
      'ENTITY_ADDED_SUCCESS': { FR: 'Élément ajouté et sélectionné avec succès !', EN: 'Item successfully added and selected!', ES: '¡Elemento agregado y seleccionado con éxito!', PT: 'Item adicionado e selecionado com sucesso!' },
      'ENTITY_ALREADY_EXISTS': { FR: 'Cet élément existe déjà et a été sélectionné.', EN: 'This item already exists and was selected.', ES: 'Este elemento ya existe y ha sido seleccionado.', PT: 'Este item já existe e foi selecionado.' },
      'ALL_ACCOUNTS': { FR: 'Tous les actifs', EN: 'All assets', ES: 'Todos los activos', PT: 'Todos os ativos' },
      'BTN_CLOSE_AUDIT': {`;

engineSrc = engineSrc.replace("'BTN_CLOSE_AUDIT': {", extraDictKeys);

// 4. Update addAccount, addCategory, addTiers
const enhancedAddMethods = `  addAccount(nom, classe = 'LIQUIDITE', devise = this.currency) {
    if (!this.data.comptes) this.data.comptes = [];
    const cleanNom = this.cleanText(nom);
    const existing = this.existeDeja(this.data.comptes, cleanNom);
    if (existing) {
      return { ...existing, isNew: false };
    }
    const nextId = this.genererNouvelId(this.data.comptes);
    const newCpt = { id: nextId, nom: cleanNom, classe, devise, actif: 'OUI' };
    this.data.comptes.push(newCpt);
    this.logHistory('ACCOUNT_ADD', 'account', nextId, \`Ajout compte \${cleanNom} (\${classe})\`, newCpt);
    this.save();
    this.notify();
    return { ...newCpt, isNew: true };
  }

  addCategory(nom, type = 'DEPENSE') {
    if (!this.data.categories) this.data.categories = [];
    const cleanNom = this.cleanText(nom);
    const existing = this.existeDeja(this.data.categories, cleanNom);
    if (existing) {
      return { ...existing, isNew: false };
    }
    const nextId = this.genererNouvelId(this.data.categories);
    const newCat = { id: nextId, nom: cleanNom, type };
    this.data.categories.push(newCat);
    this.logHistory('CATEGORY_ADD', 'category', nextId, \`Ajout catégorie \${cleanNom} (\${type})\`, newCat);
    this.save();
    this.notify();
    return { ...newCat, isNew: true };
  }

  addTiers(nom, type = 'COMMERCE') {
    if (!this.data.tiers) this.data.tiers = [];
    const cleanNom = this.cleanText(nom);
    const existing = this.existeDeja(this.data.tiers, cleanNom);
    if (existing) {
      return { ...existing, isNew: false };
    }
    const nextId = this.genererNouvelId(this.data.tiers);
    const newTiers = { id: nextId, nom: cleanNom, type };
    this.data.tiers.push(newTiers);
    this.logHistory('TIERS_ADD', 'tiers', nextId, \`Ajout tiers \${cleanNom} (\${type})\`, newTiers);
    this.save();
    this.notify();
    return { ...newTiers, isNew: true };
  }`;

const oldAddMethodsRegex = /addAccount\(nom, classe = 'LIQUIDITE', devise = 'MUR'\) \{[\s\S]*?return newTiers;\s*\}/;
engineSrc = engineSrc.replace(oldAddMethodsRegex, enhancedAddMethods);

// 5. In computeCashflow, add translated entity fields
const oldCashflowMapping = `        // Multi-level Filters
        const nomCompte = cptDict[tx.id_compte] || 'Compte';
        const nomCat = catDict[tx.id_categorie] || 'Catégorie';
        const nomTiers = tiersDict[tx.id_tiers] || 'Tiers';
        const fullSearchStr = \`\${tx.date} \${nomCompte} \${nomCat} \${nomTiers} \${tx.notes || ''}\`.toLowerCase();

        const matchSearch = (!this.searchTerm) || fullSearchStr.includes(this.searchTerm.toLowerCase());
        const matchFlux = (this.fluxFilter === 'ALL') || (typeFlux === this.fluxFilter);
        const matchCatInteractive = (!this.selectedCategoryFilter) || (nomCat === this.selectedCategoryFilter);
        const matchAccount = (this.accountFilter === 'ALL') || (String(tx.id_compte) === String(this.accountFilter));

        if (matchSearch && matchFlux && matchCatInteractive && matchAccount) {
          filteredTx.push({
            ...tx,
            nomCompte,
            nomCategorie: nomCat,
            nomTiers,
            typeFlux,
            montantConverti: converted
          });
        }`;

const newCashflowMapping = `        // Multi-level Filters & Translation
        const nomCompte = cptDict[tx.id_compte] || 'Compte';
        const nomCat = catDict[tx.id_categorie] || 'Catégorie';
        const nomTiers = tiersDict[tx.id_tiers] || 'Tiers';
        const nomCompteTraduit = this.trEntity(nomCompte);
        const nomCategorieTraduit = this.trEntity(nomCat);
        const nomTiersTraduit = this.trEntity(nomTiers);
        const notesTraduites = this.trEntity(tx.notes);
        const fullSearchStr = \`\${tx.date} \${nomCompte} \${nomCat} \${nomTiers} \${tx.notes || ''} \${nomCompteTraduit} \${nomCategorieTraduit} \${nomTiersTraduit} \${notesTraduites}\`.toLowerCase();

        const matchSearch = (!this.searchTerm) || fullSearchStr.includes(this.searchTerm.toLowerCase());
        const matchFlux = (this.fluxFilter === 'ALL') || (typeFlux === this.fluxFilter);
        const matchCatInteractive = (!this.selectedCategoryFilter) || (nomCat === this.selectedCategoryFilter || nomCategorieTraduit === this.selectedCategoryFilter);
        const matchAccount = (this.accountFilter === 'ALL') || (String(tx.id_compte) === String(this.accountFilter));

        if (matchSearch && matchFlux && matchCatInteractive && matchAccount) {
          filteredTx.push({
            ...tx,
            nomCompte,
            nomCategorie: nomCat,
            nomTiers,
            nomCompteTraduit,
            nomCategorieTraduit,
            nomTiersTraduit,
            notesTraduites,
            typeFlux,
            montantConverti: converted
          });
        }`;

engineSrc = engineSrc.replace(oldCashflowMapping, newCashflowMapping);

// Write to scratch first to test
fs.writeFileSync('scratch/patched_engine.js', engineSrc, 'utf8');
console.log('Patched engine written to scratch/patched_engine.js');
