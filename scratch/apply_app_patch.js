const fs = require('fs');

let appSrc = fs.readFileSync('mobile-web/app.js', 'utf8');

// 1. Replace filterCategories & populateSelects
const oldPopulate = `  function filterCategories(selectId, targetType) {
    const catSel = document.getElementById(selectId);
    if (!catSel || !engine.data) return;

    const filtered = (engine.data.categories || []).filter(c => {
      if (targetType === 'ALL') return true;
      return c.type === targetType;
    });

    catSel.innerHTML = filtered.map(c => \`
      <option value="\${c.id}">\${c.nom}</option>
    \`).join('');
  }

  function populateSelects() {
    if (!engine.data) return;

    const fillSelects = (prefix) => {
      const cptSel = document.getElementById(\`\${prefix}-compte\`);
      const tiersSel = document.getElementById(\`\${prefix}-tiers\`);
      const devSel = document.getElementById(\`\${prefix}-devise\`);

      if (cptSel) {
        cptSel.innerHTML = (engine.data.comptes || []).map(c => \`
          <option value="\${c.id}">\${engine.cleanAssetName(c.nom)}</option>
        \`).join('');
      }

      if (tiersSel) {
        tiersSel.innerHTML = (engine.data.tiers || []).map(t => \`
          <option value="\${t.id}">\${t.nom}</option>
        \`).join('');
      }

      if (devSel) {
        const keys = Object.keys(engine.data.devises || { MUR: 1, EUR: 1, USD: 1 });
        devSel.innerHTML = keys.map(k => \`
          <option value="\${k}" \${k === engine.currency ? 'selected' : ''}>\${k}</option>
        \`).join('');
      }
    };

    fillSelects('tx');
    fillSelects('exp-tx');

    // Filter categories based on active type selection
    filterCategories('tx-categorie', activeTxType);
    filterCategories('exp-tx-categorie', activeExpType);

    // Cashflow account filter dropdown
    const cfAccSel = document.getElementById('cf-account-filter');
    if (cfAccSel) {
      cfAccSel.innerHTML = '<option value="ALL">Tous les actifs</option>' + (engine.data.comptes || []).map(c => \`
        <option value="\${c.id}" \${String(engine.accountFilter) === String(c.id) ? 'selected' : ''}>\${engine.cleanAssetName(c.nom)}</option>
      \`).join('');
    }

    // Budget Alloc Category Select
    const allocCat = document.getElementById('alloc-categorie');
    if (allocCat) {
      const deps = (engine.data.categories || []).filter(c => c.type === 'DEPENSE');
      allocCat.innerHTML = deps.map(c => \`
        <option value="\${c.id}">\${c.nom}</option>
      \`).join('');
    }

    // Set default dates to today
    const todayStr = new Date().toISOString().split('T')[0];
    const txDate = document.getElementById('tx-date');
    if (txDate && !txDate.value) txDate.value = todayStr;
    const expDate = document.getElementById('exp-tx-date');
    if (expDate && !expDate.value) expDate.value = todayStr;
  }`;

const newPopulate = `  function filterCategories(selectId, targetType) {
    const catSel = document.getElementById(selectId);
    if (!catSel || !engine.data) return;

    const filtered = (engine.data.categories || []).filter(c => {
      if (targetType === 'ALL') return true;
      return c.type === targetType;
    });

    const addOptionText = engine.tr('AJOUTER_NOUVEAU') || 'Ajouter nouveau';
    catSel.innerHTML = filtered.map(c => \`
      <option value="\${c.id}">\${engine.trEntity(c.nom)}</option>
    \`).join('') + \`<option value="__NEW__" style="color:var(--c-yellow-royal); font-weight:800;">+ \${addOptionText}...</option>\`;
  }

  function populateSelects() {
    if (!engine.data) return;

    const addOptionText = engine.tr('AJOUTER_NOUVEAU') || 'Ajouter nouveau';

    const fillSelects = (prefix) => {
      const cptSel = document.getElementById(\`\${prefix}-compte\`);
      const tiersSel = document.getElementById(\`\${prefix}-tiers\`);
      const devSel = document.getElementById(\`\${prefix}-devise\`);

      if (cptSel) {
        cptSel.innerHTML = (engine.data.comptes || []).map(c => \`
          <option value="\${c.id}">\${engine.cleanAssetName(c.nom)}</option>
        \`).join('') + \`<option value="__NEW__" style="color:var(--c-yellow-royal); font-weight:800;">+ \${addOptionText}...</option>\`;
      }

      if (tiersSel) {
        tiersSel.innerHTML = (engine.data.tiers || []).map(t => \`
          <option value="\${t.id}">\${engine.trEntity(t.nom)}</option>
        \`).join('') + \`<option value="__NEW__" style="color:var(--c-yellow-royal); font-weight:800;">+ \${addOptionText}...</option>\`;
      }

      if (devSel) {
        const keys = Object.keys(engine.data.devises || { MUR: 1, EUR: 1, USD: 1 });
        devSel.innerHTML = keys.map(k => \`
          <option value="\${k}" \${k === engine.currency ? 'selected' : ''}>\${k}</option>
        \`).join('');
      }
    };

    fillSelects('tx');
    fillSelects('exp-tx');

    // Filter categories based on active type selection
    filterCategories('tx-categorie', activeTxType);
    filterCategories('exp-tx-categorie', activeExpType);

    // Cashflow account filter dropdown
    const cfAccSel = document.getElementById('cf-account-filter');
    if (cfAccSel) {
      const allAccText = engine.tr('ALL_ACCOUNTS') || 'Tous les actifs';
      cfAccSel.innerHTML = \`<option value="ALL">\${allAccText}</option>\` + (engine.data.comptes || []).map(c => \`
        <option value="\${c.id}" \${String(engine.accountFilter) === String(c.id) ? 'selected' : ''}>\${engine.cleanAssetName(c.nom)}</option>
      \`).join('');
    }

    // Budget Alloc Category Select
    const allocCat = document.getElementById('alloc-categorie');
    if (allocCat) {
      const deps = (engine.data.categories || []).filter(c => c.type === 'DEPENSE');
      allocCat.innerHTML = deps.map(c => \`
        <option value="\${c.id}">\${engine.trEntity(c.nom)}</option>
      \`).join('') + \`<option value="__NEW__" style="color:var(--c-yellow-royal); font-weight:800;">+ \${addOptionText}...</option>\`;
    }

    // Set default dates to today
    const todayStr = new Date().toISOString().split('T')[0];
    const txDate = document.getElementById('tx-date');
    if (txDate && !txDate.value) txDate.value = todayStr;
    const expDate = document.getElementById('exp-tx-date');
    if (expDate && !expDate.value) expDate.value = todayStr;
  }`;

appSrc = appSrc.replace(oldPopulate, newPopulate);

// 2. Add Dynamic Entity Dialog Controller before section 9
const entityDialogCode = `  // --- 8B. DYNAMIC ENTITY EXTENSION GATEKEEPER DIALOG (GOLD STANDARD) ---
  const addEntityDialog = document.getElementById('modal-add-entity-dialog');
  const addEntityForm = document.getElementById('form-add-entity');
  let currentTargetSelect = null;
  let activeEntityType = 'compte'; // 'compte', 'categorie', 'tiers'

  function switchEntityType(type) {
    activeEntityType = type;
    document.querySelectorAll('#add-entity-type-selector .type-segment-btn').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.entityType === type);
    });

    const fieldsCompte = document.getElementById('add-entity-fields-compte');
    const fieldsCat = document.getElementById('add-entity-fields-categorie');
    const fieldsTiers = document.getElementById('add-entity-fields-tiers');

    if (fieldsCompte) fieldsCompte.style.display = (type === 'compte') ? 'flex' : 'none';
    if (fieldsCat) fieldsCat.style.display = (type === 'categorie') ? 'flex' : 'none';
    if (fieldsTiers) fieldsTiers.style.display = (type === 'tiers') ? 'flex' : 'none';

    // Update modal title
    const titleElem = document.getElementById('add-entity-dialog-title');
    if (titleElem) {
      const titles = {
        compte: engine.tr('PILL_ENTITY_COMPTE') || 'Actif / Compte',
        categorie: engine.tr('PILL_ENTITY_CATEGORIE') || 'Catégorie',
        tiers: engine.tr('PILL_ENTITY_TIERS') || 'Tiers / Bénéficiaire'
      };
      const prefix = engine.tr('MODAL_ADD_ENTITY_TITLE') || 'NOUVEL ÉLÉMENT';
      titleElem.textContent = \`\${prefix} — \${titles[type] || ''}\`.toUpperCase();
    }
  }

  function openAddEntityDialog(type = 'compte', targetSelect = null) {
    currentTargetSelect = targetSelect;
    triggerHaptic(15);
    switchEntityType(type);

    const nameInput = document.getElementById('add-entity-name');
    if (nameInput) {
      nameInput.value = '';
    }

    if (addEntityDialog) {
      if (typeof addEntityDialog.showModal === 'function') {
        addEntityDialog.showModal();
      } else {
        addEntityDialog.setAttribute('open', '');
      }
      setTimeout(() => nameInput?.focus(), 50);
    }
  }

  function closeAddEntityDialog() {
    if (addEntityDialog) {
      if (typeof addEntityDialog.close === 'function') {
        addEntityDialog.close();
      } else {
        addEntityDialog.removeAttribute('open');
      }
    }
    // If the target select was left on __NEW__, reset to first option
    if (currentTargetSelect && currentTargetSelect.value === '__NEW__') {
      const firstOpt = currentTargetSelect.options[0];
      if (firstOpt) currentTargetSelect.value = firstOpt.value;
    }
    currentTargetSelect = null;
  }

  // Type selector pills inside dialog
  document.querySelectorAll('#add-entity-type-selector .type-segment-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      triggerHaptic(15);
      switchEntityType(btn.dataset.entityType);
    });
  });

  // Close buttons
  document.getElementById('btn-close-add-entity')?.addEventListener('click', closeAddEntityDialog);
  document.getElementById('btn-cancel-add-entity')?.addEventListener('click', closeAddEntityDialog);

  // Hook all "+ Nouveau" buttons in forms
  document.querySelectorAll('.btn-field-add').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      const entity = btn.dataset.entity; // 'compte', 'categorie', 'tiers'
      let targetSel = null;
      if (btn.id.includes('main')) {
        targetSel = document.getElementById(\`tx-\${entity}\`);
      } else if (btn.id.includes('exp')) {
        targetSel = document.getElementById(\`exp-tx-\${entity}\`);
      } else if (btn.id.includes('alloc')) {
        targetSel = document.getElementById('alloc-categorie');
      }
      openAddEntityDialog(entity, targetSel);
    });
  });

  // Hook change on selects for __NEW__ or "Autre (Préciser...)"
  const hookSelectNew = (selectId, entityType) => {
    const sel = document.getElementById(selectId);
    if (!sel) return;
    sel.addEventListener('change', () => {
      const optText = sel.options[sel.selectedIndex]?.text || '';
      if (sel.value === '__NEW__' || optText.includes('(Préciser...)')) {
        openAddEntityDialog(entityType, sel);
      }
    });
  };

  hookSelectNew('tx-compte', 'compte');
  hookSelectNew('tx-categorie', 'categorie');
  hookSelectNew('tx-tiers', 'tiers');
  hookSelectNew('exp-tx-compte', 'compte');
  hookSelectNew('exp-tx-categorie', 'categorie');
  hookSelectNew('exp-tx-tiers', 'tiers');
  hookSelectNew('alloc-categorie', 'categorie');

  // Submit new entity
  addEntityForm?.addEventListener('submit', (e) => {
    e.preventDefault();
    triggerHaptic(25);

    const nameVal = document.getElementById('add-entity-name')?.value?.trim();
    if (!nameVal) {
      showToast('Veuillez renseigner un nom valide', 'error');
      return;
    }

    let result = null;
    if (activeEntityType === 'compte') {
      const classe = document.getElementById('add-entity-classe')?.value || 'LIQUIDITE';
      const devise = document.getElementById('add-entity-devise')?.value || engine.currency;
      result = engine.addAccount(nameVal, classe, devise);
    } else if (activeEntityType === 'categorie') {
      const type = document.getElementById('add-entity-cat-type')?.value || 'DEPENSE';
      result = engine.addCategory(nameVal, type);
    } else if (activeEntityType === 'tiers') {
      const secteur = document.getElementById('add-entity-tiers-secteur')?.value || 'COMMERCE';
      result = engine.addTiers(nameVal, secteur);
    }

    if (result) {
      populateSelects();
      if (currentTargetSelect) {
        currentTargetSelect.value = String(result.id);
      }
      const msg = result.isNew ? (engine.tr('ENTITY_ADDED_SUCCESS') || 'Élément ajouté et sélectionné avec succès !') : (engine.tr('ENTITY_ALREADY_EXISTS') || 'Cet élément existe déjà et a été sélectionné.');
      showToast(msg, 'success');
    }

    closeAddEntityDialog();
  });

  // --- 9. FULL-SCREEN IMMERSIVE MODAL (SAISIE EXPRESS TRANSACTION) ---`;

appSrc = appSrc.replace('  // --- 9. FULL-SCREEN IMMERSIVE MODAL (SAISIE EXPRESS TRANSACTION) ---', entityDialogCode);

// 3. Update Form Submits with validation against __NEW__
const oldExpressSubmit = `    const dateVal = document.getElementById('exp-tx-date').value;
    const cptVal = document.getElementById('exp-tx-compte').value;
    const catVal = document.getElementById('exp-tx-categorie').value;
    const tiersVal = document.getElementById('exp-tx-tiers').value;
    const montantVal = document.getElementById('exp-tx-montant').value;
    const devVal = document.getElementById('exp-tx-devise').value;
    const notesVal = document.getElementById('exp-tx-notes').value;

    if (!montantVal || parseFloat(montantVal) <= 0) {
      showToast('Veuillez saisir un montant valide', 'error');
      return;
    }`;

const newExpressSubmit = `    const dateVal = document.getElementById('exp-tx-date').value;
    const cptVal = document.getElementById('exp-tx-compte').value;
    const catVal = document.getElementById('exp-tx-categorie').value;
    const tiersVal = document.getElementById('exp-tx-tiers').value;
    const montantVal = document.getElementById('exp-tx-montant').value;
    const devVal = document.getElementById('exp-tx-devise').value;
    const notesVal = document.getElementById('exp-tx-notes').value;

    if (!montantVal || parseFloat(montantVal) <= 0) {
      showToast('Veuillez saisir un montant valide', 'error');
      return;
    }

    if (cptVal === '__NEW__' || !cptVal) {
      showToast('Veuillez sélectionner ou créer un actif', 'error');
      openAddEntityDialog('compte', document.getElementById('exp-tx-compte'));
      return;
    }
    if (catVal === '__NEW__' || !catVal) {
      showToast('Veuillez sélectionner ou créer une catégorie', 'error');
      openAddEntityDialog('categorie', document.getElementById('exp-tx-categorie'));
      return;
    }
    if (tiersVal === '__NEW__' || !tiersVal) {
      showToast('Veuillez sélectionner ou créer un tiers', 'error');
      openAddEntityDialog('tiers', document.getElementById('exp-tx-tiers'));
      return;
    }`;

appSrc = appSrc.replace(oldExpressSubmit, newExpressSubmit);

const oldMainSubmit = `    const dateVal = document.getElementById('tx-date').value;
    let cptVal = document.getElementById('tx-compte').value;
    let catVal = document.getElementById('tx-categorie').value;
    let tiersVal = document.getElementById('tx-tiers').value;
    const montantVal = document.getElementById('tx-montant').value;
    const devVal = document.getElementById('tx-devise').value;
    const notesVal = document.getElementById('tx-notes').value;

    if (!montantVal || parseFloat(montantVal) <= 0) {
      showToast('Veuillez saisir un montant valide', 'error');
      return;
    }

    // Self-learning fields
    const newCpt = document.getElementById('tx-new-compte')?.value?.trim();
    if (newCpt) cptVal = engine.addAccount(newCpt).id;
    const newCat = document.getElementById('tx-new-categorie')?.value?.trim();
    if (newCat) catVal = engine.addCategory(newCat).id;
    const newTiers = document.getElementById('tx-new-tiers')?.value?.trim();
    if (newTiers) tiersVal = engine.addTiers(newTiers).id;`;

const newMainSubmit = `    const dateVal = document.getElementById('tx-date').value;
    let cptVal = document.getElementById('tx-compte').value;
    let catVal = document.getElementById('tx-categorie').value;
    let tiersVal = document.getElementById('tx-tiers').value;
    const montantVal = document.getElementById('tx-montant').value;
    const devVal = document.getElementById('tx-devise').value;
    const notesVal = document.getElementById('tx-notes').value;

    if (!montantVal || parseFloat(montantVal) <= 0) {
      showToast('Veuillez saisir un montant valide', 'error');
      return;
    }

    if (cptVal === '__NEW__' || !cptVal) {
      showToast('Veuillez sélectionner ou créer un actif', 'error');
      openAddEntityDialog('compte', document.getElementById('tx-compte'));
      return;
    }
    if (catVal === '__NEW__' || !catVal) {
      showToast('Veuillez sélectionner ou créer une catégorie', 'error');
      openAddEntityDialog('categorie', document.getElementById('tx-categorie'));
      return;
    }
    if (tiersVal === '__NEW__' || !tiersVal) {
      showToast('Veuillez sélectionner ou créer un tiers', 'error');
      openAddEntityDialog('tiers', document.getElementById('tx-tiers'));
      return;
    }`;

appSrc = appSrc.replace(oldMainSubmit, newMainSubmit);

const oldAllocSubmit = `    const mois = document.getElementById('alloc-mois').value;
    const cat = document.getElementById('alloc-categorie').value;
    const amt = document.getElementById('alloc-montant').value;

    engine.allocateBudget(cat, amt, mois);`;

const newAllocSubmit = `    const mois = document.getElementById('alloc-mois').value;
    const cat = document.getElementById('alloc-categorie').value;
    const amt = document.getElementById('alloc-montant').value;

    if (cat === '__NEW__' || !cat) {
      showToast('Veuillez sélectionner ou créer une catégorie', 'error');
      openAddEntityDialog('categorie', document.getElementById('alloc-categorie'));
      return;
    }

    engine.allocateBudget(cat, amt, mois);`;

appSrc = appSrc.replace(oldAllocSubmit, newAllocSubmit);

// 4. Update Financial Report Modal top expenses translation
appSrc = appSrc.replace(
  `          <span>\${e.nom}</span>`,
  `          <span>\${engine.trEntity(e.nom)}</span>`
);
appSrc = appSrc.replace(
  `\${e.nom}: \${e.montant}`,
  `\${engine.trEntity(e.nom)}: \${e.montant}`
);

// 5. Update Donut Chart translation in legend & title
appSrc = appSrc.replace(
  `<title>\${item.nom}: \${engine.formatMoney(item.montant)} (\${Math.round(item.pct)}%) - Cliquer pour filtrer</title>`,
  `<title>\${engine.trEntity(item.nom)}: \${engine.formatMoney(item.montant)} (\${Math.round(item.pct)}%) - Cliquer pour filtrer</title>`
);
appSrc = appSrc.replace(
  `<span style="\${isSelected ? 'color:#FADA5E; font-weight:800;' : ''}">\${item.nom}</span>`,
  `<span style="\${isSelected ? 'color:#FADA5E; font-weight:800;' : ''}">\${engine.trEntity(item.nom)}</span>`
);

// 6. Update Hub Central top assets, budget, and recent transactions
appSrc = appSrc.replace(
  `<span class="badge-classe \${ast.classe}" style="font-size:0.62rem; padding:2px 6px;">\${ast.classe}</span>
              <span style="font-size:0.82rem; font-weight:700; color:#FFF;">\${ast.nom}</span>`,
  `<span class="badge-classe \${ast.classe}" style="font-size:0.62rem; padding:2px 6px;">\${engine.trEntity(ast.classe)}</span>
              <span style="font-size:0.82rem; font-weight:700; color:#FFF;">\${engine.trEntity(ast.nom)}</span>`
);
appSrc = appSrc.replace(
  `<span style="font-size:0.8rem; font-weight:700; color:#FFF;">\${bg.nom}</span>`,
  `<span style="font-size:0.8rem; font-weight:700; color:#FFF;">\${engine.trEntity(bg.nom)}</span>`
);
appSrc = appSrc.replace(
  `<div style="font-weight:700; font-size:0.8rem; color:#FFF;">\${tx.categorie}</div>
                <div style="font-size:0.68rem; color:var(--text-secondary);">\${tx.compte}</div>`,
  `<div style="font-weight:700; font-size:0.8rem; color:#FFF;">\${engine.trEntity(tx.categorie)}</div>
                <div style="font-size:0.68rem; color:var(--text-secondary);">\${engine.cleanAssetName(tx.compte)}</div>`
);

// 7. Update Cashflow View table row translations
const oldCashflowRow = `        <td>
          <div style="font-weight:700;">\${tx.nomCategorie}</div>
          <div style="font-size:0.7rem; opacity:0.8;">\${engine.cleanAssetName(tx.nomCompte)}</div>
        </td>
        <td>
          <div>\${tx.nomTiers}</div>
          <span class="badge-flux \${tx.typeFlux}">\${tx.typeFlux}</span>
        </td>`;

const newCashflowRow = `        <td>
          <div style="font-weight:700;">\${tx.nomCategorieTraduit || engine.trEntity(tx.nomCategorie)}</div>
          <div style="font-size:0.7rem; opacity:0.8;">\${tx.nomCompteTraduit || engine.cleanAssetName(tx.nomCompte)}</div>
        </td>
        <td>
          <div>\${tx.nomTiersTraduit || engine.trEntity(tx.nomTiers)}</div>
          <span class="badge-flux \${tx.typeFlux}">\${engine.trEntity(tx.typeFlux)}</span>
          \${tx.notes ? \`<div style="font-size:0.68rem; color:var(--text-muted); font-style:italic;">\${tx.notesTraduites || engine.trEntity(tx.notes)}</div>\` : ''}
        </td>`;

appSrc = appSrc.replace(oldCashflowRow, newCashflowRow);

// 8. Update Net Worth View row translations
const oldNwRow = `        <td>
          <div style="font-weight:700;">\${engine.cleanAssetName(cpt.nom)}</div>
          <div style="font-size:0.68rem; opacity:0.8;">\${cpt.isDette ? engine.tr('ASSET_LIABILITY') : engine.tr('ASSET_ACTIVE')}</div>
        </td>
        <td>
          <span class="badge-flux \${cpt.classe}">\${cpt.classe}</span>
        </td>`;

const newNwRow = `        <td>
          <div style="font-weight:700;">\${cpt.nomTraduit || engine.cleanAssetName(cpt.nom)}</div>
          <div style="font-size:0.68rem; opacity:0.8;">\${cpt.isDette ? engine.tr('ASSET_LIABILITY') : engine.tr('ASSET_ACTIVE')}</div>
        </td>
        <td>
          <span class="badge-flux \${cpt.classe}">\${cpt.classeTraduit || engine.trEntity(cpt.classe)}</span>
        </td>`;

appSrc = appSrc.replace(oldNwRow, newNwRow);

// 9. Update Budget View row translations
appSrc = appSrc.replace(
  `<div style="font-weight:700;">\${c.nom}</div>`,
  `<div style="font-weight:700;">\${c.nomTraduit || engine.trEntity(c.nom)}</div>`
);

// 10. Update renderCurrentView to translate .btn-field-add and dialog labels
const additionInRenderCurrentView = `    // 13. Dynamic Entity Dialog & Quick Add Buttons Translations
    document.querySelectorAll('.btn-field-add').forEach(btn => {
      btn.textContent = engine.tr('BTN_FIELD_ADD') || '+ Nouveau';
    });
    const addEntityTitle = document.getElementById('add-entity-dialog-title');
    if (addEntityTitle) addEntityTitle.textContent = engine.tr('MODAL_ADD_ENTITY_TITLE') || 'NOUVEL ÉLÉMENT';
    const pillCompte = document.getElementById('pill-entity-compte');
    if (pillCompte) pillCompte.textContent = engine.tr('PILL_ENTITY_COMPTE') || 'Actif / Compte';
    const pillCat = document.getElementById('pill-entity-categorie');
    if (pillCat) pillCat.textContent = engine.tr('PILL_ENTITY_CATEGORIE') || 'Catégorie';
    const pillTiers = document.getElementById('pill-entity-tiers');
    if (pillTiers) pillTiers.textContent = engine.tr('PILL_ENTITY_TIERS') || 'Tiers';
    const lblAddName = document.getElementById('lbl-add-entity-name');
    if (lblAddName) lblAddName.textContent = engine.tr('LBL_ADD_ENTITY_NAME') || 'Désignation / Nom *';
    const lblAddClasse = document.getElementById('lbl-add-entity-classe');
    if (lblAddClasse) lblAddClasse.textContent = engine.tr('LBL_ADD_ENTITY_CLASSE') || 'Classe d\\'Actif';
    const lblAddDev = document.getElementById('lbl-add-entity-devise');
    if (lblAddDev) lblAddDev.textContent = engine.tr('LBL_ADD_ENTITY_DEVISE') || 'Devise du Compte';
    const lblAddCatType = document.getElementById('lbl-add-entity-cat-type');
    if (lblAddCatType) lblAddCatType.textContent = engine.tr('LBL_ADD_ENTITY_CAT_TYPE') || 'Type Budgétaire';
    const lblAddTiersSec = document.getElementById('lbl-add-entity-tiers-secteur');
    if (lblAddTiersSec) lblAddTiersSec.textContent = engine.tr('LBL_ADD_ENTITY_TIERS_SECTEUR') || 'Secteur / Typologie';
    const btnConfirmAdd = document.getElementById('btn-confirm-add-entity');
    if (btnConfirmAdd) btnConfirmAdd.textContent = engine.tr('BTN_CONFIRM_ADD_ENTITY') || 'Ajouter & Sélectionner';
    const btnCancelAdd = document.getElementById('btn-cancel-add-entity');
    if (btnCancelAdd) btnCancelAdd.textContent = engine.tr('BTN_CANCEL') || 'Annuler';

    // 14. Dispatch to active view renderer`;

appSrc = appSrc.replace('    // 13. Dispatch to active view renderer', additionInRenderCurrentView);

fs.writeFileSync('scratch/patched_app.js', appSrc, 'utf8');
console.log('Patched app written to scratch/patched_app.js');
