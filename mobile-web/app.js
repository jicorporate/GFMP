/**
 * GFMP — Global Finance Management Platform (SFP v1.0)
 * Ultra-Modern Mobile Application Controller & DOM Binder
 * Cutting-Edge Display, Interactivity, Real-Time Calculations & Offline Storage
 */

document.addEventListener('DOMContentLoaded', () => {
  const engine = window.gfmpEngine;
  engine.init();

  // --- 1. TACTILE HAPTIC FEEDBACK ---
  function triggerHaptic(duration = 15) {
    if (localStorage.getItem('gfmp_haptic') === 'false') return;
    if (navigator.vibrate) {
      try {
        navigator.vibrate(duration);
      } catch (e) {
        // Ignored on unsupported devices
      }
    }
  }

  // --- 1.1 SECURITY HELPER: HTML ENTITY ESCAPING FOR DYNAMIC DOM SINKS ---
  function escapeHTML(val) {
    if (val === null || val === undefined) return '';
    return String(val)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  // --- 2. FLOATING TOAST NOTIFICATION ---
  function showToast(message, type = 'success') {
    let toast = document.getElementById('gfmp-toast');
    if (!toast) {
      toast = document.createElement('div');
      toast.id = 'gfmp-toast';
      toast.setAttribute('role', 'status');
      toast.setAttribute('aria-live', 'polite');
      toast.setAttribute('aria-atomic', 'true');
      toast.style.cssText = `
        position: fixed;
        top: 20px;
        left: 50%;
        transform: translateX(-50%);
        background: #10B981;
        color: #FFFFFF;
        padding: 10px 20px;
        border-radius: 20px;
        font-size: 0.82rem;
        font-weight: 700;
        box-shadow: 0 10px 30px rgba(0,0,0,0.5);
        z-index: 9999;
        transition: opacity 0.3s ease, transform 0.3s ease;
        pointer-events: none;
        opacity: 0;
      `;
      document.body.appendChild(toast);
    }
    toast.style.background = (type === 'error') ? '#EF4444' : (type === 'info') ? '#4169E1' : '#10B981';
    toast.textContent = message;
    toast.style.opacity = '1';
    toast.style.transform = 'translateX(-50%) translateY(0)';

    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateX(-50%) translateY(-10px)';
    }, 2500);
  }

  // --- 2.1 MODERN ASYNC GOLD CONFIRMATION DIALOG (NON-BLOCKING) ---
  function showConfirmDialog({
    title = 'Confirmation',
    message = 'Êtes-vous sûr de vouloir continuer ?',
    confirmText = 'Confirmer',
    cancelText = 'Annuler',
    type = 'warning' // 'warning' | 'danger' | 'info'
  } = {}) {
    return new Promise((resolve) => {
      const dialog = document.getElementById('modal-confirm-dialog');
      if (!dialog || typeof dialog.showModal !== 'function') {
        resolve(window.confirm(`${title}\n\n${message}`));
        return;
      }

      const titleEl = document.getElementById('confirm-dialog-title');
      const msgEl = document.getElementById('confirm-dialog-message');
      const btnConfirm = document.getElementById('confirm-dialog-btn-confirm');
      const btnCancel = document.getElementById('confirm-dialog-btn-cancel');
      const iconContainer = document.getElementById('confirm-dialog-icon');

      if (titleEl) titleEl.textContent = title;
      if (msgEl) msgEl.textContent = message;
      if (btnConfirm) {
        btnConfirm.textContent = confirmText;
        btnConfirm.className = `btn-confirm-action ${type}`;
      }
      if (btnCancel) btnCancel.textContent = cancelText;

      if (iconContainer) {
        iconContainer.className = `confirm-dialog-icon ${type}`;
        if (type === 'danger') {
          iconContainer.innerHTML = '<svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="#EF4444" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>';
        } else if (type === 'info') {
          iconContainer.innerHTML = '<svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="#4169E1" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/></svg>';
        } else {
          iconContainer.innerHTML = '<svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="#FADA5E" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>';
        }
      }

      let settled = false;
      const settle = (val) => {
        if (settled) return;
        settled = true;
        cleanup();
        if (dialog.open) dialog.close();
        resolve(val);
      };

      const onConfirm = () => settle(true);
      const onCancel = () => settle(false);
      const onClose = () => settle(dialog.returnValue === 'confirm');

      function cleanup() {
        btnConfirm?.removeEventListener('click', onConfirm);
        btnCancel?.removeEventListener('click', onCancel);
        dialog.removeEventListener('close', onClose);
      }

      btnConfirm?.addEventListener('click', onConfirm);
      btnCancel?.addEventListener('click', onCancel);
      dialog.addEventListener('close', onClose, { once: true });

      // Fallback for browsers without native closedby support
      if (!('closedBy' in HTMLDialogElement.prototype) && !dialog._hasLightDismissBound) {
        dialog._hasLightDismissBound = true;
        dialog.addEventListener('click', (event) => {
          if (event.target !== dialog) return;
          const rect = dialog.getBoundingClientRect();
          const isDialogContent = (
            rect.top <= event.clientY &&
            event.clientY <= rect.top + rect.height &&
            rect.left <= event.clientX &&
            event.clientX <= rect.left + rect.width
          );
          if (!isDialogContent) dialog.close('cancel');
        });
      }

      triggerHaptic(20);
      try {
        dialog.showModal();
      } catch (e) {
        resolve(window.confirm(`${title}\n\n${message}`));
      }
    });
  }

  // --- 2.2 UNIFIED AUDIT HISTORY MODAL ---
  function openAuditHistoryModal() {
    triggerHaptic(20);
    const modalAudit = document.getElementById('modal-audit-history');
    const timelineEl = document.getElementById('history-modal-timeline');
    if (timelineEl) {
      const logs = engine.getHistory();
      if (!logs || logs.length === 0) {
        timelineEl.innerHTML = '<div style="text-align:center; padding:40px; color:var(--text-muted); font-size:0.85rem;">Aucun événement d\'audit enregistré.</div>';
      } else {
        timelineEl.innerHTML = logs.map(log => `
          <div class="history-timeline-item">
            <div class="history-header">
              <span class="history-badge ${escapeHTML(log.action)}">${escapeHTML(log.action)}</span>
              <span class="history-time">${log.timestamp ? escapeHTML(log.timestamp.replace('T', ' ').substring(0, 19)) : ''}</span>
            </div>
            <div class="history-desc">${escapeHTML(log.description)}</div>
            <div style="font-size:0.68rem; color:var(--text-muted);">${escapeHTML(log.entity)} ${log.entityId ? `#${escapeHTML(log.entityId)}` : ''}</div>
          </div>
        `).join('');
      }
    }
    modalAudit?.classList.add('open');
  }

  // --- 3. VIEW ROUTING SYSTEM ---
  function navigateTo(viewName) {
    document.querySelectorAll('.view-panel').forEach(p => p.style.display = 'none');
    const target = document.getElementById(`view-${viewName}`);
    if (target) {
      target.style.display = 'block';
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    if (viewName === 'transaction' && !txDeviseUserModified) {
      const cptId = document.getElementById('tx-compte')?.value;
      const account = (engine.data?.comptes || []).find(c => String(c.id) === String(cptId));
      const devSel = document.getElementById('tx-devise');
      if (devSel) {
        devSel.value = (account && account.devise) ? account.devise : engine.currency;
      }
    }

    // Update bottom navigation bar
    document.querySelectorAll('.bottom-nav .nav-tab').forEach(t => {
      if (t.dataset.view === viewName) {
        t.classList.add('active');
        t.setAttribute('aria-current', 'page');
      } else {
        t.classList.remove('active');
        t.removeAttribute('aria-current');
      }
    });

    engine.currentView = viewName;
    renderCurrentView();
  }

  // Bind Bottom Nav Tabs
  document.querySelectorAll('.bottom-nav .nav-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      triggerHaptic(12);
      navigateTo(tab.dataset.view);
    });
  });

  // Bind Hub Quick Transaction Button
  document.getElementById('btn-hub-quick-saisie')?.addEventListener('click', (e) => {
    e.preventDefault();
    e.stopPropagation();
    openExpressModal();
  });

  // Bind Live Rates Refresh Button (Icon Only, Gold Cutting-Edge & Multilingual)
  document.getElementById('btn-refresh-rates')?.addEventListener('click', async () => {
    triggerHaptic(20);
    const btn = document.getElementById('btn-refresh-rates');
    btn?.classList.add('loading');
    showToast(engine.tr('RATES_UPDATING'), 'info');
    const res = await engine.fetchLiveExchangeRates(true);
    btn?.classList.remove('loading');
    if (res.success) {
      showToast(engine.tr('RATES_UPDATED_SUCCESS'), 'success');
    } else {
      showToast(`${engine.tr('RATES_OFFLINE')} (${res.error || 'Cache'})`, 'info');
    }
    updateLiveRatesDisplay();
  });

  function updateLiveRatesDisplay() {
    const status = engine.getExchangeRateStatus();
    const btn = document.getElementById('btn-refresh-rates');
    const dot = document.getElementById('live-pulse-dot');
    const isOnline = (typeof navigator === 'undefined' || navigator.onLine);

    if (dot) {
      dot.className = (isOnline && status.isLive) ? 'live-pulse-dot' : 'live-pulse-dot offline';
    }

    if (btn) {
      const baseTitle = engine.tr('BTN_REFRESH_RATES_TITLE');
      if (status.lastSync) {
        const time = new Date(status.lastSync).toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' });
        const lastSyncLabel = engine.tr('RATES_LAST_SYNC') || 'Sync';
        const fullTitle = `${baseTitle} • ${lastSyncLabel}: ${time}`;
        btn.setAttribute('title', fullTitle);
        btn.setAttribute('aria-label', fullTitle);
      } else {
        btn.setAttribute('title', baseTitle);
        btn.setAttribute('aria-label', baseTitle);
      }
    }
  }

  if (typeof window !== 'undefined') {
    window.addEventListener('online', () => {
      showToast('Connexion Internet rétablie', 'success');
      engine.fetchLiveExchangeRates(false);
      updateLiveRatesDisplay();
    });

    window.addEventListener('offline', () => {
      showToast('Passage en mode hors-ligne', 'info');
      updateLiveRatesDisplay();
    });
  }

  // Bind Back Buttons (< Hub Central)
  document.querySelectorAll('.btn-nav-home').forEach(btn => {
    btn.addEventListener('click', () => {
      triggerHaptic(12);
      navigateTo('home');
    });
  });


  // --- 4. TIME SLIDER & MONTH STRIP CONTROLS ---
  document.getElementById('btn-prev-year')?.addEventListener('click', () => { triggerHaptic(15); engine.prevYear(); });
  document.getElementById('btn-next-year')?.addEventListener('click', () => { triggerHaptic(15); engine.nextYear(); });
  
  // Interactive Month Strip Chips
  document.querySelectorAll('#header-month-strip .month-chip').forEach(chip => {
    chip.addEventListener('click', () => {
      triggerHaptic(15);
      engine.setMonth(chip.dataset.month);
      showToast(`Période : ${engine.getMonthLabel()} ${engine.year}`, 'info');
    });
  });

  // --- 5. MULTI-CURRENCY O(1) CONTROLS ---
  document.querySelectorAll('.curr-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      triggerHaptic(20);
      engine.setCurrency(btn.dataset.curr);
      showToast(`Devise active : ${btn.dataset.curr}`, 'info');
    });
  });

  // --- 6. LANGUAGE CONTROLS ---
  document.querySelectorAll('.lang-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      triggerHaptic(15);
      engine.setLanguage(btn.dataset.lang);
      showToast(`Langue : ${btn.dataset.lang}`, 'info');
    });
  });

  // --- 7. POPULATE FORM DROPDOWNS & FILTER BY TYPE ---
  let activeTxType = 'DEPENSE';
  let activeExpType = 'DEPENSE';
  let txDeviseUserModified = false;
  let expTxDeviseUserModified = false;
  let allocDeviseUserModified = false;

  function filterCategories(selectId, targetType) {
    const catSel = document.getElementById(selectId);
    if (!catSel || !engine.data) return;

    const prev = catSel.value;
    const filtered = (engine.data.categories || []).filter(c => {
      if (targetType === 'ALL') return true;
      return c.type === targetType;
    });

    const optAddNewText = `+ ${engine.tr('OPT_ADD_NEW') || 'Ajouter nouveau...'}`;
    catSel.innerHTML = filtered.map(c => `
      <option value="${c.id}">${engine.trEntity(c.nom)}</option>
    `).join('') + `<option value="__NEW__" style="color:var(--c-yellow-royal); font-weight:700;">${optAddNewText}</option>`;

    if (prev && prev !== '__NEW__' && filtered.some(c => String(c.id) === String(prev))) {
      catSel.value = prev;
    }
  }

  function populateSelects() {
    if (!engine.data) return;

    const optAddNewText = `+ ${engine.tr('OPT_ADD_NEW') || 'Ajouter nouveau...'}`;

    const fillSelects = (prefix) => {
      const cptSel = document.getElementById(`${prefix}-compte`);
      const tiersSel = document.getElementById(`${prefix}-tiers`);
      const devSel = document.getElementById(`${prefix}-devise`);

      if (cptSel) {
        const prevCpt = cptSel.value;
        cptSel.innerHTML = (engine.data.comptes || []).map(c => `
          <option value="${c.id}">${engine.cleanAssetName(c.nom)}</option>
        `).join('') + `<option value="__NEW__" style="color:var(--c-yellow-royal); font-weight:700;">${optAddNewText}</option>`;
        if (prevCpt && prevCpt !== '__NEW__' && (engine.data.comptes || []).some(c => String(c.id) === String(prevCpt))) {
          cptSel.value = prevCpt;
        }
      }

      if (tiersSel) {
        const prevTiers = tiersSel.value;
        tiersSel.innerHTML = (engine.data.tiers || []).map(t => `
          <option value="${t.id}">${engine.trEntity(t.nom)}</option>
        `).join('') + `<option value="__NEW__" style="color:var(--c-yellow-royal); font-weight:700;">${optAddNewText}</option>`;
        if (prevTiers && prevTiers !== '__NEW__' && (engine.data.tiers || []).some(t => String(t.id) === String(prevTiers))) {
          tiersSel.value = prevTiers;
        }
      }

      if (devSel) {
        const isModified = (prefix === 'tx') ? txDeviseUserModified : (prefix === 'exp-tx') ? expTxDeviseUserModified : false;
        const selectedCpt = cptSel ? (engine.data.comptes || []).find(c => String(c.id) === String(cptSel.value)) : null;
        const defaultDev = (selectedCpt && selectedCpt.devise) ? selectedCpt.devise : (engine.currency || 'MUR');
        const activeDev = (isModified && devSel.value) ? devSel.value : defaultDev;
        const keys = Object.keys(engine.data.devises || { MUR: 1, EUR: 1, USD: 1, GBP: 1, ZAR: 1, XOF: 1 });
        devSel.innerHTML = keys.map(k => `
          <option value="${k}" ${k === activeDev ? 'selected' : ''}>${k}</option>
        `).join('');
        devSel.value = activeDev;
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
      const prevAcc = cfAccSel.value || String(engine.accountFilter);
      const allText = engine.tr('FILTER_ALL_ASSETS') || 'Tous les actifs';
      cfAccSel.innerHTML = `<option value="ALL">${allText}</option>` + (engine.data.comptes || []).map(c => `
        <option value="${c.id}" ${String(prevAcc) === String(c.id) ? 'selected' : ''}>${engine.cleanAssetName(c.nom)}</option>
      `).join('');
    }

    // Budget Alloc Devise Select
    const allocDevSel = document.getElementById('alloc-devise');
    if (allocDevSel) {
      const activeAllocDev = (allocDeviseUserModified && allocDevSel.value) ? allocDevSel.value : (engine.currency || 'MUR');
      const keys = Object.keys(engine.data.devises || { MUR: 1, EUR: 1, USD: 1, GBP: 1, ZAR: 1, XOF: 1 });
      allocDevSel.innerHTML = keys.map(k => `
        <option value="${k}" ${k === activeAllocDev ? 'selected' : ''}>${k}</option>
      `).join('');
      allocDevSel.value = activeAllocDev;
    }

    // Budget Alloc Category Select
    const allocCat = document.getElementById('alloc-categorie');
    if (allocCat) {
      const prevAlloc = allocCat.value;
      const deps = (engine.data.categories || []).filter(c => c.type === 'DEPENSE');
      allocCat.innerHTML = deps.map(c => `
        <option value="${c.id}">${engine.trEntity(c.nom)}</option>
      `).join('') + `<option value="__NEW__" style="color:var(--c-yellow-royal); font-weight:700;">${optAddNewText}</option>`;
      if (prevAlloc && prevAlloc !== '__NEW__' && deps.some(c => String(c.id) === String(prevAlloc))) {
        allocCat.value = prevAlloc;
      }
    }

    // Set default dates to system reference date or today
    const todayStr = engine.referenceDate || new Date().toISOString().split('T')[0];
    const txDate = document.getElementById('tx-date');
    if (txDate && !txDate.value) txDate.value = todayStr;
    const expDate = document.getElementById('exp-tx-date');
    if (expDate && !expDate.value) expDate.value = todayStr;
  }

  // --- 8. TYPE SEGMENTED CONTROLS & QUICK AMOUNT CHIPS ---
  // Main form segmented type selector
  document.querySelectorAll('#tx-type-selector .type-segment-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      triggerHaptic(15);
      document.querySelectorAll('#tx-type-selector .type-segment-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      activeTxType = btn.dataset.type;
      filterCategories('tx-categorie', activeTxType);
    });
  });

  // Express modal segmented type selector
  document.querySelectorAll('#exp-type-selector .type-segment-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      triggerHaptic(15);
      document.querySelectorAll('#exp-type-selector .type-segment-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      activeExpType = btn.dataset.type;
      filterCategories('exp-tx-categorie', activeExpType);
    });
  });

  // Quick Amount Chips on Main Form
  document.querySelectorAll('.quick-amount-chip:not(.exp-chip)').forEach(chip => {
    chip.addEventListener('click', () => {
      triggerHaptic(15);
      const addVal = parseFloat(chip.dataset.add) || 0;
      const input = document.getElementById('tx-montant');
      const cur = parseFloat(input.value) || 0;
      input.value = (cur + addVal).toFixed(2);
    });
  });

  // Quick Amount Chips on Express Modal
  document.querySelectorAll('.exp-chip').forEach(chip => {
    chip.addEventListener('click', () => {
      triggerHaptic(15);
      const addVal = parseFloat(chip.dataset.add) || 0;
      const input = document.getElementById('exp-tx-montant');
      const cur = parseFloat(input.value) || 0;
      input.value = (cur + addVal).toFixed(2);
    });
  });

  // Quick Amount Chips on Budget Allocation Modal
  document.querySelectorAll('.alloc-chip').forEach(chip => {
    chip.addEventListener('click', () => {
      triggerHaptic(15);
      const addVal = parseFloat(chip.dataset.add) || 0;
      const input = document.getElementById('alloc-montant');
      const cur = parseFloat(input.value) || 0;
      input.value = (cur + addVal).toFixed(2);
    });
  });

  // --- 8B. DYNAMIC ENTITY EXTENSION GATEKEEPER DIALOG (GOLD STANDARD) ---
  const addEntityDialog = document.getElementById('modal-add-entity-dialog');
  const addEntityForm = document.getElementById('form-add-entity');
  let currentTargetSelect = null;
  let activeEntityType = 'compte'; // 'compte', 'categorie', 'tiers'

  function updateAddEntityModalTranslations() {
    // 1. Update Title
    const titleElem = document.getElementById('add-entity-dialog-title');
    if (titleElem) {
      const titles = {
        compte: engine.tr('PILL_ENTITY_COMPTE') || 'Actif / Compte',
        categorie: engine.tr('PILL_ENTITY_CATEGORIE') || 'Catégorie',
        tiers: engine.tr('PILL_ENTITY_TIERS') || 'Tiers'
      };
      const prefix = engine.tr('MODAL_ADD_ENTITY_TITLE') || 'NOUVEL ÉLÉMENT';
      titleElem.textContent = `${prefix} — ${titles[activeEntityType] || ''}`.toUpperCase();
    }

    // 2. Segmented Pill Buttons
    const pillCompte = document.getElementById('pill-entity-compte');
    if (pillCompte) pillCompte.textContent = engine.tr('PILL_ENTITY_COMPTE') || 'Actif / Compte';
    const pillCat = document.getElementById('pill-entity-categorie');
    if (pillCat) pillCat.textContent = engine.tr('PILL_ENTITY_CATEGORIE') || 'Catégorie';
    const pillTiers = document.getElementById('pill-entity-tiers');
    if (pillTiers) pillTiers.textContent = engine.tr('PILL_ENTITY_TIERS') || 'Tiers';

    // 3. Close Button
    const btnClose = document.getElementById('btn-close-add-entity');
    if (btnClose) btnClose.setAttribute('aria-label', engine.tr('BTN_CLOSE') || 'Fermer');

    // 4. Form Labels
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

    // 5. Dynamic Contextual Placeholder for Name Input
    const nameInput = document.getElementById('add-entity-name');
    if (nameInput) {
      if (activeEntityType === 'compte') {
        nameInput.placeholder = engine.tr('PLACEHOLDER_NEW_COMPTE') || 'Ex: Livret A, Compte Titres...';
      } else if (activeEntityType === 'categorie') {
        nameInput.placeholder = engine.tr('PLACEHOLDER_NEW_CAT') || 'Ex: Abonnements Streaming...';
      } else if (activeEntityType === 'tiers') {
        nameInput.placeholder = engine.tr('PLACEHOLDER_NEW_TIERS') || 'Ex: Supermarché, Employeur...';
      }
    }

    // 6. Dynamic Localized Options for Select Dropdowns (Preserves current selection)
    const selClasse = document.getElementById('add-entity-classe');
    if (selClasse) {
      const curVal = selClasse.value || 'LIQUIDITE';
      selClasse.innerHTML = `
        <option value="LIQUIDITE">${engine.tr('OPT_CLASSE_LIQUIDITE') || 'Liquidité (Banque, Espèces)'}</option>
        <option value="INVESTISSEMENT">${engine.tr('OPT_CLASSE_INVESTISSEMENT') || 'Investissement (Actions, Crypto, Immo)'}</option>
        <option value="DETTE">${engine.tr('OPT_CLASSE_DETTE') || 'Dette & Emprunt (Prêt, Différé)'}</option>
      `;
      selClasse.value = curVal;
    }

    const selCatType = document.getElementById('add-entity-cat-type');
    if (selCatType) {
      const curVal = selCatType.value || 'DEPENSE';
      selCatType.innerHTML = `
        <option value="DEPENSE">${engine.tr('OPT_TYPE_DEPENSE') || 'Dépense (Consommation)'}</option>
        <option value="REVENU">${engine.tr('OPT_TYPE_REVENU') || 'Revenu (Entrée)'}</option>
        <option value="TRANSFERT">${engine.tr('OPT_TYPE_TRANSFERT') || 'Transfert / Épargne'}</option>
      `;
      selCatType.value = curVal;
    }

    const selSecteur = document.getElementById('add-entity-tiers-secteur');
    if (selSecteur) {
      const curVal = selSecteur.value || 'COMMERCE';
      selSecteur.innerHTML = `
        <option value="COMMERCE">${engine.tr('OPT_SECTEUR_COMMERCE') || 'Commerce / Supermarché'}</option>
        <option value="ENTREPRISE">${engine.tr('OPT_SECTEUR_ENTREPRISE') || 'Entreprise / Employeur'}</option>
        <option value="FINANCE">${engine.tr('OPT_SECTEUR_FINANCE') || 'Banque & Finance'}</option>
        <option value="INSTITUTION">${engine.tr('OPT_SECTEUR_INSTITUTION') || 'Institution Publique'}</option>
        <option value="IMMOBILIER">${engine.tr('OPT_SECTEUR_IMMOBILIER') || 'Immobilier / Propriétaire'}</option>
        <option value="SANTE">${engine.tr('OPT_SECTEUR_SANTE') || 'Santé & Pharmacie'}</option>
        <option value="AUTRE">${engine.tr('OPT_SECTEUR_AUTRE') || 'Autre Secteur'}</option>
      `;
      selSecteur.value = curVal;
    }

    // 7. Action Buttons
    const btnCancel = document.getElementById('btn-cancel-add-entity');
    if (btnCancel) btnCancel.textContent = engine.tr('BTN_CANCEL') || 'Annuler';
    const btnConfirm = document.getElementById('btn-confirm-add-entity');
    if (btnConfirm) btnConfirm.textContent = engine.tr('BTN_CONFIRM_ADD_ENTITY') || 'Ajouter & Sélectionner';
  }

  function updateFieldAddButtonsTranslations() {
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
  }

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

    updateAddEntityModalTranslations();
  }

  function openAddEntityDialog(type = 'compte', targetSelect = null) {
    currentTargetSelect = targetSelect;
    triggerHaptic(15);
    switchEntityType(type);
    updateAddEntityModalTranslations();

    const nameInput = document.getElementById('add-entity-name');
    if (nameInput) {
      nameInput.value = '';
    }
    const addDev = document.getElementById('add-entity-devise');
    if (addDev) {
      addDev.value = engine.currency || 'MUR';
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
        targetSel = document.getElementById(`tx-${entity}`);
      } else if (btn.id.includes('exp')) {
        targetSel = document.getElementById(`exp-tx-${entity}`);
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

  // Auto-sync transaction currency with account's native currency on selection
  document.getElementById('tx-compte')?.addEventListener('change', (e) => {
    const cptId = e.target.value;
    if (cptId === '__NEW__') return;
    const account = (engine.data?.comptes || []).find(c => String(c.id) === String(cptId));
    if (account && account.devise) {
      const devSel = document.getElementById('tx-devise');
      if (devSel) {
        devSel.value = account.devise;
        txDeviseUserModified = false;
      }
    }
  });

  document.getElementById('exp-tx-compte')?.addEventListener('change', (e) => {
    const cptId = e.target.value;
    if (cptId === '__NEW__') return;
    const account = (engine.data?.comptes || []).find(c => String(c.id) === String(cptId));
    if (account && account.devise) {
      const devSel = document.getElementById('exp-tx-devise');
      if (devSel) {
        devSel.value = account.devise;
        expTxDeviseUserModified = false;
      }
    }
  });

  // Track explicit manual user selection of currency in forms
  document.getElementById('tx-devise')?.addEventListener('change', () => {
    txDeviseUserModified = true;
  });
  document.getElementById('exp-tx-devise')?.addEventListener('change', () => {
    expTxDeviseUserModified = true;
  });
  document.getElementById('alloc-devise')?.addEventListener('change', (e) => {
    allocDeviseUserModified = true;
    const badgeCurr = document.getElementById('alloc-currency-badge');
    if (badgeCurr) badgeCurr.textContent = e.target.value;
  });

  // Submit new entity
  addEntityForm?.addEventListener('submit', (e) => {
    e.preventDefault();
    triggerHaptic(25);

    const nameVal = document.getElementById('add-entity-name')?.value?.trim();
    if (!nameVal) {
      showToast(engine.tr('ERR_NAME_REQUIRED') || 'Veuillez renseigner un nom valide.', 'error');
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

  // --- 9. FULL-SCREEN IMMERSIVE MODAL (SAISIE EXPRESS TRANSACTION) ---
  const fabBtn = document.getElementById('btn-fab-add');
  const modalExpress = document.getElementById('modal-bottom-sheet-tx');
  const btnCloseExpress = document.getElementById('btn-close-express-tx');

  function openExpressModal(defaultAccountId = null) {
    triggerHaptic(25);
    populateSelects();
    const validAccountId = (typeof defaultAccountId === 'string' || typeof defaultAccountId === 'number') ? String(defaultAccountId) : null;
    if (validAccountId) {
      const cptSel = document.getElementById('exp-tx-compte');
      if (cptSel) {
        cptSel.value = validAccountId;
      }
    }
    if (!expTxDeviseUserModified) {
      const cptId = document.getElementById('exp-tx-compte')?.value;
      const account = (engine.data?.comptes || []).find(c => String(c.id) === String(cptId));
      const devSel = document.getElementById('exp-tx-devise');
      if (devSel) {
        devSel.value = (account && account.devise) ? account.devise : engine.currency;
      }
    }
    modalExpress?.classList.add('open');
  }

  function closeExpressModal() {
    modalExpress?.classList.remove('open');
  }

  fabBtn?.addEventListener('click', () => openExpressModal());
  btnCloseExpress?.addEventListener('click', closeExpressModal);

  // Also bind clicking "+ Saisie Express" on Hub Central card to open express modal
  document.getElementById('btn-hub-quick-saisie')?.addEventListener('click', (e) => {
    e.preventDefault();
    e.stopPropagation();
    openExpressModal();
  });

  // Express Form Submit
  document.getElementById('form-express-tx')?.addEventListener('submit', (e) => {
    e.preventDefault();
    triggerHaptic(30);

    const dateVal = document.getElementById('exp-tx-date').value;
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

    engine.addTransaction({
      date: dateVal,
      id_compte: cptVal,
      id_categorie: catVal,
      id_tiers: tiersVal,
      montant: montantVal,
      devise: devVal,
      notes: notesVal
    });

    expTxDeviseUserModified = false;
    document.getElementById('exp-tx-montant').value = '';
    document.getElementById('exp-tx-notes').value = '';
    modalExpress?.classList.remove('open');

    showToast('Transaction express enregistrée !', 'success');
    navigateTo('cashflow');
  });

  // Main Form Submit (USF_Transaction)
  document.getElementById('form-transaction')?.addEventListener('submit', (e) => {
    e.preventDefault();
    triggerHaptic(30);

    const dateVal = document.getElementById('tx-date').value;
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
    if (newTiers) tiersVal = engine.addTiers(newTiers).id;

    engine.addTransaction({
      date: dateVal,
      id_compte: cptVal,
      id_categorie: catVal,
      id_tiers: tiersVal,
      montant: montantVal,
      devise: devVal,
      notes: notesVal
    });

    txDeviseUserModified = false;
    document.getElementById('tx-montant').value = '';
    document.getElementById('tx-notes').value = '';
    showToast('Transaction enregistrée avec succès !', 'success');
    navigateTo('cashflow');
  });

  // --- 9. BUDGET ALLOCATION MODAL (FULLSCREEN QUESTION-OPTIMIZED) ---
  const modalAlloc = document.getElementById('modal-budget-alloc');

  function updateAllocMoisLocalized() {
    const allocMois = document.getElementById('alloc-mois');
    const locDiv = document.getElementById('alloc-mois-localized');
    if (!allocMois || !locDiv) return;
    const val = allocMois.value;
    if (val && val.includes('-')) {
      const parts = val.split('-');
      const y = parts[0];
      const m = parts[1];
      const monthNames = {
        FR: ['Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre'],
        EN: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
        ES: ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'],
        PT: ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
      };
      const mIdx = parseInt(m, 10) - 1;
      const lang = engine.lang || 'FR';
      const mName = (monthNames[lang] && monthNames[lang][mIdx]) || m;
      locDiv.textContent = `${mName} ${y}`;
    }
  }

  document.getElementById('alloc-mois')?.addEventListener('input', updateAllocMoisLocalized);
  document.getElementById('alloc-mois')?.addEventListener('change', updateAllocMoisLocalized);

  document.getElementById('btn-open-alloc-modal')?.addEventListener('click', () => {
    triggerHaptic(15);
    populateSelects();
    modalAlloc?.classList.add('open');
    const targetMois = (engine.month === '00') ? `${engine.year}-02` : `${engine.year}-${engine.month}`;
    const allocMois = document.getElementById('alloc-mois');
    if (allocMois) {
      allocMois.value = targetMois;
      updateAllocMoisLocalized();
    }
    const allocDev = document.getElementById('alloc-devise');
    if (allocDev) allocDev.value = engine.currency;
    const badgeCurr = document.getElementById('alloc-currency-badge');
    if (badgeCurr) badgeCurr.textContent = engine.currency;
  });

  document.getElementById('alloc-devise')?.addEventListener('change', (e) => {
    const badgeCurr = document.getElementById('alloc-currency-badge');
    if (badgeCurr) badgeCurr.textContent = e.target.value;
  });

  document.getElementById('btn-close-alloc')?.addEventListener('click', () => {
    modalAlloc?.classList.remove('open');
  });

  modalAlloc?.addEventListener('click', (e) => {
    if (e.target === modalAlloc) modalAlloc.classList.remove('open');
  });

  document.getElementById('form-budget-alloc')?.addEventListener('submit', (e) => {
    e.preventDefault();
    triggerHaptic(25);
    const mois = document.getElementById('alloc-mois').value;
    const cat = document.getElementById('alloc-categorie').value;
    const amt = document.getElementById('alloc-montant').value;
    const dev = document.getElementById('alloc-devise')?.value || engine.currency;

    engine.allocateBudget(cat, amt, mois, dev);
    modalAlloc?.classList.remove('open');
    const toastMsg = {
      FR: 'Allocation budgétaire enregistrée !',
      EN: 'Budget allocation saved!',
      ES: '¡Asignación presupuestaria guardada!',
      PT: 'Alocação orçamentária salva com sucesso!'
    }[engine.lang] || 'Allocation budgétaire enregistrée !';
    showToast(toastMsg, 'success');
    renderBudgetView();
  });

  // --- HUB CONSOLIDATED FILTER CHIPS ---
  document.querySelectorAll('#hub-filter-bar .hub-filter-chip').forEach(chip => {
    chip.addEventListener('click', () => {
      triggerHaptic(15);
      document.querySelectorAll('#hub-filter-bar .hub-filter-chip').forEach(c => c.classList.remove('active'));
      chip.classList.add('active');
      engine.hubFilter = chip.dataset.hubFilter;
      renderHomeView();
    });
  });

  // --- DEDICATED DATABASE MANAGEMENT SYSTEM (SFP v1.0) ---
  // A. Refresh
  document.getElementById('btn-db-refresh')?.addEventListener('click', () => {
    triggerHaptic(20);
    engine.refreshDatabase();
    showToast('Base rafraîchie : calculs synchronisés', 'success');
    renderCurrentView();
  });

  // B. Backup
  document.getElementById('btn-db-backup')?.addEventListener('click', () => {
    triggerHaptic(20);
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(engine.backupDatabase());
    const dlAnchor = document.createElement('a');
    dlAnchor.setAttribute("href", dataStr);
    dlAnchor.setAttribute("download", `GFMP_Backup_${new Date().toISOString().split('T')[0]}.json`);
    document.body.appendChild(dlAnchor);
    dlAnchor.click();
    dlAnchor.remove();
    showToast('Sauvegarde de la base exportée', 'success');
  });

  // C. Initialiser
  document.getElementById('btn-db-init')?.addEventListener('click', async () => {
    triggerHaptic(25);
    const confirmed = await showConfirmDialog({
      title: 'Réinitialiser la Base',
      message: 'Réinitialiser la base rechargera les données officielles GFMP SFP v1.0 avec la configuration d\'usine. Confirmer ?',
      confirmText: 'Réinitialiser',
      cancelText: 'Annuler',
      type: 'warning'
    });
    if (confirmed) {
      engine.initDatabase();
      showToast('Base réinitialisée aux données de référence', 'success');
      populateSelects();
      renderCurrentView();
    }
  });

  // D. Purger
  document.getElementById('btn-db-purge')?.addEventListener('click', async () => {
    triggerHaptic(25);
    const confirmed = await showConfirmDialog({
      title: 'Purger la Base',
      message: 'Attention : Voulez-vous purger l\'ensemble des transactions et budgets de la base de données ? Les comptes et catégories seront conservés.',
      confirmText: 'Purger Définitivement',
      cancelText: 'Annuler',
      type: 'danger'
    });
    if (confirmed) {
      engine.purgeDatabase();
      showToast('Transactions et budgets purgés avec succès', 'info');
      populateSelects();
      renderCurrentView();
    }
  });

  // E. Audit History Modal
  const modalAudit = document.getElementById('modal-audit-history');
  document.getElementById('btn-db-audit')?.addEventListener('click', openAuditHistoryModal);

  document.getElementById('btn-close-history')?.addEventListener('click', () => modalAudit?.classList.remove('open'));
  document.getElementById('btn-close-history-footer')?.addEventListener('click', () => modalAudit?.classList.remove('open'));
  modalAudit?.addEventListener('click', (e) => {
    if (e.target === modalAudit) modalAudit.classList.remove('open');
  });

  // F. Executive Report shortcut from DB bar
  document.getElementById('btn-db-report')?.addEventListener('click', () => {
    triggerHaptic(20);
    document.getElementById('btn-open-report')?.click();
  });

  // --- 10. FINANCIAL EXECUTIVE REPORT MODAL ---
  const modalReport = document.getElementById('modal-financial-report');

  function renderFinancialReportModal() {
    const bodyEl = document.getElementById('report-modal-body');
    const titleEl = document.getElementById('lbl-report-modal-title');
    const btnClose = document.getElementById('btn-close-report');
    const btnCopy = document.getElementById('btn-copy-report');

    if (titleEl) titleEl.textContent = engine.tr('REPORT_TITLE');
    if (btnClose) btnClose.textContent = engine.tr('BTN_CLOSE');
    if (btnCopy) btnCopy.textContent = engine.tr('BTN_REPORT_COPY');

    if (!bodyEl) return;
    const rep = engine.generateFinancialReport();
    const summaryLabel = engine.tr('REPORT_SUMMARY', { period: rep.period });
    const consolCurrLabel = engine.tr('REPORT_CONSOL_CURR', { curr: rep.currency });
    const topExpLabel = engine.tr('REPORT_TOP_DEP');
    const netWorthLabel = engine.tr('REPORT_NET_WORTH');
    const assetsDebtsLabel = engine.tr('REPORT_ASSETS_DEBTS', { assets: rep.totalAssets, debts: rep.totalLiabilities });

    let topExpHtml = '';
    if (rep.topExpenses && rep.topExpenses.length > 0) {
      topExpHtml = rep.topExpenses.map(e => `
        <div style="display:flex; justify-content:space-between; font-size:0.78rem; padding:4px 0; border-bottom:1px solid rgba(255,255,255,0.06);">
          <span>${escapeHTML(engine.trEntity(e.nom))}</span>
          <span style="font-weight:700; color:var(--c-yellow-royal); font-feature-settings:'tnum';">${escapeHTML(e.montant)} (${escapeHTML(e.pct)})</span>
        </div>
      `).join('');
    } else {
      topExpHtml = `
        <div class="empty-state-rich" style="padding:14px 12px; margin-top:4px;">
          <div class="empty-state-title">${escapeHTML(engine.tr('NO_DEPENSE_PERIOD'))}</div>
          <p class="empty-state-desc">${escapeHTML(engine.tr('DONUT_EMPTY_HINT'))}</p>
          <div class="empty-state-actions">
            <button type="button" class="btn-empty-cta" data-empty-action="add-exp">${escapeHTML(engine.tr('EMPTY_CTA_ADD_EXP'))}</button>
            ${engine.month !== '00' ? `<button type="button" class="btn-empty-cta secondary" data-empty-action="view-full-year">${escapeHTML(engine.tr('EMPTY_CTA_FULL_YEAR'))}</button>` : ''}
          </div>
        </div>
      `;
    }

    const emptyDiagnosticHtml = !rep.hasData ? `
      <div class="empty-state-rich" style="margin-bottom:12px; padding:14px 12px; border-color:rgba(250,218,94,0.35); background:rgba(250,218,94,0.05);">
        <div class="empty-state-icon-box" aria-hidden="true">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
        </div>
        <div class="empty-state-title">${escapeHTML(engine.tr('REPORT_EMPTY_DIAG_TITLE'))}</div>
        <p class="empty-state-desc">${escapeHTML(engine.tr('REPORT_EMPTY_DIAG_DESC'))}</p>
      </div>
    ` : '';

    bodyEl.innerHTML = `
      <div style="background:rgba(255,255,255,0.05); padding:12px; border-radius:10px; margin-bottom:12px; border:1px solid rgba(250,218,94,0.3);">
        <div style="font-weight:800; color:var(--c-yellow-royal); font-size:1.05rem;">${summaryLabel}</div>
        <div style="font-size:0.75rem; color:var(--text-muted);">${consolCurrLabel}</div>
      </div>
      ${emptyDiagnosticHtml}
      <div style="display:grid; grid-template-columns:1fr 1fr; gap:8px; margin-bottom:12px;">
        <div style="background:rgba(16,185,129,0.1); border-left:3px solid #10B981; padding:8px; border-radius:6px;">
          <div style="font-size:0.7rem; color:#A7F3D0;">${engine.tr('REPORT_TOT_REV')}</div>
          <div style="font-weight:700; color:#FFF; font-feature-settings:'tnum';">${rep.revenus}</div>
        </div>
        <div style="background:rgba(239,68,68,0.1); border-left:3px solid #EF4444; padding:8px; border-radius:6px;">
          <div style="font-size:0.7rem; color:#FECACA;">${engine.tr('REPORT_TOT_DEP')}</div>
          <div style="font-weight:700; color:#FFF; font-feature-settings:'tnum';">${rep.depenses}</div>
        </div>
        <div style="background:rgba(65,105,225,0.1); border-left:3px solid #4169E1; padding:8px; border-radius:6px;">
          <div style="font-size:0.7rem; color:#BFDBFE;">${engine.tr('REPORT_NET_CF')}</div>
          <div style="font-weight:700; color:#FFF; font-feature-settings:'tnum';">${rep.netCashflow}</div>
        </div>
        <div style="background:rgba(250,218,94,0.1); border-left:3px solid #FADA5E; padding:8px; border-radius:6px;">
          <div style="font-size:0.7rem; color:#FEF08A;">${engine.tr('REPORT_SAVINGS_RATE')}</div>
          <div style="font-weight:700; color:#FFF; font-feature-settings:'tnum';">${rep.savingsRate}</div>
        </div>
      </div>
      <div style="background:rgba(255,255,255,0.04); padding:10px; border-radius:8px; margin-bottom:12px;">
        <div style="font-weight:700; color:#FFF; margin-bottom:4px;">${netWorthLabel}</div>
        <div style="font-size:1.15rem; font-weight:800; color:var(--c-yellow-royal); font-feature-settings:'tnum';">${rep.netWorth}</div>
        <div style="font-size:0.75rem; color:var(--text-muted);">${assetsDebtsLabel}</div>
      </div>
      <div style="font-weight:700; margin-bottom:6px; color:#FFF;">${topExpLabel}</div>
      ${topExpHtml}
    `;
  }

  document.getElementById('btn-open-report')?.addEventListener('click', () => {
    triggerHaptic(20);
    renderFinancialReportModal();
    modalReport?.classList.add('open');
  });

  document.getElementById('btn-close-report')?.addEventListener('click', () => {
    modalReport?.classList.remove('open');
  });

  modalReport?.addEventListener('click', (e) => {
    if (e.target === modalReport) modalReport.classList.remove('open');
  });

  document.getElementById('btn-copy-report')?.addEventListener('click', () => {
    triggerHaptic(15);
    const rep = engine.generateFinancialReport();
    const text = `=== GFMP ${engine.tr('REPORT_TITLE')} (${rep.period}) ===\n${engine.tr('REPORT_CONSOL_CURR', { curr: rep.currency })}\n${engine.tr('REPORT_TOT_REV')} : ${rep.revenus}\n${engine.tr('REPORT_TOT_DEP')} : ${rep.depenses}\n${engine.tr('REPORT_NET_CF')} : ${rep.netCashflow}\n${engine.tr('REPORT_SAVINGS_RATE')} : ${rep.savingsRate}\n${engine.tr('REPORT_NET_WORTH')} : ${rep.netWorth}\n${engine.tr('REPORT_ASSETS_DEBTS', { assets: rep.totalAssets, debts: rep.totalLiabilities })}\n${engine.tr('BUDGET_CONSUMPTION')} : ${rep.budgetConsumption}\n========================================`;
    if (navigator.clipboard) {
      navigator.clipboard.writeText(text).then(() => {
        showToast(engine.tr('TOAST_REPORT_COPIED'), 'success');
      }).catch(() => {
        showToast(engine.tr('TOAST_REPORT_GENERATED'), 'info');
      });
    } else {
      showToast(engine.tr('TOAST_REPORT_GENERATED'), 'info');
    }
  });

  // --- 11. SEARCH & FILTER INTERACTIVITY ---
  // --- 11. SEARCH & FILTERS INTERACTIVITY ---
  // A. Cashflow Search
  const searchInput = document.getElementById('input-tx-search');
  searchInput?.addEventListener('input', (e) => {
    engine.searchTerm = e.target.value;
    renderCashflowView();
  });

  // B. Cashflow Flux Filter Chips (Tous / Revenus / Dépenses / Virements)
  document.querySelectorAll('.search-filter-box .filter-chip[data-filter]').forEach(chip => {
    chip.addEventListener('click', () => {
      triggerHaptic(12);
      document.querySelectorAll('.search-filter-box .filter-chip[data-filter]').forEach(c => c.classList.remove('active'));
      chip.classList.add('active');
      engine.fluxFilter = chip.dataset.filter;
      renderCashflowView();
    });
  });

  // C. Cashflow Account Filter Dropdown
  document.getElementById('cf-account-filter')?.addEventListener('change', (e) => {
    triggerHaptic(12);
    engine.accountFilter = e.target.value;
    renderCashflowView();
  });

  // D. Cashflow Sort Dropdown
  document.getElementById('cf-sort-select')?.addEventListener('change', (e) => {
    triggerHaptic(12);
    const val = e.target.value;
    if (val === 'date-desc') { engine.sortField = 'date'; engine.sortDir = 'desc'; }
    else if (val === 'date-asc') { engine.sortField = 'date'; engine.sortDir = 'asc'; }
    else if (val === 'amount-desc') { engine.sortField = 'montant'; engine.sortDir = 'desc'; }
    else if (val === 'amount-asc') { engine.sortField = 'montant'; engine.sortDir = 'asc'; }
    renderCashflowView();
  });

  // E. Patrimoine Account Class Filter Chips
  document.querySelectorAll('#nw-class-chips .filter-chip').forEach(chip => {
    chip.addEventListener('click', () => {
      triggerHaptic(12);
      document.querySelectorAll('#nw-class-chips .filter-chip').forEach(c => c.classList.remove('active'));
      chip.classList.add('active');
      engine.patrimoineClassFilter = chip.dataset.class;
      renderNetWorthView();
    });
  });

  // F. Budget Status Filter Chips
  document.querySelectorAll('#bg-status-chips .filter-chip').forEach(chip => {
    chip.addEventListener('click', () => {
      triggerHaptic(12);
      document.querySelectorAll('#bg-status-chips .filter-chip').forEach(c => c.classList.remove('active'));
      chip.classList.add('active');
      engine.budgetStatusFilter = chip.dataset.bstatus;
      renderBudgetView();
    });
  });

  // --- 12. INTERACTIVE SVG CHART GENERATORS (BIDIRECTIONAL CROSS-FILTERING) ---

  // A. Cashflow Multi-Month Trend Chart (6 Months SVG, Interactive Click to switch Month & Resilient Zero-State)
  function renderTrendChart(containerId) {
    const container = document.getElementById(containerId);
    if (!container) return;

    const trend = engine.computeCashflowTrend(6);
    if (!trend || trend.length === 0) {
      container.innerHTML = `
        <div class="empty-state-rich">
          <div class="empty-state-title">${escapeHTML(engine.tr('TREND_EMPTY_TITLE'))}</div>
          <p class="empty-state-desc">${escapeHTML(engine.tr('TREND_EMPTY_SUB'))}</p>
          <div class="empty-state-actions">
            <button type="button" class="btn-empty-cta" data-empty-action="add-tx">${escapeHTML(engine.tr('EMPTY_CTA_ADD_TX'))}</button>
          </div>
        </div>
      `;
      return;
    }

    let maxVal = 0;
    let hasAnyTrendFlux = false;
    trend.forEach(t => {
      if (t.revenus > 0 || t.depenses > 0) hasAnyTrendFlux = true;
      if (t.revenus > maxVal) maxVal = t.revenus;
      if (t.depenses > maxVal) maxVal = t.depenses;
    });
    if (maxVal === 0) maxVal = 1000;

    const width = 340;
    const height = 180;
    const padX = 25;
    const padBottom = 28;
    const padTop = 20;
    const chartHeight = height - padBottom - padTop;
    const step = (width - padX * 2) / trend.length;
    const barW = Math.max(8, step * 0.32);

    let barsHtml = '';
    trend.forEach((t, i) => {
      const centerX = padX + i * step + step / 2;
      const hRev = t.revenus > 0 ? Math.max(4, (t.revenus / maxVal) * chartHeight) : 3;
      const hDep = t.depenses > 0 ? Math.max(4, (t.depenses / maxVal) * chartHeight) : 3;

      const yRev = height - padBottom - hRev;
      const yDep = height - padBottom - hDep;

      const xRev = centerX - barW - 1.5;
      const xDep = centerX + 1.5;

      const isCurrentMonth = t.yearMonth === `${engine.year}-${engine.month}`;
      const revFill = t.revenus > 0 ? '#10B981' : 'rgba(16,185,129,0.28)';
      const depFill = t.depenses > 0 ? '#EF4444' : 'rgba(239,68,68,0.28)';

      barsHtml += `
        <!-- Month ${t.label} Group (Clickable to switch active month) -->
        <g class="chart-bar-group ${isCurrentMonth ? 'active' : ''}" data-ym="${t.yearMonth}" data-label="${t.label}" style="cursor:pointer;">
          ${isCurrentMonth ? `<rect x="${centerX - step / 2 + 2}" y="${padTop - 8}" width="${step - 4}" height="${chartHeight + 14}" rx="8" fill="rgba(250,218,94,0.12)" stroke="#FADA5E" stroke-width="1.5" />` : ''}
          <!-- Revenus (Vert) -->
          <rect x="${xRev}" y="${yRev}" width="${barW}" height="${hRev}" rx="3" fill="${revFill}" filter="drop-shadow(0 2px 4px rgba(16,185,129,0.3))">
            <title>${t.label} Revenus : ${engine.formatMoney(t.revenus)}</title>
          </rect>
          <!-- Dépenses (Rouge) -->
          <rect x="${xDep}" y="${yDep}" width="${barW}" height="${hDep}" rx="3" fill="${depFill}" filter="drop-shadow(0 2px 4px rgba(239,68,68,0.3))">
            <title>${t.label} Dépenses : ${engine.formatMoney(t.depenses)}</title>
          </rect>
          <!-- Label Mois -->
          <text x="${centerX}" y="${height - 8}" text-anchor="middle" fill="${isCurrentMonth ? '#FADA5E' : '#94A3B8'}" font-size="9" font-weight="${isCurrentMonth ? '900' : '600'}" font-family="Inter, sans-serif">${t.label}</text>
        </g>
      `;
    });

    const zeroTrendCallout = !hasAnyTrendFlux ? `
      <div class="trend-empty-callout">
        <div style="text-align:left; flex:1; min-width:180px;">
          <div style="font-size:0.76rem; font-weight:800; color:#FFF;">${escapeHTML(engine.tr('TREND_EMPTY_TITLE'))}</div>
          <div style="font-size:0.68rem; color:var(--text-secondary); margin-top:2px;">${escapeHTML(engine.tr('TREND_EMPTY_SUB'))}</div>
        </div>
        <button type="button" class="btn-empty-cta" data-empty-action="add-tx">${escapeHTML(engine.tr('EMPTY_CTA_ADD_TX'))}</button>
      </div>
    ` : '';

    container.innerHTML = `
      <svg viewBox="0 0 ${width} ${height}" style="width:100%; height:100%; overflow:visible;">
        <!-- Horizontal grid lines -->
        <line x1="${padX}" y1="${padTop}" x2="${width - padX}" y2="${padTop}" stroke="rgba(255,255,255,0.08)" stroke-dasharray="3 3" />
        <line x1="${padX}" y1="${padTop + chartHeight / 2}" x2="${width - padX}" y2="${padTop + chartHeight / 2}" stroke="rgba(255,255,255,0.08)" stroke-dasharray="3 3" />
        <line x1="${padX}" y1="${height - padBottom}" x2="${width - padX}" y2="${height - padBottom}" stroke="rgba(255,255,255,0.2)" />
        
        <!-- Bars -->
        ${barsHtml}
      </svg>
      <div style="display:flex; justify-content:center; gap:16px; margin-top:6px; font-size:0.7rem;">
        <span style="display:inline-flex; align-items:center; gap:5px; color:#A7F3D0;"><span style="width:8px; height:8px; border-radius:2px; background:#10B981;"></span> ${escapeHTML(engine.tr('REVENUES') || 'Revenus')}</span>
        <span style="display:inline-flex; align-items:center; gap:5px; color:#FECACA;"><span style="width:8px; height:8px; border-radius:2px; background:#EF4444;"></span> ${escapeHTML(engine.tr('DEPENSES') || 'Dépenses')}</span>
      </div>
      ${zeroTrendCallout}
    `;

    // Bind Bar Click -> Switch Active Month Bidirectionally
    container.querySelectorAll('.chart-bar-group').forEach(group => {
      group.addEventListener('click', () => {
        triggerHaptic(20);
        const ym = group.dataset.ym;
        if (ym) {
          const parts = ym.split('-');
          engine.setYear(parts[0]);
          engine.setMonth(parts[1]);
          showToast(`Période sélectionnée : ${engine.getMonthLabel()} ${parts[0]}`, 'info');
        }
      });
    });
  }

  // B. Category Expense Donut Chart (SVG, Interactive Click to filter table & Golden Standard Styling)
  function renderDonutChart(containerId, breakdown, totalDepenses) {
    const container = document.getElementById(containerId);
    if (!container) return;

    if (!breakdown || breakdown.length === 0 || totalDepenses <= 0) {
      const emptyMsg = engine.tr('NO_EXPENSE_PERIOD', { period: `${engine.getMonthLabel()} ${engine.year}` }) ||
        `Aucune dépense enregistrée pour ${engine.getMonthLabel()} ${engine.year}`;
      const emptyHint = engine.tr('DONUT_EMPTY_HINT');
      const size = 124;
      const strokeW = 14;
      const radius = (size - strokeW) / 2;
      const center = size / 2;
      container.innerHTML = `
        <div class="empty-state-donut-wrap">
          <div class="donut-ring-box" style="position:relative; width:${size}px; height:${size}px; flex-shrink:0;">
            <svg width="${size}" height="${size}" viewBox="0 0 ${size} ${size}" role="img" aria-label="${escapeHTML(emptyMsg)}">
              <circle cx="${center}" cy="${center}" r="${radius}" fill="transparent" stroke="rgba(255,255,255,0.06)" stroke-width="${strokeW}" />
              <circle cx="${center}" cy="${center}" r="${radius}" fill="transparent" stroke="rgba(250,218,94,0.22)" stroke-width="2" stroke-dasharray="5 5" />
            </svg>
            <div class="donut-center-box" style="position:absolute; top:0; left:0; width:100%; height:100%; display:flex; flex-direction:column; align-items:center; justify-content:center; pointer-events:none; padding:8px; text-align:center;">
              <div class="donut-center-lbl" style="font-size:0.58rem; font-weight:700; color:var(--text-muted); text-transform:uppercase;">${escapeHTML(engine.tr('TOTAL') || 'TOTAL')}</div>
              <div class="donut-center-val" style="font-size:0.76rem; font-weight:900; color:var(--text-secondary); font-feature-settings:'tnum'; margin-top:2px;">${escapeHTML(engine.formatMoney(0))}</div>
            </div>
          </div>
          <div class="empty-state-rich" style="flex:1; min-width:200px; padding:14px 12px; margin:0;">
            <div class="empty-state-title">${escapeHTML(emptyMsg)}</div>
            <p class="empty-state-desc">${escapeHTML(emptyHint)}</p>
            <div class="empty-state-actions">
              <button type="button" class="btn-empty-cta" data-empty-action="add-exp">${escapeHTML(engine.tr('EMPTY_CTA_ADD_EXP'))}</button>
              ${engine.month !== '00' ? `<button type="button" class="btn-empty-cta secondary" data-empty-action="view-full-year">${escapeHTML(engine.tr('EMPTY_CTA_FULL_YEAR'))}</button>` : ''}
            </div>
          </div>
        </div>
      `;
      return;
    }

    const size = 140;
    const strokeW = 16;
    const radius = (size - strokeW) / 2;
    const center = size / 2;
    const circ = 2 * Math.PI * radius;

    const formatCatPct = (pct) => {
      if (!pct || pct <= 0) return '0';
      if (pct < 1) return pct.toFixed(1);
      return String(Math.round(pct));
    };

    let accumPct = 0;
    let pathsHtml = '';

    breakdown.forEach(item => {
      const slicePct = item.pct / 100;
      const rawArc = slicePct * circ;
      const effectiveGap = breakdown.length > 1 ? Math.min(2, rawArc * 0.25) : 0;
      const sliceLength = Math.max(3, rawArc - effectiveGap);
      const strokeDash = `${sliceLength} ${circ}`;
      const strokeOffset = -(accumPct * circ);
      accumPct += slicePct;

      const isSelected = (engine.selectedCategoryFilter === item.nom);
      const catTraduit = item.nomTraduit || engine.trEntity(item.nom);
      const pctDisplay = formatCatPct(item.pct);
      const titleText = `${catTraduit}: ${engine.formatMoney(item.montant)} (${pctDisplay}%) - ${engine.tr('CLICK_TO_FILTER') || 'Cliquer pour filtrer'}`;

      pathsHtml += `
        <circle cx="${center}" cy="${center}" r="${radius}"
          class="donut-slice ${isSelected ? 'active' : ''}"
          data-cat="${item.nom}"
          data-amount="${item.montant}"
          data-pct="${pctDisplay}"
          data-color="${item.color}"
          data-label="${catTraduit}"
          fill="transparent"
          stroke="${item.color}"
          stroke-width="${isSelected ? 20 : strokeW}"
          stroke-dasharray="${strokeDash}"
          stroke-dashoffset="${strokeOffset}"
          transform="rotate(-90 ${center} ${center})"
          tabindex="0"
          role="graphics-symbol"
          aria-label="${titleText}">
          <title>${titleText}</title>
        </circle>
      `;
    });

    // Display ALL categories in legend with 100% interactivity, exact amounts, percentages, and micro-bars
    let legendItemsHtml = '';
    breakdown.forEach(item => {
      const isSelected = (engine.selectedCategoryFilter === item.nom);
      const catTraduit = item.nomTraduit || engine.trEntity(item.nom);
      const pctDisplay = formatCatPct(item.pct);
      const barWidth = Math.max(3, Math.min(100, item.pct));
      legendItemsHtml += `
        <div class="legend-item ${isSelected ? 'active' : ''}"
             data-cat="${item.nom}"
             data-amount="${item.montant}"
             data-pct="${pctDisplay}"
             data-color="${item.color}"
             data-label="${catTraduit}"
             style="--cat-color: ${item.color};"
             tabindex="0"
             role="button"
             aria-pressed="${isSelected ? 'true' : 'false'}"
             title="${catTraduit}: ${engine.formatMoney(item.montant)} (${pctDisplay}%) - ${engine.tr('CLICK_TO_FILTER') || 'Cliquer pour filtrer'}">
          <div class="legend-item-main">
            <span class="legend-label-wrap">
              <span class="legend-color-dot" style="background:${item.color}; box-shadow: 0 0 6px ${item.color}88;"></span>
              <span class="legend-text ${isSelected ? 'selected' : ''}">${catTraduit}</span>
            </span>
            <span class="legend-metrics">
              <span class="legend-amount">${engine.formatMoney(item.montant)}</span>
              <span class="legend-pct">${pctDisplay}%</span>
            </span>
          </div>
          <div class="legend-mini-bar-track" aria-hidden="true">
            <div class="legend-mini-bar-fill" style="width:${barWidth}%; background:${item.color};"></div>
          </div>
        </div>
      `;
    });

    const activeItem = breakdown.find(b => b.nom === engine.selectedCategoryFilter);
    const initialLabel = activeItem ? (activeItem.nomTraduit || engine.trEntity(activeItem.nom)) : (engine.tr('TOTAL') || 'TOTAL');
    const initialVal = activeItem ? engine.formatMoney(activeItem.montant) : engine.formatMoney(totalDepenses);
    const initialPct = activeItem ? `${formatCatPct(activeItem.pct)}%` : '';

    container.innerHTML = `
      <div class="donut-ring-box" style="position:relative; width:${size}px; height:${size}px; flex-shrink:0;">
        <svg width="${size}" height="${size}" viewBox="0 0 ${size} ${size}" role="img" aria-label="${engine.tr('HUB_DONUT_TITLE') || 'Répartition des dépenses par catégorie'}">
          <circle cx="${center}" cy="${center}" r="${radius}" fill="transparent" stroke="rgba(255,255,255,0.06)" stroke-width="${strokeW}" />
          ${pathsHtml}
        </svg>
        <div class="donut-center-box" id="${containerId}-center-box" style="position:absolute; top:0; left:0; width:100%; height:100%; display:flex; flex-direction:column; align-items:center; justify-content:center; pointer-events:none; padding:10px; text-align:center;">
          <div class="donut-center-lbl" style="font-size:0.6rem; font-weight:700; color:var(--text-muted); text-transform:uppercase; letter-spacing:0.5px; max-width:92px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">
            ${initialLabel}
          </div>
          <div class="donut-center-val" style="font-size:0.8rem; font-weight:900; color:var(--c-yellow-royal); max-width:98px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; font-feature-settings:'tnum'; margin-top:1px;">
            ${initialVal}
          </div>
          <div class="donut-center-sub" style="font-size:0.62rem; font-weight:700; color:#93C5FD; margin-top:1px; ${initialPct ? '' : 'display:none;'}">
            ${initialPct}
          </div>
        </div>
      </div>
      <div class="donut-legend" role="list" aria-label="${engine.tr('HUB_DONUT_TITLE') || 'Catégories de dépenses'}">
        ${legendItemsHtml}
      </div>
    `;

    // Interactive Hover & Focus Linked Behavior
    const centerBox = document.getElementById(`${containerId}-center-box`);
    const lblEl = centerBox?.querySelector('.donut-center-lbl');
    const valEl = centerBox?.querySelector('.donut-center-val');
    const subEl = centerBox?.querySelector('.donut-center-sub');

    const updateCenter = (label, value, pct) => {
      if (!lblEl || !valEl) return;
      lblEl.textContent = label;
      valEl.textContent = value;
      if (subEl) {
        if (pct) {
          subEl.textContent = `${pct}%`;
          subEl.style.display = 'block';
        } else {
          subEl.textContent = '';
          subEl.style.display = 'none';
        }
      }
    };

    const resetCenter = () => {
      if (engine.selectedCategoryFilter) {
        const curActive = breakdown.find(b => b.nom === engine.selectedCategoryFilter);
        if (curActive) {
          updateCenter(curActive.nomTraduit || engine.trEntity(curActive.nom), engine.formatMoney(curActive.montant), formatCatPct(curActive.pct));
          return;
        }
      }
      updateCenter(engine.tr('TOTAL') || 'TOTAL', engine.formatMoney(totalDepenses), '');
    };

    const slices = container.querySelectorAll('.donut-slice');
    const legendItems = container.querySelectorAll('.legend-item[data-cat]');

    const highlightCat = (catNom, isHover) => {
      slices.forEach(s => {
        if (s.dataset.cat === catNom) {
          s.classList.toggle('hovered', isHover);
          s.classList.remove('dimmed');
        } else {
          s.classList.remove('hovered');
          s.classList.toggle('dimmed', isHover);
        }
      });
      legendItems.forEach(li => {
        if (li.dataset.cat === catNom) {
          li.classList.toggle('hovered', isHover);
        } else {
          li.classList.remove('hovered');
        }
      });
    };

    // Bind Hover & Focus on Slices
    slices.forEach(slice => {
      const cat = slice.dataset.cat;
      const amount = parseFloat(slice.dataset.amount);
      const pct = slice.dataset.pct;
      const label = slice.dataset.label;

      slice.addEventListener('mouseenter', () => {
        highlightCat(cat, true);
        updateCenter(label, engine.formatMoney(amount), pct);
      });
      slice.addEventListener('mouseleave', () => {
        highlightCat(cat, false);
        resetCenter();
      });
      slice.addEventListener('focus', () => {
        highlightCat(cat, true);
        updateCenter(label, engine.formatMoney(amount), pct);
      });
      slice.addEventListener('blur', () => {
        highlightCat(cat, false);
        resetCenter();
      });
    });

    // Bind Hover & Focus on Legend Items
    legendItems.forEach(li => {
      const cat = li.dataset.cat;
      const amount = parseFloat(li.dataset.amount);
      const pct = li.dataset.pct;
      const label = li.dataset.label;

      li.addEventListener('mouseenter', () => {
        highlightCat(cat, true);
        updateCenter(label, engine.formatMoney(amount), pct);
      });
      li.addEventListener('mouseleave', () => {
        highlightCat(cat, false);
        resetCenter();
      });
      li.addEventListener('focus', () => {
        highlightCat(cat, true);
        updateCenter(label, engine.formatMoney(amount), pct);
      });
      li.addEventListener('blur', () => {
        highlightCat(cat, false);
        resetCenter();
      });
    });

    // Bind Click to Toggle Category Filter
    container.querySelectorAll('.donut-slice, .legend-item[data-cat]').forEach(el => {
      const clickHandler = () => {
        triggerHaptic(20);
        const cat = el.dataset.cat;
        if (engine.selectedCategoryFilter === cat) {
          engine.selectedCategoryFilter = null; // Clear filter
          showToast(engine.tr('FILTER_CLEARED') || 'Filtre de catégorie effacé', 'info');
        } else {
          engine.selectedCategoryFilter = cat; // Apply filter
          const catTrad = engine.trEntity(cat);
          showToast(`${engine.tr('FILTER_ACTIVE') || 'Filtre actif'} : ${catTrad}`, 'info');
        }

        if (engine.currentView === 'home') {
          renderHomeView();
        } else if (engine.currentView === 'cashflow') {
          renderCashflowView();
        } else {
          renderCurrentView();
        }
      };

      el.addEventListener('click', clickHandler);
      el.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          clickHandler();
        }
      });
    });
  }

  // --- 13. RENDER ACTIVE VIEWS & REACTIVE FILTER PROPAGATION ---
  function renderCurrentView() {
    // 1. Header Year Stepper
    const lblYear = document.getElementById('lbl-year');
    if (lblYear) lblYear.textContent = engine.year;

    // 2. Header Month Strip (Active state & Multilingual labels)
    const monthNames = {
      FR: ['Jan', 'Fév', 'Mar', 'Avr', 'Mai', 'Juin', 'Juil', 'Août', 'Sep', 'Oct', 'Nov', 'Déc', 'Année'],
      EN: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Year'],
      ES: ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic', 'Año'],
      PT: ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez', 'Ano']
    };
    const langMonths = monthNames[engine.lang] || monthNames['FR'];
    document.querySelectorAll('#header-month-strip .month-chip').forEach(chip => {
      const m = parseInt(chip.dataset.month, 10);
      if (m >= 1 && m <= 12) chip.textContent = langMonths[m - 1];
      else if (chip.dataset.month === '00') chip.textContent = langMonths[12];

      if (chip.dataset.month === engine.month) {
        chip.classList.add('active');
        chip.scrollIntoView({ behavior: 'smooth', inline: 'center', block: 'nearest' });
      } else {
        chip.classList.remove('active');
      }
    });

    // 3. Header Date Display (Dynamically reflects active year & month)
    const dateDisplayEl = document.getElementById('date-display');
    if (dateDisplayEl) {
      if (engine.month === '00') {
        dateDisplayEl.textContent = engine.tr('FULL_YEAR', { year: engine.year });
      } else {
        dateDisplayEl.textContent = `${engine.getMonthLabel()} ${engine.year}`;
      }
    }

    // 4. Header Currency Buttons
    document.querySelectorAll('.curr-btn').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.curr === engine.currency);
    });

    // 5. Header Language Buttons & Document Human Language Attribute
    if (typeof document !== 'undefined' && document.documentElement) {
      document.documentElement.lang = (engine.lang || 'FR').toLowerCase();
    }
    document.querySelectorAll('.lang-btn').forEach(btn => {
      const isAct = (btn.dataset.lang === engine.lang);
      btn.classList.toggle('active', isAct);
      btn.setAttribute('aria-pressed', isAct ? 'true' : 'false');
    });

    // 6. Currency Synchronization across all Views, Forms and Modals (Dynamic and Responsive)
    const txDevise = document.getElementById('tx-devise');
    if (txDevise && !txDeviseUserModified) {
      const activeCpt = (engine.data?.comptes || []).find(c => String(c.id) === String(document.getElementById('tx-compte')?.value));
      txDevise.value = (activeCpt && activeCpt.devise) ? activeCpt.devise : engine.currency;
    }
    const expTxDevise = document.getElementById('exp-tx-devise');
    if (expTxDevise && !expTxDeviseUserModified) {
      const activeExpCpt = (engine.data?.comptes || []).find(c => String(c.id) === String(document.getElementById('exp-tx-compte')?.value));
      expTxDevise.value = (activeExpCpt && activeExpCpt.devise) ? activeExpCpt.devise : engine.currency;
    }
    const allocDevise = document.getElementById('alloc-devise');
    if (allocDevise && !allocDeviseUserModified) {
      allocDevise.value = engine.currency;
    }
    const allocBadge = document.getElementById('alloc-currency-badge');
    if (allocBadge) allocBadge.textContent = (allocDevise ? allocDevise.value : engine.currency);
    const setCurrSel = document.getElementById('settings-currency-select');
    if (setCurrSel && setCurrSel.value !== engine.currency) setCurrSel.value = engine.currency;
    const setLangSel = document.getElementById('settings-lang-select');
    if (setLangSel && setLangSel.value !== engine.lang) setLangSel.value = engine.lang;

    // 7. Bottom Navigation Tab Labels
    const tabHome = document.getElementById('tab-label-home');
    if (tabHome) tabHome.textContent = engine.tr('NAV_HOME');
    const tabTx = document.getElementById('tab-label-transaction');
    if (tabTx) tabTx.textContent = engine.tr('NAV_SAISIE');
    const tabCf = document.getElementById('tab-label-cashflow');
    if (tabCf) tabCf.textContent = engine.tr('NAV_CASHFLOW');
    const tabNw = document.getElementById('tab-label-networth');
    if (tabNw) tabNw.textContent = engine.tr('NAV_NETWORTH');
    const tabBg = document.getElementById('tab-label-budget');
    if (tabBg) tabBg.textContent = engine.tr('NAV_BUDGET');
    const tabSet = document.getElementById('tab-label-settings');
    if (tabSet) tabSet.textContent = engine.tr('NAV_SETTINGS');

    // 8. Back Buttons across Sub-Views
    document.querySelectorAll('.btn-nav-home span').forEach(span => {
      span.textContent = engine.tr('BACK_HUB');
    });

    // 9. Dynamic Titles, Greetings & Descriptions
    const appTitle = document.getElementById('app-title-display');
    if (appTitle) appTitle.textContent = engine.tr('APP_TITLE') || 'SYSTÈME FINANCIER PERSONNEL';
    const txtWelcome = document.getElementById('txt-welcome');
    if (txtWelcome) txtWelcome.textContent = engine.tr('WELCOME');
    const lblCardSaisie = document.getElementById('lbl-card-saisie');
    if (lblCardSaisie) lblCardSaisie.textContent = engine.tr('SAISIE_T');
    const descCardSaisie = document.getElementById('desc-card-saisie');
    if (descCardSaisie) descCardSaisie.textContent = engine.tr('SAISIE_D');
    const lblCardCf = document.getElementById('lbl-card-cashflow');
    if (lblCardCf) lblCardCf.textContent = engine.tr('DASH_T');
    const descCardCf = document.getElementById('desc-card-cashflow');
    if (descCardCf) descCardCf.textContent = engine.tr('DASH_D');
    const lblCardNw = document.getElementById('lbl-card-networth');
    if (lblCardNw) lblCardNw.textContent = engine.tr('NETW_T');
    const descCardNw = document.getElementById('desc-card-networth');
    if (descCardNw) descCardNw.textContent = engine.tr('NETW_D');
    const lblCardBg = document.getElementById('lbl-card-budget');
    if (lblCardBg) lblCardBg.textContent = engine.tr('BUDG_T');
    const descCardBg = document.getElementById('desc-card-budget');
    if (descCardBg) descCardBg.textContent = engine.tr('BUDG_D');

    // 10. Hub Central Labels & Section Titles
    const hubChipAll = document.getElementById('hub-chip-all');
    if (hubChipAll) hubChipAll.textContent = engine.tr('HUB_ALL');
    const hubChipCf = document.getElementById('hub-chip-cashflow');
    if (hubChipCf) hubChipCf.textContent = engine.tr('HUB_CASHFLOW');
    const hubChipPat = document.getElementById('hub-chip-patrimoine');
    if (hubChipPat) hubChipPat.textContent = engine.tr('HUB_PATRIMOINE');
    const hubChipBg = document.getElementById('hub-chip-budget');
    if (hubChipBg) hubChipBg.textContent = engine.tr('HUB_BUDGET');

    const portCons = document.getElementById('lbl-portfolio-consolidated');
    if (portCons) portCons.textContent = engine.tr('PORTFOLIO_CONSOLIDATED');
    const saisieExp = document.getElementById('lbl-card-saisie-express');
    if (saisieExp) saisieExp.textContent = engine.tr('SAISIE_EXPRESS');
    const saisieSub = document.getElementById('lbl-card-saisie-sub');
    if (saisieSub) saisieSub.textContent = engine.tr('DIRECT_LOCAL');
    const hubCf = document.getElementById('lbl-hub-cf');
    if (hubCf) hubCf.textContent = engine.tr('NET_CASHFLOW');
    const hubBg = document.getElementById('lbl-hub-bg');
    if (hubBg) hubBg.textContent = engine.tr('BUDGET_CONSUMPTION');
    const hubAssets = document.getElementById('lbl-hub-assets');
    if (hubAssets) hubAssets.textContent = engine.tr('TOTAL_ASSETS');
    const hubDebts = document.getElementById('lbl-hub-debts');
    if (hubDebts) hubDebts.textContent = engine.tr('TOTAL_DEBTS');
    const hubSolvencyTitle = document.getElementById('lbl-hub-solvency-title');
    if (hubSolvencyTitle) hubSolvencyTitle.textContent = engine.tr('ASSETS_VS_DEBTS');

    const hubTrendTitle = document.getElementById('hub-trend-title');
    if (hubTrendTitle) hubTrendTitle.textContent = engine.tr('HUB_TREND_TITLE');
    const hubDonutTitle = document.getElementById('hub-donut-title');
    if (hubDonutTitle) hubDonutTitle.textContent = engine.tr('HUB_DONUT_TITLE');
    const hubPatTitle = document.getElementById('hub-patrimoine-title');
    if (hubPatTitle) hubPatTitle.textContent = engine.tr('HUB_PATRIMOINE_TITLE');
    const hubTopAssetsTitle = document.getElementById('lbl-hub-top-assets-title');
    if (hubTopAssetsTitle) hubTopAssetsTitle.textContent = engine.tr('TOP_ASSETS_LABEL');
    const hubBgTitle = document.getElementById('hub-budget-title');
    if (hubBgTitle) hubBgTitle.textContent = engine.tr('HUB_BUDGET_TITLE');
    const hubRecentTitle = document.getElementById('hub-recent-title');
    if (hubRecentTitle) hubRecentTitle.textContent = engine.tr('HUB_RECENT_TITLE');
    const btnHubQuickSaisie = document.getElementById('btn-hub-quick-saisie');
    if (btnHubQuickSaisie) btnHubQuickSaisie.textContent = engine.tr('SAISIE_EXPRESS') || '+ SAISIE EXPRESS';

    const hubThDate = document.getElementById('hub-th-date');
    if (hubThDate) hubThDate.textContent = engine.tr('TH_DATE');
    const hubThOp = document.getElementById('hub-th-op');
    if (hubThOp) hubThOp.textContent = engine.tr('TH_OPERATION_ACTIF');
    const hubThAmt = document.getElementById('hub-th-amt');
    if (hubThAmt) hubThAmt.textContent = engine.tr('TH_MONTANT');

    const btnRep = document.getElementById('lbl-btn-report');
    if (btnRep) btnRep.textContent = engine.tr('BTN_REPORT');
    const btnBkp = document.getElementById('lbl-btn-backup');
    if (btnBkp) btnBkp.textContent = engine.tr('BTN_BACKUP');
    const btnRst = document.getElementById('lbl-btn-restore');
    if (btnRst) btnRst.textContent = engine.tr('BTN_RESTORE');

    // 11. Saisie (Entry Form) Labels
    const txTitle = document.getElementById('lbl-tx-screen-title');
    if (txTitle) {
      const icon = txTitle.querySelector('svg');
      txTitle.textContent = '';
      if (icon) txTitle.appendChild(icon);
      txTitle.appendChild(document.createTextNode(' ' + engine.tr('SAISIE_T')));
    }
    document.querySelectorAll('#tx-type-selector .type-segment-btn').forEach(btn => {
      if (btn.dataset.type === 'REVENU') btn.textContent = engine.tr('TYPE_REVENU');
      else if (btn.dataset.type === 'DEPENSE') btn.textContent = engine.tr('TYPE_DEPENSE');
      else if (btn.dataset.type === 'TRANSFERT') btn.textContent = engine.tr('TYPE_VIREMENT');
    });
    const lblFormAmt = document.getElementById('lbl-form-montant');
    if (lblFormAmt) lblFormAmt.textContent = engine.tr('FIELD_MONTANT');
    const lblFormDate = document.getElementById('lbl-form-date');
    if (lblFormDate) lblFormDate.textContent = engine.tr('FIELD_DATE');
    const lblFormCpt = document.getElementById('lbl-form-compte');
    if (lblFormCpt) lblFormCpt.textContent = engine.tr('FIELD_ACTIFS');
    const lblFormCat = document.getElementById('lbl-form-categorie');
    if (lblFormCat) lblFormCat.textContent = engine.tr('FIELD_CATEGORIE');
    const lblFormTiers = document.getElementById('lbl-form-tiers');
    if (lblFormTiers) lblFormTiers.textContent = engine.tr('FIELD_TIERS');
    const lblFormNotes = document.getElementById('lbl-form-notes');
    if (lblFormNotes) lblFormNotes.textContent = engine.tr('FIELD_NOTES');
    const lblBtnValider = document.getElementById('lbl-btn-valider');
    if (lblBtnValider) lblBtnValider.textContent = engine.tr('BTN_VALIDATE_TX');

    // 12. Express Saisie Modal Labels
    const expTitle = document.getElementById('lbl-express-modal-title');
    if (expTitle) expTitle.textContent = engine.tr('SAISIE_EXPRESS_TITLE');
    const expSub = document.getElementById('lbl-express-modal-sub');
    if (expSub) expSub.textContent = engine.tr('SAISIE_EXPRESS_SUB');
    document.querySelectorAll('#exp-type-selector .type-segment-btn').forEach(btn => {
      if (btn.dataset.type === 'REVENU') btn.textContent = engine.tr('TYPE_REVENU');
      else if (btn.dataset.type === 'DEPENSE') btn.textContent = engine.tr('TYPE_DEPENSE');
      else if (btn.dataset.type === 'TRANSFERT') btn.textContent = engine.tr('TYPE_VIREMENT');
    });
    const q1Title = document.getElementById('q1-title') || document.getElementById('lbl-q1-title');
    if (q1Title) q1Title.textContent = engine.tr('Q1_TITLE');
    const q2Title = document.getElementById('q2-title') || document.getElementById('lbl-q2-title');
    if (q2Title) q2Title.textContent = engine.tr('Q2_TITLE');
    const q3Title = document.getElementById('q3-title') || document.getElementById('lbl-q3-title');
    if (q3Title) q3Title.textContent = engine.tr('Q3_TITLE');
    const q4Title = document.getElementById('q4-title') || document.getElementById('lbl-q4-title');
    if (q4Title) q4Title.textContent = engine.tr('Q4_TITLE');
    const q5Title = document.getElementById('q5-title') || document.getElementById('lbl-q5-title');
    if (q5Title) q5Title.textContent = engine.tr('Q5_TITLE');
    const lblBtnExp = document.getElementById('lbl-btn-save-express') || document.getElementById('lbl-btn-submit-express');
    if (lblBtnExp) lblBtnExp.textContent = engine.tr('BTN_SAVE_EXPRESS');

    // 13. Budget Allocation Modal Labels
    const allocModalTitle = document.getElementById('lbl-alloc-modal-title');
    if (allocModalTitle) allocModalTitle.textContent = engine.tr('ALLOC_MODAL_TITLE');
    const qAlloc1 = document.getElementById('q-alloc-1');
    if (qAlloc1) qAlloc1.textContent = engine.tr('Q_ALLOC_1');
    const qAlloc2 = document.getElementById('q-alloc-2');
    if (qAlloc2) qAlloc2.textContent = engine.tr('Q_ALLOC_2');
    const qAlloc3 = document.getElementById('q-alloc-3');
    if (qAlloc3) qAlloc3.textContent = engine.tr('Q_ALLOC_3');
    const lblBtnSaveAlloc = document.getElementById('lbl-btn-save-alloc');
    if (lblBtnSaveAlloc) lblBtnSaveAlloc.textContent = engine.tr('BTN_SAVE_ALLOC');
    if (typeof updateAllocMoisLocalized === 'function') updateAllocMoisLocalized();

    // 14. Audit History Modal Labels
    const histTitle = document.getElementById('lbl-history-modal-title');
    if (histTitle) histTitle.textContent = engine.tr('AUDIT_TITLE');
    const btnCloseHist = document.querySelector('#btn-close-history-footer span');
    if (btnCloseHist) btnCloseHist.textContent = engine.tr('BTN_CLOSE_AUDIT');

    // 15. Cashflow View Labels & Placeholders
    const lblKpiRev = document.getElementById('lbl-kpi-rev');
    if (lblKpiRev) lblKpiRev.textContent = engine.tr('REVENUES');
    const lblKpiDep = document.getElementById('lbl-kpi-dep');
    if (lblKpiDep) lblKpiDep.textContent = engine.tr('DEPENSES');
    const lblKpiNet = document.getElementById('lbl-kpi-net');
    if (lblKpiNet) lblKpiNet.textContent = engine.tr('NET_CASHFLOW_SHORT');
    const cfTrendTitle = document.getElementById('cf-chart-trend-title');
    if (cfTrendTitle) cfTrendTitle.textContent = engine.tr('TREND_CHART_TITLE');
    const cfTrendSub = document.getElementById('cf-chart-trend-sub');
    if (cfTrendSub) cfTrendSub.textContent = engine.tr('TREND_6M');
    const cfDonutTitle = document.getElementById('cf-chart-donut-title');
    if (cfDonutTitle) cfDonutTitle.textContent = engine.tr('DONUT_TITLE');
    const cfChipAll = document.getElementById('cf-chip-all');
    if (cfChipAll) cfChipAll.textContent = engine.tr('FILTER_ALL');
    const cfChipRev = document.getElementById('cf-chip-rev');
    if (cfChipRev) cfChipRev.textContent = engine.tr('FILTER_REV');
    const cfChipDep = document.getElementById('cf-chip-dep');
    if (cfChipDep) cfChipDep.textContent = engine.tr('FILTER_DEP');
    const cfChipTra = document.getElementById('cf-chip-tra');
    if (cfChipTra) cfChipTra.textContent = engine.tr('FILTER_TRA');
    const lblCfSort = document.getElementById('lbl-cf-sort');
    if (lblCfSort) lblCfSort.textContent = engine.tr('SORT_PREFIX');
    const lblCfTable = document.getElementById('lbl-cf-table-title');
    if (lblCfTable) lblCfTable.textContent = engine.tr('TX_CONSOLIDATED');
    const cfThDate = document.getElementById('cf-th-date');
    if (cfThDate) cfThDate.textContent = engine.tr('TH_DATE');
    const cfThCat = document.getElementById('cf-th-cat');
    if (cfThCat) cfThCat.textContent = engine.tr('TH_CATEGORIE_COMPTE');
    const cfThTiers = document.getElementById('cf-th-tiers');
    if (cfThTiers) cfThTiers.textContent = engine.tr('TH_TIERS');
    const cfThAmt = document.getElementById('cf-th-amt');
    if (cfThAmt) cfThAmt.textContent = engine.tr('TH_MONTANT');
    const txSearchInput = document.getElementById('input-tx-search');
    if (txSearchInput) txSearchInput.placeholder = engine.tr('SEARCH_TX_PLACEHOLDER');

    // 16. Net Worth View Labels
    const lblNwTitle = document.getElementById('lbl-networth-title');
    if (lblNwTitle) lblNwTitle.textContent = engine.tr('NETWORTH_TITLE');
    const lblNwAssets = document.getElementById('lbl-nw-assets');
    if (lblNwAssets) lblNwAssets.textContent = engine.tr('TOTAL_ASSETS');
    const lblNwDebts = document.getElementById('lbl-nw-debts');
    if (lblNwDebts) lblNwDebts.textContent = engine.tr('TOTAL_DEBTS');
    const nwChipAll = document.getElementById('nw-chip-all');
    if (nwChipAll) nwChipAll.textContent = engine.tr('FILTER_ALL');
    const nwChipLiq = document.getElementById('nw-chip-liq');
    if (nwChipLiq) nwChipLiq.textContent = engine.tr('LIQUIDITIES');
    const nwChipInv = document.getElementById('nw-chip-inv');
    if (nwChipInv) nwChipInv.textContent = engine.tr('INVESTMENTS');
    const nwChipDet = document.getElementById('nw-chip-det');
    if (nwChipDet) nwChipDet.textContent = engine.tr('COMMITMENTS');
    const lblNwTable = document.getElementById('lbl-nw-table-title');
    if (lblNwTable) lblNwTable.textContent = engine.tr('NETWORTH_INVENTORY');
    const nwThCpt = document.getElementById('nw-th-cpt');
    if (nwThCpt) nwThCpt.textContent = engine.tr('TH_COMPTE');
    const nwThClasse = document.getElementById('nw-th-classe');
    if (nwThClasse) nwThClasse.textContent = engine.tr('TH_CLASSE');
    const nwThNative = document.getElementById('nw-th-native');
    if (nwThNative) nwThNative.textContent = engine.tr('TH_SOLDE_NATIF');
    const nwThConv = document.getElementById('nw-th-conv');
    if (nwThConv) nwThConv.textContent = engine.tr('TH_SOLDE_CONVERTI');

    // 17. Budget View Labels
    const lblBgEnvelope = document.getElementById('lbl-budget-envelope');
    if (lblBgEnvelope) lblBgEnvelope.textContent = engine.tr('BUDGET_GLOBAL_ENVELOPE');
    const btnAlloc = document.getElementById('btn-open-alloc-modal');
    if (btnAlloc) btnAlloc.textContent = engine.tr('BTN_ALLOC_BUDGET');
    const bgChipAll = document.getElementById('bg-chip-all');
    if (bgChipAll) bgChipAll.textContent = engine.tr('FILTER_ALL');
    const bgChipOver = document.getElementById('bg-chip-over');
    if (bgChipOver) bgChipOver.textContent = engine.tr('STATUS_OVER');
    const bgChipOk = document.getElementById('bg-chip-ok');
    if (bgChipOk) bgChipOk.textContent = engine.tr('STATUS_OK');
    const bgTitle = document.getElementById('lbl-budget-title');
    if (bgTitle) bgTitle.textContent = engine.tr('BUDGET_TITLE');
    const bgThCat = document.getElementById('bg-th-cat');
    if (bgThCat) bgThCat.textContent = engine.tr('TH_CATEGORIE');
    const bgThAlloue = document.getElementById('bg-th-alloue');
    if (bgThAlloue) bgThAlloue.textContent = engine.tr('TH_ALLOUE');
    const bgThConsomme = document.getElementById('bg-th-consomme');
    if (bgThConsomme) bgThConsomme.textContent = engine.tr('TH_CONSOMME');
    const bgThConso = document.getElementById('bg-th-conso');
    if (bgThConso) bgThConso.textContent = engine.tr('TH_STATUT');

    // 17.5 Always re-populate all form and modal selects in the active language
    populateSelects();

    // 17.6 Update all .btn-field-add buttons to active language
    updateFieldAddButtonsTranslations();

    // 17.65 Update Dynamic Entity Extension Gatekeeper Modal to active language
    updateAddEntityModalTranslations();

    // 17.7 Update Live Rates Refresh Button Tooltip & Accessibility
    updateLiveRatesDisplay();

    // 18. Dispatch to Current View
    if (engine.currentView === 'home') renderHomeView();
    else if (engine.currentView === 'cashflow') renderCashflowView();
    else if (engine.currentView === 'networth') renderNetWorthView();
    else if (engine.currentView === 'budget') renderBudgetView();
    else if (engine.currentView === 'settings') renderSettingsView();

    // 19. Real-Time Refresh for Open Modals
    if (modalReport && modalReport.classList.contains('open')) {
      renderFinancialReportModal();
    }
  }

  function renderHomeView() {
    const data = engine.getConsolidatedHubData();

    // Update active filter chip state
    document.querySelectorAll('#hub-filter-bar .hub-filter-chip').forEach(chip => {
      if (chip.dataset.hubFilter === engine.hubFilter) chip.classList.add('active');
      else chip.classList.remove('active');
    });

    // Consolidated Portfolio Hero Card
    const heroNw = document.getElementById('hub-hero-networth');
    if (heroNw) heroNw.textContent = engine.formatMoney(data.netWorth);

    const heroCurr = document.getElementById('hub-hero-currency');
    if (heroCurr) heroCurr.textContent = engine.currency;

    const hubCfVal = document.getElementById('hub-val-cashflow');
    if (hubCfVal) {
      hubCfVal.textContent = engine.formatMoney(data.netCashflow);
      hubCfVal.style.color = data.netCashflow >= 0 ? '#10B981' : '#EF4444';
    }

    const hubBgVal = document.getElementById('hub-val-budget');
    if (hubBgVal) {
      hubBgVal.textContent = `${data.consumptionRate}%`;
      hubBgVal.style.color = data.consumptionRate > 100 ? '#EF4444' : data.consumptionRate >= 80 ? '#F59E0B' : '#10B981';
    }

    const hubAssetsVal = document.getElementById('hub-val-assets');
    if (hubAssetsVal) hubAssetsVal.textContent = engine.formatMoney(data.totalAssets);

    const hubDebtsVal = document.getElementById('hub-val-debts');
    if (hubDebtsVal) hubDebtsVal.textContent = engine.formatMoney(data.totalLiabilities);

    const solvencyEl = document.getElementById('hub-solvency-ratio');
    if (solvencyEl) {
      const sol = data.totalAssets > 0 ? Math.max(0, 100 - data.debtRatio) : (data.totalLiabilities > 0 ? 0 : 100);
      if (!data.hasData && data.totalAssets === 0 && data.totalLiabilities === 0) {
        solvencyEl.textContent = `${engine.tr('SOLVENCY_SCORE', { sol, score: data.healthScore })} • ${engine.tr('SOLVENCY_EMPTY')}`;
      } else {
        solvencyEl.textContent = engine.tr('SOLVENCY_SCORE', { sol, score: data.healthScore });
      }
    }

    // Mini stack bar: Actifs vs Dettes
    const hubStack = document.getElementById('hub-asset-stack-bar');
    if (hubStack) {
      const totalPat = data.totalAssets + data.totalLiabilities;
      if (totalPat > 0) {
        const aPct = Math.round((data.totalAssets / totalPat) * 100);
        const dPct = Math.max(0, 100 - aPct);
        hubStack.innerHTML = `
          <div class="stack-segment" style="width:${aPct}%; background:#10B981;" title="Actifs : ${aPct}%"></div>
          <div class="stack-segment" style="width:${dPct}%; background:#EF4444;" title="Dettes : ${dPct}%"></div>
        `;
      } else {
        hubStack.innerHTML = `<div class="stack-segment" style="width:100%; background:rgba(255,255,255,0.1);"></div>`;
      }
    }

    // Card Visibility Filtering by Hub Filter Chip
    const cardTrend = document.getElementById('hub-card-trend');
    const cardDonut = document.getElementById('hub-card-donut');
    const cardPatrimoine = document.getElementById('hub-card-patrimoine');
    const cardBudget = document.getElementById('hub-card-budget');
    const cardRecent = document.getElementById('hub-card-recent-tx');

    if (cardTrend) cardTrend.style.display = (engine.hubFilter === 'ALL' || engine.hubFilter === 'CASHFLOW') ? 'block' : 'none';
    if (cardDonut) cardDonut.style.display = (engine.hubFilter === 'ALL' || engine.hubFilter === 'CASHFLOW' || engine.hubFilter === 'BUDGET') ? 'block' : 'none';
    if (cardPatrimoine) cardPatrimoine.style.display = (engine.hubFilter === 'ALL' || engine.hubFilter === 'PATRIMOINE') ? 'block' : 'none';
    if (cardBudget) cardBudget.style.display = (engine.hubFilter === 'ALL' || engine.hubFilter === 'BUDGET') ? 'block' : 'none';
    if (cardRecent) cardRecent.style.display = 'block';

    // Section 1: Cashflow 6-Month Trend Chart
    const trendPeriod = document.getElementById('hub-trend-period');
    if (trendPeriod) trendPeriod.textContent = `${engine.getMonthLabel()} ${engine.year}`;
    renderTrendChart('hub-trend-chart-box');

    // Section 2: Expense Category Donut
    const donutPeriod = document.getElementById('hub-donut-period');
    if (donutPeriod) donutPeriod.textContent = `${engine.getMonthLabel()} ${engine.year}`;

    const hubFilterBadgeEl = document.getElementById('hub-donut-active-filter');
    if (hubFilterBadgeEl) {
      if (engine.selectedCategoryFilter) {
        hubFilterBadgeEl.style.display = 'inline-flex';
        hubFilterBadgeEl.className = 'active-filter-badge';
        hubFilterBadgeEl.innerHTML = `
          <span>${engine.tr('TH_CATEGORIE')} : <strong>${engine.trEntity(engine.selectedCategoryFilter)}</strong></span>
          <button type="button" class="clear-btn" id="btn-hub-clear-cat-filter" title="${engine.tr('CAT_FILTER_CLEAR') || 'Effacer le filtre'}">✕</button>
        `;
        document.getElementById('btn-hub-clear-cat-filter')?.addEventListener('click', () => {
          engine.selectedCategoryFilter = null;
          renderHomeView();
        });
      } else {
        hubFilterBadgeEl.style.display = 'none';
      }
    }

    renderDonutChart('hub-donut-chart-box', data.breakdown, data.totalDepenses);

    // Section 3: Asset Allocation & Top Assets (Clean header without redundant Net Assets clutter)
    const patTotal = document.getElementById('hub-patrimoine-total');
    if (patTotal) {
      patTotal.textContent = '';
      patTotal.style.display = 'none';
    }

    const patBarsBox = document.getElementById('hub-patrimoine-bars-box');
    if (patBarsBox) {
      const liq = (data.allocation && data.allocation.LIQUIDITE) || 0;
      const inv = (data.allocation && data.allocation.INVESTISSEMENT) || 0;
      const det = data.totalLiabilities || 0;
      const totalAlloc = liq + inv + det;
      const pLiq = totalAlloc > 0 ? Math.round((liq / totalAlloc) * 100) : 0;
      const pInv = totalAlloc > 0 ? Math.round((inv / totalAlloc) * 100) : 0;
      const pDet = totalAlloc > 0 ? Math.max(0, 100 - pLiq - pInv) : 0;

      patBarsBox.innerHTML = `
        <div style="display:flex; height:12px; border-radius:6px; overflow:hidden; background:rgba(255,255,255,0.08); margin-bottom:10px;">
          <div style="width:${pLiq}%; background:#10B981; transition:width 0.4s ease;" title="${engine.tr('LIQUIDITIES')}: ${pLiq}%"></div>
          <div style="width:${pInv}%; background:#4169E1; transition:width 0.4s ease;" title="${engine.tr('INVESTMENTS')}: ${pInv}%"></div>
          <div style="width:${pDet}%; background:#EF4444; transition:width 0.4s ease;" title="${engine.tr('COMMITMENTS')}: ${pDet}%"></div>
        </div>
        <div style="display:grid; grid-template-columns:repeat(3, 1fr); gap:8px; font-size:0.72rem;">
          <div style="background:rgba(16,185,129,0.1); border:1px solid rgba(16,185,129,0.3); border-radius:8px; padding:6px 8px;">
            <div style="color:#A7F3D0; font-weight:700;">${engine.tr('LIQUIDITIES')}</div>
            <div style="font-weight:900; color:#FFF; font-size:0.8rem; margin-top:2px; font-feature-settings:'tnum';">${engine.formatMoney(liq)}</div>
            <div style="font-size:0.65rem; color:#94A3B8;">${engine.tr('OF_TOTAL', { pct: pLiq })}</div>
          </div>
          <div style="background:rgba(65,105,225,0.1); border:1px solid rgba(65,105,225,0.3); border-radius:8px; padding:6px 8px;">
            <div style="color:#93C5FD; font-weight:700;">${engine.tr('INVESTMENTS')}</div>
            <div style="font-weight:900; color:#FFF; font-size:0.8rem; margin-top:2px; font-feature-settings:'tnum';">${engine.formatMoney(inv)}</div>
            <div style="font-size:0.65rem; color:#94A3B8;">${engine.tr('OF_TOTAL', { pct: pInv })}</div>
          </div>
          <div style="background:rgba(239,68,68,0.1); border:1px solid rgba(239,68,68,0.3); border-radius:8px; padding:6px 8px;">
            <div style="color:#FECACA; font-weight:700;">${engine.tr('COMMITMENTS')}</div>
            <div style="font-weight:900; color:#FFF; font-size:0.8rem; margin-top:2px; font-feature-settings:'tnum';">${engine.formatMoney(det)}</div>
            <div style="font-size:0.65rem; color:#94A3B8;">${engine.tr('OF_TOTAL', { pct: pDet })}</div>
          </div>
        </div>
      `;
    }

    const topAssetsList = document.getElementById('hub-top-assets-list');
    if (topAssetsList) {
      if (!data.topAssets || data.topAssets.length === 0) {
        topAssetsList.innerHTML = `
          <div class="empty-state-rich">
            <div class="empty-state-icon-box" aria-hidden="true">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 2L2 7l10 5 10-5-10-5z"/><path d="M2 17l10 5 10-5"/><path d="M2 12l10 5 10-5"/></svg>
            </div>
            <div class="empty-state-title">${escapeHTML(engine.tr('NO_ASSETS'))}</div>
            <p class="empty-state-desc">${escapeHTML(engine.tr('EMPTY_PATRIMOINE_SUB'))}</p>
            <div class="empty-state-actions">
              <button type="button" class="btn-empty-cta" data-empty-action="add-tx">${escapeHTML(engine.tr('EMPTY_CTA_ADD_TX'))}</button>
              <button type="button" class="btn-empty-cta secondary" data-empty-action="add-account">${escapeHTML(engine.tr('EMPTY_CTA_ADD_ACC'))}</button>
            </div>
          </div>
        `;
      } else {
        topAssetsList.innerHTML = data.topAssets.map(ast => `
          <div class="hub-asset-row" style="display:flex; justify-content:space-between; align-items:center; background:rgba(255,255,255,0.03); border:1px solid rgba(255,255,255,0.07); border-radius:8px; padding:8px 12px;">
            <div style="display:flex; align-items:center; gap:8px;">
              <span class="badge-classe ${escapeHTML(ast.classe)}" style="font-size:0.62rem; padding:2px 6px;">${escapeHTML(ast.classe)}</span>
              <span style="font-size:0.82rem; font-weight:700; color:#FFF;">${escapeHTML(ast.nom)}</span>
            </div>
            <span style="font-size:0.85rem; font-weight:900; color:var(--c-yellow-royal); font-feature-settings:'tnum';">${escapeHTML(ast.solde)}</span>
          </div>
        `).join('');
      }
    }

    // Section 4: Budget Envelopes Progress Gauges
    const budgetPctEl = document.getElementById('hub-budget-global-pct');
    if (budgetPctEl) {
      budgetPctEl.textContent = `${data.consumptionRate}% ${engine.tr('TH_STATUT')}`;
      budgetPctEl.style.color = data.consumptionRate > 100 ? '#EF4444' : data.consumptionRate >= 80 ? '#F59E0B' : '#10B981';
    }

    const budgetGaugesBox = document.getElementById('hub-budget-gauges-box');
    if (budgetGaugesBox) {
      if (!data.topBudgetCategories || data.topBudgetCategories.length === 0) {
        budgetGaugesBox.innerHTML = `
          <div class="empty-state-rich">
            <div class="empty-state-icon-box" aria-hidden="true">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="6"/><circle cx="12" cy="12" r="2"/></svg>
            </div>
            <div class="empty-state-title">${escapeHTML(engine.tr('NO_BUDGET_PERIOD'))}</div>
            <p class="empty-state-desc">${escapeHTML(engine.tr('EMPTY_BUDGET_SUB'))}</p>
            <div class="empty-state-actions">
              <button type="button" class="btn-empty-cta" data-empty-action="open-alloc">${escapeHTML(engine.tr('EMPTY_CTA_ALLOC_BG'))}</button>
              ${engine.month !== '00' ? `<button type="button" class="btn-empty-cta secondary" data-empty-action="view-full-year">${escapeHTML(engine.tr('EMPTY_CTA_FULL_YEAR'))}</button>` : ''}
            </div>
          </div>
        `;
      } else {
        budgetGaugesBox.innerHTML = data.topBudgetCategories.map(bg => {
          const color = bg.rawPct > 100 ? '#EF4444' : bg.rawPct >= 80 ? '#F59E0B' : '#10B981';
          const badgeClass = bg.rawPct > 100 ? 'status-over' : bg.rawPct >= 80 ? 'status-warn' : 'status-ok';
          const badgeText = bg.rawPct > 100 ? engine.tr('STATUS_OVER') : bg.rawPct >= 80 ? engine.tr('STATUS_WARN') : engine.tr('STATUS_OK');
          return `
            <div class="hub-budget-row" style="background:rgba(255,255,255,0.02); border:1px solid rgba(255,255,255,0.06); border-radius:10px; padding:10px 12px;">
              <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
                <span style="font-size:0.8rem; font-weight:700; color:#FFF;">${escapeHTML(engine.trEntity(bg.nom))}</span>
                <span class="badge-status ${badgeClass}" style="font-size:0.62rem; padding:2px 8px;">${escapeHTML(badgeText)} (${Math.round(bg.rawPct)}%)</span>
              </div>
              <div style="height:6px; background:rgba(255,255,255,0.08); border-radius:3px; overflow:hidden; margin-bottom:6px;">
                <div style="width:${Math.min(100, bg.rawPct)}%; height:100%; background:${color}; transition:width 0.4s ease;"></div>
              </div>
              <div style="display:flex; justify-content:space-between; font-size:0.68rem; color:var(--text-secondary);">
                <span>${escapeHTML(engine.tr('CONSUMED'))} <strong style="color:#FFF; font-feature-settings:'tnum';">${escapeHTML(engine.formatMoney(bg.consomme))}</strong></span>
                <span>${escapeHTML(engine.tr('ALLOCATED'))} <strong style="color:var(--c-yellow-royal); font-feature-settings:'tnum';">${escapeHTML(engine.formatMoney(bg.alloue))}</strong></span>
              </div>
            </div>
          `;
        }).join('');
      }
    }

    // Section 5: Recent Transactions Table
    const recentTxBody = document.getElementById('hub-recent-tx-body');
    if (recentTxBody) {
      if (!data.recentTransactions || data.recentTransactions.length === 0) {
        const isCatFiltered = Boolean(engine.selectedCategoryFilter);
        recentTxBody.innerHTML = `
          <tr>
            <td colspan="3" class="empty-state">
              <div class="empty-state-rich">
                <div class="empty-state-icon-box" aria-hidden="true">
                  <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>
                </div>
                <div class="empty-state-title">${escapeHTML(engine.tr('NO_TX'))}</div>
                <p class="empty-state-desc">${escapeHTML(engine.tr('EMPTY_TX_SUB'))}</p>
                <div class="empty-state-actions">
                  <button type="button" class="btn-empty-cta" data-empty-action="add-tx">${escapeHTML(engine.tr('EMPTY_CTA_ADD_TX'))}</button>
                  ${isCatFiltered ? `<button type="button" class="btn-empty-cta secondary" data-empty-action="reset-cf-filters">${escapeHTML(engine.tr('EMPTY_CTA_RESET_FILTER'))}</button>` : ''}
                </div>
              </div>
            </td>
          </tr>
        `;
      } else {
        recentTxBody.innerHTML = data.recentTransactions.map(tx => {
          const color = tx.type === 'REVENU' ? '#10B981' : tx.type === 'VIREMENT' ? '#60A5FA' : '#EF4444';
          const sign = tx.type === 'REVENU' ? '+' : tx.type === 'DEPENSE' ? '-' : '⇄';
          return `
            <tr>
              <td>
                <div style="font-weight:700; font-size:0.78rem; color:#FFF;">${escapeHTML(tx.date)}</div>
                <div style="font-size:0.65rem; color:var(--text-muted);">#${escapeHTML(tx.id)}</div>
              </td>
              <td>
                <div style="font-weight:700; font-size:0.8rem; color:#FFF;">${escapeHTML(tx.categorie)}</div>
                <div style="font-size:0.68rem; color:var(--text-secondary);">${escapeHTML(tx.compte)}</div>
              </td>
              <td style="text-align:right;">
                <div style="font-weight:900; font-size:0.85rem; color:${color}; font-feature-settings:'tnum';">${sign} ${escapeHTML(tx.montant)}</div>
              </td>
            </tr>
          `;
        }).join('');
      }
    }

    // Update Live Forex Rates Status
    updateLiveRatesDisplay();
  }

  function renderCashflowView() {
    const cf = engine.computeCashflow();

    document.getElementById('kpi-cashflow-rev').textContent = engine.formatMoney(cf.totalRevenus);
    document.getElementById('kpi-cashflow-dep').textContent = engine.formatMoney(cf.totalDepenses);
    
    const netEl = document.getElementById('kpi-cashflow-net');
    const netCard = document.getElementById('card-kpi-net');
    netEl.textContent = engine.formatMoney(cf.netCashflow);
    if (cf.netCashflow >= 0) {
      netCard.className = 'kpi-card kpi-mini-card kpi-net kpi-net-card positive';
    } else {
      netCard.className = 'kpi-card kpi-mini-card kpi-net kpi-net-card negative';
    }

    const lblCfSubRev = document.getElementById('lbl-cf-sub-rev');
    if (lblCfSubRev) lblCfSubRev.textContent = engine.tr('SUB_INFLOW') || 'Flux Entrant';
    const lblCfSubDep = document.getElementById('lbl-cf-sub-dep');
    if (lblCfSubDep) lblCfSubDep.textContent = engine.tr('SUB_OUTFLOW') || 'Flux Sortant';
    const lblCfSubNet = document.getElementById('lbl-cf-sub-net');
    if (lblCfSubNet) lblCfSubNet.textContent = cf.netCashflow >= 0 ? (engine.tr('SUB_SURPLUS') || 'Excédent Net') : (engine.tr('SUB_DEFICIT') || 'Déficit Net');

    // Render Dynamic SVG Charts
    renderTrendChart('svg-trend-chart-box');
    renderDonutChart('svg-donut-chart-box', cf.breakdown, cf.totalDepenses);

    const donutPeriodLabel = document.getElementById('donut-period-label');
    if (donutPeriodLabel) {
      donutPeriodLabel.textContent = `${engine.getMonthLabel()} ${engine.year}`;
    }

    // Transactions Count
    document.getElementById('cashflow-count-label').textContent = engine.tr('FLUX_COUNT', { count: cf.transactions.length });

    // Active Category Filter Indicator Badge
    const filterBadgeEl = document.getElementById('cf-active-filter-badge');
    if (filterBadgeEl) {
      if (engine.selectedCategoryFilter) {
        filterBadgeEl.style.display = 'inline-flex';
        filterBadgeEl.className = 'active-filter-badge';
        filterBadgeEl.innerHTML = `
          <span>${engine.tr('TH_CATEGORIE')} : <strong>${engine.trEntity(engine.selectedCategoryFilter)}</strong></span>
          <button type="button" class="clear-btn" id="btn-clear-cat-filter" title="${engine.tr('CAT_FILTER_CLEAR')}">✕</button>
        `;
        document.getElementById('btn-clear-cat-filter')?.addEventListener('click', () => {
          engine.selectedCategoryFilter = null;
          renderCashflowView();
        });
      } else {
        filterBadgeEl.style.display = 'none';
      }
    }

    // Render Modern Table
    const tbody = document.getElementById('cashflow-table-body');
    if (cf.transactions.length === 0) {
      const hasActiveFilters = Boolean(engine.searchTerm || engine.fluxFilter !== 'ALL' || engine.selectedCategoryFilter || engine.accountFilter !== 'ALL');
      tbody.innerHTML = `
        <tr>
          <td colspan="5" class="empty-state">
            <div class="empty-state-rich">
              <div class="empty-state-icon-box" aria-hidden="true">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
              </div>
              <div class="empty-state-title">${escapeHTML(engine.tr('EMPTY_TX_FILTER'))}</div>
              <p class="empty-state-desc">${escapeHTML(engine.tr('EMPTY_TX_SUB'))}</p>
              <div class="empty-state-actions">
                <button type="button" class="btn-empty-cta" data-empty-action="add-tx">${escapeHTML(engine.tr('EMPTY_CTA_ADD_TX'))}</button>
                ${hasActiveFilters ? `<button type="button" class="btn-empty-cta secondary" data-empty-action="reset-cf-filters">${escapeHTML(engine.tr('EMPTY_CTA_RESET_FILTER'))}</button>` : ''}
                ${engine.month !== '00' ? `<button type="button" class="btn-empty-cta secondary" data-empty-action="view-full-year">${escapeHTML(engine.tr('EMPTY_CTA_FULL_YEAR'))}</button>` : ''}
              </div>
            </div>
          </td>
        </tr>
      `;
      return;
    }

    tbody.innerHTML = cf.transactions.map(tx => `
      <tr>
        <td>
          <div style="font-weight:700;">${escapeHTML(tx.date)}</div>
          <div style="font-size:0.68rem; opacity:0.8;">#${escapeHTML(tx.id)}</div>
        </td>
        <td>
          <div style="font-weight:700;">${escapeHTML(tx.nomCategorieTraduit || engine.trEntity(tx.nomCategorie))}</div>
          <div style="font-size:0.7rem; opacity:0.8;">${escapeHTML(tx.nomCompteTraduit || engine.cleanAssetName(tx.nomCompte))}</div>
        </td>
        <td>
          <div>${escapeHTML(tx.nomTiersTraduit || engine.trEntity(tx.nomTiers))}</div>
          <span class="badge-flux ${escapeHTML(tx.typeFlux)}">${escapeHTML(engine.trEntity(tx.typeFlux))}</span>
        </td>
        <td class="td-amount">
          <div>${escapeHTML(engine.formatMoney(tx.montantConverti))}</div>
          ${tx.devise !== engine.currency ? `<div style="font-size:0.68rem; opacity:0.75;">(${escapeHTML(tx.montant)} ${escapeHTML(tx.devise)})</div>` : ''}
        </td>
        <td style="text-align:center;">
          <button class="btn-del-tx" data-id="${escapeHTML(tx.id)}" aria-label="Supprimer" title="Supprimer">
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <polyline points="3 6 5 6 21 6"/>
              <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/>
            </svg>
          </button>
        </td>
      </tr>
    `).join('');

    // Bind delete buttons
    tbody.querySelectorAll('.btn-del-tx').forEach(btn => {
      btn.addEventListener('click', async () => {
        triggerHaptic(20);
        const id = btn.dataset.id;
        const confirmed = await showConfirmDialog({
          title: 'Supprimer l\'Opération',
          message: `Confirmez-vous la suppression définitive de l'opération #${id} ?`,
          confirmText: 'Supprimer',
          cancelText: 'Annuler',
          type: 'danger'
        });
        if (confirmed) {
          engine.deleteTransaction(id);
          showToast(`Opération #${id} supprimée`, 'info');
        }
      });
    });
  }

  function renderNetWorthView() {
    const nw = engine.computeNetWorth();

    const heroCard = document.getElementById('networth-hero-card');
    const heroVal = document.getElementById('nw-hero-val');
    heroVal.textContent = engine.formatMoney(nw.netWorth);

    if (heroCard) {
      heroCard.className = 'kpi-row kpi-networth-row ' + (nw.netWorth >= 0 ? 'positive' : 'negative');
    }
    const nwCard = document.getElementById('card-nw-networth');
    if (nwCard) {
      nwCard.className = 'kpi-card kpi-networth-card ' + (nw.netWorth >= 0 ? 'positive' : 'negative');
    }

    const lblNwCard = document.getElementById('lbl-nw-card-networth');
    if (lblNwCard) lblNwCard.textContent = engine.tr('NET_WORTH') || 'VALEUR NETTE';
    const lblNwSub = document.getElementById('lbl-nw-sub-networth');
    if (lblNwSub) lblNwSub.textContent = engine.tr('SUB_NET_WORTH') || 'Patrimoine Net';

    const lblAssetsCard = document.getElementById('lbl-nw-card-assets');
    if (lblAssetsCard) lblAssetsCard.textContent = engine.tr('TOTAL_ASSETS') || 'TOTAL ACTIFS';
    const lblAssetsSub = document.getElementById('lbl-nw-sub-assets');
    if (lblAssetsSub) lblAssetsSub.textContent = engine.tr('SUB_GROSS_ASSETS') || 'Actifs Bruts';

    const lblDebtsCard = document.getElementById('lbl-nw-card-debts');
    if (lblDebtsCard) lblDebtsCard.textContent = engine.tr('TOTAL_DEBTS') || 'TOTAL DETTES';
    const lblDebtsSub = document.getElementById('lbl-nw-sub-debts');
    if (lblDebtsSub) lblDebtsSub.textContent = engine.tr('SUB_LIABILITIES') || 'Engagements';

    const lblAllocTitle = document.getElementById('lbl-nw-alloc-title');
    if (lblAllocTitle) lblAllocTitle.textContent = engine.tr('ALLOC_TITLE') || 'RÉPARTITION & STRUCTURE DU PATRIMOINE';

    document.getElementById('nw-hero-assets').textContent = engine.formatMoney(nw.totalAssets);
    document.getElementById('nw-hero-liabilities').textContent = engine.formatMoney(nw.totalLiabilities);

    // Translated As of Month Label
    const asOfLabel = document.getElementById('lbl-as-of-date');
    if (asOfLabel) {
      asOfLabel.textContent = engine.tr('AS_OF_MONTH', {
        month: engine.getMonthLabel(),
        year: engine.year
      });
    }

    // Asset Allocation Stack Bar
    const stackBar = document.getElementById('nw-asset-stack-bar');
    const stackLegend = document.getElementById('nw-asset-legend');
    if (stackBar && nw.totalAssets > 0) {
      const otherAssets = Math.max(0, engine.roundMoney(nw.totalAssets - (nw.allocation.LIQUIDITE || 0) - (nw.allocation.INVESTISSEMENT || 0)));
      const liqPct = Math.round(((nw.allocation.LIQUIDITE || 0) / nw.totalAssets) * 100);
      const invPct = (otherAssets === 0)
        ? Math.max(0, 100 - liqPct)
        : Math.round(((nw.allocation.INVESTISSEMENT || 0) / nw.totalAssets) * 100);
      const restPct = Math.max(0, 100 - liqPct - invPct);

      stackBar.innerHTML = `
        <div class="stack-segment" style="width:${liqPct}%; background:#10B981;" title="${engine.tr('LIQUIDITIES')} : ${liqPct}%"></div>
        <div class="stack-segment" style="width:${invPct}%; background:#4169E1;" title="${engine.tr('INVESTMENTS')} : ${invPct}%"></div>
        <div class="stack-segment" style="width:${restPct}%; background:#FADA5E;" title="${engine.tr('OTHERS')} : ${restPct}%"></div>
      `;

      if (stackLegend) {
        stackLegend.innerHTML = `
          <span><span style="color:#10B981;">●</span> ${engine.tr('LIQUIDITIES')} (${liqPct}%)</span>
          <span><span style="color:#4169E1;">●</span> ${engine.tr('INVESTMENTS')} (${invPct}%)</span>
          <span><span style="color:#FADA5E;">●</span> ${engine.tr('OTHERS')} (${restPct}%)</span>
        `;
      }
    } else if (stackBar) {
      stackBar.innerHTML = `<div class="stack-segment" style="width:100%; background:linear-gradient(90deg, rgba(212,175,55,0.25) 0%, rgba(250,218,94,0.4) 100%);"></div>`;
      if (stackLegend) {
        stackLegend.innerHTML = `
          <span style="display:inline-flex; align-items:center; gap:8px; flex-wrap:wrap;">
            <span style="color:var(--c-gold-light); font-weight:600;">● ${escapeHTML(engine.tr('NW_STACK_EMPTY_LABEL'))}</span>
            <button type="button" class="btn-empty-cta" style="padding:4px 10px; font-size:0.66rem;" data-empty-action="add-tx">${escapeHTML(engine.tr('EMPTY_CTA_ADD_TX'))}</button>
          </span>
        `;
      }
    }

    // Modern Accounts & Liabilities Inventory Table Header & Readiness Callout
    const asOfDateEl = document.getElementById('lbl-as-of-date');
    if (asOfDateEl) {
      asOfDateEl.textContent = `${engine.tr('AS_OF') || 'Au'} ${engine.systemDate}`;
    }
    const accBadge = document.getElementById('nw-accounts-summary-badge');
    if (accBadge) {
      accBadge.textContent = (nw.accounts.length === 0)
        ? `0 ${engine.tr('NW_ACCOUNTS_COUNT')}`
        : `${nw.accounts.length} ${engine.tr('NW_ACCOUNTS_COUNT')}`;
    }
    const btnAddAcc = document.getElementById('btn-nw-add-account');
    if (btnAddAcc) {
      btnAddAcc.textContent = engine.tr('BTN_NW_ADD_ACCOUNT');
    }

    const readinessSlot = document.getElementById('nw-table-readiness-slot');
    if (readinessSlot) {
      readinessSlot.innerHTML = '';
    }

    // Accounts Modern Table Body
    const tbody = document.getElementById('networth-table-body');
    if (nw.accounts.length === 0) {
      const hasClassFilter = (engine.patrimoineClassFilter && engine.patrimoineClassFilter !== 'ALL');
      tbody.innerHTML = `
        <tr>
          <td colspan="4" class="empty-state">
            <div class="empty-state-rich">
              <div class="empty-state-icon-box" aria-hidden="true">
                <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 2L2 7l10 5 10-5-10-5z"/><path d="M2 17l10 5 10-5"/><path d="M2 12l10 5 10-5"/></svg>
              </div>
              <div class="empty-state-title">${escapeHTML(hasClassFilter ? engine.tr('NO_ACCOUNT_MATCH') : (engine.tr('NO_ACCOUNT_OR_TX') || engine.tr('NO_ACCOUNT_MATCH')))}</div>
              <p class="empty-state-desc">${escapeHTML(engine.tr('EMPTY_PATRIMOINE_SUB'))}</p>
              <div class="empty-state-actions">
                <button type="button" class="btn-empty-cta" data-empty-action="add-tx">${escapeHTML(engine.tr('EMPTY_CTA_ADD_TX'))}</button>
                <button type="button" class="btn-empty-cta secondary" data-empty-action="add-account">${escapeHTML(engine.tr('EMPTY_CTA_ADD_ACC'))}</button>
                ${hasClassFilter ? `<button type="button" class="btn-empty-cta secondary" data-empty-action="reset-nw-filters">${escapeHTML(engine.tr('EMPTY_CTA_RESET_FILTER'))}</button>` : ''}
              </div>
            </div>
          </td>
        </tr>
      `;
      return;
    }

    tbody.innerHTML = nw.accounts.map(cpt => {
      let statusPillHtml = '';
      if (cpt.isDette) {
        if (cpt.soldeConverti === 0) {
          statusPillHtml = `<span class="cpt-status-pill clear">${escapeHTML(engine.tr('ACCOUNT_STATUS_DEBT_FREE'))}</span>`;
        } else {
          const debtPct = (nw.totalLiabilities > 0) ? Math.round((cpt.soldeConverti / nw.totalLiabilities) * 100) : 0;
          statusPillHtml = `<span class="cpt-status-pill debt">${escapeHTML(engine.tr('ACCOUNT_STATUS_DEBT'))} • ${debtPct}%</span>`;
        }
      } else {
        if (cpt.soldeConverti === 0) {
          statusPillHtml = `<span class="cpt-status-pill ready">${escapeHTML(engine.tr('ACCOUNT_STATUS_READY'))}</span>`;
        } else {
          const assetPct = (nw.totalAssets > 0) ? Math.round((cpt.soldeConverti / nw.totalAssets) * 100) : 0;
          statusPillHtml = `<span class="cpt-status-pill funded">${escapeHTML(engine.tr('ACCOUNT_STATUS_FUNDED'))} • ${assetPct}%</span>`;
        }
      }

      return `
        <tr>
          <td>
            <div style="font-weight:700; display:flex; align-items:center; gap:6px; flex-wrap:wrap;">
              <span>${escapeHTML(engine.cleanAssetName(cpt.nom))}</span>
              ${statusPillHtml}
            </div>
            <div style="font-size:0.68rem; opacity:0.8; margin-top:2px;">${cpt.isDette ? engine.tr('ASSET_LIABILITY') : engine.tr('ASSET_ACTIVE')}</div>
          </td>
          <td>
            <span class="badge-flux ${escapeHTML(cpt.classe)}">${escapeHTML(cpt.classeTraduit || engine.trEntity(cpt.classe))}</span>
          </td>
          <td>
            <div style="font-family:monospace, tabular-nums; font-weight:600;">${cpt.soldeNatif.toLocaleString('fr-FR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ${escapeHTML(cpt.deviseNative)}</div>
          </td>
          <td class="td-amount td-amount-with-action" style="color: ${cpt.isDette ? 'var(--c-yellow-royal)' : '#FFF'};">
            <span class="cpt-converted-val">${escapeHTML(engine.formatMoney(cpt.soldeConverti))}</span>
            <button type="button" class="btn-cpt-quick-add" data-account-id="${escapeHTML(cpt.id)}" title="${escapeHTML(engine.tr('ACCOUNT_QUICK_ADD_TX'))}" aria-label="${escapeHTML(engine.tr('ACCOUNT_QUICK_ADD_TX'))} ${escapeHTML(engine.cleanAssetName(cpt.nom))}">
              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
            </button>
          </td>
        </tr>
      `;
    }).join('');
  }

  function renderBudgetView() {
    const bg = engine.computeBudgetZBB();

    document.getElementById('budget-hero-total').textContent = engine.formatMoney(bg.totalAlloue);
    document.getElementById('budget-hero-sub').textContent = `${engine.tr('TH_CONSOMME')} : ${engine.formatMoney(bg.totalConsomme)} • ${engine.tr('TH_ECART')} : ${engine.formatMoney(bg.totalEcart)}`;
    document.getElementById('budget-month-label').textContent = bg.mois;

    const tbody = document.getElementById('budget-table-body');
    if (bg.categories.length === 0) {
      const hasStatusFilter = (engine.budgetStatusFilter && engine.budgetStatusFilter !== 'ALL');
      tbody.innerHTML = `
        <tr>
          <td colspan="4" class="empty-state">
            <div class="empty-state-rich">
              <div class="empty-state-icon-box" aria-hidden="true">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="6"/><circle cx="12" cy="12" r="2"/></svg>
              </div>
              <div class="empty-state-title">${escapeHTML(engine.tr('NO_BUDGET_MONTH'))}</div>
              <p class="empty-state-desc">${escapeHTML(engine.tr('EMPTY_BUDGET_SUB'))}</p>
              <div class="empty-state-actions">
                <button type="button" class="btn-empty-cta" data-empty-action="open-alloc">${escapeHTML(engine.tr('EMPTY_CTA_ALLOC_BG'))}</button>
                ${hasStatusFilter ? `<button type="button" class="btn-empty-cta secondary" data-empty-action="reset-bg-filters">${escapeHTML(engine.tr('EMPTY_CTA_RESET_FILTER'))}</button>` : ''}
                ${engine.month !== '00' ? `<button type="button" class="btn-empty-cta secondary" data-empty-action="view-full-year">${escapeHTML(engine.tr('EMPTY_CTA_FULL_YEAR'))}</button>` : ''}
              </div>
            </div>
          </td>
        </tr>
      `;
      return;
    }

    tbody.innerHTML = bg.categories.map(c => {
      let gaugeClass = 'gauge-normal';
      let statusClass = 'status-normal';
      let statusText = engine.tr('STATUS_OK');

      if (c.isOverBudget) {
        gaugeClass = 'gauge-over';
        statusClass = 'status-over';
        statusText = engine.tr('STATUS_OVER');
      } else if (c.rawPct >= 80) {
        gaugeClass = 'gauge-warning';
        statusClass = 'status-warning';
        statusText = engine.tr('STATUS_WARN');
      }

      const catNomTraduit = escapeHTML(engine.trEntity(c.nom));
      return `
        <tr>
          <td>
            <div style="font-weight:700;">${catNomTraduit}</div>
            <div class="budget-gauge-container">
              <div class="budget-gauge-bar ${gaugeClass}" style="width: ${c.pct}%;"></div>
            </div>
          </td>
          <td>${escapeHTML(engine.formatMoney(c.alloue))}</td>
          <td>${escapeHTML(engine.formatMoney(c.consomme))}</td>
          <td class="td-amount">
            <span class="budget-status-pill ${statusClass}">${Math.round(c.rawPct)}% • ${statusText}</span>
          </td>
        </tr>
      `;
    }).join('');
  }

  // --- 14. BACKUP IMPORT & EXPORT ---
  document.getElementById('btn-export-backup')?.addEventListener('click', () => {
    triggerHaptic(20);
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(engine.exportBackupJSON());
    const dlAnchor = document.createElement('a');
    dlAnchor.setAttribute("href", dataStr);
    dlAnchor.setAttribute("download", `GFMP_Backup_${new Date().toISOString().split('T')[0]}.json`);
    document.body.appendChild(dlAnchor);
    dlAnchor.click();
    dlAnchor.remove();
    showToast('Sauvegarde exportée avec succès !', 'success');
  });

  document.getElementById('btn-import-backup')?.addEventListener('click', () => {
    triggerHaptic(20);
    document.getElementById('file-import-input')?.click();
  });

  document.getElementById('file-import-input')?.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      const res = engine.importBackupJSON(event.target.result);
      if (res.success) {
        showToast(`Restauration réussie (${res.count} flux) !`, 'success');
        renderCurrentView();
      } else {
        showToast(`Erreur d'importation : ${res.error}`, 'error');
      }
    };
    reader.readAsText(file);
  });

  // --- 15. SETTINGS & APP CUSTOMIZATION CONTROLS ---
  function renderSettingsView() {
    const setCurr = document.getElementById('settings-currency-select');
    if (setCurr) setCurr.value = engine.currency;
    const setLng = document.getElementById('settings-lang-select');
    if (setLng) setLng.value = engine.lang;

    const currentTheme = localStorage.getItem('gfmp_theme') || 'blue';
    document.querySelectorAll('#settings-theme-selector .theme-choice-btn').forEach(btn => {
      if (btn.dataset.theme === currentTheme) {
        btn.classList.add('active');
        btn.style.borderColor = '#FADA5E';
      } else {
        btn.classList.remove('active');
        btn.style.borderColor = 'rgba(255,255,255,0.15)';
      }
    });

    const hapticOn = localStorage.getItem('gfmp_haptic') !== 'false';
    const hapticBtn = document.getElementById('btn-toggle-haptic');
    if (hapticBtn) {
      hapticBtn.textContent = hapticOn ? 'ACTIVÉ' : 'DÉSACTIVÉ';
      hapticBtn.style.background = hapticOn ? '#10B981' : '#64748B';
    }

  }

  // Delegated Click Handler for Inventory Quick-Add, Empty States & Action CTAs
  document.addEventListener('click', (e) => {
    const quickAddBtn = e.target.closest('.btn-cpt-quick-add');
    if (quickAddBtn) {
      e.stopPropagation();
      triggerHaptic(18);
      const accId = quickAddBtn.getAttribute('data-account-id');
      openExpressModal(accId);
      return;
    }

    const btn = e.target.closest('[data-empty-action]');
    if (!btn) return;
    triggerHaptic(18);
    const action = btn.getAttribute('data-empty-action');

    if (action === 'add-tx') {
      modalReport?.classList.remove('open');
      openExpressModal();
    } else if (action === 'add-exp') {
      modalReport?.classList.remove('open');
      activeExpType = 'DEPENSE';
      document.querySelectorAll('#exp-type-selector .type-segment-btn').forEach(b => {
        b.classList.toggle('active', b.dataset.type === 'DEPENSE');
      });
      filterCategories('exp-tx-categorie', 'DEPENSE');
      openExpressModal();
    } else if (action === 'open-alloc') {
      document.getElementById('btn-open-alloc-modal')?.click();
    } else if (action === 'add-account') {
      openAddEntityDialog('compte');
    } else if (action === 'reset-cf-filters') {
      engine.searchTerm = '';
      engine.fluxFilter = 'ALL';
      engine.selectedCategoryFilter = null;
      engine.accountFilter = 'ALL';
      const searchIn = document.getElementById('input-tx-search');
      if (searchIn) searchIn.value = '';
      const accSel = document.getElementById('cf-account-filter');
      if (accSel) accSel.value = 'ALL';
      document.querySelectorAll('.search-filter-box .filter-chip[data-filter]').forEach(c => {
        c.classList.toggle('active', c.dataset.filter === 'ALL');
      });
      renderCurrentView();
      showToast(engine.tr('FILTER_CLEARED'), 'info');
    } else if (action === 'reset-nw-filters') {
      engine.patrimoineClassFilter = 'ALL';
      document.querySelectorAll('#nw-class-chips .filter-chip').forEach(c => {
        c.classList.toggle('active', c.dataset.class === 'ALL');
      });
      renderNetWorthView();
    } else if (action === 'reset-bg-filters') {
      engine.budgetStatusFilter = 'ALL';
      document.querySelectorAll('#bg-status-chips .filter-chip').forEach(c => {
        c.classList.toggle('active', c.dataset.bstatus === 'ALL');
      });
      renderBudgetView();
    } else if (action === 'view-full-year') {
      engine.setMonth('00');
      showToast(engine.tr('FULL_YEAR', { year: engine.year }), 'info');
    }
  });

  // Net Worth Inline Add Account Button
  document.getElementById('btn-nw-add-account')?.addEventListener('click', () => {
    triggerHaptic(18);
    openAddEntityDialog('compte');
  });

  // Theme selector
  document.querySelectorAll('#settings-theme-selector .theme-choice-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      triggerHaptic(18);
      const theme = btn.dataset.theme;
      document.body.className = (theme === 'blue') ? '' : `theme-${theme}`;
      localStorage.setItem('gfmp_theme', theme);
      renderSettingsView();
      showToast(`Thème appliqué : ${btn.textContent.trim()}`, 'success');
    });
  });

  // Apply saved theme on startup
  const savedTheme = localStorage.getItem('gfmp_theme');
  if (savedTheme && savedTheme !== 'blue') {
    document.body.className = `theme-${savedTheme}`;
  }

  // Settings Currency select
  document.getElementById('settings-currency-select')?.addEventListener('change', (e) => {
    triggerHaptic(15);
    engine.setCurrency(e.target.value);
    showToast(`Devise de base : ${e.target.value}`, 'info');
  });

  // Settings Language select
  document.getElementById('settings-lang-select')?.addEventListener('change', (e) => {
    triggerHaptic(15);
    engine.setLanguage(e.target.value);
    showToast(`Langue système : ${e.target.value}`, 'info');
  });

  // Settings Haptic toggle
  document.getElementById('btn-toggle-haptic')?.addEventListener('click', () => {
    const cur = localStorage.getItem('gfmp_haptic') !== 'false';
    const next = !cur;
    localStorage.setItem('gfmp_haptic', next ? 'true' : 'false');
    if (next) triggerHaptic(25);
    renderSettingsView();
    showToast(`Retour haptique : ${next ? 'Activé' : 'Désactivé'}`, 'info');
  });

  // Settings Database buttons
  document.getElementById('btn-set-init')?.addEventListener('click', () => {
    triggerHaptic(18);
    selectedInitLang = engine.lang || 'FR';
    selectedInitCurr = engine.currency || 'MUR';
    selectedInitDate = engine.referenceDate || '2026-02-27';
    selectedInitYear = engine.year || '2026';
    selectedInitMonth = engine.month || '02';
    selectedInitTheme = localStorage.getItem('gfmp_theme') || engine.theme || 'blue';
    selectedInitDataMode = 'empty';
    modalInit?.classList.add('open');
    updateInitModalUI();
  });

  document.getElementById('btn-set-refresh')?.addEventListener('click', async () => {
    const btn = document.getElementById('btn-set-refresh');
    if (btn) btn.disabled = true;
    triggerHaptic(18);
    showToast('Actualisation des données et taux via API...', 'info');
    try {
      await engine.fetchLiveExchangeRates(true);
      engine.refreshDatabase();
      renderCurrentView();
      showToast('Données et cours actualisés avec succès !', 'success');
    } catch (e) {
      showToast('Erreur lors de l\'actualisation des cours', 'error');
    } finally {
      if (btn) btn.disabled = false;
    }
  });

  document.getElementById('btn-set-backup')?.addEventListener('click', () => {
    triggerHaptic(18);
    try {
      const jsonStr = engine.backupDatabase();
      const blob = new Blob([jsonStr], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `GFMP_Backup_${new Date().toISOString().substring(0, 10)}.json`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 2000);
      showToast('Sauvegarde JSON générée et téléchargée !', 'success');
    } catch (e) {
      showToast('Erreur lors de la création de la sauvegarde', 'error');
    }
  });

  document.getElementById('btn-set-purge')?.addEventListener('click', async () => {
    triggerHaptic(30);
    const confirmed = await showConfirmDialog({
      title: 'Purger les Données',
      message: 'Attention : Voulez-vous vraiment purger toutes les transactions et enveloppes budgétaires ? Les comptes et catégories seront conservés.',
      confirmText: 'Purger Définitivement',
      cancelText: 'Annuler',
      type: 'danger'
    });
    if (confirmed) {
      engine.purgeDatabase();
      showToast('Toutes les transactions et enveloppes ont été purgées avec succès', 'info');
      populateSelects();
      renderCurrentView();
    }
  });

  document.getElementById('btn-set-audit')?.addEventListener('click', openAuditHistoryModal);

  document.getElementById('btn-set-restore')?.addEventListener('click', () => {
    triggerHaptic(18);
    document.getElementById('settings-file-import')?.click();
  });

  document.getElementById('settings-file-import')?.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const confirmed = await showConfirmDialog({
      title: 'Restaurer la Base',
      message: `Voulez-vous remplacer les données actuelles par la sauvegarde "${file.name}" ?`,
      confirmText: 'Restaurer la Base',
      cancelText: 'Annuler',
      type: 'warning'
    });
    if (!confirmed) {
      e.target.value = '';
      return;
    }
    const reader = new FileReader();
    reader.onload = (event) => {
      const res = engine.restoreDatabase(event.target.result);
      if (res.success) {
        showToast(`Restauration réussie (${res.count} flux) !`, 'success');
        populateSelects();
        renderCurrentView();
      } else {
        showToast(`Erreur d'importation : ${res.error}`, 'error');
      }
      e.target.value = '';
    };
    reader.onerror = () => {
      showToast('Erreur de lecture du fichier', 'error');
      e.target.value = '';
    };
    reader.readAsText(file);
  });

  // --- 16. FULLSCREEN INITIALIZATION MODAL (4-STEP GOLD-STANDARD PRODUCTION v1.01 ONBOARDING & CONFIG) ---
  const modalInit = document.getElementById('modal-initialization');
  let selectedInitLang = engine.lang || 'FR';
  let selectedInitCurr = engine.currency || 'MUR';
  let selectedInitDate = engine.referenceDate || '2026-02-27';
  let selectedInitYear = engine.year || '2026';
  let selectedInitMonth = engine.month || '02';
  let selectedInitTheme = localStorage.getItem('gfmp_theme') || engine.theme || 'blue';
  let selectedInitDataMode = 'empty';

  function trInit(key, params = {}) {
    return engine.tr(key, params, selectedInitLang);
  }

  function getInitMonthNames(langCode) {
    const names = {
      FR: ['Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre'],
      EN: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
      ES: ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'],
      PT: ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
    };
    return names[langCode] || names.FR;
  }

  function updateInitModalUI() {
    // 1. Language Cards Active State
    document.querySelectorAll('#init-lang-grid .init-choice-card').forEach(card => {
      const isSel = (card.dataset.initLang === selectedInitLang);
      card.classList.toggle('active', isSel);
      card.setAttribute('aria-pressed', isSel ? 'true' : 'false');
      const code = card.querySelector('.init-choice-code');
      if (code) code.style.color = isSel ? 'var(--c-yellow-royal)' : '#94A3B8';
    });

    // 2. Currency Cards Active State
    document.querySelectorAll('#init-curr-grid .init-choice-card').forEach(card => {
      const isSel = (card.dataset.initCurr === selectedInitCurr);
      card.classList.toggle('active', isSel);
      card.setAttribute('aria-pressed', isSel ? 'true' : 'false');
      const code = card.querySelector('.init-choice-code');
      if (code) code.style.color = isSel ? 'var(--c-yellow-royal)' : '#94A3B8';
    });

    // 3. Theme Cards Active State
    document.querySelectorAll('#init-theme-grid .init-choice-card').forEach(card => {
      const isSel = (card.dataset.initTheme === selectedInitTheme);
      card.classList.toggle('active', isSel);
      card.setAttribute('aria-pressed', isSel ? 'true' : 'false');
      const code = card.querySelector('.init-choice-code');
      if (code) code.style.color = isSel ? 'var(--c-yellow-royal)' : '#94A3B8';
    });

    // 4. Live Quadrilingual Translation of Configuration Steps
    const setTxt = (id, key) => {
      const el = document.getElementById(id);
      if (el) el.textContent = trInit(key);
    };
    setTxt('init-modal-badge', 'INIT_BADGE');
    setTxt('init-modal-title', 'INIT_TITLE');
    setTxt('init-modal-sub', 'INIT_SUBTITLE');
    setTxt('init-step1-badge', 'INIT_STEP1_BADGE');
    setTxt('init-step1-title', 'INIT_STEP1_TITLE');
    setTxt('init-step1-desc', 'INIT_STEP1_DESC');
    setTxt('init-step2-badge', 'INIT_STEP2_BADGE');
    setTxt('init-step2-title', 'INIT_STEP2_TITLE');
    setTxt('init-step2-desc', 'INIT_STEP2_DESC');
    setTxt('init-step3-badge', 'INIT_STEP3_BADGE');
    setTxt('init-step3-title', 'INIT_STEP3_TITLE');
    setTxt('init-step3-desc', 'INIT_STEP3_DESC');
    setTxt('init-step4-badge', 'INIT_STEP4_BADGE');
    setTxt('init-step4-title', 'INIT_STEP4_TITLE');
    setTxt('init-step4-desc', 'INIT_STEP4_DESC');
    setTxt('lbl-init-date-picker', 'INIT_DATE_LABEL');
    setTxt('lbl-init-month-picker', 'INIT_MONTH_LABEL');
    setTxt('lbl-init-year-picker', 'INIT_YEAR_LABEL');
    setTxt('lbl-init-btn-first', 'INIT_BTN_FIRST');
    setTxt('lbl-init-btn-today', 'INIT_BTN_TODAY');
    setTxt('lbl-btn-confirm-init', 'INIT_BTN_CONFIRM');

    // 5. Translate Currency Card Names
    document.querySelectorAll('#init-curr-grid [data-curr-name]').forEach(el => {
      const curr = el.getAttribute('data-curr-name');
      if (curr) el.textContent = trInit(`CURR_NAME_${curr}`);
    });

    // 6. Translate Theme Card Names & Subtitles
    document.querySelectorAll('#init-theme-grid [data-theme-name]').forEach(el => {
      const th = el.getAttribute('data-theme-name');
      if (th) el.textContent = trInit(`THEME_NAME_${th}`);
    });
    document.querySelectorAll('#init-theme-grid [data-theme-sub]').forEach(el => {
      const th = el.getAttribute('data-theme-sub');
      if (th) el.textContent = trInit(`THEME_SUB_${th}`);
    });

    // 7. Synchronize & Translate Date / Month / Year Controls
    const dateInput = document.getElementById('init-date-input');
    if (dateInput && dateInput.value !== selectedInitDate) {
      dateInput.value = selectedInitDate;
    }

    const monthSelect = document.getElementById('init-month-select');
    const mNames = getInitMonthNames(selectedInitLang);
    if (monthSelect) {
      Array.from(monthSelect.options).forEach(opt => {
        if (opt.value === '00') {
          opt.textContent = `00 — ${trInit('FULL_YEAR', { year: selectedInitYear })}`;
        } else {
          const idx = parseInt(opt.value, 10) - 1;
          if (idx >= 0 && idx < 12) {
            opt.textContent = `${opt.value} — ${mNames[idx]}`;
          }
        }
      });
      monthSelect.value = selectedInitMonth;
    }

    const yearSelect = document.getElementById('init-year-select');
    if (yearSelect) {
      if (!Array.from(yearSelect.options).some(o => o.value === String(selectedInitYear))) {
        const opt = document.createElement('option');
        opt.value = String(selectedInitYear);
        opt.textContent = String(selectedInitYear);
        yearSelect.appendChild(opt);
      }
      yearSelect.value = String(selectedInitYear);
    }

    // Highlight active date shortcut chip
    const firstChip = document.getElementById('btn-init-date-first');
    const todayChip = document.getElementById('btn-init-date-today');
    const realToday = new Date().toISOString().split('T')[0];
    const activeMonthCode = (selectedInitMonth === '00') ? '01' : selectedInitMonth;
    const firstOfMonth = `${selectedInitYear}-${activeMonthCode}-01`;
    if (firstChip) firstChip.classList.toggle('active', selectedInitDate === firstOfMonth);
    if (todayChip) todayChip.classList.toggle('active', selectedInitDate === realToday);

    // 8. Update Live Summary Pill (4 Production Pillars: Language • Currency • Period • Theme)
    const langLabels = { FR: 'Français (FR)', EN: 'English (EN)', ES: 'Español (ES)', PT: 'Português (PT)' };
    const summaryText = document.getElementById('init-summary-text');
    if (summaryText) {
      const prefix = trInit('INIT_SUMMARY_PREFIX');
      const themeUpper = String(selectedInitTheme || 'blue').toUpperCase();
      const themeLabel = trInit(`THEME_NAME_${themeUpper}`);
      const mIdx = parseInt(selectedInitMonth, 10) - 1;
      const periodLabel = (selectedInitMonth === '00')
        ? trInit('FULL_YEAR', { year: selectedInitYear })
        : `${(mNames[mIdx] || selectedInitMonth)} ${selectedInitYear}`;
      summaryText.innerHTML = `${prefix} <strong>${langLabels[selectedInitLang] || selectedInitLang}</strong> • <strong>${selectedInitCurr}</strong> • <strong>${periodLabel}</strong> • <strong>${themeLabel}</strong>`;
    }
  }

  // Step 1: Language Card Click
  document.querySelectorAll('#init-lang-grid .init-choice-card').forEach(card => {
    card.addEventListener('click', () => {
      triggerHaptic(15);
      selectedInitLang = card.dataset.initLang;
      updateInitModalUI();
    });
  });

  // Step 2: Currency Card Click
  document.querySelectorAll('#init-curr-grid .init-choice-card').forEach(card => {
    card.addEventListener('click', () => {
      triggerHaptic(15);
      selectedInitCurr = card.dataset.initCurr;
      updateInitModalUI();
    });
  });

  // Step 3: Date Input & Period Selectors
  document.getElementById('init-date-input')?.addEventListener('change', (e) => {
    const val = e.target.value;
    if (val && /^\d{4}-\d{2}-\d{2}$/.test(val)) {
      triggerHaptic(12);
      selectedInitDate = val;
      const parts = val.split('-');
      selectedInitYear = parts[0];
      selectedInitMonth = parts[1];
      updateInitModalUI();
    }
  });

  document.getElementById('init-month-select')?.addEventListener('change', (e) => {
    triggerHaptic(12);
    selectedInitMonth = e.target.value;
    if (selectedInitMonth !== '00') {
      const dayPart = (selectedInitDate && /^\d{4}-\d{2}-\d{2}$/.test(selectedInitDate)) ? selectedInitDate.split('-')[2] : '01';
      const maxDay = new Date(parseInt(selectedInitYear, 10), parseInt(selectedInitMonth, 10), 0).getDate();
      const safeDay = String(Math.min(parseInt(dayPart, 10) || 1, maxDay)).padStart(2, '0');
      selectedInitDate = `${selectedInitYear}-${selectedInitMonth}-${safeDay}`;
    }
    updateInitModalUI();
  });

  document.getElementById('init-year-select')?.addEventListener('change', (e) => {
    triggerHaptic(12);
    selectedInitYear = e.target.value;
    const monthPart = (selectedInitMonth === '00') ? '02' : selectedInitMonth;
    const dayPart = (selectedInitDate && /^\d{4}-\d{2}-\d{2}$/.test(selectedInitDate)) ? selectedInitDate.split('-')[2] : '01';
    const maxDay = new Date(parseInt(selectedInitYear, 10), parseInt(monthPart, 10), 0).getDate();
    const safeDay = String(Math.min(parseInt(dayPart, 10) || 1, maxDay)).padStart(2, '0');
    selectedInitDate = `${selectedInitYear}-${monthPart}-${safeDay}`;
    updateInitModalUI();
  });

  document.getElementById('btn-init-date-first')?.addEventListener('click', () => {
    triggerHaptic(15);
    const activeMonthCode = (selectedInitMonth === '00') ? '01' : selectedInitMonth;
    selectedInitDate = `${selectedInitYear}-${activeMonthCode}-01`;
    updateInitModalUI();
  });

  document.getElementById('btn-init-date-today')?.addEventListener('click', () => {
    triggerHaptic(15);
    const now = new Date();
    const y = String(now.getFullYear());
    const m = String(now.getMonth() + 1).padStart(2, '0');
    const d = String(now.getDate()).padStart(2, '0');
    selectedInitDate = `${y}-${m}-${d}`;
    selectedInitYear = y;
    selectedInitMonth = m;
    updateInitModalUI();
  });

  // Step 4: Theme Card Click (with Instant Live Theme Preview)
  document.querySelectorAll('#init-theme-grid .init-choice-card').forEach(card => {
    card.addEventListener('click', () => {
      triggerHaptic(15);
      selectedInitTheme = card.dataset.initTheme || 'blue';
      document.body.className = (selectedInitTheme === 'blue') ? '' : `theme-${selectedInitTheme}`;
      updateInitModalUI();
    });
  });

  // Confirm Initialization Button
  document.getElementById('btn-confirm-init')?.addEventListener('click', () => {
    triggerHaptic(25);
    engine.initializeSystem(
      selectedInitCurr,
      selectedInitLang,
      selectedInitDate,
      selectedInitTheme,
      selectedInitYear,
      selectedInitMonth,
      false
    );
    document.body.className = (selectedInitTheme === 'blue') ? '' : `theme-${selectedInitTheme}`;
    localStorage.setItem('gfmp_theme', selectedInitTheme);

    const txDate = document.getElementById('tx-date');
    if (txDate) txDate.value = selectedInitDate;
    const expDate = document.getElementById('exp-tx-date');
    if (expDate) expDate.value = selectedInitDate;

    txDeviseUserModified = false;
    expTxDeviseUserModified = false;
    allocDeviseUserModified = false;

    const txDev = document.getElementById('tx-devise');
    if (txDev) txDev.value = selectedInitCurr;
    const expDev = document.getElementById('exp-tx-devise');
    if (expDev) expDev.value = selectedInitCurr;
    const allocDev = document.getElementById('alloc-devise');
    if (allocDev) allocDev.value = selectedInitCurr;
    const allocBadge = document.getElementById('alloc-currency-badge');
    if (allocBadge) allocBadge.textContent = selectedInitCurr;
    const addEntityDev = document.getElementById('add-entity-devise');
    if (addEntityDev) addEntityDev.value = selectedInitCurr;

    document.querySelectorAll('.curr-btn').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.curr === selectedInitCurr);
    });

    document.documentElement.classList.remove('gfmp-onboarding-active');
    modalInit?.classList.remove('open');
    showToast(`Système configuré : ${selectedInitLang} • ${selectedInitCurr} • ${selectedInitMonth === '00' ? selectedInitYear : selectedInitMonth + '/' + selectedInitYear}`, 'success');
    populateSelects();
    renderSettingsView();
    renderCurrentView();
  });

  // Automatically show on first app launch / after APK installation (Zero-FOUC)
  if (engine.needsOnboarding && modalInit) {
    document.documentElement.classList.add('gfmp-onboarding-active');
    modalInit.classList.add('open');
    try {
      updateInitModalUI();
    } catch (err) {
      console.error('Init modal UI update warning:', err);
    }
  } else {
    document.documentElement.classList.remove('gfmp-onboarding-active');
  }

  // --- 17. PWA BEFOREINSTALLPROMPT HANDLER ---
  let deferredPrompt = null;
  window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    deferredPrompt = e;
    const banner = document.getElementById('pwa-install-banner');
    if (banner) banner.style.display = 'flex';
  });

  document.getElementById('btn-pwa-install')?.addEventListener('click', () => {
    if (deferredPrompt) {
      deferredPrompt.prompt();
      deferredPrompt.userChoice.then(() => {
        deferredPrompt = null;
        const banner = document.getElementById('pwa-install-banner');
        if (banner) banner.style.display = 'none';
      });
    }
  });

  // --- 18. GOLD STANDARD CUTTING-EDGE RESPONSIVE SCROLL CONTROLLER ---
  function initScrollResponsiveness() {
    const header = document.querySelector('.gfmp-header');
    if (!header) return;

    let ticking = false;
    const SCROLL_DOWN_THRESHOLD = 28;
    const SCROLL_UP_THRESHOLD = 10;

    function handleScroll() {
      const scrollY = window.scrollY || document.documentElement.scrollTop || 0;
      if (scrollY > SCROLL_DOWN_THRESHOLD) {
        if (!header.classList.contains('is-compact')) {
          header.classList.add('is-compact');
        }
      } else if (scrollY < SCROLL_UP_THRESHOLD) {
        if (header.classList.contains('is-compact')) {
          header.classList.remove('is-compact');
        }
      }
      ticking = false;
    }

    window.addEventListener('scroll', () => {
      if (!ticking) {
        window.requestAnimationFrame(handleScroll);
        ticking = true;
      }
    }, { passive: true });

    // Initial check in case page starts scrolled
    handleScroll();
  }

  // Subscribe engine notifications (Time slider, Currency switch, Language switch, Add/Delete mutations)
  engine.subscribe(() => {
    renderCurrentView();
  });

  // Initial population & rendering
  populateSelects();
  renderCurrentView();
  initScrollResponsiveness();
});
