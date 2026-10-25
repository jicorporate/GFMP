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
    if (navigator.vibrate) {
      try {
        navigator.vibrate(duration);
      } catch (e) {
        // Ignored on unsupported devices
      }
    }
  }

  // --- 2. FLOATING TOAST NOTIFICATION ---
  function showToast(message, type = 'success') {
    let toast = document.getElementById('gfmp-toast');
    if (!toast) {
      toast = document.createElement('div');
      toast.id = 'gfmp-toast';
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

  // --- 3. VIEW ROUTING SYSTEM ---
  function navigateTo(viewName) {
    document.querySelectorAll('.view-panel').forEach(p => p.style.display = 'none');
    const target = document.getElementById(`view-${viewName}`);
    if (target) {
      target.style.display = 'block';
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    // Update bottom navigation bar
    document.querySelectorAll('.bottom-nav .nav-tab').forEach(t => {
      if (t.dataset.view === viewName) {
        t.classList.add('active');
      } else {
        t.classList.remove('active');
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

  // Bind Hub Central Solid Cards
  document.getElementById('card-nav-saisie')?.addEventListener('click', () => { triggerHaptic(18); navigateTo('transaction'); });
  document.getElementById('card-nav-cashflow')?.addEventListener('click', () => { triggerHaptic(18); navigateTo('cashflow'); });
  document.getElementById('card-nav-networth')?.addEventListener('click', () => { triggerHaptic(18); navigateTo('networth'); });
  document.getElementById('card-nav-budget')?.addEventListener('click', () => { triggerHaptic(18); navigateTo('budget'); });

  // Bind Back Buttons (< Hub Central)
  document.querySelectorAll('.btn-nav-home').forEach(btn => {
    btn.addEventListener('click', () => {
      triggerHaptic(12);
      navigateTo('home');
    });
  });

  // --- 4. TIME SLIDER CONTROLS ---
  document.getElementById('btn-prev-year')?.addEventListener('click', () => { triggerHaptic(15); engine.prevYear(); });
  document.getElementById('btn-next-year')?.addEventListener('click', () => { triggerHaptic(15); engine.nextYear(); });
  document.getElementById('btn-prev-month')?.addEventListener('click', () => { triggerHaptic(15); engine.prevMonth(); });
  document.getElementById('btn-next-month')?.addEventListener('click', () => { triggerHaptic(15); engine.nextMonth(); });

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

  // --- 7. POPULATE FORM DROPDOWNS ---
  function populateSelects() {
    if (!engine.data) return;

    const fillSelects = (prefix) => {
      const cptSel = document.getElementById(`${prefix}-compte`);
      const catSel = document.getElementById(`${prefix}-categorie`);
      const tiersSel = document.getElementById(`${prefix}-tiers`);
      const devSel = document.getElementById(`${prefix}-devise`);

      if (cptSel) {
        cptSel.innerHTML = (engine.data.comptes || []).map(c => `
          <option value="${c.id}">${c.nom} (${c.devise || 'MUR'})</option>
        `).join('');
      }

      if (catSel) {
        const revs = (engine.data.categories || []).filter(c => c.type === 'REVENU');
        const deps = (engine.data.categories || []).filter(c => c.type === 'DEPENSE');
        const others = (engine.data.categories || []).filter(c => c.type !== 'REVENU' && c.type !== 'DEPENSE');

        catSel.innerHTML = `
          <optgroup label="Revenus">
            ${revs.map(c => `<option value="${c.id}">${c.nom}</option>`).join('')}
          </optgroup>
          <optgroup label="Dépenses">
            ${deps.map(c => `<option value="${c.id}">${c.nom}</option>`).join('')}
          </optgroup>
          ${others.length > 0 ? `
            <optgroup label="Autres">
              ${others.map(c => `<option value="${c.id}">${c.nom}</option>`).join('')}
            </optgroup>
          ` : ''}
        `;
      }

      if (tiersSel) {
        tiersSel.innerHTML = (engine.data.tiers || []).map(t => `
          <option value="${t.id}">${t.nom}</option>
        `).join('');
      }

      if (devSel) {
        const keys = Object.keys(engine.data.devises || { MUR: 1, EUR: 1, USD: 1 });
        devSel.innerHTML = keys.map(k => `
          <option value="${k}" ${k === engine.currency ? 'selected' : ''}>${k}</option>
        `).join('');
      }
    };

    fillSelects('tx');
    fillSelects('exp-tx');

    // Budget Alloc Category Select
    const allocCat = document.getElementById('alloc-categorie');
    if (allocCat) {
      const deps = (engine.data.categories || []).filter(c => c.type === 'DEPENSE');
      allocCat.innerHTML = deps.map(c => `
        <option value="${c.id}">${c.nom}</option>
      `).join('');
    }

    // Set default dates to today
    const todayStr = new Date().toISOString().split('T')[0];
    const txDate = document.getElementById('tx-date');
    if (txDate && !txDate.value) txDate.value = todayStr;
    const expDate = document.getElementById('exp-tx-date');
    if (expDate && !expDate.value) expDate.value = todayStr;
  }

  // --- 8. FLOATING ACTION BUTTON (FAB) & EXPRESS BOTTOM SHEET ---
  const fabBtn = document.getElementById('btn-fab-add');
  const modalExpress = document.getElementById('modal-bottom-sheet-tx');
  const btnCloseExpress = document.getElementById('btn-close-express-tx');

  fabBtn?.addEventListener('click', () => {
    triggerHaptic(25);
    populateSelects();
    modalExpress?.classList.add('open');
  });

  btnCloseExpress?.addEventListener('click', () => {
    modalExpress?.classList.remove('open');
  });

  modalExpress?.addEventListener('click', (e) => {
    if (e.target === modalExpress) {
      modalExpress.classList.remove('open');
    }
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

    document.getElementById('tx-montant').value = '';
    document.getElementById('tx-notes').value = '';
    showToast('Transaction enregistrée avec succès !', 'success');
    navigateTo('cashflow');
  });

  // --- 9. BUDGET ALLOCATION MODAL (USF_Budget) ---
  const modalAlloc = document.getElementById('modal-budget-alloc');
  document.getElementById('btn-open-alloc-modal')?.addEventListener('click', () => {
    triggerHaptic(15);
    modalAlloc?.classList.add('open');
    const targetMois = (engine.month === '00') ? `${engine.year}-02` : `${engine.year}-${engine.month}`;
    const allocMois = document.getElementById('alloc-mois');
    if (allocMois) allocMois.value = targetMois;
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

    engine.allocateBudget(cat, amt, mois);
    modalAlloc?.classList.remove('open');
    showToast('Allocation budgétaire enregistrée !', 'success');
    renderBudgetView();
  });

  // --- 10. FINANCIAL EXECUTIVE REPORT MODAL ---
  const modalReport = document.getElementById('modal-financial-report');
  document.getElementById('btn-open-report')?.addEventListener('click', () => {
    triggerHaptic(20);
    const rep = engine.generateFinancialReport();
    const bodyEl = document.getElementById('report-modal-body');
    if (bodyEl) {
      bodyEl.innerHTML = `
        <div style="background:rgba(255,255,255,0.05); padding:12px; border-radius:10px; margin-bottom:12px; border:1px solid rgba(250,218,94,0.3);">
          <div style="font-weight:800; color:var(--c-yellow-royal); font-size:1.05rem;">Synthèse ${rep.period}</div>
          <div style="font-size:0.75rem; color:var(--text-muted);">Devise de consolidation : ${rep.currency}</div>
        </div>
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:8px; margin-bottom:12px;">
          <div style="background:rgba(16,185,129,0.1); border-left:3px solid #10B981; padding:8px; border-radius:6px;">
            <div style="font-size:0.7rem; color:#A7F3D0;">Revenus Totaux</div>
            <div style="font-weight:700; color:#FFF;">${rep.revenus}</div>
          </div>
          <div style="background:rgba(239,68,68,0.1); border-left:3px solid #EF4444; padding:8px; border-radius:6px;">
            <div style="font-size:0.7rem; color:#FECACA;">Dépenses Totales</div>
            <div style="font-weight:700; color:#FFF;">${rep.depenses}</div>
          </div>
          <div style="background:rgba(65,105,225,0.1); border-left:3px solid #4169E1; padding:8px; border-radius:6px;">
            <div style="font-size:0.7rem; color:#BFDBFE;">Cashflow Net</div>
            <div style="font-weight:700; color:#FFF;">${rep.netCashflow}</div>
          </div>
          <div style="background:rgba(250,218,94,0.1); border-left:3px solid #FADA5E; padding:8px; border-radius:6px;">
            <div style="font-size:0.7rem; color:#FEF08A;">Taux d'Épargne</div>
            <div style="font-weight:700; color:#FFF;">${rep.savingsRate}</div>
          </div>
        </div>
        <div style="background:rgba(255,255,255,0.04); padding:10px; border-radius:8px; margin-bottom:12px;">
          <div style="font-weight:700; color:#FFF; margin-bottom:4px;">Patrimoine Net</div>
          <div style="font-size:1.15rem; font-weight:800; color:var(--c-yellow-royal);">${rep.netWorth}</div>
          <div style="font-size:0.75rem; color:var(--text-muted);">Actifs : ${rep.totalAssets} • Dettes : ${rep.totalLiabilities}</div>
        </div>
        <div style="font-weight:700; margin-bottom:6px; color:#FFF;">Top 5 Postes de Dépenses :</div>
        ${rep.topExpenses.map(e => `
          <div style="display:flex; justify-content:space-between; font-size:0.78rem; padding:4px 0; border-bottom:1px solid rgba(255,255,255,0.06);">
            <span>${e.nom}</span>
            <span style="font-weight:700; color:var(--c-yellow-royal);">${e.montant} (${e.pct})</span>
          </div>
        `).join('')}
      `;
    }
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
    const text = `=== GFMP RAPPORT FINANCIER (${rep.period}) ===\nDevise : ${rep.currency}\nRevenus : ${rep.revenus}\nDépenses : ${rep.depenses}\nCashflow Net : ${rep.netCashflow}\nTaux d'Épargne : ${rep.savingsRate}\nValeur Nette : ${rep.netWorth}\nActifs : ${rep.totalAssets} | Dettes : ${rep.totalLiabilities}\nConsommation Budget : ${rep.budgetConsumption}\n========================================`;
    if (navigator.clipboard) {
      navigator.clipboard.writeText(text).then(() => {
        showToast('Rapport copié dans le presse-papier !', 'success');
      }).catch(() => {
        showToast('Rapport généré avec succès', 'info');
      });
    } else {
      showToast('Rapport généré avec succès', 'info');
    }
  });

  // --- 11. SEARCH & FILTER INTERACTIVITY ---
  const searchInput = document.getElementById('input-tx-search');
  searchInput?.addEventListener('input', (e) => {
    engine.searchTerm = e.target.value;
    renderCashflowView();
  });

  document.querySelectorAll('.filter-chip').forEach(chip => {
    chip.addEventListener('click', () => {
      triggerHaptic(12);
      document.querySelectorAll('.filter-chip').forEach(c => c.classList.remove('active'));
      chip.classList.add('active');
      engine.fluxFilter = chip.dataset.filter;
      renderCashflowView();
    });
  });

  // --- 12. INTERACTIVE SVG CHART GENERATORS ---

  // A. Cashflow Multi-Month Trend Chart (6 Months SVG)
  function renderTrendChart(containerId) {
    const container = document.getElementById(containerId);
    if (!container) return;

    const trend = engine.computeCashflowTrend(6);
    if (!trend || trend.length === 0) {
      container.innerHTML = '<div style="color:var(--text-muted); font-size:0.8rem; text-align:center; padding-top:60px;">Aucune donnée de tendance</div>';
      return;
    }

    let maxVal = 0;
    trend.forEach(t => {
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
      const hRev = Math.max(3, (t.revenus / maxVal) * chartHeight);
      const hDep = Math.max(3, (t.depenses / maxVal) * chartHeight);

      const yRev = height - padBottom - hRev;
      const yDep = height - padBottom - hDep;

      const xRev = centerX - barW - 1.5;
      const xDep = centerX + 1.5;

      const isCurrentMonth = t.yearMonth === `${engine.year}-${engine.month}`;

      barsHtml += `
        <!-- Month ${t.label} Group -->
        <g class="chart-bar-group" data-label="${t.label}" style="cursor:pointer;">
          ${isCurrentMonth ? `<rect x="${centerX - step / 2 + 2}" y="${padTop - 8}" width="${step - 4}" height="${chartHeight + 14}" rx="6" fill="rgba(250,218,94,0.08)" stroke="rgba(250,218,94,0.3)" stroke-width="1" />` : ''}
          <!-- Revenus (Vert) -->
          <rect x="${xRev}" y="${yRev}" width="${barW}" height="${hRev}" rx="3" fill="#10B981" filter="drop-shadow(0 2px 4px rgba(16,185,129,0.3))">
            <title>${t.label} Revenus : ${engine.formatMoney(t.revenus)}</title>
          </rect>
          <!-- Dépenses (Rouge) -->
          <rect x="${xDep}" y="${yDep}" width="${barW}" height="${hDep}" rx="3" fill="#EF4444" filter="drop-shadow(0 2px 4px rgba(239,68,68,0.3))">
            <title>${t.label} Dépenses : ${engine.formatMoney(t.depenses)}</title>
          </rect>
          <!-- Label Mois -->
          <text x="${centerX}" y="${height - 8}" text-anchor="middle" fill="${isCurrentMonth ? '#FADA5E' : '#94A3B8'}" font-size="9" font-weight="${isCurrentMonth ? '800' : '600'}" font-family="Inter, sans-serif">${t.label}</text>
        </g>
      `;
    });

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
        <span style="display:inline-flex; align-items:center; gap:5px; color:#A7F3D0;"><span style="width:8px; height:8px; border-radius:2px; background:#10B981;"></span> Revenus</span>
        <span style="display:inline-flex; align-items:center; gap:5px; color:#FECACA;"><span style="width:8px; height:8px; border-radius:2px; background:#EF4444;"></span> Dépenses</span>
      </div>
    `;
  }

  // B. Category Expense Donut Chart (SVG)
  function renderDonutChart(containerId, breakdown, totalDepenses) {
    const container = document.getElementById(containerId);
    if (!container) return;

    if (!breakdown || breakdown.length === 0 || totalDepenses <= 0) {
      container.innerHTML = `
        <div style="width:100%; text-align:center; padding:30px 10px; color:var(--text-muted); font-size:0.8rem;">
          Aucune dépense enregistrée pour ${engine.getMonthLabel()} ${engine.year}
        </div>
      `;
      return;
    }

    const size = 130;
    const strokeW = 16;
    const radius = (size - strokeW) / 2;
    const center = size / 2;
    const circ = 2 * Math.PI * radius;

    let accumPct = 0;
    let pathsHtml = '';

    breakdown.forEach(item => {
      const slicePct = item.pct / 100;
      const strokeDash = `${slicePct * circ} ${circ}`;
      const strokeOffset = -(accumPct * circ);
      accumPct += slicePct;

      pathsHtml += `
        <circle cx="${center}" cy="${center}" r="${radius}"
          fill="transparent"
          stroke="${item.color}"
          stroke-width="${strokeW}"
          stroke-dasharray="${strokeDash}"
          stroke-dashoffset="${strokeOffset}"
          transform="rotate(-90 ${center} ${center})"
          style="transition: stroke-dasharray 0.4s ease;">
          <title>${item.nom}: ${engine.formatMoney(item.montant)} (${Math.round(item.pct)}%)</title>
        </circle>
      `;
    });

    // Top 4 items in legend
    const legendItemsHtml = breakdown.slice(0, 4).map(item => `
      <div class="legend-item">
        <span style="display:flex; align-items:center; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; max-width:110px;">
          <span class="legend-color-dot" style="background:${item.color};"></span>
          ${item.nom}
        </span>
        <span style="font-weight:700; color:#FFF;">${Math.round(item.pct)}%</span>
      </div>
    `).join('');

    container.innerHTML = `
      <div style="position:relative; width:${size}px; height:${size}px; flex-shrink:0;">
        <svg width="${size}" height="${size}" viewBox="0 0 ${size} ${size}">
          <circle cx="${center}" cy="${center}" r="${radius}" fill="transparent" stroke="rgba(255,255,255,0.06)" stroke-width="${strokeW}" />
          ${pathsHtml}
        </svg>
        <div style="position:absolute; top:0; left:0; width:100%; height:100%; display:flex; flex-direction:column; align-items:center; justify-content:center; pointer-events:none;">
          <div style="font-size:0.6rem; color:var(--text-muted); text-transform:uppercase;">Total</div>
          <div style="font-size:0.8rem; font-weight:800; color:var(--c-yellow-royal); max-width:90px; text-align:center; overflow:hidden; text-overflow:ellipsis;">
            ${engine.formatMoney(totalDepenses)}
          </div>
        </div>
      </div>
      <div class="donut-legend">
        ${legendItemsHtml}
      </div>
    `;
  }

  // --- 13. RENDER ACTIVE VIEWS ---
  function renderCurrentView() {
    // Header labels
    document.getElementById('lbl-year').textContent = engine.year;
    document.getElementById('lbl-month').textContent = engine.getMonthLabel();

    // Currency button active state
    document.querySelectorAll('.curr-btn').forEach(btn => {
      if (btn.dataset.curr === engine.currency) btn.classList.add('active');
      else btn.classList.remove('active');
    });

    // Language button active state
    document.querySelectorAll('.lang-btn').forEach(btn => {
      if (btn.dataset.lang === engine.lang) btn.classList.add('active');
      else btn.classList.remove('active');
    });

    // Translate dynamic headers with robust fallbacks
    document.getElementById('app-title-display').textContent = engine.tr('APP_TITLE') || 'GFMP • SFP v1.0';
    document.getElementById('txt-welcome').textContent = engine.tr('WELCOME') || 'Sélectionnez un module d\'application ci-dessous pour démarrer.';
    document.getElementById('lbl-card-saisie').textContent = engine.tr('SAISIE_T') || 'SAISIE TRANSACTION';
    document.getElementById('desc-card-saisie').textContent = engine.tr('SAISIE_D') || 'Ajouter un revenu, dépense ou virement interne avec validation d\'intégrité.';
    document.getElementById('lbl-card-cashflow').textContent = engine.tr('DASH_T') || 'DASHBOARD CASHFLOW';
    document.getElementById('desc-card-cashflow').textContent = engine.tr('DASH_D') || 'Analyser les flux mensuels consolidés et le net de trésorerie.';
    document.getElementById('lbl-card-networth').textContent = engine.tr('NETW_T') || 'BILAN PATRIMONIAL';
    document.getElementById('desc-card-networth').textContent = engine.tr('NETW_D') || 'Calcul de la Valeur Nette et photographie patrimoniale à date.';
    document.getElementById('lbl-card-budget').textContent = engine.tr('BUDG_T') || 'PILOTAGE BUDGÉTAIRE';
    document.getElementById('desc-card-budget').textContent = engine.tr('BUDG_D') || 'Suivi des enveloppes (ZBB) avec jauges de consommation dorées.';

    if (engine.currentView === 'home') renderHomeView();
    else if (engine.currentView === 'cashflow') renderCashflowView();
    else if (engine.currentView === 'networth') renderNetWorthView();
    else if (engine.currentView === 'budget') renderBudgetView();
    else if (engine.currentView === 'transaction') populateSelects();
  }

  function renderHomeView() {
    const cf = engine.computeCashflow();
    const nw = engine.computeNetWorth();
    const bg = engine.computeBudgetZBB();

    const previewCf = document.getElementById('preview-cashflow-net');
    if (previewCf) {
      previewCf.textContent = engine.formatMoney(cf.netCashflow);
      previewCf.style.color = cf.netCashflow >= 0 ? '#FADA5E' : '#EF4444';
    }

    const previewNw = document.getElementById('preview-networth-val');
    if (previewNw) {
      previewNw.textContent = engine.formatMoney(nw.netWorth);
    }

    const previewBg = document.getElementById('preview-budget-pct');
    if (previewBg) {
      const pct = bg.totalAlloue > 0 ? Math.round((bg.totalConsomme / bg.totalAlloue) * 100) : 0;
      previewBg.textContent = `${pct}%`;
    }
  }

  function renderCashflowView() {
    const cf = engine.computeCashflow();

    document.getElementById('kpi-cashflow-rev').textContent = engine.formatMoney(cf.totalRevenus);
    document.getElementById('kpi-cashflow-dep').textContent = engine.formatMoney(cf.totalDepenses);
    
    const netEl = document.getElementById('kpi-cashflow-net');
    const netCard = document.getElementById('card-kpi-net');
    netEl.textContent = engine.formatMoney(cf.netCashflow);
    if (cf.netCashflow >= 0) {
      netCard.className = 'kpi-mini-card kpi-net positive';
    } else {
      netCard.className = 'kpi-mini-card kpi-net negative';
    }

    // Render Dynamic SVG Charts
    renderTrendChart('svg-trend-chart-box');
    renderDonutChart('svg-donut-chart-box', cf.breakdown, cf.totalDepenses);

    const donutPeriodLabel = document.getElementById('donut-period-label');
    if (donutPeriodLabel) {
      donutPeriodLabel.textContent = `${engine.getMonthLabel()} ${engine.year}`;
    }

    // Transactions Count
    document.getElementById('cashflow-count-label').textContent = `${cf.transactions.length} flux`;

    // Render Violet Zebra Table
    const tbody = document.getElementById('cashflow-table-body');
    if (cf.transactions.length === 0) {
      tbody.innerHTML = `<tr><td colspan="5" class="empty-state">Aucune transaction trouvée pour cette période / filtre</td></tr>`;
      return;
    }

    tbody.innerHTML = cf.transactions.map(tx => `
      <tr>
        <td>
          <div style="font-weight:700;">${tx.date}</div>
          <div style="font-size:0.68rem; opacity:0.8;">#${tx.id}</div>
        </td>
        <td>
          <div style="font-weight:700;">${tx.nomCategorie}</div>
          <div style="font-size:0.7rem; opacity:0.8;">${tx.nomCompte}</div>
        </td>
        <td>
          <div>${tx.nomTiers}</div>
          <span class="badge-flux ${tx.typeFlux}">${tx.typeFlux}</span>
        </td>
        <td class="td-amount">
          <div>${engine.formatMoney(tx.montantConverti)}</div>
          ${tx.devise !== engine.currency ? `<div style="font-size:0.68rem; opacity:0.75;">(${tx.montant} ${tx.devise})</div>` : ''}
        </td>
        <td style="text-align:center;">
          <button class="btn-del-tx" data-id="${tx.id}" aria-label="Supprimer" title="Supprimer">
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
      btn.addEventListener('click', () => {
        triggerHaptic(20);
        const id = btn.dataset.id;
        if (confirm(`Confirmez-vous la suppression de l'opération #${id} ?`)) {
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

    if (nw.netWorth >= 0) {
      heroCard.className = 'networth-hero positive';
    } else {
      heroCard.className = 'networth-hero negative';
    }

    document.getElementById('nw-hero-assets').textContent = engine.formatMoney(nw.totalAssets);
    document.getElementById('nw-hero-liabilities').textContent = engine.formatMoney(nw.totalLiabilities);

    // Asset Allocation Stack Bar
    const stackBar = document.getElementById('nw-asset-stack-bar');
    const stackLegend = document.getElementById('nw-asset-legend');
    if (stackBar && nw.totalAssets > 0) {
      const liqPct = Math.round((nw.allocation.LIQUIDITE / nw.totalAssets) * 100);
      const invPct = Math.round((nw.allocation.INVESTISSEMENT / nw.totalAssets) * 100);
      const restPct = Math.max(0, 100 - liqPct - invPct);

      stackBar.innerHTML = `
        <div class="stack-segment" style="width:${liqPct}%; background:#10B981;" title="Liquidités : ${liqPct}%"></div>
        <div class="stack-segment" style="width:${invPct}%; background:#4169E1;" title="Investissements : ${invPct}%"></div>
        <div class="stack-segment" style="width:${restPct}%; background:#FADA5E;" title="Autres : ${restPct}%"></div>
      `;

      if (stackLegend) {
        stackLegend.innerHTML = `
          <span><span style="color:#10B981;">●</span> Liquidités (${liqPct}%)</span>
          <span><span style="color:#4169E1;">●</span> Investissements (${invPct}%)</span>
          <span><span style="color:#FADA5E;">●</span> Autres (${restPct}%)</span>
        `;
      }
    } else if (stackBar) {
      stackBar.innerHTML = `<div class="stack-segment" style="width:100%; background:rgba(255,255,255,0.1);"></div>`;
      if (stackLegend) stackLegend.innerHTML = `<span>Aucun actif</span>`;
    }

    // Accounts Violet Zebra Table
    const tbody = document.getElementById('networth-table-body');
    if (nw.accounts.length === 0) {
      tbody.innerHTML = `<tr><td colspan="4" class="empty-state">Aucun compte actif</td></tr>`;
      return;
    }

    tbody.innerHTML = nw.accounts.map(cpt => `
      <tr>
        <td>
          <div style="font-weight:700;">${cpt.nom}</div>
          <div style="font-size:0.68rem; opacity:0.8;">${cpt.isDette ? 'Passif / Dette' : 'Actif'}</div>
        </td>
        <td>
          <span class="badge-flux ${cpt.classe}">${cpt.classe}</span>
        </td>
        <td>
          <div>${cpt.soldeNatif.toLocaleString('fr-FR', { minimumFractionDigits: 2 })} ${cpt.deviseNative}</div>
        </td>
        <td class="td-amount" style="color: ${cpt.isDette ? 'var(--c-yellow-royal)' : '#FFF'};">
          ${engine.formatMoney(cpt.soldeConverti)}
        </td>
      </tr>
    `).join('');
  }

  function renderBudgetView() {
    const bg = engine.computeBudgetZBB();

    document.getElementById('budget-hero-total').textContent = engine.formatMoney(bg.totalAlloue);
    document.getElementById('budget-hero-sub').textContent = `Consommé : ${engine.formatMoney(bg.totalConsomme)} • Écart : ${engine.formatMoney(bg.totalEcart)}`;
    document.getElementById('budget-month-label').textContent = bg.mois;

    const tbody = document.getElementById('budget-table-body');
    if (bg.categories.length === 0) {
      tbody.innerHTML = `<tr><td colspan="4" class="empty-state">Aucune enveloppe allouée pour ce mois</td></tr>`;
      return;
    }

    tbody.innerHTML = bg.categories.map(c => `
      <tr>
        <td>
          <div style="font-weight:700;">${c.nom}</div>
          <div class="databar-container">
            <div class="databar-fill ${c.isOverBudget ? 'overbudget' : ''}" style="width: ${c.pct}%;"></div>
          </div>
        </td>
        <td>${engine.formatMoney(c.alloue)}</td>
        <td>${engine.formatMoney(c.consomme)}</td>
        <td class="td-amount" style="color: ${c.isOverBudget ? 'var(--c-red-alert)' : 'var(--c-yellow-royal)'};">
          ${Math.round(c.rawPct)}%
        </td>
      </tr>
    `).join('');
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

  // --- 15. PWA BEFOREINSTALLPROMPT HANDLER ---
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

  // Subscribe engine notifications (Time slider, Currency switch, Language switch, Add/Delete mutations)
  engine.subscribe(() => {
    renderCurrentView();
  });

  // Initial population & rendering
  populateSelects();
  renderCurrentView();
});
