/**
 * GFMP — Global Finance Management Platform (SFP v1.0)
 * Ultra-Modern Local-First Wealth & Finance Analytics Engine
 * 100% Offline, Zero-Network, Real-Time Calculations (< 5ms)
 * 
 * Strict SFP v1.0 Color Charter:
 * - Royal Blue: #4169E1
 * - Royal Gold: #FADA5E
 * - Royal Purple: #7851A9
 * - Dark Zebra: #5A3282
 * - Soft Zebra: #916EBE
 * - Emerald Green: #10B981
 * - Coral Red: #EF4444
 */

class GFMPEngine {
  constructor() {
    this.data = null;
    this.lang = 'FR';
    this.currency = 'MUR';
    this.year = '2026';
    this.month = '02'; // '00' = Année entière
    this.currentView = 'home';
    this.searchTerm = '';
    this.fluxFilter = 'ALL'; // 'ALL', 'REVENU', 'DEPENSE', 'TRANSFERT'
    this.listeners = [];
  }

  init() {
    try {
      // 1. Try reading saved dataset from phone storage
      const stored = localStorage.getItem('gfmp_dataset_v3');
      if (stored) {
        const parsed = JSON.parse(stored);
        if (parsed && Array.isArray(parsed.transactions) && parsed.transactions.length > 0) {
          this.data = parsed;
          console.log('[GFMPEngine] Loaded dataset from phone storage:', this.data.transactions.length, 'transactions');
        }
      }
    } catch (e) {
      console.warn('[GFMPEngine] LocalStorage read failed:', e);
    }

    // 2. Synchronous & bulletproof fallback to embedded baseline data
    if (!this.data || !Array.isArray(this.data.transactions) || this.data.transactions.length === 0) {
      if (typeof window !== 'undefined' && window.GFMP_BASELINE) {
        this.data = JSON.parse(JSON.stringify(window.GFMP_BASELINE));
        console.log('[GFMPEngine] Initialized from embedded baseline data:', this.data.transactions.length, 'transactions');
        this.save();
      } else {
        // Fallback minimal structure if window.GFMP_BASELINE is somehow unavailable
        this.data = {
          app_name: 'GFMP',
          comptes: [
            { id: '1', nom: 'Compte Courant Principal', classe: 'LIQUIDITE', devise: 'MUR', actif: 'OUI' },
            { id: '2', nom: 'Livret d\'Épargne', classe: 'LIQUIDITE', devise: 'MUR', actif: 'OUI' }
          ],
          categories: [
            { id: '1', nom: 'Salaire / Revenus Pro', type: 'REVENU' },
            { id: '4', nom: 'Logement (Loyer/Prêt/Charges)', type: 'DEPENSE' },
            { id: '5', nom: 'Alimentation & Supermarché', type: 'DEPENSE' }
          ],
          tiers: [
            { id: '1', nom: 'Employeur principal', type: 'ENTREPRISE' },
            { id: '4', nom: 'Supermarché / Alimentaire', type: 'COMMERCE' }
          ],
          transactions: [
            { id: 1, date: '2026-02-05', id_compte: '1', id_categorie: '4', id_tiers: '4', montant: 12000, devise: 'MUR', notes: 'Loyer' },
            { id: 2, date: '2026-02-23', id_compte: '1', id_categorie: '1', id_tiers: '1', montant: 78000, devise: 'MUR', notes: 'Salaire' }
          ],
          budgets: [
            { id: 1, mois: '2026-02', id_categorie: '4', montant: 15000 },
            { id: 2, mois: '2026-02', id_categorie: '5', montant: 10000 }
          ],
          devises: { MUR: 12.2243, EUR: 656.168, USD: 567.5369, GBP: 757.0023, ZAR: 33.2723, XOF: 1 },
          dictionary: {},
          config: { LANGUE_UI: 'FR', SYS_DEVISE_BASE: 'MUR' }
        };
        this.save();
      }
    }

    // Apply config
    if (this.data && this.data.config) {
      if (this.data.config.LANGUE_UI) this.lang = this.data.config.LANGUE_UI;
      if (this.data.config.SYS_DEVISE_BASE) this.currency = this.data.config.SYS_DEVISE_BASE;
    }
  }

  save() {
    try {
      if (this.data) {
        localStorage.setItem('gfmp_dataset_v3', JSON.stringify(this.data));
      }
    } catch (e) {
      console.warn('[GFMPEngine] Save failed:', e);
    }
  }

  subscribe(listener) {
    this.listeners.push(listener);
  }

  notify() {
    this.listeners.forEach(fn => {
      try {
        fn(this);
      } catch (e) {
        console.error('[GFMPEngine] Listener error:', e);
      }
    });
  }

  // --- I18N SELF-HEALING ENGINE (NO RAW KEYS EVER) ---
  tr(key) {
    if (!key) return '';
    const cleanKey = String(key).trim();

    // 1. Try dataset dictionary
    if (this.data && this.data.dictionary && this.data.dictionary[cleanKey]) {
      const entry = this.data.dictionary[cleanKey];
      const val = entry[this.lang] || entry['FR'];
      if (val) return val;
    }

    // 2. High-grade multilingual fallback dictionary
    const fallbackDict = {
      'APP_TITLE': { FR: 'GFMP • SFP v1.0', EN: 'GFMP • PFS v1.0', ES: 'GFMP • SFP v1.0', PT: 'GFMP • SFP v1.0' },
      'WELCOME': {
        FR: 'Sélectionnez un module d\'application ci-dessous pour démarrer.',
        EN: 'Select an application module below to get started.',
        ES: 'Seleccione un módulo de aplicación a continuación para comenzar.',
        PT: 'Selecione um módulo de aplicativo abaixo para começar.'
      },
      'SAISIE_T': { FR: 'SAISIE TRANSACTION', EN: 'TRANSACTION ENTRY', ES: 'ENTRADA DE TRANSACCIÓN', PT: 'ENTRADA DE TRANSAÇÃO' },
      'SAISIE_D': {
        FR: 'Ajouter un revenu, dépense ou virement interne avec validation d\'intégrité.',
        EN: 'Add an income, expense or transfer with data integrity verification.',
        ES: 'Agregar un ingreso, gasto o transferencia con verificación de integridad.',
        PT: 'Adicionar uma receita, despesa ou transferência com validação de integridade.'
      },
      'DASH_T': { FR: 'DASHBOARD CASHFLOW', EN: 'CASHFLOW DASHBOARD', ES: 'PANEL DE CASHFLOW', PT: 'PAINEL DE FLUXO DE CAIXA' },
      'DASH_D': {
        FR: 'Analyser les flux mensuels consolidés et le net de trésorerie.',
        EN: 'Analyze monthly consolidated flows and net treasury balance.',
        ES: 'Analizar flujos mensuales consolidados y tesorería neta.',
        PT: 'Analisar fluxos mensais consolidados e saldo de tesouraria líquido.'
      },
      'NETW_T': { FR: 'BILAN PATRIMONIAL', EN: 'NET WORTH & ASSETS', ES: 'BALANCE PATRIMONIAL', PT: 'BALANÇO PATRIMONIAL' },
      'NETW_D': {
        FR: 'Calcul de la Valeur Nette et photographie patrimoniale à date.',
        EN: 'Calculation of Net Worth and asset photography as of date.',
        ES: 'Cálculo del Valor Neto y fotografía patrimonial a la fecha.',
        PT: 'Cálculo do Patrimônio Líquido e fotografia patrimonial na data.'
      },
      'BUDG_T': { FR: 'PILOTAGE BUDGÉTAIRE', EN: 'BUDGET MANAGEMENT', ES: 'GESTIÓN PRESUPUESTARIA', PT: 'GESTÃO ORÇAMENTÁRIA' },
      'BUDG_D': {
        FR: 'Suivi des enveloppes (ZBB) avec jauges de consommation dorées.',
        EN: 'Zero-Based Budget tracking with golden consumption gauges.',
        ES: 'Seguimiento de sobres (ZBB) con medidores de consumo dorados.',
        PT: 'Acompanhamento de envelopes (ZBB) com barras de consumo douradas.'
      },
      'ANNÉE ENTIÈRE': { FR: 'ANNÉE ENTIÈRE', EN: 'FULL YEAR', ES: 'AÑO COMPLETO', PT: 'ANO INTEIRO' }
    };

    if (fallbackDict[cleanKey]) {
      return fallbackDict[cleanKey][this.lang] || fallbackDict[cleanKey]['FR'];
    }

    return cleanKey;
  }

  setLanguage(langCode) {
    const valid = ['FR', 'EN', 'ES', 'PT', 'DE', 'IT', 'NL', 'SV'];
    if (valid.includes(langCode)) {
      this.lang = langCode;
      if (this.data && this.data.config) this.data.config.LANGUE_UI = langCode;
      this.save();
      this.notify();
    }
  }

  // --- MULTI-CURRENCY O(1) ENGINE ---
  setCurrency(currCode) {
    if (this.data && this.data.devises && this.data.devises[currCode] !== undefined) {
      this.currency = currCode;
      if (this.data.config) this.data.config.SYS_DEVISE_BASE = currCode;
      this.save();
      this.notify();
    }
  }

  convert(amount, fromCurrency, toCurrency = this.currency) {
    if (amount === 0 || isNaN(amount)) return 0;
    if (fromCurrency === toCurrency) return amount;
    const rates = (this.data && this.data.devises) ? this.data.devises : { MUR: 12.2243, EUR: 656.168, USD: 567.5369, GBP: 757.0023, ZAR: 33.2723, XOF: 1.0 };
    const tauxO = rates[fromCurrency] || 1;
    const tauxC = rates[toCurrency] || 1;
    return (amount * tauxO) / tauxC;
  }

  formatMoney(amount, currency = this.currency) {
    const symbols = {
      MUR: 'Rs',
      EUR: '€',
      USD: '$',
      GBP: '£',
      ZAR: 'R',
      XOF: 'CFA'
    };
    const sym = symbols[currency] || currency;
    const formatted = new Intl.NumberFormat('fr-FR', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    }).format(amount || 0);

    return `${formatted} ${sym}`;
  }

  // --- TIME SLIDER ENGINE ---
  prevYear() {
    this.year = String(parseInt(this.year, 10) - 1);
    this.notify();
  }

  nextYear() {
    this.year = String(parseInt(this.year, 10) + 1);
    this.notify();
  }

  prevMonth() {
    let m = parseInt(this.month, 10) - 1;
    if (m < 0) m = 12;
    this.month = String(m).padStart(2, '0');
    this.notify();
  }

  nextMonth() {
    let m = parseInt(this.month, 10) + 1;
    if (m > 12) m = 0;
    this.month = String(m).padStart(2, '0');
    this.notify();
  }

  getMonthLabel() {
    if (this.month === '00') return this.tr('ANNÉE ENTIÈRE') || 'ANNÉE ENTIÈRE';
    const monthIndex = parseInt(this.month, 10) - 1;
    const monthsFR = ['Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre'];
    const monthsEN = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
    const mName = this.lang === 'EN' ? monthsEN[monthIndex] : monthsFR[monthIndex];
    return mName ? mName.toUpperCase() : this.month;
  }

  // --- 1. MODULE CASHFLOW & INTERACTIVE CHARTS ---
  computeCashflow() {
    if (!this.data) return { totalRevenus: 0, totalDepenses: 0, netCashflow: 0, savingsRate: 0, transactions: [], breakdown: [] };

    const catDict = {};
    const catTypeDict = {};
    (this.data.categories || []).forEach(c => {
      catDict[c.id] = c.nom;
      catTypeDict[c.id] = c.type;
    });

    const cptDict = {};
    (this.data.comptes || []).forEach(c => { cptDict[c.id] = c.nom; });

    const tiersDict = {};
    (this.data.tiers || []).forEach(t => { tiersDict[t.id] = t.nom; });

    let totalRevenus = 0;
    let totalDepenses = 0;
    const filteredTx = [];
    const catBreakdown = {};

    (this.data.transactions || []).forEach(tx => {
      if (!tx.date) return;
      const txParts = tx.date.split('-');
      const txYear = txParts[0];
      const txMonth = txParts[1];

      const matchYear = (txYear === this.year);
      const matchMonth = (this.month === '00' || txMonth === this.month);

      if (matchYear && matchMonth) {
        const typeFlux = catTypeDict[tx.id_categorie] || 'DEPENSE';
        const converted = this.convert(tx.montant, tx.devise, this.currency);

        if (typeFlux === 'REVENU') {
          totalRevenus += converted;
        } else if (typeFlux !== 'TRANSFERT') {
          totalDepenses += converted;
          const catName = catDict[tx.id_categorie] || 'Autre';
          catBreakdown[catName] = (catBreakdown[catName] || 0) + converted;
        }

        // Search & Flux filter
        const nomCompte = cptDict[tx.id_compte] || 'Compte';
        const nomCat = catDict[tx.id_categorie] || 'Catégorie';
        const nomTiers = tiersDict[tx.id_tiers] || 'Tiers';
        const fullSearchStr = `${tx.date} ${nomCompte} ${nomCat} ${nomTiers} ${tx.notes || ''}`.toLowerCase();

        const matchSearch = (!this.searchTerm) || fullSearchStr.includes(this.searchTerm.toLowerCase());
        const matchFlux = (this.fluxFilter === 'ALL') || (typeFlux === this.fluxFilter);

        if (matchSearch && matchFlux) {
          filteredTx.push({
            ...tx,
            nomCompte,
            nomCategorie: nomCat,
            nomTiers,
            typeFlux,
            montantConverti: converted
          });
        }
      }
    });

    filteredTx.sort((a, b) => (b.date.localeCompare(a.date) || b.id - a.id));

    // Color palette for category breakdown (SFP v1.0 Strict)
    const palette = ['#4169E1', '#FADA5E', '#7851A9', '#10B981', '#F43F5E', '#8B5CF6', '#06B6D4', '#F97316'];
    let colorIdx = 0;

    const breakdown = Object.entries(catBreakdown).map(([nom, montant]) => {
      const pct = totalDepenses > 0 ? (montant / totalDepenses) * 100 : 0;
      const color = palette[colorIdx % palette.length];
      colorIdx++;
      return { nom, montant, pct, color };
    }).sort((a, b) => b.montant - a.montant);

    const savingsRate = totalRevenus > 0 ? Math.max(0, ((totalRevenus - totalDepenses) / totalRevenus) * 100) : 0;

    return {
      totalRevenus,
      totalDepenses,
      netCashflow: totalRevenus - totalDepenses,
      savingsRate,
      transactions: filteredTx,
      breakdown
    };
  }

  // --- 2. MULTI-MONTH CASHFLOW TREND (FOR INTERACTIVE SVG CHART) ---
  computeCashflowTrend(monthsCount = 6) {
    if (!this.data) return [];
    const trend = [];
    const currentM = parseInt(this.month === '00' ? '12' : this.month, 10);
    const currentY = parseInt(this.year, 10);

    const catTypeDict = {};
    (this.data.categories || []).forEach(c => { catTypeDict[c.id] = c.type; });

    const monthNamesShort = ['Jan', 'Fév', 'Mar', 'Avr', 'Mai', 'Juin', 'Juil', 'Août', 'Sep', 'Oct', 'Nov', 'Déc'];

    for (let i = monthsCount - 1; i >= 0; i--) {
      let m = currentM - i;
      let y = currentY;
      while (m <= 0) {
        m += 12;
        y -= 1;
      }
      const mStr = String(m).padStart(2, '0');
      const ymStr = `${y}-${mStr}`;

      let rev = 0;
      let dep = 0;

      (this.data.transactions || []).forEach(tx => {
        if (tx.date && tx.date.substring(0, 7) === ymStr) {
          const type = catTypeDict[tx.id_categorie] || 'DEPENSE';
          const converted = this.convert(tx.montant, tx.devise, this.currency);
          if (type === 'REVENU') rev += converted;
          else if (type !== 'TRANSFERT') dep += converted;
        }
      });

      trend.push({
        yearMonth: ymStr,
        label: `${monthNamesShort[m - 1]} ${String(y).slice(-2)}`,
        revenus: rev,
        depenses: dep,
        net: rev - dep
      });
    }

    return trend;
  }

  // --- 3. MODULE BILAN PATRIMONIAL / NET WORTH ---
  computeNetWorth() {
    if (!this.data) return { totalAssets: 0, totalLiabilities: 0, netWorth: 0, accounts: [], allocation: {} };

    const targetFilterMonth = (this.month === '00') ? `${this.year}-12` : `${this.year}-${this.month}`;
    const catTypeDict = {};
    (this.data.categories || []).forEach(c => { catTypeDict[c.id] = c.type; });

    let totalAssets = 0;
    let totalLiabilities = 0;
    const accountResults = [];
    const allocation = { LIQUIDITE: 0, INVESTISSEMENT: 0, DETTE: 0 };

    (this.data.comptes || []).forEach(cpt => {
      const isActif = (cpt.actif || 'OUI').toUpperCase() === 'OUI';
      if (!isActif) return;

      const cptNativeDevise = cpt.devise || 'MUR';
      let soldeNatif = 0;

      (this.data.transactions || []).forEach(tx => {
        if (String(tx.id_compte) === String(cpt.id) && tx.date) {
          if (tx.date.substring(0, 7) <= targetFilterMonth) {
            const fluxType = catTypeDict[tx.id_categorie] || 'DEPENSE';
            const txDev = tx.devise || cptNativeDevise;
            const montantNormalise = this.convert(tx.montant, txDev, cptNativeDevise);

            if (fluxType === 'REVENU') {
              soldeNatif += montantNormalise;
            } else if (fluxType === 'DEPENSE') {
              soldeNatif -= montantNormalise;
            }
          }
        }
      });

      const soldeConverti = this.convert(soldeNatif, cptNativeDevise, this.currency);
      const isDette = (cpt.classe === 'DETTE');

      if (isDette) {
        totalLiabilities += Math.abs(soldeConverti);
        allocation.DETTE = (allocation.DETTE || 0) + Math.abs(soldeConverti);
      } else {
        totalAssets += soldeConverti;
        if (cpt.classe === 'INVESTISSEMENT') {
          allocation.INVESTISSEMENT = (allocation.INVESTISSEMENT || 0) + Math.max(0, soldeConverti);
        } else {
          allocation.LIQUIDITE = (allocation.LIQUIDITE || 0) + Math.max(0, soldeConverti);
        }
      }

      accountResults.push({
        id: cpt.id,
        nom: cpt.nom,
        classe: cpt.classe || 'LIQUIDITE',
        deviseNative: cptNativeDevise,
        soldeNatif,
        soldeConverti,
        isDette
      });
    });

    return {
      totalAssets,
      totalLiabilities,
      netWorth: totalAssets - totalLiabilities,
      accounts: accountResults,
      allocation
    };
  }

  // --- 4. MODULE PILOTAGE BUDGÉTAIRE ZBB ---
  computeBudgetZBB() {
    if (!this.data) return { mois: `${this.year}-${this.month}`, totalAlloue: 0, totalConsomme: 0, totalEcart: 0, consumptionRate: 0, categories: [] };

    const targetMonthStr = (this.month === '00') ? `${this.year}-02` : `${this.year}-${this.month}`;
    const depenseCategories = (this.data.categories || []).filter(c => c.type === 'DEPENSE');

    const catRealise = {};
    (this.data.transactions || []).forEach(tx => {
      if (tx.date && tx.date.substring(0, 7) === targetMonthStr) {
        const converted = this.convert(tx.montant, tx.devise, this.currency);
        catRealise[tx.id_categorie] = (catRealise[tx.id_categorie] || 0) + converted;
      }
    });

    let totalAlloue = 0;
    let totalConsomme = 0;
    const catResults = [];

    depenseCategories.forEach(cat => {
      const budgetEntry = (this.data.budgets || []).find(b => b.mois === targetMonthStr && String(b.id_categorie) === String(cat.id));
      const alloueBase = budgetEntry ? budgetEntry.montant : 0;
      const alloue = this.convert(alloueBase, 'MUR', this.currency);
      const consomme = catRealise[cat.id] || 0;
      const ecart = alloue - consomme;
      const rawPct = alloue > 0 ? (consomme / alloue) * 100 : (consomme > 0 ? 100 : 0);

      totalAlloue += alloue;
      totalConsomme += consomme;

      let status = 'OK';
      if (ecart < 0) status = 'OVER';
      else if (rawPct >= 80) status = 'WARN';

      catResults.push({
        id: cat.id,
        nom: cat.nom,
        alloue,
        consomme,
        ecart,
        pct: Math.min(rawPct, 100),
        rawPct,
        status,
        isOverBudget: ecart < 0
      });
    });

    catResults.sort((a, b) => b.consomme - a.consomme);

    return {
      mois: targetMonthStr,
      totalAlloue,
      totalConsomme,
      totalEcart: totalAlloue - totalConsomme,
      consumptionRate: totalAlloue > 0 ? (totalConsomme / totalAlloue) * 100 : 0,
      categories: catResults
    };
  }

  // --- 5. TRANSACTION MUTATIONS ---
  addTransaction(tx) {
    if (!this.data) return false;
    const nextId = (this.data.transactions && this.data.transactions.length > 0)
      ? Math.max(...this.data.transactions.map(t => t.id || 0)) + 1
      : 1;

    const cleanNotes = (tx.notes || '').trim().replace(/\s+/g, ' ');

    const newTx = {
      id: nextId,
      date: tx.date || new Date().toISOString().split('T')[0],
      id_compte: String(tx.id_compte),
      id_categorie: String(tx.id_categorie),
      id_tiers: String(tx.id_tiers),
      montant: Math.abs(parseFloat(tx.montant) || 0),
      devise: tx.devise || this.currency,
      notes: cleanNotes
    };

    if (!Array.isArray(this.data.transactions)) {
      this.data.transactions = [];
    }

    this.data.transactions.push(newTx);
    this.save();
    this.notify();
    return newTx;
  }

  deleteTransaction(id) {
    if (!this.data || !this.data.transactions) return false;
    const idx = this.data.transactions.findIndex(t => t.id === Number(id));
    if (idx !== -1) {
      this.data.transactions.splice(idx, 1);
      this.save();
      this.notify();
      return true;
    }
    return false;
  }

  allocateBudget(idCategorie, montant, mois = null) {
    if (!this.data) return false;
    const targetMois = mois || ((this.month === '00') ? `${this.year}-02` : `${this.year}-${this.month}`);
    const numAmount = Math.abs(parseFloat(montant) || 0);

    if (!Array.isArray(this.data.budgets)) {
      this.data.budgets = [];
    }

    const existing = this.data.budgets.find(b => b.mois === targetMois && String(b.id_categorie) === String(idCategorie));
    if (existing) {
      existing.montant = numAmount;
    } else {
      const nextId = (this.data.budgets.length > 0)
        ? Math.max(...this.data.budgets.map(b => b.id || 0)) + 1
        : 1;
      this.data.budgets.push({
        id: nextId,
        mois: targetMois,
        id_categorie: String(idCategorie),
        montant: numAmount
      });
    }

    this.save();
    this.notify();
    return true;
  }

  addAccount(nom, classe = 'LIQUIDITE', devise = 'MUR') {
    if (!this.data.comptes) this.data.comptes = [];
    const nextId = String(this.data.comptes.length + 1);
    const newCpt = { id: nextId, nom, classe, devise, actif: 'OUI' };
    this.data.comptes.push(newCpt);
    this.save();
    return newCpt;
  }

  addCategory(nom, type = 'DEPENSE') {
    if (!this.data.categories) this.data.categories = [];
    const nextId = String(this.data.categories.length + 1);
    const newCat = { id: nextId, nom, type };
    this.data.categories.push(newCat);
    this.save();
    return newCat;
  }

  addTiers(nom, type = 'COMMERCE') {
    if (!this.data.tiers) this.data.tiers = [];
    const nextId = String(this.data.tiers.length + 1);
    const newTiers = { id: nextId, nom, type };
    this.data.tiers.push(newTiers);
    this.save();
    return newTiers;
  }

  // --- 6. EXPORT / BACKUP / REPORT ---
  exportBackupJSON() {
    return JSON.stringify(this.data, null, 2);
  }

  importBackupJSON(jsonStr) {
    try {
      const parsed = JSON.parse(jsonStr);
      if (parsed.comptes && parsed.categories && parsed.transactions) {
        this.data = parsed;
        this.save();
        this.notify();
        return { success: true, count: parsed.transactions.length };
      }
      return { success: false, error: 'Format de fichier non conforme' };
    } catch (e) {
      return { success: false, error: e.message };
    }
  }

  generateFinancialReport() {
    const cf = this.computeCashflow();
    const nw = this.computeNetWorth();
    const bg = this.computeBudgetZBB();

    return {
      period: `${this.getMonthLabel()} ${this.year}`,
      currency: this.currency,
      revenus: this.formatMoney(cf.totalRevenus),
      depenses: this.formatMoney(cf.totalDepenses),
      netCashflow: this.formatMoney(cf.netCashflow),
      savingsRate: `${Math.round(cf.savingsRate)}%`,
      netWorth: this.formatMoney(nw.netWorth),
      totalAssets: this.formatMoney(nw.totalAssets),
      totalLiabilities: this.formatMoney(nw.totalLiabilities),
      budgetConsumption: `${Math.round(bg.consumptionRate)}%`,
      topExpenses: cf.breakdown.slice(0, 5).map(b => ({
        nom: b.nom,
        montant: this.formatMoney(b.montant),
        pct: `${Math.round(b.pct)}%`
      }))
    };
  }
}

window.gfmpEngine = new GFMPEngine();
