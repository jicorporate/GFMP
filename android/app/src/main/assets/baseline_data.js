window.GFMP_BASELINE = {
  "app_name": "GFMP — Global Finance Management Platform",
  "app_version": "1.0",
  "base_currency": "MUR",
  "default_language": "FR",
  "comptes": [
    {
      "id": "1",
      "nom": "Compte Courant Principal",
      "classe": "LIQUIDITE",
      "devise": "MUR",
      "actif": "OUI"
    },
    {
      "id": "2",
      "nom": "Livret d'Épargne",
      "classe": "LIQUIDITE",
      "devise": "MUR",
      "actif": "OUI"
    },
    {
      "id": "3",
      "nom": "Portefeuille Espèces",
      "classe": "LIQUIDITE",
      "devise": "MUR",
      "actif": "OUI"
    },
    {
      "id": "4",
      "nom": "Assurance Vie",
      "classe": "INVESTISSEMENT",
      "devise": "EUR",
      "actif": "OUI"
    },
    {
      "id": "5",
      "nom": "PEA / Actions",
      "classe": "INVESTISSEMENT",
      "devise": "EUR",
      "actif": "OUI"
    },
    {
      "id": "6",
      "nom": "Portefeuille Crypto",
      "classe": "INVESTISSEMENT",
      "devise": "USD",
      "actif": "OUI"
    },
    {
      "id": "7",
      "nom": "Prêt Immobilier",
      "classe": "DETTE",
      "devise": "MUR",
      "actif": "OUI"
    },
    {
      "id": "8",
      "nom": "Carte de Crédit (Différé)",
      "classe": "DETTE",
      "devise": "MUR",
      "actif": "OUI"
    },
    {
      "id": "9",
      "nom": "Autre (Préciser...)",
      "classe": "AUTRE",
      "devise": "MUR",
      "actif": "OUI"
    }
  ],
  "categories": [
    {
      "id": "1",
      "nom": "Salaire / Revenus Pro",
      "type": "REVENU"
    },
    {
      "id": "2",
      "nom": "Intérêts / Dividendes",
      "type": "REVENU"
    },
    {
      "id": "3",
      "nom": "Aides / Bourses / Allocations",
      "type": "REVENU"
    },
    {
      "id": "4",
      "nom": "Logement (Loyer/Prêt/Charges)",
      "type": "DEPENSE"
    },
    {
      "id": "5",
      "nom": "Alimentation & Supermarché",
      "type": "DEPENSE"
    },
    {
      "id": "6",
      "nom": "Transports (Essence/Assurance)",
      "type": "DEPENSE"
    },
    {
      "id": "7",
      "nom": "Santé & Mutuelle",
      "type": "DEPENSE"
    },
    {
      "id": "8",
      "nom": "Loisirs, Sorties & Vacances",
      "type": "DEPENSE"
    },
    {
      "id": "9",
      "nom": "Impôts & Taxes",
      "type": "DEPENSE"
    },
    {
      "id": "10",
      "nom": "Virement Interne (Épargne)",
      "type": "TRANSFERT"
    },
    {
      "id": "11",
      "nom": "Autre (Préciser...)",
      "type": "AUTRE"
    }
  ],
  "tiers": [
    {
      "id": "1",
      "nom": "Employeur principal",
      "type": "ENTREPRISE"
    },
    {
      "id": "2",
      "nom": "Banque / Courtier",
      "type": "FINANCE"
    },
    {
      "id": "3",
      "nom": "État / Impôts",
      "type": "INSTITUTION"
    },
    {
      "id": "4",
      "nom": "Supermarché / Alimentaire",
      "type": "COMMERCE"
    },
    {
      "id": "5",
      "nom": "Propriétaire / Syndic",
      "type": "IMMOBILIER"
    },
    {
      "id": "6",
      "nom": "Station Service",
      "type": "COMMERCE"
    },
    {
      "id": "7",
      "nom": "Autre (Préciser...)",
      "type": "AUTRE"
    }
  ],
  "transactions": [
    {
      "id": 1,
      "date": "2026-01-25",
      "id_compte": "1",
      "id_categorie": "1",
      "id_tiers": "1",
      "montant": 78000.0,
      "devise": "MUR",
      "notes": "Salaire Mensuel KIMYA"
    },
    {
      "id": 2,
      "date": "2026-01-26",
      "id_compte": "1",
      "id_categorie": "4",
      "id_tiers": "5",
      "montant": 12000.0,
      "devise": "MUR",
      "notes": "Loyer Résidence Principale"
    },
    {
      "id": 3,
      "date": "2026-01-28",
      "id_compte": "1",
      "id_categorie": "5",
      "id_tiers": "4",
      "montant": 6200.0,
      "devise": "MUR",
      "notes": "Supermarché & Courses"
    },
    {
      "id": 4,
      "date": "2026-01-29",
      "id_compte": "1",
      "id_categorie": "6",
      "id_tiers": "4",
      "montant": 3400.0,
      "devise": "MUR",
      "notes": "Carburant & Péages"
    },
    {
      "id": 5,
      "date": "2026-01-30",
      "id_compte": "1",
      "id_categorie": "10",
      "id_tiers": "2",
      "montant": 25000.0,
      "devise": "MUR",
      "notes": "Virement vers Livret Épargne"
    },
    {
      "id": 6,
      "date": "2026-02-05",
      "id_compte": "1",
      "id_categorie": "4",
      "id_tiers": "5",
      "montant": 12000.0,
      "devise": "MUR",
      "notes": "Loyer Résidence KIMYA"
    },
    {
      "id": 7,
      "date": "2026-02-12",
      "id_compte": "5",
      "id_categorie": "2",
      "id_tiers": "2",
      "montant": 4500.0,
      "devise": "MUR",
      "notes": "Dividendes Portefeuille PEA"
    },
    {
      "id": 8,
      "date": "2026-02-15",
      "id_compte": "1",
      "id_categorie": "6",
      "id_tiers": "4",
      "montant": 2800.0,
      "devise": "MUR",
      "notes": "Assurance & Entretien Auto"
    },
    {
      "id": 9,
      "date": "2026-02-18",
      "id_compte": "1",
      "id_categorie": "8",
      "id_tiers": "4",
      "montant": 3500.0,
      "devise": "MUR",
      "notes": "Sorties & Restaurants"
    },
    {
      "id": 10,
      "date": "2026-02-23",
      "id_compte": "1",
      "id_categorie": "1",
      "id_tiers": "1",
      "montant": 78000.0,
      "devise": "MUR",
      "notes": "Salaire Février KIMYA ÉDITION"
    },
    {
      "id": 11,
      "date": "2026-02-23",
      "id_compte": "4",
      "id_categorie": "7",
      "id_tiers": "2",
      "montant": 7000.0,
      "devise": "MUR",
      "notes": "Santé & Mutuelle Famille"
    },
    {
      "id": 12,
      "date": "2026-02-23",
      "id_compte": "3",
      "id_categorie": "5",
      "id_tiers": "4",
      "montant": 4000.0,
      "devise": "MUR",
      "notes": "Courses Hebdomadaires"
    }
  ],
  "budgets": [
    {
      "id": 1,
      "mois": "2026-01",
      "id_categorie": "4",
      "montant": 12000.0
    },
    {
      "id": 2,
      "mois": "2026-01",
      "id_categorie": "5",
      "montant": 8000.0
    },
    {
      "id": 3,
      "mois": "2026-01",
      "id_categorie": "6",
      "montant": 4000.0
    },
    {
      "id": 4,
      "mois": "2026-01",
      "id_categorie": "7",
      "montant": 6000.0
    },
    {
      "id": 5,
      "mois": "2026-01",
      "id_categorie": "8",
      "montant": 5000.0
    },
    {
      "id": 6,
      "mois": "2026-02",
      "id_categorie": "4",
      "montant": 12000.0
    },
    {
      "id": 7,
      "mois": "2026-02",
      "id_categorie": "5",
      "montant": 8000.0
    },
    {
      "id": 8,
      "mois": "2026-02",
      "id_categorie": "6",
      "montant": 4000.0
    },
    {
      "id": 9,
      "mois": "2026-02",
      "id_categorie": "7",
      "montant": 7000.0
    },
    {
      "id": 10,
      "mois": "2026-02",
      "id_categorie": "8",
      "montant": 5000.0
    }
  ],
  "devises": {
    "MUR": 12.2243,
    "EUR": 656.168,
    "USD": 567.5369,
    "GBP": 757.0023,
    "ZAR": 33.2723,
    "XOF": 1.0
  },
  "dictionary": {
    "APP_TITLE": {
      "FR": "SYSTÈME FINANCIER PERSONNEL",
      "EN": "PERSONAL FINANCE SYSTEM",
      "ES": "SISTEMA FINANCIERO PERSONAL",
      "PT": "SISTEMA FINANCEIRO PESSOAL",
      "DE": "PERSÖNLICHES FINANZSYSTEM",
      "IT": "SISTEMA FINANZIARIO",
      "NL": "FINANCIEEL SYSTEEM",
      "SV": "FINANSSYSTEM"
    },
    "HUB_LOC": {
      "FR": "Hub Central",
      "EN": "Central Hub",
      "ES": "Centro Principal",
      "PT": "Hub Central",
      "DE": "Zentraler Hub",
      "IT": "Hub Centrale",
      "NL": "Centrale Hub",
      "SV": "Central Hub"
    },
    "SAISIE_T": {
      "FR": "SAISIE TRANSACTION",
      "EN": "ENTER TRANSACTION",
      "ES": "INGRESAR TRANSACCIÓN",
      "PT": "INSERIR TRANSAÇÃO",
      "DE": "TRANSAKTION ERFASSEN",
      "IT": "INSERISCI TRANSAZIONE",
      "NL": "TRANSACTIE INVOEREN",
      "SV": "ANGE TRANSAKTION"
    },
    "SAISIE_D": {
      "FR": "Ajouter un revenu, dépense ou virement.",
      "EN": "Add an income, expense, or transfer.",
      "ES": "Añadir ingreso, gasto o transferencia.",
      "PT": "Adicionar renda, despesa ou transferência.",
      "DE": "Einkommen, Ausgabe oder Transfer.",
      "IT": "Aggiungi entrata, uscita o bonifico.",
      "NL": "Voeg inkomsten, uitgaven of overboeking toe.",
      "SV": "Lägg till inkomst, utgift eller överföring."
    },
    "DASH_T": {
      "FR": "DASHBOARD CASHFLOW",
      "EN": "CASHFLOW DASHBOARD",
      "ES": "PANEL DE FLUJO DE CAJA",
      "PT": "PAINEL DE FLUXO DE CAIXA",
      "DE": "CASHFLOW-DASHBOARD",
      "IT": "DASHBOARD FLUSSI",
      "NL": "CASHFLOW DASHBOARD",
      "SV": "CASHFLOW DASHBOARD"
    },
    "DASH_D": {
      "FR": "Analyser les flux consolidés.",
      "EN": "Analyze consolidated flows.",
      "ES": "Analizar flujos consolidados.",
      "PT": "Analisar fluxos consolidados.",
      "DE": "Konsolidierte Flüsse analysieren.",
      "IT": "Analizza i flussi consolidati.",
      "NL": "Analyseer de stromen.",
      "SV": "Analysera flöden."
    },
    "BUDG_T": {
      "FR": "PILOTAGE BUDGÉTAIRE",
      "EN": "BUDGET TRACKING",
      "ES": "CONTROL PRESUPUESTARIO",
      "PT": "CONTROLE ORÇAMENTÁRIO",
      "DE": "BUDGETKONTROLLE",
      "IT": "CONTROLLO BUDGET",
      "NL": "BUDGETBEHEER",
      "SV": "BUDGETKONTROLL"
    },
    "BUDG_D": {
      "FR": "Suivi des enveloppes (ZBB).",
      "EN": "Track budget envelopes.",
      "ES": "Seguimiento de presupuestos.",
      "PT": "Acompanhamento de orçamentos.",
      "DE": "Verfolgung der Budgets.",
      "IT": "Traccia i budget.",
      "NL": "Volg uw budgetten.",
      "SV": "Spåra dina budgetar."
    },
    "NETW_T": {
      "FR": "BILAN PATRIMONIAL",
      "EN": "NET WORTH STATEMENT",
      "ES": "BALANCE PATRIMONIAL",
      "PT": "BALANÇO PATRIMONIAL",
      "DE": "VERMÖGENSBILANZ",
      "IT": "BILANCIO PATRIMONIALE",
      "NL": "VERMOGENSOVERZICHT",
      "SV": "FÖRMÖGENHETSRAPPORT"
    },
    "NETW_D": {
      "FR": "Calcul de la Valeur Nette.",
      "EN": "Net Worth calculation.",
      "ES": "Cálculo del patrimonio neto.",
      "PT": "Cálculo do patrimônio líquido.",
      "DE": "Berechnung des Nettovermögens.",
      "IT": "Calcolo del patrimonio netto.",
      "NL": "Berekening nettowaarde.",
      "SV": "Beräkning av nettovärde."
    },
    "WELCOME": {
      "FR": "Sélectionnez un module d'application ci-dessous pour démarrer.",
      "EN": "Select an application module below to get started.",
      "ES": "Seleccione un módulo de aplicación a continuación.",
      "PT": "Selecione um módulo de aplicativo abaixo.",
      "DE": "Wählen Sie unten ein Anwendungsmodul aus.",
      "IT": "Seleziona un modulo dell'applicazione di seguito.",
      "NL": "Selecteer hieronder een applicatiemodule.",
      "SV": "Välj en applikationsmodul nedan."
    },
    "TT_LANG": {
      "FR": "Changer la langue :",
      "EN": "Change language :",
      "ES": "Cambiar idioma :",
      "PT": "Mudar idioma :",
      "DE": "Sprache ändern :",
      "IT": "Cambia lingua :",
      "NL": "Taal wijzigen :",
      "SV": "Ändra språk :"
    },
    "BTN_SYNC": {
      "FR": "ACTUALISER TAUX FX",
      "EN": "UPDATE FX RATES",
      "ES": "ACTUALIZAR TIPOS FX",
      "PT": "ATUALIZAR TAXAS FX",
      "DE": "FX-KURSE AKTUALISIEREN",
      "IT": "AGGIORNA TASSI FX",
      "NL": "FX-TARIEVEN BIJWERKEN",
      "SV": "UPPDATERA FX-KURSER"
    },
    "Salaire / Revenus Pro": {
      "FR": "Salaire / Revenus Pro",
      "EN": "Salary / Pro Income",
      "ES": "Salario / Ingresos",
      "PT": "Salário / Renda",
      "DE": "Gehalt / Einkommen",
      "IT": "Stipendio / Reddito",
      "NL": "Salaris / Inkomen",
      "SV": "Lön / Inkomst"
    },
    "Intérêts / Dividendes": {
      "FR": "Intérêts / Dividendes",
      "EN": "Interests / Dividends",
      "ES": "Intereses / Dividendos",
      "PT": "Juros / Dividendos",
      "DE": "Zinsen / Dividenden",
      "IT": "Interessi / Dividendi",
      "NL": "Rente / Dividenden",
      "SV": "Ränta / Utdelning"
    },
    "Logement (Loyer/Prêt/Charges)": {
      "FR": "Logement",
      "EN": "Housing (Rent/Loan)",
      "ES": "Vivienda",
      "PT": "Moradia",
      "DE": "Wohnen",
      "IT": "Abitazione",
      "NL": "Huisvesting",
      "SV": "Bostad"
    },
    "Alimentation & Supermarché": {
      "FR": "Alimentation",
      "EN": "Groceries",
      "ES": "Alimentación",
      "PT": "Alimentação",
      "DE": "Lebensmittel",
      "IT": "Alimentari",
      "NL": "Boodschappen",
      "SV": "Mat"
    },
    "Transports (Essence/Assurance)": {
      "FR": "Transports",
      "EN": "Transport",
      "ES": "Transporte",
      "PT": "Transporte",
      "DE": "Transport",
      "IT": "Trasporti",
      "NL": "Vervoer",
      "SV": "Transport"
    },
    "Santé & Mutuelle": {
      "FR": "Santé & Mutuelle",
      "EN": "Health & Insurance",
      "ES": "Salud",
      "PT": "Saúde",
      "DE": "Gesundheit",
      "IT": "Salute",
      "NL": "Gezondheid",
      "SV": "Hälsa"
    },
    "Loisirs, Sorties & Vacances": {
      "FR": "Loisirs & Vacances",
      "EN": "Leisure & Holidays",
      "ES": "Ocio",
      "PT": "Lazer",
      "DE": "Freizeit",
      "IT": "Tempo Libero",
      "NL": "Vrije Tijd",
      "SV": "Fritid"
    },
    "Virement Interne (Épargne)": {
      "FR": "Virement Interne",
      "EN": "Internal Transfer",
      "ES": "Transferencia Interna",
      "PT": "Transferência Interna",
      "DE": "Interner Transfer",
      "IT": "Bonifico Interno",
      "NL": "Interne Overboeking",
      "SV": "Intern Överföring"
    },
    "Compte Courant Principal": {
      "FR": "Compte Courant",
      "EN": "Checking Account",
      "ES": "Cuenta Corriente",
      "PT": "Conta Corrente",
      "DE": "Girokonto",
      "IT": "Conto Corrente",
      "NL": "Betaalrekening",
      "SV": "Lönekonto"
    },
    "Livret d'Épargne": {
      "FR": "Livret d'Épargne",
      "EN": "Savings Account",
      "ES": "Cuenta de Ahorros",
      "PT": "Conta Poupança",
      "DE": "Sparkonto",
      "IT": "Conto Risparmio",
      "NL": "Spaarrekening",
      "SV": "Sparkonto"
    },
    "Carte de Crédit (Différé)": {
      "FR": "Carte de Crédit",
      "EN": "Credit Card",
      "ES": "Tarjeta de Crédito",
      "PT": "Cartão de Crédito",
      "DE": "Kreditkarte",
      "IT": "Carta di Credito",
      "NL": "Creditcard",
      "SV": "Kreditkort"
    },
    "BTN_SETTING": {
      "FR": "PARAMÈTRES (DEVISE)",
      "EN": "SETTINGS (CURRENCY)",
      "ES": "AJUSTES (DIVISA)",
      "PT": "CONFIG. (MOEDA)",
      "DE": "EINSTELLUNGEN",
      "IT": "IMPOSTAZIONI",
      "NL": "INSTELLINGEN",
      "SV": "INSTÄLLNINGAR"
    },
    "MSG_ASK_BASE": {
      "FR": "Saisissez votre Devise Principale (ex: XOF, EUR, USD, CAD) :",
      "EN": "Enter your Base Currency (e.g. USD, EUR) :",
      "ES": "Ingrese su divisa base :",
      "PT": "Insira sua moeda base :",
      "DE": "Basiswährung eingeben :",
      "IT": "Inserisci valuta base :",
      "NL": "Voer basisvaluta in :",
      "SV": "Ange basvaluta :"
    },
    "LIQUIDITE": {
      "FR": "LIQUIDITÉ",
      "EN": "LIQUIDITY",
      "ES": "LIQUIDEZ",
      "PT": "LIQUIDEZ",
      "DE": "LIQUIDITÄT",
      "IT": "LIQUIDITÀ",
      "NL": "LIQUIDITEIT",
      "SV": "LIKVIDITET"
    },
    "INVESTISSEMENT": {
      "FR": "INVESTISSEMENT",
      "EN": "INVESTMENT",
      "ES": "INVERSIÓN",
      "PT": "INVESTIMENTO",
      "DE": "INVESTITION",
      "IT": "INVESTIMENTO",
      "NL": "INVESTERING",
      "SV": "INVESTERING"
    },
    "DETTE": {
      "FR": "DETTE",
      "EN": "DEBT",
      "ES": "DEUDA",
      "PT": "DÍVIDA",
      "DE": "SCHULD",
      "IT": "DEBITO",
      "NL": "SCHULD",
      "SV": "SKULD"
    },
    "DEPENSE": {
      "FR": "DÉPENSE",
      "EN": "EXPENSE",
      "ES": "GASTO",
      "PT": "DESPESA",
      "DE": "AUSGABE",
      "IT": "USCITA",
      "NL": "UITGAVE",
      "SV": "UTGIFT"
    },
    "REVENU": {
      "FR": "REVENU",
      "EN": "INCOME",
      "ES": "INGRESO",
      "PT": "RENDA",
      "DE": "EINKOMMEN",
      "IT": "ENTRATA",
      "NL": "INKOMSTEN",
      "SV": "INKOMST"
    },
    "TRANSFERT": {
      "FR": "TRANSFERT",
      "EN": "TRANSFER",
      "ES": "TRANSFERENCIA",
      "PT": "TRANSFERÊNCIA",
      "DE": "TRANSFER",
      "IT": "TRASFERIMENTO",
      "NL": "OVERDRACHT",
      "SV": "ÖVERFÖRING"
    },
    "AUTRE": {
      "FR": "AUTRE",
      "EN": "OTHER",
      "ES": "OTRO",
      "PT": "OUTRO",
      "DE": "ANDERE",
      "IT": "ALTRO",
      "NL": "ANDERS",
      "SV": "ANNAT"
    },
    "Autre (Préciser...)": {
      "FR": "Autre (Préciser...)",
      "EN": "Other (Specify...)",
      "ES": "Otro (Especificar...)",
      "PT": "Outro (Especificar...)",
      "DE": "Andere (Angeben...)",
      "IT": "Altro (Specificare...)",
      "NL": "Anders (Specificeren...)",
      "SV": "Annat (Ange...)"
    },
    "Employeur principal": {
      "FR": "Employeur",
      "EN": "Main Employer",
      "ES": "Empleador",
      "PT": "Empregador",
      "DE": "Arbeitgeber",
      "IT": "Datore di Lavoro",
      "NL": "Werkgever",
      "SV": "Arbetsgivare"
    },
    "Banque / Courtier": {
      "FR": "Banque / Courtier",
      "EN": "Bank / Broker",
      "ES": "Banco / Bróker",
      "PT": "Banco / Corretor",
      "DE": "Bank / Makler",
      "IT": "Banca / Broker",
      "NL": "Bank / Makelaar",
      "SV": "Bank / Mäklare"
    },
    "État / Impôts": {
      "FR": "État / Impôts",
      "EN": "State / Taxes",
      "ES": "Estado / Impuestos",
      "PT": "Estado / Impostos",
      "DE": "Staat / Steuern",
      "IT": "Stato / Tasse",
      "NL": "Staat / Belastingen",
      "SV": "Stat / Skatt"
    },
    "Supermarché / Alimentaire": {
      "FR": "Supermarché",
      "EN": "Supermarket",
      "ES": "Supermercado",
      "PT": "Supermercado",
      "DE": "Supermarkt",
      "IT": "Supermercato",
      "NL": "Supermarkt",
      "SV": "Stormarknad"
    },
    "Propriétaire / Syndic": {
      "FR": "Propriétaire",
      "EN": "Landlord / HOA",
      "ES": "Propietario",
      "PT": "Senhorio",
      "DE": "Vermieter",
      "IT": "Proprietario",
      "NL": "Verhuurder",
      "SV": "Hyresvärd"
    },
    "Station Service": {
      "FR": "Station Service",
      "EN": "Gas Station",
      "ES": "Gasolinera",
      "PT": "Posto de Gasolina",
      "DE": "Tankstelle",
      "IT": "Stazione di Servizio",
      "NL": "Benzinestation",
      "SV": "Bensinstation"
    },
    "MSG_FX_OK": {
      "FR": "TAUX DE CHANGE SYNCHRONISÉS._x000D_\nMise à jour réussie.",
      "EN": "EXCHANGE RATES SYNCHRONIZED._x000D_\nUpdate successful.",
      "ES": "TIPOS DE CAMBIO SINCRONIZADOS._x000D_\nActualización exitosa.",
      "PT": "TAXAS DE CÂMBIO SINCRONIZADAS._x000D_\nAtualização bem-sucedida.",
      "DE": "WECHSELKURSE SYNCHRONISIERT._x000D_\nErfolgreich aktualisiert.",
      "IT": "TASSI DI CAMBIO SINCRONIZZATI._x000D_\nAggiornamento riuscito.",
      "NL": "WISSELKOERSEN GESYNCHRONISEERD._x000D_\nSuccesvol bijgewerkt.",
      "SV": "VÄXELKURSER SYNKRONISERADE._x000D_\nUppdatering lyckades."
    },
    "MSG_FX_ERR": {
      "FR": "Erreur de connexion au serveur de devises.",
      "EN": "Connection error to currency server.",
      "ES": "Error de conexión al servidor de divisas.",
      "PT": "Erro de conexão ao servidor de moedas.",
      "DE": "Verbindungsfehler zum Währungsserver.",
      "IT": "Errore di connessione al server delle valute.",
      "NL": "Verbindingsfout met valutaserver.",
      "SV": "Anslutningsfel till valutaserver."
    },
    "MSG_FX_TITLE": {
      "FR": "Synchronisation FX",
      "EN": "FX Synchronization",
      "ES": "Sincronización FX",
      "PT": "Sincronização FX",
      "DE": "FX-Synchronisation",
      "IT": "Sincronizzazione FX",
      "NL": "FX-synchronisatie",
      "SV": "FX-synkronisering"
    },
    "MSG_ERR_DEV_API": {
      "FR": "Devise invalide ou non reconnue par le marché boursier.",
      "EN": "Invalid currency or unrecognized by the market.",
      "ES": "Divisa inválida o no reconocida.",
      "PT": "Moeda inválida ou não reconhecida.",
      "DE": "Ungültige oder nicht erkannte Währung.",
      "IT": "Valuta non valida o non riconosciuta.",
      "NL": "Ongeldige valuta of niet herkend.",
      "SV": "Ogiltig valuta eller okänd."
    },
    "Prêt Immobilier": {
      "FR": "Prêt Immobilier",
      "EN": "Mortgage Loan",
      "ES": "Préstamo Hipotecario",
      "PT": "Empréstimo Imobiliário",
      "DE": "Immobilienkredit",
      "IT": "Mutuo Immobiliare",
      "NL": "Hypothecaire Lening",
      "SV": "Bolån"
    },
    "Impôts & Taxes": {
      "FR": "Impôts & Taxes",
      "EN": "Taxes",
      "ES": "Impuestos",
      "PT": "Impostos",
      "DE": "Steuern",
      "IT": "Tasse",
      "NL": "Belastingen",
      "SV": "Skatter"
    },
    "Portefeuille Espèces": {
      "FR": "Portefeuille Espèces",
      "EN": "Cash Wallet",
      "ES": "Cartera de Efectivo",
      "PT": "Carteira de Dinheiro",
      "DE": "Bargeldbörse",
      "IT": "Portafoglio Contanti",
      "NL": "Portemonnee",
      "SV": "Kontanter"
    },
    "Assurance Vie": {
      "FR": "Assurance Vie",
      "EN": "Life Insurance",
      "ES": "Seguro de Vida",
      "PT": "Seguro de Vida",
      "DE": "Lebensversicherung",
      "IT": "Assicurazione Vita",
      "NL": "Levensverzekering",
      "SV": "Livförsäkring"
    },
    "PEA / Actions": {
      "FR": "PEA / Actions",
      "EN": "Stock Portfolio",
      "ES": "Cartera de Acciones",
      "PT": "Carteira de Ações",
      "DE": "Aktienportfolio",
      "IT": "Portafoglio Azioni",
      "NL": "Aandelenportefeuille",
      "SV": "Aktieportfölj"
    },
    "Portefeuille Crypto": {
      "FR": "Portefeuille Crypto",
      "EN": "Crypto Wallet",
      "ES": "Cartera Crypto",
      "PT": "Carteira Crypto",
      "DE": "Krypto-Wallet",
      "IT": "Portafoglio Crypto",
      "NL": "Cryptoportemonnee",
      "SV": "Kryptoplånbok"
    },
    "FRM_TITLE": {
      "FR": "Saisie de Transaction",
      "EN": "Transaction Entry",
      "ES": "Ingreso de Transacción",
      "PT": "Registro de Transação",
      "DE": "Transaktion Erfassen",
      "IT": "Inserimento Transazione",
      "NL": "Transactie Invoer",
      "SV": "Transaktionsinmatning"
    },
    "FRM_DATE": {
      "FR": "Date (MM/JJ/AAAA) :",
      "EN": "Date (MM/DD/YYYY) :",
      "ES": "Fecha (MM/DD/AAAA) :",
      "PT": "Data (MM/DD/AAAA) :",
      "DE": "Datum (MM/TT/JJJJ) :",
      "IT": "Data (MM/GG/AAAA) :",
      "NL": "Datum (MM/DD/JJJJ) :",
      "SV": "Datum (MM/DD/ÅÅÅÅ) :"
    },
    "FRM_COMPTE": {
      "FR": "Compte :",
      "EN": "Account :",
      "ES": "Cuenta :",
      "PT": "Conta :",
      "DE": "Konto :",
      "IT": "Conto :",
      "NL": "Rekening :",
      "SV": "Konto :"
    },
    "FRM_CAT": {
      "FR": "Catégorie :",
      "EN": "Category :",
      "ES": "Categoría :",
      "PT": "Categoria :",
      "DE": "Kategorie :",
      "IT": "Categoria :",
      "NL": "Categorie :",
      "SV": "Kategori :"
    },
    "FRM_TIERS": {
      "FR": "Tiers :",
      "EN": "Payee/Payer :",
      "ES": "Tercero :",
      "PT": "Terceiro :",
      "DE": "Partei :",
      "IT": "Terzo :",
      "NL": "Partij :",
      "SV": "Part :"
    },
    "FRM_MONTANT": {
      "FR": "Montant :",
      "EN": "Amount :",
      "ES": "Monto :",
      "PT": "Valor :",
      "DE": "Betrag :",
      "IT": "Importo :",
      "NL": "Bedrag :",
      "SV": "Belopp :"
    },
    "FRM_DEVISE": {
      "FR": "Devise :",
      "EN": "Currency :",
      "ES": "Divisa :",
      "PT": "Moeda :",
      "DE": "Währung :",
      "IT": "Valuta :",
      "NL": "Valuta :",
      "SV": "Valuta :"
    },
    "FRM_DESC": {
      "FR": "Notes :",
      "EN": "Notes :",
      "ES": "Notas :",
      "PT": "Notas :",
      "DE": "Notizen :",
      "IT": "Note :",
      "NL": "Notities :",
      "SV": "Anteckningar :"
    },
    "FRM_NEW": {
      "FR": "Précisez :",
      "EN": "Specify :",
      "ES": "Especificar :",
      "PT": "Especificar :",
      "DE": "Angeben :",
      "IT": "Specifica :",
      "NL": "Specificeer :",
      "SV": "Ange :"
    },
    "FRM_SAVE": {
      "FR": "ENREGISTRER",
      "EN": "SAVE",
      "ES": "GUARDAR",
      "PT": "SALVAR",
      "DE": "SPEICHERN",
      "IT": "SALVA",
      "NL": "OPSLAAN",
      "SV": "SPARA"
    },
    "FRM_CANCEL": {
      "FR": "ANNULER",
      "EN": "CANCEL",
      "ES": "CANCELAR",
      "PT": "CANCELAR",
      "DE": "ABBRECHEN",
      "IT": "ANNULLA",
      "NL": "ANNULEREN",
      "SV": "AVBRYT"
    },
    "MSG_ERR_MISSING": {
      "FR": "Sélection incomplète.",
      "EN": "Missing selection.",
      "ES": "Selección incompleta.",
      "PT": "Seleção incompleta.",
      "DE": "Fehlende Auswahl.",
      "IT": "Selezione incompleta.",
      "NL": "Ontbrekende selectie.",
      "SV": "Saknad markering."
    },
    "MSG_ERR_AMT": {
      "FR": "Montant invalide.",
      "EN": "Invalid amount.",
      "ES": "Monto inválido.",
      "PT": "Valor inválido.",
      "DE": "Ungültiger Betrag.",
      "IT": "Importo non valido.",
      "NL": "Ongeldig bedrag.",
      "SV": "Ogiltigt belopp."
    },
    "MSG_OK": {
      "FR": "Enregistré avec succès !",
      "EN": "Saved successfully!",
      "ES": "¡Guardado con éxito!",
      "PT": "Salvo com sucesso!",
      "DE": "Erfolgreich gespeichert!",
      "IT": "Salvato con successo!",
      "NL": "Succesvol opgeslagen!",
      "SV": "Sparad!"
    },
    "TT_F_DATE": {
      "FR": "Saisissez la date.",
      "EN": "Enter date.",
      "ES": "Ingrese fecha.",
      "PT": "Insira data.",
      "DE": "Datum eingeben.",
      "IT": "Inserisci data.",
      "NL": "Voer datum in.",
      "SV": "Ange datum."
    },
    "TT_F_COMPTE": {
      "FR": "Choisissez le compte.",
      "EN": "Choose account.",
      "ES": "Elija cuenta.",
      "PT": "Escolha conta.",
      "DE": "Konto wählen.",
      "IT": "Scegli conto.",
      "NL": "Kies rekening.",
      "SV": "Välj konto."
    },
    "TT_F_CAT": {
      "FR": "Sélectionnez la catégorie.",
      "EN": "Select category.",
      "ES": "Seleccione categoría.",
      "PT": "Selecione categoria.",
      "DE": "Kategorie wählen.",
      "IT": "Seleziona categoria.",
      "NL": "Selecteer categorie.",
      "SV": "Välj kategori."
    },
    "TT_F_TIERS": {
      "FR": "Tiers lié.",
      "EN": "Related party.",
      "ES": "Tercero relacionado.",
      "PT": "Terceiro relacionado.",
      "DE": "Zugehörige Partei.",
      "IT": "Parte correlata.",
      "NL": "Gerelateerde partij.",
      "SV": "Relaterad part."
    },
    "TT_F_MONTANT": {
      "FR": "Valeur absolue.",
      "EN": "Absolute value.",
      "ES": "Valor absoluto.",
      "PT": "Valor absoluto.",
      "DE": "Absoluter Wert.",
      "IT": "Valore assoluto.",
      "NL": "Absolute waarde.",
      "SV": "Absolut värde."
    },
    "TT_F_DEVISE": {
      "FR": "Monnaie.",
      "EN": "Currency.",
      "ES": "Moneda.",
      "PT": "Moeda.",
      "DE": "Währung.",
      "IT": "Valuta.",
      "NL": "Valuta.",
      "SV": "Valuta."
    },
    "TT_F_DESC": {
      "FR": "Notes.",
      "EN": "Notes.",
      "ES": "Notas.",
      "PT": "Notas.",
      "DE": "Notizen.",
      "IT": "Note.",
      "NL": "Notities.",
      "SV": "Anteckningar."
    },
    "TT_F_SAVE": {
      "FR": "Enregistrer.",
      "EN": "Save.",
      "ES": "Guardar.",
      "PT": "Salvar.",
      "DE": "Speichern.",
      "IT": "Salva.",
      "NL": "Opslaan.",
      "SV": "Spara."
    },
    "TT_F_CANCEL": {
      "FR": "Annuler.",
      "EN": "Cancel.",
      "ES": "Cancelar.",
      "PT": "Cancelar.",
      "DE": "Abbrechen.",
      "IT": "Annulla.",
      "NL": "Annuleren.",
      "SV": "Avbryt."
    },
    "TT_F_TYPE_CPT": {
      "FR": "Type d'actif.",
      "EN": "Asset type.",
      "ES": "Tipo de activo.",
      "PT": "Tipo de ativo.",
      "DE": "Anlageklasse.",
      "IT": "Classe di attività.",
      "NL": "Activaklasse.",
      "SV": "Tillgångsklass."
    },
    "TT_F_TYPE_CAT": {
      "FR": "Type de flux.",
      "EN": "Flow type.",
      "ES": "Tipo de flujo.",
      "PT": "Tipo de fluxo.",
      "DE": "Flusstyp.",
      "IT": "Tipo di flusso.",
      "NL": "Stroomtype.",
      "SV": "Flödestyp."
    },
    "OPT_LIQ": {
      "FR": "LIQUIDITÉ",
      "EN": "LIQUIDITY",
      "ES": "LIQUIDEZ",
      "PT": "LIQUIDEZ",
      "DE": "LIQUIDITÄT",
      "IT": "LIQUIDITÀ",
      "NL": "LIQUIDITEIT",
      "SV": "LIKVIDITET"
    },
    "OPT_INV": {
      "FR": "INVESTISSEMENT",
      "EN": "INVESTMENT",
      "ES": "INVERSIÓN",
      "PT": "INVESTIMENTO",
      "DE": "INVESTITION",
      "IT": "INVESTIMENTO",
      "NL": "INVESTERING",
      "SV": "INVESTERING"
    },
    "OPT_DET": {
      "FR": "DETTE",
      "EN": "DEBT",
      "ES": "DEUDA",
      "PT": "DÍVIDA",
      "DE": "SCHULD",
      "IT": "DEBITO",
      "NL": "SCHULD",
      "SV": "SKULD"
    },
    "OPT_DEP": {
      "FR": "DÉPENSE",
      "EN": "EXPENSE",
      "ES": "GASTO",
      "PT": "DESPESA",
      "DE": "AUSGABE",
      "IT": "USCITA",
      "NL": "UITGAVE",
      "SV": "UTGIFT"
    },
    "OPT_REV": {
      "FR": "REVENU",
      "EN": "INCOME",
      "ES": "INGRESO",
      "PT": "RENDA",
      "DE": "EINKOMMEN",
      "IT": "ENTRATA",
      "NL": "INKOMSTEN",
      "SV": "INKOMST"
    },
    "OPT_TRA": {
      "FR": "TRANSFERT",
      "EN": "TRANSFER",
      "ES": "TRANSFERENCIA",
      "PT": "TRANSFERÊNCIA",
      "DE": "TRANSFER",
      "IT": "TRASFERIMENTO",
      "NL": "OVERDRACHT",
      "SV": "ÖVERFÖRING"
    },
    "FRM_DEST": {
      "FR": "Compte Destination :",
      "EN": "Destination Account :",
      "ES": "Cuenta Destino :",
      "PT": "Conta Destino :",
      "DE": "Zielkonto :",
      "IT": "Conto Destinazione :",
      "NL": "Doelrekening :",
      "SV": "Målkonto :"
    },
    "BUDG_TITLE": {
      "FR": "PILOTAGE BUDGÉTAIRE",
      "EN": "BUDGET TRACKING",
      "ES": "CONTROL PRESUPUESTARIO",
      "PT": "CONTROLE ORÇAMENTÁRIO",
      "DE": "BUDGETKONTROLLE",
      "IT": "CONTROLLO BUDGET",
      "NL": "BUDGETBEHEER",
      "SV": "BUDGETKONTROLL"
    },
    "BTN_ALLOC": {
      "FR": "ALLOUER BUDGET",
      "EN": "ALLOCATE BUDGET",
      "ES": "ASIGNAR PRESUPUESTO",
      "PT": "ALOCAR ORÇAMENTO",
      "DE": "BUDGET ZUWEISEN",
      "IT": "ASSEGNA BUDGET",
      "NL": "BUDGET TOEWIJZEN",
      "SV": "TILLDELA BUDGET"
    },
    "KPI_B_ALLO": {
      "FR": "TOTAL ALLOUÉ",
      "EN": "TOTAL ALLOCATED",
      "ES": "TOTAL ASIGNADO",
      "PT": "TOTAL ALOCADO",
      "DE": "GESAMT ZUGEWIESEN",
      "IT": "TOTALE ASSEGNATO",
      "NL": "TOTAAL TOEGEWEZEN",
      "SV": "TOTALT TILLDELAD"
    },
    "KPI_B_SPEN": {
      "FR": "TOTAL DÉPENSÉ",
      "EN": "TOTAL SPENT",
      "ES": "TOTAL GASTADO",
      "PT": "TOTAL GASTO",
      "DE": "GESAMT AUSGEGEBEN",
      "IT": "TOTALE SPESO",
      "NL": "TOTAAL UITGEGEVEN",
      "SV": "TOTALT SPENDERAT"
    },
    "KPI_B_LEFT": {
      "FR": "RESTE À DÉPENSER",
      "EN": "REMAINING",
      "ES": "RESTANTE",
      "PT": "RESTANTE",
      "DE": "VERBLEIBEND",
      "IT": "RIMANENTE",
      "NL": "RESTEREND",
      "SV": "ÅTERSTÅENDE"
    },
    "COL_B_CAT": {
      "FR": "ENVELOPPE (CATÉGORIE)",
      "EN": "ENVELOPE (CATEGORY)",
      "ES": "CATEGORÍA",
      "PT": "CATEGORIA",
      "DE": "KATEGORIE",
      "IT": "CATEGORIA",
      "NL": "CATEGORIE",
      "SV": "KATEGORI"
    },
    "COL_B_ALLO": {
      "FR": "ALLOUÉ",
      "EN": "ALLOCATED",
      "ES": "ASIGNADO",
      "PT": "ALOCADO",
      "DE": "ZUGEWIESEN",
      "IT": "ASSEGNATO",
      "NL": "TOEGEWEZEN",
      "SV": "TILLDELAD"
    },
    "COL_B_REAL": {
      "FR": "RÉALISÉ",
      "EN": "ACTUAL",
      "ES": "REAL",
      "PT": "REALIZADO",
      "DE": "IST-WERT",
      "IT": "REALE",
      "NL": "WERKELIJK",
      "SV": "FAKTISKT"
    },
    "COL_B_DIFF": {
      "FR": "ÉCART",
      "EN": "VARIANCE",
      "ES": "DIFERENCIA",
      "PT": "DIFERENÇA",
      "DE": "ABWEICHUNG",
      "IT": "VARIAZIONE",
      "NL": "VERSCHIL",
      "SV": "AVVIKELSE"
    },
    "COL_B_PROG": {
      "FR": "CONSOMMATION",
      "EN": "CONSUMPTION",
      "ES": "CONSUMO",
      "PT": "CONSUMO",
      "DE": "VERBRAUCH",
      "IT": "CONSUMO",
      "NL": "VERBRUIK",
      "SV": "FÖRBRUKNING"
    },
    "NO_BUDG": {
      "FR": "Aucun budget alloué",
      "EN": "No budget allocated",
      "ES": "Sin presupuesto",
      "PT": "Sem orçamento",
      "DE": "Kein Budget",
      "IT": "Nessun budget",
      "NL": "Geen budget",
      "SV": "Ingen budget"
    },
    "FRM_B_MOIS": {
      "FR": "Mois cible (AAAA-MM) :",
      "EN": "Target Month (YYYY-MM) :",
      "ES": "Mes objetivo (AAAA-MM) :",
      "PT": "Mês alvo (AAAA-MM) :",
      "DE": "Zielmonat (JJJJ-MM) :",
      "IT": "Mese target (AAAA-MM) :",
      "NL": "Doelmaand (JJJJ-MM) :",
      "SV": "Målmånad (ÅÅÅÅ-MM) :"
    },
    "FRM_B_CAT": {
      "FR": "Enveloppe (Catégorie) :",
      "EN": "Envelope (Category) :",
      "ES": "Sobre (Categoría) :",
      "PT": "Envelope (Categoria) :",
      "DE": "Umschlag (Kategorie) :",
      "IT": "Busta (Categoria) :",
      "NL": "Envelop (Categorie) :",
      "SV": "Kuvert (Kategori) :"
    },
    "FRM_B_AMT": {
      "FR": "Montant Alloué :",
      "EN": "Allocated Amount :",
      "ES": "Monto Asignado :",
      "PT": "Valor Alocado :",
      "DE": "Zugewiesener Betrag :",
      "IT": "Importo Assegnato :",
      "NL": "Toegewezen Bedrag :",
      "SV": "Tilldelat Belopp :"
    },
    "FRM_B_SAVE": {
      "FR": "ALLOUER",
      "EN": "ALLOCATE",
      "ES": "ASIGNAR",
      "PT": "ALOCAR",
      "DE": "ZUWEISEN",
      "IT": "ASSEGNA",
      "NL": "TOEWIJZEN",
      "SV": "TILLDELA"
    },
    "FRM_B_CANCEL": {
      "FR": "ANNULER",
      "EN": "CANCEL",
      "ES": "CANCELAR",
      "PT": "CANCELAR",
      "DE": "ABBRECHEN",
      "IT": "ANNULLA",
      "NL": "ANNULEREN",
      "SV": "AVBRYT"
    },
    "FRM_B_DEV": {
      "FR": "Devise :",
      "EN": "Currency :",
      "ES": "Divisa :",
      "PT": "Moeda :",
      "DE": "Währung :",
      "IT": "Valuta :",
      "NL": "Valuta :",
      "SV": "Valuta :"
    },
    "KPI_INC": {
      "FR": "TOTAL REVENUS",
      "EN": "TOTAL INCOME",
      "ES": "TOTAL INGRESOS",
      "PT": "RENDA TOTAL",
      "DE": "GESAMTEINKOMMEN",
      "IT": "TOTALE ENTRATE",
      "NL": "TOTALE INKOMSTEN",
      "SV": "TOTAL INKOMST"
    },
    "KPI_EXP": {
      "FR": "TOTAL DÉPENSES",
      "EN": "TOTAL EXPENSES",
      "ES": "TOTAL GASTOS",
      "PT": "DESPESAS TOTAIS",
      "DE": "GESAMTAUSGABEN",
      "IT": "TOTALE USCITE",
      "NL": "TOTALE UITGAVEN",
      "SV": "TOTALA UTGIFTER"
    },
    "KPI_NET": {
      "FR": "CASHFLOW NET",
      "EN": "NET CASHFLOW",
      "ES": "FLUJO NETO",
      "PT": "FLUXO LÍQUIDO",
      "DE": "NETTO-CASHFLOW",
      "IT": "CASHFLOW NETTO",
      "NL": "NETTO CASHFLOW",
      "SV": "NETTO KASSAFLÖDE"
    },
    "BTN_BACK": {
      "FR": "RETOUR AU MENU",
      "EN": "BACK TO MENU",
      "ES": "VOLVER AL MENÚ",
      "PT": "VOLTAR AO MENU",
      "DE": "ZURÜCK ZUM MENÜ",
      "IT": "TORNA AL MENU",
      "NL": "TERUG NAAR MENU",
      "SV": "TILLBAKA TILL MENY"
    },
    "COL_TYPE": {
      "FR": "TYPE DE FLUX",
      "EN": "FLOW TYPE",
      "ES": "TIPO DE FLUJO",
      "PT": "TIPO DE FLUXO",
      "DE": "FLUSSTYP",
      "IT": "TIPO DI FLUSSO",
      "NL": "STROOMTYPE",
      "SV": "FLÖDESTYP"
    },
    "NO_DATA": {
      "FR": "Aucun compte actif à cette date",
      "EN": "No active accounts on this date",
      "ES": "Sin cuentas en esta fecha",
      "PT": "Sem contas nesta data",
      "DE": "Keine Konten",
      "IT": "Nessun conto",
      "NL": "Geen rekeningen",
      "SV": "Inga konton"
    },
    "COL_YEAR": {
      "FR": "Année",
      "EN": "Year",
      "ES": "Año",
      "PT": "Ano",
      "DE": "Jahr",
      "IT": "Anno",
      "NL": "Jaar",
      "SV": "År"
    },
    "COL_MONTH": {
      "FR": "Mois",
      "EN": "Month",
      "ES": "Mes",
      "PT": "Mês",
      "DE": "Monat",
      "IT": "Mese",
      "NL": "Maand",
      "SV": "Månad"
    },
    "KPI_ASSETS": {
      "FR": "TOTAL ACTIFS",
      "EN": "TOTAL ASSETS",
      "ES": "TOTAL ACTIVOS",
      "PT": "TOTAL DE ATIVOS",
      "DE": "GESAMTVERMÖGEN",
      "IT": "TOTALE ATTIVITÀ",
      "NL": "TOTALE ACTIVA",
      "SV": "TOTALA TILLGÅNGAR"
    },
    "KPI_LIAB": {
      "FR": "TOTAL PASSIFS",
      "EN": "TOTAL LIABILITIES",
      "ES": "TOTAL PASIVOS",
      "PT": "TOTAL DE PASSIVOS",
      "DE": "GESAMTVERBINDLICHKEITEN",
      "IT": "TOTALE PASSIVITÀ",
      "NL": "TOTALE PASSIVA",
      "SV": "TOTALA SKULDER"
    },
    "KPI_NW": {
      "FR": "VALEUR NETTE",
      "EN": "NET WORTH",
      "ES": "PATRIMONIO NETO",
      "PT": "PATRIMÔNIO LÍQUIDO",
      "DE": "NETTOVERMÖGEN",
      "IT": "PATRIMONIO NETTO",
      "NL": "NETTOWAARDE",
      "SV": "NETTOFÖRMÖGENHET"
    },
    "NW_COMPTE": {
      "FR": "COMPTE FINANCIER",
      "EN": "FINANCIAL ACCOUNT",
      "ES": "CUENTA FINANCIERA",
      "PT": "CONTA FINANCEIRA",
      "DE": "FINANZKONTO",
      "IT": "CONTO FINANZIARIO",
      "NL": "FINANCIËLE REKENING",
      "SV": "FINANSIELLT KONTO"
    },
    "NW_CLASS": {
      "FR": "CLASSE D'ACTIF",
      "EN": "ASSET CLASS",
      "ES": "CLASE DE ACTIVO",
      "PT": "CLASSE DE ATIVO",
      "DE": "ANLAGEKLASSE",
      "IT": "CLASSE DI INVESTIMENTO",
      "NL": "ACTIVA KLASSE",
      "SV": "TILLGÅNGSKLASS"
    },
    "NW_BALANCE": {
      "FR": "SOLDE",
      "EN": "BALANCE",
      "ES": "SALDO",
      "PT": "SALDO",
      "DE": "GUTHABEN",
      "IT": "SALDO",
      "NL": "SALDO",
      "SV": "SALDO"
    },
    "MSG_WELCOME_CUR": {
      "FR": "BIENVENUE ! VEUILLEZ CONFIGURER VOTRE SYSTÈME._x000D_\n_x000D_\nSaisissez votre Devise Principale (ex: XOF, EUR, USD)._x000D_\nATTENTION : Ce choix sera définitif pour garantir l'intégrité comptable.",
      "EN": "WELCOME! PLEASE CONFIGURE YOUR SYSTEM._x000D_\n_x000D_\nEnter your Base Currency (e.g., USD, EUR)._x000D_\nWARNING: This choice is permanent to ensure accounting integrity.",
      "ES": "¡BIENVENIDO! CONFIGURE SU SISTEMA.",
      "PT": "BEM-VINDO! CONFIGURE SEU SISTEMA.",
      "DE": "WILLKOMMEN! BITTE KONFIGURIEREN SIE IHR SYSTEM.",
      "IT": "BENVENUTO! CONFIGURA IL TUO SISTEMA.",
      "NL": "WELKOM! CONFIGUREER UW SYSTEEM.",
      "SV": "VÄLKOMMEN! KONFIGURERA DITT SYSTEM."
    },
    "MSG_LOCKED_CUR": {
      "FR": "SÉCURITÉ COMPTABLE ACTIVÉE_x000D_\n_x000D_\nVotre devise principale est verrouillée sur :",
      "EN": "ACCOUNTING SECURITY ENABLED_x000D_\n_x000D_\nYour base currency is locked to :",
      "ES": "SEGURIDAD CONTABLE ACTIVADA",
      "PT": "SEGURANÇA CONTÁBIL ATIVADA",
      "DE": "BUCHHALTUNGSSICHERHEIT AKTIVIERT",
      "IT": "SICUREZZA CONTABILE ATTIVATA",
      "NL": "ACCOUNTING BEVEILIGING INGESCHAKELD",
      "SV": "BOKFÖRINGSSÄKERHET AKTIVERAD"
    },
    "MSG_LOCKED_DESC": {
      "FR": "Il est impossible de modifier la devise mère après l'initialisation afin de préserver votre bilan. (Un Factory Reset est requis pour changer).",
      "EN": "It is impossible to change the base currency after initialization to protect your ledger. (A Factory Reset is required).",
      "ES": "Imposible cambiar la divisa base.",
      "PT": "Impossível mudar a moeda base.",
      "DE": "Basiswährung kann nicht geändert werden.",
      "IT": "Impossibile cambiare la valuta base.",
      "NL": "Basisvaluta kan niet worden gewijzigd.",
      "SV": "Basvalutan kan inte ändras."
    },
    "MSG_ONB_EMPTY": {
      "FR": "La configuration initiale est obligatoire pour démarrer.",
      "EN": "Initial configuration is mandatory to start.",
      "ES": "La configuración inicial es obligatoria.",
      "PT": "A configuração inicial é obrigatória.",
      "DE": "Die Erstkonfiguration ist obligatorisch.",
      "IT": "La configurazione iniziale è obbligatoria.",
      "NL": "Initiële configuratie is verplicht.",
      "SV": "Inledande konfiguration är obligatorisk."
    },
    "MSG_ONB_GUIDE": {
      "FR": "Veuillez saisir un code ISO de 3 lettres (ex: USD, EUR, MUR)._x000D_\nINFO SYSTÈME : Votre symbole monétaire local est[",
      "EN": "Please enter a 3-letter ISO code (e.g., USD, EUR, GBP)._x000D_\nSYSTEM INFO: Your local currency symbol is[",
      "ES": "Ingrese un código ISO de 3 letras._x000D_\nINFO SISTEMA: Su símbolo local es[",
      "PT": "Insira um código ISO de 3 letras._x000D_\nINFO SISTEMA: Seu símbolo local é[",
      "DE": "Geben Sie einen 3-stelligen ISO-Code ein._x000D_\nSYSTEM-INFO: Ihr lokales Symbol ist[",
      "IT": "Inserisci un codice ISO di 3 lettere._x000D_\nINFO SISTEMA: Il tuo simbolo locale è[",
      "NL": "Voer een 3-letterige ISO-code in._x000D_\nSYSTEEMINFO: Uw lokale symbool is[",
      "SV": "Ange en 3-bokstavs ISO-kod._x000D_\nSYSTEMINFO: Din lokala symbol är["
    },
    "MSG_ONB_REJECT": {
      "FR": "Devise non reconnue par le marché boursier mondial.",
      "EN": "Currency not recognized by the global stock market.",
      "ES": "Divisa no reconocida por el mercado mundial.",
      "PT": "Moeda não reconhecida pelo mercado mundial.",
      "DE": "Währung vom Weltmarkt nicht erkannt.",
      "IT": "Valuta non riconosciuta dal mercato mondiale.",
      "NL": "Valuta niet herkend door de wereldmarkt.",
      "SV": "Valutan känns inte igen av världsmarknaden."
    },
    "MSG_ONB_OFFLINE": {
      "FR": "Mode hors-ligne détecté._x000D_\nVotre devise a été enregistrée. Vous pourrez synchroniser les taux plus tard via le bouton d'actualisation.",
      "EN": "Offline mode detected._x000D_\nYour currency is saved. You can sync rates later using the update button.",
      "ES": "Modo sin conexión. Divisa guardada.",
      "PT": "Modo offline. Moeda salva.",
      "DE": "Offline-Modus. Währung gespeichert.",
      "IT": "Modalità offline. Valuta salvata.",
      "NL": "Offline modus. Valuta opgeslagen.",
      "SV": "Offline-läge. Valuta sparad."
    }
  },
  "config": {
    "SYS_DEVISE_BASE": "MUR",
    "LANGUE_UI": "FR",
    "DASH_FILTRE_ANNEE": "2026",
    "DASH_FILTRE_MOIS_SEUL": "02",
    "DASH_FILTRE_DEV": "MUR",
    "NW_FILTRE_MOIS": "2026-02",
    "NW_FILTRE_DEV": "MUR",
    "BUDG_FILTRE_MOIS": "2026-02",
    "BUDG_FILTRE_DEV": "MUR",
    "SEQ_T_DIM_Compte": "9",
    "SEQ_T_DIM_Categorie": "11",
    "SEQ_T_DIM_Tiers": "7"
  }
};
