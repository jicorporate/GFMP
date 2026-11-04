/**
 * GFMP — Global Finance Management Platform (SFP v1.0)
 * Ultra-Modern Local-First Wealth & Finance Analytics Engine
 * 100% Offline, Zero-Network, Real-Time Calculations (< 5ms)
 * 
 * Strict SFP v1.0 Color Charter & Golden Standard
 */

const GFMP_DB_TRANSLATIONS = {
  // --- COMPTES ---
  'Compte Courant Principal': {
    FR: 'Compte Courant Principal',
    EN: 'Main Checking Account',
    ES: 'Cuenta Corriente Principal',
    PT: 'Conta Corrente Principal'
  },
  'Livret d\'Épargne': {
    FR: 'Livret d\'Épargne',
    EN: 'Savings Account',
    ES: 'Cuenta de Ahorros',
    PT: 'Conta Poupança'
  },
  'Portefeuille Espèces': {
    FR: 'Portefeuille Espèces',
    EN: 'Cash Wallet',
    ES: 'Billetera de Efectivo',
    PT: 'Carteira de Dinheiro'
  },
  'Assurance Vie': {
    FR: 'Assurance Vie',
    EN: 'Life Insurance / Wealth Fund',
    ES: 'Seguro de Vida / Inversión',
    PT: 'Seguro de Vida / Fundo'
  },
  'PEA / Actions': {
    FR: 'PEA / Actions',
    EN: 'Stock Portfolio / Equities',
    ES: 'Cartera de Acciones / Bolsa',
    PT: 'Carteira de Ações / Bolsa'
  },
  'Portefeuille Crypto': {
    FR: 'Portefeuille Crypto',
    EN: 'Crypto Wallet',
    ES: 'Cartera Cripto',
    PT: 'Carteira Cripto'
  },
  'Prêt Immobilier': {
    FR: 'Prêt Immobilier',
    EN: 'Mortgage / Property Loan',
    ES: 'Préstamo Hipotecario',
    PT: 'Empréstimo Habitação'
  },
  'Carte de Crédit (Différé)': {
    FR: 'Carte de Crédit (Différé)',
    EN: 'Credit Card (Deferred)',
    ES: 'Tarjeta de Crédito (Diferido)',
    PT: 'Cartão de Crédito (Diferido)'
  },
  'Carte de Crédit': {
    FR: 'Carte de Crédit',
    EN: 'Credit Card',
    ES: 'Tarjeta de Crédito',
    PT: 'Cartão de Crédito'
  },
  'Autre (Préciser...)': {
    FR: 'Autre (Préciser...)',
    EN: 'Other (Specify...)',
    ES: 'Otro (Especificar...)',
    PT: 'Outro (Especificar...)'
  },

  // --- CATEGORIES ---
  'Salaire / Revenus Pro': {
    FR: 'Salaire / Revenus Pro',
    EN: 'Salary / Professional Income',
    ES: 'Salario / Ingresos Profesionales',
    PT: 'Salário / Rendimentos Profissionais'
  },
  'Intérêts / Dividendes': {
    FR: 'Intérêts / Dividendes',
    EN: 'Interests / Dividends',
    ES: 'Intereses / Dividendos',
    PT: 'Juros / Dividendos'
  },
  'Aides / Bourses / Allocations': {
    FR: 'Aides / Bourses / Allocations',
    EN: 'Grants / Scholarships / Allowances',
    ES: 'Ayudas / Becas / Subsidios',
    PT: 'Ajudas / Bolsas / Subsídios'
  },
  'Logement (Loyer/Prêt/Charges)': {
    FR: 'Logement (Loyer/Prêt/Charges)',
    EN: 'Housing (Rent/Loan/Bills)',
    ES: 'Vivienda (Alquiler/Préstamo/Gastos)',
    PT: 'Habitação (Renda/Empréstimo/Despesas)'
  },
  'Alimentation & Supermarché': {
    FR: 'Alimentation & Supermarché',
    EN: 'Food & Groceries',
    ES: 'Alimentación & Supermercado',
    PT: 'Alimentação & Supermercado'
  },
  'Transports (Essence/Assurance)': {
    FR: 'Transports (Essence/Assurance)',
    EN: 'Transport (Fuel/Insurance)',
    ES: 'Transporte (Gasolina/Seguro)',
    PT: 'Transporte (Combustível/Seguro)'
  },
  'Santé & Mutuelle': {
    FR: 'Santé & Mutuelle',
    EN: 'Healthcare & Insurance',
    ES: 'Salud & Seguro Médico',
    PT: 'Saúde & Seguro'
  },
  'Loisirs, Sorties & Vacances': {
    FR: 'Loisirs, Sorties & Vacances',
    EN: 'Leisure, Outings & Vacations',
    ES: 'Ocio, Salidas & Vacaciones',
    PT: 'Lazer, Saídas & Férias'
  },
  'Impôts & Taxes': {
    FR: 'Impôts & Taxes',
    EN: 'Taxes & Duties',
    ES: 'Impuestos & Tasas',
    PT: 'Impostos & Taxas'
  },
  'Virement Interne (Épargne)': {
    FR: 'Virement Interne (Épargne)',
    EN: 'Internal Transfer (Savings)',
    ES: 'Transferencia Interna (Ahorro)',
    PT: 'Transferência Interna (Poupança)'
  },

  // --- TIERS ---
  'Employeur principal': {
    FR: 'Employeur principal',
    EN: 'Main Employer',
    ES: 'Empleador principal',
    PT: 'Empregador principal'
  },
  'Banque / Courtier': {
    FR: 'Banque / Courtier',
    EN: 'Bank / Broker',
    ES: 'Banco / Bróker',
    PT: 'Banco / Corretora'
  },
  'État / Impôts': {
    FR: 'État / Impôts',
    EN: 'Government / Tax Authority',
    ES: 'Estado / Hacienda',
    PT: 'Estado / Finanças'
  },
  'Supermarché / Alimentaire': {
    FR: 'Supermarché / Alimentaire',
    EN: 'Supermarket / Grocery',
    ES: 'Supermercado / Alimentación',
    PT: 'Supermercado / Alimentar'
  },
  'Propriétaire / Syndic': {
    FR: 'Propriétaire / Syndic',
    EN: 'Landlord / Property Manager',
    ES: 'Propietario / Administración',
    PT: 'Proprietário / Condomínio'
  },
  'Clinique & Pharmacie': {
    FR: 'Clinique & Pharmacie',
    EN: 'Clinic & Pharmacy',
    ES: 'Clínica & Farmacia',
    PT: 'Clínica & Farmácia'
  },

  // --- CLASSES & TYPES ---
  'LIQUIDITE': {
    FR: 'Liquidités',
    EN: 'Liquidity',
    ES: 'Liquidez',
    PT: 'Liquidez'
  },
  'INVESTISSEMENT': {
    FR: 'Investissements',
    EN: 'Investments',
    ES: 'Inversiones',
    PT: 'Investimentos'
  },
  'DETTE': {
    FR: 'Dettes & Passifs',
    EN: 'Debts & Liabilities',
    ES: 'Deudas & Pasivos',
    PT: 'Dívidas & Passivos'
  },
  'REVENU': {
    FR: 'Revenu',
    EN: 'Income',
    ES: 'Ingreso',
    PT: 'Receita'
  },
  'DEPENSE': {
    FR: 'Dépense',
    EN: 'Expense',
    ES: 'Gasto',
    PT: 'Despesa'
  },
  'TRANSFERT': {
    FR: 'Virement',
    EN: 'Transfer',
    ES: 'Transferencia',
    PT: 'Transferência'
  },
  'COMMERCE': {
    FR: 'Commerce / Distribution',
    EN: 'Retail & Commerce',
    ES: 'Comercio',
    PT: 'Comércio'
  },
  'ENTREPRISE': {
    FR: 'Entreprise',
    EN: 'Corporate',
    ES: 'Empresa',
    PT: 'Empresa'
  },
  'FINANCE': {
    FR: 'Banque & Finance',
    EN: 'Banking & Finance',
    ES: 'Banca & Finanzas',
    PT: 'Banca & Finanças'
  },
  'INSTITUTION': {
    FR: 'Institution Publique',
    EN: 'Public Institution',
    ES: 'Institución Pública',
    PT: 'Instituição Pública'
  },
  'IMMOBILIER': {
    FR: 'Immobilier',
    EN: 'Real Estate',
    ES: 'Inmobiliario',
    PT: 'Imobiliário'
  },
  'SANTE': {
    FR: 'Santé & Médical',
    EN: 'Health & Medical',
    ES: 'Salud & Médico',
    PT: 'Saúde & Médico'
  },
  'AUTRE': {
    FR: 'Autre',
    EN: 'Other',
    ES: 'Otro',
    PT: 'Outro'
  },

  // --- STANDARD TRANSACTION NOTES (BASELINE) ---
  'Premier loyer résidence': {
    FR: 'Premier loyer résidence',
    EN: 'First residence rent',
    ES: 'Primer alquiler residencia',
    PT: 'Primeira renda residência'
  },
  'Ameublement et électroménager installation': {
    FR: 'Ameublement et électroménager installation',
    EN: 'Furnishing & household appliances',
    ES: 'Mobiliario y electrodomésticos',
    PT: 'Mobiliário e eletrodomésticos'
  },
  'Courses d\'installation et provisions': {
    FR: 'Courses d\'installation et provisions',
    EN: 'Setup groceries & supplies',
    ES: 'Compras de instalación y víveres',
    PT: 'Compras de instalação e provisões'
  },
  'Carburant et révision véhicule': {
    FR: 'Carburant et révision véhicule',
    EN: 'Fuel and vehicle servicing',
    ES: 'Combustible y revisión de vehículo',
    PT: 'Combustível e revisão de veículo'
  },
  'Bilan médical et vaccins': {
    FR: 'Bilan médical et vaccins',
    EN: 'Medical checkup and vaccines',
    ES: 'Chequeo médico y vacunas',
    PT: 'Exame médico e vacinas'
  },
  'Salaire Septembre KIMYA ÉDITION': {
    FR: 'Salaire Septembre KIMYA ÉDITION',
    EN: 'September Salary KIMYA ÉDITION',
    ES: 'Salario Septiembre KIMYA ÉDITION',
    PT: 'Salário Setembro KIMYA ÉDITION'
  },
  'Constitution épargne de précaution': {
    FR: 'Constitution épargne de précaution',
    EN: 'Emergency fund creation',
    ES: 'Constitución fondo de emergencia',
    PT: 'Constituição reserva de emergência'
  },
  'Restauration et réception emménagement': {
    FR: 'Restauration et réception emménagement',
    EN: 'Housewarming dining & party',
    ES: 'Restauración y recepción mudanza',
    PT: 'Refeições e festa de inauguração'
  },
  'Loyer Octobre': {
    FR: 'Loyer Octobre',
    EN: 'October Rent',
    ES: 'Alquiler Octubre',
    PT: 'Renda Outubro'
  },
  'Supermarché courses hebdomadaires': {
    FR: 'Supermarché courses hebdomadaires',
    EN: 'Weekly supermarket groceries',
    ES: 'Supermercado compras semanales',
    PT: 'Supermercado compras semanais'
  },
  'Dividendes trimestriels ETF World': {
    FR: 'Dividendes trimestriels ETF World',
    EN: 'Quarterly dividends ETF World',
    ES: 'Dividendos trimestrales ETF World',
    PT: 'Dividendos trimestrais ETF World'
  },
  'Plein carburant station service': {
    FR: 'Plein carburant station service',
    EN: 'Gas station fuel fill-up',
    ES: 'Llenado de combustible estación',
    PT: 'Abastecimento posto de combustível'
  },
  'Sorties culturelles et cinémas': {
    FR: 'Sorties culturelles et cinémas',
    EN: 'Cultural outings and movies',
    ES: 'Salidas culturales y cines',
    PT: 'Passeios culturais e cinemas'
  },
  'Salaire Octobre KIMYA ÉDITION': {
    FR: 'Salaire Octobre KIMYA ÉDITION',
    EN: 'October Salary KIMYA ÉDITION',
    ES: 'Salario Octubre KIMYA ÉDITION',
    PT: 'Salário Outubro KIMYA ÉDITION'
  },
  'Virement épargne mensuelle': {
    FR: 'Virement épargne mensuelle',
    EN: 'Monthly savings transfer',
    ES: 'Transferencia ahorro mensual',
    PT: 'Transferência poupança mensal'
  },
  'Pharmacie et soins dentaires': {
    FR: 'Pharmacie et soins dentaires',
    EN: 'Pharmacy and dental care',
    ES: 'Farmacia y cuidado dental',
    PT: 'Farmácia e cuidados dentários'
  },
  'Loyer Novembre': {
    FR: 'Loyer Novembre',
    EN: 'November Rent',
    ES: 'Alquiler Noviembre',
    PT: 'Renda Novembro'
  },
  'Supermarché et produits frais': {
    FR: 'Supermarché et produits frais',
    EN: 'Supermarket and fresh produce',
    ES: 'Supermercado y productos frescos',
    PT: 'Supermercado e produtos frescos'
  },
  'Rendement fonds euros assurance vie': {
    FR: 'Rendement fonds euros assurance vie',
    EN: 'Euro fund yield life insurance',
    ES: 'Rendimiento fondo seguro de vida',
    PT: 'Rendimento fundo seguro de vida'
  },
  'Assurance auto et carburant': {
    FR: 'Assurance auto et carburant',
    EN: 'Car insurance and fuel',
    ES: 'Seguro de auto y combustible',
    PT: 'Seguro de carro e combustível'
  },
  'DCA mensuel Bitcoin & Ethereum': {
    FR: 'DCA mensuel Bitcoin & Ethereum',
    EN: 'Monthly DCA Bitcoin & Ethereum',
    ES: 'DCA mensual Bitcoin & Ethereum',
    PT: 'DCA mensal Bitcoin & Ethereum'
  },
  'Salaire Novembre KIMYA ÉDITION': {
    FR: 'Salaire Novembre KIMYA ÉDITION',
    EN: 'November Salary KIMYA ÉDITION',
    ES: 'Salario Noviembre KIMYA ÉDITION',
    PT: 'Salário Novembro KIMYA ÉDITION'
  },
  'Dîners restaurants et théâtre': {
    FR: 'Dîners restaurants et théâtre',
    EN: 'Restaurant dinners and theatre',
    ES: 'Cenas en restaurantes y teatro',
    PT: 'Jantares em restaurantes e teatro'
  },
  'Soins ophtalmologiques': {
    FR: 'Soins ophtalmologiques',
    EN: 'Ophthalmic / eye care',
    ES: 'Cuidado oftalmológico',
    PT: 'Cuidados oftalmológicos'
  },
  'Loyer Décembre': {
    FR: 'Loyer Décembre',
    EN: 'December Rent',
    ES: 'Alquiler Diciembre',
    PT: 'Renda Dezembro'
  },
  'Courses festives et repas de fête': {
    FR: 'Courses festives et repas de fête',
    EN: 'Holiday groceries & festive feast',
    ES: 'Compras festivas y cena de fiesta',
    PT: 'Compras festivas e refeição comemorativa'
  },
  'Cadeaux de Noël et célébrations': {
    FR: 'Cadeaux de Noël et célébrations',
    EN: 'Christmas gifts and celebrations',
    ES: 'Regalos de Navidad y fiestas',
    PT: 'Presentes de Natal e celebrações'
  },
  'Trajets fêtes et carburant': {
    FR: 'Trajets fêtes et carburant',
    EN: 'Holiday travel and fuel',
    ES: 'Viajes de fiestas y combustible',
    PT: 'Viagens de férias e combustível'
  },
  'Salaire Décembre et Prime Annuelle': {
    FR: 'Salaire Décembre et Prime Annuelle',
    EN: 'December Salary & Annual Bonus',
    ES: 'Salario Diciembre y Bono Anual',
    PT: 'Salário Dezembro e Bônus Anual'
  },
  'Dividendes semestriels actions': {
    FR: 'Dividendes semestriels actions',
    EN: 'Semi-annual stock dividends',
    ES: 'Dividendos semestrales de acciones',
    PT: 'Dividendos semestrais de ações'
  },
  'Placement prime sur livret épargne': {
    FR: 'Placement prime sur livret épargne',
    EN: 'Bonus deposit into savings account',
    ES: 'Depósito de bono en cuenta de ahorro',
    PT: 'Depósito de bônus na poupança'
  },
  'Pharmacie d\'hiver': {
    FR: 'Pharmacie d\'hiver',
    EN: 'Winter pharmacy remedies',
    ES: 'Farmacia de invierno',
    PT: 'Farmácia de inverno'
  },
  'Réveillon Saint-Sylvestre': {
    FR: 'Réveillon Saint-Sylvestre',
    EN: 'New Year\'s Eve celebration',
    ES: 'Celebración de Nochevieja',
    PT: 'Celebração de Réveillon'
  },
  'Loyer Janvier 2026': {
    FR: 'Loyer Janvier 2026',
    EN: 'January 2026 Rent',
    ES: 'Alquiler Enero 2026',
    PT: 'Renda Janeiro 2026'
  },
  'Supermarché rentrée': {
    FR: 'Supermarché rentrée',
    EN: 'Back-to-work groceries',
    ES: 'Supermercado vuelta al trabajo',
    PT: 'Supermercado volta ao trabalho'
  },
  'Acompte fiscal trimestriel MRA': {
    FR: 'Acompte fiscal trimestriel MRA',
    EN: 'Quarterly tax prepayment MRA',
    ES: 'Pago fraccionado impuestos MRA',
    PT: 'Adiantamento fiscal trimestral MRA'
  },
  'Entretien annuel et essence': {
    FR: 'Entretien annuel et essence',
    EN: 'Annual car maintenance and petrol',
    ES: 'Mantenimiento anual y gasolina',
    PT: 'Manutenção anual e combustível'
  },
  'Cotisation mutuelle santé': {
    FR: 'Cotisation mutuelle santé',
    EN: 'Health insurance monthly premium',
    ES: 'Cuota de seguro de salud',
    PT: 'Mensalidade do seguro de saúde'
  },
  'Sorties week-end et livres': {
    FR: 'Sorties week-end et livres',
    EN: 'Weekend outings and books',
    ES: 'Salidas de fin de semana y libros',
    PT: 'Passeios de fim de semana e livros'
  },
  'Salaire Janvier KIMYA ÉDITION': {
    FR: 'Salaire Janvier KIMYA ÉDITION',
    EN: 'January Salary KIMYA ÉDITION',
    ES: 'Salario Enero KIMYA ÉDITION',
    PT: 'Salário Janeiro KIMYA ÉDITION'
  },
  'Allocation et crédit d\'impôt': {
    FR: 'Allocation et crédit d\'impôt',
    EN: 'Family allowance & tax credit',
    ES: 'Asignación y crédito fiscal',
    PT: 'Subsídio e crédito fiscal'
  },
  'Épargne programmée vers livret': {
    FR: 'Épargne programmée vers livret',
    EN: 'Scheduled transfer to savings',
    ES: 'Ahorro programado a cuenta',
    PT: 'Poupança programada para conta'
  },
  'Loyer Février Résidence KIMYA': {
    FR: 'Loyer Février Résidence KIMYA',
    EN: 'February Rent KIMYA Residence',
    ES: 'Alquiler Febrero Residencia KIMYA',
    PT: 'Renda Fevereiro Residência KIMYA'
  },
  'Super U Grand Baie - Alimentation': {
    FR: 'Super U Grand Baie - Alimentation',
    EN: 'Super U Grand Baie - Groceries',
    ES: 'Super U Grand Baie - Alimentación',
    PT: 'Super U Grand Baie - Alimentação'
  },
  'Plein carburant TotalEnergies': {
    FR: 'Plein carburant TotalEnergies',
    EN: 'Fuel fill-up TotalEnergies',
    ES: 'Combustible TotalEnergies',
    PT: 'Abastecimento TotalEnergies'
  },
  'Dividendes Portefeuille PEA (441.28 MUR)': {
    FR: 'Dividendes Portefeuille PEA (441.28 MUR)',
    EN: 'Stock Portfolio Dividends (441.28 MUR)',
    ES: 'Dividendos Cartera Acciones (441.28 MUR)',
    PT: 'Dividendos Carteira de Ações (441.28 MUR)'
  },
  'Dîner Saint-Valentin restaurant': {
    FR: 'Dîner Saint-Valentin restaurant',
    EN: 'Valentine\'s Day dinner restaurant',
    ES: 'Cena de San Valentín restaurante',
    PT: 'Jantar de São Valentim restaurante'
  },
  'Marché central et épicerie fine': {
    FR: 'Marché central et épicerie fine',
    EN: 'Central market & gourmet grocery',
    ES: 'Mercado central y delicatessen',
    PT: 'Mercado central e charcutaria'
  },
  'Consultation clinique et pharmacie': {
    FR: 'Consultation clinique et pharmacie',
    EN: 'Clinic consultation and pharmacy',
    ES: 'Consulta clínica y farmacia',
    PT: 'Consulta clínica e farmácia'
  },
  'Achat matériel tech Apple Store (20,271 MUR - Cas Dépassement Budget)': {
    FR: 'Achat matériel tech Apple Store (20,271 MUR - Cas Dépassement Budget)',
    EN: 'Apple Store Tech Hardware (20,271 MUR - Budget Overrun Case)',
    ES: 'Compra tecnología Apple Store (20,271 MUR - Exceso de Presupuesto)',
    PT: 'Compra tecnologia Apple Store (20,271 MUR - Caso Excesso Orçamento)'
  },
  'Salaire Février KIMYA ÉDITION': {
    FR: 'Salaire Février KIMYA ÉDITION',
    EN: 'February Salary KIMYA ÉDITION',
    ES: 'Salario Febrero KIMYA ÉDITION',
    PT: 'Salário Fevereiro KIMYA ÉDITION'
  },
  'Honoraires mission consulting financier': {
    FR: 'Honoraires mission consulting financier',
    EN: 'Financial consulting mission fees',
    ES: 'Honorarios asesoría financiera',
    PT: 'Honorários consultoria financeira'
  },
  'Taxe foncière et droits communaux': {
    FR: 'Taxe foncière et droits communaux',
    EN: 'Property tax and municipal fees',
    ES: 'Impuesto de bienes inmuebles y tasas',
    PT: 'Imposto municipal e taxas locais'
  },
  'Loyer Mars 2026': {
    FR: 'Loyer Mars 2026',
    EN: 'March 2026 Rent',
    ES: 'Alquiler Marzo 2026',
    PT: 'Renda Março 2026'
  },
  'Courses début de mois': {
    FR: 'Courses début de mois',
    EN: 'Start-of-month groceries',
    ES: 'Compras de inicio de mes',
    PT: 'Compras de início de mês'
  },
  'Recharge carburant': {
    FR: 'Recharge carburant',
    EN: 'Fuel refill',
    ES: 'Recarga de combustible',
    PT: 'Abastecimento de combustível'
  },
  'Intérêts trimestriels fonds euros': {
    FR: 'Intérêts trimestriels fonds euros',
    EN: 'Quarterly Euro fund interest',
    ES: 'Intereses trimestrales fondo euros',
    PT: 'Juros trimestrais fundo euros'
  },
  'Sortie cinéma et détente': {
    FR: 'Sortie cinéma et détente',
    EN: 'Cinema outing & relaxation',
    ES: 'Salida al cine y descanso',
    PT: 'Ida ao cinema e descanso'
  },
  'Virement épargne préventive': {
    FR: 'Virement épargne préventive',
    EN: 'Preventive savings transfer',
    ES: 'Transferencia ahorro preventivo',
    PT: 'Transferência poupança preventiva'
  },
  'Salaire Mars KIMYA ÉDITION': {
    FR: 'Salaire Mars KIMYA ÉDITION',
    EN: 'March Salary KIMYA ÉDITION',
    ES: 'Salario Marzo KIMYA ÉDITION',
    PT: 'Salário Março KIMYA ÉDITION'
  }
};

// Build complete case-insensitive bidirectional lookup table for instant O(1) multilingual translation
const GFMP_DB_LOOKUP = {};
Object.keys(GFMP_DB_TRANSLATIONS).forEach(primaryKey => {
  const translations = GFMP_DB_TRANSLATIONS[primaryKey];
  GFMP_DB_LOOKUP[primaryKey.toLowerCase().trim()] = translations;
  Object.values(translations).forEach(val => {
    if (val && typeof val === 'string') {
      GFMP_DB_LOOKUP[val.toLowerCase().trim()] = translations;
    }
  });
});

// Haute Joaillerie SFP v1.0 Categorical Color Palette
const GFMP_LUXURY_PALETTE = [
  '#3B82F6', // Royal Sapphire Blue (Logement / Housing)
  '#F59E0B', // Imperial Amber Gold (Alimentation / Food)
  '#8B5CF6', // Vibrant Amethyst Violet (Transports / Auto)
  '#10B981', // Pure Emerald Green (Santé / Health)
  '#EC4899', // Radiant Ruby Rose (Loisirs / Entertainment)
  '#F97316', // Deep Topaz Orange (Impôts / Taxes)
  '#06B6D4', // Electric Cyan (Abonnements / Subscriptions)
  '#6366F1', // Royal Indigo (Formation / Education)
  '#14B8A6', // Persian Teal (Habillement / Shopping)
  '#E11D48', // Crimson Garnet (Cadeaux / Gifts)
  '#A855F7', // Rich Lavender (Équipement / Home)
  '#84CC16', // Peridot Lime (Projets)
  '#D946EF', // Fuchsia Tourmaline
  '#0284C7', // Oceanic Blue
  '#94A3B8'  // Platinum Slate (Autre / Misc)
];

// Fixed Semantic Category Color Mapping for all baseline & standard categories
const GFMP_CATEGORY_COLORS = {
  // Standard French Keys
  'logement (loyer/prêt/charges)': '#3B82F6',
  'logement': '#3B82F6',
  'alimentation & supermarché': '#F59E0B',
  'alimentation': '#F59E0B',
  'transports (essence/assurance)': '#8B5CF6',
  'transports': '#8B5CF6',
  'transport': '#8B5CF6',
  'santé & mutuelle': '#10B981',
  'santé': '#10B981',
  'sante': '#10B981',
  'loisirs, sorties & vacances': '#EC4899',
  'loisirs': '#EC4899',
  'impôts & taxes': '#F97316',
  'impôts': '#F97316',
  'impots': '#F97316',
  'abonnements & services': '#06B6D4',
  'abonnements': '#06B6D4',
  'abonnements streaming': '#06B6D4',
  'éducation & formation': '#6366F1',
  'formation': '#6366F1',
  'education': '#6366F1',
  'habillement & mode': '#14B8A6',
  'habillement': '#14B8A6',
  'cadeaux & dons': '#E11D48',
  'cadeaux': '#E11D48',
  'équipement & maison': '#A855F7',
  'autre (préciser...)': '#94A3B8',
  'autre': '#94A3B8',

  // Multilingual Canonical Keys (EN)
  'housing (rent/loan/charges)': '#3B82F6',
  'housing': '#3B82F6',
  'food & supermarket': '#F59E0B',
  'food': '#F59E0B',
  'groceries': '#F59E0B',
  'transport (fuel/insurance)': '#8B5CF6',
  'transportation': '#8B5CF6',
  'health & insurance': '#10B981',
  'health': '#10B981',
  'leisure, outings & holidays': '#EC4899',
  'leisure & entertainment': '#EC4899',
  'leisure': '#EC4899',
  'taxes & duties': '#F97316',
  'taxes': '#F97316',
  'subscriptions': '#06B6D4',
  'other': '#94A3B8',

  // Multilingual Canonical Keys (ES)
  'vivienda (alquiler/hipoteca/gastos)': '#3B82F6',
  'vivienda': '#3B82F6',
  'alimentación y supermercado': '#F59E0B',
  'alimentacion y supermercado': '#F59E0B',
  'alimentación': '#F59E0B',
  'transporte (gasolina/seguro)': '#8B5CF6',
  'transporte': '#8B5CF6',
  'salud y mutua': '#10B981',
  'salud': '#10B981',
  'ocio, salidas y vacaciones': '#EC4899',
  'ocio': '#EC4899',
  'impuestos y tasas': '#F97316',
  'impuestos': '#F97316',
  'suscripciones': '#06B6D4',
  'otro': '#94A3B8',

  // Multilingual Canonical Keys (PT)
  'habitação (renda/empréstimo/despesas)': '#3B82F6',
  'habitação': '#3B82F6',
  'habitacao': '#3B82F6',
  'alimentação & supermercado': '#F59E0B',
  'alimentacao & supermercado': '#F59E0B',
  'alimentação': '#F59E0B',
  'transportes (combustível/seguro)': '#8B5CF6',
  'transportes': '#8B5CF6',
  'saúde & seguro': '#10B981',
  'saúde': '#10B981',
  'saude': '#10B981',
  'lazer, saídas e férias': '#EC4899',
  'lazer': '#EC4899',
  'impostos & taxas': '#F97316',
  'impostos': '#F97316',
  'assinaturas': '#06B6D4',
  'outro': '#94A3B8'
};

class GFMPEngine {
  constructor() {
    this.data = null;
    this.lang = 'FR';
    this.currency = 'MUR';
    this.year = '2026';
    this.month = '02'; // '00' = Année entière
    this.referenceDate = '2026-02-27';
    this.theme = 'blue';
    this.dataMode = 'DEMO'; // 'DEMO' or 'EMPTY'
    this.currentView = 'home';
    this.hubFilter = 'ALL'; // 'ALL', 'CASHFLOW', 'PATRIMOINE', 'BUDGET'
    this.searchTerm = '';
    this.fluxFilter = 'ALL'; // 'ALL', 'REVENU', 'DEPENSE', 'TRANSFERT'
    this.selectedCategoryFilter = null; // Filter for Donut click
    this.accountFilter = 'ALL'; // Filter by account
    this.budgetStatusFilter = 'ALL'; // 'ALL', 'OVER', 'OK'
    this.patrimoineClassFilter = 'ALL'; // 'ALL', 'LIQUIDITE', 'INVESTISSEMENT', 'DETTE'
    this.sortField = 'date';
    this.sortDir = 'desc';
    this.listeners = [];
  }

  roundMoney(val) {
    if (val === null || val === undefined || isNaN(val)) return 0;
    const res = Math.round((Number(val) + Number.EPSILON) * 100) / 100;
    return res === 0 ? 0 : res;
  }

  cleanText(str) {
    if (!str) return '';
    return String(str).replace(/\s+/g, ' ').trim();
  }

  trEntity(name, lang = this.lang) {
    if (!name) return '';
    const clean = this.cleanText(name);
    const targetLang = (lang && ['FR', 'EN', 'ES', 'PT'].includes(lang)) ? lang : 'FR';
    const cleanLower = clean.toLowerCase();

    // 1. Direct primary key lookup
    if (GFMP_DB_TRANSLATIONS[clean] && GFMP_DB_TRANSLATIONS[clean][targetLang]) {
      return GFMP_DB_TRANSLATIONS[clean][targetLang];
    }

    // 2. Dynamic database custom dictionary lookup
    if (this.data && this.data.dictionary && this.data.dictionary[clean]) {
      const dictEntry = this.data.dictionary[clean];
      if (dictEntry[targetLang]) return dictEntry[targetLang];
      if (dictEntry['FR']) return dictEntry['FR'];
    }

    // 3. Bidirectional multilingual reverse lookup
    if (GFMP_DB_LOOKUP[cleanLower] && GFMP_DB_LOOKUP[cleanLower][targetLang]) {
      return GFMP_DB_LOOKUP[cleanLower][targetLang];
    }

    // 4. Stripped lookup without parenthesized suffixes
    const stripped = clean.replace(/\s*\([A-Z]{3,4}\)/g, '')
                          .replace(/\s*\(Préciser\.\.\.\)/i, '')
                          .replace(/\s*\(Différé\)/i, '')
                          .trim();
    const strippedLower = stripped.toLowerCase();
    if (GFMP_DB_TRANSLATIONS[stripped] && GFMP_DB_TRANSLATIONS[stripped][targetLang]) {
      return GFMP_DB_TRANSLATIONS[stripped][targetLang];
    }
    if (GFMP_DB_LOOKUP[strippedLower] && GFMP_DB_LOOKUP[strippedLower][targetLang]) {
      return GFMP_DB_LOOKUP[strippedLower][targetLang];
    }
    if (this.data && this.data.dictionary && this.data.dictionary[stripped] && this.data.dictionary[stripped][targetLang]) {
      return this.data.dictionary[stripped][targetLang];
    }

    return clean;
  }

  cleanAssetName(name) {
    if (!name) return '';
    const cleaned = String(name).replace(/\s*\([A-Z]{3,4}\)/g, '').replace(/\s*\(Préciser\.\.\.\)/i, '').trim();
    return this.trEntity(cleaned);
  }

  getCategoryColor(categoryName) {
    if (!categoryName) return GFMP_LUXURY_PALETTE[0];
    const raw = String(categoryName).trim();
    const clean = raw.toLowerCase();

    // 1. Direct dictionary match
    if (GFMP_CATEGORY_COLORS[clean]) {
      return GFMP_CATEGORY_COLORS[clean];
    }

    // 2. Keyword heuristic search for cross-language consistency
    if (clean.includes('logement') || clean.includes('loyer') || clean.includes('hous') || clean.includes('rent') || clean.includes('vivien') || clean.includes('alquil') || clean.includes('habit') || clean.includes('mora')) {
      return '#3B82F6';
    }
    if (clean.includes('aliment') || clean.includes('superm') || clean.includes('food') || clean.includes('groc') || clean.includes('comid') || clean.includes('restau')) {
      return '#F59E0B';
    }
    if (clean.includes('transport') || clean.includes('essenc') || clean.includes('fuel') || clean.includes('gas') || clean.includes('auto') || clean.includes('combust') || clean.includes('veíc')) {
      return '#8B5CF6';
    }
    if (clean.includes('sant') || clean.includes('mutuel') || clean.includes('health') || clean.includes('medic') || clean.includes('salud') || clean.includes('saúd') || clean.includes('pharma')) {
      return '#10B981';
    }
    if (clean.includes('loisir') || clean.includes('sorti') || clean.includes('vacan') || clean.includes('leisur') || clean.includes('entert') || clean.includes('ocio') || clean.includes('lazer') || clean.includes('voyag') || clean.includes('holi')) {
      return '#EC4899';
    }
    if (clean.includes('impot') || clean.includes('impôt') || clean.includes('tax') || clean.includes('impuest') || clean.includes('tribut') || clean.includes('fisc')) {
      return '#F97316';
    }
    if (clean.includes('abonn') || clean.includes('subscri') || clean.includes('stream') || clean.includes('assin')) {
      return '#06B6D4';
    }
    if (clean.includes('educ') || clean.includes('format') || clean.includes('school') || clean.includes('cours') || clean.includes('ensin')) {
      return '#6366F1';
    }
    if (clean.includes('habil') || clean.includes('vetem') || clean.includes('cloth') || clean.includes('fash') || clean.includes('ropa') || clean.includes('vest')) {
      return '#14B8A6';
    }
    if (clean.includes('cadeau') || clean.includes('don') || clean.includes('gift') || clean.includes('regal') || clean.includes('presen')) {
      return '#E11D48';
    }
    if (clean.includes('autre') || clean.includes('other') || clean.includes('otro') || clean.includes('outro')) {
      return '#94A3B8';
    }

    // 3. Deterministic hash for custom user categories
    let hash = 0;
    for (let i = 0; i < raw.length; i++) {
      hash = ((hash << 5) - hash) + raw.charCodeAt(i);
      hash |= 0;
    }
    const idx = Math.abs(hash) % GFMP_LUXURY_PALETTE.length;
    return GFMP_LUXURY_PALETTE[idx];
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
    const targetTrs = ['FR', 'EN', 'ES', 'PT'].map(l => this.cleanText(this.trEntity(nom, l)).toLowerCase());

    return collection.find(item => {
      const rawNom = this.cleanText(item.nom).toLowerCase();
      if (rawNom === target || targetTrs.includes(rawNom)) return true;
      const itemTrs = ['FR', 'EN', 'ES', 'PT'].map(l => this.cleanText(this.trEntity(item.nom, l)).toLowerCase());
      return itemTrs.some(tr => tr === target || targetTrs.includes(tr));
    }) || null;
  }

  getCurrencySymbol(currency = this.currency) {
    const symbols = {
      MUR: 'Rs',
      EUR: '€',
      USD: '$',
      GBP: '£',
      ZAR: 'R',
      XOF: 'CFA'
    };
    return symbols[currency] || currency;
  }

  getStorage() {
    try {
      if (typeof localStorage !== 'undefined' && localStorage) return localStorage;
      if (typeof window !== 'undefined' && window.localStorage) return window.localStorage;
    } catch (e) {}
    return null;
  }

  resolvePeriod(period = null) {
    if (period && typeof period === 'string') {
      const trimmed = period.trim();
      if (/^\d{4}-\d{2}$/.test(trimmed)) {
        const [y, m] = trimmed.split('-');
        return { year: y, month: m };
      }
      if (/^\d{4}$/.test(trimmed)) {
        return { year: trimmed, month: '00' };
      }
    }
    return {
      year: String(this.year || '2026'),
      month: String(this.month || '02').padStart(2, '0')
    };
  }

  init(customData = null) {
    const storage = this.getStorage();
    let isExplicitEmpty = false;

    if (customData && typeof customData === 'object' && Array.isArray(customData.transactions)) {
      this.data = JSON.parse(JSON.stringify(customData));
      isExplicitEmpty = (this.data.dataMode === 'EMPTY') || Boolean(this.data.config && this.data.config.SYS_DATA_MODE === 'EMPTY');
    } else {
      try {
        const storedMode = storage ? storage.getItem('gfmp_data_mode') : null;
        if (storedMode === 'EMPTY') isExplicitEmpty = true;

        // 1. Try reading saved dataset from phone storage
        const stored = storage ? storage.getItem('gfmp_dataset_v3') : null;
        if (stored) {
          const parsed = JSON.parse(stored);
          const parsedEmpty = parsed && (parsed.dataMode === 'EMPTY' || (parsed.config && parsed.config.SYS_DATA_MODE === 'EMPTY') || isExplicitEmpty);
          if (parsed && Array.isArray(parsed.transactions) && (parsed.transactions.length > 0 || parsedEmpty)) {
            this.data = parsed;
            isExplicitEmpty = Boolean(parsedEmpty && parsed.transactions.length === 0);
            console.log('[GFMPEngine] Loaded dataset from phone storage:', this.data.transactions.length, 'transactions (mode:', isExplicitEmpty ? 'EMPTY' : 'DEMO', ')');
          }
        }
      } catch (e) {
        console.warn('[GFMPEngine] LocalStorage read failed:', e);
      }
    }

    // 2. Synchronous & bulletproof fallback to embedded baseline data (unless user explicitly chose EMPTY portfolio)
    if (!this.data || !Array.isArray(this.data.transactions) || (this.data.transactions.length === 0 && !isExplicitEmpty)) {
      if (typeof window !== 'undefined' && window.GFMP_BASELINE) {
        this.data = JSON.parse(JSON.stringify(window.GFMP_BASELINE));
        if (isExplicitEmpty) {
          this.data.transactions = [];
          this.data.budgets = [];
          this.data.dataMode = 'EMPTY';
          if (this.data.config) this.data.config.SYS_DATA_MODE = 'EMPTY';
        } else {
          this.data.dataMode = 'DEMO';
          if (this.data.config) this.data.config.SYS_DATA_MODE = 'DEMO';
        }
        console.log('[GFMPEngine] Initialized from embedded baseline data:', this.data.transactions.length, 'transactions');
        this.save();
      } else {
        this.data = {
          app_name: 'GFMP',
          dataMode: isExplicitEmpty ? 'EMPTY' : 'DEMO',
          comptes: [
            { id: '1', nom: 'Compte Courant Principal', classe: 'LIQUIDITE', devise: 'MUR', actif: 'OUI' },
            { id: '2', nom: 'Livret d\'Épargne', classe: 'LIQUIDITE', devise: 'MUR', actif: 'OUI' },
            { id: '4', nom: 'Assurance Vie', classe: 'INVESTISSEMENT', devise: 'EUR', actif: 'OUI' },
            { id: '5', nom: 'PEA / Actions', classe: 'INVESTISSEMENT', devise: 'EUR', actif: 'OUI' },
            { id: '7', nom: 'Prêt Immobilier', classe: 'DETTE', devise: 'MUR', actif: 'OUI' }
          ],
          categories: [
            { id: '1', nom: 'Salaire / Revenus Pro', type: 'REVENU' },
            { id: '2', nom: 'Dividendes / Placements', type: 'REVENU' },
            { id: '4', nom: 'Logement (Loyer/Prêt/Charges)', type: 'DEPENSE' },
            { id: '5', nom: 'Alimentation & Supermarché', type: 'DEPENSE' },
            { id: '6', nom: 'Transports (Essence/Assurance)', type: 'DEPENSE' },
            { id: '7', nom: 'Santé & Mutuelle', type: 'DEPENSE' },
            { id: '8', nom: 'Loisirs, Sorties & Vacances', type: 'DEPENSE' }
          ],
          tiers: [
            { id: '1', nom: 'Employeur principal', type: 'ENTREPRISE' },
            { id: '2', nom: 'Banque / Courtier', type: 'BANQUE' },
            { id: '4', nom: 'Supermarché / Alimentaire', type: 'COMMERCE' },
            { id: '5', nom: 'Propriétaire / Syndic', type: 'IMMOBILIER' }
          ],
          transactions: isExplicitEmpty ? [] : [
            { id: 1, date: '2026-02-05', id_compte: '1', id_categorie: '4', id_tiers: '5', montant: 12000, devise: 'MUR', notes: 'Loyer Résidence KIMYA' },
            { id: 2, date: '2026-02-12', id_compte: '5', id_categorie: '2', id_tiers: '2', montant: 4500, devise: 'MUR', notes: 'Dividendes Portefeuille PEA' },
            { id: 3, date: '2026-02-15', id_compte: '1', id_categorie: '6', id_tiers: '4', montant: 2800, devise: 'MUR', notes: 'Assurance & Entretien Auto' },
            { id: 4, date: '2026-02-18', id_compte: '1', id_categorie: '8', id_tiers: '4', montant: 3500, devise: 'MUR', notes: 'Sorties & Restaurants' },
            { id: 5, date: '2026-02-23', id_compte: '1', id_categorie: '1', id_tiers: '1', montant: 78000, devise: 'MUR', notes: 'Salaire Février KIMYA ÉDITION' },
            { id: 6, date: '2026-02-23', id_compte: '4', id_categorie: '7', id_tiers: '2', montant: 7000, devise: 'MUR', notes: 'Santé & Mutuelle Famille' },
            { id: 7, date: '2026-02-23', id_compte: '3', id_categorie: '5', id_tiers: '4', montant: 4000, devise: 'MUR', notes: 'Courses Hebdomadaires' }
          ],
          budgets: isExplicitEmpty ? [] : [
            { id: 1, mois: '2026-02', id_categorie: '4', montant: 12000 },
            { id: 2, mois: '2026-02', id_categorie: '5', montant: 8000 },
            { id: 3, mois: '2026-02', id_categorie: '6', montant: 4000 },
            { id: 4, mois: '2026-02', id_categorie: '7', montant: 7000 },
            { id: 5, mois: '2026-02', id_categorie: '8', montant: 5000 }
          ],
          devises: { MUR: 12.2243, EUR: 656.168, USD: 567.5369, GBP: 757.0023, ZAR: 33.2723, XOF: 1 },
          dictionary: {},
          config: { LANGUE_UI: 'FR', SYS_DEVISE_BASE: 'MUR', SYS_DATA_MODE: isExplicitEmpty ? 'EMPTY' : 'DEMO' }
        };
        this.save();
      }
    }

    if (this.data) {
      const resolvedMode = isExplicitEmpty
        ? 'EMPTY'
        : ((this.data.dataMode === 'EMPTY' || (this.data.config && this.data.config.SYS_DATA_MODE === 'EMPTY')) ? 'EMPTY' : 'DEMO');
      this.dataMode = resolvedMode;
      this.data.dataMode = resolvedMode;
      if (!this.data.config) this.data.config = {};
      this.data.config.SYS_DATA_MODE = resolvedMode;
    }

    if (this.data && !this.data.devises) {
      this.data.devises = { MUR: 12.2243, EUR: 656.168, USD: 567.5369, GBP: 757.0023, ZAR: 33.2723, XOF: 1 };
    }

    // Apply config
    if (this.data && this.data.config) {
      if (this.data.config.LANGUE_UI) this.lang = this.data.config.LANGUE_UI;
      if (this.data.config.SYS_DEVISE_BASE) this.currency = this.data.config.SYS_DEVISE_BASE;
      if (this.data.config.SYS_DATE_REF) this.referenceDate = this.data.config.SYS_DATE_REF;
      if (this.data.config.SYS_YEAR) this.year = String(this.data.config.SYS_YEAR);
      if (this.data.config.SYS_MONTH) this.month = String(this.data.config.SYS_MONTH).padStart(2, '0');
      if (this.data.config.SYS_THEME) this.theme = this.data.config.SYS_THEME;
    }

    // Check if initial onboarding setup has been completed (versioned v5.6 + native APK install trigger)
    const onboardingFlag = storage ? storage.getItem('gfmp_onboarding_v5_6') : 'true';
    const forceOnboardingUrl = (typeof window !== 'undefined' && window.location && typeof window.location.search === 'string' && window.location.search.includes('onboarding=1'));
    this.needsOnboarding = forceOnboardingUrl || (onboardingFlag !== 'true');

    // Load cached live rates if available (normalized to XOF=1 base scale)
    try {
      const cached = storage ? storage.getItem('gfmp_live_rates') : null;
      if (cached) {
        const parsed = JSON.parse(cached);
        if (parsed && parsed.devises) {
          const cachedRates = { ...parsed.devises };
          if (cachedRates.EUR && cachedRates.EUR <= 2) {
            const scale = 656.168;
            Object.keys(cachedRates).forEach(k => {
              cachedRates[k] = Number((cachedRates[k] * scale).toFixed(4));
            });
            cachedRates.XOF = 1;
          }
          this.data.devises = Object.assign({}, this.data.devises, cachedRates);
          this.lastExchangeSync = parsed.lastSync;
        }
      }
    } catch (e) {}

    // Live exchange rate status
    this.rateStatus = {
      isLive: !!this.lastExchangeSync,
      lastSync: this.lastExchangeSync || null,
      loading: false,
      error: null
    };

    // Auto-fetch fresh rates silently if online
    if (typeof navigator !== 'undefined' && navigator.onLine) {
      setTimeout(() => {
        this.fetchLiveExchangeRates(false);
      }, 300);
    }
  }

  save() {
    try {
      const storage = this.getStorage();
      if (this.data && storage) {
        this.data.dataMode = this.dataMode || this.data.dataMode || 'DEMO';
        if (this.data.config) {
          this.data.config.SYS_DATA_MODE = this.data.dataMode;
        }
        storage.setItem('gfmp_dataset_v3', JSON.stringify(this.data));
        storage.setItem('gfmp_data_mode', this.data.dataMode);
      }
    } catch (e) {
      console.warn('[GFMPEngine] Save failed:', e);
    }
  }

  getDataMode() {
    return this.dataMode === 'EMPTY' ? 'EMPTY' : 'DEMO';
  }

  hasFinancialData(period = null) {
    if (!this.data) return false;
    if (!period) {
      const hasTx = Array.isArray(this.data.transactions) && this.data.transactions.length > 0;
      const hasBg = Array.isArray(this.data.budgets) && this.data.budgets.length > 0;
      return hasTx || hasBg;
    }
    const { year: targetYear, month: targetMonth } = this.resolvePeriod(period);
    const isAllYear = (targetMonth === '00');
    const prefix = isAllYear ? `${targetYear}-` : `${targetYear}-${targetMonth}`;
    const hasTxPeriod = Array.isArray(this.data.transactions) && this.data.transactions.some(tx => tx.date && tx.date.startsWith(prefix));
    const hasBgPeriod = Array.isArray(this.data.budgets) && this.data.budgets.some(b => {
      const bm = b.mois || b.mois_annee || '';
      return isAllYear ? bm.startsWith(`${targetYear}-`) : bm === `${targetYear}-${targetMonth}`;
    });
    return hasTxPeriod || hasBgPeriod;
  }

  setDataMode(mode = 'DEMO') {
    const cleanMode = (String(mode).trim().toUpperCase() === 'EMPTY' || mode === false) ? 'EMPTY' : 'DEMO';
    this.dataMode = cleanMode;
    this.selectedCategoryFilter = null;

    if (!this.data) {
      this.init();
    }

    if (cleanMode === 'EMPTY') {
      const prevCount = Array.isArray(this.data.transactions) ? this.data.transactions.length : 0;
      this.data.transactions = [];
      this.data.budgets = [];
      this.data.dataMode = 'EMPTY';
      if (!this.data.config) this.data.config = {};
      this.data.config.SYS_DATA_MODE = 'EMPTY';
      this.logHistory('DATA_MODE_EMPTY', 'database', 'system', `Activation du Portefeuille Personnel Vierge (${prevCount} opérations de test retirées)`);
    } else {
      const baseline = (typeof window !== 'undefined' && window.GFMP_BASELINE) ? JSON.parse(JSON.stringify(window.GFMP_BASELINE)) : null;
      if (baseline) {
        this.data.transactions = Array.isArray(baseline.transactions) ? baseline.transactions : [];
        this.data.budgets = Array.isArray(baseline.budgets) ? baseline.budgets : [];
        if (!Array.isArray(this.data.comptes) || this.data.comptes.length === 0) {
          this.data.comptes = baseline.comptes || [];
        }
        if (!Array.isArray(this.data.categories) || this.data.categories.length === 0) {
          this.data.categories = baseline.categories || [];
        }
        if (!Array.isArray(this.data.tiers) || this.data.tiers.length === 0) {
          this.data.tiers = baseline.tiers || [];
        }
      }
      this.data.dataMode = 'DEMO';
      if (!this.data.config) this.data.config = {};
      this.data.config.SYS_DATA_MODE = 'DEMO';
      this.logHistory('DATA_MODE_DEMO', 'database', 'system', `Activation des Données de Test & Démo (${this.data.transactions.length} opérations chargées)`);
    }

    this.save();
    this.notify();
    return {
      success: true,
      dataMode: this.dataMode,
      txCount: Array.isArray(this.data.transactions) ? this.data.transactions.length : 0,
      budgetCount: Array.isArray(this.data.budgets) ? this.data.budgets.length : 0
    };
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
  getDictionary() {
    return (this.data && this.data.dictionary) ? this.data.dictionary : {};
  }

  tr(key, params = {}, langOverride = null) {
    if (!key) return '';
    const cleanKey = String(key).trim();
    const targetLang = (langOverride && ['FR', 'EN', 'ES', 'PT'].includes(String(langOverride).toUpperCase()))
      ? String(langOverride).toUpperCase()
      : this.lang;

    const fallbackDict = {
      // Header & Global
      'APP_TITLE': { FR: 'SYSTÈME FINANCIER PERSONNEL', EN: 'PERSONAL FINANCE SYSTEM', ES: 'SISTEMA FINANCIERO PERSONAL', PT: 'SISTEMA FINANCEIRO PESSOAL' },
      'WELCOME': {
        FR: 'Plateforme de Gestion de Fortune & Pilotage Financier Décisionnel.',
        EN: 'Wealth Management & Financial Decision Platform.',
        ES: 'Plataforma de Gestión Patrimonial y Finanzas de Precisión.',
        PT: 'Plataforma de Gestão Patrimonial e Finanças de Precisão.'
      },
      'HERO_GREETING': {
        FR: 'Bonjour, Bienvenue sur votre Hub',
        EN: 'Welcome to your Central Hub',
        ES: 'Bienvenido a su Panel Central',
        PT: 'Bem-vindo ao seu Painel Central'
      },
      'FULL_YEAR': { FR: 'Année {year}', EN: 'Full Year {year}', ES: 'Año Completo {year}', PT: 'Ano Inteiro {year}' },
      'UPDATE_RATES': { FR: 'Actualiser Taux', EN: 'Update Rates', ES: 'Actualizar Tasas', PT: 'Atualizar Taxas' },
      'MARKETS_CONNECTED': { FR: 'Marchés Connectés • Taux en Direct (API)', EN: 'Markets Connected • Live Rates (API)', ES: 'Mercados Conectados • Tasas en Vivo (API)', PT: 'Mercados Conectados • Taxas ao Vivo (API)' },
      'OFFLINE_RATES': { FR: 'Mode Hors-Ligne • Taux sauvegardés', EN: 'Offline Mode • Saved Rates', ES: 'Modo Fuera de Línea • Tasas guardadas', PT: 'Modo Offline • Taxas salvas' },
      'BACK_HUB': { FR: 'Hub Central', EN: 'Central Hub', ES: 'Panel Central', PT: 'Painel Central' },

      // Bottom Navigation
      'NAV_HOME': { FR: 'Hub', EN: 'Hub', ES: 'Panel', PT: 'Painel' },
      'NAV_SAISIE': { FR: 'Saisie', EN: 'Entry', ES: 'Entrada', PT: 'Inserir' },
      'NAV_CASHFLOW': { FR: 'Cashflow', EN: 'Cashflow', ES: 'Cashflow', PT: 'Fluxo' },
      'NAV_NETWORTH': { FR: 'Patrimoine', EN: 'Net Worth', ES: 'Patrimonio', PT: 'Patrimônio' },
      'NAV_BUDGET': { FR: 'Budget', EN: 'Budget', ES: 'Presupuesto', PT: 'Orçamento' },
      'NAV_SETTINGS': { FR: 'Réglages', EN: 'Settings', ES: 'Ajustes', PT: 'Configurações' },

      // Hub Filter Pills
      'HUB_ALL': { FR: 'Vue Consolidée', EN: 'Consolidated View', ES: 'Vista Consolidada', PT: 'Visão Consolidada' },
      'HUB_CASHFLOW': { FR: 'Trésorerie', EN: 'Cashflow', ES: 'Tesorería', PT: 'Tesouraria' },
      'HUB_PATRIMOINE': { FR: 'Patrimoine', EN: 'Net Worth', ES: 'Patrimonio', PT: 'Patrimônio' },
      'HUB_BUDGET': { FR: 'Budgets', EN: 'Budgets', ES: 'Presupuestos', PT: 'Orçamentos' },

      // Hub Metrics & Headings
      'PORTFOLIO_CONSOLIDATED': {
        FR: 'Patrimoine Net Consolidé',
        EN: 'Consolidated Net Worth',
        ES: 'Patrimonio Neto Consolidado',
        PT: 'Patrimônio Líquido Consolidado'
      },
      'NET_CASHFLOW': { FR: 'FLUX NET MENSUEL', EN: 'MONTHLY NET FLOW', ES: 'FLUJO NETO MENSUAL', PT: 'FLUXO LÍQUIDO MENSAL' },
      'BUDGET_CONSUMPTION': { FR: 'CONSOMMATION BUDGET', EN: 'BUDGET CONSUMPTION', ES: 'CONSUMO PRESUPUESTO', PT: 'CONSUMO ORÇAMENTO' },
      'TOTAL_ASSETS': { FR: 'TOTAL ACTIFS', EN: 'TOTAL ASSETS', ES: 'TOTAL ACTIVOS', PT: 'TOTAL ATIVOS' },
      'TOTAL_DEBTS': { FR: 'TOTAL DETTES', EN: 'TOTAL DEBTS', ES: 'TOTAL DEUDAS', PT: 'TOTAL DÍVIDAS' },
      'ASSETS_VS_DEBTS': { FR: 'Répartition Actifs vs Dettes', EN: 'Assets vs Debts Breakdown', ES: 'Distribución Activos vs Deudas', PT: 'Distribuição Ativos vs Dívidas' },
      'SOLVENCY_SCORE': {
        FR: 'Solvabilité : {sol}% • Score : {score}/100',
        EN: 'Solvency: {sol}% • Score: {score}/100',
        ES: 'Solvencia: {sol}% • Puntuación: {score}/100',
        PT: 'Solvência: {sol}% • Pontuação: {score}/100'
      },
      'HUB_TREND_TITLE': { FR: 'ÉVOLUTION DES FLUX MENSUELS (TRÉSORERIE)', EN: 'MONTH CASHFLOW TREND', ES: 'EVOLUCIÓN MENSUAL DE FLUJOS', PT: 'EVOLUÇÃO MENSAL DOS FLUXOS' },
      'HUB_DONUT_TITLE': { FR: 'RÉPARTITION DES DÉPENSES PAR CATÉGORIE', EN: 'EXPENSE BREAKDOWN BY CATEGORY', ES: 'DISTRIBUCIÓN DE GASTOS POR CATEGORÍA', PT: 'DISTRIBUIÇÃO DE DESPESAS POR CATEGORIA' },
      'HUB_PATRIMOINE_TITLE': { FR: 'ALLOCATION D\'ACTIFS & STRUCTURE PATRIMONIALE', EN: 'ASSET ALLOCATION & WEALTH STRUCTURE', ES: 'DISTRIBUCIÓN DE ACTIVOS Y PATRIMONIO', PT: 'ALOCAÇÃO DE ATIVOS E PATRIMÔNIO' },
      'HUB_BUDGET_TITLE': { FR: 'PILOTAGE & SANTÉ DES ENVELOPPES BUDGÉTAIRES', EN: 'BUDGET ENVELOPES HEALTH & TRACKING', ES: 'GESTIÓN Y SALUD PRESUPUESTARIA', PT: 'GESTÃO E SAÚDE ORÇAMENTÁRIA' },
      'HUB_RECENT_TITLE': { FR: 'DERNIÈRES TRANSACTIONS ENREGISTRÉES', EN: 'LATEST RECORDED TRANSACTIONS', ES: 'ÚLTIMAS TRANSACCIONES REGISTRADAS', PT: 'ÚLTIMAS TRANSAÇÕES REGISTRADAS' },
      'INIT_BADGE': { FR: 'SFP v1.0 • ÉDITION GOLD PRIVÉE', EN: 'SFP v1.0 • PRIVATE GOLD EDITION', ES: 'SFP v1.0 • EDICIÓN GOLD PRIVADA', PT: 'SFP v1.0 • EDIÇÃO GOLD PRIVADA' },
      'INIT_TITLE': { FR: 'INITIALISATION DU SYSTÈME', EN: 'SYSTEM INITIALIZATION', ES: 'INICIALIZACIÓN DEL SISTEMA', PT: 'INICIALIZAÇÃO DO SISTEMA' },
      'INIT_SUBTITLE': {
        FR: 'Configurez votre environnement de gestion patrimoniale et financière haute précision.',
        EN: 'Configure your high-precision wealth and personal finance management environment.',
        ES: 'Configure su entorno de gestión patrimonial y financiera de alta precisión.',
        PT: 'Configure o seu ambiente de gestão patrimonial e financeira de alta precisão.'
      },
      'INIT_TRUST_1': { FR: '100% Hors-Ligne & Privé', EN: '100% Offline & Private', ES: '100% Sin Conexión y Privado', PT: '100% Offline e Privado' },
      'INIT_TRUST_2': { FR: 'Multi-Devises Live API', EN: 'Live Multi-Currency API', ES: 'API Multidivisa en Vivo', PT: 'API Multimoeda ao Vivo' },
      'INIT_TRUST_3': { FR: 'Architecture ZBB & Audit', EN: 'ZBB & Audit Architecture', ES: 'Arquitectura ZBB y Auditoría', PT: 'Arquitetura ZBB e Auditoria' },
      'INIT_STEP1_BADGE': { FR: 'ÉTAPE 01 / 05', EN: 'STEP 01 / 05', ES: 'PASO 01 / 05', PT: 'PASSO 01 / 05' },
      'INIT_STEP1_TITLE': { FR: 'Sélectionnez la langue du système', EN: 'Select your system language', ES: 'Seleccione el idioma del sistema', PT: 'Selecione o idioma do sistema' },
      'INIT_STEP1_DESC': {
        FR: 'L\'ensemble de vos synthèses, rapports et interfaces s\'adapteront immédiatement.',
        EN: 'All your dashboards, financial reports, and interfaces will adapt immediately.',
        ES: 'Todos sus paneles, informes financieros e interfaces se adaptarán de inmediato.',
        PT: 'Todos os seus painéis, relatórios financeiros e interfaces adaptar-se-ão imediatamente.'
      },
      'INIT_STEP2_BADGE': { FR: 'ÉTAPE 02 / 05', EN: 'STEP 02 / 05', ES: 'PASO 02 / 05', PT: 'PASSO 02 / 05' },
      'INIT_STEP2_TITLE': { FR: 'Sélectionnez votre devise de référence', EN: 'Select your primary reference currency', ES: 'Seleccione su divisa principal de referencia', PT: 'Selecione a sua moeda principal de referência' },
      'INIT_STEP2_DESC': {
        FR: 'Devise mère sur laquelle l\'ensemble de votre patrimoine et flux seront consolidés.',
        EN: 'Base currency in which all your net worth, assets, and cashflows will be consolidated.',
        ES: 'Divisa base en la que se consolidarán todo su patrimonio, activos y flujos.',
        PT: 'Moeda base na qual todo o seu patrimônio, ativos e fluxos serão consolidados.'
      },
      'INIT_STEP3_BADGE': { FR: 'ÉTAPE 03 / 05', EN: 'STEP 03 / 05', ES: 'PASO 03 / 05', PT: 'PASSO 03 / 05' },
      'INIT_STEP3_TITLE': { FR: 'Configurez la date et la période financière', EN: 'Configure date & active financial period', ES: 'Configure la fecha y el período financiero', PT: 'Configure a data e o período financeiro' },
      'INIT_STEP3_DESC': {
        FR: 'Définissez la date système par défaut pour vos saisies et le mois/année d\'ouverture du tableau de bord.',
        EN: 'Set the default system date for your entries and the opening month/year for your dashboard.',
        ES: 'Defina la fecha del sistema por defecto para sus registros y el mes/año de apertura del panel.',
        PT: 'Defina a data do sistema padrão para os seus registos e o mês/ano de abertura do painel.'
      },
      'INIT_DATE_LABEL': { FR: 'Date de référence du système', EN: 'System reference date', ES: 'Fecha de referencia del sistema', PT: 'Data de referência do sistema' },
      'INIT_MONTH_LABEL': { FR: 'Mois de pilotage', EN: 'Active month', ES: 'Mes de control', PT: 'Mês de controlo' },
      'INIT_YEAR_LABEL': { FR: 'Exercice (Année)', EN: 'Fiscal year', ES: 'Ejercicio (Año)', PT: 'Exercício (Ano)' },
      'INIT_BTN_TODAY': { FR: 'Aujourd\'hui', EN: 'Today', ES: 'Hoy', PT: 'Hoje' },
      'INIT_BTN_DEMO': { FR: 'Période Démo SFP (Fév 2026)', EN: 'SFP Demo Period (Feb 2026)', ES: 'Período Demo SFP (Feb 2026)', PT: 'Período Demo SFP (Fev 2026)' },
      'INIT_STEP4_BADGE': { FR: 'ÉTAPE 04 / 05', EN: 'STEP 04 / 05', ES: 'PASO 04 / 05', PT: 'PASSO 04 / 05' },
      'INIT_STEP4_TITLE': { FR: 'Choisissez le thème visuel de l\'application', EN: 'Choose the application visual theme', ES: 'Elija el tema visual de la aplicación', PT: 'Escolha o tema visual da aplicação' },
      'INIT_STEP4_DESC': {
        FR: 'Personnalisez l\'ambiance Haute Joaillerie de votre interface avec aperçu instantané en direct.',
        EN: 'Customize the Haute Joaillerie ambiance of your interface with instant live preview.',
        ES: 'Personalice el ambiente de Alta Joyería de su interfaz con vista previa instantánea en vivo.',
        PT: 'Personalize o ambiente de Alta Joalharia da sua interface com pré-visualização instantânea ao vivo.'
      },
      'INIT_STEP5_BADGE': { FR: 'ÉTAPE 05 / 05', EN: 'STEP 05 / 05', ES: 'PASO 05 / 05', PT: 'PASSO 05 / 05' },
      'INIT_STEP5_TITLE': { FR: 'Choisissez le mode de données financières', EN: 'Choose your financial data mode', ES: 'Elija el modo de datos financieros', PT: 'Escolha o modo de dados financeiros' },
      'INIT_STEP5_DESC': {
        FR: 'Explorez l\'application avec le portefeuille de test complet ou démarrez avec un portefeuille vierge prêt pour vos saisies.',
        EN: 'Explore the application with the full test portfolio or start with a blank portfolio ready for your entries.',
        ES: 'Explore la aplicación con la cartera de prueba completa o comience con una cartera en blanco lista para sus registros.',
        PT: 'Explore a aplicação com a carteira de teste completa ou comece com uma carteira em branco pronta para os seus lançamentos.'
      },
      'DATA_MODE_DEMO_BADGE': { FR: '60 OPÉRATIONS • 17 BUDGETS', EN: '60 TRANSACTIONS • 17 BUDGETS', ES: '60 OPERACIONES • 17 PRESUPUESTOS', PT: '60 TRANSAÇÕES • 17 ORÇAMENTOS' },
      'DATA_MODE_DEMO_TITLE': { FR: 'Données de Test & Démo SFP', EN: 'SFP Test & Demo Dataset', ES: 'Datos de Prueba y Demo SFP', PT: 'Dados de Teste e Demo SFP' },
      'DATA_MODE_DEMO_SUB': {
        FR: 'Portefeuille d\'exemple pré-rempli (Jan–Mar 2026) pour tester les graphiques, rapports et analyses.',
        EN: 'Pre-filled sample portfolio (Jan–Mar 2026) to explore charts, reports, and analytics.',
        ES: 'Cartera de ejemplo precargada (Ene–Mar 2026) para explorar gráficos, informes y análisis.',
        PT: 'Carteira de exemplo pré-preenchida (Jan–Mar 2026) para explorar gráficos, relatórios e análises.'
      },
      'DATA_MODE_EMPTY_BADGE': { FR: 'DÉMARRAGE À ZÉRO • 0,00', EN: 'CLEAN SLATE • 0.00', ES: 'DESDE CERO • 0,00', PT: 'DO ZERO • 0,00' },
      'DATA_MODE_EMPTY_TITLE': { FR: 'Portefeuille Personnel Vierge', EN: 'Blank Personal Portfolio', ES: 'Cartera Personal en Blanco', PT: 'Carteira Pessoal em Branco' },
      'DATA_MODE_EMPTY_SUB': {
        FR: 'Aucune transaction ni budget de test. Comptes, catégories et devises prêts pour vos propres opérations.',
        EN: 'Zero test transactions or budgets. Accounts, categories, and currencies ready for your own entries.',
        ES: 'Sin transacciones ni presupuestos de prueba. Cuentas, categorías y divisas listas para sus propios datos.',
        PT: 'Sem transações nem orçamentos de teste. Contas, categorias e moedas prontas para os seus próprios dados.'
      },
      'SOLVENCY_EMPTY': {
        FR: 'Portefeuille vierge • Aucune opération enregistrée sur cette période',
        EN: 'Blank portfolio • No transactions recorded for this period',
        ES: 'Cartera en blanco • Ninguna operación registrada en este período',
        PT: 'Carteira em branco • Nenhuma transação registrada neste período'
      },
      'TREND_EMPTY_TITLE': {
        FR: 'Aucun flux de trésorerie sur les 6 derniers mois',
        EN: 'No cashflow recorded over the last 6 months',
        ES: 'Sin flujos de tesorería en los últimos 6 meses',
        PT: 'Sem fluxos de caixa nos últimos 6 meses'
      },
      'TREND_EMPTY_SUB': {
        FR: 'Saisissez votre premier revenu ou dépense pour activer la courbe d\'évolution mensuelle.',
        EN: 'Record your first income or expense to activate the monthly trend chart.',
        ES: 'Registre su primer ingreso o gasto para activar el gráfico de evolución mensual.',
        PT: 'Registe a sua primeira receita ou despesa para ativar o gráfico de evolução mensal.'
      },
      'DONUT_EMPTY_HINT': {
        FR: 'Enregistrez une dépense sur cette période ou changez de mois pour visualiser la répartition.',
        EN: 'Record an expense for this period or switch months to visualize the breakdown.',
        ES: 'Registre un gasto en este período o cambie de mes para visualizar la distribución.',
        PT: 'Registe uma despesa neste período ou mude de mês para visualizar a distribuição.'
      },
      'EMPTY_PATRIMOINE_SUB': {
        FR: 'Tous vos comptes affichent un solde de 0,00. Saisissez un revenu ou apport initial pour valoriser votre patrimoine.',
        EN: 'All accounts currently have a 0.00 balance. Record an income or initial deposit to value your net worth.',
        ES: 'Todas sus cuentas muestran un saldo de 0,00. Registre un ingreso o aporte inicial para valorar su patrimonio.',
        PT: 'Todas as suas contas apresentam saldo de 0,00. Registe uma receita ou aporte inicial para valorizar o seu patrimônio.'
      },
      'EMPTY_BUDGET_SUB': {
        FR: 'Définissez vos enveloppes budgétaires mensuelles (ZBB) pour suivre vos plafonds de dépenses en temps réel.',
        EN: 'Define your monthly budget envelopes (ZBB) to track your spending limits in real time.',
        ES: 'Defina sus sobres presupuestarios mensuales (ZBB) para controlar sus límites de gasto en tiempo real.',
        PT: 'Defina os seus envelopes orçamentários mensais (ZBB) para acompanhar os seus limites de gastos em tempo real.'
      },
      'EMPTY_TX_SUB': {
        FR: 'Votre journal d\'opérations est vierge pour cette sélection. Ajoutez une opération ou ajustez la période.',
        EN: 'Your transaction log is empty for this selection. Add a transaction or adjust the active period.',
        ES: 'Su registro de operaciones está vacío para esta selección. Añada una operación o ajuste el período.',
        PT: 'O seu registo de transações está vazio para esta seleção. Adicione uma transação ou ajuste o período.'
      },
      'EMPTY_CTA_ADD_TX': { FR: '+ Saisir une opération', EN: '+ Add transaction', ES: '+ Registrar operación', PT: '+ Inserir transação' },
      'EMPTY_CTA_ADD_EXP': { FR: '+ Saisir une dépense', EN: '+ Add expense', ES: '+ Registrar gasto', PT: '+ Inserir despesa' },
      'EMPTY_CTA_ALLOC_BG': { FR: '+ Allouer un budget', EN: '+ Allocate budget', ES: '+ Asignar presupuesto', PT: '+ Alocar orçamento' },
      'EMPTY_CTA_ADD_ACC': { FR: '+ Créer un compte', EN: '+ Create account', ES: '+ Crear cuenta', PT: '+ Criar conta' },
      'EMPTY_CTA_RESET_FILTER': { FR: 'Réinitialiser les filtres', EN: 'Reset filters', ES: 'Restablecer filtros', PT: 'Repor filtros' },
      'EMPTY_CTA_FULL_YEAR': { FR: 'Voir toute l\'année', EN: 'View full year', ES: 'Ver año completo', PT: 'Ver ano inteiro' },
      'EMPTY_CTA_LOAD_DEMO': { FR: 'Charger les données de test', EN: 'Load test dataset', ES: 'Cargar datos de prueba', PT: 'Carregar dados de teste' },
      'EMPTY_CTA_SWITCH_EMPTY': { FR: 'Vider les données de test', EN: 'Clear test data (Start at 0)', ES: 'Vaciar datos de prueba', PT: 'Limpar dados de teste' },
      'BANNER_DEMO_ACTIVE': {
        FR: 'Mode Données de Test & Démo actif (60 opérations d\'exemple)',
        EN: 'Test & Demo Dataset Mode active (60 sample transactions)',
        ES: 'Modo Datos de Prueba y Demo activo (60 operaciones de ejemplo)',
        PT: 'Modo Dados de Teste e Demo ativo (60 transações de exemplo)'
      },
      'BANNER_EMPTY_ACTIVE': {
        FR: 'Portefeuille Personnel Vierge actif — Prêt pour vos opérations',
        EN: 'Blank Personal Portfolio active — Ready for your transactions',
        ES: 'Cartera Personal en Blanco activa — Lista para sus operaciones',
        PT: 'Carteira Pessoal em Branco ativa — Pronta para as suas transações'
      },
      'SETTINGS_DATAMODE_TITLE': { FR: 'MODE DE DONNÉES FINANCIÈRES', EN: 'FINANCIAL DATA MODE', ES: 'MODO DE DATOS FINANCIEROS', PT: 'MODO DE DADOS FINANCEIROS' },
      'SETTINGS_DATAMODE_DESC': {
        FR: 'Basculez librement entre les données de test SFP et un portefeuille personnel vierge sans perdre vos comptes ni vos catégories.',
        EN: 'Switch freely between the SFP test dataset and a blank personal portfolio without losing your accounts or categories.',
        ES: 'Cambie libremente entre los datos de prueba SFP y una cartera personal en blanco sin perder sus cuentas ni categorías.',
        PT: 'Alterne livremente entre os dados de teste SFP e uma carteira pessoal em branco sem perder as suas contas ou categorias.'
      },
      'REPORT_EMPTY_DIAG_TITLE': {
        FR: 'DIAGNOSTIC EXÉCUTIF : AUCUNE OPÉRATION SUR LA PÉRIODE',
        EN: 'EXECUTIVE DIAGNOSTIC: NO OPERATIONS FOR THIS PERIOD',
        ES: 'DIAGNÓSTICO EJECUTIVO: SIN OPERACIONES EN EL PERÍODO',
        PT: 'DIAGNÓSTICO EXECUTIVO: SEM TRANSAÇÕES NO PERÍODO'
      },
      'REPORT_EMPTY_DIAG_DESC': {
        FR: 'Aucun flux de revenus, de dépenses ou d\'enveloppe budgétaire n\'est enregistré pour cette période. Saisissez vos opérations ou sélectionnez l\'année entière pour générer une synthèse financière complète.',
        EN: 'No income, expense, or budget envelope is recorded for this period. Enter transactions or select the full year to generate a complete financial executive summary.',
        ES: 'No se han registrado ingresos, gastos ni presupuestos en este período. Registre operaciones o seleccione el año completo para generar un informe ejecutivo completo.',
        PT: 'Nenhuma receita, despesa ou orçamento foi registado neste período. Insira transações ou selecione o ano inteiro para gerar um relatório executivo completo.'
      },
      'THEME_NAME_BLUE': { FR: 'Bleu Royal', EN: 'Royal Blue', ES: 'Azul Real', PT: 'Azul Real' },
      'THEME_SUB_BLUE': { FR: 'Standard SFP • Or & Saphir', EN: 'SFP Standard • Gold & Sapphire', ES: 'Estándar SFP • Oro y Zafiro', PT: 'Padrão SFP • Ouro e Safira' },
      'THEME_NAME_GOLD': { FR: 'Or & Noir Impérial', EN: 'Imperial Gold & Black', ES: 'Oro y Negro Imperial', PT: 'Ouro e Preto Imperial' },
      'THEME_SUB_GOLD': { FR: 'Prestige • Obsidienne & Or', EN: 'Prestige • Obsidian & Gold', ES: 'Prestigio • Obsidiana y Oro', PT: 'Prestígio • Obsidiana e Ouro' },
      'THEME_NAME_EMERALD': { FR: 'Émeraude Wealth', EN: 'Emerald Wealth', ES: 'Esmeralda Wealth', PT: 'Esmeralda Wealth' },
      'THEME_SUB_EMERALD': { FR: 'Banque Privée • Émeraude', EN: 'Private Banking • Emerald', ES: 'Banca Privada • Esmeralda', PT: 'Private Banking • Esmeralda' },
      'THEME_NAME_AMETHYST': { FR: 'Améthyste Royale', EN: 'Royal Amethyst', ES: 'Amatista Real', PT: 'Ametista Real' },
      'THEME_SUB_AMETHYST': { FR: 'Haute Joaillerie • Pourpre', EN: 'Haute Joaillerie • Purple', ES: 'Alta Joyería • Púrpura', PT: 'Alta Joalharia • Púrpura' },
      'CURR_NAME_MUR': { FR: 'Roupie (Rs)', EN: 'Rupee (Rs)', ES: 'Rupia (Rs)', PT: 'Rupia (Rs)' },
      'CURR_NAME_EUR': { FR: 'Euro (€)', EN: 'Euro (€)', ES: 'Euro (€)', PT: 'Euro (€)' },
      'CURR_NAME_USD': { FR: 'Dollar ($)', EN: 'US Dollar ($)', ES: 'Dólar ($)', PT: 'Dólar ($)' },
      'CURR_NAME_GBP': { FR: 'Livre (£)', EN: 'Pound (£)', ES: 'Libra (£)', PT: 'Libra (£)' },
      'CURR_NAME_ZAR': { FR: 'Rand (R)', EN: 'Rand (R)', ES: 'Rand (R)', PT: 'Rand (R)' },
      'CURR_NAME_XOF': { FR: 'Franc CFA', EN: 'CFA Franc', ES: 'Franco CFA', PT: 'Franco CFA' },
      'INIT_SUMMARY_PREFIX': { FR: 'Configuration active :', EN: 'Active configuration:', ES: 'Configuración activa:', PT: 'Configuração ativa:' },
      'INIT_BTN_CONFIRM': { FR: 'CONFIRMER & LANCER L\'APPLICATION', EN: 'CONFIRM & LAUNCH APPLICATION', ES: 'CONFIRMAR E INICIAR APLICACIÓN', PT: 'CONFIRMAR E INICIAR APLICAÇÃO' },
      'TOP_ASSETS_LABEL': { FR: 'Principaux Actifs :', EN: 'Main Assets:', ES: 'Activos Principales:', PT: 'Principais Ativos:' },
      'NET_ASSETS': { FR: 'Actifs Nets', EN: 'Net Assets', ES: 'Activos Netos', PT: 'Ativos Líquidos' },
      'LIQUIDITIES': { FR: 'Liquidités', EN: 'Liquidities', ES: 'Liquidez', PT: 'Liquidez' },
      'INVESTMENTS': { FR: 'Placements', EN: 'Investments', ES: 'Inversiones', PT: 'Investimentos' },
      'COMMITMENTS': { FR: 'Engagements', EN: 'Liabilities', ES: 'Compromisos', PT: 'Compromissos' },
      'OTHERS': { FR: 'Autres', EN: 'Others', ES: 'Otros', PT: 'Outros' },
      'OF_TOTAL': { FR: '{pct}% du total', EN: '{pct}% of total', ES: '{pct}% del total', PT: '{pct}% do total' },
      'CONSUMED': { FR: 'Consommé :', EN: 'Spent:', ES: 'Consumido:', PT: 'Consumido:' },
      'ALLOCATED': { FR: 'Alloué :', EN: 'Allocated:', ES: 'Asignado:', PT: 'Alocado:' },
      'NO_BUDGET_PERIOD': { FR: 'Aucun budget défini pour cette période', EN: 'No budget defined for this period', ES: 'Sin presupuesto para este período', PT: 'Nenhum orçamento para este período' },
      'NO_TX': { FR: 'Aucune transaction enregistrée', EN: 'No transactions recorded', ES: 'Ninguna transacción registrada', PT: 'Nenhuma transação registrada' },
      'NO_ASSETS': { FR: 'Aucun actif enregistré', EN: 'No assets recorded', ES: 'Ningún activo registrado', PT: 'Nenhum ativo registrado' },
      'NO_ACCOUNT_MATCH': { FR: 'Aucun compte correspondant', EN: 'No matching account', ES: 'Ninguna cuenta coincidente', PT: 'Nenhuma conta correspondente' },
      'NO_DEPENSE_PERIOD': { FR: 'Aucune dépense pour cette période', EN: 'No expenses for this period', ES: 'Sin gastos para este período', PT: 'Sem despesas para este período' },
      'TOTAL': { FR: 'Total', EN: 'Total', ES: 'Total', PT: 'Total' },
      'NO_EXPENSE_PERIOD': { FR: 'Aucune dépense enregistrée pour {period}', EN: 'No expenses recorded for {period}', ES: 'Ningún gasto registrado para {period}', PT: 'Nenhuma despesa registrada para {period}' },
      'FILTER_ACTIVE': { FR: 'Filtre actif', EN: 'Active filter', ES: 'Filtro activo', PT: 'Filtro ativo' },
      'FILTER_CLEARED': { FR: 'Filtre de catégorie effacé', EN: 'Category filter cleared', ES: 'Filtro de categoría eliminado', PT: 'Filtro de categoria limpo' },
      'CLICK_TO_FILTER': { FR: 'Cliquer pour filtrer', EN: 'Click to filter', ES: 'Clic para filtrar', PT: 'Clique para filtrar' },

      // Saisie View & Modals
      'SAISIE_T': { FR: 'SAISIE TRANSACTION', EN: 'TRANSACTION ENTRY', ES: 'ENTRADA DE TRANSACCIÓN', PT: 'ENTRADA DE TRANSAÇÃO' },
      'SAISIE_D': {
        FR: 'Ajouter un revenu, dépense ou virement interne avec validation d\'intégrité.',
        EN: 'Add an income, expense or transfer with data integrity verification.',
        ES: 'Agregar un ingreso, gasto o transferencia con verificación de integridad.',
        PT: 'Adicionar uma receita, despesa ou transferência com validação de integridade.'
      },
      'SAISIE_EXPRESS': { FR: '+ Saisie Express', EN: '+ Express Entry', ES: '+ Entrada Rápida', PT: '+ Entrada Expressa' },
      'SAISIE_EXPRESS_TITLE': { FR: 'SAISIE EXPRESS TRANSACTION', EN: 'EXPRESS TRANSACTION ENTRY', ES: 'ENTRADA RÁPIDA DE TRANSACCIÓN', PT: 'ENTRADA EXPRESSA DE TRANSAÇÃO' },
      'SAISIE_EXPRESS_SUB': {
        FR: 'Enregistrement ultra-rapide de votre opération financière',
        EN: 'Fast financial operation record',
        ES: 'Registro ultrarrápido de operación financiera',
        PT: 'Registro ultrarrápido de operação financeira'
      },
      'DIRECT_LOCAL': { FR: 'Enregistrement direct local', EN: 'Direct local record', ES: 'Registro directo local', PT: 'Registro direto local' },
      'TYPE_REVENU': { FR: 'Revenu', EN: 'Income', ES: 'Ingreso', PT: 'Receita' },
      'TYPE_DEPENSE': { FR: 'Dépense', EN: 'Expense', ES: 'Gasto', PT: 'Despesa' },
      'TYPE_VIREMENT': { FR: 'Virement', EN: 'Transfer', ES: 'Transferencia', PT: 'Transferência' },
      'FIELD_MONTANT': { FR: 'Montant', EN: 'Amount', ES: 'Monto', PT: 'Valor' },
      'FIELD_DATE': { FR: 'Date de l\'opération', EN: 'Operation Date', ES: 'Fecha de operación', PT: 'Data da operação' },
      'FIELD_ACTIFS': { FR: 'Actifs', EN: 'Assets / Account', ES: 'Activos / Cuenta', PT: 'Ativos / Conta' },
      'FIELD_CATEGORIE': { FR: 'Catégorie Budgétaire', EN: 'Budget Category', ES: 'Categoría Presupuestaria', PT: 'Categoria Orçamentária' },
      'FIELD_TIERS': { FR: 'Tiers (Bénéficiaire / Émetteur)', EN: 'Party (Beneficiary / Payer)', ES: 'Tercero (Beneficiario / Emisor)', PT: 'Terceiro (Beneficiário / Emissor)' },
      'FIELD_NOTES': { FR: 'Description / Notes', EN: 'Description / Notes', ES: 'Descripción / Notas', PT: 'Descrição / Notas' },
      'BTN_VALIDATE_TX': { FR: 'VALIDER L\'OPÉRATION', EN: 'SUBMIT TRANSACTION', ES: 'VALIDAR OPERACIÓN', PT: 'CONFIRMAR OPERAÇÃO' },

      // Cashflow View
      'DASH_T': { FR: 'DASHBOARD CASHFLOW', EN: 'CASHFLOW DASHBOARD', ES: 'PANEL DE CASHFLOW', PT: 'PAINEL DE FLUXO DE CAIXA' },
      'DASH_D': {
        FR: 'Analyser les flux mensuels consolidés et le net de trésorerie.',
        EN: 'Analyze monthly consolidated flows and net treasury balance.',
        ES: 'Analizar flujos mensuales consolidados y tesorería neta.',
        PT: 'Analisar fluxos mensais consolidados e saldo de tesouraria líquido.'
      },
      'REVENUES': { FR: 'REVENUS', EN: 'INCOME', ES: 'INGRESOS', PT: 'RECEITAS' },
      'DEPENSES': { FR: 'DÉPENSES', EN: 'EXPENSES', ES: 'GASTOS', PT: 'DESPESAS' },
      'NET_CASHFLOW_SHORT': { FR: 'CASHFLOW NET', EN: 'NET CASHFLOW', ES: 'FLUJO NETO', PT: 'FLUXO LÍQUIDO' },
      'TREND_CHART_TITLE': { FR: 'TENDANCE MULTI-MOIS (Revenus vs Dépenses)', EN: 'MULTI-MONTH TREND (Income vs Expenses)', ES: 'TENDENCIA MULTIMES (Ingresos vs Gastos)', PT: 'TENDÊNCIA MULTIMÊS (Receitas vs Despesas)' },
      'TREND_6M': { FR: '6 Derniers Mois', EN: 'Last 6 Months', ES: 'Últimos 6 Meses', PT: 'Últimos 6 Meses' },
      'DONUT_TITLE': { FR: 'RÉPARTITION PAR CATÉGORIE', EN: 'BREAKDOWN BY CATEGORY', ES: 'DISTRIBUCIÓN POR CATEGORÍA', PT: 'DISTRIBUIÇÃO POR CATEGORIA' },
      'SEARCH_TX_PLACEHOLDER': { FR: 'Rechercher catégorie, tiers, compte ou note...', EN: 'Search category, party, account or note...', ES: 'Buscar categoría, tercero, cuenta o nota...', PT: 'Pesquisar categoria, terceiro, conta ou nota...' },
      'TX_CONSOLIDATED': { FR: 'TRANSACTIONS CONSOLIDÉES', EN: 'CONSOLIDATED TRANSACTIONS', ES: 'TRANSACCIONES CONSOLIDADAS', PT: 'TRANSAÇÕES CONSOLIDADAS' },
      'FLUX_COUNT': { FR: '{count} flux', EN: '{count} flows', ES: '{count} flujos', PT: '{count} fluxos' },
      'EMPTY_TX_FILTER': { FR: 'Aucune transaction trouvée pour cette période / filtre', EN: 'No transactions found for this period / filter', ES: 'No se encontraron transacciones para este período / filtro', PT: 'Nenhuma transação encontrada para este período / filtro' },
      'SORT_PREFIX': { FR: 'Trier :', EN: 'Sort:', ES: 'Ordenar:', PT: 'Ordenar:' },
      'ALL_ASSETS': { FR: 'Tous les comptes', EN: 'All accounts', ES: 'Todas las cuentas', PT: 'Todas as contas' },

      // Net Worth View
      'NETW_T': { FR: 'BILAN PATRIMONIAL', EN: 'NET WORTH & ASSETS', ES: 'BALANCE PATRIMONIAL', PT: 'BALANÇO PATRIMONIAL' },
      'NETW_D': {
        FR: 'Calcul de la Valeur Nette et photographie patrimoniale à date.',
        EN: 'Calculation of Net Worth and asset photography as of date.',
        ES: 'Cálculo del Valor Neto y fotografía patrimonial a la fecha.',
        PT: 'Cálculo do Patrimônio Líquido e fotografia patrimonial na data.'
      },
      'NETWORTH_TITLE': { FR: 'VALEUR NETTE PATRIMONIALE', EN: 'NET WORTH & WEALTH', ES: 'VALOR NETO PATRIMONIAL', PT: 'VALOR PATRIMONIAL LÍQUIDO' },
      'NETWORTH_INVENTORY': { FR: 'INVENTAIRE DES COMPTES ET DETTES', EN: 'ACCOUNTS & LIABILITIES INVENTORY', ES: 'INVENTARIO DE CUENTAS Y DEUDAS', PT: 'INVENTÁRIO DE CONTAS E DÍVIDAS' },
      'ASSET_ACTIVE': { FR: 'Actif', EN: 'Asset', ES: 'Activo', PT: 'Ativo' },
      'ASSET_LIABILITY': { FR: 'Passif / Dette', EN: 'Liability / Debt', ES: 'Pasivo / Deuda', PT: 'Passivo / Dívida' },

      // Budget View
      'BUDG_T': { FR: 'PILOTAGE BUDGÉTAIRE', EN: 'BUDGET MANAGEMENT', ES: 'GESTIÓN PRESUPUESTARIA', PT: 'GESTÃO ORÇAMENTÁRIA' },
      'BUDG_D': {
        FR: 'Suivi des enveloppes budgétaires avec jauges de consommation.',
        EN: 'Budget envelope tracking with consumption gauges.',
        ES: 'Seguimiento de sobres presupuestarios con medidores de consumo.',
        PT: 'Acompanhamento de envelopes orçamentários com barras de consumo.'
      },
      'BUDGET_TITLE': {
        FR: 'SUIVI DES ENVELOPPES BUDGÉTAIRES',
        EN: 'BUDGETARY ENVELOPES TRACKING',
        ES: 'SEGUIMIENTO DE SOBRES PRESUPUESTARIOS',
        PT: 'ACOMPANHAMENTO DE ENVELOPES ORÇAMENTÁRIOS'
      },
      'BUDGET_GLOBAL_ENVELOPE': { FR: 'Enveloppe Globale', EN: 'Global Envelope', ES: 'Presupuesto Global', PT: 'Envelope Global' },
      'BTN_ALLOC_BUDGET': { FR: '+ ALLOUER', EN: '+ ALLOCATE', ES: '+ ASIGNAR', PT: '+ ALOCAR' },
      'NO_BUDGET_MONTH': { FR: 'Aucune enveloppe allouée pour ce mois', EN: 'No budget envelope allocated for this month', ES: 'Ningún sobre asignado para este mes', PT: 'Nenhum envelope alocado para este mês' },

      // Table Headers
      'TH_DATE': { FR: 'Date', EN: 'Date', ES: 'Fecha', PT: 'Data' },
      'TH_OPERATION_ACTIF': { FR: 'Opération / Actif', EN: 'Operation / Asset', ES: 'Operación / Activo', PT: 'Operação / Ativo' },
      'TH_MONTANT': { FR: 'Montant', EN: 'Amount', ES: 'Monto', PT: 'Valor' },
      'TH_CATEGORIE_COMPTE': { FR: 'Catégorie / Compte', EN: 'Category / Account', ES: 'Categoría / Cuenta', PT: 'Categoria / Conta' },
      'TH_TIERS': { FR: 'Tiers', EN: 'Party', ES: 'Tercero', PT: 'Terceiro' },
      'TH_COMPTE': { FR: 'Compte', EN: 'Account', ES: 'Cuenta', PT: 'Conta' },
      'TH_CLASSE': { FR: 'Classe', EN: 'Class', ES: 'Clase', PT: 'Classe' },
      'TH_SOLDE_NATIF': { FR: 'Solde Natif', EN: 'Native Balance', ES: 'Saldo Nativo', PT: 'Saldo Nativo' },
      'TH_SOLDE_CONVERTI': { FR: 'Solde Converti', EN: 'Converted Balance', ES: 'Saldo Convertido', PT: 'Saldo Convertido' },
      'TH_CATEGORIE': { FR: 'Catégorie', EN: 'Category', ES: 'Categoría', PT: 'Categoria' },
      'TH_ALLOUE': { FR: 'Alloué', EN: 'Allocated', ES: 'Asignado', PT: 'Alocado' },
      'TH_CONSOMME': { FR: 'Réalisé', EN: 'Spent', ES: 'Consumido', PT: 'Consumido' },
      'TH_STATUT': { FR: 'Consommation', EN: 'Usage', ES: 'Consumo', PT: 'Consumo' },
      'TH_ECART': { FR: 'Écart', EN: 'Variance', ES: 'Diferencia', PT: 'Diferença' },

      // Periods & Dates
      'ANNÉE ENTIÈRE': { FR: 'ANNÉE ENTIÈRE', EN: 'FULL YEAR', ES: 'AÑO COMPLETO', PT: 'ANO INTEIRO' },
      'ALL_YEAR_SHORT': { FR: 'TOUT', EN: 'ALL', ES: 'TODO', PT: 'TUDO' },
      'AS_OF_MONTH': {
        FR: 'Situation en {month} {year}',
        EN: 'As of {month} {year}',
        ES: 'A fecha de {month} de {year}',
        PT: 'Posição em {month} de {year}'
      },
      'AS_OF_YEAR': {
        FR: 'Situation pour l\'année {year}',
        EN: 'Position for Year {year}',
        ES: 'Situación para el año {year}',
        PT: 'Posição para o ano {year}'
      },
      'AS_OF_DATE': {
        FR: 'Situation au {date}',
        EN: 'As of {date}',
        ES: 'A fecha de {date}',
        PT: 'Posição em {date}'
      },

      // Buttons & Actions
      'BTN_REPORT': { FR: 'Rapport', EN: 'Report', ES: 'Informe', PT: 'Relatório' },
      'BTN_BACKUP': { FR: 'Sauvegarder', EN: 'Backup', ES: 'Guardar', PT: 'Salvar' },
      'BTN_RESTORE': { FR: 'Restaurer', EN: 'Restore', ES: 'Restaurar', PT: 'Restaurar' },
      'FILTER_ALL': { FR: 'Tous', EN: 'All', ES: 'Todos', PT: 'Todos' },
      'FILTER_REV': { FR: 'Revenus', EN: 'Income', ES: 'Ingresos', PT: 'Receitas' },
      'FILTER_DEP': { FR: 'Dépenses', EN: 'Expenses', ES: 'Gastos', PT: 'Despesas' },
      'FILTER_TRA': { FR: 'Virements', EN: 'Transfers', ES: 'Transferencias', PT: 'Transferências' },
      'ACTIVE_PERIOD': { FR: 'Période Active', EN: 'Active Period', ES: 'Período Activo', PT: 'Período Ativo' },
      'SORT_DATE_DESC': { FR: 'Date (Récent)', EN: 'Date (Recent)', ES: 'Fecha (Reciente)', PT: 'Data (Recente)' },
      'SORT_DATE_ASC': { FR: 'Date (Ancien)', EN: 'Date (Oldest)', ES: 'Fecha (Antiguo)', PT: 'Data (Antiga)' },
      'SORT_AMT_DESC': { FR: 'Montant (Élevé)', EN: 'Amount (High)', ES: 'Monto (Alto)', PT: 'Valor (Alto)' },
      'SORT_AMT_ASC': { FR: 'Montant (Faible)', EN: 'Amount (Low)', ES: 'Monto (Bajo)', PT: 'Valor (Baixo)' },
      'CAT_FILTER_CLEAR': { FR: 'Réinitialiser filtre catégorie', EN: 'Reset category filter', ES: 'Reiniciar filtro', PT: 'Limpar filtro' },
      'STATUS_OVER': { FR: 'Dépassement', EN: 'Over Budget', ES: 'Superado', PT: 'Excedido' },
      'STATUS_WARN': { FR: 'Sous tension', EN: 'Warning', ES: 'Alerta', PT: 'Alerta' },
      'STATUS_OK': { FR: 'Sous contrôle', EN: 'Under Control', ES: 'Bajo Control', PT: 'Sob Controle' },

      // Executive Report Modal
      'REPORT_TITLE': { FR: 'RAPPORT FINANCIER EXÉCUTIF', EN: 'EXECUTIVE FINANCIAL REPORT', ES: 'INFORME FINANCIERO EJECUTIVO', PT: 'RELATÓRIO FINANCEIRO EXECUTIVO' },
      'REPORT_SUMMARY': { FR: 'Synthèse {period}', EN: 'Summary {period}', ES: 'Resumen {period}', PT: 'Síntese {period}' },
      'REPORT_CONSOL_CURR': { FR: 'Devise de consolidation : {curr}', EN: 'Consolidation currency: {curr}', ES: 'Moneda de consolidación: {curr}', PT: 'Moeda de consolidação: {curr}' },
      'REPORT_TOT_REV': { FR: 'Revenus Totaux', EN: 'Total Income', ES: 'Ingresos Totales', PT: 'Receitas Totais' },
      'REPORT_TOT_DEP': { FR: 'Dépenses Totales', EN: 'Total Expenses', ES: 'Gastos Totales', PT: 'Despesas Totais' },
      'REPORT_NET_CF': { FR: 'Cashflow Net', EN: 'Net Cashflow', ES: 'Flujo Neto', PT: 'Fluxo Líquido' },
      'REPORT_SAVINGS_RATE': { FR: 'Taux d\'Épargne', EN: 'Savings Rate', ES: 'Tasa de Ahorro', PT: 'Taxa de Poupança' },
      'REPORT_NET_WORTH': { FR: 'Patrimoine Net', EN: 'Net Worth', ES: 'Patrimonio Neto', PT: 'Patrimônio Líquido' },
      'REPORT_ASSETS_DEBTS': { FR: 'Actifs : {assets} • Dettes : {debts}', EN: 'Assets: {assets} • Liabilities: {debts}', ES: 'Activos: {assets} • Deudas: {debts}', PT: 'Ativos: {assets} • Dívidas: {debts}' },
      'REPORT_TOP_DEP': { FR: 'Top 5 Postes de Dépenses :', EN: 'Top 5 Expense Categories:', ES: 'Top 5 Categorías de Gastos:', PT: 'Top 5 Categorias de Despesas:' },
      'BTN_REPORT_COPY': { FR: 'Copier', EN: 'Copy', ES: 'Copiar', PT: 'Copiar' },
      'BTN_CLOSE': { FR: 'Fermer', EN: 'Close', ES: 'Cerrar', PT: 'Fechar' },
      'TOAST_REPORT_COPIED': { FR: 'Rapport copié dans le presse-papier !', EN: 'Report copied to clipboard!', ES: '¡Informe copiado al portapapeles!', PT: 'Relatório copiado para a área de transferência!' },
      'TOAST_REPORT_GENERATED': { FR: 'Rapport généré avec succès', EN: 'Report generated successfully', ES: 'Informe generado con éxito', PT: 'Relatório gerado com sucesso' },

      // Database management buttons
      'DB_ADMIN_TITLE': { FR: 'ADMINISTRATION DE LA BASE DE DONNÉES', EN: 'DATABASE ADMINISTRATION', ES: 'ADMINISTRACIÓN DE LA BASE DE DATOS', PT: 'ADMINISTRAÇÃO DO BANCO DE DADOS' },
      'DB_INIT': { FR: 'Réinitialiser', EN: 'Reset', ES: 'Reiniciar', PT: 'Reiniciar' },
      'DB_REFRESH': { FR: 'Actualiser', EN: 'Refresh', ES: 'Actualizar', PT: 'Atualizar' },
      'DB_BACKUP': { FR: 'Sauvegarder', EN: 'Backup', ES: 'Guardar', PT: 'Salvar' },
      'DB_RESTORE': { FR: 'Restaurer', EN: 'Restore', ES: 'Restaurar', PT: 'Restaurar' },
      'DB_PURGE': { FR: 'Purger les Données', EN: 'Purge Data', ES: 'Purgar Datos', PT: 'Purgar Dados' },
      'DB_AUDIT': { FR: 'Journal d\'Audit', EN: 'Audit Log', ES: 'Registro de Auditoría', PT: 'Diário de Auditoria' },

      // Modals Questions
      'Q1_TITLE': { FR: 'Quel est le type et le montant ?', EN: 'What is the type and amount?', ES: '¿Cuál es el tipo y monto?', PT: 'Qual é o tipo e valor?' },
      'Q2_TITLE': { FR: 'Quel actif utiliser ?', EN: 'Which asset to use?', ES: '¿Qué activo utilizar?', PT: 'Qual ativo utilizar?' },
      'Q3_TITLE': { FR: 'Quelle est la catégorie ?', EN: 'What is the category?', ES: '¿Cuál es la categoría?', PT: 'Qual é a categoria?' },
      'Q4_TITLE': { FR: 'Qui est le tiers concerné ?', EN: 'Who is the party involved?', ES: '¿Quién es el tercero involucrado?', PT: 'Quem é o terceiro envolvido?' },
      'Q5_TITLE': { FR: 'Date & Justificatif', EN: 'Date & Description', ES: 'Fecha y Justificante', PT: 'Data e Justificativa' },
      'BTN_SAVE_EXPRESS': { FR: 'ENREGISTRER LA TRANSACTION', EN: 'SAVE TRANSACTION', ES: 'GUARDAR TRANSACCIÓN', PT: 'SALVAR TRANSAÇÃO' },
      'ALLOC_MODAL_TITLE': { FR: 'ALLOUER UNE ENVELOPPE BUDGÉTAIRE', EN: 'ALLOCATE BUDGET ENVELOPE', ES: 'ASIGNAR SOBRE PRESUPUESTARIO', PT: 'ALOCAR ENVELOPE ORÇAMENTÁRIO' },
      'Q_ALLOC_1': { FR: 'Pour quel mois budgétaire ?', EN: 'For which budget month?', ES: '¿Para qué mes presupuestario?', PT: 'Para qual mês orçamentário?' },
      'Q_ALLOC_2': { FR: 'Quelle catégorie de dépense allouer ?', EN: 'Which expense category to allocate?', ES: '¿Qué categoría de gasto asignar?', PT: 'Qual categoria de despesa alocar?' },
      'Q_ALLOC_3': { FR: 'Quel montant allouer à cette enveloppe ?', EN: 'What amount to allocate?', ES: '¿Qué monto asignar?', PT: 'Qual valor alocar?' },
      'BTN_SAVE_ALLOC': { FR: 'VALIDER L\'ALLOCATION', EN: 'CONFIRM ALLOCATION', ES: 'VALIDAR ASIGNACIÓN', PT: 'CONFIRMAR ALOCAÇÃO' },
      'AUDIT_TITLE': { FR: 'JOURNAL D\'AUDIT & HISTORIQUE', EN: 'AUDIT LOG & HISTORY', ES: 'REGISTRO DE AUDITORÍA E HISTORIAL', PT: 'DIÁRIO DE AUDITORIA E HISTÓRICO' },
            'MODAL_ADD_ENTITY_TITLE': { FR: 'NOUVEL ÉLÉMENT', EN: 'NEW ELEMENT', ES: 'NUEVO ELEMENTO', PT: 'NOVO ELEMENTO' },
      'PILL_ENTITY_COMPTE': { FR: 'Actif / Compte', EN: 'Asset / Account', ES: 'Activo / Cuenta', PT: 'Ativo / Conta' },
      'PILL_ENTITY_CATEGORIE': { FR: 'Catégorie', EN: 'Category', ES: 'Categoría', PT: 'Categoria' },
      'PILL_ENTITY_TIERS': { FR: 'Tiers', EN: 'Third Party', ES: 'Tercero', PT: 'Terceiro' },
      'LBL_ADD_ENTITY_NAME': { FR: 'Désignation / Nom *', EN: 'Name / Designation *', ES: 'Nombre / Designación *', PT: 'Nome / Designação *' },
      'LBL_ADD_ENTITY_CLASSE': { FR: 'Classe d\'Actif', EN: 'Asset Class', ES: 'Clase de Activo', PT: 'Classe de Ativo' },
      'LBL_ADD_ENTITY_DEVISE': { FR: 'Devise du Compte', EN: 'Account Currency', ES: 'Moneda de la Cuenta', PT: 'Moeda da Conta' },
      'LBL_ADD_ENTITY_CAT_TYPE': { FR: 'Type Budgétaire', EN: 'Budget Type', ES: 'Tipo de Presupuesto', PT: 'Tipo de Orçamento' },
      'LBL_ADD_ENTITY_TIERS_SECTEUR': { FR: 'Secteur / Typologie', EN: 'Sector / Type', ES: 'Sector / Tipología', PT: 'Setor / Tipologia' },
      'PLACEHOLDER_NEW_COMPTE': {
        FR: 'Ex: Livret A, Compte Titres, Wallet Crypto...',
        EN: 'Ex: Savings Account, Brokerage, Crypto Wallet...',
        ES: 'Ej: Cuenta de Ahorro, Cartera de Inversión...',
        PT: 'Ex: Conta Poupança, Carteira de Investimentos...'
      },
      'PLACEHOLDER_NEW_CAT': {
        FR: 'Ex: Abonnements Streaming, Formations Pro...',
        EN: 'Ex: Streaming Subscriptions, Pro Courses...',
        ES: 'Ej: Suscripciones, Cursos Profesionales...',
        PT: 'Ex: Assinaturas de Streaming, Cursos Pro...'
      },
      'PLACEHOLDER_NEW_TIERS': {
        FR: 'Ex: Netflix, Super U, Propriétaire, Employeur...',
        EN: 'Ex: Netflix, Supermarket, Landlord, Employer...',
        ES: 'Ej: Netflix, Supermercado, Propietario, Empleador...',
        PT: 'Ex: Netflix, Supermercado, Senhorio, Empregador...'
      },
      'OPT_CLASSE_LIQUIDITE': {
        FR: 'Liquidité (Banque, Espèces)',
        EN: 'Liquidity (Bank, Cash)',
        ES: 'Liquidez (Banco, Efectivo)',
        PT: 'Liquidez (Banco, Dinheiro)'
      },
      'OPT_CLASSE_INVESTISSEMENT': {
        FR: 'Investissement (Actions, Crypto, Immo)',
        EN: 'Investment (Stocks, Crypto, Real Estate)',
        ES: 'Inversión (Acciones, Cripto, Inmuebles)',
        PT: 'Investimento (Ações, Cripto, Imóveis)'
      },
      'OPT_CLASSE_DETTE': {
        FR: 'Dette & Emprunt (Prêt, Différé)',
        EN: 'Debt & Liability (Loan, Deferred)',
        ES: 'Deuda & Pasivo (Préstamo, Diferido)',
        PT: 'Dívida & Passivo (Empréstimo, Diferido)'
      },
      'OPT_TYPE_DEPENSE': {
        FR: 'Dépense (Consommation)',
        EN: 'Expense (Consumption)',
        ES: 'Gasto (Consumo)',
        PT: 'Despesa (Consumo)'
      },
      'OPT_TYPE_REVENU': {
        FR: 'Revenu (Entrée)',
        EN: 'Income (Revenue)',
        ES: 'Ingreso (Entrada)',
        PT: 'Receita (Entrada)'
      },
      'OPT_TYPE_TRANSFERT': {
        FR: 'Transfert / Épargne',
        EN: 'Transfer / Savings',
        ES: 'Transferencia / Ahorro',
        PT: 'Transferência / Poupança'
      },
      'OPT_SECTEUR_COMMERCE': {
        FR: 'Commerce / Supermarché',
        EN: 'Retail / Supermarket',
        ES: 'Comercio / Supermercado',
        PT: 'Comércio / Supermercado'
      },
      'OPT_SECTEUR_ENTREPRISE': {
        FR: 'Entreprise / Employeur',
        EN: 'Corporate / Employer',
        ES: 'Empresa / Empleador',
        PT: 'Empresa / Empregador'
      },
      'OPT_SECTEUR_FINANCE': {
        FR: 'Banque & Finance',
        EN: 'Banking & Finance',
        ES: 'Banca & Finanzas',
        PT: 'Banco & Finanças'
      },
      'OPT_SECTEUR_INSTITUTION': {
        FR: 'Institution Publique',
        EN: 'Public Institution',
        ES: 'Institución Pública',
        PT: 'Instituição Pública'
      },
      'OPT_SECTEUR_IMMOBILIER': {
        FR: 'Immobilier / Propriétaire',
        EN: 'Real Estate / Landlord',
        ES: 'Inmobiliaria / Propietario',
        PT: 'Imobiliário / Proprietário'
      },
      'OPT_SECTEUR_SANTE': {
        FR: 'Santé & Pharmacie',
        EN: 'Healthcare & Pharmacy',
        ES: 'Salud & Farmacia',
        PT: 'Saúde & Farmácia'
      },
      'OPT_SECTEUR_AUTRE': {
        FR: 'Autre Secteur',
        EN: 'Other Sector',
        ES: 'Otro Sector',
        PT: 'Outro Setor'
      },
      'BTN_CONFIRM_ADD_ENTITY': {
        FR: 'Ajouter & Sélectionner',
        EN: 'Add & Select',
        ES: 'Agregar y Seleccionar',
        PT: 'Adicionar e Selecionar'
      },
      'BTN_CANCEL': {
        FR: 'Annuler',
        EN: 'Cancel',
        ES: 'Cancelar',
        PT: 'Cancelar'
      },
      'ERR_NAME_REQUIRED': {
        FR: 'Veuillez renseigner un nom valide.',
        EN: 'Please enter a valid name.',
        ES: 'Por favor, ingrese un nombre válido.',
        PT: 'Por favor, insira um nome válido.'
      },
      'TITLE_NEW_COMPTE': {
        FR: 'Créer un nouvel actif',
        EN: 'Create a new asset',
        ES: 'Crear un nuevo activo',
        PT: 'Criar um novo ativo'
      },
      'TITLE_NEW_CAT': {
        FR: 'Créer une nouvelle catégorie',
        EN: 'Create a new category',
        ES: 'Crear una nueva categoría',
        PT: 'Criar uma nova categoria'
      },
      'TITLE_NEW_TIERS': {
        FR: 'Créer un nouveau tiers',
        EN: 'Create a new third party',
        ES: 'Crear un nuevo tercero',
        PT: 'Criar um novo terceiro'
      },
      'AJOUTER_NOUVEAU': { FR: 'Ajouter nouveau', EN: 'Add new', ES: 'Agregar nuevo', PT: 'Adicionar novo' },
      'BTN_FIELD_ADD': { FR: '+ Nouveau', EN: '+ New', ES: '+ Nuevo', PT: '+ Novo' },
      'ENTITY_ADDED_SUCCESS': { FR: 'Élément ajouté et sélectionné avec succès !', EN: 'Item successfully added and selected!', ES: '¡Elemento agregado y seleccionado con éxito!', PT: 'Item adicionado e selecionado com sucesso!' },
      'ENTITY_ALREADY_EXISTS': { FR: 'Cet élément existe déjà et a été sélectionné.', EN: 'This item already exists and was selected.', ES: 'Este elemento ya existe y ha sido seleccionado.', PT: 'Este item já existe e foi selecionado.' },
      'ALL_ACCOUNTS': { FR: 'Tous les actifs', EN: 'All assets', ES: 'Todos los activos', PT: 'Todos os ativos' },
      'FILTER_ALL_ASSETS': { FR: 'Tous les actifs', EN: 'All assets', ES: 'Todos los activos', PT: 'Todos os ativos' },
      'OPT_ADD_NEW': { FR: 'Ajouter nouveau...', EN: 'Add new...', ES: 'Agregar nuevo...', PT: 'Adicionar novo...' },
      'OPERATION_DEFAULT': { FR: 'Opération', EN: 'Transaction', ES: 'Operación', PT: 'Operação' },
      'ASSETS_DEFAULT': { FR: 'Actifs', EN: 'Assets', ES: 'Activos', PT: 'Ativos' },
      'BTN_CLOSE_AUDIT': { FR: 'FERMER LE JOURNAL', EN: 'CLOSE LOG', ES: 'CERRAR REGISTRO', PT: 'FECHAR DIÁRIO' },
      'BTN_REFRESH_RATES_TITLE': { FR: 'Actualiser les taux de change (Marchés connectés)', EN: 'Refresh live exchange rates (Connected markets)', ES: 'Actualizar tipos de cambio (Mercados conectados)', PT: 'Atualizar taxas de câmbio (Mercados conectados)' },
      'RATES_UPDATING': { FR: 'Actualisation des cours en direct...', EN: 'Updating live exchange rates...', ES: 'Actualizando tipos de cambio en directo...', PT: 'Atualizando taxas de câmbio em tempo real...' },
      'RATES_UPDATED_SUCCESS': { FR: 'Taux de change actualisés en direct !', EN: 'Exchange rates updated live!', ES: '¡Tipos de cambio actualizados en directo!', PT: 'Taxas de câmbio atualizadas em tempo real!' },
      'RATES_OFFLINE': { FR: 'Mode Hors-Ligne : Derniers taux sauvegardés', EN: 'Offline Mode: Using cached rates', ES: 'Modo Fuera de Línea: Últimos tipos guardados', PT: 'Modo Offline: Últimas taxas salvas' },
      'RATES_LAST_SYNC': { FR: 'Dernière sync', EN: 'Last sync', ES: 'Última sinc', PT: 'Última sinc' }
    };

    let template = '';
    if (fallbackDict[cleanKey]) {
      template = fallbackDict[cleanKey][targetLang] || fallbackDict[cleanKey]['FR'];
    } else if (this.data && this.data.dictionary && this.data.dictionary[cleanKey]) {
      const entry = this.data.dictionary[cleanKey];
      template = entry[targetLang] || entry['FR'] || cleanKey;
    } else {
      template = cleanKey;
    }

    // Interpolate params (e.g. {month}, {year}, {date})
    Object.keys(params).forEach(p => {
      template = template.replace(new RegExp(`\\{${p}\\}`, 'g'), params[p]);
    });

    return template;
  }

  setLanguage(langCode) {
    const code = (typeof langCode === 'string') ? langCode.trim().toUpperCase() : 'FR';
    const valid = ['FR', 'EN', 'ES', 'PT'];
    if (valid.includes(code)) {
      this.lang = code;
      if (this.data && this.data.config) this.data.config.LANGUE_UI = code;
      this.save();
      this.notify();
      return true;
    }
    return false;
  }

  // --- REAL-TIME FOREX EXCHANGE RATE API ENGINE ---
  async fetchLiveExchangeRates(showNotification = true) {
    if (typeof navigator !== 'undefined' && !navigator.onLine) {
      this.rateStatus = {
        isLive: false,
        loading: false,
        lastSync: this.lastExchangeSync || null,
        error: 'Mode Hors-Ligne (Taux sauvegardés)'
      };
      this.notify();
      return { success: false, offline: true, rates: this.data ? this.data.devises : {} };
    }

    if (!this.rateStatus) this.rateStatus = {};
    this.rateStatus.loading = true;
    this.notify();

    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 6500);

      const res = await fetch('https://open.er-api.com/v6/latest/EUR', {
        signal: controller.signal,
        cache: 'no-cache'
      });
      clearTimeout(timeoutId);

      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const json = await res.json();

      if (json && json.rates) {
        const supported = ['MUR', 'EUR', 'USD', 'GBP', 'ZAR', 'XOF'];
        const newDevises = {};
        const eurBaseXof = (this.data && this.data.devises && this.data.devises.EUR > 10)
          ? this.data.devises.EUR
          : 656.168;

        supported.forEach(curr => {
          if (curr === 'EUR') {
            newDevises[curr] = eurBaseXof;
          } else if (curr === 'XOF') {
            newDevises[curr] = 1;
          } else if (json.rates[curr] && json.rates[curr] > 0) {
            newDevises[curr] = Number((eurBaseXof / json.rates[curr]).toFixed(4));
          }
        });

        if (!newDevises['XOF']) newDevises['XOF'] = 1;

        if (!this.data) this.data = {};
        this.data.devises = Object.assign({}, this.data.devises, newDevises);
        if (this.data.config) {
          this.data.config.TAUX_CHANGE = { ...this.data.devises };
        }
        this.lastExchangeSync = new Date().toISOString();
        this.rateStatus = {
          isLive: true,
          loading: false,
          lastSync: this.lastExchangeSync,
          error: null
        };

        const storage = this.getStorage();
        if (storage) {
          storage.setItem('gfmp_live_rates', JSON.stringify({
            devises: this.data.devises,
            lastSync: this.lastExchangeSync
          }));
        }

        this.logHistory('SYNC_RATES', 'FOREX_API', 'LIVE', `Taux de change actualisés en direct via API (MUR, EUR, USD, GBP, ZAR, XOF)`);
        this.save();
        this.notify();
        return { success: true, rates: this.data.devises, lastSync: this.lastExchangeSync };
      }
      throw new Error('Réponse API invalide');
    } catch (err) {
      console.warn('[GFMPEngine] Live Forex API fetch error:', err);
      this.rateStatus = {
        isLive: false,
        loading: false,
        lastSync: this.lastExchangeSync || null,
        error: err.name === 'AbortError' ? 'Délai d\'attente dépassé' : err.message
      };
      this.notify();
      return { success: false, error: err.message, rates: this.data ? this.data.devises : {} };
    }
  }

  // --- APP INITIALIZATION & RECONFIGURATION ---
  initializeSystem(defaultCurrency = 'MUR', defaultLanguage = 'FR', defaultDate = null, defaultTheme = null, defaultYear = null, defaultMonth = null, useDemoData = true) {
    const sanitizedCurrency = (typeof defaultCurrency === 'string') ? defaultCurrency.trim().toUpperCase() : 'MUR';
    const sanitizedLanguage = (typeof defaultLanguage === 'string') ? defaultLanguage.trim().toUpperCase() : 'FR';
    const validThemes = ['blue', 'gold', 'emerald', 'amethyst'];
    const sanitizedTheme = (typeof defaultTheme === 'string' && validThemes.includes(defaultTheme.trim().toLowerCase()))
      ? defaultTheme.trim().toLowerCase()
      : (this.theme || 'blue');
    const isDemo = !(useDemoData === false || String(useDemoData).trim().toUpperCase() === 'EMPTY');
    const resolvedMode = isDemo ? 'DEMO' : 'EMPTY';

    this.lang = sanitizedLanguage;
    this.currency = sanitizedCurrency;
    this.theme = sanitizedTheme;
    this.dataMode = resolvedMode;
    this.selectedCategoryFilter = null;

    if (defaultDate && /^\d{4}-\d{2}-\d{2}$/.test(String(defaultDate).trim())) {
      this.referenceDate = String(defaultDate).trim();
      const [y, m] = this.referenceDate.split('-');
      this.year = defaultYear ? String(defaultYear) : y;
      this.month = defaultMonth ? String(defaultMonth).padStart(2, '0') : m;
    } else {
      if (defaultYear) this.year = String(defaultYear);
      if (defaultMonth) this.month = String(defaultMonth).padStart(2, '0');
    }

    // Reset dataset to baseline with chosen defaults
    if (typeof window !== 'undefined' && window.GFMP_BASELINE) {
      this.data = JSON.parse(JSON.stringify(window.GFMP_BASELINE));
    }

    if (this.data && !this.data.devises) {
      this.data.devises = { MUR: 12.2243, EUR: 656.168, USD: 567.5369, GBP: 757.0023, ZAR: 33.2723, XOF: 1 };
    }

    if (this.data) {
      this.data.dataMode = resolvedMode;
      if (!isDemo) {
        this.data.transactions = [];
        this.data.budgets = [];
      }
      if (!this.data.config) this.data.config = {};
      this.data.config.SYS_DEVISE_BASE = sanitizedCurrency;
      this.data.config.LANGUE_UI = sanitizedLanguage;
      this.data.config.SYS_DATE_REF = this.referenceDate;
      this.data.config.SYS_YEAR = this.year;
      this.data.config.SYS_MONTH = this.month;
      this.data.config.SYS_THEME = sanitizedTheme;
      this.data.config.SYS_DATA_MODE = resolvedMode;
    }

    const storage = this.getStorage();
    if (storage) {
      storage.setItem('gfmp_initialized', 'true');
      storage.setItem('gfmp_onboarding_v5_6', 'true');
      storage.setItem('gfmp_theme', sanitizedTheme);
      storage.setItem('gfmp_data_mode', resolvedMode);
    }
    this.needsOnboarding = false;

    // Trigger immediate background rate sync
    this.fetchLiveExchangeRates(false);

    this.logHistory('INIT_SYSTEM', 'CONFIG', sanitizedCurrency, `Initialisation du système : Devise ${sanitizedCurrency}, Langue ${sanitizedLanguage}, Période ${this.month}/${this.year}, Thème ${sanitizedTheme}, Données ${resolvedMode}`);
    this.save();
    this.notify();
    return {
      success: true,
      currency: sanitizedCurrency,
      lang: sanitizedLanguage,
      date: this.referenceDate,
      year: this.year,
      month: this.month,
      theme: sanitizedTheme,
      dataMode: resolvedMode
    };
  }

  getExchangeRateStatus() {
    return this.rateStatus || { isLive: false, loading: false, lastSync: null };
  }

  // --- MULTI-CURRENCY O(1) ENGINE ---
  setCurrency(currCode) {
    const code = (typeof currCode === 'string') ? currCode.trim().toUpperCase() : 'MUR';
    const supported = ['MUR', 'EUR', 'USD', 'GBP', 'ZAR', 'XOF'];
    if (this.data) {
      if (!this.data.devises) {
        this.data.devises = { MUR: 12.2243, EUR: 656.168, USD: 567.5369, GBP: 757.0023, ZAR: 33.2723, XOF: 1 };
      }
      if (this.data.devises[code] !== undefined || supported.includes(code)) {
        this.currency = code;
        if (this.data.config) this.data.config.SYS_DEVISE_BASE = code;
        this.save();
        this.notify();
        return true;
      }
    }
    return false;
  }

  convert(amount, fromCurrency, toCurrency = this.currency) {
    if (amount === 0 || isNaN(amount)) return 0;
    if (fromCurrency === toCurrency) return this.roundMoney(amount);
    const rates = (this.data && this.data.devises) ? this.data.devises : { MUR: 12.2243, EUR: 656.168, USD: 567.5369, GBP: 757.0023, ZAR: 33.2723, XOF: 1.0 };
    const tauxO = (rates[fromCurrency] !== undefined && rates[fromCurrency] > 0) ? rates[fromCurrency] : 1;
    const tauxC = (rates[toCurrency] !== undefined && rates[toCurrency] > 0) ? rates[toCurrency] : 1;
    return this.roundMoney((amount * tauxO) / tauxC);
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
    const safeAmount = this.roundMoney(amount || 0);
    const formatted = new Intl.NumberFormat('fr-FR', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    }).format(safeAmount);

    return `${formatted} ${sym}`;
  }

  // --- TIME SLIDER & MONTH STRIP ENGINE ---
  setYear(newYear) {
    this.year = String(newYear);
    this.notify();
  }

  setMonth(newMonth) {
    this.month = String(newMonth).padStart(2, '0');
    this.notify();
  }

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

  getMonthLabel(targetMonth = this.month) {
    if (targetMonth === '00') return this.tr('ANNÉE ENTIÈRE');
    const monthIndex = parseInt(targetMonth, 10) - 1;
    const months = {
      FR: ['Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre'],
      EN: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
      ES: ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'],
      PT: ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
    };
    const langList = months[this.lang] || months['FR'];
    return (langList[monthIndex] || targetMonth).toUpperCase();
  }

  // --- 1. MODULE CASHFLOW & INTERACTIVE CHARTS ---
  computeCashflow(period = null) {
    if (!this.data) return { totalRevenus: 0, totalDepenses: 0, netCashflow: 0, savingsRate: 0, transactions: [], breakdown: [] };

    const { year: targetYear, month: targetMonth } = this.resolvePeriod(period);

    const catDict = {};
    const catTypeDict = {};
    (this.data.categories || []).forEach(c => {
      catDict[c.id] = c.nom;
      catTypeDict[c.id] = c.type;
    });

    const cptDict = {};
    (this.data.comptes || []).forEach(c => { cptDict[c.id] = this.cleanAssetName(c.nom); });

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

      const matchYear = (txYear === targetYear);
      const matchMonth = (targetMonth === '00' || txMonth === targetMonth);

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

        // Multi-level Filters & Translation
        const nomCompte = cptDict[tx.id_compte] || 'Compte';
        const nomCat = catDict[tx.id_categorie] || 'Catégorie';
        const nomTiers = tiersDict[tx.id_tiers] || 'Tiers';
        const nomCompteTraduit = this.trEntity(nomCompte);
        const nomCategorieTraduit = this.trEntity(nomCat);
        const nomTiersTraduit = this.trEntity(nomTiers);
        const notesTraduites = this.trEntity(tx.notes);
        const fullSearchStr = `${tx.date} ${nomCompte} ${nomCat} ${nomTiers} ${tx.notes || ''} ${nomCompteTraduit} ${nomCategorieTraduit} ${nomTiersTraduit} ${notesTraduites}`.toLowerCase();

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
        }
      }
    });

    // Sorting
    filteredTx.sort((a, b) => {
      if (this.sortField === 'montant') {
        return (this.sortDir === 'asc') ? a.montantConverti - b.montantConverti : b.montantConverti - a.montantConverti;
      }
      // default: date
      return (this.sortDir === 'asc') ? a.date.localeCompare(b.date) : b.date.localeCompare(a.date);
    });

    totalRevenus = this.roundMoney(totalRevenus);
    totalDepenses = this.roundMoney(totalDepenses);
    const netCashflow = this.roundMoney(totalRevenus - totalDepenses);

    // Category expense breakdown with deterministic Haute Joaillerie SFP v1.0 Palette
    const breakdown = Object.entries(catBreakdown).map(([nom, rawMontant]) => {
      const montant = this.roundMoney(rawMontant);
      const pct = totalDepenses > 0 ? this.roundMoney((montant / totalDepenses) * 100) : 0;
      const color = this.getCategoryColor(nom);
      return {
        nom,
        nomTraduit: this.trEntity(nom),
        montant,
        pct,
        color
      };
    }).sort((a, b) => b.montant - a.montant);

    const savingsRate = totalRevenus > 0 ? this.roundMoney(Math.max(0, ((totalRevenus - totalDepenses) / totalRevenus) * 100)) : 0;

    return {
      totalRevenus,
      totalDepenses,
      netCashflow,
      savingsRate,
      transactions: filteredTx,
      breakdown
    };
  }

  // --- 2. MULTI-MONTH CASHFLOW TREND (INTERACTIVE) ---
  computeCashflowTrend(monthsCount = 6, period = null) {
    if (!this.data) return [];
    const trend = [];
    const { year: targetYear, month: targetMonth } = this.resolvePeriod(period);
    const currentM = parseInt(targetMonth === '00' ? '12' : targetMonth, 10);
    const currentY = parseInt(targetYear, 10);

    const catTypeDict = {};
    (this.data.categories || []).forEach(c => { catTypeDict[c.id] = c.type; });

    const monthNamesShort = {
      FR: ['Jan', 'Fév', 'Mar', 'Avr', 'Mai', 'Juin', 'Juil', 'Août', 'Sep', 'Oct', 'Nov', 'Déc'],
      EN: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
      ES: ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'],
      PT: ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
    };
    const langShort = monthNamesShort[this.lang] || monthNamesShort['FR'];

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

      const roundedRev = this.roundMoney(rev);
      const roundedDep = this.roundMoney(dep);

      trend.push({
        year: String(y),
        month: mStr,
        yearMonth: ymStr,
        label: `${langShort[m - 1]} ${String(y).slice(-2)}`,
        revenus: roundedRev,
        depenses: roundedDep,
        net: this.roundMoney(roundedRev - roundedDep)
      });
    }

    return trend;
  }

  // --- 3. MODULE BILAN PATRIMONIAL / NET WORTH ---
  computeNetWorth(period = null) {
    if (!this.data) return { totalAssets: 0, totalLiabilities: 0, netWorth: 0, accounts: [], allocation: {} };

    const { year: targetYear, month: targetMonth } = this.resolvePeriod(period);
    const targetFilterMonth = (targetMonth === '00') ? `${targetYear}-12` : `${targetYear}-${targetMonth}`;
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
        if (!tx.date || tx.date.substring(0, 7) > targetFilterMonth) return;
        const fluxType = catTypeDict[tx.id_categorie] || 'DEPENSE';
        const txDev = tx.devise || cptNativeDevise;
        const montantNormalise = this.convert(tx.montant, txDev, cptNativeDevise);

        if (String(tx.id_compte) === String(cpt.id)) {
          if (fluxType === 'REVENU') {
            soldeNatif += montantNormalise;
          } else if (fluxType === 'DEPENSE') {
            soldeNatif -= montantNormalise;
          } else if (fluxType === 'TRANSFERT') {
            // Transfers into investments (DCA/Crypto/PEA) increase balance
            if (cpt.classe === 'INVESTISSEMENT') {
              soldeNatif += montantNormalise;
            } else {
              // Outflow from checking/liquid account to savings/investments
              soldeNatif -= montantNormalise;
            }
          }
        } else if (String(cpt.id) === '2' && String(tx.id_compte) === '1' && fluxType === 'TRANSFERT') {
          // Cross-account counterpart: Transfers from Checking (Account 1) to Savings (Account 2 Livret d'Épargne)
          const notes = (tx.notes || '').toLowerCase();
          if (notes.includes('épargne') || notes.includes('epargne') || notes.includes('livret') || notes.includes('precaution') || notes.includes('précaution')) {
            soldeNatif += montantNormalise;
          }
        }
      });

      soldeNatif = this.roundMoney(soldeNatif);

      // Exclude empty template trigger placeholder ("Autre (Préciser...)") from Net Worth inventory
      const isPlaceholder = (cpt.classe === 'AUTRE' && (cpt.nom || '').toLowerCase().includes('préciser') && Math.abs(soldeNatif) < 0.005);
      if (isPlaceholder) return;

      const soldeConverti = this.roundMoney(this.convert(soldeNatif, cptNativeDevise, this.currency));
      const isDette = (cpt.classe === 'DETTE');

      if (isDette) {
        const debtAmt = Math.abs(soldeConverti);
        totalLiabilities += debtAmt;
        allocation.DETTE = (allocation.DETTE || 0) + debtAmt;
      } else {
        if (cpt.classe === 'INVESTISSEMENT') {
          const invAmt = Math.max(0, soldeConverti);
          totalAssets += invAmt;
          allocation.INVESTISSEMENT = (allocation.INVESTISSEMENT || 0) + invAmt;
        } else {
          const liqAmt = Math.max(0, soldeConverti);
          totalAssets += liqAmt;
          allocation.LIQUIDITE = (allocation.LIQUIDITE || 0) + liqAmt;
        }
      }

      // Filter by class if applicable
      const matchClass = (this.patrimoineClassFilter === 'ALL') || (cpt.classe === this.patrimoineClassFilter);
      if (matchClass) {
        accountResults.push({
          id: cpt.id,
          nom: this.cleanAssetName(cpt.nom),
          nomTraduit: this.trEntity(cpt.nom),
          classe: cpt.classe || 'LIQUIDITE',
          classeTraduit: this.trEntity(cpt.classe || 'LIQUIDITE'),
          deviseNative: cptNativeDevise,
          soldeNatif,
          soldeConverti,
          isDette
        });
      }
    });

    totalAssets = this.roundMoney(totalAssets);
    totalLiabilities = this.roundMoney(totalLiabilities);
    allocation.LIQUIDITE = this.roundMoney(allocation.LIQUIDITE || 0);
    allocation.INVESTISSEMENT = this.roundMoney(allocation.INVESTISSEMENT || 0);
    allocation.DETTE = this.roundMoney(allocation.DETTE || 0);

    return {
      totalAssets,
      totalLiabilities,
      netWorth: this.roundMoney(totalAssets - totalLiabilities),
      accounts: accountResults,
      allocation
    };
  }

  // --- 4. MODULE PILOTAGE BUDGÉTAIRE ZBB ---
  computeBudgetZBB(period = null) {
    const { year: targetYear, month: targetMonth } = this.resolvePeriod(period);
    if (!this.data) return { mois: `${targetYear}-${targetMonth}`, totalAlloue: 0, totalConsomme: 0, totalEcart: 0, consumptionRate: 0, categories: [] };

    const isAllYear = (targetMonth === '00');
    const targetMonthStr = isAllYear ? targetYear : `${targetYear}-${targetMonth}`;
    const depenseCategories = (this.data.categories || []).filter(c => c.type === 'DEPENSE');

    const catRealise = {};
    (this.data.transactions || []).forEach(tx => {
      const matchPeriod = isAllYear
        ? (tx.date && tx.date.startsWith(`${targetYear}-`))
        : (tx.date && tx.date.startsWith(`${targetYear}-${targetMonth}`));
      if (matchPeriod) {
        const converted = this.convert(tx.montant, tx.devise || this.currency, this.currency);
        catRealise[tx.id_categorie] = (catRealise[tx.id_categorie] || 0) + converted;
      }
    });

    let totalAlloue = 0;
    let totalConsomme = 0;
    const catResults = [];

    depenseCategories.forEach(cat => {
      let alloue = 0;
      (this.data.budgets || []).forEach(b => {
        const bMois = b.mois || b.mois_annee || '';
        const matchMonth = isAllYear ? bMois.startsWith(`${targetYear}-`) : (bMois === targetMonthStr);
        if (matchMonth && String(b.id_categorie) === String(cat.id)) {
          const rawAmt = (b.montant !== undefined) ? b.montant : (b.montant_alloue !== undefined ? b.montant_alloue : 0);
          const bDev = b.devise || 'MUR';
          alloue += this.convert(rawAmt, bDev, this.currency);
        }
      });

      alloue = this.roundMoney(alloue);
      const consomme = this.roundMoney(catRealise[cat.id] || 0);
      if (alloue === 0 && consomme === 0) return;

      const ecart = this.roundMoney(alloue - consomme);
      const rawPct = alloue > 0 ? (consomme / alloue) * 100 : (consomme > 0 ? 100 : 0);

      totalAlloue += alloue;
      totalConsomme += consomme;

      let status = 'OK';
      if (ecart < 0) status = 'OVER';
      else if (rawPct >= 80) status = 'WARN';

      // Status filter
      const matchStatus = (this.budgetStatusFilter === 'ALL')
        || (this.budgetStatusFilter === 'OVER' && status === 'OVER')
        || (this.budgetStatusFilter === 'WARN' && status === 'WARN')
        || (this.budgetStatusFilter === 'OK' && status === 'OK');

      if (matchStatus) {
        catResults.push({
          id: cat.id,
          nom: cat.nom,
          nomTraduit: this.trEntity(cat.nom),
          alloue,
          consomme,
          ecart,
          pct: Math.min(this.roundMoney(rawPct), 100),
          rawPct: this.roundMoney(rawPct),
          status,
          isOverBudget: ecart < 0
        });
      }
    });

    catResults.sort((a, b) => b.rawPct - a.rawPct);

    totalAlloue = this.roundMoney(totalAlloue);
    totalConsomme = this.roundMoney(totalConsomme);
    const totalEcart = this.roundMoney(totalAlloue - totalConsomme);
    const consumptionRate = totalAlloue > 0 ? this.roundMoney((totalConsomme / totalAlloue) * 100) : (totalConsomme > 0 ? 100 : 0);

    return {
      mois: isAllYear ? `${targetYear} (${this.tr('ANNÉE ENTIÈRE')})` : targetMonthStr,
      totalAlloue,
      totalConsomme,
      totalEcart,
      consumptionRate,
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
    this.logHistory('TX_ADD', 'transaction', newTx.id, `${newTx.montant} ${newTx.devise} - ${cleanNotes || 'Opération'}`, newTx);
    this.save();
    this.notify();
    return newTx;
  }

  deleteTransaction(id) {
    if (!this.data || !this.data.transactions) return false;
    const idx = this.data.transactions.findIndex(t => t.id === Number(id));
    if (idx !== -1) {
      const removed = this.data.transactions[idx];
      this.data.transactions.splice(idx, 1);
      this.logHistory('TX_DELETE', 'transaction', id, `Suppression #${id} (${removed.montant} ${removed.devise})`, removed);
      this.save();
      this.notify();
      return true;
    }
    return false;
  }

  computeBudget(period = null) {
    return this.computeBudgetZBB(period);
  }

  allocateBudget(idCategorie, montant, mois = null) {
    if (!this.data) return false;
    const targetMois = mois || ((this.month === '00') ? `${this.year}-01` : `${this.year}-${this.month}`);
    const numAmount = this.roundMoney(Math.abs(parseFloat(montant) || 0));

    if (!Array.isArray(this.data.budgets)) {
      this.data.budgets = [];
    }

    const existing = this.data.budgets.find(b => (b.mois === targetMois || b.mois_annee === targetMois) && String(b.id_categorie) === String(idCategorie));
    if (existing) {
      existing.montant = numAmount;
      existing.montant_alloue = numAmount;
      existing.mois = targetMois;
      existing.mois_annee = targetMois;
      existing.devise = this.currency;
      this.logHistory('BUDGET_ALLOC', 'budget', existing.id, `Mise à jour enveloppe ${targetMois} : ${numAmount} ${this.currency}`, existing);
    } else {
      const nextId = (this.data.budgets.length > 0)
        ? Math.max(...this.data.budgets.map(b => b.id || 0)) + 1
        : 1;
      const newB = {
        id: nextId,
        mois: targetMois,
        mois_annee: targetMois,
        id_categorie: String(idCategorie),
        montant: numAmount,
        montant_alloue: numAmount,
        devise: this.currency
      };
      this.data.budgets.push(newB);
      this.logHistory('BUDGET_ALLOC', 'budget', nextId, `Nouvelle enveloppe ${targetMois} : ${numAmount} ${this.currency}`, newB);
    }

    this.save();
    this.notify();
    return true;
  }

  registerEntityTranslations(name) {
    if (!this.data) return;
    if (!this.data.dictionary) this.data.dictionary = {};
    const cleanNom = this.cleanText(name);
    const lower = cleanNom.toLowerCase();
    if (GFMP_DB_LOOKUP[lower]) {
      this.data.dictionary[cleanNom] = { ...GFMP_DB_LOOKUP[lower] };
    } else if (!this.data.dictionary[cleanNom]) {
      this.data.dictionary[cleanNom] = {
        FR: cleanNom,
        EN: cleanNom,
        ES: cleanNom,
        PT: cleanNom
      };
    }
  }

  addAccount(nom, classe = 'LIQUIDITE', devise = this.currency) {
    if (!this.data.comptes) this.data.comptes = [];
    const cleanNom = this.cleanText(nom);
    const existing = this.existeDeja(this.data.comptes, cleanNom);
    if (existing) {
      return { ...existing, isNew: false };
    }
    this.registerEntityTranslations(cleanNom);
    const nextId = this.genererNouvelId(this.data.comptes);
    const newCpt = { id: nextId, nom: cleanNom, classe, devise, actif: 'OUI' };
    this.data.comptes.push(newCpt);
    this.logHistory('ACCOUNT_ADD', 'account', nextId, `Ajout compte ${cleanNom} (${classe})`, newCpt);
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
    this.registerEntityTranslations(cleanNom);
    const nextId = this.genererNouvelId(this.data.categories);
    const newCat = { id: nextId, nom: cleanNom, type };
    this.data.categories.push(newCat);
    this.logHistory('CATEGORY_ADD', 'category', nextId, `Ajout catégorie ${cleanNom} (${type})`, newCat);
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
    this.registerEntityTranslations(cleanNom);
    const nextId = this.genererNouvelId(this.data.tiers);
    const newTiers = { id: nextId, nom: cleanNom, type };
    this.data.tiers.push(newTiers);
    this.logHistory('TIERS_ADD', 'tiers', nextId, `Ajout tiers ${cleanNom} (${type})`, newTiers);
    this.save();
    this.notify();
    return { ...newTiers, isNew: true };
  }

  // --- 6. HISTORISATION & AUDIT LOG ---
  logHistory(action, entity, entityId, description, metadata = {}) {
    if (!this.data) return null;
    if (!Array.isArray(this.data.audit_history)) {
      this.data.audit_history = [];
    }

    const entry = {
      id: 'HIST_' + Date.now() + '_' + Math.random().toString(36).substr(2, 4),
      timestamp: new Date().toISOString(),
      dateDisplay: new Date().toLocaleDateString(this.lang === 'FR' ? 'fr-FR' : 'en-US', {
        day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit'
      }),
      action, // 'TX_ADD', 'TX_DELETE', 'BUDGET_ALLOC', 'ACCOUNT_ADD', 'CATEGORY_ADD', 'TIERS_ADD', 'DB_INIT', 'DB_REFRESH', 'DB_PURGE', 'DB_BACKUP', 'DB_RESTORE'
      entity, // 'transaction', 'budget', 'account', 'database'
      entityId: String(entityId || ''),
      description,
      currency: this.currency,
      metadata
    };

    this.data.audit_history.unshift(entry);
    // Keep max 500 audit items
    if (this.data.audit_history.length > 500) {
      this.data.audit_history = this.data.audit_history.slice(0, 500);
    }
    return entry;
  }

  getHistory(limit = 100) {
    if (!this.data || !Array.isArray(this.data.audit_history)) return [];
    return this.data.audit_history.slice(0, limit);
  }

  // --- 7. DATABASE OPERATIONS (INITIALISER, RAFRAÎCHIR, PURGER, BACKUP) ---
  initDatabase() {
    if (typeof window !== 'undefined' && window.GFMP_BASELINE) {
      this.data = JSON.parse(JSON.stringify(window.GFMP_BASELINE));
      this.dataMode = 'DEMO';
      this.data.dataMode = 'DEMO';
      if (!this.data.config) this.data.config = {};
      this.data.config.SYS_DATA_MODE = 'DEMO';
      this.logHistory('DB_INIT', 'database', 'system', 'Réinitialisation de la base de données aux paramètres d\'usine');
      this.save();
      this.notify();
      return true;
    }
    return false;
  }

  refreshDatabase() {
    if (!this.data) return false;
    // Re-verify and clean data integrity
    if (Array.isArray(this.data.comptes)) {
      this.data.comptes.forEach(c => {
        c.nom = this.cleanAssetName(c.nom);
      });
    }
    this.logHistory('DB_REFRESH', 'database', 'system', 'Rafraîchissement complet des calculs et du cache analytique');
    this.save();
    this.notify();
    return true;
  }

  purgeDatabase() {
    if (!this.data) return false;
    const txCount = Array.isArray(this.data.transactions) ? this.data.transactions.length : 0;
    this.data.transactions = [];
    this.data.budgets = [];
    this.dataMode = 'EMPTY';
    this.data.dataMode = 'EMPTY';
    if (!this.data.config) this.data.config = {};
    this.data.config.SYS_DATA_MODE = 'EMPTY';
    this.selectedCategoryFilter = null;
    this.logHistory('DB_PURGE', 'database', 'system', `Purge des données : ${txCount} transactions et budgets réinitialisés`);
    this.save();
    this.notify();
    return true;
  }

  purgeTransactions() {
    return this.purgeDatabase();
  }

  backupDatabase() {
    this.logHistory('DB_BACKUP', 'database', 'system', 'Sauvegarde intégrale de la base de données exportée');
    this.save();
    return JSON.stringify(this.data, null, 2);
  }

  restoreDatabase(input) {
    try {
      let parsed = input;
      if (typeof input === 'string') {
        if (!input.trim()) return { success: false, error: 'Données JSON vides' };
        parsed = JSON.parse(input);
      }
      if (parsed && typeof parsed === 'object' && Array.isArray(parsed.comptes) && Array.isArray(parsed.categories)) {
        this.data = parsed;
        if (!Array.isArray(this.data.transactions)) this.data.transactions = [];
        if (!Array.isArray(this.data.budgets)) this.data.budgets = [];
        if (!Array.isArray(this.data.audit_history)) this.data.audit_history = [];
        if (!this.data.config) this.data.config = {};
        if (this.data.config.SYS_DEVISE_BASE) this.currency = this.data.config.SYS_DEVISE_BASE;
        if (this.data.config.LANGUE_UI) this.lang = this.data.config.LANGUE_UI;
        this.logHistory('DB_RESTORE', 'database', 'system', `Restauration de la base : ${this.data.transactions.length} transactions chargées`);
        this.save();
        this.notify();
        return { success: true, count: this.data.transactions.length };
      }
      return { success: false, error: 'Format de fichier non conforme (structure attendue non reconnue)' };
    } catch (e) {
      return { success: false, error: e.message || 'Erreur d\'analyse JSON' };
    }
  }

  exportBackupJSON() {
    return this.backupDatabase();
  }

  importBackupJSON(jsonStr) {
    return this.restoreDatabase(jsonStr);
  }

  // --- 8. HUB CONSOLIDATED DASHBOARD ANALYTICS ---
  getConsolidatedHubData(period = null) {
    const { year: targetYear, month: targetMonth } = this.resolvePeriod(period);
    const cf = this.computeCashflow(period);
    const nw = this.computeNetWorth(period);
    const bg = this.computeBudgetZBB(period);

    // Financial health score (10 to 100) — neutral when netCashflow === 0
    let healthScore = 50;
    if (cf.netCashflow > 0) healthScore += 15;
    else if (cf.netCashflow < 0) healthScore -= 15;

    if (cf.savingsRate >= 20) healthScore += 15;
    else if (cf.savingsRate >= 10) healthScore += 8;

    if (bg.consumptionRate > 0 && bg.consumptionRate <= 85) healthScore += 10;
    else if (bg.consumptionRate <= 100) healthScore += 5;
    else if (bg.consumptionRate > 100) healthScore -= 15;

    const totalAssets = nw.totalAssets || 0;
    const totalLiabilities = nw.totalLiabilities || 0;
    const debtRatio = totalAssets > 0 ? Math.round((totalLiabilities / totalAssets) * 100) : (totalLiabilities > 0 ? 100 : 0);
    if (debtRatio <= 25) healthScore += 10;
    else if (debtRatio > 60) healthScore -= 10;

    const savingsRate = Math.round(cf.savingsRate) || 0;
    const totalTxCount = (this.data && Array.isArray(this.data.transactions)) ? this.data.transactions.length : 0;
    const totalBudgetCount = (this.data && Array.isArray(this.data.budgets)) ? this.data.budgets.length : 0;

    return {
      hubFilter: this.hubFilter,
      period: `${this.getMonthLabel(targetMonth)} ${targetYear}`,
      currency: this.currency,
      dataMode: this.getDataMode(),
      hasData: (totalTxCount > 0 || totalBudgetCount > 0),
      hasCashflowData: (cf.totalRevenus > 0 || cf.totalDepenses > 0 || cf.transactions.length > 0),
      hasPatrimoineData: (totalAssets > 0 || totalLiabilities > 0),
      hasBudgetData: (bg.totalAlloue > 0 || bg.totalConsomme > 0 || (bg.categories && bg.categories.length > 0)),
      totalTxCount,
      totalBudgetCount,
      healthScore: Math.max(10, Math.min(100, healthScore)),
      // Net Worth
      netWorth: nw.netWorth,
      totalAssets,
      totalLiabilities,
      debtRatio,
      allocation: nw.allocation,
      // Cashflow
      totalRevenus: cf.totalRevenus,
      totalDepenses: cf.totalDepenses,
      netCashflow: cf.netCashflow,
      savingsRate,
      txCount: cf.transactions.length,
      breakdown: cf.breakdown,
      // Budget
      totalAlloue: bg.totalAlloue,
      totalConsomme: bg.totalConsomme,
      totalEcart: bg.totalEcart,
      consumptionRate: Math.round(bg.consumptionRate),
      isOverBudget: bg.totalEcart < 0,
      topBudgetCategories: (bg.categories || []).slice(0, 4),
      // Top 3 Expense categories
      topCategories: cf.breakdown.slice(0, 3).map(b => ({
        ...b,
        nom: this.trEntity(b.nom)
      })),
      // Top Asset accounts (only positive active balances so empty portfolios don't render 0,00 rows)
      topAssets: (nw.accounts || [])
        .filter(a => a.classe !== 'DETTE' && a.soldeConverti > 0)
        .sort((a, b) => b.soldeConverti - a.soldeConverti)
        .slice(0, 4)
        .map(a => ({
          nom: this.cleanAssetName(a.nom),
          classe: this.trEntity(a.classe),
          solde: this.formatMoney(a.soldeConverti),
          rawSolde: a.soldeConverti
        })),
      // Recent 5 Transactions (filtered by selectedCategoryFilter if active)
      recentTransactions: ((this.data && this.data.transactions) ? [...this.data.transactions] : [])
        .filter(tx => {
          if (!this.selectedCategoryFilter) return true;
          const cat = (this.data.categories || []).find(c => String(c.id) === String(tx.id_categorie));
          return cat && cat.nom === this.selectedCategoryFilter;
        })
        .sort((a, b) => (b.date || '').localeCompare(a.date || ''))
        .slice(0, 5)
        .map(tx => {
          const cpt = (this.data.comptes || []).find(c => String(c.id) === String(tx.id_compte));
          const cat = (this.data.categories || []).find(c => String(c.id) === String(tx.id_categorie));
          const montantConv = this.convert(tx.montant, tx.devise || this.currency, this.currency);
          return {
            id: tx.id,
            date: tx.date,
            rawCategorie: cat ? cat.nom : '',
            categorie: cat ? this.trEntity(cat.nom) : (this.tr('OPERATION_DEFAULT') || 'Opération'),
            compte: cpt ? this.cleanAssetName(cpt.nom) : (this.tr('ASSETS_DEFAULT') || 'Actifs'),
            type: cat ? cat.type : 'DEPENSE',
            montant: this.formatMoney(montantConv),
            rawMontant: montantConv
          };
        })
    };
  }

  generateFinancialReport(period = null) {
    const { year: targetYear, month: targetMonth } = this.resolvePeriod(period);
    const cf = this.computeCashflow(period);
    const nw = this.computeNetWorth(period);
    const bg = this.computeBudgetZBB(period);
    const hasData = (cf.totalRevenus > 0 || cf.totalDepenses > 0 || nw.totalAssets > 0 || nw.totalLiabilities > 0 || bg.totalAlloue > 0);

    return {
      period: `${this.getMonthLabel(targetMonth)} ${targetYear}`,
      currency: this.currency,
      dataMode: this.getDataMode(),
      hasData,
      txCount: cf.transactions.length,
      rawRevenus: cf.totalRevenus,
      rawDepenses: cf.totalDepenses,
      rawNetWorth: nw.netWorth,
      revenus: this.formatMoney(cf.totalRevenus),
      depenses: this.formatMoney(cf.totalDepenses),
      netCashflow: this.formatMoney(cf.netCashflow),
      savingsRate: `${Math.round(cf.savingsRate)}%`,
      netWorth: this.formatMoney(nw.netWorth),
      totalAssets: this.formatMoney(nw.totalAssets),
      totalLiabilities: this.formatMoney(nw.totalLiabilities),
      budgetConsumption: `${Math.round(bg.consumptionRate)}%`,
      topExpenses: cf.breakdown.slice(0, 5).map(b => ({
        nom: this.trEntity(b.nom),
        montant: this.formatMoney(b.montant),
        pct: `${Math.round(b.pct)}%`
      }))
    };
  }
}

window.gfmpEngine = new GFMPEngine();
