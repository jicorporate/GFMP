const fs = require('fs');

const dbEntityDict = {
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

function trEntity(name, lang = 'FR') {
  if (!name) return '';
  const clean = String(name).replace(/\s+/g, ' ').trim();
  const entry = dbEntityDict[clean];
  if (entry && entry[lang]) {
    return entry[lang];
  }
  return clean;
}

console.log('Testing translation of "Salaire / Revenus Pro":');
['FR', 'EN', 'ES', 'PT'].forEach(l => console.log(`  ${l}: ${trEntity('Salaire / Revenus Pro', l)}`));

console.log('Testing translation of "Alimentation & Supermarché":');
['FR', 'EN', 'ES', 'PT'].forEach(l => console.log(`  ${l}: ${trEntity('Alimentation & Supermarché', l)}`));

console.log('Testing translation of "Compte Courant Principal":');
['FR', 'EN', 'ES', 'PT'].forEach(l => console.log(`  ${l}: ${trEntity('Compte Courant Principal', l)}`));

console.log('Testing translation of "LIQUIDITE":');
['FR', 'EN', 'ES', 'PT'].forEach(l => console.log(`  ${l}: ${trEntity('LIQUIDITE', l)}`));

console.log('Testing translation of "Premier loyer résidence":');
['FR', 'EN', 'ES', 'PT'].forEach(l => console.log(`  ${l}: ${trEntity('Premier loyer résidence', l)}`));
