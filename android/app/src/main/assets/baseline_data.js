window.GFMP_BASELINE = {
  "app_name": "GFMP — Global Finance Management Platform",
  "version": "1.0",
  "dataset_version": "2.0",
  "export_date": "2026-03-01T12:00:00.000Z",
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
      "nom": "Clinique & Pharmacie",
      "type": "SANTE"
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
      "date": "2025-09-02",
      "id_compte": "1",
      "id_categorie": "4",
      "id_tiers": "5",
      "montant": 15000,
      "devise": "MUR",
      "notes": "Premier loyer résidence"
    },
    {
      "id": 2,
      "date": "2025-09-05",
      "id_compte": "1",
      "id_categorie": "8",
      "id_tiers": "4",
      "montant": 125000,
      "devise": "MUR",
      "notes": "Ameublement et électroménager installation"
    },
    {
      "id": 3,
      "date": "2025-09-10",
      "id_compte": "1",
      "id_categorie": "5",
      "id_tiers": "4",
      "montant": 12200,
      "devise": "MUR",
      "notes": "Courses d'installation et provisions"
    },
    {
      "id": 4,
      "date": "2025-09-15",
      "id_compte": "1",
      "id_categorie": "6",
      "id_tiers": "7",
      "montant": 4800,
      "devise": "MUR",
      "notes": "Carburant et révision véhicule"
    },
    {
      "id": 5,
      "date": "2025-09-20",
      "id_compte": "1",
      "id_categorie": "7",
      "id_tiers": "6",
      "montant": 3200,
      "devise": "MUR",
      "notes": "Bilan médical et vaccins"
    },
    {
      "id": 6,
      "date": "2025-09-25",
      "id_compte": "1",
      "id_categorie": "1",
      "id_tiers": "1",
      "montant": 85000,
      "devise": "MUR",
      "notes": "Salaire Septembre KIMYA ÉDITION"
    },
    {
      "id": 7,
      "date": "2025-09-28",
      "id_compte": "1",
      "id_categorie": "10",
      "id_tiers": "2",
      "montant": 20000,
      "devise": "MUR",
      "notes": "Constitution épargne de précaution"
    },
    {
      "id": 8,
      "date": "2025-09-29",
      "id_compte": "8",
      "id_categorie": "8",
      "id_tiers": "4",
      "montant": 20000,
      "devise": "MUR",
      "notes": "Restauration et réception emménagement"
    },
    {
      "id": 9,
      "date": "2025-10-02",
      "id_compte": "1",
      "id_categorie": "4",
      "id_tiers": "5",
      "montant": 15000,
      "devise": "MUR",
      "notes": "Loyer Octobre"
    },
    {
      "id": 10,
      "date": "2025-10-07",
      "id_compte": "1",
      "id_categorie": "5",
      "id_tiers": "4",
      "montant": 8900,
      "devise": "MUR",
      "notes": "Supermarché courses hebdomadaires"
    },
    {
      "id": 11,
      "date": "2025-10-12",
      "id_compte": "5",
      "id_categorie": "2",
      "id_tiers": "2",
      "montant": 190.65,
      "devise": "EUR",
      "notes": "Dividendes trimestriels ETF World"
    },
    {
      "id": 12,
      "date": "2025-10-16",
      "id_compte": "1",
      "id_categorie": "6",
      "id_tiers": "7",
      "montant": 4200,
      "devise": "MUR",
      "notes": "Plein carburant station service"
    },
    {
      "id": 13,
      "date": "2025-10-21",
      "id_compte": "1",
      "id_categorie": "8",
      "id_tiers": "4",
      "montant": 6500,
      "devise": "MUR",
      "notes": "Sorties culturelles et cinémas"
    },
    {
      "id": 14,
      "date": "2025-10-25",
      "id_compte": "1",
      "id_categorie": "1",
      "id_tiers": "1",
      "montant": 85000,
      "devise": "MUR",
      "notes": "Salaire Octobre KIMYA ÉDITION"
    },
    {
      "id": 15,
      "date": "2025-10-28",
      "id_compte": "1",
      "id_categorie": "10",
      "id_tiers": "2",
      "montant": 15000,
      "devise": "MUR",
      "notes": "Virement épargne mensuelle"
    },
    {
      "id": 16,
      "date": "2025-10-30",
      "id_compte": "1",
      "id_categorie": "7",
      "id_tiers": "6",
      "montant": 5400,
      "devise": "MUR",
      "notes": "Pharmacie et soins dentaires"
    },
    {
      "id": 17,
      "date": "2025-11-02",
      "id_compte": "1",
      "id_categorie": "4",
      "id_tiers": "5",
      "montant": 15000,
      "devise": "MUR",
      "notes": "Loyer Novembre"
    },
    {
      "id": 18,
      "date": "2025-11-06",
      "id_compte": "1",
      "id_categorie": "5",
      "id_tiers": "4",
      "montant": 9400,
      "devise": "MUR",
      "notes": "Supermarché et produits frais"
    },
    {
      "id": 19,
      "date": "2025-11-12",
      "id_compte": "4",
      "id_categorie": "2",
      "id_tiers": "2",
      "montant": 3482,
      "devise": "MUR",
      "notes": "Rendement fonds euros assurance vie"
    },
    {
      "id": 20,
      "date": "2025-11-16",
      "id_compte": "1",
      "id_categorie": "6",
      "id_tiers": "7",
      "montant": 4600,
      "devise": "MUR",
      "notes": "Assurance auto et carburant"
    },
    {
      "id": 21,
      "date": "2025-11-20",
      "id_compte": "6",
      "id_categorie": "10",
      "id_tiers": "2",
      "montant": 250,
      "devise": "USD",
      "notes": "DCA mensuel Bitcoin & Ethereum"
    },
    {
      "id": 22,
      "date": "2025-11-25",
      "id_compte": "1",
      "id_categorie": "1",
      "id_tiers": "1",
      "montant": 85000,
      "devise": "MUR",
      "notes": "Salaire Novembre KIMYA ÉDITION"
    },
    {
      "id": 23,
      "date": "2025-11-27",
      "id_compte": "1",
      "id_categorie": "8",
      "id_tiers": "4",
      "montant": 6200,
      "devise": "MUR",
      "notes": "Dîners restaurants et théâtre"
    },
    {
      "id": 24,
      "date": "2025-11-29",
      "id_compte": "1",
      "id_categorie": "7",
      "id_tiers": "6",
      "montant": 2500,
      "devise": "MUR",
      "notes": "Soins ophtalmologiques"
    },
    {
      "id": 25,
      "date": "2025-12-02",
      "id_compte": "1",
      "id_categorie": "4",
      "id_tiers": "5",
      "montant": 15000,
      "devise": "MUR",
      "notes": "Loyer Décembre"
    },
    {
      "id": 26,
      "date": "2025-12-08",
      "id_compte": "1",
      "id_categorie": "5",
      "id_tiers": "4",
      "montant": 14500,
      "devise": "MUR",
      "notes": "Courses festives et repas de fête"
    },
    {
      "id": 27,
      "date": "2025-12-14",
      "id_compte": "1",
      "id_categorie": "8",
      "id_tiers": "4",
      "montant": 18500,
      "devise": "MUR",
      "notes": "Cadeaux de Noël et célébrations"
    },
    {
      "id": 28,
      "date": "2025-12-18",
      "id_compte": "1",
      "id_categorie": "6",
      "id_tiers": "7",
      "montant": 5200,
      "devise": "MUR",
      "notes": "Trajets fêtes et carburant"
    },
    {
      "id": 29,
      "date": "2025-12-20",
      "id_compte": "1",
      "id_categorie": "1",
      "id_tiers": "1",
      "montant": 120000,
      "devise": "MUR",
      "notes": "Salaire Décembre et Prime Annuelle"
    },
    {
      "id": 30,
      "date": "2025-12-22",
      "id_compte": "5",
      "id_categorie": "2",
      "id_tiers": "2",
      "montant": 4200,
      "devise": "MUR",
      "notes": "Dividendes semestriels actions"
    },
    {
      "id": 31,
      "date": "2025-12-26",
      "id_compte": "1",
      "id_categorie": "10",
      "id_tiers": "2",
      "montant": 40000,
      "devise": "MUR",
      "notes": "Placement prime sur livret épargne"
    },
    {
      "id": 32,
      "date": "2025-12-28",
      "id_compte": "1",
      "id_categorie": "7",
      "id_tiers": "6",
      "montant": 2800,
      "devise": "MUR",
      "notes": "Pharmacie d'hiver"
    },
    {
      "id": 33,
      "date": "2025-12-30",
      "id_compte": "1",
      "id_categorie": "8",
      "id_tiers": "4",
      "montant": 4000,
      "devise": "MUR",
      "notes": "Réveillon Saint-Sylvestre"
    },
    {
      "id": 34,
      "date": "2026-01-02",
      "id_compte": "1",
      "id_categorie": "4",
      "id_tiers": "5",
      "montant": 15000,
      "devise": "MUR",
      "notes": "Loyer Janvier 2026"
    },
    {
      "id": 35,
      "date": "2026-01-05",
      "id_compte": "1",
      "id_categorie": "5",
      "id_tiers": "4",
      "montant": 9800,
      "devise": "MUR",
      "notes": "Supermarché rentrée"
    },
    {
      "id": 36,
      "date": "2026-01-10",
      "id_compte": "1",
      "id_categorie": "9",
      "id_tiers": "3",
      "montant": 10000,
      "devise": "MUR",
      "notes": "Acompte fiscal trimestriel MRA"
    },
    {
      "id": 37,
      "date": "2026-01-14",
      "id_compte": "1",
      "id_categorie": "6",
      "id_tiers": "7",
      "montant": 4400,
      "devise": "MUR",
      "notes": "Entretien annuel et essence"
    },
    {
      "id": 38,
      "date": "2026-01-18",
      "id_compte": "1",
      "id_categorie": "7",
      "id_tiers": "6",
      "montant": 3800,
      "devise": "MUR",
      "notes": "Cotisation mutuelle santé"
    },
    {
      "id": 39,
      "date": "2026-01-22",
      "id_compte": "1",
      "id_categorie": "8",
      "id_tiers": "4",
      "montant": 7400,
      "devise": "MUR",
      "notes": "Sorties week-end et livres"
    },
    {
      "id": 40,
      "date": "2026-01-25",
      "id_compte": "1",
      "id_categorie": "1",
      "id_tiers": "1",
      "montant": 85000,
      "devise": "MUR",
      "notes": "Salaire Janvier KIMYA ÉDITION"
    },
    {
      "id": 41,
      "date": "2026-01-27",
      "id_compte": "1",
      "id_categorie": "3",
      "id_tiers": "3",
      "montant": 8052,
      "devise": "MUR",
      "notes": "Allocation et crédit d'impôt"
    },
    {
      "id": 42,
      "date": "2026-01-29",
      "id_compte": "1",
      "id_categorie": "10",
      "id_tiers": "2",
      "montant": 25000,
      "devise": "MUR",
      "notes": "Épargne programmée vers livret"
    },
    {
      "id": 43,
      "date": "2026-02-02",
      "id_compte": "1",
      "id_categorie": "4",
      "id_tiers": "5",
      "montant": 15000,
      "devise": "MUR",
      "notes": "Loyer Février Résidence KIMYA"
    },
    {
      "id": 44,
      "date": "2026-02-05",
      "id_compte": "1",
      "id_categorie": "5",
      "id_tiers": "4",
      "montant": 4800,
      "devise": "MUR",
      "notes": "Super U Grand Baie - Alimentation"
    },
    {
      "id": 45,
      "date": "2026-02-09",
      "id_compte": "1",
      "id_categorie": "6",
      "id_tiers": "7",
      "montant": 4500,
      "devise": "MUR",
      "notes": "Plein carburant TotalEnergies"
    },
    {
      "id": 46,
      "date": "2026-02-12",
      "id_compte": "5",
      "id_categorie": "2",
      "id_tiers": "2",
      "montant": 8.71,
      "devise": "EUR",
      "notes": "Dividendes Portefeuille PEA (441.28 MUR)"
    },
    {
      "id": 47,
      "date": "2026-02-14",
      "id_compte": "1",
      "id_categorie": "8",
      "id_tiers": "4",
      "montant": 3000,
      "devise": "MUR",
      "notes": "Dîner Saint-Valentin restaurant"
    },
    {
      "id": 48,
      "date": "2026-02-17",
      "id_compte": "1",
      "id_categorie": "5",
      "id_tiers": "4",
      "montant": 5000,
      "devise": "MUR",
      "notes": "Marché central et épicerie fine"
    },
    {
      "id": 49,
      "date": "2026-02-19",
      "id_compte": "1",
      "id_categorie": "7",
      "id_tiers": "6",
      "montant": 3200,
      "devise": "MUR",
      "notes": "Consultation clinique et pharmacie"
    },
    {
      "id": 50,
      "date": "2026-02-21",
      "id_compte": "8",
      "id_categorie": "8",
      "id_tiers": "4",
      "montant": 400.0,
      "devise": "EUR",
      "notes": "Achat matériel tech Apple Store (20,271 MUR - Cas Dépassement Budget)"
    },
    {
      "id": 51,
      "date": "2026-02-24",
      "id_compte": "1",
      "id_categorie": "1",
      "id_tiers": "1",
      "montant": 78000,
      "devise": "MUR",
      "notes": "Salaire Février KIMYA ÉDITION"
    },
    {
      "id": 52,
      "date": "2026-02-25",
      "id_compte": "1",
      "id_categorie": "1",
      "id_tiers": "1",
      "montant": 25000,
      "devise": "MUR",
      "notes": "Honoraires mission consulting financier"
    },
    {
      "id": 53,
      "date": "2026-02-27",
      "id_compte": "1",
      "id_categorie": "9",
      "id_tiers": "3",
      "montant": 10000,
      "devise": "MUR",
      "notes": "Taxe foncière et droits communaux"
    },
    {
      "id": 54,
      "date": "2026-03-01",
      "id_compte": "1",
      "id_categorie": "4",
      "id_tiers": "5",
      "montant": 15000,
      "devise": "MUR",
      "notes": "Loyer Mars 2026"
    },
    {
      "id": 55,
      "date": "2026-03-03",
      "id_compte": "1",
      "id_categorie": "5",
      "id_tiers": "4",
      "montant": 3500,
      "devise": "MUR",
      "notes": "Courses début de mois"
    },
    {
      "id": 56,
      "date": "2026-03-05",
      "id_compte": "1",
      "id_categorie": "6",
      "id_tiers": "7",
      "montant": 2000,
      "devise": "MUR",
      "notes": "Recharge carburant"
    },
    {
      "id": 57,
      "date": "2026-03-10",
      "id_compte": "4",
      "id_categorie": "2",
      "id_tiers": "2",
      "montant": 5500,
      "devise": "MUR",
      "notes": "Intérêts trimestriels fonds euros"
    },
    {
      "id": 58,
      "date": "2026-03-15",
      "id_compte": "1",
      "id_categorie": "8",
      "id_tiers": "4",
      "montant": 2500,
      "devise": "MUR",
      "notes": "Sortie cinéma et détente"
    },
    {
      "id": 59,
      "date": "2026-03-20",
      "id_compte": "1",
      "id_categorie": "10",
      "id_tiers": "2",
      "montant": 15000,
      "devise": "MUR",
      "notes": "Virement épargne préventive"
    },
    {
      "id": 60,
      "date": "2026-03-25",
      "id_compte": "1",
      "id_categorie": "1",
      "id_tiers": "1",
      "montant": 85000,
      "devise": "MUR",
      "notes": "Salaire Mars KIMYA ÉDITION"
    }
  ],
  "budgets": [
    {
      "id": 1,
      "mois_annee": "2026-01",
      "id_categorie": "4",
      "montant_alloue": 15000,
      "notes": "Budget Logement Janvier"
    },
    {
      "id": 2,
      "mois_annee": "2026-01",
      "id_categorie": "5",
      "montant_alloue": 12000,
      "notes": "Budget Alimentation Janvier"
    },
    {
      "id": 3,
      "mois_annee": "2026-01",
      "id_categorie": "6",
      "montant_alloue": 6000,
      "notes": "Budget Transports Janvier"
    },
    {
      "id": 4,
      "mois_annee": "2026-01",
      "id_categorie": "7",
      "montant_alloue": 5000,
      "notes": "Budget Santé Janvier"
    },
    {
      "id": 5,
      "mois_annee": "2026-01",
      "id_categorie": "8",
      "montant_alloue": 8000,
      "notes": "Budget Loisirs Janvier"
    },
    {
      "id": 6,
      "mois_annee": "2026-02",
      "id_categorie": "4",
      "montant_alloue": 15000,
      "notes": "Budget Logement (100% OK)"
    },
    {
      "id": 7,
      "mois_annee": "2026-02",
      "id_categorie": "5",
      "montant_alloue": 12000,
      "notes": "Budget Alimentation (81.7% Jauge Dorée)"
    },
    {
      "id": 8,
      "mois_annee": "2026-02",
      "id_categorie": "6",
      "montant_alloue": 6000,
      "notes": "Budget Transports (75% Vert)"
    },
    {
      "id": 9,
      "mois_annee": "2026-02",
      "id_categorie": "7",
      "montant_alloue": 5000,
      "notes": "Budget Santé (64% Vert)"
    },
    {
      "id": 10,
      "mois_annee": "2026-02",
      "id_categorie": "8",
      "montant_alloue": 7000,
      "notes": "Budget Loisirs (332% Dépassement Rouge)"
    },
    {
      "id": 11,
      "mois_annee": "2026-02",
      "id_categorie": "9",
      "montant_alloue": 10000,
      "notes": "Budget Impôts (100% OK)"
    },
    {
      "id": 12,
      "mois_annee": "2026-03",
      "id_categorie": "4",
      "montant_alloue": 15000,
      "notes": "Budget Logement Mars"
    },
    {
      "id": 13,
      "mois_annee": "2026-03",
      "id_categorie": "5",
      "montant_alloue": 12000,
      "notes": "Budget Alimentation Mars"
    },
    {
      "id": 14,
      "mois_annee": "2026-03",
      "id_categorie": "6",
      "montant_alloue": 6000,
      "notes": "Budget Transports Mars"
    },
    {
      "id": 15,
      "mois_annee": "2026-03",
      "id_categorie": "7",
      "montant_alloue": 5000,
      "notes": "Budget Santé Mars"
    },
    {
      "id": 16,
      "mois_annee": "2026-03",
      "id_categorie": "8",
      "montant_alloue": 8000,
      "notes": "Budget Loisirs Mars"
    },
    {
      "id": 17,
      "mois_annee": "2026-03",
      "id_categorie": "9",
      "montant_alloue": 10000,
      "notes": "Budget Impôts Mars"
    }
  ],
  "config": {
    "SYS_DEVISE_BASE": "MUR",
    "LANGUE_UI": "FR",
    "TAUX_CHANGE": {
      "MUR": 1.0,
      "EUR": 50.6787,
      "USD": 46.5214,
      "GBP": 59.4321,
      "ZAR": 2.5412,
      "XOF": 0.0772
    },
    "TAUX_DATE_MAJ": "2026-03-01T12:00:00.000Z",
    "THEME": "DARK_ROYAL",
    "DATE_FORMAT": "YYYY-MM-DD",
    "SYS_ADMIN_HASH": "SFP_ADMIN_2026"
  },
  "audit_history": [
    {
      "timestamp": "2025-09-01T08:00:00.000Z",
      "action": "SYS_INIT",
      "entity": "SYSTEM",
      "entityId": "1",
      "description": "Initialisation du Système Financier Personnel (SFP v1.0)"
    },
    {
      "timestamp": "2025-09-02T10:00:00.000Z",
      "action": "CREATE_TX",
      "entity": "TRANSACTION",
      "entityId": "1",
      "description": "Saisie premier loyer résidence (15000 MUR)"
    },
    {
      "timestamp": "2026-01-01T00:00:00.000Z",
      "action": "ZBB_INIT",
      "entity": "BUDGET",
      "entityId": "1",
      "description": "Activation du pilotage Zero-Based Budgeting 2026"
    },
    {
      "timestamp": "2026-02-01T08:00:00.000Z",
      "action": "CALC_UPDATE",
      "entity": "CASHFLOW",
      "entityId": "2026-02",
      "description": "Calcul du cashflow mensuel Février 2026"
    },
    {
      "timestamp": "2026-03-01T12:00:00.000Z",
      "action": "TEST_DATA_EMBED",
      "entity": "DATABASE",
      "entityId": "60",
      "description": "Intégration du jeu de données de test SFP v1.0 complet (60 transactions, 17 budgets)"
    }
  ]
};
