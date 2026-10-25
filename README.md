# 🏛️ GFMP — Global Finance Management Platform
### Système Financier Personnel & Entreprise-Grade (SFP v1.0)
*Wealth Management • Star Schema Kimball • Multi-Plateforme (Excel VBA, Android Natif & PWA Offline)*

---

## 📋 Présentation Générale

**GFMP (Global Finance Management Platform)** est une suite financière et décisionnelle de standard entreprise-grade, conçue pour le pilotage rigoureux des flux financiers, du patrimoine net et des budgets. 

Initialement développée sous forme d'application maîtresse **Excel/VBA relationnelle en mémoire (Kimball Star Schema)**, la plateforme est désormais déployée en une solution **"Local-First" multi-plateforme** :
1. **Application Maîtresse Excel / VBA** : [GFMP_V1.0.xlsm](file:///c:/Users/KIMYAÉDITION/source/repos/GFMP/GFMP_V1.0.xlsm) (et version démo [GFMP_V1.xlsm](file:///c:/Users/KIMYAÉDITION/source/repos/GFMP/GFMP_V1.xlsm)).
2. **Application Mobile Android Autonome** : [dist/GFMP-v1.0.apk](file:///c:/Users/KIMYAÉDITION/source/repos/GFMP/dist/GFMP-v1.0.apk) (Installable par sideloading sans Play Store, zéro serveur, zéro PC).
3. **Application Web Mobile PWA** : [mobile-web/](file:///c:/Users/KIMYAÉDITION/source/repos/GFMP/mobile-web) (100% hors-ligne via Service Worker cache-first).

---

## 🗂️ Structure Organisée du Projet (Best Practices)

```
GFMP/
├── README.md                      # Présentation et documentation officielle GFMP
├── .gitignore                     # Configuration Git complète (Android, Gradle, Office)
├── GFMP_V1.0.xlsm                 # Classeur maître de production Excel / VBA
├── GFMP_V1.xlsm                   # Classeur de démonstration avec transactions exemples
├── GFMP-v1.0.apk                  # Package Android autonome racine (accès immédiat)
├── run_mobile_preview.ps1         # Script de lancement du serveur Wi-Fi smartphone
│
├── android/                       # Projet natif Android (API 34 / Gradle 9.1 / AGP 9.0.1)
│   ├── app/
│   │   ├── build.gradle.kts       # Namespace & ApplicationId "com.jicorporate.gfmp"
│   │   └── src/main/
│   │       ├── assets/            # Bundle web offline embarqué
│   │       ├── kotlin/com/jicorporate/gfmp/MainActivity.kt
│   │       └── res/               # Thème Theme.GFMP, icônes mipmap
│   ├── gradle/
│   ├── build.gradle.kts
│   └── settings.gradle.kts
│
├── mobile-web/                    # Application Web Mobile & Progressive Web App (PWA)
│   ├── index.html                 # Interface tactile (4 Solid Cards, Gatekeeper, Violet Zebra)
│   ├── styles.css                 # Charte graphique stricte SFP v1.0
│   ├── app.js                     # Contrôleur DOM, événements et modales
│   ├── analytics-engine.js        # Moteur analytique vectorisé pur JS (< 5ms)
│   ├── baseline_data.json         # Données relationnelles Kimball de référence
│   ├── manifest.json              # Manifeste PWA GFMP
│   ├── sw.js                      # Service Worker hors-ligne
│   └── assets/                    # Favicon et identité visuelle
│
├── vba/                           # Sources modulaires VBA exportées
│   ├── MOD_00_Architecture_v1.0.bas   # Définitions et constantes architecturales
│   ├── MOD_01_CoreEngine_v1.0.bas     # Moteur central, séquences d'ID, regex clean
│   ├── MOD_02_AppHome_Global_v1.0.bas # Hub Central, routage UCR, matrice i18n
│   ├── MOD_03_Gatekeeper_v1.0.bas     # Moteur transactionnel et intégrité
│   ├── MOD_04_Dashboard_ETL_v1.0.bas  # ETL Cashflow, Time Slider et Violet Zebra
│   ├── MOD_05_Advanced_Modules_v1.0.bas # Bilan Patrimonial et snapshot cumulatif
│   ├── MOD_06_Budget_ZBB_v1.0.bas     # Pilotage budgétaire ZBB et DataBars
│   ├── MOD_99_SystemAdmin_v1.0.bas    # Procédures d'administration (Soft/Hard Reset)
│   ├── M99_Sales_KeyGen_V1.0.bas      # Générateur de clés
│   ├── GFMP_v1.0.cls                  # Classe principale
│   ├── USF_Transaction_V1.0.frm/.frx  # Formulaire Gatekeeper Saisie
│   └── USF_Budget_V1.0.frm/.frx       # Formulaire Allocation Budgétaire
│
├── docs/                          # Spécifications et manuels
│   ├── SFP_SPEC_V1.0.docx         # Spécifications d'ingénierie d'entreprise SFP v1.0
│   ├── INSTALL_ANDROID_GUIDE.md   # Guide d'installation sideloading APK & PWA
│   └── Keya-Raymond-TOP_10.pdf    # Documentation exécutive
│
├── licenses/                      # Gestion et génération de licences
│   ├── KeyFile.xlsm
│   └── SFP_Generateur_Licences.xlsm
│
└── dist/                          # Livrables exécutables prêts au déploiement
    └── GFMP-v1.0.apk              # Package Android signé autonome (2.46 Mo)
```

---

## ⚙️ Modules & Fonctionnalités Clés

### 1. 🏛️ Hub Central (`APP_HOME` / Vue Accueil Mobile)
- Navigation inspirée des *Single Page Applications* (SPA).
- **4 Solid Cards massives** : Saisie Transaction, Dashboard Cashflow, Bilan Patrimonial, Pilotage Budgétaire.
- **Matrice i18n auto-cicatrisante** : Support de 8 langues (Français, Anglais, Espagnol, Portugais, Allemand, Italien, Néerlandais, Suédois).

### 2. 💳 Le Gatekeeper (`USF_Transaction` — MOD_03)
- Saisie sécurisée des opérations (Revenu, Dépense, Transfert).
- **Auto-apprentissage immédiat** : Détection de l'option *"Autre (Préciser...)"* pour créer à la volée de nouveaux comptes, catégories ou tiers.
- **Intégrité absolue** : Nettoyage regex (`CLEAN_TEXT`), formatage automatique, protection anti-doublon en mémoire.

### 3. 📊 Dashboard Cashflow (MOD_04)
- **Time Slider vectoriel** : Navigation temporelle fluide par Année et par Mois (ou Année Entière).
- **Multi-Devises $O(1)$** : Conversion en temps réel entre `MUR` (Roupie Mauricienne), `EUR` (€), `USD` ($), `GBP` (£), `ZAR` (R) et `XOF` (CFA).
- **3 Cartes KPI** : Total Revenus (Bleu Royal), Total Dépenses (Violet Royal), Cashflow Net (Vert/Rouge conditionnel).
- **Tableau Violet Zebra** : En-tête Violet Sombre (`#5A3282`), alternance Violet Royal (`#7851A9`) et Violet Doux (`#916EBE`).

### 4. 💰 Bilan Patrimonial & Net Worth (MOD_05)
- **Mode Snapshot Temporel ("As of date")** : Photographie patrimoniale cumulative des soldes bancaires à la date choisie.
- **Ventilation par classe d'actifs** : Liquidités, Investissements, et **mise en avant des Dettes en Jaune Royal**.
- **Hero Card dynamique** : Affichage de la Valeur Nette totale (Actifs − Dettes).

### 5. 🎯 Pilotage Budgétaire ZBB (Zero-Based Budgeting — MOD_06)
- Suivi des enveloppes allouées (`T_FACT_Budget`) vs consommations réelles (`T_FACT_Transaction`).
- **DataBars Jaune Royal** : Jauges de consommation dorées intégrées dans le tableau avec alertes visuelles de dépassement.
- **Formulaire d'allocation rapide (+ ALLOUER)** avec logique UPSERT en mémoire.

---

## 🎨 Charte Graphique Stricte (SFP_SPEC_V1.0.docx)

- **Bleu Royal** : `RGB(65, 105, 225)` (`#4169E1`)
- **Jaune Royal** : `RGB(250, 218, 94)` (`#FADA5E`)
- **Violet Royal** : `RGB(120, 81, 169)` (`#7851A9`)
- **Violet Sombre (Zebra Header)** : `RGB(90, 50, 130)` (`#5A3282`)
- **Violet Doux (Zebra Row)** : `RGB(145, 110, 190)` (`#916EBE`)
- **Vert Émeraude** : `RGB(46, 204, 113)` (`#2ECC71`)
- **Rouge Alerte** : `RGB(231, 76, 60)` (`#E74C3C`)
- **Typographie** : ADLaM Display & Inter

---

## 🚀 Démarrage Rapide

### Déploiement Mobile Android (Sans Store)
1. Transférez [dist/GFMP-v1.0.apk](file:///c:/Users/KIMYAÉDITION/source/repos/GFMP/dist/GFMP-v1.0.apk) vers votre smartphone (câble USB, WhatsApp, Drive).
2. Touchez le fichier `.apk` pour l'installer (autorisez la source si demandé).
3. L'application **GFMP** s'exécute en 100% hors-ligne sans aucun serveur ni PC.

### Prévisualisation Web Locale & Wi-Fi
Dans un terminal PowerShell :
```powershell
.\run_mobile_preview.ps1
```
Accédez ensuite à l'adresse indiquée sur votre navigateur PC ou smartphone connecté au même réseau Wi-Fi.

### Utilisation du Classeur Maître Excel
Ouvrez [GFMP_V1.0.xlsm](file:///c:/Users/KIMYAÉDITION/source/repos/GFMP/GFMP_V1.0.xlsm) avec Microsoft Excel (macros activées). L'application s'initialise automatiquement sur le Hub Central.