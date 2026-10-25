# Guide d'Installation Android Sans Play Store — GFMP (SFP v1.0)

Ce guide vous explique comment installer et utiliser la plateforme financière **GFMP (Global Finance Management Platform / Système Financier Personnel SFP v1.0)** sur votre smartphone Android, **sans passer par Google Play Store** ni aucun magasin d'applications, et en **totale autonomie (zéro serveur, zéro PC requis)**.

---

## Deux Méthodes d'Installation au Choix

| Méthode | Format | Avantages | Procédure |
| :--- | :---: | :--- | :--- |
| **Méthode 1 (Recommandée)** | **APK Natif (.apk)** | • Véritable application Android native (`GFMP`)<br>• Accélération GPU & icône dédiée sur l'écran d'accueil<br>• Mode plein écran immersif sans barre d'adresse<br>• Contrôle matériel du bouton retour Android | Transférer le fichier `GFMP-v1.0.apk` sur le téléphone et cliquer dessus pour installer. |
| **Méthode 2 (Instantanée)** | **PWA Autonome** | • Zéro fichier APK à gérer<br>• S'installe via Chrome Mobile / Samsung Internet<br>• Cache 100% hors-ligne via Service Worker (`sw.js`) | Ouvrir `http://<IP_PC>:8089` sur le mobile et faire *"Ajouter à l'écran d'accueil"*. |

---

## Méthode 1 : Installation par APK (Sideloading Direct)

### Étape 1 : Transférer l'APK sur votre smartphone
Copiez le fichier **`GFMP-v1.0.apk`** (2.46 Mo) sur votre téléphone par l'un des moyens suivants :
- **Par câble USB** (copier dans le dossier *Téléchargements* ou *Download*).
- **Par WhatsApp / Telegram / Signal** (envoi direct du fichier en pièce jointe à vous-même ou à un contact).
- **Par Email ou Google Drive / OneDrive / Bluetooth**.

### Étape 2 : Autoriser l'installation (sources inconnues)
1. Sur votre smartphone Android, ouvrez votre gestionnaire de fichiers ou WhatsApp.
2. Touchez le fichier **`GFMP-v1.0.apk`**.
3. Si Android affiche : *"Pour des raisons de sécurité, votre téléphone n'est pas autorisé à installer des applications inconnues depuis cette source"* :
   - Cliquez sur **Paramètres**.
   - Activez le curseur **"Autoriser cette source"**.
   - Revenez en arrière et cliquez sur **"Installer"**.

### Étape 3 : Lancer l'application
L'icône **GFMP** apparaît sur votre écran d'accueil.
Touchez l'icône : l'application s'ouvre instantanément en **100% hors-ligne**, sans aucun écran de chargement serveur !

---

## Modules GFMP & Fonctionnalités Disponibles sur Téléphone

1. **🏛️ Hub Central (APP_HOME)** :
   - 4 Solid Cards massives : **Saisie Transaction**, **Dashboard Cashflow**, **Bilan Patrimonial**, **Pilotage Budgétaire**.
   - Sélecteur 8 langues : Français (FR), Anglais (EN), Espagnol (ES), Portugais (PT), Allemand (DE), Italien (IT), Néerlandais (NL), Suédois (SV).

2. **💳 Saisie Transaction (Gatekeeper - USF_Transaction)** :
   - Enregistrement immédiat de revenus, dépenses et virements avec assainissement regex et anti-doublon.
   - Auto-apprentissage instantané : possibilité d'ajouter un nouveau compte, une catégorie ou un tiers directement dans le formulaire.

3. **📊 Dashboard Cashflow (MOD_04)** :
   - Time Slider interactif avec sélecteur d'Année et de Mois (ou Année Entière).
   - Multi-devises O(1) : Roupie Mauricienne (`MUR`), Euro (`EUR`), Dollar (`USD`), Livre (`GBP`), Rand (`ZAR`), Franc CFA (`XOF`).
   - 3 KPIs : Total Revenus (Bleu Royal), Total Dépenses (Violet Royal), Cashflow Net (Vert si positif, Rouge si négatif).
   - Tableau "Violet Zebra" conforme aux spécifications SFP v1.0.

4. **💰 Bilan Patrimonial / Net Worth (MOD_05)** :
   - Photographie patrimoniale "As of Date" (calcul cumulatif à la date sélectionnée).
   - Inventaire des comptes par classe d'actifs : Liquidités, Investissements, et mise en évidence des Dettes en Jaune Royal.
   - KPI Valeur Nette dynamique.

5. **🎯 Pilotage Budgétaire ZBB (MOD_06)** :
   - Comparaison en temps réel des enveloppes allouées (`T_FACT_Budget`) et des dépenses réelles (`T_FACT_Transaction`).
   - Bouton d'allocation rapide (+ ALLOUER BUDGET) avec logique UPSERT.
   - Jauges de consommation dorées (DataBars Jaune Royal) et alertes de dépassement.

6. **💾 Sauvegarde & Souveraineté Totale (Zéro Cloud Externe)** :
   - Boutons de Sauvegarde et Restauration JSON/Excel directement depuis le téléphone.
   - Vos données financières restent strictement confinées dans votre appareil.
