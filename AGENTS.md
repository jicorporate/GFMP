# Règles de Gouvernance et Standards d'Ingénierie GFMP

## 1. Sanctité Absolue des Binaires de Référence
- `GFMP-v1.0.apk` est la version stable scellée. Elle ne doit sous aucun prétexte être modifiée, recompilée, supprimée ou écrasée.
  - Taille requise : 3 083 929 octets
  - Empreinte SHA-256 : 4E29C86CDC7D18B2585849F17474F721D8185F7CC37A22F5FE00D1E27CC80DB9
- Tout cycle de développement concerne exclusivement la version de production `v1.01` (`GFMP-v1.01.apk`).

## 2. Clarification Active et Non-Supposition
- Si une demande utilisateur présente la moindre ambiguïté ou divergence d'interprétation, ne jamais supposer ni improviser : poser immédiatement les questions de clarification nécessaires avant d'engager des modifications.

## 3. Dynamisme Monétaire des Formulaires (Zéro Figeage en Dur)
- Les formulaires de saisie (`#tx-devise`, `#exp-tx-devise`, `#alloc-devise`, `#add-entity-devise`) doivent adopter dynamiquement la devise de base configurée par l'utilisateur (`SYS_DEVISE_BASE` / `engine.currency`) et la devise native du compte sélectionné.
- Interdiction formelle de figer `MUR` ou toute autre devise par défaut dans les sélecteurs.
- Respect absolu du choix ponctuel de l'utilisateur : s'il sélectionne une devise étrangère pour une transaction, cette sélection ne doit jamais être écrasée avant la validation.

## 4. Pureté de l'Interface Haute Joaillerie (Zéro Étiquette Parasite)
- Ne jamais ajouter d'étiquettes superflues, de badges de statut non requis, de tableaux de cours manuels ou de calculateurs de conversion dans l'application.
- Les conversions de devises doivent opérer à 100% en arrière-plan via l'API connectée, de manière fluide, dynamique et responsive.

