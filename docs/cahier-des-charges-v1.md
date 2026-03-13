# Cahier des charges v1 - Outil d'aide a la maintenance routiere

## 1. Contexte

Le client (Port de Douala-Bonaberi) souhaite une application d'aide a la decision pour la maintenance routiere.
Les donnees de reference sont issues du fichier Excel `programme ayissi.xlsx` et d'un memo fonctionnel.

Objectif principal: assister rapidement un decideur pour choisir une action de maintenance sur une voie donnee.

## 2. Objectifs metier

1. Centraliser le catalogue des voies du domaine portuaire.
2. Associer chaque voie a son secteur (SAP), son lineaire et ses bornes (pk debut / pk fin ou debut/fin).
3. Permettre la selection d'une degradation et afficher automatiquement:
   - causes probables,
   - solution(s) de maintenance recommandee(s),
   - recommandations assainissement/caniveaux.
4. Fournir un premier MVP hors ligne, simple d'utilisation.

## 3. Perimetre fonctionnel MVP

### 3.1 Donnees a couvrir

1. Reseau viaire: 39 voies environ (SAP1, SAP2, SAP3, SAP4).
2. Informations voie: code, designation, debut, fin, longueur, type de revetement, etat chaussee, etat assainissement.
3. Catalogue des degradations (demarrage avec 3 degradations prioritaires).
4. Causes probables et solutions de maintenance associees.

### 3.2 Fonctions MVP

1. Ecran liste des voies:
   - filtre par SAP,
   - recherche par code/nom.
2. Ecran fiche voie:
   - informations geometriques et etat,
   - resume assainissement.
3. Ecran aide a la decision:
   - choix degradation,
   - affichage automatique causes probables,
   - affichage automatique solution(s) de maintenance,
   - suggestion de curage/entretien caniveaux.
4. Ecran evaluation rapide:
   - saisie valeur de deflexion D,
   - niveau d'etat (faible/moyen/fort/tres fort),
   - type d'intervention recommande.
5. Extension progressive du catalogue degradations:
   - ajout progressif de toutes les degradations restantes.
6. Evolution du moteur de regles:
   - moteur de regles enrichi par vagues successives.
7. Suivi et restitution:
   - modules de suivi, historique et reporting (niveau de base Lot 1).

### 3.3 Hors perimetre MVP

1. SIG cartographique avance (cartes interactives detaillees).
2. Planification budgetaire multi-annees complete.
3. Mobile natif Android/iOS.
4. Workflow d'approbation complexe multi-direction.

### 3.4 Socle metier non negociable (respect du besoin client)

1. L'application est utilisable hors ligne.
2. Le catalogue complet des voies est consultable (SAP, pk debut, pk fin/debut-fin).
3. L'utilisateur peut selectionner une voie.
4. L'utilisateur peut selectionner une degradation.
5. L'application affiche automatiquement:
   - causes probables,
   - solution(s) de maintenance,
   - recommandation assainissement (curage/caniveaux).
6. Le demarrage se fait avec 3 degradations prioritaires, extensible ensuite sans refonte de l'application.

## 4. Profils utilisateurs

1. Administrateur:
   - importe et met a jour les donnees,
   - gere catalogue degradations/causes/solutions.
2. Decideur maintenance:
   - consulte une voie,
   - lance l'evaluation,
   - obtient la recommandation.
3. Operateur terrain (lecture/saisie simple):
   - enregistre observation (option MVP+).

## 5. Regles metier v1

1. Une voie appartient a un seul SAP.
2. Une degradation peut avoir plusieurs causes probables.
3. Une degradation peut avoir plusieurs solutions de maintenance.
4. Une observation est rattachee a une voie et a une degradation.
5. La recommandation est calculee a partir:
   - degradation selectionnee,
   - gravite/etat,
   - valeur de deflexion D (si disponible),
   - etat d'assainissement.
6. Au MVP, la logique de recommandation est basee sur des regles deterministes (pas de modele IA).

## 6. Exigences non fonctionnelles

1. Mode hors ligne obligatoire (pas d'internet requis).
2. Temps de reponse cible: < 2 secondes sur poste standard.
3. Interface simple, lisible, orientee operationnel.
4. Export de resultats CSV inclus au Lot 1 (PDF en option MVP+).
5. Journalisation locale minimale (audit: qui a consulte/cree une evaluation).

## 7. Criteres d'acceptation MVP

1. Le systeme charge le referentiel de voies et les affiche par SAP.
2. L'utilisateur selectionne une voie et voit ses informations essentielles.
3. L'utilisateur selectionne une degradation et obtient:
   - causes probables,
   - solution(s) de maintenance.
4. L'application propose une recommendation assainissement si caniveaux obstrues ou etat mauvais.
5. L'application fonctionne totalement hors ligne.
6. Les 3 degradations prioritaires sont operationnelles de bout en bout.

## 8. Donnees a valider avec le client avant sprint de dev

1. Liste exacte des 3 degradations a inclure au MVP.
2. Nomenclature definitive des etats (Bon, Moyen, Mauvais, A determiner).
3. Format attendu pour pk debut/pk fin quand la source est textuelle.
4. Regles de priorisation quand plusieurs solutions sont possibles.

## 9. Roadmap proposee

1. Semaine 1: nettoyage donnees + schema base + import initial.
2. Semaine 2: ecrans catalogue/fiche voie + moteur de recommandation v1.
3. Semaine 3: ecran evaluation, tests metier, livraison MVP.

## 10. Technologies retenues (MVP)

1. Application desktop hors ligne:
   - Electron (shell desktop Windows),
   - React + TypeScript (interface utilisateur).
2. Base de donnees locale:
   - SQLite (stockage embarque sur poste),
   - acces via `better-sqlite3` cote application.
3. Import et transformation des donnees:
   - scripts Node.js/TypeScript pour importer les donnees Excel,
   - normalisation des champs et gestion des valeurs manquantes.
4. Packaging et livraison:
   - build Windows installable (`.exe`) via Electron Builder.
5. Outils qualite:
   - ESLint + Prettier,
   - tests de regles metier prioritaires (unitaires).

### 10.1 Hebergement et deploiement de l'application

1. Lot 1 (MVP) est une application desktop hors ligne: pas d'hebergement web/cloud requis.
2. L'application est installee sur un ou plusieurs postes Windows via un installateur (`.exe`).
3. Les donnees sont stockees localement sur chaque poste dans une base SQLite.
4. Aucun serveur internet n'est necessaire pour l'utilisation quotidienne.
5. Les mises a jour se font par livraison d'une nouvelle version installable.
6. Les sauvegardes de la base SQLite sont faites en local (copie vers disque externe ou dossier partage interne).

## 11. Offre commerciale retenue (Lot 1)

Devise: FCFA, estimation hors TVA.

1. Montant forfaitaire Lot 1: 650 000 FCFA.
2. Cette offre couvre le socle metier non negociable defini en section 3.4.
3. Cette offre couvre un MVP lite: priorite au coeur metier, reduction des fonctions non essentielles.
4. Le Lot 1 inclut aussi:
   - ajout progressif de toutes les degradations restantes,
   - moteur de regles enrichi,
   - modules de suivi, historique et reporting (version de base).

## 12. Delai de livraison Lot 1

Hypothese: validation des 3 degradations prioritaires par le client au demarrage.

1. Cadrage final et validation echantillon data: 1 jour ouvre.
2. Setup technique + base SQLite + import initial: 2 jours ouvres.
3. Ecrans catalogue/fiches voies + aide a la decision: 3 jours ouvres.
4. Tests fonctionnels + corrections + livraison: 2 jours ouvres.

Delai total cible Lot 1: 8 jours ouvres (entre 7 et 10 jours ouvres selon retours recette).

## 13. Detail du cout Lot 1 (650 000 FCFA)

1. Cadrage final + specification executable: 80 000 FCFA.
2. Preparation et normalisation donnees initiales: 90 000 FCFA.
3. Developpement application desktop offline (coeur metier): 360 000 FCFA.
4. Tests, recette, packaging installable Windows: 80 000 FCFA.
5. Mini formation passation (1 session): 40 000 FCFA.

Total: 650 000 FCFA.

## 14. Authentification et PIN admin

1. Pas de login
2. Gestion d'un seul compte

## 15. Modalites de paiement proposees (Lot 1)

1. 40% a la commande (210 000 FCFA) au demarrage du projet.
2. 40% a la livraison de la version fonctionnelle pre-recette (210 000 FCFA) a convenir.
3. 20% a la recette finale et passation (110 000 FCFA).

## 16. Exclusions Lot 1 et options futures

1. Le Lot 1 n'inclut pas:
   - workflow de validation complexe,
   - application mobile,
   - cartographie SIG avancee.

2. Extension Phase 2 (optionnelle):
   - reporting avance (tableaux de bord decisionnels),
   - workflow de validation multi-acteurs,
   - synchronisation multi-postes / multi-sites.
