# Projet PAD - Outil d'aide a la maintenance routiere

Application desktop Electron + React + SQLite, utilisable hors ligne.

## Demarrage
1. Installer les dependances:
   `npm install`
2. Recompiler le module SQLite pour Electron (si necessaire):
   `npm run rebuild:electron`
3. Lancer l'application desktop:
   `npm run dev`
4. Verifier le typecheck + build web:
   `npm run build`
5. Generer l'installateur Windows:
   `npm run dist:win`

## Fonctionnalites disponibles
1. Import Excel (`programme ayissi.xlsx`) avec bouton `Parcourir`.
2. Ecran `Aide decision`:
   - selection d'une voie,
   - selection d'une degradation,
   - saisie de la deflexion D,
   - restitution automatique de la cause probable,
   - proposition de solution de maintenance,
   - recommandation assainissement/caniveaux,
   - enregistrement automatique dans l'historique.
3. Ecran `Catalogue` des voies (filtres SAP + recherche + action directe vers evaluation).
4. Ecran `Degradations` (causes detaillees + solution associee).
   - source de solution visible (`Modele`, `Personnalisee`, `A parametrer`),
   - affectation d'un modele de solution par degradation,
   - personnalisation editable de la solution par degradation.
5. Ecran `Historique` avec filtres et export CSV.
   - purge complete de l'historique avec bouton `Vider`.
6. Ecran `Suivi` des entretiens.
   - enregistrement d'une intervention par voie,
   - statut `Prevu`, `En cours`, `Termine`,
   - etat avant / apres, solution appliquee, prestataire, cout, observations,
   - historique dynamique par route.
7. Onglets `Feuil1..Feuil7` avec CRUD complet (lister, rechercher, ajouter, modifier, supprimer).
8. Stockage local SQLite (mode offline).
9. Pas de login/PIN au MVP (demarrage direct).

## Depannage
1. Erreur `Port 5173 is already in use`:
   - fermer le process Vite existant, puis relancer `npm run dev`.
2. Erreur `better_sqlite3.node ... compiled against a different Node.js version`:
   - executer `npm run rebuild:electron`,
   - puis relancer `npm run dev`.
   - si l'erreur arrive au lancement Electron, un message d'aide est affiche automatiquement.
3. Packaging Windows (`npm run dist:win`) en erreur de droits symlink:
   - relancer le terminal en mode administrateur, ou
   - activer le mode Developpeur Windows.

## Source de donnees Excel
Chemins auto-detectes au demarrage:
1. `PAD_EXCEL_PATH` (si defini)
2. `C:\Users\harfl\OneDrive\Desktop\pad\programme ayissi.xlsx`
3. `C:\Users\harfl\OneDrive\Desktop\programme ayissi.xlsx`
