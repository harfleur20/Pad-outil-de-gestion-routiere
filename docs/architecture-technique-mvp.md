# Architecture technique MVP - PAD maintenance routiere

## 1. Stack retenue
1. Desktop offline: Electron.
2. Interface: React + TypeScript (Vite).
3. Base locale: SQLite (`better-sqlite3`).
4. Import Excel: `xlsx`.

## 2. Macro architecture
1. `Electron Main`:
   - creation de la fenetre,
   - ouverture des dialogues systeme (selection fichier Excel),
   - handlers IPC,
   - orchestration du `data-layer` SQLite.
2. `Data layer` (`app/electron/services/data-layer.cjs`):
   - schema SQLite,
   - import des 7 feuilles Excel,
   - CRUD feuille par feuille,
   - construction catalogue voies/degradations,
   - moteur de decision,
   - historique des decisions.
3. `Preload`:
   - exposition d'une API securisee `window.padApp`.
4. `Renderer React` (`app/web/src/App.tsx`):
   - onglets metier,
   - formulaire d'aide a la decision,
   - catalogues,
   - historique + export CSV,
   - ecrans CRUD `Feuil1..Feuil7`.

## 3. Perimetre securite MVP
1. Application locale offline.
2. Pas de login/PIN au MVP (demande utilisateur en cours de projet).
3. Les fonctions de protection avancee (roles, audit utilisateur, pin admin) sont replanifiees en phase 2.

## 4. Flux applicatif MVP
1. Au demarrage, l'application ouvre la base SQLite locale.
2. Si la base est vide, import automatique du fichier Excel detecte.
3. L'utilisateur peut:
   - importer un fichier Excel,
   - naviguer entre `Aide decision`, `Catalogue`, `Degradations`, `Historique`, `Feuil1..Feuil7`.
4. Le calcul de decision retourne:
   - cause probable,
   - solution de maintenance,
   - recommandation deflexion,
   - recommandation assainissement.
5. Chaque decision est enregistree dans `decision_history`.
6. Le module historique permet filtrage, export CSV et purge complete.

## 5. API IPC actuellement exposee
1. Data:
   - `data:status()`
   - `data:importFromExcel(excelPath?)`
   - `data:pickExcelFile()`
2. Feuilles Excel (CRUD):
   - `sheet:definitions()`
   - `sheet:list(sheetName, filters)`
   - `sheet:create(sheetName, payload)`
   - `sheet:update(sheetName, rowId, payload)`
   - `sheet:delete(sheetName, rowId)`
3. Metier:
   - `sap:list()`
   - `roads:list({ sapCode, search })`
   - `degradations:list()`
   - `decision:evaluate({ roadId, degradationId, deflectionValue, askDrainage })`
4. Reporting:
   - `reporting:listHistory({ sapCode, search, limit })`
   - `reporting:clearHistory()`

## 6. Regles de decision v1 (deflexion)
1. `D < 60` -> `FAIBLE` -> `PAS D'ENTRETIEN`.
2. `60 <= D < 80` -> `MOYEN` -> `RENFORCEMENT LEGER`.
3. `80 <= D < 90` -> `FORT` -> `RENFORCEMENT LOURD`.
4. `D >= 90` -> `TRES FORT` -> `REHABILITATION COUCHE DE ROULEMENT ET DE BASE`.
5. Si assainissement degrade (obstrue, mauvais, non fonctionnel, etc.):
   - action prioritaire de curage/nettoyage caniveaux.

## 7. Mapping Excel (7 feuilles)
1. `Feuil1`: mesures de deflexion.
2. `Feuil2`: listing des sections.
3. `Feuil3`: etat chaussee + intervention.
4. `Feuil4`: programme d'evaluation / donnees d'entree.
5. `Feuil5`: sections avec assainissement/stationnement.
6. `Feuil6`: repertoire codifie des voies.
7. `Feuil7`: degradations + causes.

## 8. Structure projet
```text
projet-PAD/
  app/
    electron/
      main.cjs
      preload.cjs
      services/
        data-layer.cjs
    web/
      src/
        App.tsx
        lib/pad-api.ts
        styles/theme.css
        types/pad-api.ts
  docs/
    cahier-des-charges-v1.md
    architecture-technique-mvp.md
    schema-sqlite-v1.sql
```

## 9. Risques restants
1. Homogeneite des donnees Excel (valeurs textuelles heterogenes).
2. Regles metier a enrichir pour l'ensemble des degradations.
3. Packaging Windows a stabiliser selon environnement cible client.
