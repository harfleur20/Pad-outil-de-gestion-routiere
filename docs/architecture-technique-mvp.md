# Architecture technique MVP - PAD maintenance routiere

## 1. Choix de stack (recommande)
1. Desktop offline: Electron.
2. Interface: React + TypeScript.
3. Base locale: SQLite.
4. Acces DB: Prisma ou better-sqlite3 (simple et rapide).
5. UI: composant sobre et lisible (tableaux, filtres, formulaires).

Pourquoi ce choix:
1. Fonctionne sans internet.
2. Installation simple sur postes Windows.
3. Bon compromis rapidite de developpement / robustesse.

## 2. Macro architecture
1. `Electron Main`:
   - gestion fenetre,
   - acces filesystem,
   - acces SQLite,
   - API locale (IPC securise).
2. `Renderer React`:
   - pages metier (catalogue, fiche voie, evaluation),
   - logique UI et formulaires.
3. `Rule Engine v1`:
   - regles deterministes sur degradation, severite, deflexion, drainage.
4. `Import module`:
   - lecture Excel,
   - normalisation,
   - insertion base locale.

## 3. Structure de projet proposee
```text
projet-PAD/
  docs/
    cahier-des-charges-v1.md
    modele-de-donnees-v1.md
    schema-sqlite-v1.sql
    architecture-technique-mvp.md
  app/
    electron/
      src/
        main/
          main.ts
          db/
            client.ts
            migrations/
          services/
            road.service.ts
            recommendation.service.ts
            import.service.ts
          ipc/
            road.ipc.ts
            evaluation.ipc.ts
    web/
      src/
        pages/
          RoadsPage.tsx
          RoadDetailPage.tsx
          EvaluationPage.tsx
        components/
          RoadTable.tsx
          DegradationSelector.tsx
          RecommendationPanel.tsx
        store/
          useRoadStore.ts
          useEvaluationStore.ts
```

## 4. Flux applicatif MVP
1. L'utilisateur ouvre l'app.
2. L'app charge les secteurs SAP et voies depuis SQLite.
3. L'utilisateur choisit une voie.
4. L'app affiche la fiche voie (etat chaussee + drainage).
5. L'utilisateur choisit une degradation.
6. Le moteur de regles renvoie:
   - causes probables,
   - solution principale,
   - action assainissement.
7. L'utilisateur peut enregistrer l'observation et la recommandation.

## 5. API locale (IPC) minimale
1. `road:list({ sapCode, search })`
2. `road:getById({ id })`
3. `degradation:listActive()`
4. `evaluation:compute({ roadId, degradationId, deflectionValue, severityLevel })`
5. `observation:create(payload)`

## 6. Regles de decision v1 (exemple)
1. Si `deflectionValue < 60`:
   - niveau: `FAIBLE`
   - action: `PAS D'ENTRETIEN / SURVEILLANCE`
2. Si `60 <= deflectionValue < 80`:
   - niveau: `MOYEN`
   - action: `RENFORCEMENT LEGER`
3. Si `80 <= deflectionValue < 90`:
   - niveau: `FORT`
   - action: `RENFORCEMENT LOURD`
4. Si `deflectionValue >= 90`:
   - niveau: `TRES FORT`
   - action: `REHABILITATION COUCHE DE ROULEMENT ET BASE`
5. Si `drainage_state` contient `Mau` ou `obstrue`:
   - ajouter action assainissement: `CURAGE / NETTOYAGE CANIVEAUX`.

## 7. Plan de mise en oeuvre
1. Sprint 0 (2-3 jours):
   - creation squelette Electron + React,
   - creation DB SQLite avec schema v1,
   - script import Excel initial.
2. Sprint 1 (5 jours):
   - ecran liste voies + fiche voie,
   - catalogue degradations actives (3 au depart).
3. Sprint 2 (5 jours):
   - ecran evaluation,
   - moteur de recommandation,
   - sauvegarde observation + recommandation.
4. Sprint 3 (3 jours):
   - stabilisation, tests fonctionnels, package installable Windows.

## 8. Risques et mitigation
1. Risque: donnees heterogenes (`A determiner`, `-`, valeurs manquantes).
   - mitigation: phase de normalisation et validation obligatoire avant import final.
2. Risque: ambiguite des regles metier par degradation.
   - mitigation: valider une table de regles simple avec le client avant dev.
3. Risque: extension rapide du catalogue.
   - mitigation: conception data-driven (ajout sans changer le code UI).
