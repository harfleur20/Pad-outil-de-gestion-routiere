# Modele de donnees v1 - Maintenance routiere PAD

## 1. Objectif
Definir un modele relationnel simple pour:
1. stocker le referentiel des voies,
2. stocker degradations/causes/solutions,
3. enregistrer observations et recommandations.

## 2. Entites principales

### 2.1 `sap_sector`
Represente un secteur d'activite portuaire.

Champs:
1. `id` (PK)
2. `code` (unique, ex: SAP1)
3. `name`
4. `description`
5. `created_at`
6. `updated_at`

### 2.2 `road`
Represente une voie du reseau.

Champs:
1. `id` (PK)
2. `sap_sector_id` (FK -> sap_sector.id)
3. `road_code` (ex: Rue.07, Bvd.02)
4. `road_type` (Rue, Boulevard, Avenue)
5. `name`
6. `start_label`
7. `end_label`
8. `length_m`
9. `width_m`
10. `surface_type` (BB, Mixte, Paves, BB/Paves...)
11. `pavement_state` (Bon, Moy, Mau, AD...)
12. `drainage_type`
13. `drainage_state`
14. `sidewalk_left_m`
15. `sidewalk_right_m`
16. `parking_left`
17. `parking_right`
18. `notes`
19. `is_active`
20. `created_at`
21. `updated_at`

Contraintes:
1. (`sap_sector_id`, `road_code`) unique.

### 2.3 `degradation`
Catalogue des degradations.

Champs:
1. `id` (PK)
2. `code` (unique, ex: DEG_FISSURE_TR)
3. `name` (ex: FISSURES TRANSVERSALES)
4. `description`
5. `is_active`
6. `created_at`
7. `updated_at`

### 2.4 `degradation_cause`
Liste des causes probables d'une degradation.

Champs:
1. `id` (PK)
2. `degradation_id` (FK -> degradation.id)
3. `cause_text`
4. `weight` (0..100, optionnel)
5. `is_active`
6. `created_at`
7. `updated_at`

### 2.5 `maintenance_solution`
Solutions possibles associees a une degradation.

Champs:
1. `id` (PK)
2. `degradation_id` (FK -> degradation.id)
3. `title`
4. `description`
5. `intervention_level` (leger, moyen, lourd)
6. `estimated_duration_days`
7. `is_active`
8. `created_at`
9. `updated_at`

### 2.6 `road_observation`
Observation ponctuelle sur une voie.

Champs:
1. `id` (PK)
2. `road_id` (FK -> road.id)
3. `degradation_id` (FK -> degradation.id)
4. `observed_at`
5. `observed_by`
6. `severity_level` (faible, moyen, fort, tres_fort)
7. `deflection_value`
8. `pavement_note`
9. `drainage_note`
10. `attachments_json`
11. `created_at`

### 2.7 `recommendation`
Resultat genere par le moteur de regles.

Champs:
1. `id` (PK)
2. `observation_id` (FK -> road_observation.id)
3. `primary_cause_id` (FK -> degradation_cause.id, nullable)
4. `primary_solution_id` (FK -> maintenance_solution.id, nullable)
5. `recommended_action`
6. `assainissement_action`
7. `confidence_score` (0..1)
8. `rule_version`
9. `created_at`

## 3. Relations
1. Un `sap_sector` a plusieurs `road`.
2. Une `degradation` a plusieurs `degradation_cause`.
3. Une `degradation` a plusieurs `maintenance_solution`.
4. Une `road_observation` concerne une `road` et une `degradation`.
5. Une `recommendation` est rattachee a une `road_observation`.

## 4. Mapping minimum depuis les feuilles Excel
1. `Feuil6` -> `sap_sector`, `road` (referentiel codifie).
2. `Feuil2` / `Feuil5` -> enrichissement attributs de `road`.
3. `Feuil7` -> `degradation_cause` (et initialisation `degradation`).
4. `Feuil3` -> pre-remplissage `maintenance_solution` / textes d'intervention.
5. `Feuil4` -> regles de decision pour le moteur v1 (deflexion D).

## 5. Donnees cibles MVP
1. Charger toutes les voies.
2. Activer seulement 3 degradations au demarrage.
3. Associer causes et solutions pour ces 3 degradations.
4. Conserver le reste du catalogue en statut inactif pour extension.
