# Mapping feuilles Excel -> modele metier PAD

## 1. Objectif
Ce document fixe la logique de correspondance entre les 7 feuilles Excel et le noyau metier actuel de l'application.

But:
- ne pas casser le systeme deja en place,
- garder les feuilles comme source d'entree,
- consolider les donnees dans des tables relationnelles stables,
- definir les priorites quand plusieurs feuilles parlent de la meme voie.

## 2. Tables metier actuelles
Le noyau SQLite deja en place s'appuie sur les tables suivantes:

| Table | Role |
| --- | --- |
| `sap_sector` | referentiel des SAP |
| `road` | referentiel unifie des voies |
| `road_section` | sections/troncons par feuille source |
| `road_measurement` | mesures Feuil1 par PK |
| `degradation` | catalogue unifie des degradations |
| `degradation_cause` | causes probables par degradation |
| `road_degradation` | lien voie <-> degradation |
| `degradation_definition` | metadonnees Feuil7 |
| `decision_profile_input` | donnees brutes Feuil4 |
| `decision_history` | historique des decisions |
| `maintenance_intervention` | historique des entretiens |
| `drainage_rule` | regles de decision assainissement |

## 3. Cle metier centrale
La cle centrale du systeme est la `voie`.

Une voie apparait sous plusieurs formes dans les feuilles:
- `code voie` (`Rue.01`, `Bvd.02`, `Av.01`),
- `voie` (`Rue 01`, `Bvd 02`),
- `designation` (`Rue des Archives`, `Boulevard des Portiques`),
- `debut` / `fin`,
- `SAP`.

La correspondance ne doit donc pas reposer sur un seul libelle texte.

## 4. Regle de correspondance recommandee
### 4.1 Cle de voie cible
La voie doit etre resolue dans cet ordre:

1. `road_code` normalise.
2. `designation` normalisee.
3. `designation + debut + fin`.
4. `SAP + designation`.

### 4.2 Sources de verite par domaine
| Domaine metier | Source prioritaire | Source secondaire | Commentaire |
| --- | --- | --- | --- |
| Identite de la voie | `Feuil6` | `Feuil2`, `Feuil5`, `Feuil3` | `Feuil6` porte le code stable et le nom propose |
| SAP et sections | `Feuil2` | `Feuil5`, `Feuil6` | `Feuil2` porte le vrai listing structurel |
| Profil technique de chaussee | `Feuil3` | `Feuil5` | `Feuil3` porte l'etat metier et la nature d'intervention |
| Assainissement | `Feuil3` | `Feuil5` | `Feuil3` plus narratif, `Feuil5` plus structure |
| Stationnement | `Feuil5` | aucune | seulement present dans `Feuil5` |
| Mesures de deflexion | `Feuil1` | aucune | donnees techniques de mesure |
| Degradations et causes | `Feuil7` | aucune | catalogue metier de reference |
| Ecran d'evaluation | calcule depuis les autres | `Feuil4` comme source brute | `Feuil4` n'est pas une source de verite autonome |

## 5. Mapping feuille par feuille
## 5.1 Feuil1 - Mesures de deflexion
### Role metier
Campagne de mesure sur une voie ou un troncon, avec valeurs par PK.

### Colonnes utiles
- date de campagne,
- voie/troncon mesure,
- `PK` lecture,
- lecture comparateur `Gauche`, `Axe`, `Droit`,
- `PK` deflexion,
- deflexion `Gauche`, `Axe`, `Droit`,
- `Defl.Brute.Moy`,
- `ecart type`,
- `Deflexion caracteristique Dc`.

### Cible metier
- `road_measurement`
- `measurement_campaign`

### Dependances
- la voie doit d'abord etre resolue dans `road`
- les bornes `debut` / `fin` viennent de `Feuil2` / `Feuil6`

### Regle de correspondance
1. identifier la voie mesuree par code ou designation,
2. rattacher la mesure a `road.id`,
3. stocker chaque ligne PK dans `road_measurement`,
4. stocker la date de campagne au niveau d'un lot de mesure, pas sur chaque cellule brute.

### Point critique
Le schema doit porter explicitement:
- la date de campagne,
- un identifiant de campagne,
- le nom complet du troncon mesure.

Cette couche doit vivre dans `measurement_campaign`, avec rattachement fin vers `road_measurement`.

## 5.2 Feuil2 - Listing des sections
### Role metier
Referentiel structurel des sections par SAP.

### Cible metier
- `road_section`
- `road`
- `sap_sector`

### Donnees portees
- numero troncon,
- numero section,
- voie,
- designation,
- debut,
- fin,
- longueur,
- regroupement SAP.

### Priorite
`Feuil2` est la source principale pour:
- `sap_code`,
- `debut`,
- `fin`,
- `longueur`,
- structure par troncon/section.

## 5.3 Feuil3 - Etat chaussee et intervention
### Role metier
Diagnostic metier par voie.

### Cible metier
- `road`
- `road_section`

### Donnees portees
- largeur,
- nature du revetement,
- etat de la chaussee,
- type de caniveaux,
- description assainissement,
- largeur minimale trottoirs,
- nature de l'intervention.

### Priorite
`Feuil3` est la source principale pour:
- `surface_type`,
- `pavement_state`,
- `drainage_type`,
- `drainage_state`,
- `intervention_hint`.

### Interpretation
La colonne `Nature de l'intervention` n'est pas une execution d'entretien.
Elle sert de recommandation ou d'orientation contextuelle par voie.
L'execution reelle d'un entretien doit rester dans `maintenance_intervention`.

## 5.4 Feuil4 - Donnees d'entree / evaluation
### Role metier
Vue d'entree et de simulation pour l'aide a la decision.

### Cible metier
- `decision_profile_input` pour les valeurs brutes
- alimentation indirecte du moteur de decision

### Dependances
`Feuil4` consomme:
- la voie depuis `road`,
- le SAP depuis `sap_sector`,
- les bornes depuis `road`,
- la degradation depuis `degradation`,
- la cause depuis `degradation_cause`,
- la valeur `D` depuis l'utilisateur ou une mesure `Feuil1`,
- la recommandation via les regles de decision.

### Regle de lecture
`Feuil4` n'est pas une source de verite metier.
C'est une projection des autres feuilles et du moteur.

## 5.5 Feuil5 - Sections avec assainissement et stationnement
### Role metier
Profil de section plus riche que `Feuil2`, avec assainissement et stationnement.

### Cible metier
- `road`
- `road_section`

### Donnees portees
- largeur min. facade,
- nature du revetement,
- etat de la chaussee,
- type assainissement,
- etat assainissement,
- largeur trottoirs,
- stationnement gauche,
- stationnement droit,
- autre stationnement.

### Priorite
`Feuil5` complete:
- `width_m`,
- `sidewalk_min_m`,
- `parking_left`,
- `parking_right`,
- `parking_other`.

### Regle de conflit avec Feuil3
- `etat de la chaussee`: priorite `Feuil3`
- `assainissement`: priorite `Feuil3`, fallback `Feuil5`
- `stationnement`: priorite `Feuil5`

## 5.6 Feuil6 - Repertoire codifie des voies
### Role metier
Referentiel de nommage et de codification des voies.

### Cible metier
- `road`
- `sap_sector` comme fallback de regroupement

### Donnees portees
- type de voie,
- code voie,
- lineaire,
- nom propose,
- itineraire,
- justification.

### Priorite
`Feuil6` est la source principale pour:
- `road_code`,
- `designation` cible affichee,
- `itinerary`,
- `justification`.

### Point important
`Feuil6` doit devenir la base des alias de voies.
Exemples:
- `Rue 02` <-> `Rue.02`
- `Bvd 03` <-> `Bvd.03`
- `Rue des Cimenteries` <-> `Rue de la Cimenterie`

## 5.7 Feuil7 - Degradations et causes
### Role metier
Catalogue metier des degradations.

### Cible metier
- `degradation`
- `degradation_cause`
- `degradation_definition`
- `road_degradation`

### Donnees portees
- categorie,
- reference,
- degradation,
- famille,
- sous-famille,
- notes,
- cause probable.

### Priorite
`Feuil7` est la source exclusive pour:
- le libelle de degradation,
- les causes probables,
- les metadonnees de classement.

## 6. Priorites de fusion recommandees
Quand plusieurs feuilles portent la meme information, la priorite cible doit etre la suivante.

| Champ cible | Priorite 1 | Priorite 2 | Priorite 3 |
| --- | --- | --- | --- |
| `road_code` | `Feuil6` | `Feuil2` / `Feuil5` | `Feuil3` |
| `designation` | `Feuil6` | `Feuil2` / `Feuil5` | `Feuil3` |
| `sap_code` | `Feuil2` | `Feuil5` | `Feuil6` |
| `start_label` | `Feuil2` | `Feuil3` / `Feuil5` | `Feuil6.itinerary` |
| `end_label` | `Feuil2` | `Feuil3` / `Feuil5` | `Feuil6.itinerary` |
| `length_m` | `Feuil2` | `Feuil5` | `Feuil6` |
| `width_m` | `Feuil3` | `Feuil5` | aucune |
| `surface_type` | `Feuil3` | `Feuil5` | aucune |
| `pavement_state` | `Feuil3` | `Feuil5` | aucune |
| `drainage_type` | `Feuil3` | `Feuil5` | aucune |
| `drainage_state` | `Feuil3` | `Feuil5` | aucune |
| `sidewalk_min_m` | `Feuil5` | `Feuil3` | aucune |
| `parking_*` | `Feuil5` | aucune | aucune |
| `intervention_hint` | `Feuil3` | valeur par defaut | aucune |

## 7. Ce que cela implique pour Feuil1
La vraie mise a jour de `Feuil1` ne doit pas juste renommer des colonnes.
Elle doit faire trois choses:

1. reconnaitre la voie mesuree a partir du referentiel `road`,
2. relier chaque ligne de mesure a cette voie et a sa section,
3. permettre a `Feuil4` et au moteur de decision de consommer la bonne campagne de mesure.

## 8. Ajustements systeme a prevoir
Avant de finaliser `Feuil1`, les ajustements suivants sont recommandes:

1. Ajouter une notion de `measurement_campaign`
- `campaign_key`
- `road_id`
- `measurement_date`
- `source_sheet`
- `source_header_label`

2. Ajouter une table ou une logique d'alias de voies
- pour absorber les variantes `Rue 02` / `Rue.02`
- pour absorber les variantes de designation

Statut actuel:
- `measurement_campaign` implemente dans le noyau SQLite
- `road_alias` implemente dans le noyau SQLite

3. Rendre explicite la priorite de fusion
- aujourd'hui le code fusionne deja les feuilles,
- mais cette priorite doit etre stabilisee et documentee, comme dans ce document.

4. Faire lire `Feuil4` comme une vue dynamique
- valeurs d'entree utilisateur,
- causes issues de `Feuil7`,
- niveau `D` issu soit de l'utilisateur soit de `Feuil1`,
- intervention issue du moteur et du contexte voie.

## 9. Decision de conception recommandee
La conception cible doit etre:

1. `Feuil6` = identite de la voie.
2. `Feuil2` = structure SAP / debut / fin / longueur.
3. `Feuil3` = etat metier de reference.
4. `Feuil5` = completement technique complementaire.
5. `Feuil1` = mesures techniques par campagne.
6. `Feuil7` = catalogue des degradations.
7. `Feuil4` = interface de calcul, pas referentiel autonome.

## 10. Conclusion
Le noyau actuel va dans la bonne direction.

La suite logique n'est pas de reconstruire tout le projet.
La suite logique est:

1. figer cette matrice de correspondance,
2. ajuster le modele de `Feuil1` autour d'une campagne de mesure,
3. renforcer la resolution de la voie via `Feuil6` + alias,
4. laisser `Feuil4` comme vue de decision alimentee par les autres feuilles.
