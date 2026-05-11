# Solution digitalisée TRACE-PORT

## Objectif

Transformer le classeur mensuel des arrêts en application de pilotage pour le Poste de Commande de la Direction des Embarquements du Port de Casablanca. La solution reprend la logique du rapport PFE : suivi des anomalies, bilan des arrêts, performance manutention, coordination exploitation-maintenance, traçabilité SMQE et amélioration DMAIC.

## Périmètre repris depuis Excel

| Feuille Excel | Rôle dans la solution |
| --- | --- |
| `Bilan` | Base principale des arrêts : S/E, équipement, nature, début, fin, durée, description, affectation, qualité, destination. |
| `Synthèses` | Calculs mensuels par S/E et par famille d'arrêt : totaux, pourcentages, TRS exploitation, maintenance et global. |
| `Tonnage` | Tonnage pesage et draft par journée et par qualité. |
| `Trains` | Déchargement trains : nombre de trains, wagons, durée, tonnage, cadence, retards. |
| `Navire` | Chargement navires : navire, qualité, dates, bascule, connaissement, écart. |
| `Familles arrêts` | Référentiel des familles et exemples. |
| `EXPORTER` / `Feuil*` | Requêtes et références d'interventions importées dans le catalogue. |

## Modules fonctionnels

| Module | Besoin du rapport | Fonction digitale |
| --- | --- | --- |
| Tableau de bord | Suivi des KPI du processus manutention | Tonnage, cadence, TRS, TRG, Pareto, familles critiques. |
| Journal des arrêts | Suivi et saisie des anomalies et arrêts | Table filtrable par S/E, famille, qualité et recherche texte. |
| Saisie | Remplacer la saisie Excel manuelle | Formulaire reprenant les colonnes du `Bilan`, avec durée automatique. |
| Tonnage | Rapport journalier/mensuel d'activité | Pesage, draft, qualité, graphique journalier, écarts. |
| Trains & navires | Coordination flux réel installation/quai/trafic | Déchargement trains, chargement navires, bascule/connaissement. |
| Formules & requêtes | Garder la traçabilité des calculs Excel | Audit des 3 582 formules extraites et mapping vers le moteur digital. |
| Besoins PFE | Cadrage Define/Measure/Analyze | Matrice entre besoins métier, données, KPI et modules. |

## Moteur de calcul repris

| Indicateur | Formule Excel | Formule digitale |
| --- | --- | --- |
| Durée arrêt | `fin - début` | `(end - start) * 24` en heures. |
| Total par S/E et famille | `SUMIFS(Bilan!F:F, Bilan!A:A, S/E, Bilan!C:C, famille)` | Somme des durées filtrées par `sectionKey` et `family`. |
| Total arrêts | `SUM(D:AI)` | Somme des familles d'arrêts. |
| Maintenance | `SUM(électrique: bande)` | `électrique + instrumentation + mécanique + bande`. |
| TRS exploitation | `(durée affectation - arrêt exploitation) / durée affectation` | Même logique en heures. |
| TRS maintenance | `(durée affectation - arrêts maintenance) / durée affectation` | Même logique en heures. |
| TRS global | `(durée affectation - exploitation - maintenance) / durée affectation` | Même logique en heures. |
| Tonnage | `SUM(C:L)` et `SUM(P:Y)` | Somme des qualités pesage et draft. |
| Ecart navire | `(connaissement - bascule) / connaissement` | Ratio d'écart par navire et global. |

## Modèle de données cible

| Table | Champs principaux |
| --- | --- |
| `events_arrets` | `sectionKey`, `subEquipment`, `family`, `start`, `end`, `durationHours`, `description`, `assignment`, `quality`, `destination`. |
| `tonnage_daily` | `day`, `qualities`, `pesageTotal`, `draftTotal`. |
| `train_daily` | `trains`, `wagons`, `durationHours`, `totalTonnage`, `cadenceTph`, `delayHours`. |
| `ship_loading` | `name`, `quality`, `start`, `end`, `bascule`, `connaissement`, `gapRatio`. |
| `families` | `name`, `examples`. |
| `formula_audit` | `sheet`, `address`, `formula`, `cached`. |

## Recommandation d'évolution

Pour une version industrielle, la maquette peut évoluer vers :

- Backend `PostgreSQL` ou `SQL Server`.
- Authentification par rôle : Poste de Commande, Maintenance, Exploitation, HSE, Direction.
- Import Excel mensuel automatisé.
- Workflow de validation des arrêts : saisi, vérifié, validé, clôturé.
- Export PDF/Excel du rapport journalier, mensuel et annuel.
- Module actions DMAIC : cause racine, action corrective, responsable, échéance, gain.
