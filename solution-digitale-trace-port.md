# Solution digitalisée TRACE-PORT

## Objectif

Transformer le classeur mensuel des arrêts en application de pilotage pour le Poste de Commande de la Direction des Embarquements du Port de Casablanca. La solution reprend la logique du rapport PFE : suivi des anomalies, bilan des arrêts, performance manutention, coordination exploitation-maintenance, traçabilité SMQE et amélioration DMAIC.

## Présentation et vision cible

TRACE-PORT est conçu comme un système d'information digitalisé destiné à moderniser et centraliser la traçabilité des arrêts de manutention au sein de la Direction Logistique Portuaire de l'OCP Casablanca. La solution répond aux limites observées dans le système actuel : fichiers Excel dispersés, traitements manuels, double saisie papier/Excel, risques d'erreurs, pertes de temps et faible disponibilité des informations.

L'objectif principal est de transformer un processus manuel, lent et fragmenté en un système centralisé, automatisé et proche du temps réel. La plateforme doit permettre une meilleure qualité des données et un pilotage opérationnel plus réactif des activités portuaires.

La solution est pensée comme une plateforme accessible aux différents acteurs selon leurs rôles :

| Acteur | Rôle dans TRACE-PORT |
| --- | --- |
| Agent de quart | Saisie directe des arrêts depuis une interface simple avec contrôles automatiques. |
| Chef d'équipe | Vérification, correction et validation des informations avant exploitation. |
| Responsable exploitation / maintenance | Suivi des arrêts, performances des circuits, tendances et alertes critiques. |
| Direction | Consultation des KPI, synthèses et rapports pour la prise de décision. |

TRACE-PORT intègre une base de données centralisée regroupant les équipements, circuits, natures d'arrêts, utilisateurs, tonnages, navires, trains et indicateurs de performance. Cette centralisation élimine les fichiers dispersés et garantit une meilleure cohérence des données.

L'un des apports majeurs concerne l'automatisation des traitements actuellement réalisés manuellement. Après chaque saisie ou validation d'arrêt, la solution recalcule automatiquement les temps d'arrêt, les TRS, les débits, les Pareto et les synthèses journalières ou mensuelles. Cette automatisation réduit les délais de traitement et supprime les risques liés aux recalculs Excel répétitifs.

La solution prévoit également des dashboards décisionnels permettant le suivi temps réel des opérations : arrêts en cours, performances des circuits, KPI principaux, tendances d'évolution et alertes critiques. Les responsables disposent ainsi d'une vision globale et instantanée de l'état des opérations de manutention.

Enfin, TRACE-PORT assure une traçabilité complète grâce à l'historisation automatique des actions : création, modification, validation, rejet et suppression. Cette fonctionnalité renforce la fiabilité des données et facilite les audits, les analyses futures et la démarche SMQE.

La valeur ajoutée principale est la disponibilité immédiate de l'information. Là où le système actuel rend les rapports disponibles après plusieurs étapes de consolidation manuelle, TRACE-PORT fournit des indicateurs et synthèses quasi instantanés. La solution supprime ainsi plusieurs tâches sans valeur ajoutée identifiées dans la phase Analyze : double saisie, consolidation manuelle, recalculs répétitifs et vérifications Excel. Elle s'inscrit pleinement dans une logique Lean de réduction des gaspillages et d'amélioration de l'efficacité globale.

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
| Workflow de validation | Fiabiliser les données avant exploitation | Statuts cible : saisi, vérifié, validé, rejeté, clôturé. |
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
| `audit_log` | `user`, `action`, `entity`, `oldValue`, `newValue`, `timestamp`. |
| `users_roles` | `userId`, `role`, `permissions`, `team`. |

## Recommandation d'évolution

Pour une version industrielle, la maquette peut évoluer vers :

- Backend `PostgreSQL` ou `SQL Server`.
- Authentification par rôle : Poste de Commande, Maintenance, Exploitation, HSE, Direction.
- Import Excel mensuel automatisé.
- Workflow de validation des arrêts : saisi, vérifié, validé, clôturé.
- Export PDF/Excel du rapport journalier, mensuel et annuel.
- Module actions DMAIC : cause racine, action corrective, responsable, échéance, gain.
