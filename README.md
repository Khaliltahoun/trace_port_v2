# TRACE-PORT Digital

Application locale de digitalisation du fichier `Bilan et synthéses des arrêts du JANVIER 2026 VF.xlsx`.

TRACE-PORT vise à remplacer le suivi manuel des arrêts de manutention par une plateforme centralisée pour la saisie, la validation, le calcul automatique des KPI, le suivi temps réel et l'historisation des actions.

Le fichier Excel source représente un seul mois. TRACE-PORT conserve cette logique : les données sont saisies jour par jour, puis la solution génère une synthèse journalière et une synthèse mensuelle équivalente à la première feuille Excel envoyée aux services en fin de mois.

## Contenu livré

- `index.html` : interface de pilotage Poste de Commande.
- `app.js` : moteur de calcul des KPI et vues interactives.
- `app-data.js` : données extraites du classeur Excel.
- `styles.css` : mise en forme responsive.
- `tools/extract-workbook.ps1` : extracteur Excel vers base JavaScript.
- `static-server.mjs` : serveur local optionnel.
- `solution-digitale-trace-port.md` : note de conception PFE.

## Vision fonctionnelle

- Saisie directe des arrêts par les agents de quart.
- Synthèse journalière recalculée à partir des arrêts, trains, navires et tonnages du jour.
- Synthèse mensuelle consolidant les données quotidiennes du mois.
- Ajout des trains reçus tout au long de la journée avec wagons, tonnage, durée, retard et cadence.
- Ajout des navires avec qualité, début/fin de chargement, bascule, connaissement et écart.
- Vérification et validation par les chefs d'équipe.
- Base centralisée pour équipements, circuits, familles d'arrêts, utilisateurs et indicateurs.
- Recalcul automatique des temps d'arrêt, TRS, débits, Pareto et synthèses.
- Dashboards de pilotage pour arrêts en cours, performances, tendances et alertes.
- Historisation des créations, modifications, validations, rejets et suppressions.
- Réduction des gaspillages Lean : double saisie, consolidation manuelle, recalculs répétitifs et contrôles Excel.

## Lancement

Le fichier `index.html` peut être ouvert directement dans le navigateur.

Option serveur local :

```powershell
node static-server.mjs
```

Puis ouvrir :

```text
http://127.0.0.1:4173
```

## Actualiser les données depuis Excel

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\extract-workbook.ps1
```

L'extraction actuelle contient :

- 863 arrêts et anomalies depuis `Bilan`.
- 3 582 formules Excel inventoriées.
- 58 lignes de requêtes/références depuis `EXPORTER` et les feuilles `Feuil*`.
- 31 jours de tonnage.
- 22 lignes de déchargement trains.
- 10 navires.
