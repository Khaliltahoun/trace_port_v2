# TRACE-PORT Digital

Application locale de digitalisation du fichier `Bilan et synthéses des arrêts du JANVIER 2026 VF.xlsx`.

## Contenu livré

- `index.html` : interface de pilotage Poste de Commande.
- `app.js` : moteur de calcul des KPI et vues interactives.
- `app-data.js` : données extraites du classeur Excel.
- `styles.css` : mise en forme responsive.
- `tools/extract-workbook.ps1` : extracteur Excel vers base JavaScript.
- `static-server.mjs` : serveur local optionnel.
- `solution-digitale-trace-port.md` : note de conception PFE.

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
