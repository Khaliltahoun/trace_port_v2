# Architecture TRACE-PORT

TRACE-PORT est conçu comme une plateforme industrielle de pilotage des arrêts de manutention. La solution remplace les fichiers Excel mensuels par une application centralisée, traçable et orientée temps réel.

## Architecture 3 tiers

### Frontend

- React.js ou Next.js
- TypeScript
- TailwindCSS
- Shadcn/UI
- Recharts ou Chart.js
- Modules: dashboard, nouvel arrêt, mes arrêts, validation, détail arrêt, KPI, Pareto, rapports, référentiels, utilisateurs, logs

### Backend

- Node.js Express
- API REST
- JWT pour l'authentification
- RBAC pour les rôles Agent, Chef d'équipe, Superviseur, Responsable, Administrateur
- Services métier: stops, validations, kpis, reports, referentials, logs

### Base de données

- PostgreSQL
- Tables principales: users, roles, equipments, circuits, stop_types, stops, validations, kpis, reports, logs
- Historisation de chaque action sensible

## Workflow métier

1. L'agent de quart saisit un arrêt depuis le formulaire digital.
2. TRACE-PORT contrôle automatiquement les champs obligatoires et la cohérence début/fin.
3. L'arrêt passe au statut `pending`.
4. Le chef d'équipe valide ou rejette l'arrêt avec commentaire.
5. Les KPI, dashboards, Pareto et rapports sont recalculés automatiquement.
6. Chaque action est enregistrée dans les logs.

## API REST cible

| Méthode | Endpoint | Rôle | Description |
| --- | --- | --- | --- |
| POST | `/api/auth/login` | Public | Authentification JWT |
| GET | `/api/stops` | Agent+ | Liste des arrêts |
| POST | `/api/stops` | Agent+ | Création d'un arrêt |
| GET | `/api/stops/:id` | Agent+ | Détail d'un arrêt |
| POST | `/api/stops/:id/validate` | ChefEquipe+ | Validation |
| POST | `/api/stops/:id/reject` | ChefEquipe+ | Rejet |
| GET | `/api/kpis` | Superviseur+ | KPI temps réel |
| GET | `/api/reports/:type` | Superviseur+ | Rapport journalier ou mensuel |
| GET | `/api/logs` | Administrateur | Historique des actions |

## KPI digitalisés depuis Excel

- Durée arrêt = heure fin - heure début
- Synthèse par famille = agrégation type `SUMIFS`
- TRS exploitation = `(durée affectation - arrêts exploitation) / durée affectation`
- TRS maintenance = `(durée affectation - arrêts maintenance) / durée affectation`
- TRS global = `(durée affectation - maintenance - exploitation) / durée affectation`
- Débit global = `tonnage / heures de marche`
- Ecart navire = `(connaissement - bascule) / connaissement`

## Livrables dans ce dépôt

- Application web fonctionnelle dans `index.html`, `styles.css`, `app.js`, `app-data.js`
- Serveur local statique dans `static-server.mjs`
- Schéma PostgreSQL dans `database/schema.sql`
- Squelette API Express/JWT/RBAC dans `backend/src/server.js`
- Documentation fonctionnelle dans `solution-digitale-trace-port.md`
